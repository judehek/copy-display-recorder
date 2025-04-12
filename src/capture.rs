use std::sync::mpsc::{channel, Receiver, Sender};
use std::thread;
use std::time::{Duration, Instant};

use windows::Foundation::TimeSpan;
use windows::Win32::System::Performance::QueryPerformanceCounter;
use windows::{
    core::{ComInterface, Result},
    Win32::{
        Foundation::{E_FAIL, RECT},
        Graphics::{
            Direct3D11::{ID3D11Device, ID3D11Texture2D, D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT},
            Dxgi::{
                IDXGIAdapter, IDXGIDevice, IDXGIOutput, IDXGIOutput1, IDXGIOutputDuplication, IDXGIResource, 
                DXGI_ERROR_ACCESS_LOST, DXGI_ERROR_NOT_FOUND, DXGI_ERROR_WAIT_TIMEOUT, 
                DXGI_OUTDUPL_FRAME_INFO, DXGI_OUTPUT_DESC
            },
            Gdi::HMONITOR,
        },
    },
};

// Helper function to get IDXGIOutput1 from HMONITOR
fn get_dxgi_output_from_hmonitor(
    d3d_device: &ID3D11Device,
    monitor_handle: HMONITOR,
) -> Result<IDXGIOutput1> {
    let dxgi_device: IDXGIDevice = d3d_device.cast()?;
    let adapter: IDXGIAdapter = unsafe { dxgi_device.GetAdapter()? };

    let mut output_index = 0;
    loop {
        let output: IDXGIOutput = match unsafe { adapter.EnumOutputs(output_index) } {
            Ok(output) => output,
            Err(err) if err.code() == DXGI_ERROR_NOT_FOUND => {
                return Err(windows::core::Error::new(
                    DXGI_ERROR_ACCESS_LOST,
                    "Monitor not found on adapter".into(),
                ));
            }
            Err(err) => return Err(err),
        };

        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        unsafe { output.GetDesc(&mut desc)? };
        
        if desc.Monitor == monitor_handle {
            return output.cast();
        }

        output_index += 1;
    }
}

#[derive(Clone)]
pub struct AcquiredFrame {
    pub texture: ID3D11Texture2D,
    pub frame_info: DXGI_OUTDUPL_FRAME_INFO,
    pub present_time: TimeSpan,
    // Add timing information to track performance
    pub acquisition_time_ms: u64,
}

// Mimics GraphicsCaptureSession from the Windows API
pub struct CustomGraphicsCaptureSession {
    sender: Sender<bool>, // To signal start/stop
    running: bool,
}

impl CustomGraphicsCaptureSession {
    fn new(sender: Sender<bool>) -> Self {
        Self {
            sender,
            running: false,
        }
    }

    pub fn StartCapture(&mut self) -> Result<()> {
        if !self.running {
            self.sender.send(true).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = true;
        }
        Ok(())
    }

    pub fn Close(&mut self) -> Result<()> {
        if self.running {
            self.sender.send(false).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = false;
        }
        Ok(())
    }
}

impl Clone for CustomGraphicsCaptureSession {
    fn clone(&self) -> Self {
        Self {
            sender: self.sender.clone(),
            running: self.running,
        }
    }
}

pub struct CaptureFrameGenerator {
    d3d_device: ID3D11Device,
    sender: Sender<Option<AcquiredFrame>>,
    receiver: Receiver<Option<AcquiredFrame>>,
    session: CustomGraphicsCaptureSession,
    // Performance stats
    success_frames: u64,
    timeout_frames: u64,
    avg_success_time_ms: f64,
    avg_timeout_time_ms: f64,
}

impl CaptureFrameGenerator {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
    ) -> Result<Self> {
        // Create channels for frames and control
        let (frame_sender, frame_receiver) = channel();
        let (control_sender, control_receiver) = channel();

        let frame_sender_for_struct = frame_sender.clone();
        
        // Create session
        let session = CustomGraphicsCaptureSession::new(control_sender.clone());
        
        // Get output and create duplication
        let output = get_dxgi_output_from_hmonitor(&d3d_device, monitor_handle)?;
        let duplication = unsafe { output.DuplicateOutput(&d3d_device)? };
        
        // Get output dimensions
        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        unsafe { output.GetDesc(&mut desc)? };
        println!("Output dimensions: {}x{}", 
                 desc.DesktopCoordinates.right - desc.DesktopCoordinates.left, 
                 desc.DesktopCoordinates.bottom - desc.DesktopCoordinates.top);
        
        // Clone necessary values for the capture thread
        let d3d_device_clone = d3d_device.clone();
        
        // Start background thread to poll for frames
        thread::spawn(move || {
            println!("Capture thread started");
            let mut running = false;
            let mut buffer_texture: Option<ID3D11Texture2D> = None;
            let mut frame_count = 0;
            let mut last_texture: Option<ID3D11Texture2D> = None; // Track last texture for duplication
            
            // Performance tracking
            let mut timeout_count = 0;
            let mut success_count = 0;
            let mut total_success_time_ms = 0u64;
            let mut total_timeout_time_ms = 0u64;
            
            'outer: loop {
                // Check for control messages first
                match control_receiver.try_recv() {
                    Ok(start_signal) => {
                        running = start_signal;
                        if running {
                            println!("Capture started");
                        } else {
                            println!("Capture stopped, exiting thread");
                            
                            // Print final performance stats
                            if success_count > 0 || timeout_count > 0 {
                                println!("Performance statistics:");
                                println!("  Total frames processed: {}", frame_count);
                                println!("  Successful acquisitions: {} ({:.2}%)", 
                                         success_count, 
                                         (success_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                                println!("  Timeouts: {} ({:.2}%)", 
                                         timeout_count, 
                                         (timeout_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                                
                                if success_count > 0 {
                                    println!("  Avg time for successful acquisitions: {:.2}ms", 
                                             total_success_time_ms as f64 / success_count as f64);
                                }
                                
                                if timeout_count > 0 {
                                    println!("  Avg time for timeouts: {:.2}ms", 
                                             total_timeout_time_ms as f64 / timeout_count as f64);
                                }
                            }
                            
                            // Signal end of capture
                            let _ = frame_sender.send(None);
                            break;
                        }
                    },
                    Err(std::sync::mpsc::TryRecvError::Empty) => {
                        // No control messages, continue
                    },
                    Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                        println!("Control channel disconnected, exiting thread");
                        break; // Exit if channel is closed
                    }
                }
                
                if !running {
                    thread::sleep(Duration::from_millis(10));
                    continue;
                }
                
                // Start timing this acquisition attempt
                let start_time = Instant::now();
                
                // Try to get a frame with timeout
                let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
                let mut desktop_resource: Option<IDXGIResource> = None;
                
                let acquire_start_time = Instant::now();
                let acquire_result = unsafe {
                    duplication.AcquireNextFrame(0, &mut frame_info, &mut desktop_resource)
                };
                let acquire_elapsed_ms = acquire_start_time.elapsed().as_millis() as u64;
                
                match acquire_result {
                    Ok(_) => {
                        // Successful frame acquisition
                        success_count += 1;
                        frame_count += 1;

                        // Log just the AcquireNextFrame timing for successes
                        if frame_count % 100 == 0 {
                            println!("Success AcquireNextFrame time: {}ms", acquire_elapsed_ms);
                        }
                        
                        let desktop_resource = desktop_resource
                            .expect("AcquireNextFrame succeeded but returned null resource");
                        
                        let acquired_texture: ID3D11Texture2D = match desktop_resource.cast() {
                            Ok(texture) => texture,
                            Err(e) => {
                                eprintln!("Failed to cast resource to texture: {:?}", e);
                                unsafe { duplication.ReleaseFrame().unwrap_or_default() };
                                continue;
                            }
                        };
                        
                        // Create or reuse buffer texture for copying
                        let source_desc = {
                            let mut desc = D3D11_TEXTURE2D_DESC::default();
                            unsafe { acquired_texture.GetDesc(&mut desc) };
                            desc
                        };
                        
                        if buffer_texture.is_none() {
                            println!("Creating buffer texture with dimensions: {}x{}, format: {:?}", 
                                    source_desc.Width, source_desc.Height, source_desc.Format);
                        }
                        
                        let buffer_desc = D3D11_TEXTURE2D_DESC {
                            Width: source_desc.Width,
                            Height: source_desc.Height,
                            MipLevels: 1,
                            ArraySize: 1,
                            Format: source_desc.Format,
                            SampleDesc: source_desc.SampleDesc,
                            Usage: D3D11_USAGE_DEFAULT,
                            BindFlags: 0,
                            CPUAccessFlags: 0,
                            MiscFlags: 0,
                            ..Default::default()
                        };
                        
                        let target_texture = match &buffer_texture {
                            Some(texture) => texture.clone(),
                            None => {
                                println!("Creating new buffer texture");
                                let new_texture = unsafe {
                                    let mut texture = None;
                                    match d3d_device_clone.CreateTexture2D(&buffer_desc, None, Some(&mut texture)) {
                                        Ok(_) => texture.unwrap(),
                                        Err(e) => {
                                            eprintln!("Failed to create buffer texture: {:?}", e);
                                            unsafe { duplication.ReleaseFrame().unwrap_or_default() };
                                            continue;
                                        }
                                    }
                                };
                                buffer_texture = Some(new_texture.clone());
                                new_texture
                            }
                        };
                        
                        // Copy the acquired frame to our buffer
                        let context = match unsafe { d3d_device_clone.GetImmediateContext() } {
                            Ok(ctx) => ctx,
                            Err(e) => {
                                eprintln!("Failed to get immediate context: {:?}", e);
                                continue;
                            }
                        };
                        unsafe { context.CopyResource(&target_texture, &acquired_texture) };
                        
                        // Release the frame back to duplication
                        if let Err(e) = unsafe { duplication.ReleaseFrame() } {
                            eprintln!("Failed to release frame: {:?}", e);
                            continue;
                        }
                        
                        // Get our own QPC timestamp
                        let present_time = match get_qpc_timestamp() {
                            Ok(timestamp) => timestamp,
                            Err(e) => {
                                eprintln!("Failed to get QPC timestamp: {:?}", e);
                                continue;
                            }
                        };
                        
                        // Calculate elapsed time
                        let elapsed = start_time.elapsed();
                        let elapsed_ms = elapsed.as_millis() as u64;
                        total_success_time_ms += elapsed_ms;
                        
                        // Log performance stats periodically
                        if frame_count % 100 == 0 {
                            println!("Frame #{}: Success time: {}ms, Avg success: {:.2}ms, Avg timeout: {:.2}ms, Success ratio: {:.2}%",
                                     frame_count, 
                                     elapsed_ms,
                                     total_success_time_ms as f64 / success_count as f64,
                                     if timeout_count > 0 { total_timeout_time_ms as f64 / timeout_count as f64 } else { 0.0 },
                                     (success_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                        }
                        
                        // Create and send the frame with timing information
                        let frame = AcquiredFrame {
                            texture: target_texture.clone(),
                            frame_info,
                            present_time,
                            acquisition_time_ms: elapsed_ms,
                        };
                        
                        // Store this texture for duplication in case of timeout
                        last_texture = Some(target_texture);
                        
                        if frame_sender.send(Some(frame)).is_err() {
                            println!("Failed to send frame, receiver disconnected");
                            break 'outer; // Exit if channel is closed
                        }
                    },
                    Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => {
                        // Handle timeout case with timing
                        timeout_count += 1;

                        if timeout_count % 100 == 0 {
                            println!("Timeout AcquireNextFrame time: {}ms", acquire_elapsed_ms);
                        }
                        
                        // Calculate elapsed time for timeout
                        let elapsed = start_time.elapsed();
                        let elapsed_ms = elapsed.as_millis() as u64;
                        total_timeout_time_ms += elapsed_ms;
                        
                        if timeout_count % 100 == 0 {
                            println!("Timeout #{}: Time: {}ms, Avg timeout: {:.2}ms, Avg success: {:.2}ms, Timeout ratio: {:.2}%", 
                                    timeout_count, 
                                    elapsed_ms,
                                    total_timeout_time_ms as f64 / timeout_count as f64,
                                    if success_count > 0 { total_success_time_ms as f64 / success_count as f64 } else { 0.0 },
                                    (timeout_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                        }
                        
                        // No new frame available - use last frame with a new timestamp if we have one
                        if let Some(last_tex) = &last_texture {
                            // Get a new QPC timestamp
                            let present_time = match get_qpc_timestamp() {
                                Ok(timestamp) => timestamp,
                                Err(e) => {
                                    eprintln!("Failed to get QPC timestamp for duplicate frame: {:?}", e);
                                    thread::sleep(Duration::from_millis(1));
                                    continue;
                                }
                            };
                            
                            // Create a duplicate frame with the new timestamp
                            let frame = AcquiredFrame {
                                texture: last_tex.clone(),
                                frame_info: Default::default(), // Default frame info for duplicate
                                present_time,
                                acquisition_time_ms: elapsed_ms,
                            };
                            
                            if frame_sender.send(Some(frame)).is_err() {
                                println!("Failed to send duplicate frame, receiver disconnected");
                                break 'outer;
                            }
                        thread::sleep(Duration::from_millis(1));
                        } else {
                            // No last texture yet, just wait a bit
                            thread::sleep(Duration::from_millis(1));
                        }
                    },
                    Err(err) if err.code() == DXGI_ERROR_ACCESS_LOST => {
                        println!("Access lost to output duplication, ending capture");
                        // Critical error, signal end of capture
                        let _ = frame_sender.send(None);
                        break;
                    },
                    Err(err) => {
                        // Log other errors but continue
                        eprintln!("Error acquiring frame: {:?}", err);
                    }
                }
            }
            println!("Capture thread exiting, processed {} frames, had {} timeouts", frame_count, timeout_count);
            
            // Print final performance stats
            if success_count > 0 || timeout_count > 0 {
                println!("Final performance statistics:");
                println!("  Total frames: {}", frame_count);
                if success_count > 0 {
                    println!("  Successful acquisitions: {} ({:.2}%)", 
                             success_count, 
                             (success_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                    println!("  Avg time for successful acquisitions: {:.2}ms", 
                             total_success_time_ms as f64 / success_count as f64);
                }
                
                if timeout_count > 0 {
                    println!("  Timeouts: {} ({:.2}%)", 
                             timeout_count, 
                             (timeout_count as f32 / (success_count + timeout_count) as f32) * 100.0);
                    println!("  Avg time for timeouts: {:.2}ms", 
                             total_timeout_time_ms as f64 / timeout_count as f64);
                }
                
                println!("  Time difference: {:.2}ms", 
                         (total_success_time_ms as f64 / success_count as f64) - 
                         (total_timeout_time_ms as f64 / timeout_count as f64));
            }
        });

        Ok(Self {
            d3d_device,
            sender: frame_sender_for_struct,
            receiver: frame_receiver,
            session,
            // Initialize performance stats
            success_frames: 0,
            timeout_frames: 0,
            avg_success_time_ms: 0.0,
            avg_timeout_time_ms: 0.0,
        })
    }

    pub fn session(&self) -> &CustomGraphicsCaptureSession {
        &self.session
    }

    // New simplified function that just receives from the channel
    pub fn try_get_next_frame(&mut self) -> Result<Option<AcquiredFrame>> {
        match self.receiver.recv() {
            Ok(Some(frame)) => {
                // Update our stats based on the frame
                if frame.frame_info.LastPresentTime > 0 {
                    // This was a new frame
                    self.success_frames += 1;
                    // Update running average
                    self.avg_success_time_ms = ((self.avg_success_time_ms * (self.success_frames - 1) as f64) + 
                                                frame.acquisition_time_ms as f64) / self.success_frames as f64;
                } else {
                    // This was a duplicated frame (timeout)
                    self.timeout_frames += 1;
                    // Update running average
                    self.avg_timeout_time_ms = ((self.avg_timeout_time_ms * (self.timeout_frames - 1) as f64) + 
                                               frame.acquisition_time_ms as f64) / self.timeout_frames as f64;
                }
                
                // Periodically show stats on the receiver side
                if (self.success_frames + self.timeout_frames) % 1000 == 0 {
                    println!("Receiver stats: {} successes (avg: {:.2}ms), {} timeouts (avg: {:.2}ms), diff: {:.2}ms",
                             self.success_frames,
                             self.avg_success_time_ms,
                             self.timeout_frames,
                             self.avg_timeout_time_ms,
                             self.avg_success_time_ms - self.avg_timeout_time_ms);
                }
                
                Ok(Some(frame))
            },
            Ok(None) => {
                // End of capture signal
                println!("Received end-of-capture signal");
                println!("Final receiver stats: {} successes (avg: {:.2}ms), {} timeouts (avg: {:.2}ms), diff: {:.2}ms",
                         self.success_frames,
                         self.avg_success_time_ms,
                         self.timeout_frames,
                         self.avg_timeout_time_ms,
                         self.avg_success_time_ms - self.avg_timeout_time_ms);
                Ok(None)
            },
            Err(e) => {
                // Channel closed, end of capture
                println!("Frame channel error: {:?}, signaling end of capture", e);
                Ok(None)
            }
        }
    }

    pub fn stop_capture(&mut self) -> Result<()> {
        println!("Stopping capture");
        self.session.Close()
    }

    // Return performance statistics
    pub fn get_performance_stats(&self) -> (u64, u64, f64, f64) {
        (self.success_frames, self.timeout_frames, self.avg_success_time_ms, self.avg_timeout_time_ms)
    }
}

fn get_qpc_timestamp() -> Result<TimeSpan> {
    let mut qpc_timestamp: i64 = 0;
    unsafe {
        QueryPerformanceCounter(&mut qpc_timestamp)?;
    }
    
    // Create a TimeSpan using the same pattern as your example
    let timestamp = TimeSpan {
        Duration: qpc_timestamp,
    };
    
    Ok(timestamp)
}