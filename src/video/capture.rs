use std::sync::mpsc::{channel, Receiver, Sender};
use std::thread;
use std::time::{Duration, Instant};
use std::sync::Arc;
use std::sync::atomic::{AtomicI64, Ordering};

use windows::core::Error;
use windows::Foundation::TimeSpan;
use windows::Win32::System::Performance::QueryPerformanceCounter;
use windows::{
    core::{Interface, Result},
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
                return Err(Error::new(
                    DXGI_ERROR_ACCESS_LOST,
                    "Monitor not found on adapter",
                ));
            }
            Err(err) => return Err(err),
        };

        let desc = unsafe { output.GetDesc()? };
        
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
}

// Mimics GraphicsCaptureSession from the Windows API
pub struct CustomGraphicsCaptureSession {
    sender: Sender<(bool, i64)>, // Modified to include QPC timestamp
    running: bool,
}

impl CustomGraphicsCaptureSession {
    fn new(sender: Sender<(bool, i64)>) -> Self {
        Self {
            sender,
            running: false,
        }
    }

    pub fn StartCapture(&mut self, start_qpc: i64) -> Result<()> {
        if !self.running {
            self.sender.send((true, start_qpc)).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = true;
        }
        Ok(())
    }

    pub fn Close(&mut self) -> Result<()> {
        if self.running {
            self.sender.send((false, 0)).map_err(|_| windows::core::Error::from(E_FAIL))?;
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
    start_qpc: Arc<AtomicI64>,  // Added to store the reference QPC value
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
        
        // Create atomic for storing the start QPC timestamp
        let start_qpc = Arc::new(AtomicI64::new(0));
        let thread_start_qpc = start_qpc.clone();
        
        // Create session
        let session = CustomGraphicsCaptureSession::new(control_sender.clone());
        
        // Get output and create duplication
        let output = get_dxgi_output_from_hmonitor(&d3d_device, monitor_handle)?;
        let duplication = unsafe { output.DuplicateOutput(&d3d_device)? };
        
        // Get output dimensions
        // Get output dimensions
        let desc = unsafe { output.GetDesc()? };
        // Clone necessary values for the capture thread
        let d3d_device_clone = d3d_device.clone();
        
        // Start background thread to poll for frames
        thread::spawn(move || {
            let mut running = false;
            let mut buffer_texture: Option<ID3D11Texture2D> = None;
            let mut last_texture: Option<ID3D11Texture2D> = None; // Track last texture for duplication
            
            'outer: loop {
                // Check for control messages first
                match control_receiver.try_recv() {
                    Ok((start_signal, new_qpc)) => {
                        running = start_signal;
                        
                        if running {
                            // Update the start QPC value when starting capture
                            thread_start_qpc.store(new_qpc, Ordering::SeqCst);
                            println!("Video capture: Updated start_qpc to: {}", new_qpc);
                        }
                        
                        if !running {
                            // Signal end of capture
                            let _ = frame_sender.send(None);
                            break;
                        }
                    },
                    Err(std::sync::mpsc::TryRecvError::Empty) => {
                        // No control messages, continue
                    },
                    Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                        break; // Exit if channel is closed
                    }
                }
                
                if !running {
                    thread::sleep(Duration::from_millis(10));
                    continue;
                }
                
                // Try to get a frame with timeout
                let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
                let mut desktop_resource: Option<IDXGIResource> = None;
                
                let acquire_result = unsafe {
                    duplication.AcquireNextFrame(0, &mut frame_info, &mut desktop_resource)
                };
                
                match acquire_result {
                    Ok(_) => {
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
                        let qpc_timestamp = match get_raw_qpc_timestamp() {
                            Ok(timestamp) => timestamp,
                            Err(e) => {
                                eprintln!("Failed to get QPC timestamp: {:?}", e);
                                continue;
                            }
                        };
                        
                        // Get the current reference QPC value 
                        let current_start_qpc = thread_start_qpc.load(Ordering::SeqCst);
                        
                        // Create a relative timestamp based on the start_qpc
                        let relative_timestamp = qpc_timestamp - current_start_qpc;
                        let present_time = TimeSpan { Duration: relative_timestamp };
                        
                        // Create and send the frame
                        let frame = AcquiredFrame {
                            texture: target_texture.clone(),
                            frame_info,
                            present_time,
                        };
                        
                        // Store this texture for duplication in case of timeout
                        last_texture = Some(target_texture);
                        
                        if frame_sender.send(Some(frame)).is_err() {
                            break 'outer; // Exit if channel is closed
                        }
                    },
                    Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => {
                        // No new frame available - use last frame with a new timestamp if we have one
                        if let Some(last_tex) = &last_texture {
                            // Get a new QPC timestamp
                            let qpc_timestamp = match get_raw_qpc_timestamp() {
                                Ok(timestamp) => timestamp,
                                Err(e) => {
                                    eprintln!("Failed to get QPC timestamp for duplicate frame: {:?}", e);
                                    thread::sleep(Duration::from_millis(1));
                                    continue;
                                }
                            };
                            
                            // Calculate relative timestamp
                            let current_start_qpc = thread_start_qpc.load(Ordering::SeqCst);
                            let relative_timestamp = qpc_timestamp - current_start_qpc;
                            let present_time = TimeSpan { Duration: relative_timestamp };
                            
                            // Create a duplicate frame with the new timestamp
                            let frame = AcquiredFrame {
                                texture: last_tex.clone(),
                                frame_info: Default::default(), // Default frame info for duplicate
                                present_time,
                            };
                            
                            if frame_sender.send(Some(frame)).is_err() {
                                break 'outer;
                            }
                        
                            // IMPORANT: prevent busy wait due to 0 timeout on acquirenextframe()
                            thread::sleep(Duration::from_millis(1));
                        } else {
                            // No last texture yet, just wait a bit
                            thread::sleep(Duration::from_millis(1));
                        }
                    },
                    Err(err) if err.code() == DXGI_ERROR_ACCESS_LOST => {
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
        });

        Ok(Self {
            d3d_device,
            sender: frame_sender_for_struct,
            receiver: frame_receiver,
            session,
            start_qpc,
        })
    }

    pub fn session(&self) -> &CustomGraphicsCaptureSession {
        &self.session
    }

    // Convenience method to start capture with the given QPC
    pub fn start_capture(&mut self, start_qpc: i64) -> Result<()> {
        self.session.StartCapture(start_qpc)
    }

    // Simplified function that just receives from the channel
    pub fn try_get_next_frame(&mut self) -> Result<Option<AcquiredFrame>> {
        // First wait for at least one frame (or end signal)
        let mut latest_frame = match self.receiver.recv() {
            Ok(Some(frame)) => Some(frame),
            Ok(None) => return Ok(None), // End of capture signal
            Err(_) => return Ok(None),   // Channel closed
        };
        
        // Now drain any additional frames that arrived
        loop {
            match self.receiver.try_recv() {
                Ok(Some(frame)) => {
                    // Keep updating with newer frames
                    latest_frame = Some(frame);
                },
                Ok(None) => {
                    // End of capture signal - return None regardless of what we've seen before
                    return Ok(None);
                },
                Err(std::sync::mpsc::TryRecvError::Empty) => {
                    // No more frames in the channel, break the loop
                    break;
                },
                Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                    // Channel closed, end of capture
                    return Ok(None);
                }
            }
        }
        
        // Return the latest frame we found
        Ok(latest_frame)
    }

    pub fn stop_capture(&mut self) -> Result<()> {
        self.session.Close()
    }
    
    // Getter for the current start_qpc value
    pub fn get_start_qpc(&self) -> i64 {
        self.start_qpc.load(Ordering::SeqCst)
    }
}

// Modified to return raw QPC value without creating TimeSpan
fn get_raw_qpc_timestamp() -> Result<i64> {
    let mut qpc_timestamp: i64 = 0;
    unsafe {
        QueryPerformanceCounter(&mut qpc_timestamp)?;
    }
    Ok(qpc_timestamp)
}

// Original function kept for backward compatibility
fn get_qpc_timestamp() -> Result<TimeSpan> {
    let qpc_timestamp = get_raw_qpc_timestamp()?;
    
    // Create a TimeSpan using the same pattern as your example
    let timestamp = TimeSpan {
        Duration: qpc_timestamp,
    };
    
    Ok(timestamp)
}