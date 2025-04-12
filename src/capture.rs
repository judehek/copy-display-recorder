use std::sync::mpsc::{channel, Receiver, Sender};
use std::thread;
use std::time::Duration;

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
    println!("Entering get_dxgi_output_from_hmonitor");
    
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
}

impl CaptureFrameGenerator {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
    ) -> Result<Self> {
        println!("Initializing CaptureFrameGenerator");
        // Create channels for frames and control
        let (frame_sender, frame_receiver) = channel();
        let (control_sender, control_receiver) = channel();

        let frame_sender_for_struct = frame_sender.clone();
        
        println!("Creating capture session");
        // Create session
        let session = CustomGraphicsCaptureSession::new(control_sender.clone());
        
        println!("Getting DXGI output from monitor handle");
        // Get output and create duplication
        let output = get_dxgi_output_from_hmonitor(&d3d_device, monitor_handle)?;
        println!("Creating output duplication");
        let duplication = unsafe { output.DuplicateOutput(&d3d_device)? };
        
        // Get output dimensions
        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        unsafe { output.GetDesc(&mut desc)? };
        println!("Output dimensions: {:?}x{:?}", desc.DesktopCoordinates.right - desc.DesktopCoordinates.left, 
                                                desc.DesktopCoordinates.bottom - desc.DesktopCoordinates.top);
        
        // Clone necessary values for the capture thread
        let d3d_device_clone = d3d_device.clone();
        
        println!("Starting capture thread");
        // Start background thread to poll for frames
        thread::spawn(move || {
            println!("Capture thread started");
            let mut running = false;
            let mut buffer_texture: Option<ID3D11Texture2D> = None;
            let mut frame_count = 0;
            
            'outer: loop {
                // Check for control messages first
                match control_receiver.try_recv() {
                    Ok(start_signal) => {
                        running = start_signal;
                        if running {
                            println!("Capture started");
                        } else {
                            println!("Capture stopped, exiting thread");
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
                
                // Try to get a frame with timeout
                let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
                let mut desktop_resource: Option<IDXGIResource> = None;
                
                let acquire_result = unsafe {
                    duplication.AcquireNextFrame(100, &mut frame_info, &mut desktop_resource)
                };
                
                match acquire_result {
                    Ok(_) => {
                        frame_count += 1;
                        if frame_count % 100 == 0 {
                            println!("Acquired frame #{}: LastPresentTime={:?}, LastMouseUpdateTime={:?}",
                                    frame_count, frame_info.LastPresentTime, frame_info.LastMouseUpdateTime);
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
                                            duplication.ReleaseFrame().unwrap_or_default();
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
                                // Either continue the loop, break, or take another appropriate action
                                continue;
                            }
                        };
                        unsafe { context.CopyResource(&target_texture, &acquired_texture) };
                        
                        // Release the frame back to duplication
                        if let Err(e) = unsafe { duplication.ReleaseFrame() } {
                            eprintln!("Failed to release frame: {:?}", e);
                            continue;
                        }
                        
                        // Create and send the frame
                        let frame = AcquiredFrame {
                            texture: target_texture,
                            frame_info,
                        };
                        
                        if frame_sender.send(Some(frame)).is_err() {
                            println!("Failed to send frame, receiver disconnected");
                            break 'outer; // Exit if channel is closed
                        }
                    },
                    Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => {
                        // No frame available, just continue
                        thread::sleep(Duration::from_millis(1));
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
            println!("Capture thread exiting, processed {} frames", frame_count);
        });

        println!("CaptureFrameGenerator initialization complete");
        Ok(Self {
            d3d_device,
            sender: frame_sender_for_struct,
            receiver: frame_receiver,
            session,
        })
    }

    pub fn session(&self) -> &CustomGraphicsCaptureSession {
        &self.session
    }

    // New simplified function that just receives from the channel
    pub fn try_get_next_frame(&mut self) -> Result<Option<AcquiredFrame>> {
        match self.receiver.recv() {
            Ok(frame) => {
                match &frame {
                    Some(_) => println!("Received new frame"),
                    None => println!("Received end-of-capture signal"),
                }
                Ok(frame)
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
}