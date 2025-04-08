use std::time::Duration;

use windows::{
    core::{ComInterface, Result},
    Win32::{
        Foundation::RECT,
        Graphics::{
            Direct3D11::{ID3D11Device, ID3D11Texture2D},
            Dxgi::{
                IDXGIAdapter, IDXGIAdapter1, IDXGIDevice, IDXGIOutput, IDXGIOutput1, IDXGIOutputDuplication, IDXGIResource, DXGI_ERROR_ACCESS_LOST, DXGI_ERROR_NOT_FOUND, DXGI_ERROR_WAIT_TIMEOUT, DXGI_OUTDUPL_FRAME_INFO, DXGI_OUTPUT_DESC, DXGI_RESOURCE_PRIORITY_MAXIMUM
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
    
    let dxgi_device: IDXGIDevice = match d3d_device.cast() {
        Ok(device) => {
            println!("Successfully cast to IDXGIDevice");
            device
        },
        Err(e) => {
            println!("Failed to cast to IDXGIDevice: {:?}", e);
            return Err(e);
        }
    };
    
    let adapter: IDXGIAdapter = match unsafe { dxgi_device.GetAdapter() } {
        Ok(adapter) => {
            println!("Successfully got IDXGIAdapter");
            adapter
        },
        Err(e) => {
            println!("Failed to get adapter: {:?}", e);
            return Err(e);
        }
    };

    let mut output_index = 0;
    loop {
        println!("Trying to enumerate output {}", output_index);
        let output: IDXGIOutput = match unsafe { adapter.EnumOutputs(output_index) } {
            Ok(output) => {
                println!("Successfully got output {}", output_index);
                output
            },
            Err(err) if err.code() == DXGI_ERROR_NOT_FOUND => {
                println!("No more outputs found at index {}", output_index);
                return Err(windows::core::Error::new(
                    DXGI_ERROR_ACCESS_LOST,
                    "Monitor not found on adapter".into(),
                ));
            }
            Err(err) => {
                println!("Error enumerating output {}: {:?}", output_index, err);
                return Err(err.into());
            }
        };

        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        match unsafe { output.GetDesc(&mut desc) } {
            Ok(_) => {
                println!("Got output description for output {}", output_index);
                if desc.Monitor == monitor_handle {
                    println!("Found matching monitor at output index {}", output_index);
                    return output.cast();
                }
            },
            Err(e) => {
                println!("Failed to get output description: {:?}", e);
                return Err(e);
            }
        }

        output_index += 1;
    }
}

pub struct AcquiredFrame {
    pub texture: ID3D11Texture2D,
    pub frame_info: DXGI_OUTDUPL_FRAME_INFO,
}

pub struct CaptureFrameGenerator {
    _d3d_device: ID3D11Device, // Keep the device alive
    duplication: IDXGIOutputDuplication,
    buffer_texture: Option<ID3D11Texture2D>, // Reusable texture for copying
    width: u32,
    height: u32,
}

impl CaptureFrameGenerator {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
    ) -> Result<Self> {
        println!("Getting DXGI output from monitor handle");
        let output = get_dxgi_output_from_hmonitor(&d3d_device, monitor_handle)?;
        println!("Successfully got DXGI output");
        
        println!("Attempting to duplicate output");
        let duplication_result = unsafe { output.DuplicateOutput(&d3d_device) };
        match &duplication_result {
            Ok(_) => println!("Successfully duplicated output"),
            Err(e) => {
                eprintln!("Failed to duplicate output: {:?}. Ensure monitor supports duplication.", e);
                return Err(e.clone());
            }
        }
        let duplication = duplication_result?;
        println!("Successfully duplicated output");
        
        // Get output dimensions
        println!("Getting output description");
        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        let desc_result = unsafe { output.GetDesc(&mut desc) };
        match &desc_result {
            Ok(_) => println!("Get output description succeeded"),
            Err(e) => println!("Get output description failed with error: {:?}", e),
        }
        desc_result?;
        
        let width = (desc.DesktopCoordinates.right - desc.DesktopCoordinates.left) as u32;
        let height = (desc.DesktopCoordinates.bottom - desc.DesktopCoordinates.top) as u32;
        println!("Output dimensions: {}x{}", width, height);

        Ok(Self {
            _d3d_device: d3d_device,
            duplication,
            buffer_texture: None,
            width,
            height,
        })
    }

    pub fn resolution(&self) -> (u32, u32) {
        (self.width, self.height)
    }

    // Timeout is in milliseconds
    pub fn try_get_next_frame(&mut self, timeout_ms: u32) -> Result<Option<AcquiredFrame>> {
        println!("Attempting to acquire next frame with timeout: {} ms", timeout_ms);
        
        let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
        let mut desktop_resource: Option<IDXGIResource> = None;
    
        println!("Calling AcquireNextFrame...");
        let acquire_result = unsafe {
            self.duplication
                .AcquireNextFrame(timeout_ms, &mut frame_info, &mut desktop_resource)
        };
    
        match acquire_result {
            Ok(_) => {
                println!("AcquireNextFrame succeeded, frame info: {:?}", frame_info);
                
                let desktop_resource = desktop_resource
                    .expect("AcquireNextFrame succeeded but returned null resource");
                println!("Got desktop resource, casting to ID3D11Texture2D");
                
                let texture: ID3D11Texture2D = desktop_resource.cast()?;
                println!("Successfully cast resource to texture");
                
                // Release the frame immediately after getting the texture
                println!("Releasing frame");
                unsafe { self.duplication.ReleaseFrame()? };
                
                println!("Frame acquisition complete");
                Ok(Some(AcquiredFrame {
                    texture,
                    frame_info,
                }))
            }
            Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => {
                println!("Frame acquisition timed out after {} ms", timeout_ms);
                Ok(None) // Timeout is not an error condition here
            }
            Err(err) if err.code() == DXGI_ERROR_ACCESS_LOST => {
                println!("ERROR: Access lost to desktop duplication: {:?}", err);
                // Re-initialization might be needed, signal this as an error
                Err(err.into())
            }
            Err(err) => {
                println!("ERROR: Failed to acquire frame: {:?}", err);
                Err(err.into()) // Other errors
            }
        }
    }

    // Optional: Provide a blocking version
    pub fn get_next_frame(&mut self) -> Result<AcquiredFrame> {
        loop {
            match self.try_get_next_frame(u32::MAX) { // Use a very long timeout for blocking
                Ok(Some(frame)) => return Ok(frame),
                Ok(None) => continue, // Should not happen with infinite timeout, but handle defensively
                Err(e) if e.code() == DXGI_ERROR_ACCESS_LOST => {
                    // Handle lost access, maybe try to reinitialize or return error
                    return Err(e);
                }
                Err(e) => return Err(e), // Propagate other errors
            }
        }
    }

    // The stop signal mechanism is removed as Desktop Duplication doesn't use callbacks.
    // Stopping is handled by dropping the CaptureFrameGenerator.
}

impl Drop for CaptureFrameGenerator {
    fn drop(&mut self) {
        // IDXGIOutputDuplication is released automatically when dropped
        // No explicit Close needed like Graphics Capture
        println!("Dropping CaptureFrameGenerator and releasing duplication.");
    }
}

// Remove CaptureFrameGeneratorStopSignal as it's no longer needed
