use std::time::Duration;

use windows::{
    core::{ComInterface, Result},
    Win32::{
        Foundation::RECT,
        Graphics::{
            Direct3D11::{ID3D11Device, ID3D11Texture2D},
            Dxgi::{
                IDXGIAdapter1, IDXGIDevice, IDXGIOutput, IDXGIOutput1, IDXGIOutputDuplication,
                IDXGIResource, DXGI_ERROR_ACCESS_LOST, DXGI_ERROR_NOT_FOUND, DXGI_ERROR_WAIT_TIMEOUT,
                DXGI_OUTDUPL_FRAME_INFO, DXGI_OUTPUT_DESC, DXGI_RESOURCE_PRIORITY_MAXIMUM,
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
    let adapter: IDXGIAdapter1 = unsafe { dxgi_device.GetAdapter()? }.cast()?;

    let mut output_index = 0;
    loop {
        let output: IDXGIOutput = match unsafe { adapter.EnumOutputs(output_index) } {
            Ok(output) => output,
            Err(err) if err.code() == DXGI_ERROR_NOT_FOUND => {
                return Err(windows::core::Error::new(
                    DXGI_ERROR_ACCESS_LOST, // Or a more specific error
                    "Monitor not found on adapter".into(),
                ));
            }
            Err(err) => return Err(err.into()),
        };

        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        unsafe { output.GetDesc(&mut desc)? };

        if desc.Monitor == monitor_handle {
            return output.cast();
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
        let output = get_dxgi_output_from_hmonitor(&d3d_device, monitor_handle)?;
        let duplication = unsafe { output.DuplicateOutput(&d3d_device)? };

        // Set max resource priority to ensure duplication continues during fullscreen games
        let resource: IDXGIResource = duplication.cast()?;
        unsafe { resource.SetEvictionPriority(DXGI_RESOURCE_PRIORITY_MAXIMUM)? };

        // Get output dimensions
        let mut desc: DXGI_OUTPUT_DESC = Default::default();
        unsafe { output.GetDesc(&mut desc)? };
        let width = (desc.DesktopCoordinates.right - desc.DesktopCoordinates.left) as u32;
        let height = (desc.DesktopCoordinates.bottom - desc.DesktopCoordinates.top) as u32;

        Ok(Self {
            _d3d_device: d3d_device,
            duplication,
            buffer_texture: None, // Initialize later if needed, or handle directly
            width,
            height,
        })
    }

    pub fn resolution(&self) -> (u32, u32) {
        (self.width, self.height)
    }

    // Timeout is in milliseconds
    pub fn try_get_next_frame(&mut self, timeout_ms: u32) -> Result<Option<AcquiredFrame>> {
        let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
        let mut desktop_resource: Option<IDXGIResource> = None;

        let acquire_result = unsafe {
            self.duplication
                .AcquireNextFrame(timeout_ms, &mut frame_info, &mut desktop_resource)
        };

        match acquire_result {
            Ok(_) => {
                let desktop_resource = desktop_resource
                    .expect("AcquireNextFrame succeeded but returned null resource");
                let texture: ID3D11Texture2D = desktop_resource.cast()?;
                // Release the frame immediately after getting the texture
                unsafe { self.duplication.ReleaseFrame()? };
                Ok(Some(AcquiredFrame {
                    texture,
                    frame_info,
                }))
            }
            Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => Ok(None), // Timeout is not an error condition here
            Err(err) if err.code() == DXGI_ERROR_ACCESS_LOST => {
                // Re-initialization might be needed, signal this as an error
                Err(err.into())
            }
            Err(err) => Err(err.into()), // Other errors
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
