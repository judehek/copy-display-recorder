use std::time::Duration;

use windows::{
    core::{ComInterface, Result},
    Win32::{
        Foundation::RECT,
        Graphics::{
            Direct3D11::{ID3D11Device, ID3D11Texture2D, D3D11_CPU_ACCESS_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_MAP_READ, D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT, D3D11_USAGE_STAGING},
            Dxgi::{
                IDXGIAdapter, IDXGIAdapter1, IDXGIDevice, IDXGIOutput, IDXGIOutput1, IDXGIOutputDuplication, IDXGIResource, DXGI_ERROR_ACCESS_LOST, DXGI_ERROR_NOT_FOUND, DXGI_ERROR_WAIT_TIMEOUT, DXGI_OUTDUPL_FRAME_INFO, DXGI_OUTPUT_DESC, DXGI_RESOURCE_PRIORITY_MAXIMUM
            },
            Gdi::HMONITOR,
        },
    },
};


pub struct FrameAnalysisResult {
    pub is_black: bool,
    pub non_black_percentage: f32,
    pub max_brightness: u8,
    pub sample_pixels: Vec<(u32, u32, [u8; 3])>, // Coordinates and RGB values
    pub analyzed_pixels: u32,
}

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
    _d3d_device: ID3D11Device,
    duplication: IDXGIOutputDuplication,
    buffer_texture: Option<ID3D11Texture2D>,
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
        //println!("Attempting to acquire next frame with timeout: {} ms", timeout_ms); // Keep if needed
    
        let mut frame_info: DXGI_OUTDUPL_FRAME_INFO = Default::default();
        let mut desktop_resource: Option<IDXGIResource> = None;
    
        //println!("Calling AcquireNextFrame..."); // Keep if needed
        let acquire_result = unsafe {
            self.duplication
                .AcquireNextFrame(timeout_ms, &mut frame_info, &mut desktop_resource)
        };
    
        match acquire_result {
            Ok(_) => {
                //println!("AcquireNextFrame succeeded, frame info: {:?}", frame_info); // Keep if needed
    
                let desktop_resource = desktop_resource
                    .expect("AcquireNextFrame succeeded but returned null resource");
                //println!("Got desktop resource, casting to ID3D11Texture2D"); // Keep if needed
    
                // --- Start Change ---
                let acquired_texture: ID3D11Texture2D = desktop_resource.cast()?;
    
                // Ensure buffer_texture exists and matches description
                let mut desc = D3D11_TEXTURE2D_DESC::default();
                unsafe { acquired_texture.GetDesc(&mut desc) };

                let acquired_texture: ID3D11Texture2D = desktop_resource.cast()?;
                /*println!("Analyzing acquired_texture BEFORE copy...");
                match self.analyze_frame(&acquired_texture) { // Analyze the ORIGINAL
                    Ok(analysis) => println!(" -> BEFORE analysis result: black={}, non_black%={:.2}, max_brightness={}", analysis.is_black, analysis.non_black_percentage, analysis.max_brightness),
                    Err(e) => println!(" -> Error analyzing BEFORE copy: {:?}", e),
                }*/
    
                // We need a texture we can copy *to*, it doesn't need CPU access here
                // It should be usable as a source later, so default usage is fine.
                let buffer_desc = D3D11_TEXTURE2D_DESC {
                    Width: desc.Width,
                    Height: desc.Height,
                    MipLevels: 1,        // Must be 1 for CopyResource
                    ArraySize: 1,        // Must be 1 for DuplicateOutput textures
                    Format: desc.Format, // MUST match the source format
                    SampleDesc: desc.SampleDesc, // MUST match the source sample desc
                    Usage: D3D11_USAGE_DEFAULT, // Suitable for GPU-GPU copies
                    BindFlags: 0,        // NO specific bind flags needed
                    CPUAccessFlags: 0,   // NO CPU access needed
                    MiscFlags: 0,        // NO misc flags needed (clear potentially problematic ones like SHARED)
                    ..Default::default() // Ensure other fields are zeroed
                };
    
                // Check if buffer needs creation or recreation
                let mut create_new_buffer = true;
                if let Some(buffer) = &self.buffer_texture {
                    let mut buffer_current_desc = D3D11_TEXTURE2D_DESC::default();
                    unsafe { buffer.GetDesc(&mut buffer_current_desc) };
                    if buffer_current_desc.Width == buffer_desc.Width &&
                       buffer_current_desc.Height == buffer_desc.Height &&
                       buffer_current_desc.Format == buffer_desc.Format {
                        create_new_buffer = false;
                    }
                }
    
                if create_new_buffer {
                    println!("Creating buffer texture for captured frame copy.");
                    self.buffer_texture = Some(unsafe {
                         let mut texture = None;
                         self._d3d_device.CreateTexture2D(&buffer_desc, None, Some(&mut texture))?;
                         texture.unwrap()
                    });
                }
    
                let target_texture = self.buffer_texture.as_ref().unwrap();
    
                // Get context and copy
                let context = unsafe { self._d3d_device.GetImmediateContext()? };
                unsafe { context.CopyResource(target_texture, &acquired_texture) };
    
                // NOW release the frame
                //println!("Releasing frame"); // Keep if needed
                unsafe { self.duplication.ReleaseFrame()? };
    
                // Return the *cloned handle* to our owned buffer texture
                // Cloning the COM pointer just increases the ref count
                let texture_copy = target_texture.clone();
                println!("test");
                // Frame analysis (optional, can be done on texture_copy or target_texture)
                /*match self.analyze_frame(&texture_copy) { // Analyze the copy
                    Ok(analysis) => {
                        if analysis.is_black {
                             println!("Frame analysis: FRAME IS BLACK (Black Pixels: {:.2}%, Max Brightness: {})", analysis.non_black_percentage, analysis.max_brightness);
                        } else {
                             println!("Frame analysis: Non-Black Pixels: {:.2}%, Max Brightness: {}", analysis.non_black_percentage, analysis.max_brightness);
                            // if !analysis.sample_pixels.is_empty() {
                            //     println!("Sample non-black pixels:");
                            //     for (i, (x, y, rgb)) in analysis.sample_pixels.iter().enumerate().take(2) { // Limit logging
                            //         println!("  Pixel {}: ({}, {}) RGB={:?}", i, x, y, rgb);
                            //     }
                            // }
                        }
                    },
                    Err(e) => println!("Error analyzing frame: {:?}", e),
                }*/
    
                //println!("Frame acquisition complete, returning copied texture."); // Keep if needed
                Ok(Some(AcquiredFrame {
                    texture: texture_copy, // Return the copy
                    frame_info,
                }))
                // --- End Change ---
    
            }
            Err(err) if err.code() == DXGI_ERROR_WAIT_TIMEOUT => {
                //println!("Frame acquisition timed out after {} ms", timeout_ms); // Keep if needed
                Ok(None) // Timeout is not an error condition here
            }
            Err(err) if err.code() == DXGI_ERROR_ACCESS_LOST => {
                println!("ERROR: Access lost to desktop duplication: {:?}", err);
                Err(err.into())
            }
            Err(err) => {
                println!("ERROR: Failed to acquire frame: {:?}", err);
                Err(err.into()) // Other errors
            }
        }
    }
    
    pub fn analyze_frame(&self, texture: &ID3D11Texture2D) -> Result<FrameAnalysisResult> {
        // Get the D3D11 device and immediate context
        let device = &self._d3d_device;
        let context = unsafe { device.GetImmediateContext() }.unwrap();
    
        // Create a staging texture that we can read from CPU
        let mut texture_desc = unsafe { std::mem::zeroed::<D3D11_TEXTURE2D_DESC>() };
        unsafe { texture.GetDesc(&mut texture_desc) };
        
        // Modify the description for our staging texture
        texture_desc.Usage = D3D11_USAGE_STAGING;
        texture_desc.BindFlags = 0;
        texture_desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ.0 as u32;
        texture_desc.MiscFlags = 0;
    
        // Create the staging texture
        let staging_texture = unsafe {
            let mut texture_out = None;
            let result = device.CreateTexture2D(
                &texture_desc, 
                None, 
                Some(&mut texture_out)
            );
            
            if result.is_err() {
                println!("Failed to create staging texture: {:?}", result);
                return Err(result.unwrap_err());
            }
            
            texture_out.unwrap()
        };
    
        // Copy the frame texture to our staging texture
        unsafe { context.CopyResource(&staging_texture, texture) };
    
        // Map the staging texture to get access to its data
        let mut mapped_resource = unsafe { std::mem::zeroed::<D3D11_MAPPED_SUBRESOURCE>() };
        let map_result = unsafe {
            context.Map(
                &staging_texture,
                0, // Subresource index
                D3D11_MAP_READ,
                0, // MapFlags
                Some(&mut mapped_resource),
            )
        };
        
        if map_result.is_err() {
            println!("Failed to map texture: {:?}", map_result);
            return Err(map_result.unwrap_err());
        }
    
        // Analyze the pixel data
        let analysis = unsafe {
            let row_pitch = mapped_resource.RowPitch;
            let data_ptr = mapped_resource.pData as *const u8;
            
            let mut is_black = true;
            let mut non_black_count = 0;
            let mut max_brightness = 0u8;
            let mut sample_pixels = Vec::new();
            let sample_step = 16; // Check every Nth pixel for performance
            let mut analyzed_pixels = 0;
            
            for y in 0..texture_desc.Height {
                if y % sample_step != 0 {
                    continue;
                }
                
                let row_start = data_ptr.add((y * row_pitch) as usize);
                
                for x in 0..texture_desc.Width {
                    if x % sample_step != 0 {
                        continue;
                    }
                    
                    analyzed_pixels += 1;
                    let pixel_offset = (x * 4) as usize; // 4 bytes per pixel
                    let pixel = row_start.add(pixel_offset);
                    
                    // BGRA format - B, G, R values at offsets 0, 1, 2
                    let b = *pixel;
                    let g = *pixel.add(1);
                    let r = *pixel.add(2);
                    
                    // Calculate max brightness across RGB channels
                    let brightness = b.max(g).max(r);
                    max_brightness = max_brightness.max(brightness);
                    
                    // Check if pixel is non-black (allowing for some near-black noise)
                    if brightness > 5 { // Threshold for "black enough"
                        is_black = false;
                        non_black_count += 1;
                        
                        // Store some samples for debugging (up to 10)
                        if sample_pixels.len() < 10 {
                            sample_pixels.push((x, y, [r, g, b]));
                        }
                    }
                }
            }
            
            // Calculate non-black percentage based on our sampling
            let non_black_percentage = if analyzed_pixels > 0 {
                (non_black_count as f32 / analyzed_pixels as f32) * 100.0
            } else {
                0.0
            };
            
            // Unmap when done
            context.Unmap(&staging_texture, 0);
            
            FrameAnalysisResult {
                is_black,
                non_black_percentage,
                max_brightness,
                sample_pixels,
                analyzed_pixels,
            }
        };
    
        Ok(analysis)
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
