use std::sync::Arc;

use windows::{
    core::{ComInterface, Result, HSTRING}, // Ensure ComInterface is imported for .cast()
    Foundation::TimeSpan,
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Graphics::{
            Direct3D11::{
                ID3D11Device, ID3D11DeviceContext, ID3D11RenderTargetView, ID3D11Resource, ID3D11Texture2D, D3D11_BIND_RENDER_TARGET, D3D11_BIND_SHADER_RESOURCE, D3D11_BOX, D3D11_CPU_ACCESS_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_MAP_READ, D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT, D3D11_USAGE_STAGING
            },
            Dxgi::{
                Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_NV12, DXGI_SAMPLE_DESC},
                DXGI_ERROR_ACCESS_LOST, // Ensure this is imported
            },
            Gdi::HMONITOR,
        },
        Media::MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFCreateAttributes,
            MFCreateMFByteStreamOnStreamEx, MFCreateSinkWriterFromURL,
        },
        System::Performance::QueryPerformanceFrequency, // Correct path
    },
};

use crate::{
    capture::{AcquiredFrame, CaptureFrameGenerator},
    video::{
        encoding_session::{VideoEncoderSessionFactory, VideoEncodingSession},
        util::ensure_even_size,
        CLEAR_COLOR,
    },
};

use super::{
    encoder::{VideoEncoder, VideoEncoderInputSample},
    encoder_device::VideoEncoderDevice,
    processor::VideoProcessor,
};

struct MFVideoEncodingSession {
    video_encoder: VideoEncoder,
    sample_writer: Arc<SampleWriter>,
}

// Helper function moved before SampleGenerator impl
fn convert_qpc_to_timespan(qpc_time: i64) -> Result<TimeSpan> {
    if qpc_time == 0 {
        return Ok(TimeSpan::default());
    }
    let mut frequency = 0;
    unsafe {
        QueryPerformanceFrequency(&mut frequency)?;
    }
    if frequency == 0 {
        return Err(windows::core::Error::new(
            windows::Win32::Foundation::E_FAIL,
            "QueryPerformanceFrequency returned zero".into(),
        ));
    }
    let duration = (qpc_time as i128 * 10_000_000) / (frequency as i128);
    Ok(TimeSpan {
        Duration: duration as i64,
    })
}

struct SampleGenerator {
    d3d_device: ID3D11Device,
    d3d_context: ID3D11DeviceContext,

    video_processor: VideoProcessor,
    compose_texture: ID3D11Texture2D,
    render_target_view: ID3D11RenderTargetView,

    frame_generator: CaptureFrameGenerator,

    seen_first_time_stamp: bool,
    first_timestamp: TimeSpan,
}

struct SampleWriter {
    _stream: IRandomAccessStream,
    sink_writer: IMFSinkWriter,
    sink_writer_stream_index: u32,
}

impl MFVideoEncodingSession {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        encoder_device: &VideoEncoderDevice,
        resolution: SizeInt32,
        bit_rate: u32,
        frame_rate: u32,
        stream: IRandomAccessStream,
    ) -> Result<Self> {
        println!("Starting MFVideoEncodingSession::new");
        let input_size = ensure_even_size(resolution);
        let output_size = ensure_even_size(resolution);
        println!("Input size: {}x{}, Output size: {}x{}", 
                 input_size.Width, input_size.Height, 
                 output_size.Width, output_size.Height);
    
        println!("Creating VideoEncoder");
        let mut video_encoder = match VideoEncoder::new(
            encoder_device,
            d3d_device.clone(),
            output_size,
            output_size,
            bit_rate,
            frame_rate,
        ) {
            Ok(encoder) => {
                println!("VideoEncoder created successfully");
                encoder
            },
            Err(e) => {
                println!("Failed to create VideoEncoder: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
        
        let output_type = video_encoder.output_type().clone();
    
        println!("Creating SampleGenerator");
        let mut sample_generator = match SampleGenerator::new(
            d3d_device.clone(),
            monitor_handle,
            input_size,
            output_size,
        ) {
            Ok(generator) => {
                println!("SampleGenerator created successfully");
                generator
            },
            Err(e) => {
                println!("Failed to create SampleGenerator: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };

        video_encoder.set_sample_requested_callback(
            move || -> Result<Option<VideoEncoderInputSample>> { sample_generator.generate() },
        );

        let sample_writer = Arc::new(SampleWriter::new(stream, &output_type)?);
        video_encoder.set_sample_rendered_callback({
            let sample_writer = sample_writer.clone();
            move |sample| -> Result<()> { sample_writer.write(sample.sample()) }
        });

        Ok(Self {
            video_encoder,
            sample_writer,
        })
    }
}

impl VideoEncodingSession for MFVideoEncodingSession {
    fn start(&mut self) -> Result<()> {
        self.sample_writer.start()?;
        assert!(self.video_encoder.try_start()?);
        Ok(())
    }

    fn stop(&mut self) -> Result<()> {
        self.video_encoder.stop()?;
        self.sample_writer.stop()?;
        Ok(())
    }
}

pub struct MFVideoEncodingSessionFactory {
    encoder_device: VideoEncoderDevice,
}

impl MFVideoEncodingSessionFactory {
    pub fn new(encoder_device: VideoEncoderDevice) -> Self {
        Self { encoder_device }
    }
}

impl VideoEncoderSessionFactory for MFVideoEncodingSessionFactory {
    fn create_session(
        &self,
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        resolution: SizeInt32,
        bit_rate: u32,
        frame_rate: u32,
        stream: IRandomAccessStream,
    ) -> Result<Box<dyn VideoEncodingSession>> {
        let session = Box::new(MFVideoEncodingSession::new(
            d3d_device,
            monitor_handle,
            &self.encoder_device,
            resolution,
            bit_rate,
            frame_rate,
            stream,
        )?);
        Ok(session)
    }
}

unsafe impl Send for SampleGenerator {}
impl SampleGenerator {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        input_size: SizeInt32,
        output_size: SizeInt32,
    ) -> Result<Self> {
        println!("SampleGenerator::new starting...");
        println!("Input size: {}x{}, Output size: {}x{}", 
            input_size.Width, input_size.Height, 
            output_size.Width, output_size.Height);
        
        println!("Getting D3D context...");
        let d3d_context = match unsafe { d3d_device.GetImmediateContext() } {
            Ok(context) => {
                println!("Successfully got D3D context");
                context
            },
            Err(e) => {
                println!("Failed to get D3D context: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
    
        println!("Creating VideoProcessor...");
        let video_processor = match VideoProcessor::new(
            d3d_device.clone(),
            DXGI_FORMAT_B8G8R8A8_UNORM,
            input_size,
            DXGI_FORMAT_NV12,
            output_size,
        ) {
            Ok(processor) => {
                println!("Successfully created VideoProcessor");
                processor
            },
            Err(e) => {
                println!("Failed to create VideoProcessor: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
    
        println!("Creating compose texture...");
        let texture_desc = D3D11_TEXTURE2D_DESC {
            Width: input_size.Width as u32,
            Height: input_size.Height as u32,
            ArraySize: 1,
            MipLevels: 1,
            Format: DXGI_FORMAT_B8G8R8A8_UNORM,
            SampleDesc: DXGI_SAMPLE_DESC {
                Count: 1,
                ..Default::default()
            },
            Usage: D3D11_USAGE_DEFAULT,
            BindFlags: (D3D11_BIND_RENDER_TARGET.0 | D3D11_BIND_SHADER_RESOURCE.0) as u32,
            ..Default::default()
        };
        
        println!("Texture description: Width={}, Height={}, Format={:?}, BindFlags={:x}",
            texture_desc.Width, texture_desc.Height, texture_desc.Format, texture_desc.BindFlags);
        
        let compose_texture = match unsafe {
            let mut texture = None;
            let result = d3d_device.CreateTexture2D(&texture_desc, None, Some(&mut texture));
            if result.is_ok() {
                println!("Successfully created texture");
                Ok(texture.unwrap())
            } else {
                println!("Failed to create texture: {:?}", result);
                Err(result.unwrap_err())
            }
        } {
            Ok(texture) => texture,
            Err(e) => {
                println!("Error creating compose texture: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
    
        println!("Creating render target view...");
        let render_target_view = match unsafe {
            let mut rtv = None;
            let result = d3d_device.CreateRenderTargetView(&compose_texture, None, Some(&mut rtv));
            if result.is_ok() {
                println!("Successfully created render target view");
                Ok(rtv.unwrap())
            } else {
                println!("Failed to create render target view: {:?}", result);
                Err(result.unwrap_err())
            }
        } {
            Ok(view) => view,
            Err(e) => {
                println!("Error creating render target view: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
    
        println!("Creating CaptureFrameGenerator...");
        let frame_generator = match CaptureFrameGenerator::new(d3d_device.clone(), monitor_handle) {
            Ok(generator) => {
                println!("Successfully created CaptureFrameGenerator");
                generator
            },
            Err(e) => {
                println!("Failed to create CaptureFrameGenerator: {:?} - {}", e.code(), e.message());
                return Err(e);
            }
        };
        
        let (gen_width, gen_height) = frame_generator.resolution();
        println!("Frame generator resolution: {}x{}", gen_width, gen_height);
    
        if input_size.Width as u32 != gen_width || input_size.Height as u32 != gen_height {
            eprintln!(
                "Warning: Specified input size ({}, {}) does not match monitor resolution ({}, {}). Using monitor resolution.",
                input_size.Width, input_size.Height, gen_width, gen_height
            );
        }
    
        println!("SampleGenerator::new completed successfully");
        Ok(Self {
            d3d_device,
            d3d_context,
            video_processor,
            compose_texture,
            render_target_view,
            frame_generator,
            seen_first_time_stamp: false,
            first_timestamp: TimeSpan::default(),
        })
    }

    pub fn generate(&mut self) -> Result<Option<VideoEncoderInputSample>> {
        match self.frame_generator.try_get_next_frame(100) {
            Ok(Some(frame)) => {
                match self.generate_from_frame(&frame) {
                    Ok(sample) => Ok(Some(sample)),
                    Err(error) => {
                        eprintln!(
                            "Error during input sample generation: {:?} - {}",
                            error.code(),
                            error.message()
                        );
                        Ok(None)
                    }
                }
            }
            Ok(None) => {
                Ok(None) // Timeout
            }
            Err(e) if e.code() == DXGI_ERROR_ACCESS_LOST => {
                eprintln!("DXGI Access Lost in frame generation: {:?}", e);
                Ok(None)
            }
            Err(e) => {
                eprintln!("Error getting next frame: {:?}", e);
                Ok(None)
            }
        }
    }

    fn generate_from_frame(
        &mut self,
        frame: &AcquiredFrame,
    ) -> Result<VideoEncoderInputSample> {
        let frame_qpc_time = frame.frame_info.LastPresentTime;
        let timestamp = convert_qpc_to_timespan(frame_qpc_time)?;
    
        if !self.seen_first_time_stamp {
            self.first_timestamp = timestamp;
            self.seen_first_time_stamp = true;
        }
        let relative_timestamp = TimeSpan {
            Duration: timestamp.Duration - self.first_timestamp.Duration,
        };
    
        let frame_texture = &frame.texture;
        let desc = unsafe {
            let mut desc = D3D11_TEXTURE2D_DESC::default();
            frame_texture.GetDesc(&mut desc);
            desc
        };
    
        let width = desc.Width;
        let height = desc.Height;
    
        let region = D3D11_BOX {
            left: 0,
            right: width,
            top: 0,
            bottom: height,
            back: 1,
            front: 0,
        };
    
        unsafe {
            self.d3d_context
                .ClearRenderTargetView(&self.render_target_view, &CLEAR_COLOR);
            self.d3d_context.CopySubresourceRegion(
                &self.compose_texture,
                0,
                0,
                0,
                0,
                &*frame_texture,
                0,
                Some(&region),
            );
    
            // Process our back buffer
            self.video_processor
                .process_texture(&self.compose_texture)?;
    
            // Get our NV12 texture
            let video_output_texture = self.video_processor.output_texture();
    
            // Make a copy for the sample
            let desc = {
                let mut desc = D3D11_TEXTURE2D_DESC::default();
                video_output_texture.GetDesc(&mut desc);
                desc
            };
            let sample_texture = {
                let mut texture = None;
                self.d3d_device
                    .CreateTexture2D(&desc, None, Some(&mut texture))?;
                texture.unwrap()
            };
            self.d3d_context
                .CopyResource(&sample_texture, video_output_texture);
    
            Ok(VideoEncoderInputSample::new(relative_timestamp, sample_texture))
        }
    }
}

unsafe impl Send for SampleWriter {}
unsafe impl Sync for SampleWriter {}
impl SampleWriter {
    pub fn new(stream: IRandomAccessStream, output_type: &IMFMediaType) -> Result<Self> {
        let empty_attributes = unsafe {
            let mut attributes = None;
            MFCreateAttributes(&mut attributes, 0)?;
            attributes.unwrap()
        };
        let sink_writer = unsafe {
            let byte_stream = MFCreateMFByteStreamOnStreamEx(&stream)?;
            MFCreateSinkWriterFromURL(&HSTRING::from(".mp4"), &byte_stream, &empty_attributes)?
        };
        let sink_writer_stream_index = unsafe { sink_writer.AddStream(output_type)? };
        unsafe {
            sink_writer.SetInputMediaType(
                sink_writer_stream_index,
                output_type,
                &empty_attributes,
            )?
        };

        Ok(Self {
            _stream: stream,
            sink_writer,
            sink_writer_stream_index,
        })
    }

    pub fn start(&self) -> Result<()> {
        unsafe { self.sink_writer.BeginWriting() }
    }

    pub fn stop(&self) -> Result<()> {
        unsafe { self.sink_writer.Finalize() }
    }

    pub fn write(&self, sample: &IMFSample) -> Result<()> {
        unsafe {
            self.sink_writer
                .WriteSample(self.sink_writer_stream_index, sample)
        }
    }
}
