use std::{sync::Arc, time::{SystemTime, UNIX_EPOCH}};

use windows::{
    core::{Result, HSTRING},
    Foundation::TimeSpan,
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Graphics::{
            Direct3D11::{
                ID3D11Device, ID3D11DeviceContext, ID3D11RenderTargetView, ID3D11Texture2D, 
                D3D11_BIND_RENDER_TARGET, D3D11_BIND_SHADER_RESOURCE, D3D11_BOX, 
                D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT
            },
            Dxgi::Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_NV12, DXGI_SAMPLE_DESC},
            Gdi::HMONITOR,
        },
        Media::MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFCreateAttributes, 
            MFCreateMFByteStreamOnStreamEx, MFCreateSinkWriterFromURL
        },
        System::Performance::{QueryPerformanceCounter, QueryPerformanceFrequency}
    },
};

use crate::capture::{AcquiredFrame, CaptureFrameGenerator, CustomGraphicsCaptureSession};

use super::{
    encoder::{VideoEncoder, VideoEncoderInputSample},
    encoder_device::VideoEncoderDevice,
    processor::VideoProcessor,
};

pub struct VideoEncodingSession {
    video_encoder: VideoEncoder,
    capture_session: CustomGraphicsCaptureSession,
    sample_writer: Arc<SampleWriter>,
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

    frame_period: i64,
    next_frame_time: TimeSpan,
    qpc_frequency: i64,
}

pub struct SampleWriter {
    _stream: IRandomAccessStream,
    sink_writer: IMFSinkWriter,
    sink_writer_stream_index: u32,
}

impl VideoEncodingSession {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        encoder_device: &VideoEncoderDevice,
        resolution: SizeInt32,
        bit_rate: u32,
        frame_rate: u32,
        stream: IRandomAccessStream,
    ) -> Result<Self> {
        let input_size = ensure_even_size(resolution);
        let output_size = ensure_even_size(resolution);

        let mut video_encoder = VideoEncoder::new(
            encoder_device,
            d3d_device.clone(),
            output_size,
            output_size,
            bit_rate,
            frame_rate,
        )?;
        let output_type = video_encoder.output_type().clone();

        let mut sample_generator = SampleGenerator::new(
            d3d_device, 
            monitor_handle,
            input_size, 
            output_size,
            frame_rate,
        )?;
        let capture_session = sample_generator.capture_session().clone();
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
            capture_session,
            sample_writer,
        })
    }

    pub fn start(&mut self) -> Result<()> {
        self.sample_writer.start()?;
        self.capture_session.StartCapture()?;
        assert!(self.video_encoder.try_start()?);
        Ok(())
    }

    pub fn stop(&mut self) -> Result<()> {
        self.video_encoder.stop()?;
        self.sample_writer.stop()?;
        Ok(())
    }
}

unsafe impl Send for SampleGenerator {}
impl SampleGenerator {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        input_size: SizeInt32,
        output_size: SizeInt32,
        frame_rate: u32,
    ) -> Result<Self> {
        let d3d_context = unsafe { d3d_device.GetImmediateContext()? };

        let video_processor = VideoProcessor::new(
            d3d_device.clone(),
            DXGI_FORMAT_B8G8R8A8_UNORM,
            input_size,
            DXGI_FORMAT_NV12,
            output_size,
        )?;

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

        let mut qpc_frequency: i64 = 0;
        unsafe {
            QueryPerformanceFrequency(&mut qpc_frequency)?;
        }
        
        // Calculate frame period in QPC units (performance counter ticks)
        let frame_period = qpc_frequency / (frame_rate as i64);
        
        let compose_texture = unsafe {
            let mut texture = None;
            d3d_device.CreateTexture2D(&texture_desc, None, Some(&mut texture))?;
            texture.unwrap()
        };
        
        let render_target_view = unsafe {
            let mut rtv = None;
            d3d_device.CreateRenderTargetView(&compose_texture, None, Some(&mut rtv))?;
            rtv.unwrap()
        };

        // Create frame generator
        let frame_generator = CaptureFrameGenerator::new(d3d_device.clone(), monitor_handle)?;

        Ok(Self {
            d3d_device,
            d3d_context,

            video_processor,
            compose_texture,
            render_target_view,

            frame_generator,

            seen_first_time_stamp: false,
            first_timestamp: TimeSpan::default(),

            frame_period,
            next_frame_time: TimeSpan::default(),
            qpc_frequency,
        })
    }

    pub fn capture_session(&self) -> &CustomGraphicsCaptureSession {
        self.frame_generator.session()
    }

    pub fn generate(&mut self) -> Result<Option<VideoEncoderInputSample>> {
        // Loop until we find a frame that meets our timing requirements or get None
        loop {
            match self.frame_generator.try_get_next_frame()? {
                Some(frame) => {
                    // If this is our first frame, initialize timing
                    if !self.seen_first_time_stamp {
                        self.first_timestamp = frame.present_time;
                        self.seen_first_time_stamp = true;
                        // Set the next expected frame time
                        self.next_frame_time = TimeSpan {
                            Duration: self.first_timestamp.Duration + self.frame_period,
                        };
                        
                        //println!("First frame at QPC time: {}, frame_period: {}", 
                        //         self.first_timestamp.Duration, self.frame_period);
                        
                        // Process this first frame
                        return match self.generate_from_frame(&frame) {
                            Ok(sample) => Ok(Some(sample)),
                            Err(err) => Err(err),
                        };
                    }
                    
                    // Calculate relative time from first frame
                    let relative_time = TimeSpan {
                        Duration: frame.present_time.Duration - self.first_timestamp.Duration,
                    };
                    
                    let expected_time = self.next_frame_time.Duration - self.first_timestamp.Duration;
                    
                    // Check if this frame is at or after our expected time
                    if relative_time.Duration >= expected_time {
                        // Calculate timing precision metrics
                        let time_delta = relative_time.Duration - expected_time;
                        let time_error_ms = (time_delta as f64 * 1000.0) / self.qpc_frequency as f64;
                        
                        // Create a timestamp for the actual fps calculation
                        static mut LAST_FRAME_TIME: i64 = 0;
                        static mut FRAME_COUNT: u32 = 0;
                        static mut TOTAL_TIME_ERROR_MS: f64 = 0.0;
                        
                        unsafe {
                            // Calculate actual frame interval
                            if LAST_FRAME_TIME > 0 {
                                let actual_interval = frame.present_time.Duration - LAST_FRAME_TIME;
                                let actual_frame_time_ms = (actual_interval as f64 * 1000.0) / self.qpc_frequency as f64;
                                let target_frame_time_ms = (self.frame_period as f64 * 1000.0) / self.qpc_frequency as f64;
                                
                                // Update running statistics
                                FRAME_COUNT += 1;
                                TOTAL_TIME_ERROR_MS += time_error_ms;
                                let avg_error_ms = TOTAL_TIME_ERROR_MS / FRAME_COUNT as f64;
                                
                                if FRAME_COUNT % 30 == 0 {  // Print detailed stats every 30 frames
                                    println!("FRAME STATS: Frame #{}", FRAME_COUNT);
                                    println!("  Target frame time: {:.2} ms", target_frame_time_ms);
                                    println!("  Actual frame time: {:.2} ms", actual_frame_time_ms);
                                    println!("  Current error: {:.2} ms", time_error_ms);
                                    println!("  Average error: {:.2} ms", avg_error_ms);
                                    println!("  Target FPS: {}", 1000.0 / target_frame_time_ms);
                                    println!("  Actual FPS: {:.2}", 1000.0 / actual_frame_time_ms);
                                } else {
                                    /*println!("Frame #{}: error {:.2}ms, actual frame time: {:.2}ms", 
                                            FRAME_COUNT, time_error_ms, actual_frame_time_ms);*/
                                }
                            }
                            LAST_FRAME_TIME = frame.present_time.Duration;
                        }
                        
                        // Update the next expected frame time
                        self.next_frame_time = TimeSpan {
                            Duration: self.next_frame_time.Duration + self.frame_period,
                        };
                        
                        // Fix type mismatch by wrapping in Some
                        return match self.generate_from_frame(&frame) {
                            Ok(sample) => Ok(Some(sample)),
                            Err(err) => Err(err),
                        };
                    } else {
                        // This frame is too early, skip it
                        /*println!("Skipping frame at time: {}, waiting for: {}", 
                                relative_time.Duration, expected_time);*/
                        
                        // Continue loop to get next frame
                        continue;
                    }
                },
                None => {
                    // No more frames, end capture
                    self.stop_capture()?;
                    return Ok(None);
                }
            }
        }
    }

    fn stop_capture(&mut self) -> Result<()> {
        self.frame_generator.stop_capture()
    }
    
    fn generate_from_frame(
        &mut self,
        frame: &AcquiredFrame,
    ) -> Result<VideoEncoderInputSample> {
        let frame_texture = &frame.texture;
        let frame_time = frame.present_time;

        if !self.seen_first_time_stamp {
            self.first_timestamp = frame_time;
            self.seen_first_time_stamp = true;
        }

        let timestamp = TimeSpan {
            Duration: frame_time.Duration - self.first_timestamp.Duration,
        };
    
        // Determine region to copy
        let desc = unsafe {
            let mut desc = D3D11_TEXTURE2D_DESC::default();
            frame_texture.GetDesc(&mut desc);
            desc
        };
        let region = D3D11_BOX {
            left: 0, right: desc.Width, top: 0, bottom: desc.Height, back: 1, front: 0,
        };
    
        // GPU Processing
        unsafe {
            // Clear render target
            self.d3d_context.ClearRenderTargetView(&self.render_target_view, &CLEAR_COLOR);
    
            // Copy the captured frame to composition texture
            self.d3d_context.CopySubresourceRegion(
                &self.compose_texture,
                0, 0, 0, 0,
                frame_texture,
                0, Some(&region),
            );
    
            // Process BGRA -> NV12
            // Fix: Call the function directly and use ? afterward
            self.video_processor.process_texture(&self.compose_texture)?;
    
            // Get the resulting NV12 texture
            let video_output_texture = self.video_processor.output_texture();
    
            // Create a new texture for the sample
            let sample_texture = {
                 let output_desc = {
                     let mut desc = D3D11_TEXTURE2D_DESC::default();
                     video_output_texture.GetDesc(&mut desc);
                     desc
                 };
                 
                 // Fix: Call the function directly and use ? afterward
                 let mut texture = None;
                 self.d3d_device.CreateTexture2D(&output_desc, None, Some(&mut texture))?;
                 texture.unwrap()
            };
            
            // Copy the processed texture to the sample texture
            self.d3d_context.CopyResource(&sample_texture, video_output_texture);
    
            // Create and return the input sample
            Ok(VideoEncoderInputSample::new(
                timestamp,
                sample_texture,
            ))
        }
    }
}

unsafe impl Send for SampleWriter {}
unsafe impl Sync for SampleWriter {}
impl SampleWriter {
    pub fn new(
        stream: IRandomAccessStream,
        output_type: &IMFMediaType,
    ) -> Result<Self> {
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
        // Get the sample time directly
        unsafe {
            let time = sample.GetSampleTime()?;
            //println!("Sample time: {}", time);
            
            // Write the sample to the sink
            self.sink_writer
                .WriteSample(self.sink_writer_stream_index, sample)
        }
    }
}


const CLEAR_COLOR: [f32; 4] = [0.0, 0.0, 0.0, 1.0];

fn ensure_even(value: i32) -> i32 {
    if value % 2 == 0 {
        value
    } else {
        value + 1
    }
}

fn ensure_even_size(size: SizeInt32) -> SizeInt32 {
    SizeInt32 {
        Width: ensure_even(size.Width),
        Height: ensure_even(size.Height),
    }
}