use std::{sync::{mpsc::{sync_channel, Receiver, RecvTimeoutError, SyncSender, TryRecvError}, Arc}, time::Duration};

use windows::{
    core::{ComInterface, Result, HSTRING}, // Ensure ComInterface is imported for .cast()
    Foundation::TimeSpan,
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Foundation::E_FAIL, Graphics::{
            Direct3D11::{
                ID3D11Device, ID3D11DeviceContext, ID3D11RenderTargetView, ID3D11Resource, ID3D11Texture2D, D3D11_BIND_RENDER_TARGET, D3D11_BIND_SHADER_RESOURCE, D3D11_BOX, D3D11_CPU_ACCESS_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_MAP_READ, D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT, D3D11_USAGE_STAGING
            },
            Dxgi::{
                Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_NV12, DXGI_SAMPLE_DESC},
                DXGI_ERROR_ACCESS_LOST, // Ensure this is imported
            },
            Gdi::HMONITOR,
        }, Media::MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFAudioFormat_AAC, MFAudioFormat_PCM, MFCreateAttributes, MFCreateMFByteStreamOnStreamEx, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFCreateSinkWriterFromURL, MFMediaType_Audio, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE
        }, System::Performance::QueryPerformanceFrequency // Correct path
    },
};

use crate::{
    audio::{create_imf_sample_from_packet, AudioCapture, AudioDataPacket, AudioSource}, capture::{AcquiredFrame, CaptureFrameGenerator}, video::{
        encoding_session::{VideoEncoderSessionFactory, VideoEncodingSession}, mf::audio_encoder::{AudioEncoder, EncodedAudioPacket}, util::ensure_even_size, CLEAR_COLOR
    }
};

use super::{
    encoder::{VideoEncoder, VideoEncoderInputSample},
    encoder_device::VideoEncoderDevice,
    processor::VideoProcessor,
};

struct MFVideoEncodingSession {
    video_encoder: VideoEncoder,
    sample_writer: Arc<SampleWriter>,
    audio_capture: Option<AudioCapture>,
    raw_audio_sender: Option<SyncSender<(u32, AudioDataPacket)>>,
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

    target_frame_duration: TimeSpan, // Calculate this in new(): 1_000_000_0 / frame_rate
    next_target_relative_timestamp: Option<TimeSpan>,
    last_returned_relative_timestamp: Option<TimeSpan>,
    seen_first_time_stamp: bool,
    first_timestamp: TimeSpan,
}


pub struct SampleWriter {
    _stream: IRandomAccessStream,
    sink_writer: IMFSinkWriter,
    video_stream_index: u32,
    audio_stream_index: Option<u32>,
    audio_packet_receiver: Receiver<(u32, AudioDataPacket)>,
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
        audio_source: AudioSource,
    ) -> Result<Self> {
        println!(
            "Starting MFVideoEncodingSession::new with AudioSource: {:?}",
            audio_source
        );
        let input_size = ensure_even_size(resolution);
        let output_size = ensure_even_size(resolution);
        println!(
            "Input size: {}x{}, Output size: {}x{}",
            input_size.Width, input_size.Height, output_size.Width, output_size.Height
        );
    
        // --- Video Encoder Setup (No change) ---
        println!("Creating VideoEncoder");
        let mut video_encoder = VideoEncoder::new(
            encoder_device,
            d3d_device.clone(),
            output_size,
            output_size,
            bit_rate,
            frame_rate,
        )?;
        let video_output_type = video_encoder.output_type().clone();
        println!("VideoEncoder created successfully");
    
        // --- Audio Capture Setup ---
        let placeholder_index: u32 = 0;
        println!("Attempting to create AudioCapture...");
        let mut audio_capture = match AudioCapture::new(audio_source, placeholder_index) {
            Ok(ac) => {
                println!("AudioCapture created successfully.");
                Some(ac)
            }
            Err(e) => {
                eprintln!("Failed to create AudioCapture: {:?}. Audio disabled.", e);
                None
            }
        };
    
        // --- Channel for RAW Audio Packets (Capture -> SinkWriter) ---
        let (raw_audio_sender, raw_audio_receiver) =
            sync_channel::<(u32, AudioDataPacket)>(100);
    
        // --- Create Audio Output Media Type (AAC) ---
        let audio_output_type = if let Some(ref ac) = audio_capture {
            match unsafe {
                // Create a standard AAC media type for the output
                let aac_type = MFCreateMediaType()?;
                aac_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
                aac_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
                aac_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?;
                aac_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, 44100)?;
                aac_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2)?;
                // Set bitrate to 128kbps (bytes per second = bitrate / 8)
                aac_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, 16000)?;
                
                Ok::<IMFMediaType, windows::core::Error>(aac_type)
            } {
                Ok(aac_type) => {
                    println!("Created AAC output media type for MP4 container");
                    Some(aac_type)
                }
                Err(e) => {
                    eprintln!("Failed to create AAC media type: {:?}. Disabling audio.", e);
                    audio_capture = None;
                    None
                }
            }
        } else {
            None
        };
    
        // --- Create PCM Audio Input Media Type ---
        let audio_input_type = if let Some(ref ac) = audio_capture {
            match unsafe {
                // Create a PCM media type for the input
                let pcm_type = MFCreateMediaType()?;
                pcm_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
                pcm_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)?;
                pcm_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?;
                pcm_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, 44100)?;
                pcm_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2)?;
                pcm_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, 4)?; // channels * (bits/8)
                pcm_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, 176400)?; // samplerate * blockalign
                
                Ok::<IMFMediaType, windows::core::Error>(pcm_type)
            } {
                Ok(pcm_type) => {
                    println!("Created PCM input media type");
                    Some(pcm_type)
                }
                Err(e) => {
                    eprintln!("Failed to create PCM media type: {:?}. Disabling audio.", e);
                    audio_capture = None;
                    None
                }
            }
        } else {
            None
        };
    
        // --- Create Sample Writer ---
        println!("Creating SampleWriter...");
        // SampleWriter takes both the output (AAC) and input (PCM) types
        let sample_writer = Arc::new(SampleWriter::new(
            stream,
            &video_output_type,
            audio_output_type.as_ref(),
            audio_input_type.as_ref(),
            raw_audio_receiver,
        )?);
        println!("SampleWriter created.");
    
        // --- Update AudioCapture Stream Index ---
        let actual_audio_stream_index = sample_writer.audio_stream_index;
        println!("Actual audio stream index from SampleWriter: {:?}", actual_audio_stream_index);
        if let Some(ref mut ac) = audio_capture {
            if let Some(index) = actual_audio_stream_index {
                ac.set_stream_index(index);
            } else {
                println!("Audio disabled in SampleWriter, not setting stream index in AudioCapture.");
                audio_capture = None;
            }
        }
    
        // --- Setup Video Sample Generator (No change) ---
        println!("Creating SampleGenerator...");
        let mut sample_generator = SampleGenerator::new(
            d3d_device.clone(),
            monitor_handle,
            input_size,
            output_size,
            frame_rate,
        )?;
        println!("SampleGenerator created successfully");
    
        // --- Setup Video Encoder Callbacks (No change) ---
        println!("Setting video encoder callbacks...");
        video_encoder.set_sample_requested_callback(move || -> Result<Option<VideoEncoderInputSample>> {
            sample_generator.generate()
        });
        video_encoder.set_sample_rendered_callback({
            let writer_arc = sample_writer.clone();
            move |sample| -> Result<()> {
                writer_arc.write_video_sample(sample.sample())
            }
        });
        println!("Video encoder callbacks set.");
    
        // --- Return the Session ---
        Ok(Self {
            video_encoder,
            sample_writer,
            audio_capture,
            raw_audio_sender: Some(raw_audio_sender),
        })
    }
}

impl VideoEncodingSession for MFVideoEncodingSession {
    fn start(&mut self) -> Result<()> {
        println!("Starting MFVideoEncodingSession...");
    
        // Start SampleWriter first (so it's ready for samples)
        self.sample_writer.start()?;
        println!("SampleWriter started (BeginWriting called).");
    
        // Start Video Encoder Thread
        if !self.video_encoder.try_start()? {
            eprintln!("Video encoder failed to start or already started.");
            return Err(windows::core::Error::new(E_FAIL, "Failed to start video encoder".into()));
        }
        println!("VideoEncoder thread started successfully.");
    
        // Start Audio Capture Thread (if present)
        if let Some(ref mut ac) = self.audio_capture {
            if let Some(sender) = self.raw_audio_sender.take() {
                ac.start(sender)?;
                println!("AudioCapture thread started.");
            } else {
                eprintln!("Critical Error: Raw audio sender unavailable during start.");
                self.video_encoder.stop().ok();
                self.sample_writer.stop().ok();
                return Err(windows::core::Error::new(E_FAIL, "Internal error: audio sender missing".into()));
            }
        }
        
        println!("MFVideoEncodingSession start sequence complete.");
        Ok(())
    }

    fn stop(&mut self) -> Result<()> {
        println!("Stopping MFVideoEncodingSession...");
        let mut first_error: Option<windows::core::Error> = None;
    
        // Define helper to record the first error
        let mut record_error = |res: Result<()>, component: &str| {
            if let Err(e) = res {
                eprintln!("Error stopping {}: {:?}", component, e);
                if first_error.is_none() {
                    first_error = Some(e);
                }
            } else {
                println!("{} stopped successfully.", component);
            }
        };
    
        // Order: Stop Input -> Stop Processing -> Stop Output
        // 1. Stop Audio Capture first (stops feeding raw data)
        if let Some(ref mut ac) = self.audio_capture.take() {
            record_error(ac.stop(), "AudioCapture");
        } else {
            println!("AudioCapture already stopped or None.");
        }
        
        // Close the raw audio sender explicitly
        drop(self.raw_audio_sender.take());
    
        // 2. Stop Video Encoder
        record_error(self.video_encoder.stop(), "VideoEncoder");
    
        // 3. Stop Sample Writer (allows it to process any remaining samples and finalize)
        record_error(self.sample_writer.stop(), "SampleWriter");
    
        println!("MFVideoEncodingSession stop sequence complete.");
    
        // Return the first error encountered, or Ok(())
        match first_error {
            Some(e) => Err(e),
            None => Ok(()),
        }
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
        audio_source: AudioSource,
    ) -> Result<Box<dyn VideoEncodingSession>> {
        let session = Box::new(MFVideoEncodingSession::new(
            d3d_device,
            monitor_handle,
            &self.encoder_device,
            resolution,
            bit_rate,
            frame_rate,
            stream,
            audio_source,
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
        frame_rate: u32,
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

        let target_frame_duration = TimeSpan {
            Duration: 10_000_000 / (frame_rate as i64), // TimeSpan is in 100ns units
        };
        if target_frame_duration.Duration <= 0 {
             // Handle error: frame_rate too high or zero
             return Err(windows::core::Error::new(windows::Win32::Foundation::E_INVALIDARG, "Invalid frame rate".into()));
        }
    
        println!("SampleGenerator::new completed successfully");
        Ok(Self {
            d3d_device,
            d3d_context,
            video_processor,
            compose_texture,
            render_target_view,
            frame_generator,
            target_frame_duration,
            next_target_relative_timestamp: None,
            last_returned_relative_timestamp: None,
            seen_first_time_stamp: false,
            first_timestamp: TimeSpan::default(),
        })
    }

    pub fn generate(&mut self) -> Result<Option<VideoEncoderInputSample>> {
        loop {
            // Try to get the next frame with a short timeout.
            match self.frame_generator.try_get_next_frame(33) {
                Ok(Some(frame)) => {
                    // --- Timestamp Calculation ---
                    let frame_qpc_time = frame.frame_info.LastPresentTime;
                    let timestamp = convert_qpc_to_timespan(frame_qpc_time)?;
    
                    // On the first frame, set the base timestamp and initialize the scheduled target.
                    if !self.seen_first_time_stamp {
                        self.first_timestamp = timestamp;
                        // The next target is set to the target frame duration, relative to the first timestamp.
                        self.next_target_relative_timestamp = Some(TimeSpan {
                            Duration: self.target_frame_duration.Duration,
                        });
                        self.seen_first_time_stamp = true;
                    }
                    // Compute the current relative timestamp.
                    let current_relative_timestamp = TimeSpan {
                        Duration: timestamp.Duration.saturating_sub(self.first_timestamp.Duration),
                    };
                    // --- End Timestamp Calculation ---
    
                    // --- Scheduled Rate Control Logic ---
                    println!(
                        "Current relative timestamp: {} ms",
                        current_relative_timestamp.Duration / 10_000
                    );
                    // Retrieve the next scheduled target time.
                    let scheduled_time = self.next_target_relative_timestamp
                        .expect("next_target_relative_timestamp should be set on first frame");
                    println!(
                        "Next scheduled target: {} ms",
                        scheduled_time.Duration / 10_000
                    );
                    println!(
                        "Target frame duration: {} ms",
                        self.target_frame_duration.Duration / 10_000
                    );
    
                    // Accept the frame if we have reached the scheduled target.
                    if current_relative_timestamp.Duration >= scheduled_time.Duration {
                        println!("Returning frame");
                        // Update the scheduler: Add the target frame duration to the current scheduled time.
                        self.next_target_relative_timestamp = Some(TimeSpan {
                            Duration: scheduled_time.Duration + self.target_frame_duration.Duration,
                        });
                        // --- End Scheduled Rate Control Logic ---
    
                        // --- Frame Processing ---
                        let start_time = std::time::Instant::now();
                        match self.generate_from_frame(&frame, current_relative_timestamp) {
                            Ok(sample) => {
                                let elapsed = start_time.elapsed();
                                println!("Frame generation took: {:?}", elapsed);
                                self.last_returned_relative_timestamp = Some(current_relative_timestamp);
                                return Ok(Some(sample)); // Return the processed frame.
                            }
                            Err(error) => {
                                eprintln!(
                                    "Error during input sample generation: {:?} - {}",
                                    error.code(),
                                    error.message()
                                );
                                return Ok(None); // Signal end-of-stream, or you might choose to continue.
                            }
                        }
                    } else {
                        println!("Skipping frame (arrived too early)");
                        continue; // Frame arrived before scheduled target; skip it.
                    }
                }
                Ok(None) => {
                    // Timeout acquiring a frame.
                    return Ok(None);
                }
                Err(e) if e.code() == DXGI_ERROR_ACCESS_LOST => {
                    eprintln!("DXGI Access Lost in frame generation: {:?}", e);
                    return Err(e);
                }
                Err(e) => {
                    eprintln!("Error getting next frame: {:?}", e);
                    return Err(e);
                }
            }
        }
    }    

    fn generate_from_frame(
        &mut self,
        frame: &AcquiredFrame,
        relative_timestamp: TimeSpan,
    ) -> Result<VideoEncoderInputSample> {

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
    pub fn new(
        stream: IRandomAccessStream,
        video_output_type: &IMFMediaType,
        audio_output_type: Option<&IMFMediaType>,  // AAC output type
        audio_input_type: Option<&IMFMediaType>,   // PCM input type
        audio_sample_receiver: Receiver<(u32, AudioDataPacket)>,
    ) -> Result<Self> {
        unsafe {
            let empty_attributes = {
                let mut attributes = None;
                windows::Win32::Media::MediaFoundation::MFCreateAttributes(&mut attributes, 1)?;
                if let Some(attr) = attributes.as_ref() {
                    // Example: Disable throttling
                    attr.SetUINT32(&windows::Win32::Media::MediaFoundation::MF_SINK_WRITER_DISABLE_THROTTLING, 1)?;
                }
                attributes
            };
            let byte_stream = MFCreateMFByteStreamOnStreamEx(&stream)?;
            let sink_writer = MFCreateSinkWriterFromURL(&HSTRING::from(".mp4"), &byte_stream, empty_attributes.as_ref())?;

            // --- Add Video Stream ---
            let video_stream_index = sink_writer.AddStream(video_output_type)?;
            sink_writer.SetInputMediaType(
                video_stream_index,
                video_output_type,
                None,
            )?;
            println!("SampleWriter: Added Video Stream, Index: {}", video_stream_index);

            // --- Add Audio Stream (if provided) ---
            let mut audio_stream_index: Option<u32> = None;
            if let (Some(output_type), Some(input_type)) = (audio_output_type, audio_input_type) {
                // Add stream with AAC output type
                let index = sink_writer.AddStream(output_type)?;
                
                // Set input type to PCM - this allows Media Foundation to handle the conversion
                sink_writer.SetInputMediaType(
                    index,
                    input_type,
                    None,
                )?;
                
                audio_stream_index = Some(index);
                println!("SampleWriter: Added Audio Stream, Index: {}", index);
                println!("SampleWriter: Configured automatic PCM to AAC conversion");
            } else {
                println!("SampleWriter: No audio stream provided or missing types.");
            }

            Ok(Self {
                _stream: stream,
                sink_writer,
                video_stream_index,
                audio_stream_index,
                audio_packet_receiver: audio_sample_receiver,
            })
        }
    }

    pub fn start(&self) -> Result<()> {
        unsafe { self.sink_writer.BeginWriting() }
    }

    pub fn stop(&self) -> Result<()> {
        println!("SampleWriter: Stopping (Draining Audio samples & Finalizing)...");
        // 1. Drain any remaining audio samples from the channel
        match self.write_pending_audio_samples() {
            Ok(_) => println!("SampleWriter: Drained pending audio samples."),
            Err(e) => {
                eprintln!(
                    "SampleWriter: Error draining audio samples during stop: {:?}",
                    e
                );
                // Decide whether to proceed with Finalize or return error
                // For now, let's proceed to Finalize to attempt saving the file.
            }
        }

        // 2. Finalize the Sink Writer
        println!("SampleWriter: Calling Finalize...");
        let result = unsafe { self.sink_writer.Finalize() };
        match &result {
            Ok(_) => println!("SampleWriter: Finalize successful."),
            Err(e) => {
                eprintln!(
                    "SampleWriter: Finalize failed: HRESULT={:?}, Msg={}",
                    e.code(),
                    e.message()
                );
            }
        }
        result // Return the result of Finalize
    }

    pub fn write_video_sample(&self, sample: &IMFSample) -> Result<()> {
        // Interleave by trying to write pending audio first (best effort)
        // Ignore error here as it's logged in the function itself
        let _ = self.write_pending_audio_samples();

        unsafe {
            self.sink_writer
                .WriteSample(self.video_stream_index, sample)
                .map_err(|e| {
                     // Add context to error if WriteSample fails (e.g., MF_E_NOTACCEPTING)
                     eprintln!("SinkWriter::WriteSample (Video) failed: HRESULT={:?}, Msg={}", e.code(), e.message());
                     e
                })
        }
    }

    /// Drains the audio channel and writes any pending audio samples.
    /// Call this periodically (e.g., before writing video) and before Finalize.
    fn write_pending_audio_samples(&self) -> Result<()> {
        if self.audio_stream_index.is_none() { return Ok(()); }
        let target_audio_index = self.audio_stream_index.unwrap();

        loop {
            match self.audio_packet_receiver.try_recv() {
                Ok((_, packet)) => { // Receive (stream_index, AudioDataPacket)
                    // Convert AudioDataPacket to IMFSample using the helper function
                    let sample = create_imf_sample_from_packet(packet)?;
                    
                    // Write the sample directly to the sink writer
                    // The sink writer handles the PCM to AAC conversion automatically
                    unsafe {
                        let write_result = self.sink_writer.WriteSample(target_audio_index, &sample);
                        
                        if let Err(e) = write_result {
                            eprintln!("SinkWriter::WriteSample (Audio) failed: {:?}", e);
                            return Err(e); // Stop draining on error
                        }
                    }
                }
                Err(TryRecvError::Empty) => break, // Channel empty
                Err(TryRecvError::Disconnected) => {
                    println!("SampleWriter: Audio packet channel disconnected.");
                    break; // Capture stopped
                }
            }
        }
        Ok(())
    }
}

impl Drop for SampleWriter {
    fn drop(&mut self) {
        println!("SampleWriter dropped.");
        // Finalize should have been called via MFVideoEncodingSession::stop -> self.stop().
        // The IMFSinkWriter COM object releases itself when the ref count goes to zero.
    }
}
