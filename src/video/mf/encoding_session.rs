use std::{sync::{mpsc::{sync_channel, Receiver, TryRecvError, SyncSender}, Arc}, time::Duration};

use windows::{
    core::{ComInterface, Result, HSTRING},
    Foundation::TimeSpan, // Keep for calculating duration, but not for sample timestamps
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Foundation::E_FAIL, Graphics::{
            Direct3D11::{
                ID3D11Device, ID3D11DeviceContext, ID3D11RenderTargetView, ID3D11Resource, ID3D11Texture2D, D3D11_BIND_RENDER_TARGET, D3D11_BIND_SHADER_RESOURCE, D3D11_BOX, D3D11_CPU_ACCESS_READ, D3D11_MAPPED_SUBRESOURCE, D3D11_MAP_READ, D3D11_TEXTURE2D_DESC, D3D11_USAGE_DEFAULT, D3D11_USAGE_STAGING
            },
            Dxgi::{
                Common::{DXGI_FORMAT_B8G8R8A8_UNORM, DXGI_FORMAT_NV12, DXGI_SAMPLE_DESC},
                DXGI_ERROR_ACCESS_LOST,
            },
            Gdi::HMONITOR,
        }, Media::MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFAudioFormat_AAC, MFAudioFormat_PCM, MFCreateAttributes, MFCreateMFByteStreamOnStreamEx, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFCreateSinkWriterFromURL, MFMediaType_Audio, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE
        }, System::Performance::{QueryPerformanceCounter, QueryPerformanceFrequency}
    },
};

use crate::{
    // Assuming these functions/structs are updated as discussed:
    // - `convert_qpc_to_mf_timespan` exists and returns Result<i64>
    // - `create_imf_sample_from_packet` signature is `fn(AudioDataPacket, i64) -> Result<IMFSample>`
    // - `AudioDataPacket.timestamp` contains raw QPC (i64)
    audio::{convert_qpc_to_mf_timespan, create_imf_sample_from_packet, AudioCapture, AudioDataPacket, AudioSource},
    capture::{AcquiredFrame, CaptureFrameGenerator},
    video::{
        encoding_session::{VideoEncoderSessionFactory, VideoEncodingSession},
        // mf::audio_encoder::{AudioEncoder, EncodedAudioPacket}, // Not used directly here
        util::ensure_even_size,
        CLEAR_COLOR
    }
};

use super::{
    // Assuming `VideoEncoderInputSample` is updated:
    // - `fn new(timestamp_100ns: i64, duration_100ns: i64, texture: ID3D11Texture2D)`
    // - has fields `timestamp: i64`, `duration: i64`
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

struct SampleGenerator {
    d3d_device: ID3D11Device,
    d3d_context: ID3D11DeviceContext,

    video_processor: VideoProcessor,
    compose_texture: ID3D11Texture2D,
    render_target_view: ID3D11RenderTargetView,

    frame_generator: CaptureFrameGenerator,

    frame_duration_100ns: i64, // Store duration directly as i64 (100ns)
    anchor_mf_time: i64, // Store the anchor time as i64 (100ns)

     // --- Rate Control State ---
     seen_first_frame: bool,     // Have we processed the very first frame?
     // Timestamp of the first frame *relative to the anchor* (100ns)
     first_frame_relative_mf_time: i64,
     // The scheduled presentation time for the *next* frame we want to output,
     // relative to the anchor time (100ns).
     next_target_relative_mf_timestamp: i64,
}


pub struct SampleWriter {
    _stream: IRandomAccessStream,
    sink_writer: IMFSinkWriter,
    video_stream_index: u32,
    audio_stream_index: Option<u32>,
    audio_packet_receiver: Receiver<(u32, AudioDataPacket)>,
    anchor_mf_time: i64, // Store the anchor time as i64 (100ns)
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
        let input_size = ensure_even_size(resolution);
        println!("Input size set: {:?}", input_size);
        
        let output_size = ensure_even_size(resolution);
        println!("Output size set: {:?}", output_size);

        let mut initial_qpc: i64 = 0;
        unsafe { QueryPerformanceCounter(&mut initial_qpc)?; }
        println!("QueryPerformanceCounter result: {}", initial_qpc);
        
        let anchor_mf_time = convert_qpc_to_mf_timespan(initial_qpc)?;
        println!("Anchor MF time: {}", anchor_mf_time);

        let mut video_encoder = VideoEncoder::new(
            encoder_device,
            d3d_device.clone(),
            output_size,
            output_size,
            bit_rate,
            frame_rate,
        )?;
        println!("Video encoder created");
        
        let video_output_type = video_encoder.output_type().clone();
        println!("Video output type obtained");

        let placeholder_index: u32 = 0;
        let mut audio_capture = match AudioCapture::new(audio_source, placeholder_index) {
            Ok(ac) => {
                println!("Audio capture created successfully");
                Some(ac)
            },
            Err(e) => {
                println!("Failed to create audio capture: {:?}", e);
                None
            },
        };

        let (raw_audio_sender, raw_audio_receiver) = sync_channel::<(u32, AudioDataPacket)>(100);
        println!("Audio channel created with buffer size 100");

        let audio_output_type = if audio_capture.is_some() {
            unsafe {
                let aac_type = MFCreateMediaType()?;
                println!("Created AAC media type");
                
                aac_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
                println!("Set major type to Audio");
                
                aac_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
                println!("Set subtype to AAC");
                
                aac_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, 44100)?;
                println!("Set sample rate to 44100");
                
                aac_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2)?;
                println!("Set channel count to 2");
                
                aac_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, 16000)?;
                println!("Set average bytes per second to 16000");
                
                Some(aac_type)
            }
        } else {
            println!("No audio output type created (no audio capture)");
            None
        };

        let audio_input_type = if let Some(ref ac) = audio_capture {
            let input_type = ac.clone_media_type()?;
            println!("Audio input type cloned from capture device");
            Some(input_type)
        } else {
            println!("No audio input type (no audio capture)");
            None
        };

        let sample_writer = Arc::new(SampleWriter::new(
            stream,
            &video_output_type,
            audio_output_type.as_ref(),
            audio_input_type.as_ref(),
            raw_audio_receiver,
            anchor_mf_time,
        )?);
        println!("Sample writer created");

        if let Some(ref mut ac) = audio_capture {
            if let Some(index) = sample_writer.audio_stream_index {
                ac.set_stream_index(index);
                println!("Audio stream index set to {}", index);
            } else {
                audio_capture = None;
                println!("Audio capture disabled (no audio stream index)");
            }
        }

        let mut sample_generator = SampleGenerator::new(
            d3d_device.clone(),
            monitor_handle,
            input_size,
            output_size,
            frame_rate,
            anchor_mf_time,
        )?;
        println!("Sample generator created");

        video_encoder.set_sample_requested_callback(move || -> Result<Option<VideoEncoderInputSample>> {
            println!("Sample requested callback invoked");
            sample_generator.generate()
        });
        println!("Sample requested callback set");
        
        video_encoder.set_sample_rendered_callback({
            let writer_arc = sample_writer.clone();
            move |sample| -> Result<()> {
                println!("Sample rendered callback invoked");
                writer_arc.write_video_sample(sample.sample())
            }
        });
        println!("Sample rendered callback set");

        println!("New recorder instance created successfully");
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
        self.sample_writer.start()?; // Start writing (headers etc.)

        // Start encoder thread (which will call sample_requested_callback -> generate)
        if !self.video_encoder.try_start()? {
            self.sample_writer.stop().ok(); // Attempt cleanup
            return Err(windows::core::Error::new(E_FAIL, "Failed to start video encoder".into()));
        }

        // Start audio capture thread
        if let Some(ref mut ac) = self.audio_capture {
            if let Some(sender) = self.raw_audio_sender.take() {
                 // Check if audio stream was successfully added before starting
                if self.sample_writer.audio_stream_index.is_some() {
                    ac.start(sender)?;
                } else {
                    // Don't start audio if the writer doesn't have an audio stream
                    eprintln!("Audio stream not created in SinkWriter, not starting audio capture.");
                    self.audio_capture = None; // Disable audio capture instance
                }
            } else {
                 // Should not happen if logic is correct, but handle defensively
                self.video_encoder.stop().ok();
                self.sample_writer.stop().ok();
                return Err(windows::core::Error::new(E_FAIL, "Internal error: audio sender missing before start".into()));
            }
        }

        Ok(())
    }

    fn stop(&mut self) -> Result<()> {
        let mut first_error: Option<windows::core::Error> = None;

        // Helper to record the first error encountered
        let mut record_error = |res: Result<()>, _component: &str| {
            if let Err(e) = res {
                // eprintln!("Error stopping {}: {:?}", component, e); // Optional logging
                if first_error.is_none() {
                    first_error = Some(e);
                }
            }
        };

        // 1. Stop Audio Capture Thread (signal it to stop sending)
        if let Some(ref mut ac) = self.audio_capture.take() { // Take ownership
            record_error(ac.stop(), "AudioCapture");
        }

        // 2. Drop the sender to signal disconnection to SampleWriter's audio loop (if running)
        drop(self.raw_audio_sender.take());

        // 3. Stop Video Encoder Thread (signal it to stop requesting frames and finish processing)
        // This will internally signal the SampleGenerator to stop via the callback returning None/Err.
        record_error(self.video_encoder.stop(), "VideoEncoder"); // This should block until the encoder thread joins

        // 4. Stop Sample Writer (flush pending audio/video, finalize file)
        // Need to use Arc::try_unwrap or similar if SampleWriter needs exclusive access to stop,
        // but here we assume stop() can be called on the Arc reference.
        // Ensure the writer processes any final samples flushed by the encoder.
        record_error(self.sample_writer.stop(), "SampleWriter");

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
        input_size: SizeInt32, // Size for processor input (e.g., desktop size)
        output_size: SizeInt32, // Size for processor output (encoder input size)
        frame_rate: u32,
        anchor_mf_time: i64, // Receive anchor time
    ) -> Result<Self> {
        let d3d_context = unsafe { d3d_device.GetImmediateContext() }?;

        // Calculate frame duration in 100ns units
        let frame_duration_100ns = if frame_rate > 0 {
            10_000_000 / (frame_rate as i64)
        } else {
            // Return error or use a default? Error is safer.
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid frame rate (must be > 0)".into()
            ));
        };
        if frame_duration_100ns <= 0 {
             return Err(windows::core::Error::new(
                 windows::Win32::Foundation::E_INVALIDARG,
                 "Calculated frame duration is invalid (frame rate too high?)".into()
             ));
        }

        // Processor setup (Input BGRA -> Output NV12)
        let video_processor = VideoProcessor::new(
            d3d_device.clone(),
            DXGI_FORMAT_B8G8R8A8_UNORM, // Input format from desktop duplication
            input_size,                 // Input size for processor
            DXGI_FORMAT_NV12,           // Output format for encoder
            output_size,                // Output size for processor/encoder
        )?;

        // Texture for composition (matches processor input)
        let texture_desc = D3D11_TEXTURE2D_DESC {
            Width: input_size.Width as u32,
            Height: input_size.Height as u32,
            ArraySize: 1,
            MipLevels: 1,
            Format: DXGI_FORMAT_B8G8R8A8_UNORM,
            SampleDesc: DXGI_SAMPLE_DESC { Count: 1, ..Default::default() },
            Usage: D3D11_USAGE_DEFAULT,
            BindFlags: (D3D11_BIND_RENDER_TARGET.0 | D3D11_BIND_SHADER_RESOURCE.0) as u32,
            ..Default::default()
        };
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

        // Frame generator (gets raw frames)
        let frame_generator = CaptureFrameGenerator::new(d3d_device.clone(), monitor_handle)?;

        // Calculate frame duration in 100ns units
        let frame_duration_100ns = if frame_rate > 0 {
            10_000_000 / (frame_rate as i64)
        } else {
            return Err(windows::core::Error::new(
                windows::Win32::Foundation::E_INVALIDARG,
                "Invalid frame rate (must be > 0)".into()
            ));
        };
        if frame_duration_100ns <= 0 {
             return Err(windows::core::Error::new(
                 windows::Win32::Foundation::E_INVALIDARG,
                 "Calculated frame duration is invalid".into()
             ));
        }

        Ok(Self {
            d3d_device,
            d3d_context,
            video_processor,
            compose_texture,
            render_target_view,
            frame_generator,
            frame_duration_100ns, // Store calculated duration
            anchor_mf_time,        // Store anchor time
            seen_first_frame: false,
            first_frame_relative_mf_time: 0, // Will be set on first frame
            // Start by targeting the first frame immediately (relative time 0)
            // It will be updated *after* the first frame is processed.
            next_target_relative_mf_timestamp: 0,
        })
    }

    /// Tries to get the next frame, process it, and return an input sample for the encoder.
    pub fn generate(&mut self) -> Result<Option<VideoEncoderInputSample>> {
        loop {
            match self.frame_generator.try_get_next_frame(33) { // ~30fps timeout
                Ok(Some(frame)) => {
                    // --- Timestamp Calculation ---
                    let frame_qpc_time = frame.frame_info.LastPresentTime;
                    if frame_qpc_time == 0 {
                         eprintln!("Warning: Skipping frame with zero QPC timestamp.");
                         continue;
                    }
                    let frame_mf_time = convert_qpc_to_mf_timespan(frame_qpc_time)?;
                    // Calculate timestamp relative to the session anchor
                    let current_relative_mf_timestamp = frame_mf_time.saturating_sub(self.anchor_mf_time).max(0);

                    // --- Rate Control Logic ---
                    if !self.seen_first_frame {
                        // This is the very first frame we've received. Process it immediately.
                        self.seen_first_frame = true;
                        self.first_frame_relative_mf_time = current_relative_mf_timestamp;
                        // Schedule the target time for the *next* frame based on this first one.
                        self.next_target_relative_mf_timestamp = current_relative_mf_timestamp + self.frame_duration_100ns;

                        // Proceed to generate this first frame
                        match self.generate_from_frame(&frame, current_relative_mf_timestamp) {
                             Ok(sample) => return Ok(Some(sample)),
                             Err(e) => {
                                 eprintln!("Error generating first frame: {:?}", e);
                                 return Err(e); // Propagate error
                             }
                        }
                    } else {
                        // We have seen the first frame, now check against the schedule.
                        if current_relative_mf_timestamp >= self.next_target_relative_mf_timestamp {
                            // This frame meets or exceeds the scheduled time. Process it.

                            // Calculate the target time for the *next* frame based on the
                            // *current target*, not the actual frame time, to keep cadence.
                            let current_target = self.next_target_relative_mf_timestamp;
                            self.next_target_relative_mf_timestamp = current_target + self.frame_duration_100ns;

                            // Process the frame using its actual relative timestamp
                             match self.generate_from_frame(&frame, current_relative_mf_timestamp) {
                                 Ok(sample) => return Ok(Some(sample)),
                                 Err(e) => {
                                     eprintln!("Error generating frame: {:?}", e);
                                     // Consider if error should stop generation or just skip frame
                                     return Err(e);
                                 }
                            }
                        } else {
                            // Frame arrived too early. Skip it and try getting the next one.
                            //println!("Skipping frame (arrived too early)"); // Optional debug log
                            continue;
                        }
                    }
                }
                Ok(None) => {
                    // Timeout occurred, no new frame available yet. Signal encoder to wait.
                    return Ok(None);
                }
                Err(e) if e.code() == DXGI_ERROR_ACCESS_LOST => {
                    eprintln!("DXGI Access Lost in frame generation: {:?}", e);
                    return Err(e); // Critical error
                }
                Err(e) => {
                    eprintln!("Error getting next frame: {:?}", e);
                    return Err(e); // Propagate other errors
                }
            }
        }
    }

    /// Processes a single AcquiredFrame and creates a VideoEncoderInputSample.
    fn generate_from_frame(
        &mut self,
        frame: &AcquiredFrame,
        relative_mf_timestamp: i64, // Receive relative timestamp
    ) -> Result<VideoEncoderInputSample> {

        let frame_texture = &frame.texture; // This is the BGRA texture from duplication

        // Describe the region to copy (entire texture)
        let desc = unsafe {
            let mut desc = D3D11_TEXTURE2D_DESC::default();
            frame_texture.GetDesc(&mut desc);
            desc
        };
        let region = D3D11_BOX {
            left: 0, right: desc.Width, top: 0, bottom: desc.Height, back: 1, front: 0,
        };

        // --- GPU Processing ---
        unsafe {
            // 1. Clear render target (optional, good practice)
            self.d3d_context.ClearRenderTargetView(&self.render_target_view, &CLEAR_COLOR);

            // 2. Copy the captured frame (BGRA) to our composition texture
            self.d3d_context.CopySubresourceRegion(
                &self.compose_texture, // Destination (BGRA)
                0, 0, 0, 0,
                frame_texture,         // Source (BGRA)
                0, Some(&region),
            );

            // 3. Process the composition texture (BGRA -> NV12) using the VideoProcessor
            self.video_processor.process_texture(&self.compose_texture)?;

            // 4. Get the resulting NV12 texture from the processor
            let video_output_texture = self.video_processor.output_texture(); // This is NV12

            // 5. (Optional but recommended) Create a *new* texture and copy the result into it.
            //    This decouples the sample's texture lifetime from the video processor's output.
            let sample_texture = {
                 let output_desc = {
                     let mut desc = D3D11_TEXTURE2D_DESC::default();
                     video_output_texture.GetDesc(&mut desc);
                     desc
                 };
                 // Ensure the new texture has the same description
                 let mut texture = None;
                 self.d3d_device.CreateTexture2D(&output_desc, None, Some(&mut texture))?;
                 texture.unwrap()
            };
            self.d3d_context.CopyResource(&sample_texture, video_output_texture);

            let relative_timespan = TimeSpan { Duration: relative_mf_timestamp };

            // --- Create Input Sample ---
            // Use the calculated relative timestamp and frame duration
            Ok(VideoEncoderInputSample::new(
                relative_timespan, // <<< Pass the TimeSpan
                sample_texture,    // Pass the copied NV12 texture
            ))
        }
    }
}

unsafe impl Send for SampleWriter {}
unsafe impl Sync for SampleWriter {}
impl SampleWriter {
    pub fn new(
        stream: IRandomAccessStream,
        video_output_type: &IMFMediaType, // The H.264 or HEVC type
        audio_output_type: Option<&IMFMediaType>, // The AAC or other encoded type
        audio_input_type: Option<&IMFMediaType>, // The PCM type expected by SinkWriter input
        audio_packet_receiver: Receiver<(u32, AudioDataPacket)>,
        anchor_mf_time: i64, // Receive anchor time
    ) -> Result<Self> {
        unsafe {
            // Attributes to disable throttling
            let attributes = {
                let mut attributes = None;
                MFCreateAttributes(&mut attributes, 1)?;
                attributes.as_ref().unwrap().SetUINT32(&windows::Win32::Media::MediaFoundation::MF_SINK_WRITER_DISABLE_THROTTLING, 1)?;
                attributes
            };

            let byte_stream = MFCreateMFByteStreamOnStreamEx(&stream)?;
            let sink_writer = MFCreateSinkWriterFromURL(
                &HSTRING::from(".mp4"), // Or other container format HSTRING
                &byte_stream,
                attributes.as_ref() // Pass attributes
            )?;

            // --- Add Video Stream ---
            // Provide the *encoded* video format (H.264)
            let video_stream_index = sink_writer.AddStream(video_output_type)?;
            // Set the input media type for the video stream *if required* by the sink.
            // Often for video, the input type matches the output type for raw encoded data.
            // If the sink expects something else (unlikely for raw encoded), set it here.
            // sink_writer.SetInputMediaType(video_stream_index, video_output_type, None)?; // Example if input=output

            // --- Add Audio Stream ---
            let mut audio_stream_index: Option<u32> = None;
            if let (Some(output_type), Some(input_type)) = (audio_output_type, audio_input_type) {
                 match sink_writer.AddStream(output_type) { // Add stream with *encoded* format (AAC)
                    Ok(index) => {
                        // Set the *input* format the sink expects for this stream (PCM)
                        match sink_writer.SetInputMediaType(index, input_type, None) {
                            Ok(_) => {
                                audio_stream_index = Some(index);
                            }
                            Err(e) => {
                                eprintln!("Failed to set audio input media type on SinkWriter: {:?}", e);
                                // Decide how to handle: continue without audio or fail hard?
                                // return Err(e); // Fail hard example
                            }
                        }
                    }
                    Err(e) => {
                         eprintln!("Failed to add audio stream to SinkWriter: {:?}", e);
                         // Decide how to handle
                         // return Err(e);
                    }
                 }
            }

            Ok(Self {
                _stream: stream, // Keep stream alive
                sink_writer,
                video_stream_index,
                audio_stream_index,
                audio_packet_receiver,
                anchor_mf_time, // Store anchor time
            })
        }
    }

    /// Starts the writing process (writes headers).
    pub fn start(&self) -> Result<()> {
        unsafe { self.sink_writer.BeginWriting() }
    }

    /// Stops writing, finalizes the file.
    pub fn stop(&self) -> Result<()> {
        // Write any remaining audio samples that might be buffered in the channel
        let _ = self.write_pending_audio_samples(); // Use ? if error should stop finalization

        // Finalize the media file
        let result = unsafe { self.sink_writer.Finalize() };
        if let Err(e) = &result {
            eprintln!("SampleWriter: Finalize failed: HRESULT={:?}", e);
        }
        result
    }

    /// Writes an encoded video sample to the sink.
    pub fn write_video_sample(&self, sample: &IMFSample) -> Result<()> {
        // Process any pending audio first to maintain order as much as possible
        self.write_pending_audio_samples()?; // Use ? to propagate audio writing errors

        unsafe {
            self.sink_writer
                .WriteSample(self.video_stream_index, sample)
                .map_err(|e| {
                    eprintln!("SinkWriter::WriteSample (Video) failed: HRESULT={:?}", e);
                    e // Return the error
                })
        }
    }

    /// Processes and writes buffered audio packets.
    fn write_pending_audio_samples(&self) -> Result<()> {
        let Some(target_audio_index) = self.audio_stream_index else {
            // No audio stream configured, drain the channel silently if needed or just return
            // Drain example:
             while let Ok(_) = self.audio_packet_receiver.try_recv() {}
            return Ok(());
        };

        loop {
            match self.audio_packet_receiver.try_recv() {
                Ok((_stream_idx, packet)) => { // Original stream index might not be needed now
                    // Assuming packet.timestamp is raw QPC (i64)
                    if packet.timestamp == 0 {
                        // Skip packets with zero timestamp if they indicate invalid data
                        eprintln!("Warning: Skipping audio packet with zero QPC timestamp.");
                        continue;
                    }
                    // --- Timestamp Calculation ---
                    let packet_mf_time = convert_qpc_to_mf_timespan(packet.timestamp)?;
                    let relative_mf_timestamp = packet_mf_time.saturating_sub(self.anchor_mf_time).max(0);

                    // --- Create and Write Sample ---
                    // Assuming create_imf_sample_from_packet takes (packet, relative_ts)
                    match create_imf_sample_from_packet(packet, relative_mf_timestamp) {
                        Ok(sample) => {
                            unsafe {
                                let write_result = self.sink_writer.WriteSample(target_audio_index, &sample);
                                if let Err(e) = write_result {
                                    eprintln!("SinkWriter::WriteSample (Audio) failed: {:?}", e);
                                    // Decide if this error is fatal for the writer
                                    return Err(e); // Propagate error
                                }
                            }
                        }
                        Err(e) => {
                            eprintln!("Failed to create IMFSample from audio packet: {:?}", e);
                            // Decide if this error is fatal
                            // return Err(e); // Example: treat as fatal
                        }
                    }
                }
                Err(TryRecvError::Empty) => {
                    // No more packets currently available in the channel
                    break;
                }
                Err(TryRecvError::Disconnected) => {
                    // Sender has been dropped, no more packets will arrive
                    // This is expected during shutdown.
                    break;
                }
            }
        }
        Ok(())
    }
}

impl Drop for SampleWriter {
    fn drop(&mut self) {
        // Usually, explicit stop() is preferred over relying on drop for finalization,
        // as drop should not ideally panic or return errors.
        // If stop() wasn't called, attempt a best-effort finalize, but log errors.
        // Note: This might be called if MFVideoEncodingSession::stop fails partway.
        // unsafe {
        //     if let Err(e) = self.sink_writer.Finalize() {
        //          eprintln!("Error finalizing SinkWriter in Drop: {:?}", e);
        //     }
        // }
    }
}