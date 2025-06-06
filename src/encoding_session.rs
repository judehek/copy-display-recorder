use std::sync::{Arc, Mutex};
use windows::{
    core::{Result, HSTRING},
    Foundation::TimeSpan,
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Graphics::{
            Direct3D11::ID3D11Device,
            Gdi::HMONITOR,
        },
        Media::MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MF_SINK_WRITER_DISABLE_THROTTLING,
            MFCreateAttributes, MFCreateMFByteStreamOnStreamEx, MFCreateSinkWriterFromURL
        }, System::Performance::QueryPerformanceCounter,
    },
};

use crate::{
    audio::encoder_device::AudioEncoderDevice,
    audio::encoding_session::AudioEncodingSession,
    video::encoder_device::VideoEncoderDevice,
    video::encoding_session::VideoEncodingSession,
};

pub struct MediaEncodingSession {
    video_session: VideoEncodingSession,
    audio_session: AudioEncodingSession,
    sample_writer: Arc<Mutex<SampleWriter>>,
}

pub struct SampleWriter {
    _stream: IRandomAccessStream,
    sink_writer: IMFSinkWriter,
    video_stream_index: Option<u32>,
    audio_stream_index: Option<u32>,
}

unsafe impl Send for SampleWriter {}
unsafe impl Sync for SampleWriter {}
impl SampleWriter {
    pub fn new(stream: IRandomAccessStream) -> Result<Self> {
        let attributes = unsafe {
            let mut attributes = None;
            MFCreateAttributes(&mut attributes, 1)?;
            let attributes = attributes.unwrap();
            
            // Set the disable throttling attribute to TRUE
            attributes.SetUINT32(&MF_SINK_WRITER_DISABLE_THROTTLING, 1)?;
            
            attributes
        };
        
        let sink_writer = unsafe {
            let byte_stream = MFCreateMFByteStreamOnStreamEx(&stream)?;
            MFCreateSinkWriterFromURL(&HSTRING::from(".mp4"), &byte_stream, &attributes)?
        };

        Ok(Self {
            _stream: stream,
            sink_writer,
            video_stream_index: None,
            audio_stream_index: None,
        })
    }

    pub fn add_video_stream(&mut self, media_type: &IMFMediaType) -> Result<u32> {
        let stream_index = unsafe { self.sink_writer.AddStream(media_type)? };
        self.video_stream_index = Some(stream_index);
        Ok(stream_index)
    }

    pub fn add_audio_stream(&mut self, media_type: &IMFMediaType) -> Result<u32> {
        let stream_index = unsafe { self.sink_writer.AddStream(media_type)? };
        self.audio_stream_index = Some(stream_index);
        Ok(stream_index)
    }

    pub fn set_input_media_type(&self, stream_index: u32, media_type: &IMFMediaType) -> Result<()> {
        let empty_attributes = unsafe {
            let mut attributes = None;
            MFCreateAttributes(&mut attributes, 0)?;
            attributes.unwrap()
        };
        unsafe {
            self.sink_writer.SetInputMediaType(stream_index, media_type, &empty_attributes)
        }
    }

    pub fn start(&self) -> Result<()> {
        unsafe { self.sink_writer.BeginWriting() }
    }

    pub fn stop(&self) -> Result<()> {
        unsafe { self.sink_writer.Finalize() }
    }

    /// Write a video sample to the sink writer and release its buffers to avoid leaks.
    pub fn write_video_sample(&self, sample: &IMFSample) -> Result<()> {
        if let Some(stream_index) = self.video_stream_index {
            unsafe {
                // Write the sample to the sink writer
                self.sink_writer.WriteSample(stream_index, sample)?;
                // Remove all buffers from the sample to free associated memory
                sample.RemoveAllBuffers()?;
            }
            Ok(())
        } else {
            Err(windows::core::Error::from_win32())
        }
    }

    /// Write an audio sample to the sink writer and release its buffers to avoid leaks.
    pub fn write_audio_sample(&self, sample: &IMFSample) -> Result<()> {
        if let Some(stream_index) = self.audio_stream_index {
            unsafe {
                // Write the sample to the sink writer
                self.sink_writer.WriteSample(stream_index, sample)?;
                // Remove all buffers from the sample to free associated memory
                sample.RemoveAllBuffers()?;
            }
            Ok(())
        } else {
            Err(windows::core::Error::from_win32())
        }
    }
}

impl MediaEncodingSession {
    pub fn new(
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        video_encoder_device: &VideoEncoderDevice,
        audio_encoder_device: &AudioEncoderDevice,
        resolution: SizeInt32,
        video_bit_rate: u32,
        audio_bit_rate: u32,
        frame_rate: u32,
        stream: IRandomAccessStream,
    ) -> Result<Self> {
        // Create the shared sink writer
        let sample_writer = Arc::new(Mutex::new(SampleWriter::new(stream)?));
        
        // Create video session with shared sink writer
        let video_session = VideoEncodingSession::new(
            d3d_device.clone(),
            monitor_handle,
            video_encoder_device,
            resolution,
            video_bit_rate,
            frame_rate,
            sample_writer.clone(),
        )?;
        println!("created video encoder");
        
        // Create audio session with shared sink writer
        let audio_session = AudioEncodingSession::new(
            audio_encoder_device,
            audio_bit_rate,
            sample_writer.clone(),
        )?;
        
        Ok(Self {
            video_session,
            audio_session,
            sample_writer,
        })
    }
    
    pub fn start(&mut self) -> Result<()> {
        // Start the sink writer first
        self.sample_writer.lock().unwrap().start()?;

        let mut start_qpc = 0;
        unsafe { QueryPerformanceCounter(&mut start_qpc)? };
        println!("Obtained start QPC: {}", start_qpc);
        
        // Start both encoding sessions
        self.audio_session.start(start_qpc)?;
        self.video_session.start(start_qpc)?;
        
        Ok(())
    }
    
    pub fn stop(&mut self) -> Result<()> {
        // Stop both encoding sessions first
        self.video_session.stop()?;
        self.audio_session.stop()?;
        
        // Finally stop the sink writer
        self.sample_writer.lock().unwrap().stop()?;
        
        Ok(())
    }
}