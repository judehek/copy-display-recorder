use std::{sync::Arc, time::{SystemTime, UNIX_EPOCH}};

use windows::{
    core::{imp::CoTaskMemFree, ComInterface, Result, HSTRING},
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
        Media::{Audio::{IAudioClient, IMMDevice}, MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFAudioFormat_PCM, MFCreateAttributes, MFCreateMFByteStreamOnStreamEx, MFCreateSinkWriterFromURL
        }},
        System::{Com::CLSCTX_ALL, Performance::QueryPerformanceFrequency}
    },
};

use crate::{audio::capture_audio::{CaptureAudioGenerator}, encoding_session::SampleWriter};

use super::{
    capture_audio::AudioCaptureSession, capture_microphone::{CaptureMicrophoneGenerator, MicrophoneCaptureSession}, encoder::{AudioEncoder, AudioEncoderInputSample}, encoder_device::{AudioEncoderDevice}, processor::{AudioFormat, AudioProcessor}
};

#[derive(Clone)]
pub enum AudioSource {
    Desktop,
    ActiveWindow,
}

pub struct AudioEncodingSession {
    audio_encoder: AudioEncoder,
    audio_capture_session: Option<AudioCaptureSession>,
    microphone_capture_session: Option<MicrophoneCaptureSession>,
}

struct SampleGenerator {
    audio_processor: Option<AudioProcessor>,
    microphone_processor: Option<AudioProcessor>,

    audio_generator: Option<CaptureAudioGenerator>,
    microphone_generator: Option<CaptureMicrophoneGenerator>,

    seen_first_time_stamp: bool,
    first_timestamp: i64,

    frame_period: i64,
    next_frame_time: TimeSpan,
    total_samples_processed: i64,
}

impl AudioEncodingSession {
    pub fn new(
        encoder_device: &AudioEncoderDevice,
        bit_rate: u32,
        sample_writer: Arc<SampleWriter>,
    ) -> Result<Self> {
        let mut audio_encoder = AudioEncoder::new(
            &encoder_device,
48000,
            2,
            8,
        )?;

        let output_format = AudioFormat {
            sample_rate: 48000,    // 48kHz is standard for professional audio
            channels: 2,           // Stereo output
            bits_per_sample: 32,   // 16-bit PCM is widely compatible
            channel_mask: Some(3), // SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT (0x1 | 0x2)
            format: MFAudioFormat_PCM, // PCM format for compatibility
        };

        let mut sample_generator = SampleGenerator::new(
            true,
            AudioSource::Desktop,
            false, 
            None,
            output_format,
            None,
        )?;
        println!("created sample gen");
        let audio_capture_session = sample_generator.audio_capture_session().clone();
        let microphone_capture_session = sample_generator.microphone_capture_session().clone();
        audio_encoder.set_sample_requested_callback(
            move || -> Result<Option<AudioEncoderInputSample>> { sample_generator.generate() },
        );

        audio_encoder.set_sample_rendered_callback({
            let sample_writer = sample_writer.clone();
            move |sample| -> Result<()> { sample_writer.write_audio_sample(sample.sample()) }
        });

        Ok(Self {
            audio_encoder,
            audio_capture_session,
            microphone_capture_session
        })
    }

    pub fn start(&mut self) -> Result<()> {
        if let Some(session) = &mut self.audio_capture_session {
            session.StartCapture()?;
        }
        if let Some(session) = &mut self.microphone_capture_session {
           session.StartCapture()?;
       }
        assert!(self.audio_encoder.try_start()?);
        Ok(())
    }

    pub fn stop(&mut self) -> Result<()> {
        self.audio_encoder.stop()?;
        Ok(())
    }
}

unsafe impl Send for SampleGenerator {}
impl SampleGenerator {
    pub fn new(
        capture_audio: bool,
        audio_source: AudioSource,
        capture_microphone: bool,
        microphone_device: Option<IMMDevice>,
        output_format: AudioFormat,
        quality: Option<u32>,
    ) -> Result<Self> {
        // Initialize variables to be used in conditionals
        let mut audio_generator = None;
        let mut audio_processor = None;
        let mut microphone_generator = None;
        let mut microphone_processor = None;

        if capture_audio {
            // Create the audio generator
            let mut temp_audio_generator = CaptureAudioGenerator::new(audio_source)?;
            println!("created capture audio gen");
            // Start capture and wait for initialization with 500ms timeout
            temp_audio_generator.start_capture_and_wait(500)?;
            
            // Create the audio processor with the audio format from the generator
            let audio_input_format = AudioFormat {
                sample_rate: temp_audio_generator.get_sample_rate(),
                channels: temp_audio_generator.get_channels(),
                bits_per_sample: temp_audio_generator.get_bits_per_sample(),
                channel_mask: None,
                format: MFAudioFormat_PCM,
            };
            
            audio_generator = Some(temp_audio_generator);
            audio_processor = Some(AudioProcessor::new(
                audio_input_format,
                output_format.clone(),
                quality,
            )?);
        }

        /*if capture_microphone {
            if let Some(device) = &microphone_device {
                // Create the microphone generator
                let mut temp_microphone_generator = CaptureMicrophoneGenerator::new(device.clone())?;
                
                // Start capture and wait for initialization with 500ms timeout
                temp_microphone_generator.start_capture_and_wait(500)?;
                
                // Create the audio processor with the microphone format from the generator
                let microphone_input_format = AudioFormat {
                    sample_rate: temp_microphone_generator.get_sample_rate(),
                    channels: temp_microphone_generator.get_channels(),
                    bits_per_sample: temp_microphone_generator.get_bits_per_sample(),
                };
                
                microphone_generator = Some(temp_microphone_generator);
                microphone_processor = Some(AudioProcessor::new(
                    microphone_input_format,
                    output_format.clone(),
                    quality,
                )?);
            }
        }*/

        Ok(Self {
            audio_generator,
            audio_processor,
            microphone_generator,
            microphone_processor,

            seen_first_time_stamp: false,
            first_timestamp: 0,
        
            frame_period: 0,
            next_frame_time: TimeSpan::default(),
            total_samples_processed: 0,
        })
    }

    pub fn audio_capture_session(&self) -> Option<AudioCaptureSession> {
        self.audio_generator.as_ref().map(|gen| gen.session().clone())
    }

    pub fn microphone_capture_session(&self) -> Option<MicrophoneCaptureSession> {
        self.microphone_generator.as_ref().map(|gen| gen.session().clone())
    }

    pub fn generate(&mut self) -> Result<Option<AudioEncoderInputSample>> {
        // Define buffer size based on your requirements
        let max_samples = 1024; // Adjust based on your needs
        let mut audio_buffer = vec![0.0f32; max_samples];
        let mut mic_buffer = vec![0.0f32; max_samples];
        
        // 1. Try get both audio sources
        let mut audio_samples = 0;
        let mut mic_samples = 0;
        
        // Get audio samples if available
        if let Some(generator) = &mut self.audio_generator {
            if let Some(count) = generator.try_get_audio_samples(&mut audio_buffer, max_samples) {
                audio_samples = count;
            }
        }
        
        // Get microphone samples if available
        if let Some(generator) = &mut self.microphone_generator {
            if let Some(count) = generator.try_get_audio_samples(&mut mic_buffer, max_samples) {
                mic_samples = count;
            }
        }
        
        // If no samples available from either source, end capture
        if audio_samples == 0 && mic_samples == 0 {
            self.stop_capture()?;
            return Ok(None);
        }
        
        // Calculate current timestamp and duration based on sample position and rate
        let sample_rate = self.audio_generator
            .as_ref()
            .map(|g| g.get_sample_rate())
            .unwrap_or(48000) as f64;
            
        let current_time = if !self.seen_first_time_stamp {
            // Initialize timing on first batch of samples
            self.first_timestamp = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos() as i64;
            self.seen_first_time_stamp = true;
            
            TimeSpan { Duration: 0 }
        } else {
            // For subsequent samples, calculate relative time based on sample position
            let time_in_seconds = self.total_samples_processed as f64 / sample_rate;
            let duration_100ns = (time_in_seconds * 10_000_000.0) as i64;
            
            TimeSpan { Duration: duration_100ns }
        };
        
        // Calculate duration for this batch of samples (critical for AAC encoding)
        let samples_this_batch = audio_samples.max(mic_samples) as f64;
        let duration_in_seconds = samples_this_batch / sample_rate;
        let duration_100ns = (duration_in_seconds * 10_000_000.0) as i64;
        let duration = TimeSpan { Duration: duration_100ns };
        
        // Convert float samples to bytes for IMFSample creation - moved out of nested blocks
        // to avoid borrowing conflicts
        let audio_bytes = if audio_samples > 0 {
            self.float_to_bytes(&audio_buffer[0..audio_samples])
        } else {
            Vec::new()
        };
        
        let mic_bytes = if mic_samples > 0 {
            self.float_to_bytes(&mic_buffer[0..mic_samples])
        } else {
            Vec::new()
        };
        
        // 2. Process both audio sources
        let processed_audio = if audio_samples > 0 {
            if let Some(processor) = &mut self.audio_processor {
                // Create an IMFSample through from_raw
                let audio_sample = AudioEncoderInputSample::from_raw(
                    &audio_bytes,
                    current_time,
                    duration
                )?;
                
                // Process the sample and extract resulting IMFSample
                let processed_sample = processor.process_sample(&audio_sample.sample())?;
                
                processed_sample
            } else {
                None
            }
        } else {
            None
        };
        
        let processed_mic = if mic_samples > 0 {
            if let Some(processor) = &mut self.microphone_processor {
                // Create an IMFSample through from_raw
                let mic_sample = AudioEncoderInputSample::from_raw(
                    &mic_bytes,
                    current_time,
                    duration
                )?;
                
                // Process the sample and extract resulting IMFSample
                let processed_sample = processor.process_sample(&mic_sample.sample())?;
                
                processed_sample
            } else {
                None
            }
        } else {
            None
        };
        
        // Track total samples processed for timing
        self.total_samples_processed += audio_samples.max(mic_samples) as i64;
        
        // 3. Mix if needed
        if processed_audio.is_some() && processed_mic.is_some() {
            // Both sources available, mix them
            return Ok(Some(self.mix_audio_samples(
                processed_audio.unwrap(),
                processed_mic.unwrap(),
                current_time,
                duration
            )?));
        } else if let Some(audio) = processed_audio {
            // Only system audio available
            return Ok(Some(AudioEncoderInputSample::new(audio)));
        } else if let Some(mic) = processed_mic {
            // Only microphone audio available
            return Ok(Some(AudioEncoderInputSample::new(mic)));
        }
        
        // Should never reach here given the earlier check
        Ok(None)
    }
    
    fn stop_capture(&mut self) -> Result<()> {
        // Stop audio capture if it exists
        if let Some(generator) = &mut self.audio_generator {
            generator.stop_capture()?;
        }
        
        // Stop microphone capture if it exists
        if let Some(generator) = &mut self.microphone_generator {
            generator.stop_capture()?;
        }
        
        Ok(())
    }
    
    // Helper to convert float samples to bytes
    fn float_to_bytes(&self, samples: &[f32]) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(samples.len() * 4);
        for sample in samples {
            let sample_bytes = sample.to_le_bytes();
            bytes.extend_from_slice(&sample_bytes);
        }
        bytes
    }
    
    fn mix_audio_samples(
        &mut self,
        audio_sample: IMFSample,
        mic_sample: IMFSample,
        timestamp: TimeSpan,
        duration: TimeSpan
    ) -> Result<AudioEncoderInputSample> {
        // Extract audio data from both samples
        let (audio_data, audio_len) = self.extract_float_data_from_sample(&audio_sample)?;
        let (mic_data, mic_len) = self.extract_float_data_from_sample(&mic_sample)?;
        
        // Create a new buffer for the mixed audio
        let mix_len = audio_len.min(mic_len);
        let mut mixed_data = vec![0.0f32; mix_len];
        
        // Apply mixing with level control
        let audio_level = 0.7;
        let mic_level = 0.3;
        
        // Mix the data
        for i in 0..mix_len {
            mixed_data[i] = audio_data[i] * audio_level + mic_data[i] * mic_level;
        }
        
        // Convert mixed float data to bytes
        let mixed_bytes = self.float_to_bytes(&mixed_data);
        
        // Create the final sample using from_raw
        AudioEncoderInputSample::from_raw(
            &mixed_bytes,
            timestamp,
            duration
        )
    }
    
    // Helper to extract float data from an IMFSample
    fn extract_float_data_from_sample(&self, sample: &IMFSample) -> Result<(Vec<f32>, usize)> {
        unsafe {
            // Get the first buffer
            let buffer = sample.GetBufferByIndex(0)?;
            
            // Lock the buffer to read data
            let mut data_ptr = std::ptr::null_mut();
            let mut max_length = 0;
            let mut current_length = 0;
            buffer.Lock(&mut data_ptr, Some(&mut max_length), Some(&mut current_length))?;
            
            // Calculate how many float samples we have
            let sample_count = current_length as usize / std::mem::size_of::<f32>();
            
            // Create a vector to hold the samples
            let mut samples = vec![0.0f32; sample_count];
            
            // Copy the data
            std::ptr::copy_nonoverlapping(
                data_ptr as *const f32,
                samples.as_mut_ptr(),
                sample_count
            );
            
            // Unlock the buffer
            buffer.Unlock()?;
            
            Ok((samples, sample_count))
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