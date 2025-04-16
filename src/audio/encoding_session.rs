use std::{sync::{Arc, Barrier, Mutex}, time::{SystemTime, UNIX_EPOCH}};

use windows::{
    core::{imp::CoTaskMemFree, Interface, Result, HSTRING},
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
        Media::{Audio::{IAudioClient, IMMDevice}, KernelStreaming::KS_TRUECOLORINFO, MediaFoundation::{
            IMFMediaType, IMFSample, IMFSinkWriter, MFAudioFormat_AAC, MFAudioFormat_Float, MFAudioFormat_PCM, MFCreateAttributes, MFCreateMFByteStreamOnStreamEx, MFCreateSinkWriterFromURL
        }},
        System::{Com::CLSCTX_ALL, Performance::QueryPerformanceFrequency}
    },
};

use crate::{audio::capture_audio::{CaptureAudioGenerator}, encoding_session::SampleWriter};

use super::{
    capture_audio::{AudioCaptureSession, AudioSample}, capture_microphone::{CaptureMicrophoneGenerator, MicrophoneCaptureSession}, encoder::{AudioEncoder, AudioEncoderInputSample}, encoder_device::AudioEncoderDevice, processor::{AudioFormat, AudioProcessor}
};

#[derive(Clone)]
pub enum AudioSource {
    Desktop,
    ActiveWindow,
}

pub struct AudioEncodingSession {
    audio_capture_session: Option<AudioCaptureSession>,
    microphone_capture_session: Option<MicrophoneCaptureSession>,
    stop_signal: Arc<std::sync::atomic::AtomicBool>,
    start_barrier: Arc<Barrier>,
    processing_thread: Option<std::thread::JoinHandle<()>>,
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
        sample_writer: Arc<Mutex<SampleWriter>>,
    ) -> Result<Self> {
        // Your existing format setup code remains the same
        let output_format = AudioFormat {
            sample_rate: 48000,
            channels: 2,
            bits_per_sample: 16,
            channel_mask: Some(0x3),
            format: MFAudioFormat_AAC,
        };
    
        let capture_format = AudioFormat {
            sample_rate: 48000,
            channels: 2,
            bits_per_sample: 16,
            channel_mask: Some(3),
            format: MFAudioFormat_PCM,
        };
    
        // Create the sample generator outside the thread as it's still needed
        let sample_generator = SampleGenerator::new(
            true,
            AudioSource::Desktop,
            false,
            None,
            output_format.clone(),
            None,
        )?;
        
        // Store references to capture sessions
        let audio_capture_session = sample_generator.audio_capture_session().clone();
        let microphone_capture_session = sample_generator.microphone_capture_session().clone();
    
        // Only wrap the sample generator in a thread-safe wrapper
        let sample_generator = Arc::new(std::sync::Mutex::new(sample_generator));
        
        // Clone for thread
        let sample_generator_thread = sample_generator.clone();
        
        // Capture encoder configuration for the thread
        let encoder_device_clone = encoder_device.clone();
        let output_format_clone = output_format.clone();
        let capture_format_clone = capture_format.clone();
        
        // Create a barrier for 2 threads: the main thread and the worker thread
        let start_barrier = Arc::new(Barrier::new(2));
        let start_barrier_thread = start_barrier.clone();

        // Use a separate signal for stopping
        let stop_signal = Arc::new(std::sync::atomic::AtomicBool::new(false));
        let stop_signal_thread = stop_signal.clone();
        
        // Create the processing thread
        let processing_thread = std::thread::spawn(move || {
            // Create the audio encoder inside the thread
            let mut audio_encoder = match AudioEncoder::new(
                &encoder_device_clone,
                capture_format_clone,
                output_format_clone,
                None
            ) {
                Ok(encoder) => encoder,
                Err(e) => {
                    eprintln!("Failed to create audio encoder in thread: {:?}", e);
                    return; // Exit thread if encoder creation fails
                }
            };
            let _ = sample_writer.lock().unwrap().add_audio_stream(audio_encoder.output_media_type());
            println!("created audio encoder");

            println!("Audio thread waiting on barrier...");
            start_barrier_thread.wait();
            println!("Audio thread proceeding past barrier.");
            
            while !stop_signal_thread.load(std::sync::atomic::Ordering::Relaxed) {
                // Try to get the next sample
                if let Ok(mut generator) = sample_generator_thread.lock() {
                    match generator.generate() {
                        Ok(Some(sample)) => {
                            // Process the sample with the encoder (no mutex needed now)
                            match audio_encoder.process_sample(&sample) {
                                Ok(Some(encoded_sample)) => {
                                    // Write the encoded sample and remove buffers
                                    {
                                        let writer = sample_writer.lock().unwrap();
                                        if let Err(e) = writer.write_audio_sample(encoded_sample.sample()) {
                                            eprintln!("Error writing audio sample: {:?}", e);
                                        }
                                    }
                                    // Explicitly drop the sample to force COM Release
                                    drop(encoded_sample);
                                },
                                Ok(None) => {
                                    // No encoded sample was produced, perhaps buffering
                                    // This is normal for some encoders
                                },
                                Err(e) => eprintln!("Error processing audio sample: {:?}", e),
                            }
                        },
                        Ok(None) => {
                            // No sample available, sleep briefly to avoid busy-waiting
                            std::thread::sleep(std::time::Duration::from_millis(5));
                        },
                        Err(e) => eprintln!("Error generating audio sample: {:?}", e),
                    }
                }
            }
            
            // Drain any buffered samples when stopping
            match audio_encoder.drain() {
                Ok(encoded_samples) => {
                    // Write any remaining encoded samples
                    for encoded_sample in encoded_samples {
                        // Write the drained encoded sample and remove buffers
                        {
                            let writer = sample_writer.lock().unwrap();
                            if let Err(e) = writer.write_audio_sample(encoded_sample.sample()) {
                                eprintln!("Error writing drained audio sample: {:?}", e);
                            }
                        }
                        // Explicitly drop the sample to force COM Release
                        drop(encoded_sample);
                    }
                },
                Err(e) => eprintln!("Error draining audio encoder: {:?}", e),
            }
        });
        
        Ok(Self {
            audio_capture_session,
            microphone_capture_session,
            stop_signal, // Store the stop signal
            start_barrier, // Store the barrier
            processing_thread: Some(processing_thread),
        })
    }

    pub fn start(&mut self, start_qpc: i64) -> Result<()> {
        // Start the capture sessions
        if let Some(session) = &mut self.audio_capture_session {
            session.StartCapture(start_qpc)?;
        }
        if let Some(session) = &mut self.microphone_capture_session {
            session.StartCapture(start_qpc)?;
        }

        // Signal the processing thread to start its loop by waiting on the barrier
        // This call will block until the worker thread also calls wait().
        println!("Main thread waiting on barrier...");
        self.start_barrier.wait();
        println!("Main thread proceeding past barrier.");

        // Note: The stop_signal remains false here.

        Ok(())
    }

    pub fn stop(&mut self) -> Result<()> {
        // Set the stop signal to true to stop the processing thread's loop
        self.stop_signal.store(true, std::sync::atomic::Ordering::Relaxed);
        println!("Stop signal sent.");

        // Wait for the processing thread to finish
        if let Some(thread) = self.processing_thread.take() {
            println!("Joining audio processing thread...");
            if let Err(e) = thread.join() {
                eprintln!("Error joining processing thread: {:?}", e);
            }
            println!("Audio processing thread joined.");
        }

        // Stop the capture sessions
        if let Some(session) = &mut self.audio_capture_session {
            session.StopCapture()?;
        }
        if let Some(session) = &mut self.microphone_capture_session {
            session.StopCapture()?;
        }

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
            audio_generator = Some(CaptureAudioGenerator::new(audio_source)?);
            // Start capture and wait for initialization with 500ms timeout
            /*audio_processor = Some(AudioProcessor::new(
                audio_input_format,
                output_format.clone(),
                quality,
            )?);*/
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
        // Try to get audio samples from both sources
        let audio_sample = if let Some(generator) = &mut self.audio_generator {
            generator.try_get_audio_sample()
        } else {
            None
        };
        
        /*let mic_sample = if let Some(generator) = &mut self.microphone_generator {
            generator.try_get_audio_sample()
        } else {
            None
        };*/
        
        // Check if we have both audio and mic samples
        /*if let (Some(audio), Some(mic)) = (&audio_sample, &mic_sample) {
            // Both sources available, mix them
            println!("we have both");
            return Ok(Some(self.mix_audio_samples(audio.clone(), mic.clone())?));
        }*/
        
        // If we didn't have both, check for individual sources
        if let Some(audio) = audio_sample {
            // Only system audio available
            return Ok(Some(self.convert_to_encoder_input(audio)?));
        }
        
        /*if let Some(mic) = mic_sample {
            println!("we have mic");
            // Only microphone audio available
            return Ok(Some(self.convert_to_encoder_input(mic)?));
        }*/
        
        // No samples available
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
    
    // Helper to convert float samples to bytes - still needed for mixing
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
        audio_sample: &AudioSample,
        mic_sample: &AudioSample
    ) -> Result<AudioEncoderInputSample> {
        // Extract audio data from both samples
        let audio_data = self.extract_float_data_from_audio_sample(&audio_sample);
        let mic_data = self.extract_float_data_from_audio_sample(&mic_sample);
        
        // Create a new buffer for the mixed audio
        let mix_len = audio_data.len().min(mic_data.len());
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
        
        // Use timestamp and duration from audio_sample for the mixed result
        // (We could also choose the earlier timestamp if they differ, but
        // they should be synchronized in most cases)
        Ok(AudioEncoderInputSample::new(
            mixed_bytes,
            audio_sample.timestamp,
            audio_sample.duration,
            0
        ))
    }
    
    // Helper to extract float data from an AudioSample
    fn extract_float_data_from_audio_sample(&self, sample: &AudioSample) -> Vec<f32> {
        // Calculate how many float samples we have
        let bytes_per_sample = 4; // 32-bit float
        let sample_count = sample.data.len() / bytes_per_sample;
        
        // Create a vector to hold the samples
        let mut samples = vec![0.0f32; sample_count];
        
        // Convert the bytes to floats
        for i in 0..sample_count {
            let start = i * bytes_per_sample;
            let mut byte_array = [0u8; 4];
            byte_array.copy_from_slice(&sample.data[start..start+4]);
            samples[i] = f32::from_le_bytes(byte_array);
        }
        
        samples
    }
    
    // Helper to convert AudioSample to AudioEncoderInputSample
    fn convert_to_encoder_input(&self, audio_sample: AudioSample) -> Result<AudioEncoderInputSample> {
        Ok(AudioEncoderInputSample::new(
            audio_sample.data,
            audio_sample.timestamp,
            audio_sample.duration,
            0
        ))
    }
    
}