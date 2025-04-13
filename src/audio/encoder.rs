use windows::{
    core::{ComInterface, Interface, Result, GUID, HRESULT}, Foundation::TimeSpan, Win32::{
        Foundation::S_OK,
        Media::MediaFoundation::{
            // Interfaces
            IMFMediaBuffer, IMFMediaType, IMFSample, IMFTransform, MFAudioFormat_AAC, MFAudioFormat_Float, MFAudioFormat_PCM, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFMediaType_Audio, MFVideoInterlace_Progressive, MFT_OUTPUT_DATA_BUFFER, MFT_OUTPUT_STREAM_INFO, MF_E_TRANSFORM_NEED_MORE_INPUT, MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION, MF_MT_AAC_PAYLOAD_TYPE, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_CHANNEL_MASK, MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_INTERLACE_MODE, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE
        },
        System::{
            Com::StructuredStorage::PROPVARIANT,
            Variant::VT_UI4,
        },
        UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY},
    }
};

#[derive(Clone)]
pub struct AudioEncoderInputSample {
    pub data: Vec<u8>,
    pub timestamp: TimeSpan,
    pub duration: TimeSpan,
    pub frames: u32,
}

impl AudioEncoderInputSample {
    pub fn new(data: Vec<u8>, timestamp: TimeSpan, duration: TimeSpan, frames: u32) -> Self {
        Self { data, timestamp, duration, frames }
    }
}

pub struct AudioEncoderOutputSample {
    sample: IMFSample,
}

impl AudioEncoderOutputSample {
    pub fn sample(&self) -> &IMFSample {
        &self.sample
    }
}

use std::mem::{ManuallyDrop};

use crate::video::encoder;

use super::{encoder_device::AudioEncoderDevice, processor::AudioFormat};

pub struct AudioEncoder {
    encoder_transform: IMFTransform,
    input_media_type: IMFMediaType,
    output_media_type: IMFMediaType,
    input_stream_id: u32,
    output_stream_id: u32,
    output_buffer_size: u32,
}

impl AudioEncoder {
    pub fn new(
        encoder_device: &AudioEncoderDevice,
        input_format: AudioFormat,
        output_format: AudioFormat, 
        bitrate: Option<u32>,
    ) -> Result<Self> {
        // Create the encoder transform using the provided device
        let encoder_transform = encoder_device.create_transform()?;
        
        // Create Media Types
        let input_media_type = create_audio_media_type(&input_format)?;
        
        // For output, we need to set encoder-specific attributes for AAC
        let output_media_type = create_aac_output_media_type(&output_format, bitrate)?;
        
        // Set Media Types on the Transform
        unsafe {
            encoder_transform.SetInputType(0, Some(&input_media_type), 0)?; // Stream ID 0, no flags
            encoder_transform.SetOutputType(0, Some(&output_media_type), 0)?; // Stream ID 0, no flags
        }
        
        // Get Stream IDs
        let mut input_ids = [0u32; 1]; // Expecting one input stream
        let mut output_ids = [0u32; 1]; // Expecting one output stream
        let input_stream_id;
        let output_stream_id;
        
        unsafe {
            // Note: GetStreamIDs might not be implemented by all MFTs. Defaulting to 0 is common.
            match encoder_transform.GetStreamIDs(&mut input_ids[..], &mut output_ids[..]) {
                Ok(_) => {
                    input_stream_id = input_ids[0];
                    output_stream_id = output_ids[0];
                },
                Err(e) => {
                    // Don't fail completely, try assuming 0, but warn
                    println!("Warning: IMFTransform::GetStreamIDs failed ({:?}), assuming stream IDs are 0.", e.code());
                    input_stream_id = 0;
                    output_stream_id = 0;
                }
            }
        }
        
        // Get output stream info to estimate buffer size needed
        let stream_info = unsafe { encoder_transform.GetOutputStreamInfo(output_stream_id)? };
        // Ensure we have a sane minimum if cbSize is 0 (which can happen)
        let output_buffer_size = if stream_info.cbSize > 0 { stream_info.cbSize } else { 32768 }; // 32KB min for encoded audio
        
        Ok(Self {
            encoder_transform,
            input_media_type,
            output_media_type,
            input_stream_id,
            output_stream_id,
            output_buffer_size,
        })
    }
    
    /// Processes a single input audio sample and returns the corresponding output sample(s).
    /// For encoders, one input might not immediately produce an output due to buffering.
    pub fn process_sample(&mut self, input_sample: &AudioEncoderInputSample) -> Result<Option<AudioEncoderOutputSample>> {
        unsafe {
            // Create an MF sample from the input sample
            let input_mf_sample = MFCreateSample()?;
            
            // Create a buffer for the input data
            let input_buffer = MFCreateMemoryBuffer(input_sample.data.len() as u32)?;
            
            // Get the buffer and copy the data into it
            let mut buffer_data: *mut u8 = std::ptr::null_mut();
            let mut max_length: u32 = 0;
            input_buffer.Lock(&mut buffer_data, Some(&mut max_length), None)?;
            
            // Copy the data
            std::ptr::copy_nonoverlapping(
                input_sample.data.as_ptr(),
                buffer_data,
                input_sample.data.len()
            );
            
            // Set the current length and unlock
            input_buffer.SetCurrentLength(input_sample.data.len() as u32)?;
            input_buffer.Unlock()?;
            
            // Add the buffer to the sample
            input_mf_sample.AddBuffer(&input_buffer)?;
            
            // Set the sample attributes
            input_mf_sample.SetSampleTime(input_sample.timestamp.Duration)?;
            input_mf_sample.SetSampleDuration(input_sample.duration.Duration)?;
    
            // 1. Send input to the transform
            match self.encoder_transform.ProcessInput(self.input_stream_id, &input_mf_sample, 0) {
                Ok(_) => {}, // Input accepted
                Err(e) => {
                    // Handle specific errors if needed, e.g., MF_E_NOTACCEPTING
                    println!("ProcessInput failed: {:?}", e);
                    return Err(e.into());
                }
            }
            
            // 2. Try to get output - note that encoders often need multiple inputs before producing output
            // Create necessary structures for ProcessOutput
            
            // Try to create a buffer of the suggested size - encoders often need larger buffers
            let buffer_size = self.output_buffer_size;
            let output_buffer = MFCreateMemoryBuffer(buffer_size)?;
            
            let output_sample = MFCreateSample()?;
            output_sample.AddBuffer(&output_buffer)?;
            
            let mut output_data_buffer = MFT_OUTPUT_DATA_BUFFER {
                dwStreamID: self.output_stream_id,
                pSample: ManuallyDrop::new(Some(output_sample)),
                dwStatus: 0, // Will be filled by ProcessOutput
                pEvents: ManuallyDrop::new(None), // We don't process events here
            };
            let mut process_output_status: u32 = 0;
            
            // Create a slice as expected by ProcessOutput
            let mut output_buffers = [output_data_buffer]; // Array of size 1
            
            match self.encoder_transform.ProcessOutput(
                0, // Flags
                &mut output_buffers, // Array of buffers
                &mut process_output_status, // Status flags
            ) {
                Ok(_) => {
                    // Success! We got an output sample.
                    let filled_sample_option = ManuallyDrop::take(&mut output_buffers[0].pSample);
                    
                    // Ensure the MFT actually provided a sample
                    let processed_sample = filled_sample_option.ok_or_else(|| {
                        println!("Error: ProcessOutput succeeded but returned NULL sample pointer.");
                        windows::core::Error::new(HRESULT(0x8000FFFFu32 as i32), "ProcessOutput succeeded but returned null sample".into()) // E_UNEXPECTED
                    })?;
                    
                    // Update buffer length (important!)
                    let filled_buffer: IMFMediaBuffer = processed_sample.GetBufferByIndex(0)?;
                    let current_length = filled_buffer.GetCurrentLength()?;
                    filled_buffer.SetCurrentLength(current_length)?; // Ensure length is set
                    
                    // Wrap in our output type
                    let output = AudioEncoderOutputSample {
                        sample: processed_sample
                    };
                    
                    return Ok(Some(output));
                }
                Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
                    // Transform needs more input data before it can produce output.
                    // This is common for encoders that buffer frames.
                    let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                    drop(sample_to_drop);
                    
                    let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                    drop(events_to_drop);
                    
                    return Ok(None);
                }
                Err(e) => {
                    // Any other error
                    println!("ProcessOutput failed: {:?}", e);
                    let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                    drop(sample_to_drop);
                    
                    let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                    drop(events_to_drop);
                    
                    return Err(e.into());
                }
            }
        }
    }

    /// Signal that no more input will be provided
    /// This is important for encoders to flush their buffers
    pub fn drain(&mut self) -> Result<Vec<AudioEncoderOutputSample>> {
        let mut result_samples = Vec::new();
        
        unsafe {
            // First, notify the encoder we're done with input
            match self.encoder_transform.ProcessMessage(
                windows::Win32::Media::MediaFoundation::MFT_MESSAGE_TYPE(
                    windows::Win32::Media::MediaFoundation::MFT_MESSAGE_COMMAND_DRAIN.0
                ), 
                0
            ) {
                Ok(_) => {}, // Successfully set drain mode
                Err(e) => {
                    println!("Warning: Failed to set drain mode: {:?}", e);
                    // We can still try to get remaining samples
                }
            }
            
            // Now try to get all remaining output samples
            loop {
                // Create buffer for output
                let buffer_size = self.output_buffer_size;
                let output_buffer = MFCreateMemoryBuffer(buffer_size)?;
                
                let output_sample = MFCreateSample()?;
                output_sample.AddBuffer(&output_buffer)?;
                
                let mut output_data_buffer = MFT_OUTPUT_DATA_BUFFER {
                    dwStreamID: self.output_stream_id,
                    pSample: ManuallyDrop::new(Some(output_sample)),
                    dwStatus: 0,
                    pEvents: ManuallyDrop::new(None),
                };
                let mut process_output_status: u32 = 0;
                
                let mut output_buffers = [output_data_buffer];
                
                match self.encoder_transform.ProcessOutput(
                    0,
                    &mut output_buffers,
                    &mut process_output_status,
                ) {
                    Ok(_) => {
                        // Got another output sample during drain
                        let filled_sample_option = ManuallyDrop::take(&mut output_buffers[0].pSample);
                        
                        if let Some(processed_sample) = filled_sample_option {
                            // Update buffer length
                            let filled_buffer: IMFMediaBuffer = processed_sample.GetBufferByIndex(0)?;
                            let current_length = filled_buffer.GetCurrentLength()?;
                            filled_buffer.SetCurrentLength(current_length)?;
                            
                            // Wrap in our output type
                            let output = AudioEncoderOutputSample {
                                sample: processed_sample
                            };
                            
                            result_samples.push(output);
                        } else {
                            // No more samples but ProcessOutput succeeded - unusual
                            println!("Warning: ProcessOutput during drain succeeded but returned NULL sample");
                            break;
                        }
                    }
                    Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
                        // No more output can be produced, we're done draining
                        let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                        drop(sample_to_drop);
                        
                        let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                        drop(events_to_drop);
                        
                        break;
                    }
                    Err(e) => {
                        // Error during drain
                        println!("ProcessOutput during drain failed: {:?}", e);
                        let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                        drop(sample_to_drop);
                        
                        let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                        drop(events_to_drop);
                        
                        // We still return any samples we got before the error
                        break;
                    }
                }
            }
        }
        
        Ok(result_samples)
    }
    
    // Provide access to the configured output media type
    pub fn output_media_type(&self) -> &IMFMediaType {
        &self.output_media_type
    }
}

// Helper function to create an IMFMediaType for uncompressed audio input
fn create_audio_media_type(format: &AudioFormat) -> Result<IMFMediaType> {
    unsafe {
        let media_type = MFCreateMediaType()?;
        
        media_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
        media_type.SetGUID(&MF_MT_SUBTYPE, &format.format)?;
        media_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, format.sample_rate)?;
        media_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, format.channels as u32)?;
        media_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, format.bits_per_sample as u32)?;
        media_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, format.block_align() as u32)?;
        media_type.SetUINT32(
            &MF_MT_AUDIO_AVG_BYTES_PER_SECOND,
            format.avg_bytes_per_second(),
        )?;
        
        // Set interlace mode to progressive (standard for audio)
        media_type.SetUINT32(&MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive.0 as u32)?;
        
        // Required for some transforms/sinks, especially multichannel
        const SPEAKER_FRONT_LEFT: u32 = 0x1;
        const SPEAKER_FRONT_RIGHT: u32 = 0x2;
        const SPEAKER_FRONT_CENTER: u32 = 0x4;
        
        let mut mask_to_set: Option<u32> = format.channel_mask;
        
        if mask_to_set.is_none() {
            // Provide default masks for mono/stereo if not specified
            match format.channels {
                1 => mask_to_set = Some(SPEAKER_FRONT_CENTER), // Standard mono
                2 => mask_to_set = Some(SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT), // Standard stereo
                _ => {
                    println!("Warning: Creating audio media type for {} channels without an explicit channel mask.", format.channels);
                }
            };
        }
        
        if let Some(mask) = mask_to_set {
            if mask != 0 {
                media_type.SetUINT32(&MF_MT_AUDIO_CHANNEL_MASK, mask)?;
            } else if format.channels > 0 {
                println!("Warning: Audio channel mask is 0 for {} channels. This might be invalid.", format.channels);
            }
        }
        
        Ok(media_type)
    }
}

fn create_aac_output_media_type(
    format: &AudioFormat, 
    bitrate: Option<u32>
) -> Result<IMFMediaType> {
    unsafe {
        let media_type = MFCreateMediaType()?;
        
        // Set major type to Audio
        media_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
        
        // Set subtype to AAC
        media_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
        
        // Set basic audio properties
        media_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, format.sample_rate)?;
        media_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, format.channels as u32)?;
        
        // AAC usually supports 16-bit samples, set this explicitly
        media_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?;
        
        // Calculate and set bitrate
        let bitrate_value = bitrate.unwrap_or_else(|| {
            // Default bitrates based on sample rate and channels if not specified
            // These are reasonable defaults for AAC
            match (format.sample_rate, format.channels) {
                (sr, ch) if sr >= 48000 && ch >= 2 => 192000, // High quality stereo
                (sr, ch) if sr >= 44100 && ch >= 2 => 128000, // CD quality stereo
                (sr, ch) if sr >= 44100 && ch == 1 => 96000,  // CD quality mono
                (sr, _) if sr >= 32000 => 80000,              // Medium quality
                (sr, _) if sr >= 24000 => 64000,              // Lower quality
                (sr, _) if sr >= 16000 => 48000,              // Voice quality
                _ => 32000,                                   // Minimum quality
            }
        });
        
        // Set bitrate (bytes per second = bits per second / 8)
        media_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, bitrate_value / 8)?;
        
        // Set block alignment (not always required for AAC, but good practice)
        let block_align = format.block_align() as u32;
        if block_align > 0 {
            media_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, block_align)?;
        }
        
        // Set AAC-specific attributes
        
        // AAC Payload Type: 0 = Raw AAC
        // Other options include: 1 = ADTS, 2 = ADIF, 3 = LOAS
        media_type.SetUINT32(&MF_MT_AAC_PAYLOAD_TYPE, 0)?;
        
        // AAC Profile: 0x29 = AAC Low Complexity (LC) profile
        // Other common profiles: 0x2B = HE-AAC, 0x2C = HE-AACv2
        media_type.SetUINT32(&MF_MT_AAC_AUDIO_PROFILE_LEVEL_INDICATION, 0x29)?;
        
        // For channel masks (important for multi-channel audio)
        if let Some(mask) = format.channel_mask {
            if mask != 0 {
                media_type.SetUINT32(&MF_MT_AUDIO_CHANNEL_MASK, mask)?;
            }
        } else if format.channels > 2 {
            // For more than stereo, a channel mask is strongly recommended
            // But we just log a warning and continue
            println!("Warning: No channel mask specified for multi-channel AAC audio");
        }
        
        // Set interlace mode to progressive (standard for audio)
        media_type.SetUINT32(&MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive.0 as u32)?;
        
        Ok(media_type)
    }
}