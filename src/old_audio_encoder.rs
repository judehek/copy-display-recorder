// src/video/audio_encoder.rs
use std::{
    mem::ManuallyDrop, sync::{
        atomic::{AtomicBool, Ordering},
        mpsc::{Receiver, TryRecvError}, // Need Receiver for input packets
        Arc,
    }, thread::JoinHandle
};

use windows::{
    core::{Interface, Result},
    Win32::Media::MediaFoundation::{
        IMFActivate, IMFAttributes, IMFMediaEventGenerator, IMFMediaType, IMFSample, IMFTransform, METransformHaveOutput, METransformNeedInput, MFAudioFormat_AAC, MFAudioFormat_Float, MFAudioFormat_PCM, MFCreateAttributes, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFEnumDeviceSources, MFMediaType_Audio, MFStartup, MFTEnumEx, MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS, MFSTARTUP_FULL, MFT_CATEGORY_AUDIO_ENCODER, MFT_ENUM_FLAG_ALL, MFT_ENUM_FLAG_SORTANDFILTER, MFT_MESSAGE_COMMAND_DRAIN, MFT_MESSAGE_COMMAND_FLUSH, MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, MFT_MESSAGE_NOTIFY_END_OF_STREAM, MFT_MESSAGE_NOTIFY_END_STREAMING, MFT_MESSAGE_NOTIFY_START_OF_STREAM, MFT_OUTPUT_DATA_BUFFER, MFT_OUTPUT_STREAM_INFO, MFT_REGISTER_TYPE_INFO, MFT_SET_TYPE_TEST_ONLY, MF_EVENT_TYPE, MF_E_INVALIDMEDIATYPE, MF_E_NOTACCEPTING, MF_E_NO_MORE_TYPES, MF_E_TRANSFORM_NEED_MORE_INPUT, MF_E_TRANSFORM_TYPE_NOT_SET, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_CHANNEL_MASK, MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE, MF_TRANSFORM_ASYNC_UNLOCK // Keep for event loop
    },
};

use crate::{
    audio::{create_imf_sample_from_packet, AudioDataPacket}, // Need input packet type
    media::{enumerate_mfts, get_string_attribute, MF_VERSION},                                       // Keep MF version
};
use windows::core::ComInterface;
// Input to the encoder is the raw packet from AudioCapture
pub type AudioEncoderInputSample = AudioDataPacket;

// Add this struct definition (if not already present from previous step)
#[derive(Debug)]
pub struct EncodedAudioPacket { // Make pub if needed across modules
    pub data: Vec<u8>,
    pub timestamp: i64,
    pub duration: i64,
}

pub struct AudioEncoder {
    inner: Option<AudioEncoderInner>,
    output_type: IMFMediaType, // The AAC output type
    started: AtomicBool,
    should_stop: Arc<AtomicBool>,
    encoder_thread_handle: Option<JoinHandle<Result<()>>>,
}

struct AudioEncoderInner {
    transform: IMFTransform,
    input_stream_id: u32,
    output_stream_id: u32,

    sample_requested_callback:
        Option<Box<dyn Send + FnMut() -> Result<Option<AudioEncoderInputSample>>>>,
    // --- Callback now accepts EncodedAudioPacket ---
    sample_rendered_callback: Option<Box<dyn Send + FnMut(EncodedAudioPacket) -> Result<()>>>,

    should_stop: Arc<AtomicBool>,
    draining: bool,
}

// Helper to find and activate the AAC encoder MFT
unsafe fn find_and_activate_aac_encoder(
    input_type: &IMFMediaType,
    output_type: &IMFMediaType,
) -> Result<IMFTransform> {
    let input_info = MFT_REGISTER_TYPE_INFO {
        guidMajorType: input_type.GetGUID(&MF_MT_MAJOR_TYPE)?,
        guidSubtype: input_type.GetGUID(&MF_MT_SUBTYPE)?,
    };
    
    let output_info = MFT_REGISTER_TYPE_INFO {
        guidMajorType: output_type.GetGUID(&MF_MT_MAJOR_TYPE)?,
        guidSubtype: output_type.GetGUID(&MF_MT_SUBTYPE)?,
    };

    // Use proper flag types
    let flags = MFT_ENUM_FLAG_ALL | MFT_ENUM_FLAG_SORTANDFILTER;
    
    // Use your helper to get a Vec of IMFActivate objects
    let transform_activates = enumerate_mfts(
        &MFT_CATEGORY_AUDIO_ENCODER,
        flags,
        Some(&input_info),
        Some(&output_info),
    )?;
    
    // Check if any encoder was found
    if transform_activates.is_empty() {
        return Err(windows::core::Error::new(
            windows::Win32::Foundation::E_FAIL,
            "No suitable AAC encoder MFT found".into(),
        ));
    }
    
    // Use the first one found
    let activate = &transform_activates[0];
    
    // Activate it to get the transform
    activate.ActivateObject::<IMFTransform>()
}

impl AudioEncoder {
    pub fn new(
        raw_input_type: &IMFMediaType, // The PCM/Float type from AudioCapture
        // Define desired AAC output params
        aac_sample_rate: u32,
        aac_channels: u32,
        aac_bit_rate: u32,
    ) -> Result<Self> {
        unsafe {
            // 1. Create the desired AAC output media type
            let output_type = MFCreateMediaType()?;
            output_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
            output_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
            output_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, aac_sample_rate)?; // e.g., 48000
            output_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, aac_channels)?;       // e.g., 2
            output_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, aac_bit_rate / 8)?; // e.g., 16000
            // *** ADD THIS LINE based on documentation ***
            output_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?;

            // 2. Find and Activate the AAC Encoder MFT
            //    Pass the RAW input type and the DESIRED AAC output type
            let transform = find_and_activate_aac_encoder(raw_input_type, &output_type)?;
            println!("AAC Encoder MFT activated.");

            // 3. Setup Basic MFT Attributes (No D3D Manager needed)
            //let event_generator: IMFMediaEventGenerator = transform.cast()?;
            /*let attributes = transform.GetAttributes()?;
            println!("got attributes");
            attributes.SetUINT32(&MF_TRANSFORM_ASYNC_UNLOCK, 1)?; // Enable async processing
            println!("MFT Attributes set (Async Unlock).");*/

            // 4. Get Stream IDs (Same logic as VideoEncoder)
            let mut number_of_input_streams = 0;
            let mut number_of_output_streams = 0;
            transform.GetStreamCount(&mut number_of_input_streams, &mut number_of_output_streams)?;
            if number_of_input_streams == 0 || number_of_output_streams == 0 {
                 return Err(windows::core::Error::new(windows::Win32::Foundation::E_UNEXPECTED, "MFT has no input or output streams".into()));
            }
            // Assuming stream IDs are 0 for simple encoders
            let input_stream_id = 0;
            let output_stream_id = 0;
            println!("Stream IDs obtained: Input={}, Output={}", input_stream_id, output_stream_id);


            // 5. Set Media Types on the MFT
            // Set Output Type FIRST for encoders
            transform.SetOutputType(output_stream_id, &output_type, 0)?;
            println!("Output type (AAC) set successfully on MFT.");

            // Set Input Type (using the raw type passed in)
            // No need to loop/test like video, assume the enumeration found a compatible one
            transform.SetInputType(input_stream_id, raw_input_type, 0)?;
            println!("Input type (PCM/Float) set successfully on MFT.");


            // 6. Prepare Inner Struct
            let should_stop = Arc::new(AtomicBool::new(false));
            let inner = AudioEncoderInner {
                transform,
                input_stream_id,
                output_stream_id,
                sample_requested_callback: None, // Set later
                sample_rendered_callback: None, // Set later
                should_stop: should_stop.clone(),
                draining: false,
            };

            Ok(Self {
                inner: Some(inner),
                output_type, // Store the AAC output type
                started: AtomicBool::new(false),
                should_stop,
                encoder_thread_handle: None,
            })
        }
    }

    pub fn try_start(&mut self) -> Result<bool> {
        if self
            .started
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .is_ok()
        {
            let mut inner = self.inner.take().expect("Inner should exist before start");

            if inner.sample_requested_callback.is_none()
                || inner.sample_rendered_callback.is_none()
            {
                panic!("AudioEncoder callbacks must be set before starting");
            }

            println!("Starting AudioEncoder thread...");
            self.encoder_thread_handle = Some(std::thread::spawn(move || -> Result<()> {
                // Initialize COM for this thread
                unsafe { MFStartup(MF_VERSION, MFSTARTUP_FULL)? };

                let result = inner.encode_loop();
                if let Err(e) = &result {
                    eprintln!("AudioEncoder thread failed: HRESULT={:?}, Msg={}", e.code(), e.message());
                } else {
                    println!("AudioEncoder thread finished successfully.");
                }
                result
            }));
            Ok(true)
        } else {
            Ok(false) // Already started
        }
    }

    pub fn stop(&mut self) -> Result<()> {
        if self
            .started
            .compare_exchange(true, false, Ordering::SeqCst, Ordering::SeqCst) // Ensure we only stop once
            .is_ok()
        {
             println!("Stopping AudioEncoder...");
            // Signal the encoder thread to stop processing new input
             self.should_stop.store(true, Ordering::SeqCst);

            // Wait for the thread to finish processing and exit
            self.wait_for_completion()?;
            println!("AudioEncoder stopped.");
        } else {
            println!("AudioEncoder stop called but not running or already stopped.");
        }
        Ok(())
    }


    fn wait_for_completion(&mut self) -> Result<()> {
        if let Some(handle) = self.encoder_thread_handle.take() {
            println!("Waiting for audio encoder thread to join...");
            match handle.join() {
                Ok(thread_result) => {
                    println!("Audio encoder thread joined.");
                    thread_result // Propagate Result from thread
                }
                Err(e) => {
                    eprintln!("Audio encoder thread panicked: {:?}", e);
                    Err(windows::core::Error::new(
                        windows::Win32::Foundation::E_FAIL, // Indicate thread panic
                        "Audio encoder thread panic".into(),
                    ))
                }
            }
        } else {
             println!("AudioEncoder: wait_for_completion called but no thread handle exists.");
             Ok(()) // No thread was running
        }
    }


    // Callback to provide raw audio packets
    pub fn set_sample_rendered_callback<
        F: 'static + Send + FnMut(EncodedAudioPacket) -> Result<()>, // <<< CHANGED TYPE
    >(
        &mut self,
        callback: F,
    ) {
         if let Some(inner) = self.inner.as_mut() {
             inner.sample_rendered_callback = Some(Box::new(callback));
         } else {
              eprintln!("Warning: Tried to set sample_rendered_callback after encoder started.");
         }
    }

    pub fn set_sample_requested_callback<
        F: 'static + Send + FnMut() -> Result<Option<AudioEncoderInputSample>>,
    >(
        &mut self,
        callback: F,
    ) {
        if let Some(inner) = self.inner.as_mut() {
            inner.sample_requested_callback = Some(Box::new(callback));
        } else {
             eprintln!("Warning: Tried to set sample_requested_callback after encoder started.");
        }
    }

    // Returns the AAC output type
    pub fn output_type(&self) -> &IMFMediaType {
        &self.output_type
    }

     // Clone the output type
     pub fn clone_output_type(&self) -> Result<IMFMediaType> {
         unsafe {
             let new_type = MFCreateMediaType()?;
             self.output_type.CopyAllItems(&new_type)?;
             Ok(new_type)
         }
     }
}

pub unsafe fn log_media_type_details(media_type: &IMFMediaType) -> Result<()> {
    println!("  Major Type: {:?}", media_type.GetGUID(&MF_MT_MAJOR_TYPE)?);
    let subtype_guid = media_type.GetGUID(&MF_MT_SUBTYPE)?;
    println!("  Subtype: {:?}", subtype_guid);
    // Add more specific subtype names if needed
    if subtype_guid == MFAudioFormat_Float { println!("    (MFAudioFormat_Float)"); }
    else if subtype_guid == MFAudioFormat_PCM { println!("    (MFAudioFormat_PCM)"); }
    else if subtype_guid == MFAudioFormat_AAC { println!("    (MFAudioFormat_AAC)"); }
    // Add other known subtypes...

    println!(
        "  Sample Rate: {:?}",
        media_type.GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
    );
    println!(
        "  Channels: {:?}",
        media_type.GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
    );
    // Optional attributes, might not be set or relevant for all types
    match media_type.GetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE) {
        Ok(v) => println!("  Bits Per Sample: {}", v),
        Err(_) => println!("  Bits Per Sample: (Not Set or N/A)"), // Expected for Float/AAC
    }
    match media_type.GetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT) {
        Ok(v) => println!("  Block Align: {}", v),
        Err(_) => println!("  Block Align: (Not Set or N/A)"), // Might be missing for AAC before connection
    }
    match media_type.GetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND) {
        Ok(v) => println!("  Avg Bytes Per Sec: {}", v),
        Err(_) => println!("  Avg Bytes Per Sec: (Not Set or N/A)"), // Expected for PCM/Float
    }
    Ok(())
}

// Mark Inner as Send (assuming callbacks are Send)
unsafe impl Send for AudioEncoderInner {}

// Constants for event matching (avoids non_upper_case_globals warning)
const MEDIA_ENGINE_TRANFORM_NEED_INPUT: MF_EVENT_TYPE = METransformNeedInput;
const MEDIA_ENGINE_TRANFORM_HAVE_OUTPUT: MF_EVENT_TYPE = METransformHaveOutput;

impl AudioEncoderInner {
    // The main loop running on the encoder thread
    fn encode_loop(&mut self) -> Result<()> {
        unsafe {
            // Initial MFT messages
            self.transform
                .ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_START_OF_STREAM, 0)?;
            println!("Audio MFT streaming started.");

            let mut input_needed = true; // Start by requesting input
            let mut output_ready = false;

            while !(self.draining && !output_ready) { // Loop until drained and no more output pending

                 // Check stop signal BEFORE blocking on GetEvent
                 if self.should_stop.load(Ordering::SeqCst) && !self.draining {
                     println!("Stop signal received, initiating drain...");
                     self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, 0)?;
                     self.draining = true;
                     input_needed = false; // Don't request more input after drain starts
                 }


                // --- Process Input if Needed ---
                if input_needed && !self.draining {
                    match self.process_input() {
                        Ok(more_input_available) => {
                            input_needed = !more_input_available; // If callback returned None, stop asking for input (until output processed)
                            if !more_input_available && !self.draining {
                                 println!("Input source signalled end, initiating drain...");
                                 self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, 0)?;
                                 self.draining = true;
                            }
                        }
                        Err(e) if e.code() == MF_E_TRANSFORM_TYPE_NOT_SET => {
                            // This shouldn't happen if SetInputType succeeded
                            eprintln!("Error: MFT Input stream not set!");
                            return Err(e);
                        }
                        Err(e) if e.code() == MF_E_NOTACCEPTING => {
                            // MFT cannot accept input right now, need output first
                            input_needed = false;
                            println!("MFT not accepting input, waiting for output.");
                        }
                        Err(e) => {
                            eprintln!("Error during ProcessInput: {:?}", e);
                            return Err(e);
                        }
                    }
                }

                // --- Process Output if Available ---
                 // Check for output *after* trying input, or if we know output is ready
                if output_ready || !input_needed || self.draining {
                     match self.process_output() {
                         Ok(()) => {
                             // Successfully processed an output sample
                             output_ready = true; // Assume more might be ready
                             input_needed = true; // Can potentially accept input again
                         }
                         Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
                             // MFT needs more input to produce output
                             output_ready = false;
                             input_needed = true; // Request input again
                             if self.draining {
                                 println!("Drain complete (MF_E_TRANSFORM_NEED_MORE_INPUT).");
                                 // Drain is finished when this happens
                                 break; // Exit loop
                             }
                         }
                         Err(e) => {
                             eprintln!("Error during ProcessOutput: {:?}", e);
                             return Err(e);
                         }
                     }
                 }


                // --- Event Handling (Alternative to Polling Input/Output) ---
                // This part is less critical if actively polling process_input/output
                // based on input_needed/output_ready flags, but can be kept for robustness.
                // You might need to adjust the logic above if relying solely on events.

                // if !input_needed && !output_ready && !self.draining {
                //     // If we can't input and don't have output, wait for an event
                //     println!("Waiting for MFT event...");
                //     let event = self
                //         .event_generator
                //         .GetEvent(MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS(0))?; // Blocking wait
                //
                //     let event_type = MF_EVENT_TYPE(event.GetType()? as i32);
                //     match event_type {
                //         MEDIA_ENGINE_TRANFORM_NEED_INPUT => {
                //             println!("Received METransformNeedInput event.");
                //             input_needed = true;
                //         }
                //         MEDIA_ENGINE_TRANFORM_HAVE_OUTPUT => {
                //             println!("Received METransformHaveOutput event.");
                //             output_ready = true;
                //         }
                //         _ => {
                //             println!("Received unknown MFT event type: {}", event_type.0);
                //             // Handle other events if necessary
                //         }
                //     }
                // }

                 // Small sleep to prevent busy-waiting if polling flags
                 if !input_needed && !output_ready && !self.draining {
                      std::thread::sleep(std::time::Duration::from_millis(5));
                 }


            } // End while loop

            println!("AudioEncoder loop finished. Sending final messages.");
            // Final MFT messages (EndStreaming only needed if BeginStreaming succeeded)
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_END_OF_STREAM, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_END_STREAMING, 0)?;
             self.transform
                 .ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?; // Final flush

            println!("Audio MFT streaming ended.");
        } // End unsafe block
        Ok(())
    }


    // Tries to get a raw audio packet and send it to the MFT
    // Returns Ok(true) if input was processed, Ok(false) if no input was available (EOS)
    // Returns Err on MFT error
    fn process_input(&mut self) -> Result<bool> {
        match self.sample_requested_callback.as_mut().unwrap()() {
            Ok(Some(packet)) => {
                // Create IMFSample from the raw AudioDataPacket
                // Use the helper function but ensure it handles Option<Vec<u8>> correctly
                 let input_sample = match create_imf_sample_from_raw_packet(packet) {
                     Ok(sample) => sample,
                     Err(e) => {
                          eprintln!("Failed to create IMF sample from audio packet: {:?}", e);
                          // Skip this packet? Or return error?
                          return Ok(true); // Indicate we *tried* to process input
                     }
                 };

                // Send the raw sample to the MFT
                unsafe {
                    // println!("AudioEncoder: Processing Input Sample Time: {}", input_sample.GetSampleTime()?);
                    self.transform
                        .ProcessInput(self.input_stream_id, &input_sample, 0)?;
                }
                Ok(true) // Input was available and processed
            }
            Ok(None) => {
                // Input source signaled end-of-stream
                println!("AudioEncoder: Input source returned None (EOS).");
                Ok(false) // No more input available from source
            }
            Err(e) => {
                eprintln!("AudioEncoder: Error getting sample from callback: {:?}", e);
                Err(e) // Propagate error from callback
            }
        }
    }

    // Tries to get an encoded AAC sample from the MFT and send it via callback
    fn process_output(&mut self) -> Result<()> {
        unsafe {
            // ... (Prepare output_buffer_struct as before) ...
            let mut output_stream_info = self.transform.GetOutputStreamInfo(self.output_stream_id)?;
            let output_sample = MFCreateSample()?;
            let output_buffer = MFCreateMemoryBuffer(output_stream_info.cbSize)?;
            output_sample.AddBuffer(&output_buffer)?;

            let mut output_buffer_struct = MFT_OUTPUT_DATA_BUFFER {
                dwStreamID: self.output_stream_id,
                pSample: ManuallyDrop::new(Some(output_sample)), // Pass sample
                dwStatus: 0,
                pEvents: ManuallyDrop::new(None),
            };
            let mut status_flags = 0u32;
            let hr = self.transform.ProcessOutput(
                0,
                std::slice::from_mut(&mut output_buffer_struct),
                &mut status_flags,
            );

            let processed_sample: Option<IMFSample> = if !output_buffer_struct.pSample.is_none() {
                ManuallyDrop::take(&mut output_buffer_struct.pSample)
            } else {
                None
            };

            // Handle ProcessOutput result
            match hr {
                Ok(()) => {
                    if let Some(sample) = processed_sample {
                        // --- Extraction logic moved HERE ---
                        let packet_result: Result<EncodedAudioPacket> = { // Use block scope for result
                             let timestamp = sample.GetSampleTime()?;
                             let duration = sample.GetSampleDuration()?;
                             let buffer = sample.GetBufferByIndex(0)?;
                             let mut buffer_ptr = std::ptr::null_mut();
                             let mut max_len = 0;
                             let mut current_len = 0;
                             buffer.Lock(&mut buffer_ptr, Some(&mut max_len), Some(&mut current_len))?;

                             if current_len > 0 && !buffer_ptr.is_null() {
                                 let data_slice = std::slice::from_raw_parts(buffer_ptr as *const u8, current_len as usize);
                                 Ok(EncodedAudioPacket {
                                     data: data_slice.to_vec(),
                                     timestamp,
                                     duration,
                                 })
                             } else {
                                 println!("Warning: AudioEncoderInner::process_output produced zero-length buffer sample. Time: {}", timestamp);
                                 Ok(EncodedAudioPacket {
                                     data: Vec::new(),
                                     timestamp,
                                     duration,
                                 })
                             }
                        }; // End of extraction logic block

                        // Call the callback with the extracted packet
                        match packet_result {
                            Ok(packet) => {
                                self.sample_rendered_callback.as_mut().unwrap()(packet)?;
                                Ok(()) // Indicate success for this output sample
                            }
                            Err(e) => {
                                // Error during extraction
                                eprintln!("Error extracting data from encoded audio sample: {:?}", e);
                                Err(e) // Propagate extraction error
                            }
                        }
                        // --- End of moved logic ---
                    } else {
                         eprintln!("Warning: ProcessOutput returned OK but no sample pointer.");
                         Err(MF_E_TRANSFORM_NEED_MORE_INPUT.into())
                    }
                }
                Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
                     if let Some(s) = processed_sample { drop(s); }
                     Err(e)
                }
                Err(e) => {
                     if let Some(s) = processed_sample { drop(s); }
                     eprintln!("AudioEncoder: ProcessOutput failed: {:?}", e);
                     Err(e)
                }
            }
        } // End unsafe
    }
}

// Modified helper to create IMFSample from AudioDataPacket (handles Option<Vec>)
// Renamed to avoid conflict with the original one.
fn create_imf_sample_from_raw_packet(packet: AudioDataPacket) -> Result<IMFSample> {
    unsafe {
        let sample = MFCreateSample()?;
        sample.SetSampleTime(packet.timestamp)?;
        sample.SetSampleDuration(packet.duration)?;

        // packet.flags could be used here if needed

        if let Some(data) = packet.data {
            if !data.is_empty() {
                let media_buffer = MFCreateMemoryBuffer(data.len() as u32)?;
                { // Scope for lock guard
                    let mut buffer_ptr = std::ptr::null_mut();
                    let mut max_len = 0;
                    media_buffer.Lock(&mut buffer_ptr, Some(&mut max_len), None)?;

                    if (data.len() as u32) <= max_len {
                        std::ptr::copy_nonoverlapping(data.as_ptr(), buffer_ptr, data.len());
                        media_buffer.SetCurrentLength(data.len() as u32)?;
                    } else {
                        eprintln!("Error: Media buffer too small in create_imf_sample_from_raw_packet");
                        return Err(windows::core::Error::new(
                            windows::Win32::Foundation::E_FAIL,
                            "Media buffer size mismatch".into(),
                        ));
                    }
                } // Unlock guard drops here
                sample.AddBuffer(&media_buffer)?;
            } else {
                // Handle case of Some(empty_vec) - MFT might need 0-length buffer? Or skip?
                // For simplicity, we add no buffer here. Test if AAC encoder requires empty buffers for silence.
                println!("Warning: create_imf_sample_from_raw_packet got empty data vec.");
            }
        } else {
            // Handle None data (silence) - Add no buffer. MFT might generate silence or handle gap.
            // MFSampleExtension_Discontinuity could potentially be set here if silence represents a gap.
            // sample.SetUINT32(&MFSampleExtension_Discontinuity, 1)?;
            println!("AudioEncoder: Creating sample for silence (no buffer). Time: {}", packet.timestamp);

        }
        Ok(sample)
    }
}

// --- Drop Implementation ---
impl Drop for AudioEncoder {
    fn drop(&mut self) {
        if self.started.load(Ordering::SeqCst) {
             println!("AudioEncoder drop: Forcing stop...");
             // Use SeqCst for consistency, though AcqRel might suffice for should_stop
             self.should_stop.store(true, Ordering::SeqCst);
             if let Err(e) = self.wait_for_completion() {
                 eprintln!("Error during AudioEncoder drop/wait: {:?}", e);
             }
        }
        println!("AudioEncoder dropped.");
        // Inner is dropped automatically if it exists
        // COM cleanup happens in the thread guard or when MFT is dropped.
    }
}