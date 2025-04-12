// use std::{
//   mem::ManuallyDrop,
//   sync::{
//       atomic::{AtomicBool, Ordering},
//       Arc,
//   },
//   thread::JoinHandle,
// };

// use windows::{
//   core::{ComInterface, Interface, Result},
//   Win32::Media::MediaFoundation::{
//       IMFActivate, IMFAttributes, IMFMediaEventGenerator, IMFMediaType, IMFSample, IMFTransform,
//       METransformHaveOutput, METransformNeedInput, MFAudioFormat_AAC, MFAudioFormat_Float, MFAudioFormat_PCM,
//       MFCreateAttributes, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFEnumDeviceSources,
//       MFMediaType_Audio, MFStartup, MFTEnumEx, MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS, MFSTARTUP_FULL,
//       MFT_CATEGORY_AUDIO_ENCODER, MFT_ENUM_FLAG_ALL, MFT_ENUM_FLAG_SORTANDFILTER, MFT_MESSAGE_COMMAND_DRAIN,
//       MFT_MESSAGE_COMMAND_FLUSH, MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, MFT_MESSAGE_NOTIFY_END_OF_STREAM,
//       MFT_MESSAGE_NOTIFY_END_STREAMING, MFT_MESSAGE_NOTIFY_START_OF_STREAM, MFT_OUTPUT_DATA_BUFFER,
//       MFT_OUTPUT_STREAM_INFO, MFT_REGISTER_TYPE_INFO, MFT_SET_TYPE_TEST_ONLY, MF_EVENT_TYPE,
//       MF_E_INVALIDMEDIATYPE, MF_E_NOTACCEPTING, MF_E_NO_MORE_TYPES, MF_E_TRANSFORM_NEED_MORE_INPUT,
//       MF_E_TRANSFORM_TYPE_NOT_SET, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE,
//       MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_CHANNEL_MASK, MF_MT_AUDIO_NUM_CHANNELS,
//       MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE, MF_TRANSFORM_ASYNC_UNLOCK,
//   },
// };

// use crate::media::{enumerate_mfts, get_string_attribute, MF_VERSION};

// /// Represents an input sample for the audio encoder
// #[derive(Clone)]
// pub struct AudioEncoderInputSample {
//   /// PCM audio data
//   pub data: Vec<u8>,
//   /// Sample timestamp in 100ns units
//   pub timestamp: i64,
//   /// Duration of the sample in 100ns units
//   pub duration: i64,
// }

// impl AudioEncoderInputSample {
//   pub fn new(data: Vec<u8>, timestamp: i64, duration: i64) -> Self {
//       Self {
//           data,
//           timestamp,
//           duration,
//       }
//   }
// }

// /// Represents an encoded AAC audio sample from the encoder
// pub struct AudioEncoderOutputSample {
//   sample: IMFSample,
// }

// impl AudioEncoderOutputSample {
//   pub fn sample(&self) -> &IMFSample {
//       &self.sample
//   }

//   /// Extract the encoded AAC data as bytes
//   pub fn get_data(&self) -> Result<Vec<u8>> {
//       unsafe {
//           let buffer = self.sample.GetBufferByIndex(0)?;
//           let mut buffer_ptr = std::ptr::null_mut();
//           let mut max_len = 0;
//           let mut current_len = 0;
//           buffer.Lock(&mut buffer_ptr, Some(&mut max_len), Some(&mut current_len))?;
          
//           if current_len > 0 && !buffer_ptr.is_null() {
//               let data_slice = std::slice::from_raw_parts(buffer_ptr as *const u8, current_len as usize);
//               let result = data_slice.to_vec();
//               buffer.Unlock()?;
//               Ok(result)
//           } else {
//               buffer.Unlock()?;
//               Ok(Vec::new())
//           }
//       }
//   }

//   /// Get the timestamp of this sample
//   pub fn timestamp(&self) -> Result<i64> {
//       unsafe { self.sample.GetSampleTime() }
//   }

//   /// Get the duration of this sample
//   pub fn duration(&self) -> Result<i64> {
//       unsafe { self.sample.GetSampleDuration() }
//   }
// }

// /// Main audio encoder class that wraps the AAC encoding pipeline
// pub struct AudioEncoder {
//   inner: Option<AudioEncoderInner>,
//   output_type: IMFMediaType,
//   started: AtomicBool,
//   should_stop: Arc<AtomicBool>,
//   encoder_thread_handle: Option<JoinHandle<Result<()>>>,
// }

// struct AudioEncoderInner {
//   transform: IMFTransform,
//   event_generator: IMFMediaEventGenerator,
//   input_stream_id: u32,
//   output_stream_id: u32,

//   sample_requested_callback: Option<Box<dyn Send + FnMut() -> Result<Option<AudioEncoderInputSample>>>>,
//   sample_rendered_callback: Option<Box<dyn Send + FnMut(AudioEncoderOutputSample) -> Result<()>>>,

//   should_stop: Arc<AtomicBool>,
//   draining: bool,
// }

// // Helper to find and activate the AAC encoder MFT
// unsafe fn find_and_activate_aac_encoder(
//   input_type: &IMFMediaType,
//   output_type: &IMFMediaType,
// ) -> Result<IMFTransform> {
//   let input_info = MFT_REGISTER_TYPE_INFO {
//       guidMajorType: input_type.GetGUID(&MF_MT_MAJOR_TYPE)?,
//       guidSubtype: input_type.GetGUID(&MF_MT_SUBTYPE)?,
//   };
  
//   let output_info = MFT_REGISTER_TYPE_INFO {
//       guidMajorType: output_type.GetGUID(&MF_MT_MAJOR_TYPE)?,
//       guidSubtype: output_type.GetGUID(&MF_MT_SUBTYPE)?,
//   };

//   // Use proper flag types
//   let flags = MFT_ENUM_FLAG_ALL | MFT_ENUM_FLAG_SORTANDFILTER;
  
//   // Get a Vec of IMFActivate objects for AAC encoders
//   let transform_activates = enumerate_mfts(
//       &MFT_CATEGORY_AUDIO_ENCODER,
//       flags,
//       Some(&input_info),
//       Some(&output_info),
//   )?;
  
//   // Check if any encoder was found
//   if transform_activates.is_empty() {
//       return Err(windows::core::Error::new(
//           windows::Win32::Foundation::E_FAIL,
//           "No suitable AAC encoder MFT found".into(),
//       ));
//   }
  
//   // Use the first encoder found
//   let activate = &transform_activates[0];
  
//   // Activate it to get the transform
//   activate.ActivateObject::<IMFTransform>()
// }

// impl AudioEncoder {
//   pub fn new(
//       sample_rate: u32,
//       channels: u32,
//       bit_rate: u32,
//   ) -> Result<Self> {
//       unsafe {
//           // Create the input PCM media type
//           let input_type = MFCreateMediaType()?;
//           input_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
//           input_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)?;
//           input_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)?;
//           input_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, channels)?;
//           input_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?; // Standard PCM is 16-bit
          
//           // Calculate block alignment and bytes per second for PCM
//           let block_align = channels * (16 / 8); // channels * bytes per sample
//           input_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, block_align as u32)?;
//           input_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, sample_rate * block_align as u32)?;
          
//           // Create the output AAC media type
//           let output_type = MFCreateMediaType()?;
//           output_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
//           output_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
//           output_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, sample_rate)?;
//           output_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, channels)?;
//           output_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 16)?;
//           output_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, bit_rate / 8)?; // bit rate to bytes per sec

//           // Find and activate the AAC encoder MFT
//           let transform = find_and_activate_aac_encoder(&input_type, &output_type)?;
//           println!("AAC Encoder MFT activated.");

//           // Setup event generator 
//           let event_generator: IMFMediaEventGenerator = transform.cast()?;
          
//           // Setup MFT attributes
//           let attributes = transform.GetAttributes()?;
//           attributes.SetUINT32(&MF_TRANSFORM_ASYNC_UNLOCK, 1)?; // Enable async
//           println!("MFT Attributes set (Async Unlock).");

//           // Get stream IDs
//           let mut number_of_input_streams = 0;
//           let mut number_of_output_streams = 0;
//           transform.GetStreamCount(&mut number_of_input_streams, &mut number_of_output_streams)?;
//           if number_of_input_streams == 0 || number_of_output_streams == 0 {
//               return Err(windows::core::Error::new(
//                   windows::Win32::Foundation::E_UNEXPECTED,
//                   "MFT has no input or output streams".into(),
//               ));
//           }
          
//           // Assuming stream IDs are 0 for simple encoders
//           let input_stream_id = 0;
//           let output_stream_id = 0;
//           println!("Stream IDs obtained: Input={}, Output={}", input_stream_id, output_stream_id);

//           // Set output type first (this is important for encoders)
//           transform.SetOutputType(output_stream_id, &output_type, 0)?;
//           println!("Output type (AAC) set successfully on MFT.");

//           // Then set input type
//           transform.SetInputType(input_stream_id, &input_type, 0)?;
//           println!("Input type (PCM) set successfully on MFT.");

//           // Create inner state object
//           let should_stop = Arc::new(AtomicBool::new(false));
//           let inner = AudioEncoderInner {
//               transform,
//               event_generator,
//               input_stream_id,
//               output_stream_id,
//               sample_requested_callback: None,
//               sample_rendered_callback: None,
//               should_stop: should_stop.clone(),
//               draining: false,
//           };

//           Ok(Self {
//               inner: Some(inner),
//               output_type,
//               started: AtomicBool::new(false),
//               should_stop,
//               encoder_thread_handle: None,
//           })
//       }
//   }

//   pub fn try_start(&mut self) -> Result<bool> {
//       if self
//           .started
//           .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
//           .is_ok()
//       {
//           let mut inner = self.inner.take().expect("Inner should exist before start");

//           if inner.sample_requested_callback.is_none() 
//               || inner.sample_rendered_callback.is_none() 
//           {
//               panic!("AudioEncoder callbacks must be set before starting");
//           }

//           println!("Starting AudioEncoder thread...");
//           self.encoder_thread_handle = Some(std::thread::spawn(move || -> Result<()> {
//               // Initialize COM for this thread
//               unsafe { MFStartup(MF_VERSION, MFSTARTUP_FULL)? };

//               let result = inner.encode_loop();
//               if let Err(e) = &result {
//                   eprintln!("AudioEncoder thread failed: HRESULT={:?}, Msg={}", e.code(), e.message());
//               } else {
//                   println!("AudioEncoder thread finished successfully.");
//               }
//               result
//           }));
//           Ok(true)
//       } else {
//           Ok(false) // Already started
//       }
//   }

//   pub fn stop(&mut self) -> Result<()> {
//       if self
//           .started
//           .compare_exchange(true, false, Ordering::SeqCst, Ordering::SeqCst)
//           .is_ok()
//       {
//           println!("Stopping AudioEncoder...");
//           // Signal the encoder thread to stop processing new input
//           self.should_stop.store(true, Ordering::SeqCst);

//           // Wait for the thread to finish processing and exit
//           if let Err(e) = self.wait_for_completion() {
//               eprintln!("Error waiting for audio encoder completion: {:?}", e);
//               return Err(e);
//           }
//           println!("AudioEncoder stopped.");
//       } else {
//           println!("AudioEncoder stop called but not running or already stopped.");
//       }
//       Ok(())
//   }

//   fn wait_for_completion(&mut self) -> Result<()> {
//       if let Some(handle) = self.encoder_thread_handle.take() {
//           println!("Waiting for audio encoder thread to join...");
//           match handle.join() {
//               Ok(thread_result) => {
//                   println!("Audio encoder thread joined.");
//                   thread_result // Propagate Result from thread
//               }
//               Err(e) => {
//                   eprintln!("Audio encoder thread panicked: {:?}", e);
//                   Err(windows::core::Error::new(
//                       windows::Win32::Foundation::E_FAIL,
//                       "Audio encoder thread panic".into(),
//                   ))
//               }
//           }
//       } else {
//           println!("AudioEncoder: wait_for_completion called but no thread handle exists.");
//           Ok(()) // No thread was running
//       }
//   }

//   // Set callback that will be called when the encoder needs input data
//   pub fn set_sample_requested_callback<
//       F: 'static + Send + FnMut() -> Result<Option<AudioEncoderInputSample>>,
//   >(
//       &mut self,
//       callback: F,
//   ) {
//       if let Some(inner) = self.inner.as_mut() {
//           inner.sample_requested_callback = Some(Box::new(callback));
//       } else {
//           eprintln!("Warning: Tried to set sample_requested_callback after encoder started.");
//       }
//   }

//   // Set callback that will be called when the encoder has produced output data
//   pub fn set_sample_rendered_callback<
//       F: 'static + Send + FnMut(AudioEncoderOutputSample) -> Result<()>,
//   >(
//       &mut self,
//       callback: F,
//   ) {
//       if let Some(inner) = self.inner.as_mut() {
//           inner.sample_rendered_callback = Some(Box::new(callback));
//       } else {
//           eprintln!("Warning: Tried to set sample_rendered_callback after encoder started.");
//       }
//   }

//   // Returns the AAC output type
//   pub fn output_type(&self) -> &IMFMediaType {
//       &self.output_type
//   }

//   // Clone the output type
//   pub fn clone_output_type(&self) -> Result<IMFMediaType> {
//       unsafe {
//           let new_type = MFCreateMediaType()?;
//           self.output_type.CopyAllItems(&new_type)?;
//           Ok(new_type)
//       }
//   }
// }

// // Helper to log media type details for debugging
// pub unsafe fn log_media_type_details(media_type: &IMFMediaType) -> Result<()> {
//   println!("  Major Type: {:?}", media_type.GetGUID(&MF_MT_MAJOR_TYPE)?);
//   let subtype_guid = media_type.GetGUID(&MF_MT_SUBTYPE)?;
//   println!("  Subtype: {:?}", subtype_guid);
  
//   if subtype_guid == MFAudioFormat_Float { println!("    (MFAudioFormat_Float)"); }
//   else if subtype_guid == MFAudioFormat_PCM { println!("    (MFAudioFormat_PCM)"); }
//   else if subtype_guid == MFAudioFormat_AAC { println!("    (MFAudioFormat_AAC)"); }

//   println!(
//       "  Sample Rate: {:?}",
//       media_type.GetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND)
//   );
//   println!(
//       "  Channels: {:?}",
//       media_type.GetUINT32(&MF_MT_AUDIO_NUM_CHANNELS)
//   );
  
//   match media_type.GetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE) {
//       Ok(v) => println!("  Bits Per Sample: {}", v),
//       Err(_) => println!("  Bits Per Sample: (Not Set or N/A)"),
//   }
//   match media_type.GetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT) {
//       Ok(v) => println!("  Block Align: {}", v),
//       Err(_) => println!("  Block Align: (Not Set or N/A)"),
//   }
//   match media_type.GetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND) {
//       Ok(v) => println!("  Avg Bytes Per Sec: {}", v),
//       Err(_) => println!("  Avg Bytes Per Sec: (Not Set or N/A)"),
//   }
//   Ok(())
// }

// // Mark Inner as Send
// unsafe impl Send for AudioEncoderInner {}

// // Constants for event matching
// const MEDIA_ENGINE_TRANFORM_NEED_INPUT: MF_EVENT_TYPE = METransformNeedInput;
// const MEDIA_ENGINE_TRANFORM_HAVE_OUTPUT: MF_EVENT_TYPE = METransformHaveOutput;

// impl AudioEncoderInner {
//   // The main loop running on the encoder thread
//   fn encode_loop(&mut self) -> Result<()> {
//       unsafe {
//           // Initial MFT messages
//           self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?;
//           self.transform.ProcessMessage(MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, 0)?;
//           self.transform.ProcessMessage(MFT_MESSAGE_NOTIFY_START_OF_STREAM, 0)?;
//           println!("Audio MFT streaming started.");

//           let mut should_exit = false;

//           while !should_exit {
//               // Check if we should stop
//               if self.should_stop.load(Ordering::SeqCst) && !self.draining {
//                   println!("Stop signal received, initiating drain...");
//                   self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, 0)?;
//                   self.draining = true;
//               }

//               // Wait for events from the MFT
//               let event = self
//                   .event_generator
//                   .GetEvent(MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS(0))?;

//               let event_type = MF_EVENT_TYPE(event.GetType()? as i32);
//               match event_type {
//                   MEDIA_ENGINE_TRANFORM_NEED_INPUT => {
//                       if self.draining {
//                           // Drain complete when MFT asks for more input during draining
//                           should_exit = true;
//                       } else {
//                           // Request input sample and process it
//                           should_exit = self.process_input()?;
//                       }
//                   }
//                   MEDIA_ENGINE_TRANFORM_HAVE_OUTPUT => {
//                       // Process output sample
//                       self.process_output()?;
//                   }
//                   _ => {
//                       println!("Received unknown MFT event type: {}", event_type.0);
//                   }
//               }
//           }

//           println!("AudioEncoder loop finished. Sending final messages.");
//           // Final MFT messages
//           self.transform.ProcessMessage(MFT_MESSAGE_NOTIFY_END_OF_STREAM, 0)?;
//           self.transform.ProcessMessage(MFT_MESSAGE_NOTIFY_END_STREAMING, 0)?;
//           self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?;

//           println!("Audio MFT streaming ended.");
//       }
//       Ok(())
//   }

//   // Process input sample
//   // Returns true if encoding should exit
//   fn process_input(&mut self) -> Result<bool> {
//       match self.sample_requested_callback.as_mut().unwrap()() {
//           Ok(Some(input_sample)) => {
//               unsafe {
//                   // Create IMFSample from the input sample
//                   let mf_sample = MFCreateSample()?;
//                   mf_sample.SetSampleTime(input_sample.timestamp)?;
//                   mf_sample.SetSampleDuration(input_sample.duration)?;

//                   // Create buffer and copy data
//                   if !input_sample.data.is_empty() {
//                       let buffer = MFCreateMemoryBuffer(input_sample.data.len() as u32)?;
                      
//                       let mut buffer_ptr = std::ptr::null_mut();
//                       let mut max_len = 0;
//                       buffer.Lock(&mut buffer_ptr, Some(&mut max_len), None)?;
                      
//                       if (input_sample.data.len() as u32) <= max_len {
//                           std::ptr::copy_nonoverlapping(
//                               input_sample.data.as_ptr(),
//                               buffer_ptr,
//                               input_sample.data.len(),
//                           );
//                           buffer.SetCurrentLength(input_sample.data.len() as u32)?;
//                       } else {
//                           buffer.Unlock()?;
//                           return Err(windows::core::Error::new(
//                               windows::Win32::Foundation::E_FAIL,
//                               "Buffer too small for audio data".into(),
//                           ));
//                       }
//                       buffer.Unlock()?;
//                       mf_sample.AddBuffer(&buffer)?;
//                   }

//                   // Process the input
//                   match self.transform.ProcessInput(self.input_stream_id, &mf_sample, 0) {
//                       Ok(_) => Ok(false), // Continue encoding
//                       Err(e) if e.code() == MF_E_NOTACCEPTING => {
//                           // MFT cannot accept input right now, need to process output first
//                           println!("MFT not accepting input, waiting for output event.");
//                           Ok(false)
//                       }
//                       Err(e) => {
//                           eprintln!("Error processing input: {:?}", e);
//                           Err(e)
//                       }
//                   }
//               }
//           }
//           Ok(None) => {
//               // Input source signaled end-of-stream
//               println!("Input source returned None (EOS), initiating drain.");
//               unsafe { self.transform.ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, 0)? };
//               self.draining = true;
//               Ok(false) // Continue to process output samples during drain
//           }
//           Err(e) => {
//               eprintln!("Error getting sample from callback: {:?}", e);
//               Err(e)
//           }
//       }
//   }

//   // Process output sample
//   fn process_output(&mut self) -> Result<()> {
//       unsafe {
//           // Get output stream info
//           let stream_info = self.transform.GetOutputStreamInfo(self.output_stream_id)?;
          
//           // Create output sample and buffer
//           let output_sample = MFCreateSample()?;
//           let output_buffer = MFCreateMemoryBuffer(stream_info.cbSize)?;
//           output_sample.AddBuffer(&output_buffer)?;

//           // Set up output buffer struct
//           let mut output_buffer_struct = MFT_OUTPUT_DATA_BUFFER {
//               dwStreamID: self.output_stream_id,
//               pSample: ManuallyDrop::new(Some(output_sample)),
//               dwStatus: 0,
//               pEvents: ManuallyDrop::new(None),
//           };

//           // Process output
//           let mut status_flags = 0u32;
//           let hr = self.transform.ProcessOutput(
//               0,
//               std::slice::from_mut(&mut output_buffer_struct),
//               &mut status_flags,
//           );

//           // Get the sample from the struct (need to take ownership)
//           let processed_sample = if !output_buffer_struct.pSample.is_none() {
//               ManuallyDrop::take(&mut output_buffer_struct.pSample)
//           } else {
//               None
//           };

//           match hr {
//               Ok(()) => {
//                   if let Some(sample) = processed_sample {
//                       // Create wrapper and call callback
//                       let output_sample = AudioEncoderOutputSample { sample };
//                       self.sample_rendered_callback.as_mut().unwrap()(output_sample)?;
//                       Ok(())
//                   } else {
//                       eprintln!("Warning: ProcessOutput returned OK but no sample.");
//                       Ok(())
//                   }
//               }
//               Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
//                   if let Some(s) = processed_sample { drop(s); }
//                   // Normal case - just need more input
//                   Ok(())
//               }
//               Err(e) => {
//                   if let Some(s) = processed_sample { drop(s); }
//                   eprintln!("ProcessOutput failed: {:?}", e);
//                   Err(e)
//               }
//           }
//       }
//   }
// }

// impl Drop for AudioEncoder {
//   fn drop(&mut self) {
//       if self.started.load(Ordering::SeqCst) {
//           println!("AudioEncoder drop: Forcing stop...");
//           self.should_stop.store(true, Ordering::SeqCst);
//           if let Err(e) = self.wait_for_completion() {
//               eprintln!("Error during AudioEncoder drop/wait: {:?}", e);
//           }
//       }
//       println!("AudioEncoder dropped.");
//   }
// }