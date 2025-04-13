use std::{
    mem::ManuallyDrop, // Needed for ProcessOutput buffer management
    sync::{
        atomic::{AtomicBool, Ordering},
        Arc,
    },
    thread::JoinHandle,
    // time::{SystemTime, UNIX_EPOCH}, // Keep if you want debug timing prints
};

use windows::{
    core::{ComInterface, Error, Result},
    Foundation::TimeSpan,
    Win32::{
        Foundation::E_NOTIMPL,
        Media::MediaFoundation::{
            // Interfaces
            IMFAttributes, IMFMediaBuffer, IMFMediaEventGenerator, IMFMediaType, IMFSample,
            IMFTransform,
            // Functions
            MFCreateMemoryBuffer, MFCreateMediaType, MFCreateSample, MFStartup,
            // Constants & Enums
            METransformHaveOutput, METransformNeedInput, MFMediaType_Audio,
            MFSTARTUP_FULL, MFT_MESSAGE_COMMAND_FLUSH, MFT_MESSAGE_NOTIFY_BEGIN_STREAMING,
            MFT_MESSAGE_NOTIFY_END_OF_STREAM, MFT_MESSAGE_NOTIFY_END_STREAMING,
            MFT_MESSAGE_NOTIFY_START_OF_STREAM, MFT_OUTPUT_DATA_BUFFER, MFT_OUTPUT_STREAM_INFO, // Needed for buffer hints
            // Needed for ProcessOutput status
            MFT_SET_TYPE_TEST_ONLY, MF_EVENT_TYPE, MF_E_INVALIDMEDIATYPE,
            MF_E_NO_MORE_TYPES, MF_E_TRANSFORM_NEED_MORE_INPUT, MF_E_TRANSFORM_TYPE_NOT_SET,
            // MF_TRANSFORM_ASYNC_UNLOCK, // May still be useful
            MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS,
            // Audio Attributes (Add more as needed)
            MFAudioFormat_AAC, MFAudioFormat_Float, MFAudioFormat_PCM,
            MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE,
            MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_NUM_CHANNELS,
            MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE, MF_MT_USER_DATA,
        },
        // System::Performance::{QueryPerformanceFrequency, QueryPerformanceCounter}, // Keep if needed for timing
    },
};

use crate::video::encoder_device::VideoEncoderDevice;

use super::processor::AudioFormat;

// Represents one chunk of input audio data for the encoder
// Contains the IMFSample holding an IMFMediaBuffer with PCM/Float data
#[derive(Clone)]
pub struct AudioEncoderInputSample {
    sample: IMFSample, // Holds buffer + timestamp + duration
}

impl AudioEncoderInputSample {
    // Constructor now takes an IMFSample directly
    pub fn new(sample: IMFSample) -> Self {
        Self { sample }
    }

    // Helper to create a sample from raw bytes, timestamp, and duration
    // Note: Duration is important for AAC encoding
    pub fn from_raw(
        data: &[u8],
        timestamp: TimeSpan,
        duration: TimeSpan,
    ) -> Result<Self> {
        unsafe {
            let buffer = MFCreateMemoryBuffer(data.len() as u32)?;
            // Lock buffer to copy data
            let mut buffer_ptr = std::ptr::null_mut();
            let mut max_len = 0;
            buffer.Lock(&mut buffer_ptr, None, Some(&mut max_len))?;
            if max_len >= data.len() as u32 {
                 std::ptr::copy_nonoverlapping(data.as_ptr(), buffer_ptr, data.len());
            } else {
                 // Handle error: buffer too small (shouldn't happen with MFCreateMemoryBuffer)
                 buffer.Unlock()?; // Unlock before erroring
                 return Err(Error::new(windows::Win32::Foundation::E_FAIL, "Failed to lock buffer sufficiently".into()));
            }
            buffer.SetCurrentLength(data.len() as u32)?;
            buffer.Unlock()?;


            let mf_sample = MFCreateSample()?;
            mf_sample.AddBuffer(&buffer)?;
            mf_sample.SetSampleTime(timestamp.Duration)?;
            mf_sample.SetSampleDuration(duration.Duration)?;
            Ok(Self::new(mf_sample))
        }
    }

     pub fn sample(&self) -> &IMFSample {
        &self.sample
    }

    pub fn timestamp(&self) -> Result<TimeSpan> {
        let time = unsafe { self.sample.GetSampleTime()? };
        Ok(TimeSpan { Duration: time })
    }

    pub fn duration(&self) -> Result<TimeSpan> {
        let duration = unsafe { self.sample.GetSampleDuration()? };
        Ok(TimeSpan { Duration: duration })
    }
}

// Represents one chunk of encoded output audio data (e.g., an AAC frame)
pub struct AudioEncoderOutputSample {
    sample: IMFSample, // Holds buffer + timestamp + duration
}

impl AudioEncoderOutputSample {
    pub fn sample(&self) -> &IMFSample {
        &self.sample
    }
    // Add helpers to get timestamp/duration if needed
     pub fn timestamp(&self) -> Result<TimeSpan> {
        let time = unsafe { self.sample.GetSampleTime()? };
        Ok(TimeSpan { Duration: time })
    }

    pub fn duration(&self) -> Result<TimeSpan> {
        let duration = unsafe { self.sample.GetSampleDuration()? };
        Ok(TimeSpan { Duration: duration })
    }
}

// Main struct managing the audio encoding process
pub struct AudioEncoder {
    inner: Option<AudioEncoderInner>,
    output_type: IMFMediaType, // Keep track of the configured output type
    started: AtomicBool,
    should_stop: Arc<AtomicBool>,
    encoder_thread_handle: Option<JoinHandle<Result<()>>>,
}

// Holds the state needed by the encoding thread
struct AudioEncoderInner {
    transform: IMFTransform, // The actual encoder MFT (e.g., AAC encoder)
    event_generator: IMFMediaEventGenerator,
    input_stream_id: u32,
    output_stream_id: u32,

    output_buffer_size_hint: u32, // Store hint from GetOutputStreamInfo

    // Callback to request the *next sequential* audio chunk (PCM/Float)
    sample_requested_callback:
        Option<Box<dyn Send + FnMut() -> Result<Option<AudioEncoderInputSample>>>>,
    // Callback to deliver the encoded audio chunk (AAC)
    sample_rendered_callback: Option<Box<dyn Send + FnMut(AudioEncoderOutputSample) -> Result<()>>>,

    should_stop: Arc<AtomicBool>,
}

impl AudioEncoder {
    /// Creates a new AudioEncoder.
    ///
    /// # Arguments
    ///
    /// * `encoder_transform` - The instantiated audio encoder MFT (e.g., AAC).
    /// * `input_format` - The format of the raw audio data that will be fed *into* the encoder (e.g., PCM, Float).
    /// * `output_sample_rate` - Desired output sample rate (AAC supports specific rates).
    /// * `output_channels` - Desired output channels (e.g., 1 for mono, 2 for stereo).
    /// * `output_bitrate` - Desired output bitrate in bits per second (e.g., 128000).
    ///
    pub fn new(
        encoder_device: VideoEncoderDevice, // Pass in the already created MFT
        input_format: &AudioFormat,
        output_sample_rate: u32,
        output_channels: u16,
        output_bitrate: u32,
    ) -> Result<Self> {
        let transform = encoder_device.create_transform()?;
    
        // Setup MFTransform
        let event_generator: IMFMediaEventGenerator = transform.cast()?;
    
        // Get Stream IDs (usually 0 for input/output on audio encoders)
        let mut number_of_input_streams = 0;
        let mut number_of_output_streams = 0;
        unsafe {
            transform.GetStreamCount(&mut number_of_input_streams, &mut number_of_output_streams)?
        };
        // Simple MFTs often have 1 input/1 output stream
        if number_of_input_streams == 0 || number_of_output_streams == 0 {
             return Err(Error::new(windows::Win32::Foundation::E_UNEXPECTED, "Encoder MFT reported zero streams".into()));
        }
    
        let (input_stream_id, output_stream_id) = {
             // Assume stream IDs are 0 unless GetStreamIDs succeeds and returns non-zero
             let mut input_stream_ids = vec![0u32; number_of_input_streams as usize];
             let mut output_stream_ids = vec![0u32; number_of_output_streams as usize];
             match unsafe { transform.GetStreamIDs(&mut input_stream_ids, &mut output_stream_ids) } {
                 Ok(_) => (input_stream_ids[0], output_stream_ids[0]),
                 Err(e) if e.code() == E_NOTIMPL => {
                     println!("Warning: GetStreamIDs returned E_NOTIMPL, assuming stream IDs are 0.");
                     (0, 0) // Default to 0 for fixed-stream MFTs
                 }
                 Err(e) => return Err(e.into()), // Propagate other errors
             }
        };
    
        // --- Set Output Media Type (AAC) directly --
        let output_type = unsafe { MFCreateMediaType()? };
        unsafe {
            output_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
            output_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
            output_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, output_sample_rate)?;
            output_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, output_channels as u32)?;
            // Bitrate is crucial for AAC
            output_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, output_bitrate / 8)?; // Bps
    
            // Try setting the output type
            transform.SetOutputType(output_stream_id, &output_type, 0)?; // No flags
        }
    
        // --- Set Input Media Type (PCM/Float) directly --
        unsafe {
            // Create the input media type directly
            let desired_input_type = MFCreateMediaType()?;
            desired_input_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
            
            // Determine Subtype
            let subtype = if input_format.is_float {
                MFAudioFormat_Float
            } else {
                MFAudioFormat_PCM
            };
            desired_input_type.SetGUID(&MF_MT_SUBTYPE, &subtype)?;
            
            // Set audio format details
            desired_input_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, input_format.sample_rate)?;
            desired_input_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, input_format.channels as u32)?;
            desired_input_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, input_format.bits_per_sample as u32)?;
            
            // Calculate block align and avg bytes per second
            let block_align = input_format.channels as u32 * (input_format.bits_per_sample / 8) as u32;
            let avg_bytes_per_sec = input_format.sample_rate * block_align;
            
            desired_input_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, block_align)?;
            desired_input_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, avg_bytes_per_sec)?;
    
            // Find a compatible input type supported by the encoder
            let mut i = 0;
            let mut input_type_to_set: Option<IMFMediaType> = None;
            loop {
                match transform.GetInputAvailableType(input_stream_id, i) {
                    Ok(available_type) => {
                        // Check if our desired type is compatible
                        match transform.SetInputType(input_stream_id, &desired_input_type, MFT_SET_TYPE_TEST_ONLY.0 as u32) {
                            Ok(_) => {
                                // Found a compatible type (our desired type itself)
                                input_type_to_set = Some(desired_input_type.clone());
                                break;
                            }
                            Err(e) if e.code() == MF_E_INVALIDMEDIATYPE => {
                                // Our desired type isn't directly supported in this exact form.
                                i += 1;
                            }
                            Err(e) => return Err(e.into()), // Other error during test set
                        }
                        // Explicitly drop `available_type` COM ptr
                        drop(available_type);
                    }
                    Err(e) if e.code() == MF_E_NO_MORE_TYPES => {
                        // Exhausted all available types, none matched.
                        return Err(Error::new(MF_E_TRANSFORM_TYPE_NOT_SET,
                            format!("Audio encoder does not support the provided input format: {:?}", input_format).into()));
                    }
                    Err(e) => return Err(e.into()), // Error getting available type
                }
            }
    
            // Actually set the validated input type
            transform.SetInputType(input_stream_id, &input_type_to_set.unwrap(), 0)?; // No flags
        }
    
        // Get output stream info *after* types are set
        let stream_info = unsafe { transform.GetOutputStreamInfo(output_stream_id)? };
        // cbSize gives a hint for buffer allocation. AAC frames are variable size,
        // but this provides a reasonable upper bound or typical size.
        let output_buffer_size_hint = if stream_info.cbSize > 0 { stream_info.cbSize } else { 4096 }; // Default fallback
    
        let should_stop = Arc::new(AtomicBool::new(false));
        let inner = AudioEncoderInner {
            transform, // Move transform ownership
            event_generator,
            input_stream_id,
            output_stream_id,
            output_buffer_size_hint,
    
            sample_requested_callback: None, // To be set later
            sample_rendered_callback: None,  // To be set later
    
            should_stop: should_stop.clone(),
        };
    
        Ok(Self {
            inner: Some(inner),
            output_type, // Store the final output type
            started: AtomicBool::new(false),
            should_stop,
            encoder_thread_handle: None,
        })
    }

    // --- Methods for starting, stopping, setting callbacks ---
    // (These are largely identical to VideoEncoder, just replace type names)

    pub fn try_start(&mut self) -> Result<bool> {
        let mut result = false;
        if self.started.compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst).is_ok()
        {
            let mut inner = self.inner.take().expect("Encoder inner state missing unexpectedly");

            // Callbacks must both be set
            if inner.sample_rendered_callback.is_none() || inner.sample_requested_callback.is_none()
            {
                panic!("Sample requested and rendered callbacks must be set before starting audio encoder");
            }

            // Start a seperate thread to drive the transform
            self.encoder_thread_handle = Some(std::thread::spawn(move || -> Result<()> {
                // Initialize COM for this thread (MFTs often require MTA or STA)
                // Using MTA is generally safer for background threads unless MFT requires STA
                // Or rely on MFStartup/MFShutdown to potentially handle this.
                // CoInitializeEx(None, COINIT_MULTITHREADED)?; // Example MTA init
                unsafe { MFStartup(crate::media::MF_VERSION, MFSTARTUP_FULL)? }; // Use MF_VERSION

                let encode_result = inner.encode(); // Call the encoding loop

                // Shutdown Media Foundation for this thread
                // unsafe { MFShutdown()? }; // MFShutdown is global, call once at app exit
                // CoUninitialize(); // Match CoInitializeEx

                if encode_result.is_err() {
                    println!("Audio encoding stopped unexpectedly: {:?}", encode_result);
                }
                encode_result // Return the result
            }));
            result = true;
        }
        Ok(result)
    }

    pub fn stop(&mut self) -> Result<()> {
        if self.started.load(Ordering::SeqCst) {
            assert!(self
                .should_stop
                .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
                .is_ok());
            self.wait_for_completion()?;
        }
        Ok(())
    }

    fn wait_for_completion(&mut self) -> Result<()> {
        let handle = self.encoder_thread_handle.take().unwrap();
        handle.join().unwrap()
    }

    // Update callback signatures to use Audio types
    pub fn set_sample_requested_callback<
        F: 'static + Send + FnMut() -> Result<Option<AudioEncoderInputSample>>,
    >(
        &mut self,
        callback: F,
    ) {
        self.inner.as_mut().unwrap().sample_requested_callback = Some(Box::new(callback));
    }

    pub fn set_sample_rendered_callback<
        F: 'static + Send + FnMut(AudioEncoderOutputSample) -> Result<()>,
    >(
        &mut self,
        callback: F,
    ) {
        self.inner.as_mut().unwrap().sample_rendered_callback = Some(Box::new(callback));
    }

    // Accessor for the configured output media type
    pub fn output_type(&self) -> &IMFMediaType {
        &self.output_type
    }
}

// Mark inner struct as Send (assuming IMFTransform and other COM objects are thread-safe
// when used correctly - which MF generally allows for MFTs following standard patterns)
unsafe impl Send for AudioEncoderInner {}

// Rename constants for clarity if desired, or keep using MF ones
const AUDIO_ENGINE_TRANFORM_NEED_INPUT: MF_EVENT_TYPE = METransformNeedInput;
const AUDIO_ENGINE_TRANFORM_HAVE_OUTPUT: MF_EVENT_TYPE = METransformHaveOutput;

impl AudioEncoderInner {
    // The core encoding loop run by the dedicated thread
    fn encode(&mut self) -> Result<()> {
        unsafe {
            self.transform
                .ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_START_OF_STREAM, 0)?;

                let mut should_exit = false;
                while !should_exit {
                    let event = self
                        .event_generator
                        .GetEvent(MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS(0))?;
        
                    let event_type = MF_EVENT_TYPE(event.GetType()? as i32);
                    match event_type {
                        AUDIO_ENGINE_TRANFORM_NEED_INPUT => {
                            should_exit = self.on_transform_input_requested()?;
                        }
                        AUDIO_ENGINE_TRANFORM_HAVE_OUTPUT => {
                            self.on_transform_output_ready()?;
                        }
                    _ => {
                        panic!("Unknown media event type: {}", event_type.0);
                    }
                }
            }

            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_END_OF_STREAM, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_NOTIFY_END_STREAMING, 0)?;
            self.transform
                .ProcessMessage(MFT_MESSAGE_COMMAND_FLUSH, 0)?;
        }
        Ok(())
    }

    fn on_transform_input_requested(&mut self) -> Result<bool> {
        let mut should_exit = true;
        if !self.should_stop.load(Ordering::SeqCst) {
            if let Some(sample) = self.sample_requested_callback.as_mut().unwrap()()? {
                unsafe {
                    self.transform
                        .ProcessInput(self.input_stream_id, sample.sample(), 0)?;
                }
                should_exit = false;
            }
        }
        Ok(should_exit)
    }

    // Called when the MFT signals METransformHaveOutput
    fn on_transform_output_ready(&mut self) -> Result<()> {
        let mut status = 0;
        let output_buffer = MFT_OUTPUT_DATA_BUFFER {
            dwStreamID: self.output_stream_id,
            ..Default::default()
        };

        let sample = unsafe {
            let mut output_buffers = [output_buffer];
            self.transform
                .ProcessOutput(0, &mut output_buffers, &mut status)?;
            output_buffers[0].pSample.as_ref().unwrap().clone()
        };

        let output_sample = AudioEncoderOutputSample { sample };
        self.sample_rendered_callback.as_mut().unwrap()(output_sample)?;
        Ok(())
    }
}