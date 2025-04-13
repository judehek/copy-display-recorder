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
            IMFAttributes, IMFMediaEventGenerator, IMFMediaType, IMFSample,
            IMFTransform,
            // Functions
            MFCreateMemoryBuffer, MFCreateMediaType, MFCreateSample, MFStartup,
            // Constants & Enums
            METransformHaveOutput, METransformNeedInput, MFMediaType_Audio,
            MFSTARTUP_FULL, MFT_MESSAGE_COMMAND_FLUSH, MFT_MESSAGE_NOTIFY_BEGIN_STREAMING,
            MFT_MESSAGE_NOTIFY_END_OF_STREAM, MFT_MESSAGE_NOTIFY_END_STREAMING,
            MFT_MESSAGE_NOTIFY_START_OF_STREAM, MFT_OUTPUT_DATA_BUFFER,
            // Needed for ProcessOutput status
            MFT_SET_TYPE_TEST_ONLY, MF_EVENT_TYPE, MF_E_INVALIDMEDIATYPE,
            MF_E_NO_MORE_TYPES, MF_E_TRANSFORM_TYPE_NOT_SET,
            // MF_TRANSFORM_ASYNC_UNLOCK, // May still be useful
            MEDIA_EVENT_GENERATOR_GET_EVENT_FLAGS,
            // Audio Attributes (Add more as needed)
            MFAudioFormat_AAC, MFAudioFormat_PCM,
            MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE,
            MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_NUM_CHANNELS,
            MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE,
        },
        // System::Performance::{QueryPerformanceFrequency, QueryPerformanceCounter}, // Keep if needed for timing
    },
};

use crate::video::encoder_device::VideoEncoderDevice;

use super::encoder_device::AudioEncoderDevice;

// Represents one chunk of input audio data for the encoder
// Contains the IMFSample holding an IMFMediaBuffer with PCM/Float data
#[derive(Clone)]
pub struct AudioEncoderInputSample {
    sample: IMFSample,
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

pub struct AudioEncoderOutputSample {
    sample: IMFSample
}

impl AudioEncoderOutputSample {
    pub fn sample(&self) -> &IMFSample {
        &self.sample
    }
}

pub struct AudioEncoder {
    inner: Option<AudioEncoderInner>,
    output_type: IMFMediaType,
    started: AtomicBool,
    should_stop: Arc<AtomicBool>,
    encoder_thread_handle: Option<JoinHandle<Result<()>>>,
}

struct AudioEncoderInner {
    transform: IMFTransform,
    event_generator: IMFMediaEventGenerator,
    input_stream_id: u32,
    output_stream_id: u32,

    sample_requested_callback:
        Option<Box<dyn Send + FnMut() -> Result<Option<AudioEncoderInputSample>>>>,
    sample_rendered_callback: Option<Box<dyn Send + FnMut(AudioEncoderOutputSample) -> Result<()>>>,

    should_stop: Arc<AtomicBool>,
}

impl AudioEncoder {
    pub fn new(
        encoder_device: &AudioEncoderDevice,
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
        let (input_stream_ids, output_stream_ids) = {
            let mut input_stream_ids = vec![0u32; number_of_input_streams as usize];
            let mut output_stream_ids = vec![0u32; number_of_output_streams as usize];
            let result =
                unsafe { transform.GetStreamIDs(&mut input_stream_ids, &mut output_stream_ids) };
            match result {
                Ok(_) => {}
                Err(error) => {
                    // https://docs.microsoft.com/en-us/windows/win32/api/mftransform/nf-mftransform-imftransform-getstreamids
                    // This method can return E_NOTIMPL if both of the following conditions are true:
                    //   * The transform has a fixed number of streams.
                    //   * The streams are numbered consecutively from 0 to n – 1, where n is the
                    //     number of input streams or output streams. In other words, the first
                    //     input stream is 0, the second is 1, and so on; and the first output
                    //     stream is 0, the second is 1, and so on.
                    if error.code() == E_NOTIMPL {
                        for i in 0..number_of_input_streams {
                            input_stream_ids[i as usize] = i;
                        }
                        for i in 0..number_of_output_streams {
                            output_stream_ids[i as usize] = i;
                        }
                    } else {
                        return Err(error);
                    }
                }
            }
            (input_stream_ids, output_stream_ids)
        };
        let input_stream_id = input_stream_ids[0];
        let output_stream_id = output_stream_ids[0];
    
        // --- Set Output Media Type (AAC) directly --
        let output_type = unsafe {
            let output_type = MFCreateMediaType()?;
            output_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
            output_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_AAC)?;
            output_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, output_sample_rate)?;
            output_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, output_channels as u32)?;
            // Bitrate is crucial for AAC
            output_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, output_bitrate / 8)?; // Bps
    
            // Try setting the output type
            transform.SetOutputType(output_stream_id, &output_type, 0)?; // No flags
            output_type
        };
    
        // --- Set Input Media Type (PCM/Float) directly --
        let input_type: Option<IMFMediaType> = unsafe {
            let mut count = 0;
            loop {
                let result = transform.GetInputAvailableType(input_stream_id, count);
                if let Err(error) = &result {
                    if error.code() == MF_E_NO_MORE_TYPES {
                        break None;
                    }
                }

                let input_type = result?;
                let attributes: IMFAttributes = input_type.cast()?;
                input_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;
                input_type.SetGUID(&MF_MT_SUBTYPE, &MFAudioFormat_PCM)?;
                input_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, 48000)?;
                input_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, 2 as u32)?;
                input_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, 32 as u32)?;
                input_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, 4)?;
                input_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, 192000)?;
                let result = transform.SetInputType(
                    input_stream_id,
                    &input_type,
                    MFT_SET_TYPE_TEST_ONLY.0 as u32,
                );
                if let Err(error) = &result {
                    if error.code() == MF_E_INVALIDMEDIATYPE {
                        count += 1;
                        continue;
                    }
                }
                result?;
                break Some(input_type);
            }
        };
        if let Some(input_type) = input_type {
            unsafe { transform.SetInputType(input_stream_id, &input_type, 0)? };
        } else {
            return Err(Error::new(
                MF_E_TRANSFORM_TYPE_NOT_SET,
                "No suitable input type found! Try a different set of encoding settings.".into(),
            ));
        }
    
        let should_stop = Arc::new(AtomicBool::new(false));
        let inner = AudioEncoderInner {
            transform,
            event_generator,
            input_stream_id,
            output_stream_id,
    
            sample_requested_callback: None,
            sample_rendered_callback: None,

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