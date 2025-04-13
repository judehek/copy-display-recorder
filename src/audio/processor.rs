use windows::{
    core::{implement, ComInterface, Interface, Result, GUID, HRESULT, PCWSTR},
    Win32::{
        Foundation::{CloseHandle, DECIMAL, HANDLE, S_OK},
        Media::MediaFoundation::{
            // Interfaces
            IMFActivate, IMFCollection, IMFMediaBuffer, IMFMediaType, IMFSample, IMFTransform, MFAudioConstriction, MFAudioFormat_Float, MFAudioFormat_PCM, MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFMediaType_Audio, MFShutdown, MFStartup, MFVideoInterlace_Progressive, MFT_OUTPUT_DATA_BUFFER, MFT_OUTPUT_STREAM_INFO, MF_E_TRANSFORM_NEED_MORE_INPUT, MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE, MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_CHANNEL_MASK, MF_MT_AUDIO_NUM_CHANNELS, MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_INTERLACE_MODE, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE, MF_VERSION
            // MFPKEY_WMRESAMP_CHANNELMTX, // Add if custom matrix needed
            // MFPKEY_WMRESAMP_LOWPASS_BANDWIDTH, // Add if needed
        },
        System::{
            Com::{
                CoCreateInstance, CoInitializeEx, CoUninitialize,
                // Need PROPVARIANT definition
                StructuredStorage::{PropVariantClear, PROPVARIANT},
                CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED, STGM_READ
            },
            // Need VARIANT type constants like VT_UI4
            Variant::VT_UI4,
        },
        // Required for Property Store
        UI::Shell::PropertiesSystem::{IPropertyStore, PROPERTYKEY},
    },
};

use std::mem::{ManuallyDrop};

// CLSID for the Resampler DSP
const CLSID_CResamplerMediaObject: GUID = GUID::from_u128(0xf447b69e_1884_4a7e_8055_346f74d6edb3);


// Define a simple structure to hold audio format details
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AudioFormat {
    pub sample_rate: u32,
    pub channels: u16,
    pub bits_per_sample: u16,
    pub channel_mask: Option<u32>,
    pub format: GUID,
}

impl AudioFormat {
    // Calculate derived values needed for WAVEFORMATEX / MF_MT_*
    fn block_align(&self) -> u16 {
        self.channels * (self.bits_per_sample / 8)
    }

    fn avg_bytes_per_second(&self) -> u32 {
        self.sample_rate * self.block_align() as u32
    }
}

pub struct AudioProcessor {
    resampler_transform: IMFTransform,
    input_media_type: IMFMediaType,
    output_media_type: IMFMediaType,
    input_stream_id: u32,
    output_stream_id: u32,
    output_buffer_size: u32,
}

impl AudioProcessor {
    pub fn new(input_format: AudioFormat, output_format: AudioFormat, quality: Option<u32>) -> Result<Self> {
        // Create the Resampler DSP instance
        let resampler_transform: IMFTransform = unsafe {
            CoCreateInstance(&CLSID_CResamplerMediaObject, None, CLSCTX_INPROC_SERVER)?
        };

        // Create Media Types
        let input_media_type = create_audio_media_type(&input_format)?;
        let output_media_type = create_audio_media_type(&output_format)?;

        // Set Media Types on the Transform
        unsafe {
            resampler_transform.SetInputType(0, Some(&input_media_type), 0)?; // Stream ID 0, no flags
            resampler_transform.SetOutputType(0, Some(&output_media_type), 0)?; // Stream ID 0, no flags
        }

        // Get Stream IDs (usually 0 for input and 0 for output for simple MFTs like resampler)
        // The windows crate bindings for GetStreamIDs take slices for potential multiple IDs
        let mut input_ids = [0u32; 1]; // Expecting one input stream
        let mut output_ids = [0u32; 1]; // Expecting one output stream
        let input_stream_id;
        let output_stream_id;

        unsafe {
             // Note: GetStreamIDs might not be implemented by all MFTs. Defaulting to 0 is common.
             match resampler_transform.GetStreamIDs(&mut input_ids[..], &mut output_ids[..]) {
                Ok(_) => {
                    // Check if we actually got IDs back. The function might succeed but return 0 elements.
                    // This specific MFT usually uses 0 for both.
                    input_stream_id = input_ids[0];
                    output_stream_id = output_ids[0];
                },
                 Err(e) => {
                     // Don't fail completely, try assuming 0, but warn
                     println!("Warning: IMFTransform::GetStreamIDs failed ({:?}), assuming stream IDs are 0.", e.code());
                     input_stream_id = 0;
                     output_stream_id = 0;
                     // Alternatively, return Err(e.into()); if strict adherence is needed
                 }
             }
        }


        // Get output stream info to estimate buffer size needed
        let stream_info = unsafe { resampler_transform.GetOutputStreamInfo(output_stream_id)? };
        // cbSize gives a hint, might need adjustment or dynamic allocation in ProcessOutput loop
        // Ensure we have a sane minimum if cbSize is 0 (which can happen)
        let output_buffer_size = if stream_info.cbSize > 0 { stream_info.cbSize } else { 4096 } ;


        // Optional: Set Resampler Quality Property
        if let Some(q) = quality {
            if q > 60 { // Quality is 1-60
                 println!("Warning: Resampler quality ({}) out of range (1-60). Clamping to 60.", q);
                 // q = 60; // Or return error
            }
            let props: IPropertyStore = resampler_transform.cast()?;
            unsafe {
                // Fix E0308: Use PROPVARIANT instead of VARIANT
                // Initialize PROPVARIANT correctly for VT_UI4
                let mut propvar = PROPVARIANT::default();

                // The PROPVARIANT structure in windows-rs uses nested anonymous structs/unions.
                // We need to access the correct fields to set vt and the value.
                // Access path: propvar.Anonymous.Anonymous.vt & propvar.Anonymous.Anonymous.Anonymous.ulVal
                let prop_union = &mut propvar.Anonymous.Anonymous;
                prop_union.vt = windows::Win32::System::Variant::VARENUM(VT_UI4.0 as u16); // VARENUM is a struct, use .0 for the value
                prop_union.Anonymous.ulVal = q; // Set the u32 value

                // MFPKEY_WMRESAMP_FILTERQUALITY defined in Media Foundation headers, lookup its GUID/pid
                // Found via search: {AF15E843-2AA5-4AA3-8694-33814B4DEE6B}, 6
                const MFPKEY_WMRESAMP_FILTERQUALITY_GUID: GUID = GUID::from_u128(0xaf15e843_2aa5_4aa3_8694_33814b4dee6b);
                let pkey = PROPERTYKEY {
                    fmtid: MFPKEY_WMRESAMP_FILTERQUALITY_GUID,
                    pid: 6, // pid for MFPKEY_WMRESAMP_FILTERQUALITY
                };

                 // Pass a pointer to the PROPVARIANT. `&propvar` coerces to `*const PROPVARIANT`.
                 match props.SetValue(&pkey, &propvar) {
                    Ok(_) => {},
                    Err(e) => println!("Warning: Failed to set resampler quality: {:?}", e),
                 };

                 // No explicit PropVariantClear needed. The `PROPVARIANT` type from the
                 // `windows` crate should implement `Drop` and call `PropVariantClear` automatically
                 // if necessary (e.g., if it held allocated memory or COM pointers).
                 // For VT_UI4, clear is likely a no-op, but relying on Drop is correct.
            }
        }

        Ok(Self {
            resampler_transform,
            input_media_type,
            output_media_type,
            input_stream_id,
            output_stream_id,
            output_buffer_size,
        })
    }

     /// Processes a single input audio sample and returns the corresponding output sample(s).
    ///
    /// The input `sample` should contain the audio data and have the correct timestamp and duration set.
    /// Returns `Ok(Some(IMFSample))` if output was produced, `Ok(None)` if more input is needed,
    /// or `Err` on failure.
    /// Note: One input sample might not produce one output sample immediately due to buffering or rate changes.
    /// This function currently simplifies and expects one output per input call that yields output.
    /// A more robust implementation might need to return a Vec<IMFSample> or handle multiple ProcessOutput calls.
    pub fn process_sample(&mut self, input_sample: &IMFSample) -> Result<Option<IMFSample>> {
        unsafe {
            // 1. Send input to the transform
            match self.resampler_transform.ProcessInput(self.input_stream_id, input_sample, 0) {
                Ok(_) => {}, // Input accepted
                Err(e) => {
                    // Handle specific errors if needed, e.g., MF_E_NOTACCEPTING
                    println!("ProcessInput failed: {:?}", e);
                    return Err(e.into());
                }
            }

            // 2. Try to get output
            loop {
                 // Create necessary structures for ProcessOutput

                 // Try to create a buffer of the suggested size.
                 let buffer_size = self.output_buffer_size; // Use size from GetOutputStreamInfo
                 // Fix E0061: MFCreateMemoryBuffer returns Result<IMFMediaBuffer>
                 let output_buffer = MFCreateMemoryBuffer(buffer_size)?;

                 // Fix E0061: MFCreateSample returns Result<IMFSample>
                 let output_sample = MFCreateSample()?;
                 output_sample.AddBuffer(&output_buffer)?;


                 // Fix E0308, E0061: Correctly setup MFT_OUTPUT_DATA_BUFFER and call ProcessOutput
                 let mut output_data_buffer = MFT_OUTPUT_DATA_BUFFER {
                     dwStreamID: self.output_stream_id,
                     // Fix E0308: Use ManuallyDrop<Option<IMFSample>>
                     // We wrap the sample we created. If ProcessOutput succeeds, it takes ownership.
                     // If it fails, we must take it back out of ManuallyDrop and drop it ourselves.
                     pSample: ManuallyDrop::new(Some(output_sample)),
                     dwStatus: 0, // Will be filled by ProcessOutput
                     // Fix E0308: Use ManuallyDrop<Option<IMFCollection>>
                     pEvents: ManuallyDrop::new(None), // We don't process events here
                 };
                 let mut process_output_status: u32 = 0;

                 // Create a slice as expected by ProcessOutput
                 let mut output_buffers = [output_data_buffer]; // Array of size 1


                 match self.resampler_transform.ProcessOutput(
                     0, // Flags
                     // Fix E0061: Pass a mutable slice
                     &mut output_buffers, // Array of buffers
                     &mut process_output_status, // Status flags
                 ) {
                     Ok(_) => {
                         // Success! We got an output sample.
                         // Retrieve the sample from the struct.
                         // Fix E0599: Use ManuallyDrop::take to regain ownership
                         // Since ProcessOutput succeeded, it *should* have kept ownership,
                         // but the API design means pSample potentially points to the *output* sample,
                         // which might even be different from the one we passed in (though unlikely here).
                         // We take it from ManuallyDrop to prevent a double-free if it's the same COM object,
                         // and to get the handle the MFT wants us to use.
                         let filled_sample_option = ManuallyDrop::take(&mut output_buffers[0].pSample);

                         // Ensure the MFT actually provided a sample
                         let processed_sample = filled_sample_option.ok_or_else(|| {
                            println!("Error: ProcessOutput succeeded but returned NULL sample pointer.");
                            windows::core::Error::new(HRESULT(0x8000FFFFu32 as i32), "ProcessOutput succeeded but returned null sample".into()) // E_UNEXPECTED
                         })?;

                         // Update buffer length (important!)
                         // Fix E0061: GetCurrentLength returns Result<u32>
                         // Need to get the buffer from the potentially new sample
                         let filled_buffer: IMFMediaBuffer = processed_sample.GetBufferByIndex(0)?;
                         let current_length = filled_buffer.GetCurrentLength()?;
                         filled_buffer.SetCurrentLength(current_length)?; // Ensure length is set

                         // TODO: Handle MFT_PROCESS_OUTPUT_STATUS flags if needed (e.g., MFT_PROCESS_OUTPUT_STATUS_NEW_STREAMS)

                         return Ok(Some(processed_sample));
                     }
                     Err(e) if e.code() == MF_E_TRANSFORM_NEED_MORE_INPUT => {
                         // Transform needs more input data before it can produce output.
                         // We need to drop the IMFSample we created and put in MFT_OUTPUT_DATA_BUFFER,
                         // as ProcessOutput failed and didn't take ownership.
                         // Fix E0599, E0599: Use ManuallyDrop::take and let the Option drop naturally.
                         let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                         drop(sample_to_drop); // This drops the Option<IMFSample>, releasing the COM object.

                         // Also drop pEvents if it were used
                         let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                         drop(events_to_drop);

                         return Ok(None);
                    }
                    Err(e) => {
                         // Any other error
                         println!("ProcessOutput failed: {:?}", e);
                         // Clean up the sample we allocated, similar to the NEED_MORE_INPUT case.
                         // Fix E0599, E0599: Use ManuallyDrop::take and let the Option drop naturally.
                         let sample_to_drop = ManuallyDrop::take(&mut output_buffers[0].pSample);
                         drop(sample_to_drop); // This drops the Option<IMFSample>, releasing the COM object.

                         // Also drop pEvents if it were used
                         let events_to_drop = ManuallyDrop::take(&mut output_buffers[0].pEvents);
                         drop(events_to_drop);

                         return Err(e.into());
                     }
                 }

                 // Note: A real implementation might need to loop ProcessOutput if the MFT signals it can produce more
                 // without more input (e.g., status flags MFT_OUTPUT_DATA_BUFFER_INCOMPLETE or if ProcessOutput returns S_OK
                 // but MF_E_TRANSFORM_NEED_MORE_INPUT wasn't returned), but the audio resampler usually produces output
                 // fairly directly after sufficient input. Breaking the loop here simplifies.
                 // break; // Removed, loop continues implicitly until return
             }
        }
    }

    // Optional: Provide access to the configured output media type if needed
    pub fn output_media_type(&self) -> &IMFMediaType {
        &self.output_media_type
    }
}

// Helper function to create an IMFMediaType for uncompressed audio
fn create_audio_media_type(format: &AudioFormat) -> Result<IMFMediaType> {
    unsafe {
        // Fix E0061: MFCreateMediaType returns Result<IMFMediaType>
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
        // MF_INTERLACE_MODE_PROGRESSIVE is an enum variant, access its value with .0
        media_type.SetUINT32(&MF_MT_INTERLACE_MODE, MFVideoInterlace_Progressive.0 as u32)?;

        // Required for some transforms/sinks, especially multichannel
        // Use standard WAVEFORMATEX channel masks (ksmedia.h definitions, often mapped in windows crate or define manually)
        // Common speaker configurations:
        const SPEAKER_FRONT_LEFT: u32 = 0x1;
        const SPEAKER_FRONT_RIGHT: u32 = 0x2;
        const SPEAKER_FRONT_CENTER: u32 = 0x4;
        // ... other standard speaker positions ...

        let mut mask_to_set : Option<u32> = format.channel_mask;

        if mask_to_set.is_none() {
            // Provide default masks for mono/stereo if not specified
             match format.channels {
                 1 => mask_to_set = Some(SPEAKER_FRONT_CENTER), // Standard mono
                 2 => mask_to_set = Some(SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT), // Standard stereo
                 _ => {
                     // No standard default for > 2 channels without explicit mask
                     // MFT might fail later if mask is required and not provided.
                     // Consider returning an error or logging a warning here if format.channels > 2.
                     println!("Warning: Creating audio media type for {} channels without an explicit channel mask.", format.channels);
                 }
             };
        }

        if let Some(mask) = mask_to_set {
             if mask != 0 {
                 media_type.SetUINT32(&MF_MT_AUDIO_CHANNEL_MASK, mask)?;
             } else if format.channels > 0 {
                 // A mask of 0 is generally invalid for WAVEFORMATEXTENSIBLE if channels > 0
                 println!("Warning: Audio channel mask is 0 for {} channels. This might be invalid.", format.channels);
                 // Depending on MFT requirements, you might need to fail here or let the MFT validate.
             }
         }
         // Note: The documentation implies MF_MT_AUDIO_CHANNEL_MASK might be required for WAVEFORMATEXTENSIBLE subtypes
         // (which PCM/Float often imply when channels > 2 or bits > 16). Ensure it's set correctly.

        Ok(media_type)
    }
}