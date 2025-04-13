use windows::{
    core::{ComInterface, Result},
    Win32::Media::MediaFoundation::{
        IMFActivate, IMFTransform, MFMediaType_Audio, // Changed from Video
        MFAudioFormat_AAC,                             // Changed from H264
        MFT_CATEGORY_AUDIO_ENCODER,                  // Changed from Video Encoder
        MFT_FRIENDLY_NAME_Attribute,
        MFT_ENUM_FLAG_SORTANDFILTER,
        MFT_ENUM_FLAG_TRANSCODE_ONLY,
        MFT_REGISTER_TYPE_INFO,
    },
};

// Assuming these helper functions exist in your crate::media module
// and are generic enough to work with different MFT categories and attributes.
use crate::media::{enumerate_mfts, get_string_attribute};

#[derive(Clone)]
pub struct AudioEncoderDevice {
    source: IMFActivate,
    display_name: String,
}

impl AudioEncoderDevice {
    /// Enumerates available hardware AAC audio encoders.
    ///
    /// Note: Hardware *audio* encoders are less common than hardware video encoders.
    /// This function might return an empty list on many systems.
    pub fn enumerate() -> Result<Vec<AudioEncoderDevice>> {
        // Define the desired output type: Audio / AAC
        let output_info = MFT_REGISTER_TYPE_INFO {
            guidMajorType: MFMediaType_Audio, // Use Audio type
            guidSubtype: MFAudioFormat_AAC, // Specify AAC format
        };

        // Enumerate MFTs in the Audio Encoder category
        // We keep the same flags (Hardware, TranscodeOnly, SortAndFilter)
        // for consistency, although Hardware audio encoders might be rare.
        let encoders = enumerate_mfts(
            &MFT_CATEGORY_AUDIO_ENCODER, // Use Audio Encoder category
            MFT_ENUM_FLAG_TRANSCODE_ONLY | MFT_ENUM_FLAG_SORTANDFILTER,
            None, // No specific input requirement
            Some(&output_info), // Require AAC output support
        )?;

        let mut encoder_devices = Vec::new();
        for encoder in encoders {
            // Get the friendly name, same attribute key works
            let display_name = if let Some(display_name) =
                get_string_attribute(&encoder.cast()?, &MFT_FRIENDLY_NAME_Attribute)?
            {
                display_name
            } else {
                "Unknown Audio Encoder".to_owned() // Slightly more specific default
            };

            let encoder_device = AudioEncoderDevice {
                source: encoder,
                display_name,
            };
            encoder_devices.push(encoder_device);
        }
        Ok(encoder_devices)
    }

    /// Returns the human-readable display name of the audio encoder.
    pub fn display_name(&self) -> &str {
        &self.display_name
    }

    /// Creates and activates the underlying Media Foundation Transform (MFT) for this encoder.
    pub fn create_transform(&self) -> Result<IMFTransform> {
        // Activation process is the same
        unsafe { self.source.ActivateObject() }
    }
}