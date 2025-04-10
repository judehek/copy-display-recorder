// src/audio.rs
use std::{
    ffi::c_void,
    ptr::NonNull, // Not used directly, but good to know it's the underlying issue for Send errors
    sync::{
        atomic::{AtomicBool, Ordering},
        mpsc::{self, SyncSender}, // Now sends AudioDataPacket
        Arc,
    },
    thread::{self, JoinHandle},
    time::Duration, // Keep for potential future use, though not used currently
};

use clap::ValueEnum;
use windows::{
    core::{Interface, Result, HSTRING}, // Interface needed for e.g. audio_client.GetService
    Win32::{
        Foundation::{CloseHandle, E_FAIL, E_NOTIMPL, HANDLE, WAIT_OBJECT_0, WAIT_TIMEOUT},
        Media::{
            Audio::{
                eCapture, eConsole, eRender, Endpoints::IAudioEndpointVolume, IAudioCaptureClient, IAudioClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_S_BUFFER_EMPTY, WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVE_FORMAT_PCM // Added for GetDefaultAudioEndpoint role
            },
            KernelStreaming::KSDATAFORMAT_SUBTYPE_PCM,
            MediaFoundation::{
                IMFMediaBuffer, IMFMediaType, IMFSample, MFAudioFormat_Float, MFAudioFormat_PCM,
                MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFMediaType_Audio,
                MF_E_INVALIDREQUEST, // Potential error from GetCurrentLength (though not used here)
                MF_MT_AUDIO_AVG_BYTES_PER_SECOND, MF_MT_AUDIO_BITS_PER_SAMPLE,
                MF_MT_AUDIO_BLOCK_ALIGNMENT, MF_MT_AUDIO_CHANNEL_MASK, MF_MT_AUDIO_NUM_CHANNELS,
                MF_MT_AUDIO_SAMPLES_PER_SECOND, MF_MT_MAJOR_TYPE, MF_MT_SUBTYPE,
            },
            Multimedia::{KSDATAFORMAT_SUBTYPE_IEEE_FLOAT, WAVE_FORMAT_IEEE_FLOAT},
        },
        System::{
            Com::{
                CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_ALL,
                COINIT_MULTITHREADED,
            },
            Performance::{QueryPerformanceCounter, QueryPerformanceFrequency},
            Threading::{CreateEventW, SetEvent, WaitForSingleObject}, // SetEvent not used currently but keep
        },
    },
};

// --- Data Structure for Thread Communication ---
// Holds the essential data extracted from WASAPI buffers.
// Inherently Send because Vec<u8> and i64 are Send.
#[derive(Debug)]
pub struct AudioDataPacket {
    pub data: Option<Vec<u8>>, // Audio bytes (None if silent)
    pub timestamp: i64,        // MF Time (100ns units)
    pub duration: i64,         // MF Duration (100ns units)
    pub flags: u32,            // Optional: carry over flags like discontinuity if needed
}
// --- End Data Structure ---

// Parameters passed TO the audio thread. This MUST be Send.
struct AudioThreadParams {
    device_id: HSTRING,
    source: AudioSource,
    wave_format: WAVEFORMATEXTENSIBLE,
    sample_sender: SyncSender<(u32, AudioDataPacket)>,
    audio_stream_index: u32,
    qpc_frequency: i64,
    stop_signal: Arc<AtomicBool>,
}

// Main structure for managing audio capture configuration and control
pub struct AudioCapture {
    device_id: HSTRING,
    wave_format: WAVEFORMATEXTENSIBLE,
    source: AudioSource,
    audio_stream_index: u32,
    qpc_frequency: i64,
    media_type: IMFMediaType,
    thread_handle: Option<JoinHandle<Result<()>>>,
    stop_signal: Arc<AtomicBool>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, ValueEnum)]
pub enum AudioSource {
    DefaultLoopback,
    DefaultMicrophone,
    None,
}

impl AudioCapture {
    pub fn new(source: AudioSource, audio_stream_index: u32) -> Result<Self> {
        if source == AudioSource::None {
            return Err(windows::core::Error::new(
                E_NOTIMPL,
                "Audio capture disabled".into(),
            ));
        }

        unsafe {
            // Initialize COM for the constructor thread (usually main)
            // It's okay to call this multiple times on the same thread.
            CoInitializeEx(None, COINIT_MULTITHREADED)?; // Use MTA

            let enumerator: IMMDeviceEnumerator =
                CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

            let data_flow = match source {
                AudioSource::DefaultLoopback => eRender,
                AudioSource::DefaultMicrophone => eCapture,
                AudioSource::None => unreachable!(),
            };
            let device: IMMDevice = enumerator.GetDefaultAudioEndpoint(data_flow, eConsole)?; // Use Console role
            let device_id_pwstr = device.GetId()?; // Returns Result<PWSTR>
            let device_id: HSTRING = device_id_pwstr.to_string()?.into(); // Convert PWSTR -> HSTRING (copies the string)

            // Activate temporarily only to get the mix format
            let temp_audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
            let wave_format_ptr = temp_audio_client.GetMixFormat()?;
// --- START MODIFICATION ---

            // 1. Activate temporarily to get device's preferred characteristics
            let temp_audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
            let mix_format_ptr = temp_audio_client.GetMixFormat()?;

            // 2. Extract desired sample rate and channel layout from the mix format
            let mix_wave_format = if (*mix_format_ptr).cbSize >= (std::mem::size_of::<WAVEFORMATEXTENSIBLE>() - std::mem::size_of::<WAVEFORMATEX>()) as u16 {
                *(mix_format_ptr as *const WAVEFORMATEXTENSIBLE)
            } else {
                 // Handle basic WAVEFORMATEX if necessary (less likely for modern mix formats)
                 let base_format = *(mix_format_ptr as *const WAVEFORMATEX);
                 println!("Warning: Basic WAVEFORMATEX detected as mix format, extracting basic info.");
                 WAVEFORMATEXTENSIBLE {
                    Format: base_format,
                    Samples: windows::Win32::Media::Audio::WAVEFORMATEXTENSIBLE_0 { wValidBitsPerSample: base_format.wBitsPerSample },
                    dwChannelMask: match base_format.nChannels {
                        1 => 0x4, 2 => 0x3, _ => 0
                    },
                    SubFormat: if u32::from(base_format.wFormatTag) == WAVE_FORMAT_IEEE_FLOAT { KSDATAFORMAT_SUBTYPE_IEEE_FLOAT }
                               else if u32::from(base_format.wFormatTag) == WAVE_FORMAT_PCM { KSDATAFORMAT_SUBTYPE_PCM }
                               else { windows::core::GUID::zeroed() },
                 }
            };
            let target_sample_rate = mix_wave_format.Format.nSamplesPerSec;
            let target_channels = mix_wave_format.Format.nChannels;
            let target_channel_mask = mix_wave_format.dwChannelMask; // Use mask from mix format

            // Don't forget to free the pointer from GetMixFormat
            CoTaskMemFree(Some(mix_format_ptr as *const c_void));
            drop(temp_audio_client); // Drop the temporary client

            // 3. Construct the *desired* WAVEFORMATEXTENSIBLE for PCM16
            let target_bits_per_sample: u16 = 16;
            let target_block_align = target_channels * (target_bits_per_sample / 8);
            let target_avg_bytes_per_sec = target_sample_rate * (target_block_align as u32);

            // Ensure we have a valid channel mask if the mix format didn't provide one
            let final_channel_mask = if target_channel_mask != 0 {
                target_channel_mask
            } else {
                // Fallback masks for common channel counts
                match target_channels {
                    1 => 0x4, // SPEAKER_FRONT_CENTER
                    2 => 0x3, // SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT
                    _ => {
                        eprintln!("Warning: Cannot determine channel mask for {} channels. Using 0.", target_channels);
                        0
                    }
                }
            };

            let desired_wave_format = WAVEFORMATEXTENSIBLE {
                Format: WAVEFORMATEX {
                    wFormatTag: windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE as u16,
                    nChannels: target_channels,
                    nSamplesPerSec: target_sample_rate,
                    nAvgBytesPerSec: target_avg_bytes_per_sec,
                    nBlockAlign: target_block_align,
                    wBitsPerSample: target_bits_per_sample,
                    cbSize: (std::mem::size_of::<WAVEFORMATEXTENSIBLE>() - std::mem::size_of::<WAVEFORMATEX>()) as u16,
                },
                Samples: windows::Win32::Media::Audio::WAVEFORMATEXTENSIBLE_0 {
                    wValidBitsPerSample: target_bits_per_sample,
                },
                dwChannelMask: final_channel_mask,
                SubFormat: KSDATAFORMAT_SUBTYPE_PCM, // Explicitly request PCM
            };


            // Create the IMFMediaType needed by the SinkWriter/Consumer
            let media_type = Self::create_mf_media_type(&desired_wave_format)?;

            // Get QPC Frequency for timestamp calculations
            let mut frequency = 0;
            QueryPerformanceFrequency(&mut frequency)?;
            if frequency <= 0 { // Frequency should be positive
                return Err(windows::core::Error::new(E_FAIL,"Invalid QPC frequency".into()));
            }

            let stop_signal = Arc::new(AtomicBool::new(false));

            Ok(Self {
                device_id,
                wave_format: desired_wave_format,
                source,
                audio_stream_index,
                qpc_frequency: frequency,
                media_type,
                thread_handle: None, // Thread started via start()
                stop_signal,
            })
        }
    }

    // Helper to create IMFMediaType from WAVEFORMATEXTENSIBLE
    fn create_mf_media_type(wf: &WAVEFORMATEXTENSIBLE) -> Result<IMFMediaType> {
        unsafe {
            let media_type = MFCreateMediaType()?;

            media_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;

            let sub_format = wf.SubFormat;

            let subtype = if sub_format == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
                println!("MF Media Type Subtype: MFAudioFormat_Float");
                &MFAudioFormat_Float
            } else if sub_format == KSDATAFORMAT_SUBTYPE_PCM {
                // Also copy wBitsPerSample to a local variable
                let bits_per_sample = wf.Format.wBitsPerSample;
                println!("MF Media Type Subtype: MFAudioFormat_PCM (Bits: {})", bits_per_sample);
                
                if bits_per_sample == 16 {
                    &MFAudioFormat_PCM
                } else {
                    // Add support for other PCM bit depths if needed (e.g., 8, 24, 32)
                    eprintln!("Warning: Unsupported PCM bit depth {} for MF Media Type", bits_per_sample);
                    return Err(windows::core::Error::new(E_NOTIMPL, "Unsupported PCM bit depth for MF Media Type".into()));
                }
            } else {
                // Use the copied SubFormat in the error message too
                eprintln!("Warning: Unsupported audio SubFormat for MF Media Type: {:?}. Defaulting to Float.", sub_format);
                // Fallback might cause issues downstream if the actual data isn't float
                &MFAudioFormat_Float
            };
            media_type.SetGUID(&MF_MT_SUBTYPE, subtype)?;

            media_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, wf.Format.nSamplesPerSec)?;
            media_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, wf.Format.nChannels as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, wf.Format.wBitsPerSample as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, wf.Format.nBlockAlign as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, wf.Format.nAvgBytesPerSec)?;

            // Optional but recommended: Channel Mask
            if wf.dwChannelMask != 0 {
                media_type.SetUINT32(&MF_MT_AUDIO_CHANNEL_MASK, wf.dwChannelMask)?;
            } else {
                 // Channel mask might be 0 for standard mono/stereo in basic WAVEFORMATEX
                 // MF might infer it, but good to log.
                 println!("Warning: Channel mask is zero in WAVEFORMATEXTENSIBLE. MF might infer.");
            }

            // Optional: Set other attributes if needed

            Ok(media_type)
        }
    }

    // Accessor for the created media type
    pub fn media_type(&self) -> &IMFMediaType {
        &self.media_type
    }

    // Creates a clone of the media type (useful if the consumer needs ownership)
    pub fn clone_media_type(&self) -> Result<IMFMediaType> {
        unsafe {
            let new_type = MFCreateMediaType()?;
            self.media_type.CopyAllItems(&new_type)?;
            Ok(new_type)
        }
    }

    // Starts the audio capture thread
    pub fn start(&mut self, sample_sender: SyncSender<(u32, AudioDataPacket)>) -> Result<()> {
        if self.thread_handle.is_some() {
            return Err(windows::core::Error::new(E_FAIL, "Audio capture already started".into()));
        }

        // Prepare parameters for the audio thread
        let params = AudioThreadParams {
            device_id: self.device_id.clone(), // Clone HSTRING
            source: self.source,
            wave_format: self.wave_format, // WAVEFORMATEXTENSIBLE is Copy
            sample_sender, // Move sender ownership to params
            audio_stream_index: self.audio_stream_index,
            qpc_frequency: self.qpc_frequency,
            stop_signal: self.stop_signal.clone(), // Clone Arc
        };

        println!("Starting audio capture thread...");
        self.thread_handle = Some(thread::spawn(move || -> Result<()> {
            // --- Audio Thread Entry Point ---
            let thread_result = unsafe { audio_capture_thread(params) };
            // --- Audio Thread Exit Point ---

            match thread_result {
                Ok(()) => {
                    println!("Audio capture thread finished successfully.");
                    Ok(())
                }
                Err(e) => {
                    eprintln!("Audio capture thread failed: HRESULT={:?}", e.code()); // Log HRESULT
                    eprintln!("  Error message: {}", e.message());
                    Err(e) // Propagate the error
                }
            }
        }));

        Ok(())
    }

    // Stops the audio capture thread
    pub fn stop(&mut self) -> Result<()> {
        if let Some(handle) = self.thread_handle.take() {
            println!("Stopping audio capture thread...");
            self.stop_signal.store(true, Ordering::SeqCst);

            // The thread checks the stop_signal periodically via the WaitForSingleObject timeout.
            // No need to signal the event handle directly from here.

            println!("Waiting for audio thread to join...");
            match handle.join() {
                Ok(thread_result) => {
                    println!("Audio thread joined.");
                    thread_result? // Propagate Result from thread (Ok(()) or Err)
                }
                Err(e) => {
                    // This means the thread panicked
                    eprintln!("Audio thread panicked: {:?}", e);
                    return Err(windows::core::Error::new(E_FAIL, "Audio thread panic".into()));
                }
            }
            println!("Audio stop complete.");
        } else {
            println!("Audio stop called but thread was not running.");
        }
        // Reset signal for potential restart (optional, depends on use case)
        // self.stop_signal.store(false, Ordering::SeqCst);
        Ok(())
    }

    pub fn set_stream_index(&mut self, index: u32) {
        println!("AudioCapture: Setting stream index to {}", index);
        self.audio_stream_index = index;
        // This is safe because we call it *before* start() creates the thread.
   }
} // End impl AudioCapture

// --- Audio Capture Thread Function ---
// This runs entirely within the spawned thread.
unsafe fn audio_capture_thread(params: AudioThreadParams) -> Result<()> {
    // 1. Initialize COM for this thread
    CoInitializeEx(None, COINIT_MULTITHREADED)?; // Use MTA

    // RAII guard for CoUninitialize
    struct ComGuard;
    impl Drop for ComGuard {
        fn drop(&mut self) {
            println!("Uninitializing COM on audio thread.");
            unsafe { CoUninitialize(); }
        }
    }
    let _com_guard = ComGuard; // Ensures CoUninitialize is called on exit

    // Use a block to scope COM objects and handle errors, ensuring CoUninitialize runs
    let result: Result<()> = (|| {
        // 2. Get Device Enumerator and Specific Device
        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let device: IMMDevice = enumerator.GetDevice(&params.device_id)?;

        // 3. Activate and Initialize IAudioClient on *this* thread
        let audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;

        let stream_flags = if params.source == AudioSource::DefaultLoopback {
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK
        } else {
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK
        };
        // Request a buffer duration (e.g., 100ms). WASAPI uses 100ns units.
        let requested_duration_hns: i64 = 100 * 10_000;

        audio_client.Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            stream_flags,
            requested_duration_hns, // hnsBufferDuration
            0, // hnsPeriodicity (0 lets WASAPI choose based on duration)
            &params.wave_format as *const _ as *const WAVEFORMATEX,
            None, // AudioSessionGuid (optional)
        )?;
        println!("AudioClient initialized.");


        // 4. Create event handle for event-driven capture
        let capture_event = CreateEventW(None, false, false, None)?; // Auto-reset event
        if capture_event.is_invalid() {
             return Err(windows::core::Error::from_win32());
        }
        // RAII guard for the event handle
        struct HandleGuard(HANDLE);
        impl Drop for HandleGuard {
             fn drop(&mut self) { if !self.0.is_invalid() { unsafe { CloseHandle(self.0); } } }
        }
        let _event_guard = HandleGuard(capture_event); // Ensure CloseHandle is called

        audio_client.SetEventHandle(capture_event)?;


        // 5. Get Capture Client service
        let capture_client: IAudioCaptureClient = audio_client.GetService()?;


        // 6. Start Capture
        audio_client.Start()?;
        println!("Audio capture thread started successfully (WASAPI).");


        // --- Capture Loop ---
        let block_align = params.wave_format.Format.nBlockAlign as u32;
        if block_align == 0 {
             eprintln!("Error: BlockAlign is zero!");
             return Err(windows::core::Error::new(E_FAIL, "Invalid audio format: BlockAlign is zero".into()));
        }
        let sample_rate = params.wave_format.Format.nSamplesPerSec;
        if sample_rate == 0 {
             eprintln!("Error: SampleRate is zero!");
             return Err(windows::core::Error::new(E_FAIL, "Invalid audio format: SampleRate is zero".into()));
        }


        while !params.stop_signal.load(Ordering::SeqCst) {
            // Wait for data or timeout
            let wait_result = WaitForSingleObject(capture_event, 100); // 100ms timeout

            if params.stop_signal.load(Ordering::SeqCst) {
                println!("Stop signal received during wait.");
                break; // Exit loop if stop requested during wait
            }

            match wait_result {
                WAIT_OBJECT_0 => {
                    // Data is available (or buffer released) - process all available packets
                    loop {
                        if params.stop_signal.load(Ordering::SeqCst) {
                            println!("Stop signal received before GetBuffer.");
                            break; // Exit inner loop if stopped
                        }

                        let mut buffer_ptr = std::ptr::null_mut();
                        let mut frames_available = 0u32;
                        let mut flags = 0u32;
                        let mut device_position = 0u64; // Can be useful for lip sync?
                        let mut qpc_position = 0u64;    // Raw QPC value for timestamp

                        // GetBuffer returns Result<(), Error> directly
                        let result = capture_client.GetBuffer(
                            &mut buffer_ptr, &mut frames_available, &mut flags,
                            Some(&mut device_position), Some(&mut qpc_position),
                        );

                        // --- Result Handling ---
                        match result {
                            Ok(()) => {
                                // Buffer obtained successfully, continue processing
                            },
                            Err(err) if err.code() == AUDCLNT_S_BUFFER_EMPTY => {
                                // Buffer is empty, no more data in this cycle. Wait for next event.
                                break; // Exit inner loop
                            },
                            Err(err) => {
                                // An actual error occurred. Log and return.
                                eprintln!("IAudioCaptureClient::GetBuffer failed: {:?}", err);
                                return Err(err);
                            }
                        }
                        // If here, hr is S_OK, data should be available.
                        // --- End HRESULT Handling ---


                        if params.stop_signal.load(Ordering::SeqCst) {
                            println!("Stop signal received after GetBuffer.");
                            // Need to release the buffer if we successfully got one
                            if frames_available > 0 { capture_client.ReleaseBuffer(frames_available)?; }
                            break; // Exit inner loop
                        }

                        if frames_available == 0 {
                            // This shouldn't happen if hr was S_OK, but check anyway.
                            println!("Warning: GetBuffer succeeded (S_OK) but frames_available is 0.");
                            // Must still release the (empty) buffer.
                            capture_client.ReleaseBuffer(frames_available)?;
                            continue; // Try getting buffer status again immediately
                        }


                        // --- Process Buffer Data ---
                        let bytes_available = frames_available * block_align;
                        let timestamp_100ns = convert_qpc_to_mf_timespan(qpc_position as i64, params.qpc_frequency)?;
                        let duration_100ns = (frames_available as i64 * 10_000_000) / (sample_rate as i64);

                        let mut audio_packet_data: Option<Vec<u8>> = None;

                        if flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                            // WASAPI indicates silence, send None data
                            // println!("Audio Packet: SILENCE ({} frames)", frames_available); // Can be verbose
                            audio_packet_data = None;
                        } else if !buffer_ptr.is_null() {
                            // Valid data buffer, copy it into a Vec
                            let mut data_vec = vec![0u8; bytes_available as usize];
                            std::ptr::copy_nonoverlapping(
                                buffer_ptr as *const u8,
                                data_vec.as_mut_ptr(),
                                bytes_available as usize,
                            );
                            audio_packet_data = Some(data_vec);
                        } else {
                            // Non-silent flag but null buffer? This is unexpected.
                            eprintln!("Warning: Got non-silent flag from GetBuffer but null buffer pointer.");
                            audio_packet_data = None; // Treat as silence to avoid panic
                        }

                        let packet = AudioDataPacket {
                            data: audio_packet_data,
                            timestamp: timestamp_100ns,
                            duration: duration_100ns,
                            flags: 0, // Set flags like discontinuity if needed based on WASAPI flags/state
                        };
                        // --- End Process Buffer Data ---


                        // --- Send Packet via Channel ---
                        // Use try_send to avoid blocking this audio thread if the receiver is slow/full.
                        if params.sample_sender.try_send((params.audio_stream_index, packet)).is_err() {
                            eprintln!("Audio sample sender disconnected or buffer full. Stopping thread.");
                            // Release the WASAPI buffer *before* breaking out
                            capture_client.ReleaseBuffer(frames_available)?;
                            params.stop_signal.store(true, Ordering::SeqCst); // Signal self to stop fully
                            break; // Exit inner loop
                        }
                        // --- End Send Packet ---


                        // --- Release WASAPI Buffer ---
                        // Must be called after processing the data and sending the packet.
                        capture_client.ReleaseBuffer(frames_available)?;
                        // --- End Release WASAPI Buffer ---

                    } // End inner loop (processing available packets)
                }
                WAIT_TIMEOUT => {
                    // Timeout occurred, just loop again to check stop_signal and wait again.
                    // println!("Audio thread wait timeout."); // Can be verbose
                    continue;
                }
                _ => {
                    // Some other wait error occurred
                    let error = windows::core::Error::from_win32();
                    eprintln!("WaitForSingleObject error in audio thread: {:?}", error);
                    return Err(error); // Exit outer loop with error
                }
            } // End match wait_result
        } // End while !stop_signal (outer loop)


        // 7. Stop capturing & Cleanup
        println!("Audio capture loop finished. Stopping AudioClient...");
        let stop_hr = audio_client.Stop(); // Ignore error? Device might be removed.
        if stop_hr.is_err() {
            eprintln!("AudioClient::Stop failed: HRESULT={:?}", stop_hr);
        }
        println!("AudioClient stopped.");

        // Explicitly drop COM objects before CoUninitialize (optional, scope does it)
        drop(capture_client);
        drop(audio_client);
        drop(device);
        drop(enumerator);

        Ok(()) // Thread finished normally
    }) (); // End of COM object lifetime block / error handling closure

    // COM is automatically uninitialized by _com_guard dropping here
    println!("Audio capture thread function exiting.");
    result // Return the final result (Ok or Err)
}


// --- Helper function for the RECEIVING end ---
// Reconstructs an IMFSample from the received AudioDataPacket.
// This should be called in the thread that consumes data from the channel.
pub fn create_imf_sample_from_packet(packet: AudioDataPacket) -> Result<IMFSample> {
    unsafe {
        let sample = MFCreateSample()?;
        sample.SetSampleTime(packet.timestamp)?;
        sample.SetSampleDuration(packet.duration)?;

        // TODO: Handle packet.flags if you implement discontinuity detection etc.
        // Example:
        // if (packet.flags & YOUR_DISCONTINUITY_FLAG) != 0 {
        //     sample.SetUINT32(&MFSampleExtension_Discontinuity, 1)?;
        // }

        if let Some(data) = packet.data {
            // Packet contains actual audio data
            if !data.is_empty() {
                let media_buffer = MFCreateMemoryBuffer(data.len() as u32)?;

                let mut buffer_ptr = std::ptr::null_mut();
                let mut max_len = 0;
                // Lock the buffer to get a pointer to its memory
                media_buffer.Lock(&mut buffer_ptr, Some(&mut max_len), None)?; // Don't need current_len ptr

                // RAII guard for unlocking the buffer
                struct UnlockGuard<'a>(&'a IMFMediaBuffer);
                impl<'a> Drop for UnlockGuard<'a> {
                    fn drop(&mut self) { unsafe { let _ = self.0.Unlock(); } } // Ignore unlock error?
                }
                let _unlock_guard = UnlockGuard(&media_buffer);

                if (data.len() as u32) <= max_len {
                    // Copy data from the Vec into the media buffer
                    std::ptr::copy_nonoverlapping(data.as_ptr(), buffer_ptr, data.len());
                    media_buffer.SetCurrentLength(data.len() as u32)?;
                } else {
                    // This indicates an internal logic error
                    eprintln!("Error: Created media buffer is smaller than packet data size ({} > {}).", data.len(), max_len);
                    // Unlock is handled by guard dropping
                    return Err(windows::core::Error::new(E_FAIL,"Media buffer size mismatch".into()));
                }

                // Unlock is handled by guard dropping
                drop(_unlock_guard);

                // Add the buffer (now filled with data) to the sample
                sample.AddBuffer(&media_buffer)?;
            } else {
                // Data is Some, but the Vec is empty. Treat as silence?
                println!("Warning: Reconstructing sample from packet with empty data vector.");
                // No buffer added, sample has only timestamp/duration.
            }
        } else {
            // Data is None, indicating silence from WASAPI.
            // No buffer is added. Downstream (SinkWriter) should handle this.
            // println!("Reconstructing silent audio sample (Timestamp: {}, Duration: {}).", packet.timestamp, packet.duration);
        }

        Ok(sample)
    }
}


// --- Drop Implementation ---
impl Drop for AudioCapture {
    fn drop(&mut self) {
        // Ensure the thread is stopped and joined when AudioCapture goes out of scope.
        if self.thread_handle.is_some() {
            println!("AudioCapture drop: Stopping running audio thread...");
            if let Err(e) = self.stop() {
                eprintln!("Error stopping audio thread during drop: {:?}", e);
            }
        }
        println!("AudioCapture dropped.");
        // COM objects are managed within the thread or constructor, no direct cleanup needed here.
        // CoUninitialize for the constructor thread is not handled here;
        // assume the application manages overall COM lifetime.
    }
}
pub fn convert_qpc_to_mf_timespan(qpc_time: i64, qpc_frequency: i64) -> Result<i64> {
    if qpc_time == 0 {
        // Handle zero QPC time (e.g., return 0 MF time or error if needed)
        return Ok(0);
    }
    if qpc_frequency <= 0 {
        // This should not happen if frequency was queried correctly
        return Err(windows::core::Error::new(
            E_FAIL,
            "Invalid QPC frequency provided for timestamp conversion".into(),
        ));
    }
    // Calculate duration in 100ns units using i128 for intermediate calculation
    // to avoid overflow with large qpc_time values.
    let mf_time = (qpc_time as i128 * 10_000_000) / (qpc_frequency as i128);
    Ok(mf_time as i64)
}