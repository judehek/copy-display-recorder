// src/audio.rs
use std::{
    ffi::c_void,
    ptr::NonNull,
    sync::{
        atomic::{AtomicBool, Ordering},
        mpsc::{self, SyncSender},
        Arc,
    },
    thread::{self, JoinHandle},
    time::Duration,
};

use clap::ValueEnum;
use windows::{
    core::{Interface, Result, HSTRING},
    Win32::{
        Foundation::{CloseHandle, E_FAIL, E_NOTIMPL, HANDLE, WAIT_OBJECT_0, WAIT_TIMEOUT},
        Media::{
            Audio::{
                eCapture, eConsole, eRender, Endpoints::IAudioEndpointVolume, IAudioCaptureClient, IAudioClient, IMMDevice, IMMDeviceEnumerator, MMDeviceEnumerator, AUDCLNT_BUFFERFLAGS_SILENT, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK, AUDCLNT_S_BUFFER_EMPTY, WAVEFORMATEX, WAVEFORMATEXTENSIBLE, WAVE_FORMAT_PCM
            },
            KernelStreaming::KSDATAFORMAT_SUBTYPE_PCM,
            MediaFoundation::{
                IMFMediaBuffer, IMFMediaType, IMFSample, MFAudioFormat_Float, MFAudioFormat_PCM,
                MFCreateMediaType, MFCreateMemoryBuffer, MFCreateSample, MFMediaType_Audio,
                MF_E_INVALIDREQUEST,
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
            Threading::{CreateEventW, SetEvent, WaitForSingleObject},
        },
    },
};

#[derive(Debug)]
pub struct AudioDataPacket {
    pub data: Option<Vec<u8>>,
    pub timestamp: i64,
    pub duration: i64,
    pub flags: u32,
}

struct AudioThreadParams {
    device_id: HSTRING,
    source: AudioSource,
    wave_format: WAVEFORMATEXTENSIBLE,
    sample_sender: SyncSender<(u32, AudioDataPacket)>,
    audio_stream_index: u32,
    stop_signal: Arc<AtomicBool>,
}

pub struct AudioCapture {
    device_id: HSTRING,
    wave_format: WAVEFORMATEXTENSIBLE,
    source: AudioSource,
    audio_stream_index: u32,
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
            CoInitializeEx(None, COINIT_MULTITHREADED)?;

            let enumerator: IMMDeviceEnumerator =
                CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;

            let data_flow = match source {
                AudioSource::DefaultLoopback => eRender,
                AudioSource::DefaultMicrophone => eCapture,
                AudioSource::None => unreachable!(),
            };
            let device: IMMDevice = enumerator.GetDefaultAudioEndpoint(data_flow, eConsole)?;
            let device_id_pwstr = device.GetId()?;
            let device_id: HSTRING = device_id_pwstr.to_string()?.into();

            let temp_audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
            let mix_format_ptr = temp_audio_client.GetMixFormat()?;

            let mix_wave_format = if (*mix_format_ptr).cbSize >= (std::mem::size_of::<WAVEFORMATEXTENSIBLE>() - std::mem::size_of::<WAVEFORMATEX>()) as u16 {
                *(mix_format_ptr as *const WAVEFORMATEXTENSIBLE)
            } else {
                 let base_format = *(mix_format_ptr as *const WAVEFORMATEX);
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
            let target_channel_mask = mix_wave_format.dwChannelMask;

            CoTaskMemFree(Some(mix_format_ptr as *const c_void));
            drop(temp_audio_client);

            let target_bits_per_sample: u16 = 16;
            let target_block_align = target_channels * (target_bits_per_sample / 8);
            let target_avg_bytes_per_sec = target_sample_rate * (target_block_align as u32);

            let final_channel_mask = if target_channel_mask != 0 {
                target_channel_mask
            } else {
                match target_channels {
                    1 => 0x4,
                    2 => 0x3,
                    _ => 0
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
                SubFormat: KSDATAFORMAT_SUBTYPE_PCM,
            };

            let media_type = Self::create_mf_media_type(&desired_wave_format)?;

            let stop_signal = Arc::new(AtomicBool::new(false));

            Ok(Self {
                device_id,
                wave_format: desired_wave_format,
                source,
                audio_stream_index,
                media_type,
                thread_handle: None,
                stop_signal,
            })
        }
    }

    fn create_mf_media_type(wf: &WAVEFORMATEXTENSIBLE) -> Result<IMFMediaType> {
        unsafe {
            let media_type = MFCreateMediaType()?;

            media_type.SetGUID(&MF_MT_MAJOR_TYPE, &MFMediaType_Audio)?;

            let sub_format = wf.SubFormat;

            let subtype = if sub_format == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT {
                &MFAudioFormat_Float
            } else if sub_format == KSDATAFORMAT_SUBTYPE_PCM {
                let bits_per_sample = wf.Format.wBitsPerSample;
                if bits_per_sample == 16 {
                    &MFAudioFormat_PCM
                } else {
                    return Err(windows::core::Error::new(E_NOTIMPL, "Unsupported PCM bit depth for MF Media Type".into()));
                }
            } else {
                &MFAudioFormat_Float
            };
            media_type.SetGUID(&MF_MT_SUBTYPE, subtype)?;

            media_type.SetUINT32(&MF_MT_AUDIO_SAMPLES_PER_SECOND, wf.Format.nSamplesPerSec)?;
            media_type.SetUINT32(&MF_MT_AUDIO_NUM_CHANNELS, wf.Format.nChannels as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_BITS_PER_SAMPLE, wf.Format.wBitsPerSample as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_BLOCK_ALIGNMENT, wf.Format.nBlockAlign as u32)?;
            media_type.SetUINT32(&MF_MT_AUDIO_AVG_BYTES_PER_SECOND, wf.Format.nAvgBytesPerSec)?;

            if wf.dwChannelMask != 0 {
                media_type.SetUINT32(&MF_MT_AUDIO_CHANNEL_MASK, wf.dwChannelMask)?;
            }

            Ok(media_type)
        }
    }

    pub fn media_type(&self) -> &IMFMediaType {
        &self.media_type
    }

    pub fn clone_media_type(&self) -> Result<IMFMediaType> {
        unsafe {
            let new_type = MFCreateMediaType()?;
            self.media_type.CopyAllItems(&new_type)?;
            Ok(new_type)
        }
    }

    pub fn start(&mut self, sample_sender: SyncSender<(u32, AudioDataPacket)>) -> Result<()> {
        if self.thread_handle.is_some() {
            return Err(windows::core::Error::new(E_FAIL, "Audio capture already started".into()));
        }

        let params = AudioThreadParams {
            device_id: self.device_id.clone(),
            source: self.source,
            wave_format: self.wave_format,
            sample_sender,
            audio_stream_index: self.audio_stream_index,
            stop_signal: self.stop_signal.clone(),
        };

        self.thread_handle = Some(thread::spawn(move || -> Result<()> {
            let thread_result = unsafe { audio_capture_thread(params) };

            match thread_result {
                Ok(()) => Ok(()),
                Err(e) => {
                    eprintln!("Audio capture thread failed: HRESULT={:?}", e.code());
                    eprintln!("  Error message: {}", e.message());
                    Err(e)
                }
            }
        }));

        Ok(())
    }

    pub fn stop(&mut self) -> Result<()> {
        if let Some(handle) = self.thread_handle.take() {
            self.stop_signal.store(true, Ordering::SeqCst);

            match handle.join() {
                Ok(thread_result) => thread_result?,
                Err(e) => {
                    eprintln!("Audio thread panicked: {:?}", e);
                    return Err(windows::core::Error::new(E_FAIL, "Audio thread panic".into()));
                }
            }
        }
        Ok(())
    }

    pub fn set_stream_index(&mut self, index: u32) {
        self.audio_stream_index = index;
   }
}

unsafe fn audio_capture_thread(params: AudioThreadParams) -> Result<()> {
    CoInitializeEx(None, COINIT_MULTITHREADED)?;

    struct ComGuard;
    impl Drop for ComGuard {
        fn drop(&mut self) {
            unsafe { CoUninitialize(); }
        }
    }
    let _com_guard = ComGuard;

    let result: Result<()> = (|| {
        let enumerator: IMMDeviceEnumerator =
            CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
        let device: IMMDevice = enumerator.GetDevice(&params.device_id)?;

        let audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;

        let stream_flags = if params.source == AudioSource::DefaultLoopback {
            AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK
        } else {
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK
        };
        let requested_duration_hns: i64 = 100 * 10_000;

        audio_client.Initialize(
            AUDCLNT_SHAREMODE_SHARED,
            stream_flags,
            requested_duration_hns,
            0,
            &params.wave_format as *const _ as *const WAVEFORMATEX,
            None,
        )?;

        let capture_event = CreateEventW(None, false, false, None)?;
        if capture_event.is_invalid() {
             return Err(windows::core::Error::from_win32());
        }
        struct HandleGuard(HANDLE);
        impl Drop for HandleGuard {
             fn drop(&mut self) { if !self.0.is_invalid() { unsafe { CloseHandle(self.0); } } }
        }
        let _event_guard = HandleGuard(capture_event);

        audio_client.SetEventHandle(capture_event)?;

        let capture_client: IAudioCaptureClient = audio_client.GetService()?;

        audio_client.Start()?;

        let block_align = params.wave_format.Format.nBlockAlign as u32;
        if block_align == 0 {
             return Err(windows::core::Error::new(E_FAIL, "Invalid audio format: BlockAlign is zero".into()));
        }
        let sample_rate = params.wave_format.Format.nSamplesPerSec;
        if sample_rate == 0 {
             return Err(windows::core::Error::new(E_FAIL, "Invalid audio format: SampleRate is zero".into()));
        }

        let mut get_buffer_calls = 0u64;
        let mut successful_gets = 0u64;
        let mut silent_packets = 0u64;
        let mut data_packets = 0u64;
        let mut zero_frame_sok_calls = 0u64;

        while !params.stop_signal.load(Ordering::SeqCst) {
            let wait_result = WaitForSingleObject(capture_event, 100);

            if params.stop_signal.load(Ordering::SeqCst) {
                break;
            }

            match wait_result {
                WAIT_OBJECT_0 => {
                    loop {
                        if params.stop_signal.load(Ordering::SeqCst) {
                            break;
                        }

                        let mut buffer_ptr = std::ptr::null_mut();
                        let mut frames_available = 0u32;
                        let mut flags = 0u32;
                        let mut device_position = 0u64;
                        let mut qpc_position = 0u64;

                        get_buffer_calls += 1;

                        let result = capture_client.GetBuffer(
                            &mut buffer_ptr, &mut frames_available, &mut flags,
                            Some(&mut device_position), Some(&mut qpc_position),
                        );

                        match result {
                            Ok(()) => {
                                successful_gets += 1;
                            },
                            Err(err) if err.code() == AUDCLNT_S_BUFFER_EMPTY => {
                                break;
                            },
                            Err(err) => {
                                eprintln!("IAudioCaptureClient::GetBuffer failed: {:?}", err);
                                return Err(err);
                            }
                        }

                        if params.stop_signal.load(Ordering::SeqCst) {
                            if frames_available > 0 { capture_client.ReleaseBuffer(frames_available)?; }
                            break;
                        }

                        if frames_available == 0 {
                            zero_frame_sok_calls += 1;
                            capture_client.ReleaseBuffer(frames_available)?;
                            continue;
                        }

                        let bytes_available = frames_available * block_align;
                        let timestamp_100ns = convert_qpc_to_mf_timespan(qpc_position as i64)?;
                        let duration_100ns = (frames_available as i64 * 10_000_000) / (sample_rate as i64);

                        let mut audio_packet_data: Option<Vec<u8>> = None;

                        if flags & (AUDCLNT_BUFFERFLAGS_SILENT.0 as u32) != 0 {
                            silent_packets += 1;
                            audio_packet_data = None;
                        } else if !buffer_ptr.is_null() {
                            data_packets += 1;

                            let mut data_vec = vec![0u8; bytes_available as usize];
                            std::ptr::copy_nonoverlapping(
                                buffer_ptr as *const u8,
                                data_vec.as_mut_ptr(),
                                bytes_available as usize,
                            );
                            audio_packet_data = Some(data_vec);
                        } else {
                            silent_packets += 1;
                            audio_packet_data = None;
                        }

                        let packet = AudioDataPacket {
                            data: audio_packet_data,
                            timestamp: timestamp_100ns,
                            duration: duration_100ns,
                            flags: flags,
                        };

                        if params.sample_sender.try_send((params.audio_stream_index, packet)).is_err() {
                            eprintln!("Audio sample sender disconnected or buffer full. Stopping thread.");
                            capture_client.ReleaseBuffer(frames_available)?;
                            params.stop_signal.store(true, Ordering::SeqCst);
                            break;
                        }

                        capture_client.ReleaseBuffer(frames_available)?;

                    }
                }
                WAIT_TIMEOUT => {
                    continue;
                }
                _ => {
                    let error = windows::core::Error::from_win32();
                    eprintln!("WaitForSingleObject error in audio thread: {:?}", error);
                    return Err(error);
                }
            }
        }

        let stop_hr = audio_client.Stop();
        if stop_hr.is_err() {
            eprintln!("AudioClient::Stop failed: HRESULT={:?}", stop_hr);
        }

        Ok(())
    }) ();

    result
}
// --- *** CHANGE HERE *** ---
// Modify signature and remove internal relative timestamp calculation
pub fn create_imf_sample_from_packet(packet: AudioDataPacket, relative_timestamp_100ns: i64) -> Result<IMFSample> {
    // --- *** END CHANGE *** ---
        unsafe {
            let sample = MFCreateSample()?;
    
            // --- *** CHANGE HERE *** ---
            // Use the timestamp passed directly from SampleWriter
            sample.SetSampleTime(relative_timestamp_100ns)?;
            // --- *** END CHANGE *** ---
    
            sample.SetSampleDuration(packet.duration)?; // Duration is still calculated correctly in thread
    
            // --- Buffer creation and data copy remains the same ---
            if let Some(data) = packet.data {
                if !data.is_empty() {
                    let media_buffer = MFCreateMemoryBuffer(data.len() as u32)?;
    
                    let mut buffer_ptr = std::ptr::null_mut();
                    let mut max_len = 0;
                    media_buffer.Lock(&mut buffer_ptr, Some(&mut max_len), None)?;
    
                    // Use a guard to ensure Unlock is called
                    struct UnlockGuard<'a>(&'a IMFMediaBuffer);
                    impl<'a> Drop for UnlockGuard<'a> { fn drop(&mut self) { unsafe { let _ = self.0.Unlock(); } } }
                    let _unlock_guard = UnlockGuard(&media_buffer);
    
                    if (data.len() as u32) <= max_len {
                        std::ptr::copy_nonoverlapping(data.as_ptr(), buffer_ptr, data.len());
                        media_buffer.SetCurrentLength(data.len() as u32)?;
                    } else {
                        eprintln!("Error: Created media buffer is smaller than packet data size ({} > {}).", data.len(), max_len);
                        // Unlock is handled by guard
                        return Err(windows::core::Error::new(E_FAIL,"Media buffer size mismatch".into()));
                    }
                    // Unlock guard drops here
    
                    sample.AddBuffer(&media_buffer)?;
                }
                // If data is Some but empty, or if data is None, no buffer is added (representing silence)
            }
    
            Ok(sample)
        }
    }
    
    // --- convert_qpc_to_mf_timespan remains the same (used by SampleWriter now) ---
    pub fn convert_qpc_to_mf_timespan(qpc_time: i64) -> Result<i64> {
        println!("Converting QPC time: {}", qpc_time);
        
        if qpc_time == 0 {
            println!("QPC time is 0, returning 0");
            return Ok(0);
        }
    
        let mut frequency = 0;
        unsafe {
            // ERROR: The condition is checking if it IS an error but returns an error if true
            // The correct check is to verify the function succeeded, not failed
            if QueryPerformanceFrequency(&mut frequency).is_err() {
                println!("QueryPerformanceFrequency failed");
                return Err(windows::core::Error::from_win32());
            }
        }
        println!("QueryPerformanceFrequency result: {}", frequency);
    
        if frequency <= 0 {
            println!("QueryPerformanceFrequency returned non-positive frequency: {}", frequency);
            return Err(windows::core::Error::new(
                E_FAIL,
                "QueryPerformanceFrequency returned non-positive frequency".into(),
            ));
        }
    
        // Perform calculation using i128 to avoid overflow
        let mf_time = (qpc_time as i128 * 10_000_000) / (frequency as i128);
        println!("Calculated MF time: {}", mf_time);
        
        Ok(mf_time as i64)
    }