use std::sync::mpsc::{channel, Receiver, Sender};
use std::thread;

use ringbuf::traits::{Consumer, Observer};
use ringbuf::wrap::caching::Caching;
use windows::Foundation::TimeSpan;
use windows::Win32::Foundation::HANDLE;
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;
use windows::Win32::Media::Multimedia::KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
use windows::Win32::System::Performance::QueryPerformanceFrequency;
use windows::{
    core::*,
    Win32::{
        Foundation::{ E_FAIL, WAIT_OBJECT_0, WAIT_TIMEOUT},
        Media::Audio::{
            eConsole, eRender, IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator, MMDeviceEnumerator, 
            AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK,
            WAVEFORMATEX, WAVEFORMATEXTENSIBLE,
        },
        System::{
            Com::{CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED},
            Threading::{CreateEventW, WaitForSingleObject},
        },
    },
};

use ringbuf::{HeapRb, producer::Producer, traits::Split};
use std::sync::Arc;
use std::sync::atomic::{AtomicU32, AtomicU16, AtomicBool, Ordering};

use super::capture_audio::AudioSample;
use super::encoding_session::AudioSource;

// Constants used within this module
const REFTIMES_PER_SEC: i64 = 10000000; // 100ns units per second
const REFTIMES_PER_MILLISEC: i64 = 10000; // 100ns units per millisecond

// Hard-coded audio format constants
const HARD_CODED_SAMPLE_RATE: u32 = 48000;
const HARD_CODED_CHANNELS: u16 = 2;
const HARD_CODED_BITS_PER_SAMPLE: u16 = 32;  // 32-bit float

pub struct MicrophoneCaptureSession {
    sender: Sender<bool>, // To signal start/stop
    running: bool,
}

impl MicrophoneCaptureSession {
    fn new(sender: Sender<bool>) -> Self {
        Self {
            sender,
            running: false,
        }
    }

    pub fn StartCapture(&mut self) -> Result<()> {
        if !self.running {
            self.sender.send(true).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = true;
        }
        Ok(())
    }

    pub fn StopCapture(&mut self) -> Result<()> {
        if self.running {
            self.sender.send(false).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = false;
        }
        Ok(())
    }
}

impl Clone for MicrophoneCaptureSession {
    fn clone(&self) -> Self {
        Self {
            sender: self.sender.clone(),
            running: self.running,
        }
    }
}

pub struct CaptureMicrophoneGenerator {
    consumer: Caching<Arc<ringbuf::SharedRb<ringbuf::storage::Heap<AudioSample>>>, false, true>,
    session: MicrophoneCaptureSession,
    sample_rate: Arc<AtomicU32>,
    channels: Arc<AtomicU16>,
    bits_per_sample: Arc<AtomicU16>,
    initialized: Arc<AtomicBool>,
    start_qpc: i64,
    qpf_frequency: i64,
}

// Helper function to create a WAVEFORMATEXTENSIBLE struct with our hard-coded format
unsafe fn create_hardcoded_wave_format() -> WAVEFORMATEXTENSIBLE {
    let bytes_per_sample = (HARD_CODED_BITS_PER_SAMPLE / 8) as u16;
    let block_align = HARD_CODED_CHANNELS * bytes_per_sample;
    let avg_bytes_per_sec = HARD_CODED_SAMPLE_RATE * block_align as u32;
    
    let mut format = WAVEFORMATEXTENSIBLE::default();
    format.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE as u16;
    format.Format.nChannels = HARD_CODED_CHANNELS;
    format.Format.nSamplesPerSec = HARD_CODED_SAMPLE_RATE;
    format.Format.nAvgBytesPerSec = avg_bytes_per_sec;
    format.Format.nBlockAlign = block_align;
    format.Format.wBitsPerSample = HARD_CODED_BITS_PER_SAMPLE;
    format.Format.cbSize = 22; // Size of the WAVEFORMATEXTENSIBLE struct minus the size of WAVEFORMATEX
    
    format.Samples.wValidBitsPerSample = HARD_CODED_BITS_PER_SAMPLE;
    format.dwChannelMask = 3; // SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT
    
    // Set the SubFormat to IEEE_FLOAT
    format.SubFormat = KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
    
    format
}

unsafe fn initialize_audio_capture(audio_source: &AudioSource) -> Result<(IAudioClient, IAudioCaptureClient, HANDLE, u32, u16, u16)> {
    // Create device enumerator
    let device_enumerator: IMMDeviceEnumerator = CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
    
    // Get default audio endpoint
    let device = device_enumerator.GetDefaultAudioEndpoint(eRender, eConsole)?;
    
    // Activate audio client
    let client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
    
    // Create event handle
    let handle = CreateEventW(None, false, false, None)?;
    if handle.is_invalid() {
        return Err(windows::core::Error::from(E_FAIL));
    }
    
    // Set up client initialization flags based on audio source
    let mut stream_flags = AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
    
    match audio_source {
        AudioSource::Desktop => {
            stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
        },
        AudioSource::ActiveWindow => {
            // TODO: For future implementation - for now, use loopback
            stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
            println!("Warning: Active window audio capture not fully implemented yet, falling back to desktop audio");
        }
    }
    
    // Create our hard-coded WAVEFORMATEXTENSIBLE
    let wave_format_ex = create_hardcoded_wave_format();
    let wave_format_ptr = &wave_format_ex.Format as *const WAVEFORMATEX;
    
    // Initialize audio client with our hard-coded format
    let buffer_duration_ms = 10;
    let buffer_duration_100ns = buffer_duration_ms * REFTIMES_PER_MILLISEC;
    
    client.Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        stream_flags,
        buffer_duration_100ns,
        0,
        wave_format_ptr,
        None,
    )?;
    
    // Set event handle
    client.SetEventHandle(handle)?;
    
    // Get capture client
    let capture_client = client.GetService::<IAudioCaptureClient>()?;
    
    // Start audio client
    client.Start()?;
    
    println!("Audio capture initialized and started with format: {}Hz, {} channels, {}-bit",
             HARD_CODED_SAMPLE_RATE, HARD_CODED_CHANNELS, HARD_CODED_BITS_PER_SAMPLE);
    
    Ok((client, capture_client, handle, HARD_CODED_SAMPLE_RATE, HARD_CODED_CHANNELS, HARD_CODED_BITS_PER_SAMPLE))
}

impl CaptureMicrophoneGenerator {
    pub fn new(audio_source: AudioSource, start_qpc: i64) -> Result<Self> {
        // Create shared atomic variables with hard-coded values
        let sample_rate = Arc::new(AtomicU32::new(HARD_CODED_SAMPLE_RATE));
        let channels = Arc::new(AtomicU16::new(HARD_CODED_CHANNELS));
        let bits_per_sample = Arc::new(AtomicU16::new(HARD_CODED_BITS_PER_SAMPLE));
        let initialized = Arc::new(AtomicBool::new(false));
        
        // Get the QPC frequency for timestamp calculations
        let mut qpf_frequency: i64 = 0;
        unsafe {
            QueryPerformanceFrequency(&mut qpf_frequency);
        }
        println!("QPC frequency: {}", qpf_frequency);
        
        // Create buffer for AudioSample structs
        // Use a smaller capacity since AudioSample structs are larger
        let buffer_size = 100; // This should be tuned based on expected packet sizes
        
        let rb = HeapRb::<AudioSample>::new(buffer_size);
        let (mut producer, consumer) = rb.split();
        
        // Create control channel
        let (control_sender, control_receiver) = channel();
        
        // Clone references for the thread
        let thread_audio_source = audio_source.clone();
        let thread_sample_rate = sample_rate.clone();
        let thread_channels = channels.clone();
        let thread_bits_per_sample = bits_per_sample.clone();
        let thread_initialized = initialized.clone();
        let thread_start_qpc = start_qpc;
        let thread_qpf_frequency = qpf_frequency;

        // Create session object
        let session = MicrophoneCaptureSession::new(control_sender);
        println!("Created session");
        
        // Start audio capture thread
        thread::spawn(move || {
            // Initialize COM in this thread
            unsafe {
                if let Err(e) = CoInitializeEx(None, COINIT_APARTMENTTHREADED) {
                    eprintln!("COM init failed: {:?}", e);
                    return;
                }
                
                let mut running = false;
                let mut audio_client: Option<IAudioClient> = None;
                let mut audio_capture_client: Option<IAudioCaptureClient> = None;
                let mut event_handle = None;
                
                // Main loop to handle control messages and capture audio
                'outer: loop {
                    // Check for control messages first
                    match control_receiver.try_recv() {
                        Ok(start_signal) => {
                            running = start_signal;
                            
                            if running && audio_client.is_none() {
                                // Initialize audio capture using our helper function
                                match initialize_audio_capture(&thread_audio_source) {
                                    Ok((client, capture_client, handle, actual_sample_rate, actual_channels, actual_bits_per_sample)) => {
                                        // Store the actual format info
                                        thread_sample_rate.store(actual_sample_rate, Ordering::SeqCst);
                                        thread_channels.store(actual_channels, Ordering::SeqCst);
                                        thread_bits_per_sample.store(actual_bits_per_sample, Ordering::SeqCst);
                                        thread_initialized.store(true, Ordering::SeqCst);
                                        
                                        audio_client = Some(client);
                                        audio_capture_client = Some(capture_client);
                                        event_handle = Some(handle);
                                    },
                                    Err(e) => {
                                        eprintln!("Failed to initialize audio capture: {:?}", e);
                                        break;
                                    }
                                }
                            } else if !running && audio_client.is_some() {
                                // Stop audio capture
                                if let Some(client) = &audio_client {
                                    let _ = client.Stop();
                                }
                                
                                audio_client = None;
                                audio_capture_client = None;
                                event_handle = None;
                                break;
                            }
                        },
                        Err(std::sync::mpsc::TryRecvError::Empty) => {
                            // No control messages, continue
                        },
                        Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                            break 'outer; // Exit if channel is closed
                        }
                    }
                    
                    if !running || audio_capture_client.is_none() || event_handle.is_none() {
                        thread::sleep(std::time::Duration::from_millis(10));
                        continue;
                    }
                    
                    // Process audio data if running
                    let wait_result = WaitForSingleObject(event_handle.unwrap(), 100); // 100ms timeout
                    
                    match wait_result {
                        WAIT_OBJECT_0 => {
                            let capture_client = audio_capture_client.as_ref().unwrap();
                            
                            // Get the buffer from the audio client
                            let mut buffer_data_ptr = std::ptr::null_mut();
                            let mut num_frames_available = 0;
                            let mut flags = 0;
                            let mut qpc_position: u64 = 0;
                            
                            let capture_result = capture_client.GetBuffer(
                                &mut buffer_data_ptr,
                                &mut num_frames_available,
                                &mut flags,
                                Some(&mut qpc_position),
                                Some(&mut qpc_position), // Request QPC timestamp
                            );
                            
                            if capture_result.is_ok() && num_frames_available > 0 {
                                // Get current channel count and bits per sample
                                let current_channels = thread_channels.load(Ordering::SeqCst);
                                let current_bits_per_sample = thread_bits_per_sample.load(Ordering::SeqCst);
                                let bytes_per_sample = (current_bits_per_sample / 8) as usize;
                                let current_sample_rate = thread_sample_rate.load(Ordering::SeqCst);
                                
                                // Calculate buffer size in bytes
                                let buffer_size_bytes = num_frames_available as usize * current_channels as usize * bytes_per_sample;
                                
                                // Convert the buffer to a slice of bytes
                                let buffer_slice = std::slice::from_raw_parts(
                                    buffer_data_ptr as *const u8,
                                    buffer_size_bytes
                                );
                                
                                // Calculate timestamp and duration
                                // Convert QPC timestamp to relative time in 100ns units
                                let qpc_signed = qpc_position as i64;
                                let relative_timestamp_hns = ((qpc_signed - thread_start_qpc) * REFTIMES_PER_SEC) / thread_qpf_frequency;
                                let packet_duration_hns = (num_frames_available as i64 * REFTIMES_PER_SEC) / current_sample_rate as i64;
                                
                                // Create TimeSpan objects
                                let timestamp = TimeSpan { Duration: relative_timestamp_hns };
                                let duration = TimeSpan { Duration: packet_duration_hns };
                                
                                // Create an AudioSample and push it to the ring buffer
                                let audio_sample = AudioSample {
                                    data: buffer_slice.to_vec(),
                                    timestamp,
                                    duration,
                                    frames: num_frames_available,
                                };
                                
                                if producer.try_push(audio_sample).is_err() {
                                    println!("Buffer overflow - consumer not keeping up");
                                }
                                
                                // Release the buffer
                                if let Err(e) = capture_client.ReleaseBuffer(num_frames_available) {
                                    eprintln!("Failed to release buffer: {:?}", e);
                                }
                            }                          
                        },
                        WAIT_TIMEOUT => {
                            // Normal timeout, just continue waiting
                        },
                        _ => {
                            eprintln!("Wait error");
                            break;
                        }
                    }
                }
                
                // Clean up
                if let Some(client) = &audio_client {
                    let _ = client.Stop();
                }
                
                CoUninitialize();
            }
        });
        
        Ok(Self {
            consumer,
            session,
            sample_rate,
            channels,
            bits_per_sample,
            initialized,
            start_qpc,
            qpf_frequency,
        })
    }
    
    // Wait until initialization completes
    pub fn wait_for_initialization(&self, timeout_ms: u64) -> Result<()> {
        let start = std::time::Instant::now();
        while !self.initialized.load(Ordering::SeqCst) {
            // Check for timeout
            if start.elapsed() > std::time::Duration::from_millis(timeout_ms) {
                return Err(windows::core::Error::from(E_FAIL));
            }
            
            // Sleep to avoid spinning
            thread::sleep(std::time::Duration::from_millis(10));
        }
        
        Ok(())
    }
    
    // Convenience method that starts capture and can wait for initialization
    pub fn start_capture_and_wait(&mut self, timeout_ms: u64) -> Result<()> {
        self.session.StartCapture()?;
        self.wait_for_initialization(timeout_ms)
    }
    
    pub fn session(&self) -> &MicrophoneCaptureSession {
        &self.session
    }
    
    // Method to retrieve audio samples - now returns AudioSample structs
    pub fn try_get_audio_sample(&mut self) -> Option<AudioSample> {
        if !self.consumer.is_empty() {
            self.consumer.try_pop()
        } else {
            None
        }
    }
    
    pub fn stop_capture(&mut self) -> Result<()> {
        self.session.StopCapture()
    }
    
    // Check if initialization is complete
    pub fn is_initialized(&self) -> bool {
        self.initialized.load(Ordering::SeqCst)
    }
    
    // Getters for audio format information
    pub fn get_sample_rate(&self) -> u32 {
        self.sample_rate.load(Ordering::SeqCst)
    }
    
    pub fn get_channels(&self) -> u16 {
        self.channels.load(Ordering::SeqCst)
    }
    
    pub fn get_bits_per_sample(&self) -> u16 {
        self.bits_per_sample.load(Ordering::SeqCst)
    }
    
    // Method to calculate time from QPC value
    pub fn qpc_to_time(&self, qpc: i64) -> TimeSpan {
        let relative_time_hns = ((qpc - self.start_qpc) * REFTIMES_PER_SEC) / self.qpf_frequency;
        TimeSpan { Duration: relative_time_hns }
    }
}