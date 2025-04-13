use std::sync::mpsc::{channel, Receiver, Sender};
use std::thread;

use ringbuf::traits::{Consumer, Observer};
use ringbuf::wrap::caching::Caching;
use windows::Win32::Foundation::HANDLE;
use windows::{
    core::*,
    Win32::{
        Foundation::{ E_FAIL, WAIT_OBJECT_0, WAIT_TIMEOUT},
        Media::Audio::{
            eConsole, eRender, IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator, MMDeviceEnumerator, 
            AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK,
            WAVEFORMATEX
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

use super::encoding_session::AudioSource;

// Constant used within this module
const REFTIMES_PER_MILLISEC: i64 = 10000; // 100ns units per millisecond
const DEFAULT_SAMPLE_RATE: u32 = 48000;
const DEFAULT_CHANNELS: u16 = 2;
const DEFAULT_BITS_PER_SAMPLE: u16 = 32;  // Default to 32-bit float

pub struct AudioCaptureSession {
    sender: Sender<bool>, // To signal start/stop
    running: bool,
}

impl AudioCaptureSession {
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

    pub fn Close(&mut self) -> Result<()> {
        if self.running {
            self.sender.send(false).map_err(|_| windows::core::Error::from(E_FAIL))?;
            self.running = false;
        }
        Ok(())
    }
}

impl Clone for AudioCaptureSession {
    fn clone(&self) -> Self {
        Self {
            sender: self.sender.clone(),
            running: self.running,
        }
    }
}

pub struct CaptureAudioGenerator {
    consumer: Caching<Arc<ringbuf::SharedRb<ringbuf::storage::Heap<f32>>>, false, true>,
    session: AudioCaptureSession,
    sample_rate: Arc<AtomicU32>,
    channels: Arc<AtomicU16>,
    bits_per_sample: Arc<AtomicU16>,
    initialized: Arc<AtomicBool>,
    sample_position: u64,
}

unsafe fn initialize_audio_capture(audio_source: &AudioSource) -> Result<(IAudioClient, IAudioCaptureClient, HANDLE, u32, u16, u16)> {
    // Create device enumerator
    let device_enumerator: IMMDeviceEnumerator = CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
    
    // Get default audio endpoint
    let device = device_enumerator.GetDefaultAudioEndpoint(eRender, eConsole)?;
    
    // Activate audio client
    let client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
    
    // Get mix format
    let wave_format_ptr = client.GetMixFormat()?;
    
    // Extract format information before we free the pointer
    let sample_rate = (*wave_format_ptr).nSamplesPerSec;
    let channels = (*wave_format_ptr).nChannels;
    let bits_per_sample = (*wave_format_ptr).wBitsPerSample;
    
    // Create event handle
    let handle = CreateEventW(None, false, false, None)?;
    if handle.is_invalid() {
        CoTaskMemFree(Some(wave_format_ptr as *const _));
        return Err(windows::core::Error::from(E_FAIL));
    }
    
    // Set up client initialization flags based on audio source
    let mut stream_flags = AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
    
    match audio_source {
        AudioSource::Desktop => {
            stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
        },
        AudioSource::ActiveWindow => {
            // TODO: For future implementation, we would:
            // 1. Get the active window handle using GetForegroundWindow()
            // 2. Get the process ID of that window using GetWindowThreadProcessId()
            // 3. Set up an IAudioSessionManager2 and enumerate audio sessions
            // 4. Find the session matching our process ID
            // 5. Capture audio from that specific session
            
            // For now, just use loopback as a fallback
            stream_flags |= AUDCLNT_STREAMFLAGS_LOOPBACK;
            println!("Warning: Active window audio capture not fully implemented yet, falling back to desktop audio");
        }
    }
    
    // Initialize audio client
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
    
    // Free wave format memory (after we've extracted the format information)
    CoTaskMemFree(Some(wave_format_ptr as *const _));
    
    // Set event handle
    client.SetEventHandle(handle)?;
    
    // Get capture client
    let capture_client = client.GetService::<IAudioCaptureClient>()?;
    
    // Start audio client
    client.Start()?;
    
    println!("Audio capture initialized and started");
    
    Ok((client, capture_client, handle, sample_rate, channels, bits_per_sample))
}

impl CaptureAudioGenerator {
    pub fn new(audio_source: AudioSource) -> Result<Self> {
        // Create shared atomic variables with default values
        let sample_rate = Arc::new(AtomicU32::new(DEFAULT_SAMPLE_RATE));
        let channels = Arc::new(AtomicU16::new(DEFAULT_CHANNELS));
        let bits_per_sample = Arc::new(AtomicU16::new(DEFAULT_BITS_PER_SAMPLE));
        let initialized = Arc::new(AtomicBool::new(false));
        
        // Create buffer with 2 seconds capacity
        let buffer_size = DEFAULT_SAMPLE_RATE as usize * DEFAULT_CHANNELS as usize * 2; // 2 seconds
        
        let rb = HeapRb::<f32>::new(buffer_size);
        let (mut producer, consumer) = rb.split();
        
        // Create control channel
        let (control_sender, control_receiver) = channel();
        
        // Clone references for the thread
        let thread_audio_source = audio_source.clone();
        let thread_sample_rate = sample_rate.clone();
        let thread_channels = channels.clone();
        let thread_bits_per_sample = bits_per_sample.clone();
        let thread_initialized = initialized.clone();

        // Create session object
        let session = AudioCaptureSession::new(control_sender);
        println!("created session");
        
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
                            let mut position: u64 = 0;
                            
                            let capture_result = capture_client.GetBuffer(
                                &mut buffer_data_ptr,
                                &mut num_frames_available,
                                &mut flags,
                                Some(&mut position),
                                None
                            );
                            
                            if capture_result.is_ok() && num_frames_available > 0 {
                                // Get current channel count from atomic
                                let current_channels = thread_channels.load(Ordering::SeqCst);
                                
                                // Convert the buffer to a slice of f32 values
                                let buffer_slice = std::slice::from_raw_parts(
                                    buffer_data_ptr as *const f32,
                                    num_frames_available as usize * current_channels as usize
                                );
                                
                                // Push the data to the ring buffer
                                for &sample in buffer_slice {
                                    if producer.try_push(sample).is_err() {
                                        println!("Buffer overflow - consumer not keeping up");
                                    }
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
            sample_position: 0,
        })
    }
    
    // New method to wait until initialization completes
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
    
    pub fn session(&self) -> &AudioCaptureSession {
        &self.session
    }
    
    // Method to retrieve audio samples
    pub fn try_get_audio_samples(&mut self, buffer: &mut [f32], max_samples: usize) -> Option<usize> {
        let mut count = 0;
        
        while count < max_samples && !self.consumer.is_empty() {
            if let Some(sample) = self.consumer.try_pop() {
                if count < buffer.len() {
                    buffer[count] = sample;
                    count += 1;
                } else {
                    break;
                }
            }
        }
        
        // Update sample position for timing calculations
        self.sample_position += count as u64;
        
        Some(count)
    }
    
    // Get current sample position for timestamp calculations
    pub fn get_sample_position(&self) -> u64 {
        self.sample_position
    }
    
    // Get sample time in 100ns units for Media Foundation
    pub fn get_sample_time(&self) -> i64 {
        (self.sample_position * 10000000 / self.get_sample_rate() as u64) as i64
    }
    
    // Calculate duration for a given number of samples
    pub fn calculate_duration(&self, sample_count: usize) -> i64 {
        (sample_count as i64 * 10000000 / self.get_sample_rate() as i64)
    }
    
    pub fn stop_capture(&mut self) -> Result<()> {
        self.session.Close()
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
    
    // New method to get bits per sample
    pub fn get_bits_per_sample(&self) -> u16 {
        self.bits_per_sample.load(Ordering::SeqCst)
    }
}