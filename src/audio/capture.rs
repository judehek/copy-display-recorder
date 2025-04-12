// src/audio_capture.rs
// Minimal version that just prints when Windows has audio data available

use std::thread;
use windows::{
    core::*,
    Win32::{
        Foundation::{ WAIT_OBJECT_0, WAIT_TIMEOUT},
        Media::Audio::{
            eConsole, eRender, IAudioCaptureClient, IAudioClient, IMMDeviceEnumerator, MMDeviceEnumerator, AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_EVENTCALLBACK, AUDCLNT_STREAMFLAGS_LOOPBACK
        },
        System::{
            Com::{CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_ALL, COINIT_APARTMENTTHREADED},
            Threading::{CreateEventW, WaitForSingleObject},
        },
    },
};

use ringbuf::traits::Split;
use ringbuf::producer::Producer;
use ringbuf::wrap::caching::Caching;

use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};


// Constant used within this module
const REFTIMES_PER_MILLISEC: i64 = 10000; // 100ns units per millisecond

type ProducerType = Caching<Arc<ringbuf::SharedRb<ringbuf::storage::Heap<f32>>>, true, false>;


pub fn setup_audio_capture() -> Result<(thread::JoinHandle<()>, Arc<AtomicBool>)> {

    let buffer_size = 48000 * 2 * 2; // sample_rate * channels * seconds
    let rb = ringbuf::HeapRb::<f32>::new(buffer_size);
    let (mut producer, mut consumer) = rb.split(); 
    let stop_flag = Arc::new(AtomicBool::new(true));
    let thread_stop_flag = stop_flag.clone();  

    // Create a thread for audio operations
    let handle = thread::spawn(move || {
        // Initialize COM in this thread
        unsafe {
            if let Err(e) = CoInitializeEx(None, COINIT_APARTMENTTHREADED) {
                eprintln!("COM init failed: {:?}", e);
                return;
            }
            
            // Run the audio capture loop
            if let Err(e) = run_audio_capture_loop(producer, thread_stop_flag) {
                eprintln!("Audio capture loop error: {:?}", e);
            }
            
            // Clean up COM
            CoUninitialize();
        }
    });
    
    Ok((handle, stop_flag))
}

// Simple function to run the audio capture loop
unsafe fn run_audio_capture_loop(mut producer: ProducerType, stop_flag: Arc<AtomicBool>) -> Result<()> {
    // 1. Get the audio capture device (microphone)
    let device_enumerator: IMMDeviceEnumerator =
        CoCreateInstance(&MMDeviceEnumerator, None, CLSCTX_ALL)?;
    let device = device_enumerator.GetDefaultAudioEndpoint(eRender, eConsole)?;
    
    // 2. Create the audio client
    let audio_client: IAudioClient = device.Activate(CLSCTX_ALL, None)?;
    
    // 3. Get the audio format
    let wave_format_ptr = audio_client.GetMixFormat()?;
    
    // 4. Create an event for buffer notifications
    let event_handle = CreateEventW(None, false, false, None)?;
    if event_handle.is_invalid() {
        CoTaskMemFree(Some(wave_format_ptr as *const _));
        return Err(Error::from_win32());
    }
    
    // 5. Set up a 10ms buffer
    let buffer_duration_ms = 10;
    let buffer_duration_100ns = buffer_duration_ms * REFTIMES_PER_MILLISEC;
    
    // 6. Initialize the audio client
    let init_result = audio_client.Initialize(
        AUDCLNT_SHAREMODE_SHARED,           // Share audio with other apps
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_LOOPBACK,  // Use event for notifications
        buffer_duration_100ns,              // Buffer size in 100ns units
        0,                                  // No offset
        wave_format_ptr,                    // Audio format
        None,                               // No session GUID
    );
    
    // 7. Free the format pointer (no longer needed)
    CoTaskMemFree(Some(wave_format_ptr as *const _));
    init_result?; // Check if initialize succeeded
    
    // 8. Set the event to be triggered when buffer has data
    audio_client.SetEventHandle(event_handle)?;
    
    // 9. Get the buffer size
    let buffer_size_frames = audio_client.GetBufferSize()?;
    println!("Audio capture setup complete. Buffer size: {} frames", buffer_size_frames);
    
    // 10. Start audio capture
    audio_client.Start()?;
    println!("Audio capture started. Waiting for buffer events...");
    let audio_capture_client = audio_client.GetService::<IAudioCaptureClient>()?;
    
    // Simple event loop - just prints when Windows signals it has audio data

    let mut flags = 0;
    let mut position: u64 = 0;
    let channels = 2; // Set this based on your actual format   

    while !stop_flag.load(Ordering::Relaxed) {
        let wait_result = WaitForSingleObject(event_handle, 2000); // 2 second timeout
        
        match wait_result {
            WAIT_OBJECT_0 => {
                // Get the buffer from the audio client
                let mut buffer_data_ptr = std::ptr::null_mut();
                let mut num_frames_available = 0;
                let capture_result = audio_capture_client.GetBuffer(
                    &mut buffer_data_ptr,
                    &mut num_frames_available,
                    &mut flags,
                    Some(&mut position),
                    None
                );
                
                if capture_result.is_ok() && num_frames_available > 0 {
                    // Convert the buffer to a slice of f32 values
                    let buffer_slice = std::slice::from_raw_parts(
                        buffer_data_ptr as *const f32,
                        num_frames_available as usize * channels as usize
                    );
                    
                    // Push the data to the ring buffer
                    for &sample in buffer_slice {
                        if producer.try_push(sample).is_err() {
                            println!("Buffer overflow - consumer not keeping up");
                        }
                    }
                    
                    // Release the buffer
                    audio_capture_client.ReleaseBuffer(num_frames_available)?;
                    
                    println!("Pushed {} frames to the buffer", num_frames_available);
                }                          
            },
            WAIT_TIMEOUT => {
                // Normal timeout, just continue waiting
                println!("Still waiting for audio data...");
            },
            _ => {
                eprintln!("Wait error");
                break;
            }
        }
        
        // In a real app, you'd have some way to break this loop
        // For now, we'll just keep running until the program is terminated
    }
    
    // This code is never reached in this simple example
    // audio_client.Stop()?;
    // CloseHandle(event_handle);
    
    Ok(())
}