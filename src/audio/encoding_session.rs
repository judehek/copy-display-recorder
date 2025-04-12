// use std::sync::{Arc, mpsc::{self, Receiver, Sender}};
// use std::thread;
// use std::time::Duration;

// use windows::{
//     core::Result,
//     Storage::Streams::IRandomAccessStream,
//     Win32::Media::MediaFoundation::{IMFSample, IMFSinkWriter},
// };

// use ringbuf::{consumer::Consumer as RingConsumer, storage::Heap, wrap::caching::Caching, SharedRb};

// use crate::audio::encoder::{AudioEncoder, AudioEncoderInputSample};

// type ConsumerType = Caching<Arc<SharedRb<Heap<f32>>>, false, true>;

// pub struct AudioEncodingSession {
//     encoder: AudioEncoder,
//     audio_consumer: ConsumerType,
//     sample_writer: Arc<AudioSampleWriter>,
//     stop_flag: Arc<std::sync::atomic::AtomicBool>,
//     buffer_thread_handle: Option<thread::JoinHandle<()>>,
// }

// // Writer to handle sending audio samples to the sink
// struct AudioSampleWriter {
//     sink_writer: IMFSinkWriter,
//     stream_index: u32,
// }

// impl AudioEncodingSession {
//     pub fn new(
//         stream: IRandomAccessStream,
//         audio_consumer: Box<dyn RingConsumer<Item = f32> + Send>,
//         sample_rate: u32,
//         channels: u32,
//         bit_rate: u32,
//     ) -> Result<Self> {
//         // Create the encoder
//         let mut encoder = AudioEncoder::new(sample_rate, channels, bit_rate)?;
        
//         // Get output type for the sink writer
//         let output_type = encoder.output_type().clone();
        
//         // Create sample writer
//         let sample_writer = Arc::new(AudioSampleWriter::new(stream, &output_type)?);
        
//         // Create shared stop flag
//         let stop_flag = Arc::new(std::sync::atomic::AtomicBool::new(false));
        
//         // Set up callbacks
//         let (tx, rx) = mpsc::channel();
//         let sample_writer_clone = sample_writer.clone();
//         encoder.set_sample_rendered_callback(move |output_sample| {
//             sample_writer_clone.write(output_sample.sample())
//         });

//         // We'll set up the sample_requested_callback in start() when we spin up 
//         // the buffer thread
        
//         Ok(Self {
//             encoder,
//             audio_consumer: Some(audio_consumer),
//             sample_writer,
//             stop_flag,
//             buffer_thread_handle: None,
//         })
//     }
    
//     pub fn start(&mut self) -> Result<()> {
//         // Flag that we're running
//         self.stop_flag.store(false, std::sync::atomic::Ordering::SeqCst);
        
//         // Set up the buffer thread
//         let mut audio_consumer = self.audio_consumer.take()
//             .expect("Audio consumer missing, has the session already been started?");
        
//         let stop_flag = self.stop_flag.clone();
        
//         // Create a channel to send processed audio samples to the encoder
//         let (audio_tx, audio_rx) = mpsc::channel();
        
//         // Start buffer thread to read from ring buffer and prepare samples
//         self.buffer_thread_handle = Some(thread::spawn(move || {
//             let mut buffer: Vec<f32> = Vec::with_capacity(1024);
//             let mut next_timestamp = 0i64;
            
//             while !stop_flag.load(std::sync::atomic::Ordering::SeqCst) {
//                 // Try to read up to 1024 samples
//                 while buffer.len() < 1024 {
//                     match audio_consumer.pop() {
//                         Some(sample) => buffer.push(sample),
//                         None => break,  // No more samples available right now
//                     }
//                 }
                
//                 // If we got any samples, create an input sample and send it
//                 if !buffer.is_empty() {
//                     // Convert f32 samples to i16 PCM
//                     let pcm_data = convert_f32_to_i16_pcm(&buffer);
                    
//                     // Calculate duration in 100ns units
//                     // 10_000_000 (100ns in a second) / sample_rate = 100ns per sample
//                     let samples_per_channel = buffer.len() / 2; // Assuming stereo
//                     let duration = 10_000_000 * samples_per_channel as i64 / 48000; // Assuming 48kHz
                    
//                     // Create and send the input sample
//                     let input_sample = AudioEncoderInputSample::new(
//                         pcm_data,
//                         next_timestamp,
//                         duration,
//                     );
                    
//                     // Update next timestamp
//                     next_timestamp += duration;
                    
//                     // Send the sample
//                     if audio_tx.send(input_sample).is_err() {
//                         // Encoder has been dropped, exit thread
//                         break;
//                     }
                    
//                     // Clear the buffer for next batch
//                     buffer.clear();
//                 } else {
//                     // No samples available, wait a bit
//                     thread::sleep(Duration::from_millis(5));
//                 }
//             }
//         }));
        
//         // Set up the encoder's sample requested callback to pull from our channel
//         self.encoder.set_sample_requested_callback(move || -> Result<Option<AudioEncoderInputSample>> {
//             match audio_rx.recv_timeout(Duration::from_millis(100)) {
//                 Ok(sample) => Ok(Some(sample)),
//                 Err(mpsc::RecvTimeoutError::Timeout) => {
//                     // No sample available yet, but we're still running
//                     // Return empty sample to keep encoder running
//                     Ok(Some(AudioEncoderInputSample::new(
//                         Vec::new(),  // Empty data (silence)
//                         next_timestamp,
//                         10_000_000 / 100, // 10ms of silence
//                     )))
//                 },
//                 Err(mpsc::RecvTimeoutError::Disconnected) => {
//                     // Buffer thread has exited
//                     Ok(None)
//                 }
//             }
//         });
        
//         // Start the encoder
//         self.encoder.try_start()?;
        
//         Ok(())
//     }
    
//     pub fn stop(&mut self) -> Result<()> {
//         // Signal threads to stop
//         self.stop_flag.store(true, std::sync::atomic::Ordering::SeqCst);
        
//         // Stop the encoder
//         self.encoder.stop()?;
        
//         // Wait for buffer thread to finish
//         if let Some(handle) = self.buffer_thread_handle.take() {
//             let _ = handle.join();
//         }
        
//         Ok(())
//     }
// }

// impl AudioSampleWriter {
//     pub fn new(
//         stream: IRandomAccessStream,
//         output_type: &windows::Win32::Media::MediaFoundation::IMFMediaType,
//     ) -> Result<Self> {
//         // Create attributes
//         unsafe {
//             let attributes = windows::Win32::Media::MediaFoundation::MFCreateAttributes(None, 0)?;
            
//             // Create byte stream from random access stream
//             let byte_stream = windows::Win32::Media::MediaFoundation::MFCreateMFByteStreamOnStreamEx(&stream)?;
            
//             // Create sink writer
//             let sink_writer = windows::Win32::Media::MediaFoundation::MFCreateSinkWriterFromURL(
//                 &windows::core::HSTRING::from(".mp4"),
//                 &byte_stream,
//                 &attributes,
//             )?;
            
//             // Add stream for AAC audio
//             let stream_index = sink_writer.AddStream(output_type)?;
            
//             // Set input media type (same as output for audio)
//             sink_writer.SetInputMediaType(
//                 stream_index,
//                 output_type,
//                 &attributes,
//             )?;
            
//             // Begin writing
//             sink_writer.BeginWriting()?;
            
//             Ok(Self {
//                 sink_writer,
//                 stream_index,
//             })
//         }
//     }
    
//     pub fn write(&self, sample: &IMFSample) -> Result<()> {
//         unsafe {
//             self.sink_writer.WriteSample(self.stream_index, sample)
//         }
//     }
    
//     pub fn finalize(&self) -> Result<()> {
//         unsafe {
//             self.sink_writer.Finalize()
//         }
//     }
// }

// impl Drop for AudioSampleWriter {
//     fn drop(&mut self) {
//         let _ = self.finalize();
//     }
// }

// // Helper function to convert f32 audio samples to i16 PCM
// fn convert_f32_to_i16_pcm(samples: &[f32]) -> Vec<u8> {
//     let mut pcm_data = Vec::with_capacity(samples.len() * 2); // 2 bytes per sample
    
//     for &sample in samples {
//         // Clamp to [-1.0, 1.0] and convert to i16 range
//         let scaled = (sample.max(-1.0).min(1.0) * 32767.0) as i16;
        
//         // Add as little-endian bytes
//         pcm_data.push((scaled & 0xFF) as u8);        // Low byte
//         pcm_data.push(((scaled >> 8) & 0xFF) as u8); // High byte
//     }
    
//     pcm_data
// }