use windows::{
    core::Result,
    Graphics::SizeInt32,
    Storage::Streams::IRandomAccessStream,
    Win32::{
        Graphics::{Direct3D11::ID3D11Device, Gdi::HMONITOR},
    },
};

use crate::audio::AudioSource;

pub trait VideoEncoderSessionFactory {
    fn create_session(
        &self,
        d3d_device: ID3D11Device,
        monitor_handle: HMONITOR,
        resolution: SizeInt32,
        bit_rate: u32,
        frame_rate: u32,
        stream: IRandomAccessStream,
        audio_source: AudioSource,
    ) -> Result<Box<dyn VideoEncodingSession>>;
}

pub trait VideoEncodingSession {
    fn start(&mut self) -> Result<()>;
    fn stop(&mut self) -> Result<()>;
}
