use windows::core::{Interface, Result};
use windows::Graphics::DirectX::Direct3D11::{IDirect3DDevice, IDirect3DSurface};
use windows::Win32::Graphics::Direct3D11::{ID3D11Multithread, ID3D11Texture2D, D3D11_CREATE_DEVICE_DEBUG};
use windows::Win32::Graphics::Dxgi::IDXGISurface;
use windows::Win32::Graphics::{
    Direct3D::{D3D_DRIVER_TYPE, D3D_DRIVER_TYPE_HARDWARE, D3D_DRIVER_TYPE_WARP},
    Direct3D11::{
        D3D11CreateDevice, ID3D11Device, D3D11_CREATE_DEVICE_BGRA_SUPPORT,
        D3D11_CREATE_DEVICE_FLAG, D3D11_SDK_VERSION,
    },
    Dxgi::{IDXGIDevice, DXGI_ERROR_UNSUPPORTED},
};
use windows::Win32::System::WinRT::Direct3D11::{
    CreateDirect3D11DeviceFromDXGIDevice, CreateDirect3D11SurfaceFromDXGISurface,
    IDirect3DDxgiInterfaceAccess,
};

fn create_d3d_device_with_type(
    driver_type: D3D_DRIVER_TYPE,
    flags: D3D11_CREATE_DEVICE_FLAG,
    device: *mut Option<ID3D11Device>,
) -> Result<()> {
    unsafe {
        D3D11CreateDevice(
            None,
            driver_type,
            windows::Win32::Foundation::HMODULE(std::ptr::null_mut()),
            flags,
            None,
            D3D11_SDK_VERSION,
            Some(device),
            None,
            None,
        )
    }
}

pub fn create_d3d_device() -> Result<ID3D11Device> {
    let mut device = None;
    let flags = {
        let mut flags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
        if true {
            flags |= D3D11_CREATE_DEVICE_DEBUG;
        }
        flags
    };
    let mut result = create_d3d_device_with_type(D3D_DRIVER_TYPE_HARDWARE, flags, &mut device);
    if let Err(error) = &result {
        if error.code() == DXGI_ERROR_UNSUPPORTED {
            result = create_d3d_device_with_type(D3D_DRIVER_TYPE_WARP, flags, &mut device);
        }
    }
    result?;
    Ok(device.unwrap())
}

pub fn set_multithread_protected(d3d_device: &ID3D11Device, enable: bool) -> Result<bool> {
    // Cast the device to the ID3D11Multithread interface
    let multithread: ID3D11Multithread = d3d_device.cast()?;
    
    // Call SetMultithreadProtected with the provided boolean value directly
    // This returns true if multithread protection was already enabled
    let was_protected = unsafe { multithread.SetMultithreadProtected(enable) };
    
    Ok(was_protected.into())
}

pub fn create_direct3d_device(d3d_device: &ID3D11Device) -> Result<IDirect3DDevice> {
    let dxgi_device: IDXGIDevice = d3d_device.cast()?;
    let inspectable = unsafe { CreateDirect3D11DeviceFromDXGIDevice(Some(&dxgi_device))? };
    inspectable.cast()
}

pub fn create_direct3d_surface(d3d_texture: &ID3D11Texture2D) -> Result<IDirect3DSurface> {
    let dxgi_surface: IDXGISurface = d3d_texture.cast()?;
    let inspectable = unsafe { CreateDirect3D11SurfaceFromDXGISurface(Some(&dxgi_surface))? };
    inspectable.cast()
}

pub fn get_d3d_interface_from_object<S: Interface, R: Interface>(
    object: &S,
) -> Result<R> {
    let access: IDirect3DDxgiInterfaceAccess = object.cast()?;
    let object = unsafe { access.GetInterface::<R>()? };
    Ok(object)
}
