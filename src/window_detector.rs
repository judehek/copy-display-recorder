use windows::Win32::UI::Accessibility::{
    SetWinEventHook, HWINEVENTHOOK
};
use windows::Win32::Foundation::HWND;
use windows::Win32::UI::WindowsAndMessaging::{
    GetWindowTextW, GetWindowTextLengthW, WINEVENT_OUTOFCONTEXT, EVENT_SYSTEM_FOREGROUND
};

use std::ffi::OsString;
use std::os::windows::ffi::OsStringExt;
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, Ordering};

static mut RECORDING_FLAG: Option<*const Arc<AtomicBool>> = None;

fn should_record_window() -> bool {
    // Define your logic to determine if the window should be recorded
    // For example, check if the window title contains a specific keyword
    return true;
}

pub fn start_window_change_detector(is_recording_window: Arc<AtomicBool>) -> HWINEVENTHOOK {

    // TODO: memory leak?
    unsafe {
        let leaked = Box::leak(Box::new(is_recording_window));
        RECORDING_FLAG = Some(leaked);
    }

    unsafe extern "system" fn foreground_change_hook(
        _hook: HWINEVENTHOOK,
        _event: u32,
        hwnd: HWND,
        _id_object: i32,
        _id_child: i32,
        _id_event_thread: u32,
        _event_time: u32,
    ) {
        // Get our atomic from the context
        if let Some(ptr) = RECORDING_FLAG {
                let is_recording_window = &*ptr;
            
            // Get window title
            let length = GetWindowTextLengthW(hwnd);
            if length > 0 {
                let mut buffer = vec![0u16; (length + 1) as usize];
                if GetWindowTextW(hwnd, &mut buffer) > 0 {
                    if let Some(pos) = buffer.iter().position(|&c| c == 0) {
                        buffer.truncate(pos);
                    }
                    let window_title = OsString::from_wide(&buffer).to_string_lossy().into_owned();
                    
                    // Update the atomic based on window title
                    let should_record = should_record_window();
                    is_recording_window.store(should_record, Ordering::SeqCst);
                    
                    // Print for debugging
                    println!("Window changed to: {}, recording: {}", window_title, should_record);
                }
            }
        }
    }

    // TODO: resource leak? I think we need to unhook on recording stop()
    unsafe {
        SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            None,
            Some(foreground_change_hook),
            0,
            0,
            WINEVENT_OUTOFCONTEXT,
        )
    }
}