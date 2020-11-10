use winapi::um::objbase::COINIT_APARTMENTTHREADED;
use winapi::um::combaseapi::{CoInitializeEx, CoCreateInstance};
use winapi::um::commctrl::HRESULT;
use winapi::um::bits::{IBackgroundCopyManager, BackgroundCopyManager};
use winapi::shared::wtypesbase::CLSCTX_LOCAL_SERVER;
use winapi::shared::winerror::SUCCEEDED;
use winapi::{Interface, Class};
use winapi::ctypes::c_void;
use std::ptr::null_mut;

fn connect_to_bits() -> Result<(), String>{
    let mut hr : HRESULT;
    let mut g_pbcm: *mut IBackgroundCopyManager = unsafe { std::mem::zeroed() };


    hr = unsafe{
        CoInitializeEx(null_mut(),COINIT_APARTMENTTHREADED)
    };

    if !SUCCEEDED(hr){
        return Err(format!("Failed to CoInitializeEx, with error code: {}", hr))
    }

    hr = unsafe{
        CoCreateInstance(&BackgroundCopyManager::uuidof(), null_mut(), CLSCTX_LOCAL_SERVER, &IBackgroundCopyManager::uuidof(), &mut g_pbcm as *mut *mut IBackgroundCopyManager as *mut *mut c_void)
    };

    if !SUCCEEDED(hr){
        return Err(format!("Failed to CoCreateInstance, with error code: {}", hr))
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_connect_to_bits() {
        let res = connect_to_bits().unwrap();
        assert_eq!(res, ());
    }


}
