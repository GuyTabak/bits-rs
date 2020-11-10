use std::ptr::null_mut;

use winapi::{Class, Interface};
use winapi::ctypes::c_void;
use winapi::shared::guiddef::GUID;
use winapi::shared::winerror::SUCCEEDED;
use winapi::shared::wtypesbase::CLSCTX_LOCAL_SERVER;
use winapi::um::bits::{BackgroundCopyManager, BG_JOB_TYPE_DOWNLOAD, IBackgroundCopyJob, IBackgroundCopyManager};
use winapi::um::combaseapi::{CoCreateInstance, CoInitializeEx};
use winapi::um::commctrl::HRESULT;
use winapi::um::objbase::COINIT_APARTMENTTHREADED;


use url::Url;
use std::ffi::OsStr;
use std::os::windows::ffi::OsStrExt;


///This will place the address of the service in the 'g_pbcm' variable.
///If the service is not running, it will start it.
fn connect_to_bits() -> Result<*mut IBackgroundCopyManager, String> {
    let mut hr: HRESULT;
    let mut g_pbcm: *mut IBackgroundCopyManager = unsafe { std::mem::zeroed() };


    hr = unsafe {
        CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED)
    };

    if !SUCCEEDED(hr) {
        return Err(format!("Failed to CoInitializeEx, with error code: {}", hr));
    }

    hr = unsafe {
        CoCreateInstance(&BackgroundCopyManager::uuidof(), null_mut(), CLSCTX_LOCAL_SERVER, &IBackgroundCopyManager::uuidof(), &mut g_pbcm as *mut *mut IBackgroundCopyManager as *mut *mut c_void)
    };

    if !SUCCEEDED(hr) {
        return Err(format!("Failed to CoCreateInstance, with error code: {}", hr));
    }

    Ok(g_pbcm)
}

fn create_bits_job(g_pbcm: *mut IBackgroundCopyManager) -> Result<*mut IBackgroundCopyJob, String> {
    let hr: HRESULT;
    let mut job_id: GUID = unsafe { std::mem::zeroed() };
    let mut p_job: *mut IBackgroundCopyJob = unsafe { std::mem::zeroed() };

    //todo: display name
    hr = unsafe {
        (&mut *g_pbcm).CreateJob(&mut 1, BG_JOB_TYPE_DOWNLOAD, &mut job_id, &mut p_job)
    };

    if !SUCCEEDED(hr) {
        return Err(format!("Failed to create job, with error code: {}", hr));
    }

    Ok(p_job)
}

fn add_file(bits_jon: *mut IBackgroundCopyJob, file_url: &str, save_path: &str) -> Result<(), String>{
    let download_url = Url::parse(file_url).unwrap().to_string();
    let hr: HRESULT;

    hr = unsafe{ (&mut *bits_jon).AddFile(to_wchar(download_url.as_str()).as_ptr() , to_wchar(save_path).as_ptr())};

    if !SUCCEEDED(hr){
        return Err(format!("Failed to add file to job, with error code: {}", hr));
    }

    Ok(())
}


fn to_wchar(str : &str) -> Vec<u16> {
    OsStr::new(str).encode_wide(). chain(Some(0).into_iter()).collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_connect_to_bits() {
        connect_to_bits().unwrap();
    }

    #[test]
    fn test_create_bits_job() {
        let bits_service = connect_to_bits().unwrap();
        create_bits_job(bits_service).unwrap();
    }

    #[test]
    fn test_add_file(){
        // Note, you need to run this test as admin
        let bits_service = connect_to_bits().unwrap();
        let bits_job = create_bits_job(bits_service).unwrap();
        add_file(bits_job, "http://speedtest.tele2.net/", "C:\\temp\\file_name.zip").unwrap();

    }
}
