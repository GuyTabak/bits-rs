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

    //TODO: display name
    hr = unsafe {
        (&mut *g_pbcm).CreateJob(&mut 9, BG_JOB_TYPE_DOWNLOAD, &mut job_id, &mut p_job)
    };

    if !SUCCEEDED(hr) {
        return Err(format!("Failed to create job, with error code: {}", hr));
    }

    Ok(p_job)
}

fn add_file(bits_job: *mut IBackgroundCopyJob, file_url: &str, save_path: &str) -> Result<(), String>{
    let download_url = Url::parse(file_url).unwrap().to_string();
    let hr: HRESULT;

    hr = unsafe{ (&mut *bits_job).AddFile(to_wchar(download_url.as_str()).as_ptr() , to_wchar(save_path).as_ptr())};

    if !SUCCEEDED(hr){
        return Err(format!("Failed to add file to job, with error code: {}", hr));
    }

    Ok(())
}

fn start_job(bits_job: *mut IBackgroundCopyJob) -> Result<(), String>{
    let hr: HRESULT;
    hr = unsafe{(&mut *bits_job).Resume()};

    if !SUCCEEDED(hr){
        return Err(format!("Failed to start the job, with error code: {}", hr));
    }

    Ok(())
}

fn complete_job(bits_job: *mut IBackgroundCopyJob) -> Result<(), String>{
    let hr: HRESULT;
    hr = unsafe{(&mut *bits_job).Complete()};

    if !SUCCEEDED(hr){
        return Err(format!("Failed to start the job, with error code: {}", hr));
    }

    Ok(())
}

fn to_wchar(str : &str) -> Vec<u16> {
    OsStr::new(str).encode_wide(). chain(Some(0).into_iter()).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::thread::sleep;
    use winapi::_core::time::Duration;

    #[test]
    fn test_bits_download(){
        // Note, you need to run this test as admin
        let bits_service = connect_to_bits().unwrap();
        let bits_job = create_bits_job(bits_service).unwrap();
        add_file(bits_job.clone(), "http://speedtest.tele2.net/1MB.zip", "C:\\temp\\zip_file.zip").unwrap();
        start_job(bits_job.clone()).unwrap();
        sleep(Duration::from_secs(10));
        complete_job(bits_job).unwrap();
    }
}
