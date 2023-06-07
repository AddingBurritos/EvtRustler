use windows::Win32::Foundation::*;
use windows::Win32::System::EventLog::*;
use windows::core::*;

use crate::managed_variant::*;
use crate::provider::GuidWrapper;


pub fn evt_open_publisher_enum() -> Result<EVT_HANDLE> {
    let publisher_enum_result = unsafe { 
        EvtOpenPublisherEnum(
            None, 
            0
        )   
    };
    let publisher_enum_handle = match publisher_enum_result {
        Ok(enum_handle) => enum_handle,
        Err(e) => {
            println!("Failed to open publisher metadata enumerator: {}", e.message());
            return Err(e);
        }
    };
    if publisher_enum_handle.is_invalid() {
        println!("Publisher enumerator handle is invalid.");
        return Err(Error::from_win32());
    }
    Ok(publisher_enum_handle)
}

pub fn get_property(h_array: &EVT_HANDLE, array_ix: u32, property_id: EVT_PUBLISHER_METADATA_PROPERTY_ID) -> Result<ManagedEvtVariant> {
    
	// Get size of property
	let mut buffer_used: u32 = 0;
    let status = unsafe {
        EvtGetObjectArrayProperty(
            h_array.0,
            property_id.0 as u32,
            array_ix,
            0,
            0,
            None,
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        if win_error.code() != ERROR_INSUFFICIENT_BUFFER.into() {
            return Err(win_error);
        }
    }

	// Get property
    let buffer_size: u32 = buffer_used;
    let mut variant = EVT_VARIANT {
        Anonymous: EVT_VARIANT_0 {
            ByteVal: 0,
        },
        Count: 0,
        Type: EvtVarTypeByte.0 as u32,
    };
    let status = unsafe {
        EvtGetObjectArrayProperty(
            h_array.0,
            property_id.0 as u32,
            array_ix,
            0,
            buffer_size,
            Some(&mut variant),
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        return Err(win_error);
    }

    match property_id {
        EvtPublisherMetadataChannelReferencePath => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(ManagedEvtVariant::from_string(new_string))
        },
        EvtPublisherMetadataLevelName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(ManagedEvtVariant::from_string(new_string))
        },
        EvtPublisherMetadataTaskName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(ManagedEvtVariant::from_string(new_string))
        },
        EvtPublisherMetadataOpcodeName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(ManagedEvtVariant::from_string(new_string))
        },
        EvtPublisherMetadataKeywordName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(ManagedEvtVariant::from_string(new_string))
        },
        EvtPublisherMetadataTaskEventGuid => {
            let new_guid = unsafe { *variant.Anonymous.GuidVal };
            Ok(ManagedEvtVariant::from_guid(new_guid))
        },
        // Lazily adding a wild card. This should be updated if new EVT_PUBLISHER_METADATA_PROPERTY_IDs are supported
        _ => Ok(ManagedEvtVariant::from_variant(variant))
    }
}
pub fn evt_next_publisher_id(publisher_enum_handle: &EVT_HANDLE) -> Result<String> {
    let mut buffer_used: u32 = 0;
		
    // Get size of provider name
    let next_status = unsafe {
        EvtNextPublisherId(
            *publisher_enum_handle,
            None,
            &mut buffer_used
        )
    };
    
    if !next_status.as_bool() {
        let win_error = Error::from_win32();
        if win_error.code() != ERROR_INSUFFICIENT_BUFFER.into() {
            return Err(win_error);
        }
    }

    // Get provider name
    let buffer_size: u32 = buffer_used;
    let mut provider_buffer: Vec<u16> = vec![0; buffer_size.try_into().unwrap()];
    let next_status = unsafe {
        EvtNextPublisherId(
            *publisher_enum_handle,
            Some(&mut provider_buffer),
            &mut buffer_used
        )
    };
    if !next_status.as_bool() {
        let win_error = Error::from_win32();
        println!("EvtNextPublisherID Error 2: {}", win_error.message());
        return Err(win_error);
    }

    // Convert buffer to String
    let provider_name = String::from_utf16_lossy(&provider_buffer[..(buffer_used-1) as usize]);

    Ok(provider_name)
}

pub fn evt_open_publisher_metadata(provider_vec: Vec<u16>, archive_path: Option<Vec<u16>>) -> Result<EVT_HANDLE> {
    let provider_name = PCWSTR(provider_vec.as_ptr());
    let result = match archive_path {
        Some(path) =>  unsafe {
            EvtOpenPublisherMetadata(
                None,
                provider_name,
                PCWSTR(path.as_ptr()),
                0,
                0
            )
        },
        None =>  unsafe {
            EvtOpenPublisherMetadata(
                None,
                provider_name,
                None,
                0,
                0
            )
        }
    };
    
    match result {
        Ok(handle) => return Ok(handle),
        Err(e) => {
            let provider_name_str = unsafe {provider_name.to_string().unwrap()};
            if e.message() == "The system cannot find the file specified." {
                println!("Couldn't get handle to provider '{}' because of error: {}",&provider_name_str, e.message());
                let reg_path = "HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Services\\EventLog";
                println!("Look recursively in the registry at {} for {}. Dollars to doughnuts you don't have the file represented by the 'EventMessageFile' key", reg_path, provider_name_str);
            } else if e.message() == "The specified resource type cannot be found in the image file." {
                println!("Couldn't get handle to provider '{}' because of error: {}", &provider_name_str, e.message());
                println!("Full disclosure: I don't know what that error means.");
            } else {
                println!("{}", e.message());
            }
            return Err(e);
        }
    }
}

pub fn evt_get_publisher_metadata_property(h_provider: &EVT_HANDLE, property_id: EVT_PUBLISHER_METADATA_PROPERTY_ID) -> Result<EVT_HANDLE> {
    // Determine necessary size of buffer to hold array of channels for provider.
    let mut buffer_used: u32 = 0;
    let status = unsafe {
        EvtGetPublisherMetadataProperty(
            *h_provider,
            property_id,
            0,
            0,
            None,
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        if win_error.code() != ERROR_INSUFFICIENT_BUFFER.into() {
            //println!("EvtGetPublisherMetadata Error: {}", win_error.message());
            return Err(win_error);
        }
    }

    // Properly size buffer and fill with array of provider's channels.
    let buffer_size: u32 = buffer_used;
    let mut variant = EVT_VARIANT {
        Anonymous: EVT_VARIANT_0 {
            ByteVal: 0,
        },
        Count: 0,
        Type: EvtVarTypeByte.0 as u32,
    };
    let status = unsafe {
        EvtGetPublisherMetadataProperty(
            *h_provider,
            property_id,
            0,
            buffer_size,
            Some(&mut variant),
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        //println!("EvtGetPublisherMetadataPropery Error 2. Skipping provider '{}': {}", &provider, win_error.message());
        return Err(win_error);
    }
    let handle = unsafe {variant.Anonymous.EvtHandleVal};
    Ok(handle)
}

pub fn evt_get_object_array_size(h_array: &EVT_HANDLE) -> Result<u32> {
    let mut array_size: u32 = 0;
    let status = unsafe {
        EvtGetObjectArraySize(
            h_array.0,
            &mut array_size,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        return Err(win_error);
    }
    Ok(array_size)
}
