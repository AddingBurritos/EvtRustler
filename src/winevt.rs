use windows::Win32::Foundation::*;
use windows::Win32::System::EventLog::*;
use windows::core::*;
use std::{mem, ptr};

use crate::managed_variant::*;
use crate::provider::*;


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

pub fn get_property(h_array: &EVT_HANDLE, array_ix: u32, property_id: EVT_PUBLISHER_METADATA_PROPERTY_ID) -> Result<VariantBuffer> {
    
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
    let variant_ref: *mut EVT_VARIANT = unsafe_init_evt_variant(buffer_size as usize);
    let status = unsafe {
        EvtGetObjectArrayProperty(
            h_array.0,
            property_id.0 as u32,
            array_ix,
            0,
            buffer_size,
            Some(variant_ref),
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        return Err(win_error);
    }
    let variant = unsafe {&*variant_ref};
    let result = extract_value_from_evt_variant(variant);
    /*
    let result = match property_id {
        EvtPublisherMetadataChannelReferencePath => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(VariantBuffer::StringVal(new_string))
        },
        EvtPublisherMetadataLevelName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(VariantBuffer::StringVal(new_string))
        },
        EvtPublisherMetadataTaskName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(VariantBuffer::StringVal(new_string))
        },
        EvtPublisherMetadataOpcodeName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(VariantBuffer::StringVal(new_string))
        },
        EvtPublisherMetadataKeywordName => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            Ok(VariantBuffer::StringVal(new_string))
        },
        EvtPublisherMetadataTaskEventGuid => {
            let new_guid = unsafe { *variant.Anonymous.GuidVal };
            Ok(VariantBuffer::GuidVal(new_guid))
        },
        // Lazily adding a wild card. This should be updated if new EVT_PUBLISHER_METADATA_PROPERTY_IDs are supported
        _ => Ok(ManagedEvtVariant::from_variant(EVT_VARIANT::from(*variant)))
    }; */
    unsafe {libc::free(variant_ref as *mut libc::c_void)};
    Ok(result)
}
pub fn extract_value_from_evt_variant(variant: &EVT_VARIANT) -> VariantBuffer {
    let result = match variant.Type {
        0 => panic!("Null EvtVariant parsing not implemented"),
        1 => {
            let new_string = unsafe { variant.Anonymous.StringVal.to_string().unwrap() };
            VariantBuffer::StringVal(new_string)
        },
        2 => {
            let new_string = unsafe { variant.Anonymous.AnsiStringVal.to_string().unwrap() };
            VariantBuffer::StringVal(new_string)
        },
        3 => {
            let new_byte = unsafe { variant.Anonymous.SByteVal };
            VariantBuffer::SByteVal(new_byte)
        },
        4 => {
            let new_byte = unsafe { variant.Anonymous.ByteVal };
            VariantBuffer::ByteVal(new_byte)
        },
        5 => {
            let new_word = unsafe {variant.Anonymous.Int16Val};
            VariantBuffer::Int16Val(new_word)
        },
        6 => {
            let new_word = unsafe {variant.Anonymous.UInt16Val};
            VariantBuffer::UInt16Val(new_word)
        },
        7 => {
            let dword = unsafe {variant.Anonymous.Int32Val};
            VariantBuffer::Int32Val(dword)
        },
        8 => {
            let dword = unsafe {variant.Anonymous.UInt32Val};
            VariantBuffer::UInt32Val(dword)
        },
        9 => {
            let qword = unsafe {variant.Anonymous.Int64Val};
            VariantBuffer::Int64Val(qword)
        },
        10 => {
            let qword = unsafe {variant.Anonymous.UInt64Val};
            VariantBuffer::UInt64Val(qword)
        },
        11 => {
            let single = unsafe {variant.Anonymous.SingleVal};
            VariantBuffer::SingleVal(single)
        },
        12 => {
            let double = unsafe {variant.Anonymous.DoubleVal};
            VariantBuffer::DoubleVal(double)
        },
        13 => {
            let boolean = unsafe {variant.Anonymous.BooleanVal}.as_bool();
            VariantBuffer::BooleanVal(boolean)
        },
        14 => {
            panic!("Binary type not supported for EVT_VARIANT")
        },
        15 => {
            let my_128 = unsafe {*variant.Anonymous.GuidVal}.to_u128();
            VariantBuffer::GuidVal(GUID::from_u128(my_128))
        },
        32 => {
            let handle = unsafe {variant.Anonymous.EvtHandleVal};
            VariantBuffer::EvtHandleVal(handle)
        },
        other => panic!("Type {} not supported for EVT_VARIANT", other)
    };
    result
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
    let variant_ref: *mut EVT_VARIANT = unsafe_init_evt_variant(buffer_size as usize);
    let status = unsafe {
        EvtGetPublisherMetadataProperty(
            *h_provider,
            property_id,
            0,
            buffer_size,
            Some(variant_ref),
            &mut buffer_used,
        )
    };
    if !status.as_bool() {
        let win_error = Error::from_win32();
        //println!("EvtGetPublisherMetadataPropery Error 2. Skipping provider '{}': {}", &provider, win_error.message());
        return Err(win_error);
    }
    let variant = unsafe {&*variant_ref};
    let handle = unsafe {variant.Anonymous.EvtHandleVal};
    unsafe {libc::free(variant_ref as *mut libc::c_void)};
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

pub fn format_event_message(h_event: &EVT_HANDLE, h_publisher: &EVT_HANDLE, flag: EVT_FORMAT_MESSAGE_FLAGS, message_id: Option<&u32>) -> Result<String> {
    let msg_id: u32 = match message_id {
        None => 0,
        Some(id) => *id
    };

    let mut buffer_used: u32 = 0;
    let format_status = unsafe {
        EvtFormatMessage(
            *h_publisher,
            *h_event,
            msg_id,
            None,
            flag.0,
            None,
            &mut buffer_used,
        )
    };

    if !format_status.as_bool() {
        let win_error = Error::from_win32();
        if win_error.code() != ERROR_INSUFFICIENT_BUFFER.into(){
            //println!("Failed to get buffer size for EvtFormatMessage: {}", win_error.message());
            return Err(win_error);
        }
    }

    // Fill buffer
    let buffer_size: u32 = buffer_used;
    let mut message_buffer: Vec<u16> = vec![0; buffer_size.try_into().unwrap()];
    let format_status = unsafe {
        EvtFormatMessage(
            *h_publisher,
            *h_event,
            msg_id,
            None,
            flag.0,
            Some(&mut message_buffer),
            &mut buffer_used,
        )
    };
    let win_error = Error::from_win32();
    if format_status.as_bool() || win_error.code() == ERROR_EVT_UNRESOLVED_VALUE_INSERT.into() {
        let message = String::from_utf16_lossy(&message_buffer[..(buffer_used - 1).try_into().unwrap()]);
        Ok(message)
    } else {
        //message = format!("Failed to EvtFormatMessage: {}", win_error.message());
        println!("Failed to EvtFormatMessage: {}", win_error.message());
        Err(win_error)
    }

}

pub fn unsafe_init_evt_variant(n: usize) -> *mut EVT_VARIANT {
    // allocate n bytes
    let raw = unsafe { libc::calloc(n, mem::size_of::<u8>()) } as *mut EVT_VARIANT;
    if raw.is_null() {
        panic!("Failed to allocate memory");
    }
    raw
}