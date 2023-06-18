
use std::collections::{HashSet,HashMap};
use windows::Win32::System::EventLog::*;
use windows::core::*;
use windows::Win32::Foundation::*;
use serde::{Serialize, Deserialize};
use crate::provider::*;

#[derive(Debug, Serialize, Deserialize)]
pub struct EvtEventMetadata {

}
impl EvtEventMetadata {
    pub fn from_event(h_event: &EVT_HANDLE) -> Self {
        let id = 
    }

    fn evt_get_event_metadata_property(h_provider: &EVT_HANDLE, property_id: EVT_EVENT_METADATA_PROPERTY_ID, provider: &EvtProvider) -> Result<EVT_HANDLE> {
        // Determine necessary size of buffer to hold array of events for provider.
        let mut buffer_used: u32 = 0;
        let status = unsafe {
            EvtGetEventMetadataProperty(
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
    
        // Properly size buffer and fill with array of provider's events.
        let buffer_size: u32 = buffer_used;
        let mut variant = EVT_VARIANT {
            Anonymous: EVT_VARIANT_0 {
                ByteVal: 0,
            },
            Count: 0,
            Type: EvtVarTypeByte.0 as u32,
        };
        let status = unsafe {
            EvtGetEventMetadataProperty(
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
        let result = match property_id {
            EventMetadataEventID => unsafe {variant.Anonymous.UInt32Val}, // u32
            EventMetadataEventVersion => unsafe {variant.Anonymous.UInt32Val}, // u32
            EventMetadataEventChannel => 
            EventMetadataEventTemplate => unsafe {variant.Anonymous.StringVal.to_string().unwrap()}, // String
            EventMetadataEventKeyword => {
                let keywords: u64 = unsafe {variant.Anonymous.UInt64Val};
                if (keywords & 0x00FFFFFFFFFFFFFF) > 0 {
                    let names: Vec<&String> = Vec::new();
                    let keywords_map = provider.get_keywords();
                    for (&key, value_map) in keywords_map {
                        if keywords & key > 0 {
                            if let Some(name) = value_map.get("Keyword Name") {
                                names.push(name);
                            }
                        }
                    }
                    names // Vec<&String>
                }
            },
            _ => unsafe {variant.Anonymous.UInt32Val}
        };
        //let handle = unsafe {variant.Anonymous.EvtHandleVal};
        //Ok(handle)
    }
}