
use std::collections::{HashSet,HashMap};
use windows::Win32::System::EventLog::*;
use windows::core::*;
use windows::Win32::Foundation::*;
use serde::{Serialize, Deserialize};
use crate::provider::*;
use crate::winevt::*;

#[derive(Debug)]
pub enum EventPropertyTypes {
    u32_val(u32),
    string_val(String),
    string_vec(Vec<String>),
}
#[derive(Debug, Serialize, Deserialize)]
pub struct EvtEventMetadata {
    id: u32,
    version: u32,
    channel: String,
    level: String,
    opcode: String,
    task: String,
    keywords: Vec<String>,
    message: String,
    template: String
}
impl EvtEventMetadata {
    pub fn from_event(h_event: &EVT_HANDLE, provider: &EvtProvider) -> Self {
        let id = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventID) {
            Ok(num) => {
                match num {
                    EventPropertyTypes::u32_val(val) => val,
                    _ => panic!("Event ID not u32!")
                }
            },
            Err(e) => panic!("Couldn't get event ID: {}", e.message())
        };
        let version = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventVersion) {
            Ok(num) => {
                match num {
                    EventPropertyTypes::u32_val(val) => val,
                    _ => panic!("Version not u32!")
                }
            },
            Err(e) => panic!("Couldn't get version: {}", e.message())
        };
        let channel = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventChannel) {
            Ok(name) => {
                match name {
                    EventPropertyTypes::string_val(val) => val,
                    _ => panic!("Channel not string!")
                }
            },
            Err(e) => panic!("Couldn't get channel: {}", e.message())
        };
        let level = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventLevel) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::u32_val(val) => {
                        match provider.get_levels().get(&(val as u64)) {
                            Some(level_matching_val) => match level_matching_val.get("Level Name") {
                                Some(level_name) => level_name,
                                None => panic!("No level name for {}", val)
                            },
                            None => panic!("No level defined for {}", val)
                        }
                    },
                    _ => panic!("Level not u32!")
                }
            },
            Err(e) => panic!("Couldn't get level: {}", e.message())
        };
        let opcode = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventOpcode) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::u32_val(val) => {
                        match provider.get_opcodes().get(&(val as u64)) {
                            Some(opcode_matching_val) => match opcode_matching_val.get("Opcode Name") {
                                Some(opcode_name) => opcode_name,
                                None => panic!("No opcode name for {}", val)
                            },
                            None => match val {
                                0 => "Info",
                                _ => panic!("No opcode defined for {}", val)
                            }
                        }
                    },
                    _ => panic!("Opcode not u32!")
                }
            },
            Err(e) => panic!("Couldn't get opcode: {}", e.message())
        };
        let task = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventTask) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::u32_val(val) => {
                        match provider.get_tasks().get(&(val as u64)) {
                            Some(task_matching_val) => match task_matching_val.get("Task Name") {
                                Some(task_name) => task_name,
                                None => panic!("No task name for {}", val)
                            },
                            None => match val { 
                                0 => "None",
                                _ => panic!("{}: No task defined for {}", provider.get_name(), val)
                            }
                        }
                    },
                    _ => panic!("Task not u32!")
                }
            },
            Err(e) => panic!("Couldn't get task: {}", e.message())
        };
        let keywords = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventKeyword) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::string_vec(val) => val,
                    _ => panic!("Task not Vec<String>!")
                }
            },
            Err(e) => panic!("Couldn't get task: {}", e.message())
        };
        let message = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventMessageID) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::u32_val(val) => {
                        match format_event_message( &EVT_HANDLE(0), provider.get_handle(),EvtFormatMessageId, Some(&val)) {
                            Ok(message) => message,
                            Err(e) => panic!("Couldn't format message string! {}", e.message())
                        }
                    },
                    _ => panic!("Message ID not u32!")
                }
            },
            Err(e) => panic!("Couldn't get message: {}", e.message())
        };
        let template = match Self::evt_get_event_metadata_property(h_event, provider, EventMetadataEventTemplate) {
            Ok(text) => {
                match text {
                    EventPropertyTypes::string_val(val) => val,
                    _ => panic!("Task not String!")
                }
            },
            Err(e) => panic!("Couldn't get template: {}", e.message())
        };
        EvtEventMetadata { 
            id: id, 
            version: version, 
            channel: channel, 
            level: level.to_string(), 
            opcode: opcode.to_string(), 
            task: task.to_string(), 
            keywords: keywords, 
            message: message, 
            template: template,
        }
    }

    fn evt_get_event_metadata_property(h_event: &EVT_HANDLE, provider: &EvtProvider, property_id: EVT_EVENT_METADATA_PROPERTY_ID) -> Result<EventPropertyTypes> {
        // Determine necessary size of buffer to hold array of events for provider.
        let mut buffer_used: u32 = 0;
        let status = unsafe {
            EvtGetEventMetadataProperty(
                *h_event,
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
        let variant_ref: *mut EVT_VARIANT = unsafe_init_evt_variant(buffer_size as usize);
        let status = unsafe {
            EvtGetEventMetadataProperty(
                *h_event,
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
        let result: EventPropertyTypes = match property_id {
            EventMetadataEventID => EventPropertyTypes::u32_val(unsafe {variant.Anonymous.UInt32Val}), // u32
            EventMetadataEventVersion => EventPropertyTypes::u32_val(unsafe {variant.Anonymous.UInt32Val}), // u32
            EventMetadataEventChannel => {
                let val = unsafe {variant.Anonymous.UInt32Val} as u64;
                let channel = provider.get_channels().get(&val).unwrap();
                EventPropertyTypes::string_val(channel.get("Channel Name").unwrap().to_string())
            },
            EventMetadataEventTemplate => {
                let my_str = unsafe {variant.Anonymous.StringVal.to_string().unwrap()};
                println!("{}", &my_str);
                EventPropertyTypes::string_val(my_str)
            }, // String
            EventMetadataEventKeyword => {
                let keywords: u64 = unsafe {variant.Anonymous.UInt64Val};
                if (keywords & 0x00FFFFFFFFFFFFFF) > 0 {
                    let mut names: Vec<String> = Vec::new();
                    for (&key, value_map) in provider.get_keywords() {
                        if keywords & key > 0 {
                            if let Some(name) = value_map.get("Keyword Name") {
                                names.push(name.to_string());
                            }
                        }
                    }
                    EventPropertyTypes::string_vec(names) // Vec<&String>
                } else {
                    EventPropertyTypes::string_vec(Vec::new())
                }
            },
            _ => EventPropertyTypes::u32_val(unsafe {variant.Anonymous.UInt32Val})
        };
        //let handle = unsafe {variant.Anonymous.EvtHandleVal};
        unsafe {libc::free(variant_ref as *mut libc::c_void)};
        Ok(result)
    }
    fn evt_get_event_metadata_string(h_event: &EVT_HANDLE, provider: &EvtProvider, property_id: EVT_EVENT_METADATA_PROPERTY_ID) -> Result<String> {
        // Determine necessary size of buffer to hold array of events for provider.
        let mut buffer_used: u32 = 0;
        let status = unsafe {
            EvtGetEventMetadataProperty(
                *h_event,
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
        let variant_ref: *mut EVT_VARIANT = unsafe_init_evt_variant(buffer_size as usize);
        unsafe {
            EvtGetEventMetadataProperty(
                *h_event,
                property_id,
                0,
                buffer_size,
                Some(variant_ref),
                &mut buffer_used,
            )
        };
        let variant = unsafe {&*variant_ref};
        let my_str = unsafe {variant.Anonymous.StringVal.to_string().unwrap()};
        unsafe {libc::free(variant_ref as *mut libc::c_void)};
        println!("{}", &my_str);

        //let handle = unsafe {variant.Anonymous.EvtHandleVal};
        Ok(my_str)
    }
}