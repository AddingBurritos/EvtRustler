use crate::winevt::*;
use crate::event_meta::*;
use crate::managed_variant::*;
use windows::core::*;
use windows::Win32::Foundation::*;
use std::collections::{HashSet,HashMap};
use windows::Win32::System::EventLog::*;
use std::iter::once;
use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use std::hash::{Hash, Hasher};
use std::fmt;
use serde::{Serialize, Deserialize};
use std::fs::File;
use std::io::Write;
use windows::Win32::System::WindowsProgramming::{MAX_COMPUTERNAME_LENGTH, GetComputerNameW};


#[derive(Debug, Serialize, Deserialize)]
pub struct EvtProvider {
    name: String,
    hostname: String,
    #[serde(skip)]
    handle: EVT_HANDLE,
    channels: HashMap<u64, HashMap<String, String>>,
    levels: HashMap<u64, HashMap<String, String>>,
    tasks: HashMap<u64, HashMap<String, String>>,
    opcodes: HashMap<u64, HashMap<String, String>>,
    keywords: HashMap<u64, HashMap<String, String>>,
    events: Vec<EvtEventMetadata>,
}
impl EvtProvider {
    pub fn new(name: &str) -> std::result::Result<Self, Error> {
        let provider_name = name.to_string();
        //println!("{}:", &provider_name);
        let h_provider = match Self::open_handle(name) {
            Ok(handle) => handle,
            Err(e) => return Err(e)
        };
        let provider_channels = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataChannelReferences) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get channels for provider {}: {}", &provider_name, e.message());
                HashMap::new()
            }
        };
        //println!("  Levels:");
        let levels = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataLevels) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get levels for provider {}: {}", &provider_name, e.message());
                HashMap::new()
            }
        };
        //println!("  Tasks:");
        let tasks = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataTasks) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get tasks for provider {}: {}", &provider_name, e.message());
                HashMap::new()
            }
        };
        //println!("  Opcodes:");
        let opcodes = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataOpcodes) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get opcodes for provider {}: {}", &provider_name, e.message());
                HashMap::new()
            }
        };
        //println!("  Keywords:");
        let keywords = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataKeywords) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get keywords for provider {}: {}", &provider_name, e.message());
                HashMap::new()
            }
        };
        let mut temp_prv = Self {
            name: provider_name.clone(),
            hostname: Self::get_hostname(),
            handle: h_provider,
            channels: provider_channels,
            levels: levels,
            tasks: tasks,
            opcodes: opcodes,
            keywords: keywords,
            //events: events,
            events: Vec::new()
        };
        
        //println!("  Events:");
        let events = match Self::enumerate_events(&h_provider, &temp_prv) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get event metadata for provider {}: {}", &provider_name, e.message());
                vec![]
            }
        };
        temp_prv.update_events(events);
        Ok(
            temp_prv
        )
    }

    pub fn update_events(&mut self, new_events: Vec<EvtEventMetadata>) {
        self.events = new_events
    }
    pub fn get_handle(&self) -> &EVT_HANDLE {
        &self.handle
    }
    pub fn get_channels(&self) -> &HashMap<u64, HashMap<String, String>>{
        &self.channels
    }

    pub fn get_name(&self) -> &str {
        &self.name
    }

    pub fn get_levels(&self) -> &HashMap<u64, HashMap<String, String>> {
        &self.levels
    }

    pub fn get_tasks(&self) -> &HashMap<u64, HashMap<String, String>> {
        &self.tasks
    }

    pub fn get_opcodes(&self) -> &HashMap<u64, HashMap<String, String>> {
        &self.opcodes
    }

    pub fn get_keywords(&self) -> &HashMap<u64, HashMap<String, String>> {
        &self.keywords
    }

    fn get_hostname() -> String {
        let mut max_len: u32 = MAX_COMPUTERNAME_LENGTH + 1;
        let mut name_vec: Vec<u16> = vec![0; max_len as usize];
        let name_pwstr: PWSTR = PWSTR::from_raw(name_vec.as_mut_ptr());
        if unsafe { GetComputerNameW(name_pwstr, &mut max_len) }.as_bool() {
            return unsafe { name_pwstr.to_string().unwrap() };
        } {
            println!("Could not find hostname! {}", Error::from_win32().message());
            return "UNKNOWN_HOST".to_string();
        }
    }
    fn open_handle(name: &str) -> std::result::Result<EVT_HANDLE, Error> {
        let provider_pcwstr: Vec<u16> = OsString::from(&name).encode_wide().chain(once(0)).collect();
        match evt_open_publisher_metadata(provider_pcwstr, None) {
            Ok(handle) => return Ok(handle),
            Err(e) => return Err(e)  // This is fucking stupid. Handle your errors consistently. Caller/Callee
        };
    }
    pub fn to_json(&self) -> serde_json::Result<String> {
        serde_json::to_string(&self)
    }
    pub fn write_to_file(&self, path: &str) -> std::io::Result<()> {
        let json = match self.to_json() {
            Ok(json) => json,
            Err(e) => return Err(std::io::Error::new(std::io::ErrorKind::Other, e.to_string())),
        };
        
        let mut file = File::create(path)?;
        file.write_all(json.as_bytes())
    }

    fn get_metadata_property(h_provider: &EVT_HANDLE, property_flag: EVT_PUBLISHER_METADATA_PROPERTY_ID) -> Result<HashMap<u64, HashMap<String, String>>> {
        let property_array_handle = match evt_get_publisher_metadata_property(h_provider, property_flag) {
            Ok(handle) => handle,
            Err(e) => {
                println!("EvtGetPublisherMetadataProperty Error: {}", e.message());
                return Err(e);
            }
        };

        let property_array_size = match evt_get_object_array_size(&property_array_handle) {
            Ok(size) => size,
            Err(e) => {
                println!("Couldn't determine number of properties in property array: {}", e.message());
                unsafe { EvtClose(property_array_handle) };
                return Err(e);
            }
        };

        let mut property_results: HashMap<u64, HashMap<String, String>> = HashMap::new();

        for n in 0..property_array_size {
            let mut inner_map: HashMap<String, String> = HashMap::new();
            match property_flag {
                EvtPublisherMetadataChannelReferences => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataChannelReferencePath) {
                        Ok(variant_buff) => {
                            let channel_name = match variant_buff {
                                VariantBuffer::StringVal(val) => val,
                                _ => String::new()
                            };
                            match inner_map.insert("Channel Name".to_string(), channel_name) {
                                Some(thing) => println!("Channel name '{}' overwritten.", thing),
                                None => {}
                            };
                        },
                        Err(e) => {
                            println!("Couldn't get channel name: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataChannelReferenceMessageID) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    if num != 0xFFFFFFFF {
                                        match format_event_message(&EVT_HANDLE(0), h_provider, EvtFormatMessageId, Some(&num)) {
                                            Ok(message_string) => {
                                                match inner_map.insert("Channel Message".to_string(), message_string) {
                                                    Some(thing) => println!("Channel Message '{}' has been overwritten.", thing),
                                                    None => {}
                                                };
                                            },
                                            Err(e) => {
                                                println!("Failed to retrieve message: {}", e.message());
                                            }
                                        };
                                    }
                                },
                                _ => {}
                            };
                            
                        },
                        Err(e) => println!("Couldn't get channel message: {}", e.message())
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataChannelReferenceIndex) {
                        Ok(variant_buffer) => {
                            match variant_buffer {
                                VariantBuffer::UInt32Val(num) => {
                                    match inner_map.insert("Channel Index".to_string(), num.to_string()) {
                                        Some(thing) => println!("Channel index '{}' has been overwritten by '{}'", thing, num.to_string()),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }
                        },
                        Err(e) => println!("Couldn't get channel index: {}", e.message())
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataChannelReferenceFlags) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    let mut flag_string = "False".to_string();
                                    if num > 0 {
                                        flag_string = "True".to_string();
                                    }
                                    match inner_map.insert("Channel Imported".to_string(), flag_string) {
                                        Some(thing) => println!("Channel Imported flag '{}' has been overwritten.", thing),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }
                        },
                        Err(e) => println!("Couldn't get channel imported flag: {}", e.message())
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataChannelReferenceID) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    match property_results.insert(num as u64, inner_map) {
                                        Some(thing) => println!("Channel {} has overwritten \"{}\"", num, format!("{:?}", thing)),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get channel ID: {}", e.message());
                        }
                    }
                },
                EvtPublisherMetadataLevels => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelName) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::StringVal(level_name) => {
                                    match inner_map.insert("Level Name".to_string(), level_name) {
                                        Some(thing) => println!("Level name '{}' has been replaced.", thing),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get level name: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelMessageID) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    if num != 0xFFFFFFFF {
                                        match format_event_message(&EVT_HANDLE(0), h_provider, EvtFormatMessageId, Some(&num)) {
                                            Ok(message) => {
                                                match inner_map.insert("Level Message".to_string(), message) {
                                                    Some(thing) => println!("Level Message '{}' has been replaced.", thing),
                                                    None => {}
                                                };
                                            },
                                            Err(e) => {
                                                println!("Failed to retrieve message: {}", e.message());
                                            }
                                        };
                                    }
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get level message: {}", e.message())
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelValue) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    match property_results.insert(num as u64, inner_map) {
                                        Some(thing) => println!("Level {} overwritten by \"{}\"", num, format!("{:?}", thing)),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get level value: {}", e.message());
                        }
                    }
                },
                EvtPublisherMetadataTasks => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskName) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::StringVal(task_name) => {
                                    match inner_map.insert("Task Name".to_string(), task_name) {
                                        Some(thing) => println!("Task name '{}' has been overwritten.", thing),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get task name: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskEventGuid) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::GuidVal(guid) => {
                                    if guid != GUID::zeroed() {
                                        let guid_string = GuidWrapper(guid).to_string();
                                        match inner_map.insert("Task GUID".to_string(), guid_string) {
                                            Some(thing) => println!("Task GUID '{}' has been overwritten.", thing),
                                            None => {}
                                        }
                                    }
                                },
                                VariantBuffer::StringVal(guid_string) => {
                                    if guid_string != GuidWrapper(GUID::zeroed()).to_string() {
                                        match inner_map.insert("Task GUID".to_string(), guid_string) {
                                            Some(thing) => println!("Task GUID '{}' has been overwritten.", thing),
                                            None => {}
                                        }
                                    }
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get task GUID: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskMessageID) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    if num != 0xFFFFFFFF {
                                        match format_event_message(&EVT_HANDLE(0), h_provider, EvtFormatMessageId, Some(&num)) {
                                            Ok(message) => {
                                                match inner_map.insert("Task Message".to_string(), message) {
                                                    Some(thing) => println!("Task Message '{}' has been overwritten", thing),
                                                    None => {}
                                                };
                                            },
                                            Err(e) => println!("Failed to retrieve message: {}", e.message())
                                        };
                                    }
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get task message: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskValue) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    match property_results.insert(num as u64, inner_map) {
                                        Some(thing) => println!("Task {} has overwritten \"{}\"", num, format!("{:?}", thing)),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get task value: {}", e.message());
                        }
                    }
                },
                EvtPublisherMetadataOpcodes => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeName) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::StringVal(opcode_name) => {
                                    match inner_map.insert("Opcode Name".to_string(), opcode_name) {
                                        Some(thing) => println!("Opcode name '{}' has been overwritten.", thing),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get opcode name: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeMessageID){
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    if num != 0xFFFFFFFF {
                                        match format_event_message(&EVT_HANDLE(0), h_provider, EvtFormatMessageId, Some(&num)) {
                                            Ok(message) => {
                                                match inner_map.insert("Opcode Message".to_string(), message) {
                                                    Some(thing) => println!("Opcode Message '{}' has been overwritten.", thing),
                                                    None => {}
                                                };
                                            },
                                            Err(e) => println!("Failed to retrieve message: {}", e.message())
                                        };
                                    }
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get opcode message: {}", e.message())
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeValue){
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    match property_results.insert(num as u64, inner_map) {
                                        Some(thing) => println!("Opcode {} has overwritten \"{}\"", num, format!("{:?}", thing)),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get opcode value: {}", e.message());
                        }
                    }
                },
                EvtPublisherMetadataKeywords => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordName){
                        Ok(managed) => {
                            match managed {
                                VariantBuffer::StringVal(keyword_name) => {
                                    match inner_map.insert("Keyword Name".to_string(), keyword_name) {
                                        Some(thing) => println!("Keyword name '{}' has been replaced.", thing),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get keyword name: {}", e.message());
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordMessageID) {
                        Ok(variant_buff) => {
                            match variant_buff {
                                VariantBuffer::UInt32Val(num) => {
                                    if num != 0xFFFFFFFF {
                                        match format_event_message(&EVT_HANDLE(0), h_provider, EvtFormatMessageId, Some(&num)) {
                                            Ok(message) => {
                                                match inner_map.insert("Keyword Message".to_string(), message) {
                                                    Some(thing) => println!("Keyword Message '{}' has been replaced.", thing),
                                                    None => {}
                                                };
                                            },
                                            Err(e) => println!("Failed to retrieve message: {}", e.message())
                                        };
                                    }
                                },
                                _ => {}
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get keyword message: {}", e.message())
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordValue) {
                        Ok(managed_var) => {
                            match managed_var {
                                VariantBuffer::UInt64Val(num) => {
                                    match property_results.insert(num, inner_map) {
                                        Some(thing) => println!("Keyword {} has overwritten \"{}\"", num, format!("{:?}", thing)),
                                        None => {}
                                    };
                                },
                                _ => {}
                            }

                        },
                        Err(e) => {
                            println!("Couldn't get keyword value: {}", e.message());
                            continue;
                        }
                    }
                },
                _ => println!("Incompatible property")
            };
        }
        
        unsafe { EvtClose(property_array_handle) };
        Ok(property_results)
    }

    fn enumerate_events(h_publisher: &EVT_HANDLE, provider: &EvtProvider) -> Result<Vec<EvtEventMetadata>> {
        let h_events = match unsafe { EvtOpenEventMetadataEnum(*h_publisher, 0) } {
            Ok(result) => result,
            Err(e) => {
                let win_error = Error::from_win32();
                println!("Couldn't enumerate events: {}", win_error.message());
                return Ok(Vec::new());
            }
        };
        let mut events: Vec<EvtEventMetadata> = Vec::new();
        loop {
            let h_event = match unsafe { EvtNextEventMetadata(h_events, 0) } {
                Ok(result) => {
                    match result.0 {
                        0 => {
                            let win_error = Error::from_win32();
                            if win_error.code() == ERROR_NO_MORE_ITEMS.into() {
                                break;
                            }
                            println!("Skipping event metadata because of null handle: {}", win_error.message());
                            continue;
                        },
                        _ => {} // Handle was filled successfully.
                    }
                    result
                },
                Err(e) => {
                    if e.code() == ERROR_NO_MORE_ITEMS.into() {
                        break;
                    } else {
                        panic!("Error getting next event metadata: {}", e.message())
                    }
                }
            };
            
            let event: EvtEventMetadata = EvtEventMetadata::from_event(&h_event, provider);
            events.push(event);
        }
        Ok(events)
    }
}
impl Drop for EvtProvider {
    fn drop(&mut self) {
        unsafe { EvtClose(self.handle) };
    }
}
impl PartialEq for EvtProvider {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
    }
}

impl Eq for EvtProvider {}

impl Hash for EvtProvider {
    fn hash<H: Hasher>(&self, state: &mut H) {
        self.name.hash(state);
    }
}
    fn format_event_message(h_event: &EVT_HANDLE, h_publisher: &EVT_HANDLE, flag: EVT_FORMAT_MESSAGE_FLAGS, message_id: Option<&u32>) -> Result<String> {
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
        if format_status.as_bool() {
            let message = String::from_utf16_lossy(&message_buffer[..(buffer_used - 1).try_into().unwrap()]);
            Ok(message)
        } else {
            let win_error = Error::from_win32();
            //message = format!("Failed to EvtFormatMessage: {}", win_error.message());
            println!("Failed to EvtFormatMessage: {}", win_error.message());
            Err(win_error)
        }

    }
#[repr(transparent)] // Ensure it has the same layout as the original type
pub struct GuidWrapper(pub GUID);

impl fmt::Display for GuidWrapper {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        let guid = &self.0;
        write!(f, "{:08x}-{:04x}-{:04x}-{:02x}{:02x}-{:02x}{:02x}{:02x}{:02x}{:02x}{:02x}",
            guid.data1,
            guid.data2,
            guid.data3,
            guid.data4[0], guid.data4[1],
            guid.data4[2], guid.data4[3], 
            guid.data4[4], guid.data4[5], 
            guid.data4[6], guid.data4[7]
        )
    }
}

#[cfg(test)]
mod tests {
    use crate::provider::EvtProvider;
    use std::collections::{HashSet, HashMap};
    #[test]
    fn test_provider_initialization() {
        let provider_name = "Microsoft-Windows-Security-Auditing";  // Make sure this provider exists in your test environment

        let provider_result = EvtProvider::new(provider_name);

        // Ensure the provider initialization was successful
        assert!(provider_result.is_ok(), "Expected Ok, got {:?}", provider_result);
        let provider = provider_result.unwrap();

        // Check that the provider name was correctly set
        assert_eq!(provider.name, provider_name);
    }

    #[test]
    fn test_provider_channel_initialization() {
        let provider_name = "Microsoft-Windows-Security-Auditing";  // Make sure this provider exists in your test environment
        let mut expected_channels: HashMap<String, String> = HashMap::new();
        expected_channels.insert("Channel Name".to_string(), "Security".to_string());
        expected_channels.insert("Channel Message".to_string(), "Security".to_string());
        expected_channels.insert("Channel Flag".to_string(), "Channel is imported".to_string());
        expected_channels.insert("Channel Index".to_string(), "0".to_string());
        let mut outer = HashMap::new();
        outer.insert(10, expected_channels);
        let provider = EvtProvider::new(provider_name).unwrap();
        // Check that the channel data was correctly pulled
        assert_eq!(provider.get_channels(), &outer);

    }

    // More tests here...
}
