use crate::managed_variant::ManagedEvtVariant;
use crate::winevt::*;

use windows::core::*;
use windows::Win32::Foundation::*;
use std::collections::{HashSet,HashMap};
use windows::Win32::System::EventLog::*;
use std::iter::once;
use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use std::hash::{Hash, Hasher};
use std::fmt;

#[derive(Debug)]
pub struct EvtProvider {
    name: String,
    handle: EVT_HANDLE,
    channels: HashSet<String>,
    levels: HashMap<u32, String>,
    tasks: HashMap<u32, String>,
    opcodes: HashMap<u32, String>,
    keywords: HashMap<u32, String>,
}
impl EvtProvider {
    pub fn new(name: &str) -> std::result::Result<Self, Error> {
        let provider_name = name.to_string();
        println!("{}:", &provider_name);
        let h_provider = match Self::open_handle(name) {
            Ok(handle) => handle,
            Err(e) => return Err(e)
        };
        println!("  Channels:");
        let provider_channels = match Self::enumerate_channels(name, &h_provider) {
            Ok(pcps) => pcps,
            Err(e) => return Err(e)
        };

        println!("  Levels:");
        let levels = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataLevels) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get levels for provider {}: {}", &provider_name, e.message());
                return Err(e);
            }
        };
        println!("  Tasks:");
        let tasks = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataTasks) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get tasks for provider {}: {}", &provider_name, e.message());
                return Err(e);
            }
        };
        println!("  Opcodes:");
        let opcodes = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataOpcodes) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get opcodes for provider {}: {}", &provider_name, e.message());
                return Err(e);
            }
        };
        println!("  Keywords:");
        let keywords = match Self::get_metadata_property(&h_provider, EvtPublisherMetadataKeywords) {
            Ok(results) => results,
            Err(e) => {
                println!("Couldn't get keywords for provider {}: {}", &provider_name, e.message());
                return Err(e);
            }
        };
        Ok(Self {
            name: provider_name,
            handle: h_provider,
            channels: provider_channels,
            levels: levels,
            tasks: tasks,
            opcodes: opcodes,
            keywords: keywords
        })
    }

    pub fn get_channels(&self) -> &HashSet<String>{
        &self.channels
    }

    pub fn get_name(&self) -> &str {
        &self.name
    }

    fn open_handle(name: &str) -> std::result::Result<EVT_HANDLE, Error> {
        let provider_pcwstr: Vec<u16> = OsString::from(&name).encode_wide().chain(once(0)).collect();
        match evt_open_publisher_metadata(provider_pcwstr, None) {
            Ok(handle) => return Ok(handle),
            Err(e) => return Err(e)  // This is fucking stupid. Handle your errors consistently. Caller/Callee
        };
    }

    fn get_metadata_property(h_provider: &EVT_HANDLE, property_flag: EVT_PUBLISHER_METADATA_PROPERTY_ID) -> Result<HashMap<u32, String>> {
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

        let mut property_results: HashMap<u32, String> = HashMap::new();

        for n in 0..property_array_size {
            println!("Getting item number {}", n);
            //let property_message = flag_map.get(&property_flag.0).unwrap();
            match property_flag {
                EvtPublisherMetadataLevels => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelName) {
                        Ok(managed_var) => {
                            println!("Level Name is: {}", managed_var.get_string().unwrap())
                        },
                        Err(e) => {
                            println!("Couldn't get level name: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelValue) {
                        Ok(managed_var) => {
                            println!("Level Value is: {}", managed_var.get_u32())
                        },
                        Err(e) => {
                            println!("Couldn't get level value: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataLevelMessageID) {
                        Ok(managed_var) => {
                            if managed_var.get_int32() == -1 {
                                println!("No level message");
                                continue;
                            }
                            match Self::format_event_message(
                                &EVT_HANDLE(0), 
                                h_provider, 
                                EvtFormatMessageId, 
                                Some(&managed_var.get_u32())
                            ) {
                                Ok(m) => {
                                    println!("Level message is: {}", m);
                                },
                                Err(e) => {
                                    println!("Failed to retrieve message: {}", e.message());
                                    continue;
                                }
                            };
                        },
                        Err(e) => {
                            println!("Couldn't get level message: {}", e.message())
                        }
                    }
                },
                EvtPublisherMetadataTasks => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskName) {
                        Ok(managed_var) => {
                            println!("Task Name is: {}", managed_var.get_string().unwrap())
                        },
                        Err(e) => {
                            println!("Couldn't get task name: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskValue) {
                        Ok(managed_var) => {
                            println!("Task Value is: {}", managed_var.get_u32())
                        },
                        Err(e) => {
                            println!("Couldn't get task value: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskEventGuid) {
                        Ok(managed) => {
                            // Task Guids are EvtVarTypeString, not EvtVarTypeGuid
                            let myguid = GuidWrapper(managed.get_guid().unwrap()).to_string();
                            if myguid != "00000000-0000-0000-0000-000000000000" {
                                println!("Task GUID is: {}", myguid)
                            }
                        },
                        Err(e) => {
                            println!("Couldn't get task GUID: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataTaskMessageID) {
                        Ok(managed) => {
                            if managed.get_int32() == -1 { // This might not work
                                println!("No task message");
                                continue;
                            }
                            
                            match Self::format_event_message(
                                &EVT_HANDLE(0), 
                                h_provider, 
                                EvtFormatMessageId, 
                                Some(&managed.get_u32())
                            ) {
                                Ok(m) => println!("Task message is: {}", m),
                                Err(e) => println!("Failed to retrieve message: {}", e.message())
                            };
                        },
                        Err(e) => {
                            println!("Couldn't get task message: {}", e.message());
                            continue;
                        }
                    }
                },
                EvtPublisherMetadataOpcodes => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeName) {
                        Ok(managed) => {
                            println!("Opcode Name is: {}", managed.get_string().unwrap())
                        },
                        Err(e) => {
                            println!("Couldn't get opcode name: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeValue){
                        Ok(managed) => {
                            println!("Opcode Value is: {}", managed.get_u32())
                        },
                        Err(e) => {
                            println!("Couldn't get opcode value: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataOpcodeMessageID){
                        Ok(managed) => {
                            if managed.get_int32() == -1 { // This might not work
                                println!("No opcode message");
                                continue;
                            }
                            match Self::format_event_message(
                                &EVT_HANDLE(0), 
                                h_provider, 
                                EvtFormatMessageId, 
                                Some(&managed.get_u32())
                            ) {
                                Ok(m) => println!("Opcode message is: {}", m),
                                Err(e) => println!("Failed to retrieve message: {}", e.message())
                            };
                        },
                        Err(e) => {
                            println!("Couldn't get opcode message: {}", e.message())
                        }
                    }
                },
                EvtPublisherMetadataKeywords => {
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordName){
                        Ok(managed) => {
                            println!("Keyword Name is: {}", managed.get_string().unwrap())
                        },
                        Err(e) => {
                            println!("Couldn't get keyword name: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordValue) {
                        Ok(managed) => {
                            println!("Keyword Value is: {}", managed.get_u64())
                        },
                        Err(e) => {
                            println!("Couldn't get keyword value: {}", e.message());
                            continue;
                        }
                    }
                    match get_property(&property_array_handle, n, EvtPublisherMetadataKeywordMessageID) {
                        Ok(managed) => {
                            if managed.get_int32() == -1 {
                                println!("No keyword message");
                                continue;
                            }
                            match Self::format_event_message(
                                &EVT_HANDLE(0), 
                                h_provider, 
                                EvtFormatMessageId, 
                                Some(&managed.get_u32())
                            ) {
                                Ok(m) => println!("Keyword message is: {}", m),
                                Err(e) => println!("Failed to retrieve message: {}", e.message())
                            };
                        },
                        Err(e) => {
                            println!("Couldn't get keyword message: {}", e.message())
                        }
                    }
                },
                _ => println!("Incompatible property")
            };

        }
        unsafe { EvtClose(property_array_handle) };
        Ok(property_results)
    }
    fn enumerate_channels(provider_name: &str, h_provider: &EVT_HANDLE) -> std::result::Result<HashSet<String>, Error> {
        // Get handle to array of channel names from provider
        let channels_array_handle = match evt_get_publisher_metadata_property(h_provider, EvtPublisherMetadataChannelReferences) {
            Ok(handle) => handle,
            Err(e) => {
                println!("EvtGetPublisherMetadataProperty Error with provider {}: {}", provider_name, e.message());
                return Err(e);
            }
        };
        
        // Get size of channel array
        let channel_array_size = match evt_get_object_array_size(&channels_array_handle) {
            Ok(size) => size,
            Err(e) => {
                println!("Couldn't determine number of channels in provider {}: {}", provider_name, e.message());
                return Err(e);
            }
        };
        //println!("{} channels in provider {}.", channel_array_size, provider);
        let mut channel_results: HashSet<String> = HashSet::new();
        // Loop through each channel in the array
        for n in 0..channel_array_size {

            // Get channel name
            match get_property(&channels_array_handle, n, EvtPublisherMetadataChannelReferencePath) {
                Ok(managed)=> {
                    // Add channel name to HashSet
                    let name = managed.get_string().unwrap();
                    println!("    {}", &name);
                    channel_results.insert(name);
                },
                Err(e) => {
                    println!("Couldn't get channel name from the array of provider {}. Skipping name: {}", provider_name, e.message());
                    continue;
                }
            };



        }
        Ok(channel_results)
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
                println!("Failed to get buffer size for EvtFormatMessage: {}", win_error.message());
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
    use std::collections::HashSet;
    #[test]
    fn test_provider_initialization() {
        let provider_name = "Microsoft-Windows-Security-Auditing";  // Make sure this provider exists in your test environment

        let provider_result = EvtProvider::new(provider_name);

        // Ensure the provider initialization was successful
        assert!(provider_result.is_ok(), "Expected Ok, got {:?}", provider_result);
        let provider = provider_result.unwrap();

        // Check that the provider name was correctly set
        assert_eq!(provider.name, provider_name);

        // Check that the handle is valid, if possible

        // Check that the correct channels were retrieved
        // This will depend on what channels your test provider has
    }

    #[test]
    fn test_provider_channel_initialization() {
        let provider_name = "Microsoft-Windows-Security-Auditing";  // Make sure this provider exists in your test environment
        let mut expected_channels: HashSet<String> = HashSet::new();
        expected_channels.insert("Security".to_string());
        let provider = EvtProvider::new(provider_name).unwrap();
        // Check that the provider name was correctly set
        assert_eq!(provider.get_channels(), &expected_channels);

        // Check that the handle is valid, if possible

        // Check that the correct channels were retrieved
        // This will depend on what channels your test provider has
    }

    // More tests here...
}
