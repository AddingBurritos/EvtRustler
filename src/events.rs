//use crate::winevt::*;

use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::EventLog::*;
use std::num::ParseIntError;
use std::iter::once;
use std::ffi::OsString;
use std::os::windows::ffi::OsStrExt;
use xmltree::Element;
use crate::winevt::evt_open_publisher_metadata;

#[derive(Debug, Clone)]
pub struct EvtEvent {
    handle: EVT_HANDLE,
    channel: String,
    provider: String,
    xml: String,
    time_written: String,
    record_id: u32,
    message: String
}
impl EvtEvent {
    pub fn new(h_event: &EVT_HANDLE) -> std::result::Result<Self, EvtError> {
        let xml = Self::format_event_message(h_event, &EVT_HANDLE(0), EvtFormatMessageXml)?;
        let element = Element::parse(xml.as_bytes()).unwrap();
        let empty = Element::new("Empty");
        let system_element = element.get_child("System").unwrap_or(&empty);
        let provider = system_element.get_child("Provider").unwrap_or(&empty).attributes.get("Name").unwrap_or(&String::new()).to_string();
        let channel = system_element.get_child("Channel").unwrap_or(&empty).get_text().unwrap_or(std::borrow::Cow::Borrowed("")).to_string();
        let time_written = system_element.get_child("TimeCreated").unwrap_or(&empty).attributes.get("SystemTime").unwrap_or(&"1970-01-01T00:00:00.0000000Z".to_string()).to_string();
        let record = system_element.get_child("EventRecordID").unwrap().get_text().unwrap_or(std::borrow::Cow::Borrowed("")).to_string();
        let record_id = record.parse::<u32>()?;
        let message = Self::generate_event_message(h_event, &provider);
        //println!("XML: {}", &xml);
        Ok(Self {
            handle: *h_event,
            channel: channel,
            provider: provider,
            xml: xml,
            time_written: time_written,
            record_id: record_id,
            message: message
        })
    }
    fn format_event_message(h_event: &EVT_HANDLE, h_publisher: &EVT_HANDLE, flag: EVT_FORMAT_MESSAGE_FLAGS) -> Result<String> {
        if flag == EvtFormatMessageId {
            println!("EvtFormatMessageId flag not supported in this function.")
        }

        let mut buffer_used: u32 = 0;
        let format_status = unsafe {
            EvtFormatMessage(
                *h_publisher,
                *h_event,
                0,
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
                0,
                None,
                flag.0,
                Some(&mut message_buffer),
                &mut buffer_used,
            )
        };
        let mut message: String = String::new();
        if format_status.as_bool() {
            message = String::from_utf16_lossy(&message_buffer[..buffer_used.try_into().unwrap()]);
            message = message.trim_end_matches('\0').to_string();
        } else {
            let win_error = Error::from_win32();
            message = format!("Failed to EvtFormatMessage: {}", win_error.message());
        }

        Ok(message)
    }

    pub fn get_event_message(&self) -> String {
        self.message.clone()
    }
    
    fn generate_event_message(handle: &EVT_HANDLE, provider: &String) -> String {
        let provider_vec = OsString::from(provider).encode_wide().chain(once(0)).collect();
        let h_publisher = match evt_open_publisher_metadata(provider_vec, None) {
            Ok(handle) => handle,
            Err(e) => {
                println!("Failed to open publisher metadata: {}", e.message());
                return format!("Failed to open publisher metadata: {}", e.message());
            }
        };
        let message = match Self::format_event_message(handle, &h_publisher, EvtFormatMessageEvent) {
            Ok(msg) => msg,
            Err(e) => {
                println!("Failed to get event message: {}", e.message());
                return format!("Failed to get event message: {}", e.message());
            }
        };
        unsafe { EvtClose(h_publisher) };
        //println!("{}", &message);
        message
    }
    pub fn get_timestamp(&self) -> String {
        self.time_written.clone()
    }
    pub fn get_xml(&self) -> String {
        self.xml.clone()
    }
    pub fn get_record_id (&self) -> u32 {
        self.record_id
    }
}

#[derive(Debug)]
pub enum EvtError {
    Win32Error(Error),
    ParseIntError(ParseIntError)
    // Add more as needed...
}

impl From<Error> for EvtError {
    fn from(err: Error) -> EvtError {
        EvtError::Win32Error(err)
    }
}

impl From<ParseIntError> for EvtError {
    fn from(err: ParseIntError) -> EvtError {
        EvtError::ParseIntError(err)
    }
}

impl std::error::Error for EvtError {}

impl std::fmt::Display for EvtError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            EvtError::Win32Error(err) => write!(f, "Win32 error: {}", err.message()),
            EvtError::ParseIntError(err) => write!(f, "ParseIntError error: {}", err.to_string()),
            // More as needed...
        }
    }
}