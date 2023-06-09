mod winevt;
mod provider;
mod events;
mod managed_variant;
use events::EvtEvent;
use provider::EvtProvider;
use winevt::*;
use managed_variant::*;

use clap::{Arg, Command};
use std::path::Path;
use std::ffi::OsString;
use windows::core::*;
use windows::Win32::Foundation::*;
use windows::Win32::System::EventLog::*;
use std::fs;
use xmltree::Element;
use std::iter::once;
use std::io::{BufWriter, Write, SeekFrom, Seek};
use std::collections::HashMap;
use std::collections::HashSet;
use std::os::windows::ffi::OsStrExt;
use std::io;
use regex::Regex;
use std::sync::mpsc::channel;
use std::thread;
use std::fs::File;


fn main() {
    let mut flags: EVT_QUERY_FLAGS = EvtQueryChannelPath;
    let channels_from_args: HashSet<String> = parse_cmdline_args(&mut flags).unwrap();
    let providers: HashSet<EvtProvider> = enumerate_publishers().unwrap();
    let tasks: HashMap<String, HashSet<String>> = divvy_tasks_from_providers(&providers, &channels_from_args);

    let (output_sender, output_receiver) = channel();
    let (error_sender, error_receiver) = channel();

    let mut handles = vec![];

    for (channel, providers) in tasks {
        for provider in providers {
            let output_sender = output_sender.clone();
            let error_sender = error_sender.clone();
            let provider = provider.clone();
            let channel = channel.clone();

            let handle = thread::spawn(move || {
                let query_str: String = format!("*[System[Provider[@Name='{}']]]", provider);
                let channel_vec: Vec<u16> = OsString::from(&channel).encode_wide().chain(once(0)).collect();
                let p_channel: PCWSTR = PCWSTR(channel_vec.as_ptr());
                let query_vec: Vec<u16> = OsString::from(&query_str).encode_wide().chain(once(0)).collect();
                let query_handle_result = unsafe {
                    EvtQuery(
                        None,
                        p_channel,
                        PCWSTR(query_vec.as_ptr()),
                        flags.0,
                    )
                };

                if let Ok(query_handle) = query_handle_result {
                    loop {
                        let mut next_buffer: [isize; 1] = [0; 1];
                        let mut returned: u32 = 0;
                        let next_status = unsafe { 
                            EvtNext(
                                query_handle, 
                                &mut next_buffer, 
                                0, 
                                0, 
                                &mut returned,
                            ) 
                        };
                        if !next_status.as_bool() {
                            let win_error = Error::from_win32();
                            if win_error.code() == ERROR_NO_MORE_ITEMS.into() {
                                break;
                            } else {
                                println!("Error pulling next event from query: {}", win_error.message());
                                continue;
                            }
                        }
                        let h_event  = EVT_HANDLE(next_buffer[0]);

                        // Get the provider name from the XML of the event
                        let evt: EvtEvent = match EvtEvent::new(&h_event) {
                            Ok(event) => event,
                            Err(_e) => {
                                println!("Problem with event. Skipping.");
                                continue;
                            }
                        };
                        output_sender.send((evt.get_record_id(), evt)).unwrap();

                        // Free resources allocated for the current event
                        unsafe { EvtClose(h_event) };
                    }
                    unsafe { EvtClose(query_handle) };
                } else if let Err(e) = query_handle_result {
                    println!("Couldn't open query for channel '{}' because error. Skipping: {}", &channel, e.message());
                }
            });

            handles.push(handle);
        }
    }

    // Wait for all threads to finish processing
    for handle in handles {
        handle.join().unwrap();
    }

    drop(output_sender);
    drop(error_sender);

    println!("done fetching");

    // Dump output to disk
    let output_path = Path::new("output.csv");
    let error_path = Path::new("error.txt");

    let mut output_file = File::create(&output_path).unwrap();
    let mut error_file = File::create(&error_path).unwrap();

    let mut events: Vec<EvtEvent> = output_receiver.iter().map(|(_id, event)| event.clone()).collect();
    events.sort_unstable_by(|a, b| {
        let time_a = a.get_timestamp();
        let time_b = b.get_timestamp();
        time_a.cmp(&time_b)
    });
    for event in events {
        write_to_csv(&mut output_file, &event).unwrap();
    }
    for error_msg in error_receiver {
        write_to_txt(&mut error_file, &error_msg).unwrap();
    }
}

fn divvy_tasks_from_providers(providers: &HashSet<EvtProvider>, channels_from_args: &HashSet<String>) -> HashMap<String, HashSet<String>> {
    let mut tasks: HashMap<String, HashSet<String>> = HashMap::new();
    // Loop through providers
    for provider in providers {
        // Loop through channels
        for channel in provider.get_channels() {
            // Add channel to key, provider to HashSet value
            tasks.entry(channel.to_string()).or_insert_with(HashSet::new).insert(provider.get_name().to_string());
        }
    }

    if !channels_from_args.is_empty() {
        let mut used_channels = HashMap::new();
        let file_name_pattern = Regex::new(r"Archive-([^-]+)-.*").unwrap();
        for channel_entry in channels_from_args {
            let mut file_key = String::new();
            let file_path = Path::new(&channel_entry);
            let file_name = file_path.with_extension("").file_name().unwrap().to_str().unwrap().replace("%4", "/");
            if file_name.starts_with("Archive-") {
                let caps = file_name_pattern.captures(&file_name).unwrap();
                file_key = caps.get(1).unwrap().as_str().to_string();
            } else {
                file_key = file_name;
            }
            if tasks.contains_key(&file_key){
                let value = tasks.remove(&file_key).unwrap();
                //add used channel to used_channels
                //used_channels.push(channel_entry.clone());
                used_channels.insert(channel_entry.to_string(), value);
            }
        }
        tasks = used_channels;
    }
    // Return HashMap
    tasks
}

fn parse_cmdline_args(flags: &mut EVT_QUERY_FLAGS) -> std::result::Result<HashSet<String>, Error> {
    let matches = Command::new("Event Log Parser")
        .version("1.0")
        .author("Adam Boretos")
        .about("Parses Windows event logs and outputs a CSV file")
        .arg(
            Arg::new("path")
                .short('p')
                .long("path")
                .help("Path to an .evtx file or a directory containing .evtx files")
        )
        .get_matches();
        
    let mut channel_results: HashSet<String> = HashSet::new();
    // Access the "path" value
    if let Some(untrimmed_path) = matches.get_one::<String>("path") {
        // Set flag to query file instead of local logs
        flags.0 = EvtQueryFilePath.0;

        
        let path = untrimmed_path.trim();
        let input_path = Path::new(path);
        if input_path.is_dir() {
            // Process all .evtx files within the directory
            let files = match fs::read_dir(input_path) {
                Ok(files_read_dir) => files_read_dir,
                Err(error) => panic!("Problem reading directory: {:?}", error)
            };
            for file_path in files.flatten() {
                    if let Some(extension) = file_path.path().extension().and_then(|e| e.to_str()) {
                        if extension == "evtx" {
                            let str_file_path = file_path.path().as_os_str().to_string_lossy().to_string();
                            // dbg!(&str_file_path);
                            channel_results.insert(str_file_path);
                        }
                    }
                }
        } else if input_path.is_file() {
            // Process the single .evtx file
            let _extension = Path::new(&input_path).extension().and_then(std::ffi::OsStr::to_str);
            if let Some(_extension) = Some("evtx"){
                let input_path_str = input_path.to_str();
                
                let str_input = match input_path_str{
                    //Some(input_path_str) => OsString::from(input_path_str).encode_wide().chain(once(0)).collect(),
                    Some(input_path_str) => input_path_str.to_owned(),
                    //Some(input_path_str) => str_to_vec(input_path_str),
                    None => panic!("input file ending in .evtx converted to None string???")
                };

                // dbg!("Adding {} to channel vec.", input_path_str);
                channel_results.insert(str_input);
            }
        } else {
            panic!("Invalid path provided. Please provide a valid .evtx file or directory path.");
        }
    }
    Ok(channel_results)
}

fn enumerate_publishers() -> std::result::Result<HashSet<EvtProvider>, Error> {
    // Make sure to close publisher_enum_handle before you leave this function.

    // Check if we were given a set of channels to work with
    // If so, they are expected to be from evtx files passed to the cmdline
    // We only need providers that use said channels
    // Otherwise, we need all local providers because we'll use all local channels later


    // Get provider enumerator
    let publisher_enum_handle = evt_open_publisher_enum().unwrap();
	
    let mut results = HashSet::new();

    // Loop through providers
    loop {
        match evt_next_publisher_id(&publisher_enum_handle) {
            Ok(provider_name) => {
                match EvtProvider::new(&provider_name) {
                    Ok(prv) => {
                        println!("{}", &prv.to_json().unwrap());
                        results.insert(prv);
                    },
                    Err(e) => {
                        println!("Couldn't make EvtProvider for {}: {}", &provider_name, e.message());
                        continue;
                    }
                };
            }
            Err(error) => {
                if error.code() == ERROR_NO_MORE_ITEMS.into() {
                    break;
                } else {
                    unsafe { EvtClose(publisher_enum_handle) };
                    panic!("Problem enumerating local providers: {}", error.message());
                }
            }
        }
    }
    
	unsafe { EvtClose(publisher_enum_handle) };
	Ok(results)

}

fn write_to_csv(file: &mut File, event: &EvtEvent) -> std::result::Result<(), Box<dyn std::error::Error>> {
    let root = Element::parse(event.get_xml().as_bytes())?;
    let system = root.get_child("System").ok_or("Missing 'System' Element")?;

    let mut headers = Vec::new();
    let mut rows = Vec::new();

    // Explicitly save values of certain children as attributes
    let provider_name = system.get_child("Provider").and_then(|e| e.attributes.get("Name")).unwrap_or(&"".to_string()).clone();
    let system_time = system.get_child("TimeCreated").and_then(|e| e.attributes.get("SystemTime")).unwrap_or(&"".to_string()).clone();
    let activity_id = system.get_child("Correlation").and_then(|e| e.attributes.get("ActivityID")).unwrap_or(&"".to_string()).clone();
    let process_id = system.get_child("Execution").and_then(|e| e.attributes.get("ProcessID")).unwrap_or(&"".to_string()).clone();

    for child in system.children.iter() {
        let name = child.as_element().unwrap().name.to_string();
        let value = if name == "Provider" {
            provider_name.to_string()
        } else if name == "TimeCreated" {
            system_time.to_string()
        } else if name == "Correlation" {
            activity_id.to_string()
        } else if name == "Execution" {
            process_id.to_string()
        } else { 
            child.as_element().unwrap().get_text().unwrap_or(std::borrow::Cow::Borrowed("")).to_string()
        };

        if !headers.contains(&name) {
            headers.push(name);
        }

        rows.push(value.to_string());
    }

    if !headers.contains(&"Message".to_string()) {
        headers.push("Message".to_string());
    }
    let msg_string = event.get_event_message();
    let msg_lines = msg_string.lines();
    for msg_line in msg_lines {
        rows.push(msg_line.to_string());
        break;
    }
    if !headers.contains(&"XML".to_string()) {
        headers.push("XML".to_string());
    }
    rows.push(event.get_xml().replace("\r\n", "|").replace("\n", "|"));

    let mut writer = BufWriter::new(file);
    if writer.seek(SeekFrom::End(0)).is_err() {
        return Err(Box::new(io::Error::new(io::ErrorKind::Other, "Failed to seek to end of file")));
    }
    if writer.stream_position().unwrap() == 0 {
        writeln!(writer, "{}", headers.join(","))?;
    }
    writeln!(writer, "{}", rows.join(","))?;

    Ok(())
}



fn write_to_txt(file: &mut File, event: &EvtEvent) -> std::result::Result<(), Box<dyn std::error::Error>> {
    let mut writer = BufWriter::new(file);
    if writer.seek(SeekFrom::End(0)).is_err() {
        return Err(Box::new(io::Error::new(io::ErrorKind::Other, "Failed to seek to end of file")));
    }
    writeln!(writer, "{}", event.get_xml())?;

    Ok(())
}
