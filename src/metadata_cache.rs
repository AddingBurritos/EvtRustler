use std::collections::HashMap;
use crate::provider::EvtProvider;
use std::fs::File;
use std::io::BufReader;

pub struct EvtCache {
    path: String,
    data: HashMap<String, EvtProvider>,
}

impl EvtCache {
    pub fn new(path: &str) -> std::io::Result<Self> {
        let data: HashMap<String, EvtProvider>;
    
        match File::open(path) {
            Ok(file) => {
                let reader = BufReader::new(file);
                data = match serde_json::from_reader(reader) {
                    Ok(data) => data,
                    Err(_) => return Err(std::io::Error::new(std::io::ErrorKind::Other, "Invalid cache data")),
                };
            }
            Err(ref error) if error.kind() == std::io::ErrorKind::NotFound => {
                // Create a new file if it does not exist
                File::create(path)?;
                data = HashMap::new(); // Initialize with an empty data set
            }
            Err(error) => return Err(error),
        }
    
        Ok(Self {
            path: path.to_string(),
            data: data,
        })
    }
    

    pub fn add_provider(&mut self, provider: EvtProvider) {
        let provider_name = provider.get_name();

        if !self.data.contains_key(provider_name) {
            // Add the new provider
            self.data.insert(provider_name.to_string(), provider);
            
        }
    }

    pub fn get_provider(&self, name: &str) -> Option<&EvtProvider> {
        self.data.get(name)
    }

    pub fn remove_provider(&mut self, name: &str) {
        self.data.remove(name);
    }

    pub fn provider_exists(&self, name: &str) -> bool {
        self.data.contains_key(name)
    }

    pub fn get_all_providers(&self) -> Vec<&String> {
        self.data.keys().collect()
    }

    pub fn get_data(&self) -> &HashMap<String, EvtProvider> {
        &self.data
    }

    pub fn save(&self) -> std::io::Result<()> {
        let file = File::create(&self.path)?;
        serde_json::to_writer(file, &self.data)?;
        Ok(())
    }
}