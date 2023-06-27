
use windows::core::*;
use windows::Win32::System::EventLog::*;

#[derive(Debug)]
pub enum VariantBuffer {
    StringVal(String),
    SByteVal(i8),
    ByteVal(u8),
    Int16Val(i16),
    UInt16Val(u16),
    Int32Val(i32),
    UInt32Val(u32),
    Int64Val(i64),
    UInt64Val(u64),
    SingleVal(f32),
    DoubleVal(f64),
    BooleanVal(bool),
    GuidVal(GUID),
    EvtHandleVal(EVT_HANDLE),
    // Other types omitted due to complexity, fill in as necessary
}

pub struct ManagedEvtVariant {
    buffer: VariantBuffer,
    variant: EVT_VARIANT,
}

impl ManagedEvtVariant {
    pub fn new() -> Self {
        let buffer = 0;
        let variant = EVT_VARIANT {
            Anonymous: EVT_VARIANT_0 {
                ByteVal: 0,
            },
            Count: 0,
            Type: EvtVarTypeByte.0 as u32,
        };

        ManagedEvtVariant {
            buffer: VariantBuffer::ByteVal(buffer),
            variant,
        }
    }

    pub fn update_variant(&mut self, new_variant: EVT_VARIANT) {
        self.buffer = match new_variant.Type {
            1 => { // EvtVarTypeString
                panic!("You shouldn't be adding a string val to the ManagedEvtVariant like this. Bad programmer.");
            },
            4 => { // EvtVarTypeByte
                let new_byte = unsafe { new_variant.Anonymous.ByteVal };
                self.variant.Anonymous.ByteVal = new_byte;
                VariantBuffer::ByteVal(new_byte)
            },
            7 => { // EvtVarTypeInt32
                let new_int32 = unsafe { new_variant.Anonymous.Int32Val };
                self.variant.Anonymous.Int32Val = new_int32;
                VariantBuffer::Int32Val(new_int32)
            },
            8 => { // EvtVarTypeUInt32
                let new_uint32 = unsafe { new_variant.Anonymous.UInt32Val };
                self.variant.Anonymous.UInt32Val = new_uint32;
                VariantBuffer::UInt32Val(new_uint32)
            },
            10 => { // EvtVarTypeUInt64
                let new_uint64 = unsafe { new_variant.Anonymous.UInt64Val };
                self.variant.Anonymous.UInt64Val = new_uint64;
                VariantBuffer::UInt64Val(new_uint64)
            }
            15 => { // EvtVarTypeGuid
                panic!("You shouldn't be adding a GUID val to the ManagedEvtVariant like this. Bad programmer.");
                //let buffer: &[GUID] = unsafe {std::slice::from_raw_parts(new_variant.Anonymous.GuidVal, 128)};
                //let new_buffer: Vec<GUID> = buffer.to_vec();
                //let new_guid: GUID = new_buffer[0];
                //VariantBuffer::GuidVal(new_guid)
            },
            32 => { // EvtVarTypeEvtHandle
                let new_handle = unsafe { new_variant.Anonymous.EvtHandleVal };
                self.variant.Anonymous.EvtHandleVal = new_handle;
                VariantBuffer::EvtHandleVal(new_handle)
            },
            num => panic!("Unsupported variant type: {}", num),
        };
        self.variant.Type = new_variant.Type;
        self.variant.Count = new_variant.Count;
    }
    pub fn from_variant(source_variant: &EVT_VARIANT) -> Self {
        let new_variant = match source_variant.Type {
            1 => panic!("You shouldn't be adding a string val to the ManagedEvtVariant like this. Use from_string."),
            4 => ManagedEvtVariant {
                variant: EVT_VARIANT {
                    Anonymous: EVT_VARIANT_0 {
                        ByteVal: unsafe {source_variant.Anonymous.ByteVal}
                    },
                    Count: 0,
                    Type: EvtVarTypeByte.0 as u32
                },
                buffer: VariantBuffer::ByteVal(unsafe {source_variant.Anonymous.ByteVal})
            },
            7 => ManagedEvtVariant {
                variant: EVT_VARIANT {
                    Anonymous: EVT_VARIANT_0 {
                        Int32Val: unsafe {source_variant.Anonymous.Int32Val}
                    },
                    Count: 0,
                    Type: EvtVarTypeInt32.0 as u32
                },
                buffer: VariantBuffer::Int32Val(unsafe {source_variant.Anonymous.Int32Val})
            },
            8 => ManagedEvtVariant {
                variant: EVT_VARIANT {
                    Anonymous: EVT_VARIANT_0 {
                        UInt32Val: unsafe {source_variant.Anonymous.UInt32Val}
                    },
                    Count: 0,
                    Type: EvtVarTypeUInt32.0 as u32
                },
                buffer: VariantBuffer::UInt32Val(unsafe {source_variant.Anonymous.UInt32Val})
            },
            10 => ManagedEvtVariant {
                variant: EVT_VARIANT {
                    Anonymous: EVT_VARIANT_0 {
                        UInt64Val: unsafe {source_variant.Anonymous.UInt64Val}
                    },
                    Count: 0,
                    Type: EvtVarTypeUInt64.0 as u32
                },
                buffer: VariantBuffer::UInt64Val(unsafe {source_variant.Anonymous.UInt64Val})
            },
            15 => {
                panic!("You shouldn't be adding a guid val to the ManagedEvtVariant like this. Use from_guid or from_guid_128.")
            },
            32 => ManagedEvtVariant {
                variant: EVT_VARIANT {
                    Anonymous: EVT_VARIANT_0 {
                        EvtHandleVal: unsafe {source_variant.Anonymous.EvtHandleVal}
                    },
                    Count: 0,
                    Type: EvtVarTypeEvtHandle.0 as u32
                },
                buffer: VariantBuffer::EvtHandleVal(unsafe {source_variant.Anonymous.EvtHandleVal})
            },
            num => panic!("Unsupported variant type: {}", num)
        };
        return new_variant;
        //let mut temp_var = ManagedEvtVariant::new();
        //temp_var.update_variant(source_variant);
        //temp_var
    }
    pub fn from_string(source_variant: String) -> Self {
        let new_var = ManagedEvtVariant {
            variant: EVT_VARIANT {
                Anonymous: EVT_VARIANT_0 {
                    StringVal: PCWSTR::null(),
                },
                Count: 0,
                Type: EvtVarTypeString.0 as u32,
            },
            buffer: VariantBuffer::StringVal(source_variant)
        };
        new_var
    }
    pub fn from_guid_128(guid_128: u128) -> Self {
        let new_var = ManagedEvtVariant {
            variant: EVT_VARIANT {
                Anonymous: EVT_VARIANT_0 {
                    StringVal: PCWSTR::null() // Surely this doesn't matter, right?
                },
                Count: 0,
                Type: EvtVarTypeGuid.0 as u32,
            },
            buffer: VariantBuffer::GuidVal(GUID::from_u128(guid_128))
        };
        new_var
    }
    pub fn from_guid(my_guid: GUID) -> Self {
        let new_var = ManagedEvtVariant {
            variant: EVT_VARIANT {
                Anonymous: EVT_VARIANT_0 {
                    StringVal: PCWSTR::null() // Surely this doesn't matter, right?
                },
                Count: 0,
                Type: EvtVarTypeGuid.0 as u32,
            },
            buffer: VariantBuffer::GuidVal(my_guid)
        };
        new_var
    }
    pub fn get_variant(self) -> EVT_VARIANT {
        // This function consumes itself to avoid the buffer and variant from desyncing
        self.variant
    }
    pub fn get_data(&self) -> Option<VariantBuffer> {
        // This function returns the data extracted from the variant buffer
        match &self.buffer {
            VariantBuffer::StringVal(buf) => Some(VariantBuffer::StringVal(buf.clone())),
            VariantBuffer::ByteVal(val) => Some(VariantBuffer::ByteVal(*val)),
            VariantBuffer::Int32Val(val) => Some(VariantBuffer::Int32Val(*val)),
            VariantBuffer::UInt32Val(val) => Some(VariantBuffer::UInt32Val(*val)),
            VariantBuffer::UInt64Val(val) => Some(VariantBuffer::UInt64Val(*val)),
            VariantBuffer::GuidVal(buf) => Some(VariantBuffer::GuidVal(buf.clone())),
            VariantBuffer::EvtHandleVal(buf) => Some(VariantBuffer::EvtHandleVal(buf.clone())),
            _ => None
        }
    }
    pub fn get_string(&self) -> Option<String> {
        // This function returns the String extracted from the variant buffer
        match &self.buffer {
            VariantBuffer::StringVal(buf) => Some(buf.to_string()),
            _ => None,
        }
    }
    pub fn get_byte(&self) -> u8 {
        // This function returns the byte extracted from the variant buffer
        unsafe {self.variant.Anonymous.ByteVal}
    }
    pub fn get_int32(&self) -> i32 {
        // This function returns the i32 extracted from the variant buffer
        unsafe {self.variant.Anonymous.Int32Val}
    }
    pub fn get_u32(&self) -> u32 {
        // This function returns the u32 extracted from the variant buffer
        unsafe {self.variant.Anonymous.UInt32Val}
    }
    pub fn get_u64(&self) -> u64 {
        unsafe {self.variant.Anonymous.UInt64Val}
    }
    pub fn get_guid(&self) -> Option<GUID> {
        // This function returns the i32 extracted from the variant buffer
        match &self.buffer {
            VariantBuffer::GuidVal(buf) => Some(buf.clone()),
            _ => None,
        }
    }
    pub fn get_evt_handle(&self) -> EVT_HANDLE {
        // This function returns the u32 extracted from the variant buffer
        unsafe {self.variant.Anonymous.EvtHandleVal}
    }
    

}
