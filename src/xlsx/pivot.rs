//! Pivot table parsing for XLSX files

use quick_xml::{events::Event, name::QName};
use std::io::{Read, Seek};

use crate::pivot::{
    PivotCache, PivotCacheField, PivotField, PivotFieldDataType, PivotFieldType, PivotItem,
    PivotSourceType, PivotTable, PivotTableInfo,
};
use crate::{Data, Reader, XlsxError};

use super::{get_attribute, get_dimension, xml_reader, XlReader};

impl<RS: Read + Seek> super::Xlsx<RS> {
    /// Load pivot tables metadata (without loading cache data)
    pub fn load_pivot_tables(&mut self) -> Result<(), XlsxError> {
        self.read_pivot_tables()
    }

    /// Get all pivot table names
    pub fn pivot_table_names(&self) -> Vec<&str> {
        self.pivot_tables.table_names()
    }

    /// Get pivot table by name
    pub fn pivot_table_by_name(&mut self, name: &str) -> Result<PivotTable, XlsxError> {
        // Find the pivot table info
        let info = self
            .pivot_tables
            .tables_by_sheet
            .values()
            .flat_map(|tables| tables.iter())
            .find(|t| t.name == name)
            .ok_or_else(|| XlsxError::Unexpected("Pivot table not found"))?
            .clone();

        // Find the sheet name
        let sheet_name = self
            .pivot_tables
            .tables_by_sheet
            .iter()
            .find(|(_, tables)| tables.iter().any(|t| t.name == name))
            .map(|(sheet, _)| sheet.clone())
            .ok_or_else(|| XlsxError::Unexpected("Pivot table sheet not found"))?;

        // Parse the full pivot table
        self.parse_pivot_table(&sheet_name, &info.path)
    }

    /// Get pivot cache by ID with its records
    pub fn pivot_cache_with_records(&mut self, cache_id: u32) -> Result<PivotCache, XlsxError> {
        // First get the cache metadata
        let mut cache = self
            .pivot_tables
            .caches
            .get(&cache_id)
            .ok_or_else(|| XlsxError::Unexpected("Pivot cache not found"))?
            .clone();

        // If cache already has records, return it
        if cache.records.is_some() {
            return Ok(cache);
        }

        // Find and parse the pivot cache records file
        // Use the cache path to determine the records file path
        if let Some(cache_def_path) = &cache.cache_path {
            // Extract the number from the cache definition path
            if let Some(cache_num) = cache_def_path
                .split("pivotCacheDefinition")
                .nth(1)
                .and_then(|s| s.trim_end_matches(".xml").parse::<u32>().ok())
            {
                let records_path = format!("xl/pivotCache/pivotCacheRecords{}.xml", cache_num);
                if let Some(Ok(mut reader)) = xml_reader(&mut self.zip, &records_path) {
                    let records = parse_pivot_cache_records(&mut reader, &cache.fields)?;
                    cache.records = Some(records);
                }
            }
        }

        // Update the cache in the collection
        self.pivot_tables.caches.insert(cache_id, cache.clone());

        Ok(cache)
    }

    /// Read pivot table metadata
    pub(crate) fn read_pivot_tables(&mut self) -> Result<(), XlsxError> {
        // First, discover pivot cache definitions
        self.discover_pivot_caches()?;

        // Then, discover pivot tables in each sheet
        let sheet_names = self.sheet_names();
        for sheet_name in sheet_names {
            self.discover_sheet_pivot_tables(&sheet_name)?;
        }

        Ok(())
    }

    /// Discover pivot cache definitions from workbook relationships
    fn discover_pivot_caches(&mut self) -> Result<(), XlsxError> {
        let mut cache_paths = Vec::new();

        // Read workbook relationships
        if let Some(Ok(mut reader)) = xml_reader(&mut self.zip, "xl/_rels/workbook.xml.rels") {
            let mut buf = Vec::new();
            loop {
                buf.clear();
                match reader.read_event_into(&mut buf) {
                    Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                        let mut is_cache = false;
                        let mut target = None;

                        if let Some(type_attr) = get_attribute(e.attributes(), QName(b"Type"))? {
                            let type_str = std::str::from_utf8(type_attr)
                                .map_err(|_| XlsxError::Unexpected("Invalid UTF-8"))?;
                            if type_str.contains("pivotCacheDefinition") {
                                is_cache = true;
                            }
                        }

                        if let Some(target_attr) = get_attribute(e.attributes(), QName(b"Target"))?
                        {
                            target = Some(
                                std::str::from_utf8(target_attr)
                                    .map_err(|_| XlsxError::Unexpected("Invalid UTF-8"))?
                                    .to_owned(),
                            );
                        }

                        if is_cache && target.is_some() {
                            cache_paths.push(format!("xl/{}", target.unwrap()));
                        }
                    }
                    Ok(Event::Eof) => break,
                    Err(e) => return Err(XlsxError::Xml(e)),
                    _ => {}
                }
            }
        }

        // Parse each cache definition and temporarily store with sequential IDs
        // The actual cache ID will be determined by the pivot tables that reference them
        for (idx, path) in cache_paths.iter().enumerate() {
            if let Some(Ok(mut reader)) = xml_reader(&mut self.zip, path) {
                let mut cache = parse_pivot_cache_metadata(&mut reader, idx as u32)?;
                // Store the path for later use when loading records
                cache.cache_path = Some(path.clone());
                // Store the cache - we'll update the ID when we find the pivot table
                self.pivot_tables.add_cache(cache);
            }
        }

        Ok(())
    }

    /// Discover pivot tables in a sheet
    fn discover_sheet_pivot_tables(&mut self, sheet_name: &str) -> Result<(), XlsxError> {
        let sheet_path = self
            .sheets
            .iter()
            .find(|(name, _)| name == sheet_name)
            .map(|(_, path)| path.clone());

        if let Some(sheet_path) = sheet_path {
            let rels_path = sheet_path
                .replace("worksheets/", "worksheets/_rels/")
                .replace(".xml", ".xml.rels");

            let mut pivot_paths = Vec::new();

            // Read sheet relationships
            if let Some(Ok(mut reader)) = xml_reader(&mut self.zip, &rels_path) {
                let mut buf = Vec::new();
                loop {
                    buf.clear();
                    match reader.read_event_into(&mut buf) {
                        Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"Relationship" => {
                            let mut is_pivot = false;
                            let mut target = None;

                            if let Some(type_attr) = get_attribute(e.attributes(), QName(b"Type"))?
                            {
                                let type_str = std::str::from_utf8(type_attr)
                                    .map_err(|_| XlsxError::Unexpected("Invalid UTF-8"))?;
                                if type_str.contains("pivotTable") {
                                    is_pivot = true;
                                }
                            }

                            if let Some(target_attr) =
                                get_attribute(e.attributes(), QName(b"Target"))?
                            {
                                target = Some(
                                    std::str::from_utf8(target_attr)
                                        .map_err(|_| XlsxError::Unexpected("Invalid UTF-8"))?
                                        .to_owned(),
                                );
                            }

                            if is_pivot && target.is_some() {
                                let target_str = target.unwrap();
                                let full_path = if target_str.starts_with("../") {
                                    format!("xl/{}", &target_str[3..])
                                } else {
                                    format!("xl/worksheets/{}", target_str)
                                };
                                pivot_paths.push(full_path);
                            }
                        }
                        Ok(Event::Eof) => break,
                        Err(e) => return Err(XlsxError::Xml(e)),
                        _ => {}
                    }
                }
            }

            // Get basic info from each pivot table
            for path in pivot_paths {
                if let Some(Ok(mut reader)) = xml_reader(&mut self.zip, &path) {
                    if let Ok((name, cache_id)) = parse_pivot_table_info(&mut reader) {
                        let info = PivotTableInfo {
                            name,
                            path,
                            cache_id: Some(cache_id),
                        };
                        self.pivot_tables.add_table(sheet_name.to_string(), info);
                    }
                }
            }
        }

        Ok(())
    }

    /// Parse full pivot table definition
    pub(crate) fn parse_pivot_table(
        &mut self,
        sheet_name: &str,
        pivot_path: &str,
    ) -> Result<PivotTable, XlsxError> {
        let mut reader = xml_reader(&mut self.zip, pivot_path)
            .ok_or_else(|| XlsxError::FileNotFound(pivot_path.into()))??;

        let mut pivot = PivotTable {
            name: String::new(),
            sheet_name: sheet_name.to_string(),
            location: (0, 0),
            source_range: None,
            source_sheet: None,
            cache_id: 0,
            fields: Vec::new(),
            row_fields: Vec::new(),
            column_fields: Vec::new(),
            data_fields: Vec::new(),
            filters: Vec::new(),
        };

        let mut buf = Vec::new();
        let mut field_index = 0;

        loop {
            buf.clear();
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                    b"pivotTableDefinition" => {
                        if let Some(name_attr) = get_attribute(e.attributes(), QName(b"name"))? {
                            pivot.name = reader.decoder().decode(name_attr)?.into_owned();
                        }
                        if let Some(cache_attr) = get_attribute(e.attributes(), QName(b"cacheId"))?
                        {
                            pivot.cache_id = reader
                                .decoder()
                                .decode(cache_attr)?
                                .parse::<u32>()
                                .unwrap_or(0);
                        }
                    }
                    b"location" => {
                        if let Some(ref_attr) = get_attribute(e.attributes(), QName(b"ref"))? {
                            if let Ok(dims) = get_dimension(ref_attr) {
                                pivot.location = dims.start;
                            }
                        }
                    }
                    b"pivotField" => {
                        let mut field = PivotField {
                            name: format!("Field{}", field_index),
                            field_type: PivotFieldType::Hidden,
                            items: Vec::new(),
                            cache_index: Some(field_index),
                        };

                        // Check if this is a row/column field
                        if let Some(axis_attr) = get_attribute(e.attributes(), QName(b"axis"))? {
                            let axis_str = reader.decoder().decode(axis_attr)?;
                            field.field_type = match axis_str.as_ref() {
                                "axisRow" => PivotFieldType::Row,
                                "axisCol" => PivotFieldType::Column,
                                "axisPage" => PivotFieldType::Page,
                                _ => PivotFieldType::Hidden,
                            };
                        }

                        // Parse the field contents including items
                        let mut inner_buf = Vec::new();
                        loop {
                            inner_buf.clear();
                            match reader.read_event_into(&mut inner_buf) {
                                Ok(Event::Start(ref inner_e)) => {
                                    match inner_e.local_name().as_ref() {
                                        b"items" => {
                                            // Parse items
                                            field.items = parse_pivot_items(&mut reader)?;
                                        }
                                        _ => {}
                                    }
                                },
                                Ok(Event::End(ref inner_e)) if inner_e.local_name().as_ref() == b"pivotField" => {
                                    break;
                                }
                                Ok(Event::Eof) => break,
                                Err(e) => return Err(XlsxError::Xml(e)),
                                _ => {}
                            }
                        }

                        pivot.fields.push(field);
                        field_index += 1;
                    }
                    _ => {}
                },
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pivotTableDefinition" => {
                    break
                }
                Ok(Event::Eof) => break,
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => {}
            }
        }

        // Get source range and sheet from cache and update field names
        // For Phase 1, we'll check all caches since the ID mapping might not be perfect
        if pivot.source_range.is_none() {
            // Try to find a cache with source data
            for cache in self.pivot_tables.caches.values() {
                if cache.source_range.is_some() {
                    pivot.source_range = cache.source_range.clone();
                    pivot.source_sheet = cache.source_sheet.clone();
                    // Update the cache with the correct ID
                    let mut cache = cache.clone();
                    cache.id = pivot.cache_id;

                    // Update field names from cache and build items list
                    for (i, field) in pivot.fields.iter_mut().enumerate() {
                        if let Some(cache_field) = cache.fields.get(i) {
                            field.name = cache_field.name.clone();
                            
                            // Update items with resolved values from cache
                            for item in &mut field.items {
                                if item.value.is_empty() {
                                    item.value = if let Some(index) = item.cache_index {
                                        // Use shared item at this index for the original value
                                        if let Some(shared_item) = cache_field.shared_items.get(index as usize) {
                                            match shared_item {
                                                Data::String(s) => s.clone(),
                                                Data::Float(f) => f.to_string(),
                                                Data::Int(i) => i.to_string(),
                                                Data::Bool(b) => b.to_string(),
                                                Data::DateTime(dt) => dt.to_string(),
                                                Data::DateTimeIso(s) => s.clone(),
                                                Data::DurationIso(s) => s.clone(),
                                                Data::Error(e) => format!("#{:?}", e),
                                                Data::Empty => String::new(),
                                            }
                                        } else {
                                            format!("Item {}", index)
                                        }
                                    } else if let Some(ref item_type) = item.item_type {
                                        // Special items like "default" for subtotals
                                        format!("({})", item_type)
                                    } else {
                                        String::new()
                                    };
                                }
                            }
                        }
                    }

                    self.pivot_tables.caches.insert(pivot.cache_id, cache);
                    break;
                }
            }
        } else if let Some(cache) = self.pivot_tables.caches.get(&pivot.cache_id) {
            // Update field names from cache and build items list
            for (i, field) in pivot.fields.iter_mut().enumerate() {
                if let Some(cache_field) = cache.fields.get(i) {
                    field.name = cache_field.name.clone();
                    
                    // Update items with resolved values from cache
                    for item in &mut field.items {
                        if item.value.is_empty() {
                            item.value = if let Some(index) = item.cache_index {
                                // Use shared item at this index for the original value
                                if let Some(shared_item) = cache_field.shared_items.get(index as usize) {
                                    format!("{:?}", shared_item)
                                } else {
                                    format!("Item {}", index)
                                }
                            } else if let Some(ref item_type) = item.item_type {
                                // Special items like "default" for subtotals
                                format!("({})", item_type)
                            } else {
                                String::new()
                            };
                        }
                    }
                }
            }
            // Also get source sheet
            pivot.source_sheet = cache.source_sheet.clone();
        }

        Ok(pivot)
    }
}

/// Parse pivot cache metadata (without records)
fn parse_pivot_cache_metadata<RS: Read + Seek>(
    reader: &mut XlReader<'_, RS>,
    cache_id: u32,
) -> Result<PivotCache, XlsxError> {
            let mut cache = PivotCache {
            id: cache_id,
            source_type: PivotSourceType::Worksheet,
            source_range: None,
            source_sheet: None,
            fields: Vec::new(),
            has_records: false,
            records: None,
            cache_path: None,
        };

    let mut buf = Vec::new();
    let mut in_cache_source = false;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => match e.local_name().as_ref() {
                b"pivotCacheDefinition" => {
                    if let Some(count_attr) = get_attribute(e.attributes(), QName(b"recordCount"))?
                    {
                        let count_str = reader.decoder().decode(count_attr)?;
                        if let Ok(count) = count_str.parse::<u32>() {
                            cache.has_records = count > 0;
                        }
                    }
                }
                b"cacheSource" => {
                    in_cache_source = true;
                    // Get the type attribute
                    if let Some(type_attr) = get_attribute(e.attributes(), QName(b"type"))? {
                        let type_str = reader.decoder().decode(type_attr)?;
                        cache.source_type = match type_str.as_ref() {
                            "worksheet" => PivotSourceType::Worksheet,
                            "external" => PivotSourceType::External,
                            "consolidation" => PivotSourceType::Consolidation,
                            "scenario" => PivotSourceType::Scenario,
                            _ => PivotSourceType::Worksheet,
                        };
                    }
                }
                b"worksheetSource" if in_cache_source => {
                    if let Some(ref_attr) = get_attribute(e.attributes(), QName(b"ref"))? {
                        cache.source_range = Some(reader.decoder().decode(ref_attr)?.into_owned());
                    }
                    if let Some(sheet_attr) = get_attribute(e.attributes(), QName(b"sheet"))? {
                        cache.source_sheet =
                            Some(reader.decoder().decode(sheet_attr)?.into_owned());
                    }
                }
                b"cacheField" => {
                    let mut field = PivotCacheField {
                        name: String::new(),
                        data_type: PivotFieldDataType::String,
                        shared_items: Vec::new(),
                    };

                    if let Some(name_attr) = get_attribute(e.attributes(), QName(b"name"))? {
                        field.name = reader.decoder().decode(name_attr)?.into_owned();
                    }

                    // Parse the cache field contents including shared items
                    let mut inner_buf = Vec::new();
                    loop {
                        inner_buf.clear();
                        match reader.read_event_into(&mut inner_buf) {
                            Ok(Event::Start(ref inner_e)) => {
                                match inner_e.local_name().as_ref() {
                                    b"sharedItems" => {
                                        // Parse shared items
                                        field.shared_items = parse_shared_items(reader)?;
                                    }
                                    _ => {}
                                }
                            },
                            Ok(Event::End(ref inner_e)) if inner_e.local_name().as_ref() == b"cacheField" => {
                                break;
                            }
                            Ok(Event::Eof) => return Err(XlsxError::Unexpected("Unexpected EOF in cacheField")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => {}
                        }
                    }

                    cache.fields.push(field);
                }
                _ => {}
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"cacheSource" => in_cache_source = false,
                b"pivotCacheDefinition" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(cache)
}

/// Parse shared items from cache field
fn parse_shared_items<RS: Read + Seek>(
    reader: &mut XlReader<'_, RS>,
) -> Result<Vec<Data>, XlsxError> {
    let mut items = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                // Check if this is a self-closing element with attributes
                if e.attributes().count() > 0 {
                    match e.local_name().as_ref() {
                        b"s" => {
                            // String item
                            if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                                let v_str = reader.decoder().decode(v_attr)?;
                                items.push(Data::String(v_str.into_owned()));
                            }
                        }
                        b"n" => {
                            // Number item
                            if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                                let v_str = reader.decoder().decode(v_attr)?;
                                if let Ok(num) = v_str.parse::<f64>() {
                                    items.push(Data::Float(num));
                                } else {
                                    items.push(Data::String(v_str.into_owned()));
                                }
                            }
                        }
                        b"d" => {
                            // Date item
                            if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                                let v_str = reader.decoder().decode(v_attr)?;
                                items.push(Data::String(v_str.into_owned())); // TODO: Parse as date
                            }
                        }
                        b"b" => {
                            // Boolean item
                            if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                                let v_str = reader.decoder().decode(v_attr)?;
                                items.push(Data::Bool(v_str == "1" || v_str.to_lowercase() == "true"));
                            }
                        }
                        b"m" => {
                            // Missing item
                            items.push(Data::Empty);
                        }
                        _ => {}
                    }
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                b"s" => {
                    // String item
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        items.push(Data::String(v_str.into_owned()));
                    }
                }
                b"n" => {
                    // Number item
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        if let Ok(num) = v_str.parse::<f64>() {
                            items.push(Data::Float(num));
                        } else {
                            items.push(Data::String(v_str.into_owned()));
                        }
                    }
                }
                b"d" => {
                    // Date item
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        items.push(Data::String(v_str.into_owned())); // TODO: Parse as date
                    }
                }
                b"b" => {
                    // Boolean item
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        items.push(Data::Bool(v_str == "1" || v_str.to_lowercase() == "true"));
                    }
                }
                b"m" => {
                    // Missing item
                    items.push(Data::Empty);
                }
                _ => {}
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sharedItems" => {
                break;
            }
            Ok(Event::Eof) => return Err(XlsxError::Unexpected("Unexpected EOF in sharedItems")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(items)
}

/// Parse pivot items from a pivotField
fn parse_pivot_items<RS: Read + Seek>(
    reader: &mut XlReader<'_, RS>,
) -> Result<Vec<PivotItem>, XlsxError> {
    
    let mut items = Vec::new();
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                if e.local_name().as_ref() == b"item" {
                    let mut item = PivotItem {
                        value: String::new(),
                        cache_index: None,
                        custom_name: None,
                        item_type: None,
                    };

                    // Get the index reference
                    if let Some(x_attr) = get_attribute(e.attributes(), QName(b"x"))? {
                        let x_str = reader.decoder().decode(x_attr)?;
                        if let Ok(index) = x_str.parse::<u32>() {
                            item.cache_index = Some(index);
                        }
                    }

                    // Get custom name if present
                    if let Some(n_attr) = get_attribute(e.attributes(), QName(b"n"))? {
                        item.custom_name = Some(reader.decoder().decode(n_attr)?.into_owned());
                    }

                    // Get item type if present
                    if let Some(t_attr) = get_attribute(e.attributes(), QName(b"t"))? {
                        item.item_type = Some(reader.decoder().decode(t_attr)?.into_owned());
                    }

                    items.push(item);
                }
            }
            Ok(Event::Empty(ref e)) => {
                match e.local_name().as_ref() {
                b"item" => {
                    let mut item = PivotItem {
                        value: String::new(),
                        cache_index: None,
                        custom_name: None,
                        item_type: None,
                    };

                    // Get the index reference
                    if let Some(x_attr) = get_attribute(e.attributes(), QName(b"x"))? {
                        let x_str = reader.decoder().decode(x_attr)?;
                        if let Ok(index) = x_str.parse::<u32>() {
                            item.cache_index = Some(index);
                        }
                    }

                    // Get custom name if present
                    if let Some(n_attr) = get_attribute(e.attributes(), QName(b"n"))? {
                        item.custom_name = Some(reader.decoder().decode(n_attr)?.into_owned());
                    }

                    // Get item type if present
                    if let Some(t_attr) = get_attribute(e.attributes(), QName(b"t"))? {
                        item.item_type = Some(reader.decoder().decode(t_attr)?.into_owned());
                    }

                    items.push(item);
                }
                _ => {}
                }
            },
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"items" => {
                break;
            }
            Ok(Event::Eof) => return Err(XlsxError::Unexpected("Unexpected EOF in items")),
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(items)
}

/// Parse pivot cache records from pivotCacheRecords{n}.xml
fn parse_pivot_cache_records<RS: Read + Seek>(
    reader: &mut XlReader<'_, RS>,
    fields: &[PivotCacheField],
) -> Result<Vec<Vec<Data>>, XlsxError> {
    let mut records = Vec::new();
    let mut buf = Vec::new();
    let mut current_record = Vec::new();
    let mut field_index = 0;

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                match e.local_name().as_ref() {
                b"r" => {
                    // Start of a new record
                    current_record.clear();
                    field_index = 0;
                }
                b"x" => {
                    // Indexed value (references shared items)
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        if let Ok(index) = v_str.parse::<usize>() {
                            if field_index < fields.len() {
                                if let Some(item) = fields[field_index].shared_items.get(index) {
                                    current_record.push(item.clone());
                                } else {
                                    current_record.push(Data::Empty);
                                }
                            }
                        }
                    }
                    field_index += 1;
                }
                b"s" | b"n" | b"d" | b"b" => {
                    // These elements have values in 'v' attribute
                    if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                        let v_str = reader.decoder().decode(v_attr)?;
                        let data = match e.local_name().as_ref() {
                            b"s" => Data::String(v_str.into_owned()),
                            b"n" => {
                                if let Ok(num) = v_str.parse::<f64>() {
                                    Data::Float(num)
                                } else {
                                    Data::String(v_str.into_owned())
                                }
                            }
                            b"d" => Data::String(v_str.into_owned()), // TODO: Parse as date
                            b"b" => Data::Bool(v_str == "1" || v_str.to_lowercase() == "true"),
                            _ => Data::Empty,
                        };
                        current_record.push(data);
                    } else {
                        current_record.push(Data::Empty);
                    }
                    field_index += 1;
                }
                b"m" => {
                    // Missing value
                    current_record.push(Data::Empty);
                    field_index += 1;
                }
                _ => {}
                }
            },
            Ok(Event::Empty(ref e)) => {
                // Handle self-closing tags with values in attributes
                match e.local_name().as_ref() {
                    b"x" => {
                        // Indexed value (references shared items)
                        if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                            let v_str = reader.decoder().decode(v_attr)?;
                            if let Ok(index) = v_str.parse::<usize>() {
                                if field_index < fields.len() {
                                    if let Some(item) = fields[field_index].shared_items.get(index) {
                                        current_record.push(item.clone());
                                    } else {
                                        current_record.push(Data::Empty);
                                    }
                                }
                            }
                        }
                        field_index += 1;
                    }
                    b"s" | b"n" | b"d" | b"b" => {
                        // These elements have values in 'v' attribute
                        if let Some(v_attr) = get_attribute(e.attributes(), QName(b"v"))? {
                            let v_str = reader.decoder().decode(v_attr)?;
                            let data = match e.local_name().as_ref() {
                                b"s" => Data::String(v_str.into_owned()),
                                b"n" => {
                                    if let Ok(num) = v_str.parse::<f64>() {
                                        Data::Float(num)
                                    } else {
                                        Data::String(v_str.into_owned())
                                    }
                                }
                                b"d" => Data::String(v_str.into_owned()), // TODO: Parse as date
                                b"b" => Data::Bool(v_str == "1" || v_str.to_lowercase() == "true"),
                                _ => Data::Empty,
                            };
                            current_record.push(data.clone());
                            println!("DEBUG: Added data: {:?} at index {}", data, field_index);
                        } else {
                            current_record.push(Data::Empty);
                            println!("DEBUG: Added empty data at index {}", field_index);
                        }
                        field_index += 1;
                    }
                    b"m" => {
                        // Missing value
                        current_record.push(Data::Empty);
                        field_index += 1;
                    }
                    _ => {}
                }
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                b"r" => {
                    // End of record
                    if !current_record.is_empty() {
                        records.push(current_record.clone());
                    }
                }
                b"pivotCacheRecords" => break,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Ok(records)
}

/// Parse basic pivot table info
fn parse_pivot_table_info<RS: Read + Seek>(
    reader: &mut XlReader<'_, RS>,
) -> Result<(String, u32), XlsxError> {
    let mut name = String::new();
    let mut cache_id = 0;
    let mut buf = Vec::new();

    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pivotTableDefinition" => {
                if let Some(name_attr) = get_attribute(e.attributes(), QName(b"name"))? {
                    name = reader.decoder().decode(name_attr)?.into_owned();
                }
                if let Some(cache_attr) = get_attribute(e.attributes(), QName(b"cacheId"))? {
                    let cache_str = reader.decoder().decode(cache_attr)?;
                    cache_id = cache_str.parse::<u32>().unwrap_or(0);
                }

                if !name.is_empty() {
                    return Ok((name, cache_id));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(XlsxError::Xml(e)),
            _ => {}
        }
    }

    Err(XlsxError::Unexpected("Failed to parse pivot table info"))
}
