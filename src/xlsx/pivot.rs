//! Pivot table parsing for XLSX files

use quick_xml::{events::Event, name::QName};
use std::io::{Read, Seek};

use crate::pivot::{
    PivotCache, PivotCacheField, PivotField, PivotFieldDataType, PivotFieldType, PivotSourceType,
    PivotTable, PivotTableInfo,
};
use crate::{Reader, XlsxError};

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
                let cache = parse_pivot_cache_metadata(&mut reader, idx as u32)?;
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
                        // Simple field parsing for Phase 1
                        let field = PivotField {
                            name: format!("Field{}", field_index),
                            field_type: PivotFieldType::Hidden,
                            items: Vec::new(),
                            cache_index: Some(field_index),
                        };

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

                    // Update field names from cache
                    for (i, field) in pivot.fields.iter_mut().enumerate() {
                        if let Some(cache_field) = cache.fields.get(i) {
                            field.name = cache_field.name.clone();
                        }
                    }

                    self.pivot_tables.caches.insert(pivot.cache_id, cache);
                    break;
                }
            }
        } else if let Some(cache) = self.pivot_tables.caches.get(&pivot.cache_id) {
            // Update field names from cache
            for (i, field) in pivot.fields.iter_mut().enumerate() {
                if let Some(cache_field) = cache.fields.get(i) {
                    field.name = cache_field.name.clone();
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
