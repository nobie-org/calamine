use quick_xml::{
    events::{attributes::Attribute, BytesStart, Event},
    name::QName,
};
use std::{
    borrow::Borrow,
    collections::HashMap,
    io::{Read, Seek},
};

use super::{
    get_attribute, get_dimension, get_row, get_row_column, read_string, replace_cell_names,
    ColumnDefinition, ColumnWidths, Dimensions, RowDefinition, RowDefinitions, XlReader,
};
use crate::{
    datatype::DataRef,
    formats::{format_excel_f64_ref, CellFormat, CellStyle},
    Cell, XlsxError,
};

type FormulaMap = HashMap<(u32, u32), (i64, i64)>;
type CellWithFormatting<'a> = (Cell<DataRef<'a>>, Option<&'a CellStyle>);

/// An xlsx Cell Iterator
pub struct XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    xml: XlReader<'a, RS>,
    strings: &'a [String],
    formats: &'a [CellStyle],
    is_1904: bool,
    dimensions: Dimensions,
    row_index: u32,
    col_index: u32,
    buf: Vec<u8>,
    cell_buf: Vec<u8>,
    formulas: Vec<Option<(String, FormulaMap)>>,
    column_widths: ColumnWidths,
    row_definitions: RowDefinitions,
    // Spill tracking for dynamic array sources: ranges defined by <f t="array" ref="...">
    spill_sources: Vec<Dimensions>,
    // Whether the last returned cell had its own <f> formula element
    last_cell_had_formula: bool,
}

impl<'a, RS> XlsxCellReader<'a, RS>
where
    RS: Read + Seek,
{
    pub fn new(
        mut xml: XlReader<'a, RS>,
        strings: &'a [String],
        formats: &'a [CellStyle],
        is_1904: bool,
    ) -> Result<Self, XlsxError> {
        let mut buf = Vec::with_capacity(1024);
        let mut dimensions = Dimensions::default();
        let mut column_widths = ColumnWidths::new();
        let mut row_definitions = RowDefinitions::new();
        let mut sh_type = None;
        'xml: loop {
            buf.clear();
            match xml.read_event_into(&mut buf).map_err(XlsxError::Xml)? {
                Event::Start(ref e) => match e.local_name().as_ref() {
                    b"dimension" => {
                        for a in e.attributes() {
                            if let Attribute {
                                key: QName(b"ref"),
                                value: rdim,
                            } = a.map_err(XlsxError::XmlAttr)?
                            {
                                dimensions = get_dimension(&rdim)?;
                                continue 'xml;
                            }
                        }
                        return Err(XlsxError::UnexpectedNode("dimension"));
                    }
                    b"sheetFormatPr" => {
                        // Parse sheet format properties - store raw values
                        for a in e.attributes() {
                            match a.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"defaultColWidth"),
                                    value: v,
                                } => {
                                    if let Ok(width_str) = xml.decoder().decode(&v) {
                                        if let Ok(width) = width_str.parse::<f64>() {
                                            column_widths.sheet_format.default_col_width =
                                                Some(width);
                                            row_definitions.sheet_format.default_col_width =
                                                Some(width);
                                        }
                                    }
                                }
                                Attribute {
                                    key: QName(b"baseColWidth"),
                                    value: v,
                                } => {
                                    if let Ok(width_str) = xml.decoder().decode(&v) {
                                        if let Ok(width) = width_str.parse::<u8>() {
                                            column_widths.sheet_format.base_col_width = Some(width);
                                            row_definitions.sheet_format.base_col_width =
                                                Some(width);
                                        }
                                    }
                                }
                                Attribute {
                                    key: QName(b"defaultRowHeight"),
                                    value: v,
                                } => {
                                    if let Ok(height_str) = xml.decoder().decode(&v) {
                                        if let Ok(height) = height_str.parse::<f64>() {
                                            column_widths.sheet_format.default_row_height =
                                                Some(height);
                                            row_definitions.sheet_format.default_row_height =
                                                Some(height);
                                        }
                                    }
                                }
                                _ => {}
                            }
                        }
                    }
                    b"cols" => {
                        // Parse column definitions - store raw values
                        let mut inner_buf = Vec::with_capacity(512);
                        loop {
                            inner_buf.clear();
                            match xml
                                .read_event_into(&mut inner_buf)
                                .map_err(XlsxError::Xml)?
                            {
                                Event::Start(ref col) | Event::Empty(ref col)
                                    if col.local_name().as_ref() == b"col" =>
                                {
                                    let mut def = ColumnDefinition {
                                        min: 1,
                                        max: 1,
                                        width: None,
                                        style: None,
                                        custom_width: None,
                                        best_fit: None,
                                        hidden: None,
                                        outline_level: None,
                                        collapsed: None,
                                    };

                                    for a in col.attributes() {
                                        match a.map_err(XlsxError::XmlAttr)? {
                                            Attribute {
                                                key: QName(b"min"),
                                                value: v,
                                            } => {
                                                def.min = atoi_simd::parse::<u32>(&v).unwrap_or(1);
                                            }
                                            Attribute {
                                                key: QName(b"max"),
                                                value: v,
                                            } => {
                                                def.max = atoi_simd::parse::<u32>(&v).unwrap_or(1);
                                            }
                                            Attribute {
                                                key: QName(b"width"),
                                                value: v,
                                            } => {
                                                if let Ok(width_str) = xml.decoder().decode(&v) {
                                                    if let Ok(width) = width_str.parse::<f64>() {
                                                        def.width = Some(width);
                                                    }
                                                }
                                            }
                                            Attribute {
                                                key: QName(b"style"),
                                                value: v,
                                            } => {
                                                def.style = atoi_simd::parse::<u32>(&v).ok();
                                            }
                                            Attribute {
                                                key: QName(b"customWidth"),
                                                value: v,
                                            } => {
                                                def.custom_width =
                                                    Some(&*v == b"1" || &*v == b"true");
                                            }
                                            Attribute {
                                                key: QName(b"bestFit"),
                                                value: v,
                                            } => {
                                                def.best_fit = Some(&*v == b"1" || &*v == b"true");
                                            }
                                            Attribute {
                                                key: QName(b"hidden"),
                                                value: v,
                                            } => {
                                                def.hidden = Some(&*v == b"1" || &*v == b"true");
                                            }
                                            Attribute {
                                                key: QName(b"outlineLevel"),
                                                value: v,
                                            } => {
                                                def.outline_level = atoi_simd::parse::<u8>(&v).ok();
                                            }
                                            Attribute {
                                                key: QName(b"collapsed"),
                                                value: v,
                                            } => {
                                                def.collapsed = Some(&*v == b"1" || &*v == b"true");
                                            }
                                            _ => {}
                                        }
                                    }

                                    // Store raw column definition without conversion
                                    column_widths.add_column_definition(def);
                                }
                                Event::End(ref e) if e.local_name().as_ref() == b"cols" => break,
                                Event::Eof => return Err(XlsxError::XmlEof("cols")),
                                _ => {}
                            }
                        }
                    }
                    b"sheetData" => break,
                    typ => {
                        if sh_type.is_none() {
                            sh_type = Some(xml.decoder().decode(typ)?.to_string());
                        }
                    }
                },
                Event::Eof => {
                    if let Some(typ) = sh_type {
                        return Err(XlsxError::NotAWorksheet(typ));
                    } else {
                        return Err(XlsxError::XmlEof("worksheet"));
                    }
                }
                _ => (),
            }
        }
        Ok(Self {
            xml,
            strings,
            formats,
            is_1904,
            dimensions,
            row_index: 0,
            col_index: 0,
            buf: Vec::with_capacity(1024),
            cell_buf: Vec::with_capacity(1024),
            formulas: Vec::with_capacity(1024),
            column_widths,
            row_definitions,
            spill_sources: Vec::with_capacity(32),
            last_cell_had_formula: false,
        })
    }

    /// Check if an absolute position is within any recorded spill source range
    pub fn is_in_spill(&self, pos: (u32, u32)) -> bool {
        let (row, col) = pos;
        self.spill_sources.iter().any(|d| d.contains(row, col))
    }

    /// Whether the last returned cell had its own formula (<f> element)
    pub fn last_cell_had_formula(&self) -> bool {
        self.last_cell_had_formula
    }

    pub fn dimensions(&self) -> Dimensions {
        self.dimensions
    }

    /// Get column widths information
    pub fn column_widths(&self) -> &ColumnWidths {
        &self.column_widths
    }

    /// Get row definitions information
    pub fn row_definitions(&self) -> &RowDefinitions {
        &self.row_definitions
    }

    pub fn next_cell(&mut self) -> Result<Option<Cell<DataRef<'a>>>, XlsxError> {
        self.next_cell_with_formatting()
            .map(|opt| opt.map(|(cell, _)| cell))
    }

    /// Get the next cell with its formatting information
    pub fn next_cell_with_formatting(
        &mut self,
    ) -> Result<Option<CellWithFormatting<'a>>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;

                        // Parse row definition attributes
                        let mut row_def = RowDefinition {
                            r: row,
                            height: None,
                            style: None,
                            custom_height: None,
                            hidden: None,
                            outline_level: None,
                            collapsed: None,
                            thick_top: None,
                            thick_bot: None,
                        };

                        // Parse row attributes
                        for a in row_element.attributes() {
                            match a.map_err(XlsxError::XmlAttr)? {
                                Attribute {
                                    key: QName(b"ht"),
                                    value: v,
                                } => {
                                    if let Ok(height_str) = self.xml.decoder().decode(&v) {
                                        if let Ok(height) = height_str.parse::<f64>() {
                                            row_def.height = Some(height);
                                        }
                                    }
                                }
                                Attribute {
                                    key: QName(b"s"),
                                    value: v,
                                } => {
                                    row_def.style = atoi_simd::parse::<u32>(&v).ok();
                                }
                                Attribute {
                                    key: QName(b"customHeight"),
                                    value: v,
                                } => {
                                    row_def.custom_height = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"hidden"),
                                    value: v,
                                } => {
                                    row_def.hidden = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"outlineLevel"),
                                    value: v,
                                } => {
                                    row_def.outline_level = atoi_simd::parse::<u8>(&v).ok();
                                }
                                Attribute {
                                    key: QName(b"collapsed"),
                                    value: v,
                                } => {
                                    row_def.collapsed = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"thickTop"),
                                    value: v,
                                } => {
                                    row_def.thick_top = Some(&*v == b"1" || &*v == b"true");
                                }
                                Attribute {
                                    key: QName(b"thickBot"),
                                    value: v,
                                } => {
                                    row_def.thick_bot = Some(&*v == b"1" || &*v == b"true");
                                }
                                _ => {}
                            }
                        }

                        // Store row definition if it has any meaningful information
                        if row_def.height.is_some()
                            || row_def.style.is_some()
                            || row_def.custom_height == Some(true)
                            || row_def.hidden == Some(true)
                            || row_def.outline_level.is_some()
                            || row_def.collapsed == Some(true)
                            || row_def.thick_top == Some(true)
                            || row_def.thick_bot == Some(true)
                        {
                            self.row_definitions.add_row_definition(row_def);
                        }
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };
                    
                    // Extract formatting information from the cell element
                    let cell_formatting = match get_attribute(c_element.attributes(), QName(b"s")) {
                        Ok(Some(style)) => {
                            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
                            self.formats.get(id)
                        }
                        _ => None,
                    };
                    
                    let mut value = DataRef::Empty;
                    let mut had_formula = false;

                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                if e.local_name().as_ref() == b"f" {
                                    had_formula = true;
                                    if let Ok(Some(t)) = get_attribute(e.attributes(), QName(b"t"))
                                    {
                                        if t == b"array" {
                                            if let Ok(Some(r)) =
                                                get_attribute(e.attributes(), QName(b"ref"))
                                            {
                                                let dim = get_dimension(r)?;
                                                self.spill_sources.push(dim);
                                            }
                                        }
                                    }
                                }
                                let (val, _) = read_value_with_formatting(
                                    self.strings,
                                    self.formats,
                                    self.is_1904,
                                    &mut self.xml,
                                    e,
                                    c_element,
                                )?;
                                value = val;
                                // Keep the formatting we already extracted from the cell element
                            }
                            Ok(Event::Empty(ref e)) if e.local_name().as_ref() == b"f" => {
                                // Catch inline empty <f .../> tags too
                                had_formula = true;
                                if let Ok(Some(t)) = get_attribute(e.attributes(), QName(b"t")) {
                                    if t == b"array" {
                                        if let Ok(Some(r)) =
                                            get_attribute(e.attributes(), QName(b"ref"))
                                        {
                                            let dim = get_dimension(r)?;
                                            self.spill_sources.push(dim);
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    self.last_cell_had_formula = had_formula;
                    return Ok(Some((Cell::new(pos, value), cell_formatting)));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }

    /// Get formatting information by style index
    pub fn get_formatting_by_index(&self, style_index: usize) -> Option<&CellStyle> {
        self.formats.get(style_index)
    }

    pub fn next_formula(&mut self) -> Result<Option<Cell<String>>, XlsxError> {
        self.next_formula_with_formatting()
            .map(|opt| opt.map(|(cell, _)| cell))
    }

    /// Get the next formula with its formatting information
    pub fn next_formula_with_formatting(
        &mut self,
    ) -> Result<Option<(Cell<String>, Option<&CellStyle>)>, XlsxError> {
        loop {
            self.buf.clear();
            match self.xml.read_event_into(&mut self.buf) {
                Ok(Event::Start(ref row_element))
                    if row_element.local_name().as_ref() == b"row" =>
                {
                    let attribute = get_attribute(row_element.attributes(), QName(b"r"))?;
                    if let Some(range) = attribute {
                        let row = get_row(range)?;
                        self.row_index = row;
                    }
                }
                Ok(Event::End(ref row_element)) if row_element.local_name().as_ref() == b"row" => {
                    self.row_index += 1;
                    self.col_index = 0;
                }
                Ok(Event::Start(ref c_element)) if c_element.local_name().as_ref() == b"c" => {
                    let attribute = get_attribute(c_element.attributes(), QName(b"r"))?;
                    let pos = if let Some(range) = attribute {
                        let (row, col) = get_row_column(range)?;
                        self.col_index = col;
                        (row, col)
                    } else {
                        (self.row_index, self.col_index)
                    };

                    // Extract formatting information from the cell element
                    let cell_formatting = match get_attribute(c_element.attributes(), QName(b"s")) {
                        Ok(Some(style)) => {
                            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
                            self.formats.get(id)
                        }
                        _ => None,
                    };

                    let mut value = None;
                    loop {
                        self.cell_buf.clear();
                        match self.xml.read_event_into(&mut self.cell_buf) {
                            Ok(Event::Start(ref e)) => {
                                let formula = read_formula(&mut self.xml, e)?;
                                if let Some(f) = formula.borrow() {
                                    value = Some(f.clone());
                                }
                                if let Ok(Some(b"shared")) =
                                    get_attribute(e.attributes(), QName(b"t"))
                                {
                                    // shared formula
                                    let mut offset_map: HashMap<(u32, u32), (i64, i64)> =
                                        HashMap::new();
                                    // shared index
                                    let shared_index =
                                        match get_attribute(e.attributes(), QName(b"si"))? {
                                            Some(res) => match atoi_simd::parse::<usize>(res) {
                                                Ok(res) => res,
                                                Err(_) => {
                                                    return Err(XlsxError::Unexpected(
                                                        "si attribute must be a number",
                                                    ));
                                                }
                                            },
                                            None => {
                                                return Err(XlsxError::Unexpected(
                                                    "si attribute is mandatory if it is shared",
                                                ));
                                            }
                                        };
                                    // shared reference
                                    match get_attribute(e.attributes(), QName(b"ref"))? {
                                        Some(res) => {
                                            // orignal reference formula
                                            let reference = get_dimension(res)?;
                                            // dynamic arrays also use t="array" with a ref; capture those as sources
                                            if let Ok(Some(t)) =
                                                get_attribute(e.attributes(), QName(b"t"))
                                            {
                                                if t == b"array" {
                                                    self.spill_sources.push(reference);
                                                }
                                            }
                                            // build offset map for every cell in the shared-formula rectangle
                                            for r in reference.start.0..=reference.end.0 {
                                                for c in reference.start.1..=reference.end.1 {
                                                    offset_map.insert(
                                                        (r, c),
                                                        (
                                                            r as i64 - pos.0 as i64,
                                                            c as i64 - pos.1 as i64,
                                                        ),
                                                    );
                                                }
                                            }

                                            if let Some(f) = formula.borrow() {
                                                while self.formulas.len() < shared_index {
                                                    self.formulas.push(None);
                                                }
                                                self.formulas.push(Some((f.clone(), offset_map)));
                                            }
                                            value = formula;
                                        }
                                        None => {
                                            // calculated formula
                                            if let Some(Some((f, offset_map))) =
                                                self.formulas.get(shared_index)
                                            {
                                                if let Some(offset) = offset_map.get(&pos) {
                                                    value = Some(replace_cell_names(f, *offset)?);
                                                }
                                            }
                                        }
                                    }
                                }
                                // capture non-shared array formulas with ref
                                if let Ok(Some(t)) = get_attribute(e.attributes(), QName(b"t")) {
                                    if t == b"array" {
                                        if let Ok(Some(r)) =
                                            get_attribute(e.attributes(), QName(b"ref"))
                                        {
                                            let reference = get_dimension(r)?;
                                            self.spill_sources.push(reference);
                                        }
                                    }
                                }
                            }
                            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"c" => break,
                            Ok(Event::Eof) => return Err(XlsxError::XmlEof("c")),
                            Err(e) => return Err(XlsxError::Xml(e)),
                            _ => (),
                        }
                    }
                    self.col_index += 1;
                    return Ok(Some((
                        Cell::new(pos, value.unwrap_or_default()),
                        cell_formatting,
                    )));
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    return Ok(None);
                }
                Ok(Event::Eof) => return Err(XlsxError::XmlEof("sheetData")),
                Err(e) => return Err(XlsxError::Xml(e)),
                _ => (),
            }
        }
    }
}

fn read_value_with_formatting<'s, 'f, RS>(
    strings: &'s [String],
    formats: &'f [CellStyle],
    is_1904: bool,
    xml: &mut XlReader<'_, RS>,
    e: &BytesStart<'_>,
    c_element: &BytesStart<'_>,
) -> Result<(DataRef<'s>, Option<&'f CellStyle>), XlsxError>
where
    RS: Read + Seek,
{
    // Extract style information from the cell element
    let cell_formatting = match get_attribute(c_element.attributes(), QName(b"s")) {
        Ok(Some(style)) => {
            let id = atoi_simd::parse::<usize>(style).unwrap_or(0);
            formats.get(id)
        }
        _ => None,
    };

    let value = match e.local_name().as_ref() {
        b"is" => {
            // inlineStr
            read_string(xml, e.name())?.map_or(DataRef::Empty, DataRef::String)
        }
        b"v" => {
            // value
            let mut v = String::new();
            let mut v_buf = Vec::new();
            loop {
                v_buf.clear();
                match xml.read_event_into(&mut v_buf)? {
                    Event::Text(t) => v.push_str(&t.unescape()?),
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("v")),
                    _ => (),
                }
            }
            read_v(
                v,
                strings,
                cell_formatting.map(|f| &f.number_format),
                c_element,
                is_1904,
            )?
        }
        b"f" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            DataRef::Empty
        }
        _n => return Err(XlsxError::UnexpectedNode("v, f, or is")),
    };

    Ok((value, cell_formatting))
}

/// read the contents of a <v> cell
fn read_v<'s>(
    v: String,
    strings: &'s [String],
    cell_format: Option<&CellFormat>,
    c_element: &BytesStart<'_>,
    is_1904: bool,
) -> Result<DataRef<'s>, XlsxError> {
    match get_attribute(c_element.attributes(), QName(b"t"))? {
        Some(b"s") => {
            // shared string
            let idx = atoi_simd::parse::<usize>(v.as_bytes()).unwrap_or(0);
            Ok(DataRef::SharedString(&strings[idx]))
        }
        Some(b"b") => {
            // boolean
            Ok(DataRef::Bool(v != "0"))
        }
        Some(b"e") => {
            // error
            Ok(DataRef::Error(v.parse()?))
        }
        Some(b"d") => {
            // date
            Ok(DataRef::DateTimeIso(v))
        }
        Some(b"str") => {
            // string
            Ok(DataRef::String(v))
        }
        Some(b"n") => {
            // n - number
            if v.is_empty() {
                Ok(DataRef::Empty)
            } else {
                v.parse()
                    .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                    .map_err(XlsxError::ParseFloat)
            }
        }
        None => {
            // If type is not known, we try to parse as Float for utility, but fall back to
            // String if this fails.
            v.parse()
                .map(|n| format_excel_f64_ref(n, cell_format, is_1904))
                .or(Ok(DataRef::String(v)))
        }
        Some(b"is") => {
            // this case should be handled in outer loop over cell elements, in which
            // case read_inline_str is called instead. Case included here for completeness.
            Err(XlsxError::Unexpected(
                "called read_value on a cell of type inlineStr",
            ))
        }
        Some(t) => {
            let t = std::str::from_utf8(t).unwrap_or("<utf8 error>").to_string();
            Err(XlsxError::CellTAttribute(t))
        }
    }
}

fn read_formula<RS>(xml: &mut XlReader<RS>, e: &BytesStart) -> Result<Option<String>, XlsxError>
where
    RS: Read + Seek,
{
    match e.local_name().as_ref() {
        b"is" | b"v" => {
            xml.read_to_end_into(e.name(), &mut Vec::new())?;
            Ok(None)
        }
        b"f" => {
            let mut f_buf = Vec::with_capacity(512);
            let mut f = String::new();
            loop {
                match xml.read_event_into(&mut f_buf)? {
                    Event::Text(t) => f.push_str(&t.unescape()?),
                    Event::End(end) if end.name() == e.name() => break,
                    Event::Eof => return Err(XlsxError::XmlEof("f")),
                    _ => (),
                }
                f_buf.clear();
            }
            Ok(Some(f))
        }
        _ => Err(XlsxError::UnexpectedNode("v, f, or is")),
    }
}
