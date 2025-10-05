/// Raw column definition from Excel XML
#[derive(Debug, Clone)]
pub struct ColumnDefinition {
    /// min attribute: First column affected (1-based, 1-16384)
    pub min: u32,
    /// max attribute: Last column affected (1-based, 1-16384)
    pub max: u32,
    /// width attribute: Column width as a double
    pub width: Option<f64>,
    /// style attribute: Style index (0-65429)
    pub style: Option<u32>,
    /// customWidth attribute
    pub custom_width: Option<bool>,
    /// bestFit attribute
    pub best_fit: Option<bool>,
    /// hidden attribute
    pub hidden: Option<bool>,
    /// outlineLevel attribute (0-7)
    pub outline_level: Option<u8>,
    /// collapsed attribute
    pub collapsed: Option<bool>,
}

/// Raw sheet format properties from XML
#[derive(Debug, Clone, Default)]
pub struct SheetFormatProperties {
    /// defaultColWidth attribute
    pub default_col_width: Option<f64>,
    /// baseColWidth attribute
    pub base_col_width: Option<u8>,
    /// defaultRowHeight attribute
    pub default_row_height: Option<f64>,
    /// customHeight attribute
    pub custom_height: Option<bool>,
    /// zeroHeight attribute
    pub zero_height: Option<bool>,
    /// thickTop attribute
    pub thick_top: Option<bool>,
    /// thickBottom attribute
    pub thick_bottom: Option<bool>,
    /// outlineLevelRow attribute
    pub outline_level_row: Option<u8>,
    /// outlineLevelCol attribute
    pub outline_level_col: Option<u8>,
}

/// Raw column data from Excel worksheet
#[derive(Debug, Clone, Default)]
pub struct ColumnWidths {
    /// Column definitions in document order
    pub column_definitions: Vec<ColumnDefinition>,
    /// Sheet format properties
    pub sheet_format: SheetFormatProperties,
}

impl ColumnWidths {
    /// Create new empty container
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a column definition
    pub fn add_column_definition(&mut self, def: ColumnDefinition) {
        self.column_definitions.push(def);
    }

    /// Find column definitions that include a specific column (1-based)
    pub fn find_definitions_for_column(&self, col_index: u32) -> Vec<&ColumnDefinition> {
        self.column_definitions
            .iter()
            .filter(|def| col_index >= def.min && col_index <= def.max)
            .collect()
    }
}

/// Raw row definition from Excel XML
#[derive(Debug, Clone)]
pub struct RowDefinition {
    /// r attribute: Row index (1-based)
    pub r: u32,
    /// ht attribute: Row height as a double
    pub height: Option<f64>,
    /// s attribute: Style index (0-65429)
    pub style: Option<u32>,
    /// customHeight attribute
    pub custom_height: Option<bool>,
    /// hidden attribute
    pub hidden: Option<bool>,
    /// outlineLevel attribute (0-7)
    pub outline_level: Option<u8>,
    /// collapsed attribute
    pub collapsed: Option<bool>,
    /// thickTop attribute
    pub thick_top: Option<bool>,
    /// thickBot attribute
    pub thick_bot: Option<bool>,
}

/// Raw row data from Excel worksheet
#[derive(Debug, Clone, Default)]
pub struct RowDefinitions {
    /// Row definitions in document order
    pub row_definitions: Vec<RowDefinition>,
    /// Sheet format properties (shared with columns)
    pub sheet_format: SheetFormatProperties,
}

impl RowDefinitions {
    /// Create new empty container
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a row definition
    pub fn add_row_definition(&mut self, def: RowDefinition) {
        self.row_definitions.push(def);
    }

    /// Find row definition for a specific row (1-based)
    pub fn find_definition_for_row(&self, row_index: u32) -> Option<&RowDefinition> {
        self.row_definitions.iter().find(|def| def.r == row_index)
    }
}

/// Utility functions for Excel column width conversions
#[cfg(test)]
pub mod utils {
    /// Apply Excel default logic to get effective column width
    /// Returns width in Excel's character units
    pub fn get_effective_width(
        column_width: Option<f64>,
        default_col_width: Option<f64>,
        base_col_width: Option<u8>,
    ) -> f64 {
        if let Some(width) = column_width {
            return width;
        }

        if let Some(default) = default_col_width {
            return default;
        }

        if let Some(base) = base_col_width {
            // Excel's formula for default width from base
            return base as f64 + 5.0 / 7.0;
        }

        // Excel's ultimate default
        8.43
    }

    /// Convert character units to pixels  
    /// mdw: Maximum digit width in pixels
    pub fn character_units_to_pixels(width: f64, mdw: f64) -> u32 {
        ((width * mdw + 0.5).floor() + 5.0) as u32
    }

    /// Convert pixels to character units using Excel's formula
    /// mdw: Maximum digit width in pixels
    pub fn pixels_to_character_units(pixels: u32, mdw: f64) -> f64 {
        // Formula from MS docs: =Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
        ((pixels as f64 - 5.0) / mdw * 100.0 + 0.5).trunc() / 100.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_raw_storage() {
        let mut widths = ColumnWidths::new();
        widths.add_column_definition(ColumnDefinition {
            min: 1,
            max: 3,
            width: Some(10.5),
            style: Some(1),
            custom_width: Some(true),
            best_fit: Some(false),
            hidden: None,
            outline_level: None,
            collapsed: None,
        });

        let defs = widths.find_definitions_for_column(2);
        assert_eq!(defs.len(), 1);
        assert_eq!(defs[0].width.unwrap(), 10.5);
    }

    #[test]
    fn test_utility_functions() {
        // Test effective width
        assert_eq!(utils::get_effective_width(Some(10.5), None, None), 10.5);
        assert_eq!(utils::get_effective_width(None, Some(9.0), None), 9.0);
        assert_eq!(utils::get_effective_width(None, None, None), 8.43);

        // Test pixel conversions
        assert_eq!(utils::character_units_to_pixels(8.0, 7.0), 61);
        // Test pixel to character conversion
        assert_eq!(utils::pixels_to_character_units(61, 7.0), 8.0);
    }
}
