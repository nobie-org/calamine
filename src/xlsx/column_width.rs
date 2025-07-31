use std::collections::BTreeMap;

/// Column width information for a single column or range of columns
#[derive(Debug, Clone)]
pub struct ColumnWidth {
    /// Width in character units (0-255)
    pub width: u8,
    /// Whether this is a custom width (true if user changed it)
    pub custom_width: bool,
    /// Whether to use best fit width (auto-size)
    pub best_fit: bool,
}

/// Sheet format properties containing default column widths
#[derive(Debug, Clone, Default)]
pub struct SheetFormatProperties {
    /// Default column width in character units
    pub default_col_width: Option<u8>,
    /// Base column width in character units
    pub base_col_width: Option<u8>,
}

/// Container for all column width information in a worksheet
#[derive(Debug, Clone, Default)]
pub struct ColumnWidths {
    /// Column-specific widths (key is column index, 0-based)
    /// Stored as BTreeMap for efficient range queries
    pub columns: BTreeMap<u32, ColumnWidth>,
    /// Sheet format properties with defaults
    pub sheet_format: SheetFormatProperties,
}

impl ColumnWidths {
    /// Create a new empty ColumnWidths container
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a column width for a range of columns
    pub fn add_column_range(&mut self, min_col: u32, max_col: u32, width: ColumnWidth) {
        for col in min_col..=max_col {
            self.columns.insert(col, width.clone());
        }
    }

    /// Get the width for a specific column (0-based index)
    /// Returns the width in character units (0-255)
    pub fn get_column_width(&self, col_index: u32) -> u8 {
        if let Some(col_width) = self.columns.get(&col_index) {
            return col_width.width;
        }

        // Use default width if no specific width is set
        if let Some(default_width) = self.sheet_format.default_col_width {
            default_width
        } else if let Some(base_width) = self.sheet_format.base_col_width {
            // Excel default: base width + 5 pixels (converted to character units)
            // Using the default MDW of 7 pixels for Calibri 11
            base_width.saturating_add(1) // approximately 5/7
        } else {
            // Excel's ultimate default is 8.43 character units
            8
        }
    }

    /// Convert character units to pixels
    /// mdw: Maximum digit width in pixels (typically 7 for Calibri 11 at 96 DPI)
    pub fn character_units_to_pixels(width: u8, mdw: f64) -> u32 {
        // Formula from Excel spec: pixels = FLOOR(width * MDW + 0.5) + 5
        ((width as f64 * mdw + 0.5).floor() + 5.0) as u32
    }

    /// Convert pixels to character units
    /// mdw: Maximum digit width in pixels (typically 7 for Calibri 11 at 96 DPI)
    pub fn pixels_to_character_units(pixels: u32, mdw: f64) -> u8 {
        // Formula: storedWidth = TRUNC(((pixels - 5) / MDW) * 256) / 256
        let raw = ((pixels as f64 - 5.0) / mdw).clamp(0.0, 255.0);
        raw as u8
    }

    /// Get the pixel width for a specific column
    /// mdw: Maximum digit width in pixels (defaults to 7 for Calibri 11)
    pub fn get_column_width_pixels(&self, col_index: u32, mdw: Option<f64>) -> u32 {
        let width = self.get_column_width(col_index);
        let mdw = mdw.unwrap_or(7.0);
        Self::character_units_to_pixels(width, mdw)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_default_width() {
        let widths = ColumnWidths::new();
        assert_eq!(widths.get_column_width(0), 8);
    }

    #[test]
    fn test_custom_width() {
        let mut widths = ColumnWidths::new();
        widths.add_column_range(
            0,
            0,
            ColumnWidth {
                width: 10,
                custom_width: true,
                best_fit: false,
            },
        );
        assert_eq!(widths.get_column_width(0), 10);
        assert_eq!(widths.get_column_width(1), 8);
    }

    #[test]
    fn test_pixel_conversion() {
        // Test Excel default: 8 character units ≈ 61 pixels
        let pixels = ColumnWidths::character_units_to_pixels(8, 7.0);
        assert_eq!(pixels, 61);

        // Test reverse conversion
        let chars = ColumnWidths::pixels_to_character_units(64, 7.0);
        assert_eq!(chars, 8); // 64 pixels converts to 8 character units
    }
}
