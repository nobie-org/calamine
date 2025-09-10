//! Excel theme support
//!
//! This module provides support for reading and working with Excel themes from XLSX files.
//! Themes define the color schemes, font schemes, and format schemes used throughout
//! a workbook to maintain consistent styling.
//!
//! # Excel Theme Structure
//!
//! Excel themes are defined in the `xl/theme/theme1.xml` file within XLSX archives.
//! The theme contains:
//!
//! - **Color Scheme**: 12 predefined colors (dark1, light1, dark2, light2, accent1-6, hyperlink, followed hyperlink)
//! - **Font Scheme**: Major and minor font definitions
//! - **Format Scheme**: Line, fill, and effect styles
//!
//! # References
//!
//! - ECMA-376 Part 1, Section 20.1 (DrawingML - Theme)
//! - MS-XLSX: Excel (.xlsx) Extensions to the Office Open XML SpreadsheetML File Format

use crate::formats::Color;
use std::sync::Arc;

/// Complete theme information from an Excel workbook
///
/// Contains all theme elements including color scheme, font scheme, and format scheme.
/// This structure provides access to the theme colors and fonts that cells can reference.
///
/// # Examples
///
/// ```
/// use calamine::Theme;
///
/// // Access theme colors by index
/// // let theme = workbook.theme().unwrap();
/// // let accent1_color = theme.color_scheme.accent1;
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct Theme {
    /// Theme name (e.g., "Office Theme")
    pub name: Option<String>,
    /// Color scheme defining the theme's 12 colors
    pub color_scheme: ColorScheme,
    /// Font scheme defining major and minor fonts
    pub font_scheme: FontScheme,
    /// Format scheme defining line, fill, and effect styles
    pub format_scheme: Option<FormatScheme>,
}

/// Theme color scheme containing the 12 standard theme colors
///
/// The color scheme defines the standard set of colors that Excel uses for theming.
/// These colors are referenced by index in cell formatting and other styling elements.
///
/// # Color Indices
///
/// The standard theme color indices are:
/// - 0: dark1 (usually black)
/// - 1: light1 (usually white)  
/// - 2: dark2 (usually dark accent)
/// - 3: light2 (usually light accent)
/// - 4-9: accent1 through accent6
/// - 10: hyperlink (usually blue)
/// - 11: followed hyperlink (usually purple)
#[derive(Debug, Clone, PartialEq)]
pub struct ColorScheme {
    /// Scheme name (e.g., "Office")
    pub name: Option<String>,
    /// Dark color 1 (typically black, used for text)
    pub dark1: Color,
    /// Light color 1 (typically white, used for backgrounds)
    pub light1: Color,
    /// Dark color 2 (secondary dark color)
    pub dark2: Color,
    /// Light color 2 (secondary light color)
    pub light2: Color,
    /// Accent color 1
    pub accent1: Color,
    /// Accent color 2
    pub accent2: Color,
    /// Accent color 3
    pub accent3: Color,
    /// Accent color 4
    pub accent4: Color,
    /// Accent color 5
    pub accent5: Color,
    /// Accent color 6
    pub accent6: Color,
    /// Hyperlink color
    pub hyperlink: Color,
    /// Followed hyperlink color
    pub followed_hyperlink: Color,
}

impl ColorScheme {
    /// Get a theme color by index
    ///
    /// Returns the theme color for the given index, following Excel's
    /// standard theme color indexing scheme.
    ///
    /// # Arguments
    ///
    /// * `index` - The theme color index (0-11)
    ///
    /// # Returns
    ///
    /// The color for the given index, or None if the index is out of range
    ///
    /// # Examples
    ///
    /// ```
    /// use calamine::{ColorScheme, Color};
    ///
    /// let color_scheme = ColorScheme::default();
    /// let dark1 = color_scheme.get_color(0);  // Gets dark1
    /// let accent1 = color_scheme.get_color(4); // Gets accent1
    /// ```
    pub fn get_color(&self, index: u32) -> Option<&Color> {
        match index {
            0 => Some(&self.dark1),
            1 => Some(&self.light1),
            2 => Some(&self.dark2),
            3 => Some(&self.light2),
            4 => Some(&self.accent1),
            5 => Some(&self.accent2),
            6 => Some(&self.accent3),
            7 => Some(&self.accent4),
            8 => Some(&self.accent5),
            9 => Some(&self.accent6),
            10 => Some(&self.hyperlink),
            11 => Some(&self.followed_hyperlink),
            _ => None,
        }
    }

    /// Get all theme colors as a slice
    ///
    /// Returns all 12 theme colors in index order.
    pub fn colors(&self) -> [&Color; 12] {
        [
            &self.dark1,
            &self.light1,
            &self.dark2,
            &self.light2,
            &self.accent1,
            &self.accent2,
            &self.accent3,
            &self.accent4,
            &self.accent5,
            &self.accent6,
            &self.hyperlink,
            &self.followed_hyperlink,
        ]
    }
}

impl Default for ColorScheme {
    /// Create a default Office theme color scheme
    ///
    /// Uses the standard Office theme colors that match Excel's default theme.
    fn default() -> Self {
        Self {
            name: Some("Office".to_string()),
            dark1: Color::Rgb { r: 0, g: 0, b: 0 }, // Black
            light1: Color::Rgb {
                r: 255,
                g: 255,
                b: 255,
            }, // White
            dark2: Color::Rgb {
                r: 68,
                g: 84,
                b: 106,
            }, // Dark blue-gray
            light2: Color::Rgb {
                r: 238,
                g: 236,
                b: 225,
            }, // Light gray
            accent1: Color::Rgb {
                r: 68,
                g: 114,
                b: 196,
            }, // Blue
            accent2: Color::Rgb {
                r: 237,
                g: 125,
                b: 49,
            }, // Orange
            accent3: Color::Rgb {
                r: 165,
                g: 165,
                b: 165,
            }, // Gray
            accent4: Color::Rgb {
                r: 255,
                g: 192,
                b: 0,
            }, // Gold
            accent5: Color::Rgb {
                r: 91,
                g: 155,
                b: 213,
            }, // Light blue
            accent6: Color::Rgb {
                r: 112,
                g: 173,
                b: 71,
            }, // Green
            hyperlink: Color::Rgb {
                r: 5,
                g: 99,
                b: 193,
            }, // Blue
            followed_hyperlink: Color::Rgb {
                r: 149,
                g: 79,
                b: 114,
            }, // Purple
        }
    }
}

/// Theme font scheme defining major and minor fonts
///
/// The font scheme defines the typefaces used for different text elements
/// in the theme. Major fonts are typically used for headings, while minor
/// fonts are used for body text.
#[derive(Debug, Clone, PartialEq)]
pub struct FontScheme {
    /// Scheme name (e.g., "Office")
    pub name: Option<String>,
    /// Major font info (typically for headings)
    pub major_font: ThemeFont,
    /// Minor font info (typically for body text)
    pub minor_font: ThemeFont,
}

impl Default for FontScheme {
    /// Create a default Office theme font scheme
    fn default() -> Self {
        Self {
            name: Some("Office".to_string()),
            major_font: ThemeFont {
                latin: Some(Arc::from("Calibri Light")),
                east_asian: None,
                complex_script: None,
            },
            minor_font: ThemeFont {
                latin: Some(Arc::from("Calibri")),
                east_asian: None,
                complex_script: None,
            },
        }
    }
}

/// Font information for different script types
///
/// Defines font faces for different script types (Latin, East Asian, Complex Script)
/// to support international text rendering.
#[derive(Debug, Clone, PartialEq)]
pub struct ThemeFont {
    /// Latin script font (e.g., "Calibri")
    pub latin: Option<Arc<str>>,
    /// East Asian script font
    pub east_asian: Option<Arc<str>>,
    /// Complex script font (e.g., Arabic, Hebrew)
    pub complex_script: Option<Arc<str>>,
}

/// Format scheme defining line, fill, and effect styles
///
/// Contains the style definitions for various graphical elements.
/// This is typically used for charts and other graphical elements,
/// though it can also influence cell styling.
#[derive(Debug, Clone, PartialEq)]
pub struct FormatScheme {
    /// Scheme name
    pub name: Option<String>,
    /// Fill style definitions
    pub fill_styles: Vec<FillStyle>,
    /// Line style definitions  
    pub line_styles: Vec<LineStyle>,
    /// Effect style definitions
    pub effect_styles: Vec<EffectStyle>,
}

/// Fill style definition
#[derive(Debug, Clone, PartialEq)]
pub struct FillStyle {
    /// Style name or identifier
    pub name: Option<String>,
    /// Fill color
    pub color: Option<Color>,
    /// Fill pattern or gradient information
    pub pattern: Option<String>,
}

/// Line style definition
#[derive(Debug, Clone, PartialEq)]
pub struct LineStyle {
    /// Style name or identifier
    pub name: Option<String>,
    /// Line color
    pub color: Option<Color>,
    /// Line width in points
    pub width: Option<f64>,
    /// Line dash pattern
    pub dash: Option<String>,
}

/// Effect style definition
#[derive(Debug, Clone, PartialEq)]
pub struct EffectStyle {
    /// Style name or identifier
    pub name: Option<String>,
    /// Effect type (shadow, glow, etc.)
    pub effect_type: Option<String>,
    /// Effect parameters
    pub parameters: Option<String>,
}

impl Default for Theme {
    /// Create a default Office theme
    fn default() -> Self {
        Self {
            name: Some("Office Theme".to_string()),
            color_scheme: ColorScheme::default(),
            font_scheme: FontScheme::default(),
            format_scheme: None,
        }
    }
}
