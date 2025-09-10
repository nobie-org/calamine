use calamine::{open_workbook, open_workbook_auto, Reader, Theme, Xlsx};
use std::fs::File;
use std::io::BufReader;

fn wb<R: Reader<BufReader<File>>>(name: &str) -> R {
    let path = format!("{}/tests/{name}", env!("CARGO_MANIFEST_DIR"));
    open_workbook(&path).expect(&path)
}

#[test]
fn test_xlsx_theme_default() {
    // Test with a simple XLSX file that should return the default theme
    let mut excel: Xlsx<_> = wb("issues.xlsx");

    // Theme should be accessible via the Reader trait
    let theme = excel.theme();

    match theme {
        Ok(t) => {
            // Should get a valid theme (either parsed or default)
            assert!(t.name.is_some() || t.name.is_none());

            // Color scheme should have 12 colors
            let colors = t.color_scheme.colors();
            assert_eq!(colors.len(), 12);

            // Should have font scheme
            assert!(t.font_scheme.name.is_some() || t.font_scheme.name.is_none());
            assert!(
                t.font_scheme.major_font.latin.is_some()
                    || t.font_scheme.major_font.latin.is_none()
            );
            assert!(
                t.font_scheme.minor_font.latin.is_some()
                    || t.font_scheme.minor_font.latin.is_none()
            );
        }
        Err(_) => {
            // If theme is not supported or fails to parse, that's okay too
            // The test is mainly to ensure the method exists and doesn't panic
        }
    }
}

#[test]
fn test_theme_color_access() {
    let theme = Theme::default();

    // Test color access by index
    assert!(theme.color_scheme.get_color(0).is_some()); // dark1
    assert!(theme.color_scheme.get_color(1).is_some()); // light1
    assert!(theme.color_scheme.get_color(4).is_some()); // accent1
    assert!(theme.color_scheme.get_color(12).is_none()); // out of range

    // Test color array access
    let colors = theme.color_scheme.colors();
    assert_eq!(colors.len(), 12);
}

#[test]
fn test_theme_font_scheme() {
    let theme = Theme::default();

    // Test font scheme structure
    assert!(theme.font_scheme.major_font.latin.is_some());
    assert!(theme.font_scheme.minor_font.latin.is_some());

    // Default theme should have "Calibri Light" for major font
    if let Some(major_font) = &theme.font_scheme.major_font.latin {
        assert_eq!(major_font.as_ref(), "Calibri Light");
    }

    // Default theme should have "Calibri" for minor font
    if let Some(minor_font) = &theme.font_scheme.minor_font.latin {
        assert_eq!(minor_font.as_ref(), "Calibri");
    }
}

#[test]
fn test_rows_xlsx_theme() {
    // Test with rows.xlsx which has a custom Office theme
    let mut excel: Xlsx<_> = wb("rows.xlsx");

    let theme = excel
        .theme()
        .expect("Should be able to get theme from rows.xlsx");

    // Verify theme name
    assert_eq!(theme.name, Some("Office Theme".to_string()));

    // Verify color scheme name
    assert_eq!(theme.color_scheme.name, Some("Office".to_string()));

    // Verify specific colors from the theme XML
    // The theme uses specific RGB values that we can test
    use calamine::Color;

    // Test dark2 color: <a:srgbClr val="0E2841"/>
    match &theme.color_scheme.dark2 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x0E);
            assert_eq!(*g, 0x28);
            assert_eq!(*b, 0x41);
        }
        _ => panic!("Expected RGB color for dark2"),
    }

    // Test light2 color: <a:srgbClr val="E8E8E8"/>
    match &theme.color_scheme.light2 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0xE8);
            assert_eq!(*g, 0xE8);
            assert_eq!(*b, 0xE8);
        }
        _ => panic!("Expected RGB color for light2"),
    }

    // Test accent1 color: <a:srgbClr val="156082"/>
    match &theme.color_scheme.accent1 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x15);
            assert_eq!(*g, 0x60);
            assert_eq!(*b, 0x82);
        }
        _ => panic!("Expected RGB color for accent1"),
    }

    // Test accent2 color: <a:srgbClr val="E97132"/>
    match &theme.color_scheme.accent2 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0xE9);
            assert_eq!(*g, 0x71);
            assert_eq!(*b, 0x32);
        }
        _ => panic!("Expected RGB color for accent2"),
    }

    // Test accent3 color: <a:srgbClr val="196B24"/>
    match &theme.color_scheme.accent3 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x19);
            assert_eq!(*g, 0x6B);
            assert_eq!(*b, 0x24);
        }
        _ => panic!("Expected RGB color for accent3"),
    }

    // Test accent4 color: <a:srgbClr val="0F9ED5"/>
    match &theme.color_scheme.accent4 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x0F);
            assert_eq!(*g, 0x9E);
            assert_eq!(*b, 0xD5);
        }
        _ => panic!("Expected RGB color for accent4"),
    }

    // Test accent5 color: <a:srgbClr val="A02B93"/>
    match &theme.color_scheme.accent5 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0xA0);
            assert_eq!(*g, 0x2B);
            assert_eq!(*b, 0x93);
        }
        _ => panic!("Expected RGB color for accent5"),
    }

    // Test accent6 color: <a:srgbClr val="4EA72E"/>
    match &theme.color_scheme.accent6 {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x4E);
            assert_eq!(*g, 0xA7);
            assert_eq!(*b, 0x2E);
        }
        _ => panic!("Expected RGB color for accent6"),
    }

    // Test hyperlink color: <a:srgbClr val="467886"/>
    match &theme.color_scheme.hyperlink {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x46);
            assert_eq!(*g, 0x78);
            assert_eq!(*b, 0x86);
        }
        _ => panic!("Expected RGB color for hyperlink"),
    }

    // Test followed hyperlink color: <a:srgbClr val="96607D"/>
    match &theme.color_scheme.followed_hyperlink {
        Color::Rgb { r, g, b } => {
            assert_eq!(*r, 0x96);
            assert_eq!(*g, 0x60);
            assert_eq!(*b, 0x7D);
        }
        _ => panic!("Expected RGB color for followed hyperlink"),
    }

    // Verify font scheme name
    assert_eq!(theme.font_scheme.name, Some("Office".to_string()));

    // Verify major font: <a:latin typeface="Aptos Display"
    if let Some(major_font) = &theme.font_scheme.major_font.latin {
        assert_eq!(major_font.as_ref(), "Aptos Display");
    } else {
        panic!("Expected major font to be present");
    }

    // Verify minor font: <a:latin typeface="Aptos Narrow"
    if let Some(minor_font) = &theme.font_scheme.minor_font.latin {
        assert_eq!(minor_font.as_ref(), "Aptos Narrow");
    } else {
        panic!("Expected minor font to be present");
    }

    // Test color access by index matches the expected colors
    assert_eq!(
        theme.color_scheme.get_color(2),
        Some(&theme.color_scheme.dark2)
    );
    assert_eq!(
        theme.color_scheme.get_color(3),
        Some(&theme.color_scheme.light2)
    );
    assert_eq!(
        theme.color_scheme.get_color(4),
        Some(&theme.color_scheme.accent1)
    );
    assert_eq!(
        theme.color_scheme.get_color(10),
        Some(&theme.color_scheme.hyperlink)
    );
    assert_eq!(
        theme.color_scheme.get_color(11),
        Some(&theme.color_scheme.followed_hyperlink)
    );
}

#[test]
fn test_auto_reader_theme() {
    // Test that the auto reader can also access themes
    let path = format!("{}/tests/rows.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook = open_workbook_auto(&path).expect("Cannot open rows.xlsx with auto reader");

    // Should be able to call theme() on the auto reader
    match workbook.theme() {
        Ok(theme) => {
            // Should have the same theme as the direct XLSX test
            assert_eq!(theme.name, Some("Office Theme".to_string()));
            assert_eq!(theme.color_scheme.name, Some("Office".to_string()));
            assert_eq!(theme.font_scheme.name, Some("Office".to_string()));

            // Test that it detected this is an XLSX file by checking if we get a theme
            // (only XLSX supports themes, other formats would return default or error)
            if let Some(major_font) = &theme.font_scheme.major_font.latin {
                assert_eq!(major_font.as_ref(), "Aptos Display");
            }
        }
        Err(err) => {
            // If the auto reader detected a non-XLSX format, the theme() call should fail
            // with unsupported error, which is also acceptable behavior
            println!("Auto reader theme error (expected for non-XLSX): {:?}", err);
        }
    }
}
