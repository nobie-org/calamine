//! Tests for pivot table parsing functionality

use calamine::{open_workbook, Xlsx};

#[test]
fn test_load_pivot_tables() {
    let path = format!("{}/tests/pivot_test.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Get pivot table names
    let pivot_names = workbook.pivot_table_names();
    println!("Found {} pivot tables", pivot_names.len());

    for name in &pivot_names {
        println!("Pivot table: {}", name);
    }
}

#[test]
fn test_get_pivot_table() {
    let path = format!("{}/tests/pivot_test.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Get a specific pivot table
    let pivot = workbook
        .pivot_table_by_name("SalesSummary")
        .expect("Pivot table not found");
    println!("Pivot table name: {}", pivot.name);
    println!("Sheet: {}", pivot.sheet_name);
    println!("Location: {:?}", pivot.location);
    println!("Source range: {:?}", pivot.source_range);
    println!("Source sheet: {:?}", pivot.source_sheet);
    println!("Number of fields: {}", pivot.fields.len());
    println!("Fields: {:#?}", pivot.fields);

    assert_eq!(pivot.name, "SalesSummary");
    assert_eq!(pivot.sheet_name, "Reports");
    assert_eq!(pivot.source_range, Some("A1:D5".to_string()));
    assert_eq!(pivot.source_sheet, Some("Sales".to_string()));
}

#[test]
fn test_pivot_tables_on_regular_file() {
    // Test that pivot table methods work gracefully on files without pivot tables
    let path = format!("{}/tests/temperature.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables should succeed even if there are none
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Should return empty list
    let pivot_names = workbook.pivot_table_names();
    assert_eq!(pivot_names.len(), 0);
}
