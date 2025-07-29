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

    println!("Pivot table: {:#?}", pivot);

    // Check renamed items in fields
    for (i, field) in pivot.fields.iter().enumerate() {
        println!(
            "Field {}: {} - items: {:?}",
            i,
            field.name,
            field
                .items
                .iter()
                .map(|item| &item.value)
                .collect::<Vec<_>>()
        );
    }

    assert_eq!(pivot.name, "SalesSummary");
    assert_eq!(pivot.sheet_name, "Reports");
    assert_eq!(pivot.source_range, Some("A1:D5".to_string()));
    assert_eq!(pivot.source_sheet, Some("Sales".to_string()));

    // Check that the Product field has renamed Orange item
    if let Some(product_field) = pivot.fields.iter().find(|f| f.name == "Product") {
        assert!(product_field.items.iter().any(|item|
            item.value == "Orange" && item.custom_name == Some("Renamed Orange".to_string())),
                "Product field should contain item with value 'Orange' and custom_name 'Renamed Orange'");
    }
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

#[test]
fn test_pivot_cache_with_records() {
    let path = format!("{}/tests/pivot_test.xlsx", env!("CARGO_MANIFEST_DIR"));
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Get a pivot table first to find its cache ID
    let pivot = workbook
        .pivot_table_by_name("SalesSummary")
        .expect("Pivot table not found");

    let cache_id = pivot.cache_id;
    println!("Getting cache for ID: {}", cache_id);

    // Now get the cache with records
    let cache = workbook
        .pivot_cache_with_records(cache_id)
        .expect("Failed to get pivot cache");

    println!("Cache: {:#?}", cache);
    println!("Cache ID: {}", cache.id);
    println!("Source type: {:?}", cache.source_type);
    println!("Source range: {:?}", cache.source_range);
    println!("Source sheet: {:?}", cache.source_sheet);
    println!("Number of fields: {}", cache.fields.len());

    // Print field info including shared items
    for (i, field) in cache.fields.iter().enumerate() {
        println!(
            "Field {}: {} - shared items: {:?}",
            i, field.name, field.shared_items
        );
    }

    // Check if records were loaded
    assert!(cache.records.is_some(), "Cache records should be loaded");

    if let Some(records) = &cache.records {
        println!("Number of records: {}", records.len());

        // Print first few records
        for (i, record) in records.iter().take(3).enumerate() {
            println!("Record {}: {:?}", i, record);
        }
    }

    // Verify cache metadata
    assert_eq!(cache.source_range, Some("A1:D5".to_string()));
    assert_eq!(cache.source_sheet, Some("Sales".to_string()));
}

#[test]
fn test_pivot_data_on_columns() {
    let path = format!(
        "{}/tests/pivot_data_on_columns.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Get a pivot table first to find its cache ID
    let pivot = workbook
        .pivot_table_by_name("PivotTable2")
        .expect("Pivot table not found");

    println!("Pivot table: {:#?}", pivot);
}
#[test]
fn test_pivot_data_on_rows() {
    let path = format!(
        "{}/tests/pivot_data_on_rows.xlsx",
        env!("CARGO_MANIFEST_DIR")
    );
    let mut workbook: Xlsx<_> = open_workbook(&path).expect("Cannot open file");

    // Load pivot tables
    workbook
        .load_pivot_tables()
        .expect("Failed to load pivot tables");

    // Get a pivot table first to find its cache ID
    let pivot = workbook
        .pivot_table_by_name("PivotTable2")
        .expect("Pivot table not found");

    println!("Pivot table: {:#?}", pivot);
}
