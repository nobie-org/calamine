//! Example of reading pivot table information from Excel files

use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Open an Excel file
    let path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "tests/pivot_test.xlsx".to_string());

    println!("Opening file: {}", path);
    let mut workbook: Xlsx<_> = open_workbook(&path)?;

    // Load pivot table metadata
    println!("\nLoading pivot tables...");
    workbook.load_pivot_tables()?;

    // Get all pivot table names
    let pivot_names = workbook.pivot_table_names();
    println!("\nFound {} pivot tables", pivot_names.len());

    if pivot_names.is_empty() {
        println!("No pivot tables found in this workbook.");
        return Ok(());
    }

    // List all pivot tables
    println!("\nPivot tables in workbook:");
    for (i, name) in pivot_names.iter().enumerate() {
        println!("  {}. {}", i + 1, name);
    }

    // Clone names to avoid borrow checker issues
    let pivot_names_owned: Vec<String> = pivot_names.iter().map(|&s| s.to_string()).collect();

    // Get details for each pivot table
    for name in &pivot_names_owned {
        println!("\n--- Pivot Table: {} ---", name);

        match workbook.pivot_table_by_name(name) {
            Ok(pivot) => {
                println!("  Sheet: {}", pivot.sheet_name);
                println!(
                    "  Location: row {}, column {} to row {}, column {}",
                    pivot.location.start.0, pivot.location.start.1,
                    pivot.location.end.0, pivot.location.end.1
                );

                if let Some(ref source) = pivot.source_range {
                    println!("  Source range: {}", source);
                    if let Some(ref sheet) = pivot.source_sheet {
                        println!("  Source sheet: {}", sheet);
                    }
                } else {
                    println!("  Source range: <external or unknown>");
                }

                println!("  Cache ID: {}", pivot.cache_id);
                println!("  Number of fields: {}", pivot.fields.len());

                // Show field names
                if !pivot.fields.is_empty() {
                    println!("  Fields:");
                    for (i, field) in pivot.fields.iter().enumerate() {
                        println!("    {}. {}", i, field.name);
                    }
                }

                // Show row fields
                if !pivot.row_fields.is_empty() {
                    println!("  Row fields: {:?}", pivot.row_fields);
                }

                // Show column fields
                if !pivot.column_fields.is_empty() {
                    println!("  Column fields: {:?}", pivot.column_fields);
                }

                // Show data fields
                if !pivot.data_fields.is_empty() {
                    println!("  Data fields:");
                    for data_field in &pivot.data_fields {
                        let display_name =
                            data_field.display_name.as_ref().unwrap_or(&data_field.name);
                        println!(
                            "    - {} (aggregation: {:?})",
                            display_name, data_field.aggregation
                        );
                    }
                }

                // Show filters
                if !pivot.filters.is_empty() {
                    println!("  Filters: {} filter(s) defined", pivot.filters.len());
                }
            }
            Err(e) => {
                println!("  Error reading pivot table: {:?}", e);
            }
        }
    }

    Ok(())
}
