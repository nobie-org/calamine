# Pivot Table Implementation Plan for Calamine

## Overview

This document outlines the plan to implement pivot table parsing functionality for the calamine library, focusing initially on the XLSX format only. The implementation will enable users to read pivot table definitions, access pivot cache data, and retrieve the underlying data that pivot tables reference.

## Background

Pivot tables in Excel XLSX files consist of multiple interconnected XML components:
- **Pivot Table Definition** (`xl/pivotTables/pivotTable{n}.xml`): Contains the layout and structure
- **Pivot Cache Definition** (`xl/pivotCache/pivotCacheDefinition{n}.xml`): Describes the data source
- **Pivot Cache Records** (`xl/pivotCache/pivotCacheRecords{n}.xml`): Contains the actual cached data
- **Relationships** (`xl/_rels/*.rels`): Links the components together

## Implementation Phases

### Phase 1: Basic Infrastructure (MVP)

#### 1.1 Data Structures
Create new structs to represent pivot table components:

```rust
// src/pivot.rs
pub struct PivotTable {
    pub name: String,
    pub sheet_name: String,
    pub location: CellReference,
    pub source_range: Option<Range<Data>>,
    pub cache_id: u32,
    pub fields: Vec<PivotField>,
    pub row_fields: Vec<u32>,
    pub column_fields: Vec<u32>,
    pub data_fields: Vec<PivotDataField>,
    pub filters: Vec<PivotFilter>,
}

pub struct PivotField {
    pub name: String,
    pub field_type: PivotFieldType,
    pub items: Vec<String>,
}

pub struct PivotCache {
    pub id: u32,
    pub source_type: PivotSourceType,
    pub source_range: Option<String>,
    pub records: Vec<Vec<Data>>,
    pub fields: Vec<PivotCacheField>,
}
```

#### 1.2 Parser Implementation
- Add pivot table parsing methods to `src/xlsx/mod.rs`:
  - `read_pivot_tables()`: Discovers and indexes pivot tables
  - `read_pivot_cache_definitions()`: Reads cache metadata
  - `read_pivot_cache_records()`: Reads cached data
  - `parse_pivot_table_xml()`: Parses individual pivot table files

#### 1.3 Integration with Reader Trait
- Add methods to the `Reader` trait:
  - `pivot_tables(&self) -> Vec<String>`: List all pivot table names
  - `pivot_table(&mut self, name: &str) -> Result<PivotTable, Error>`: Get specific pivot table
  - `pivot_table_data(&mut self, name: &str) -> Result<Range<Data>, Error>`: Get underlying data

### Phase 2: Enhanced Functionality

#### 2.1 Advanced Features
- Support for calculated fields
- Handling of grouped fields (date grouping, numeric ranges)
- Support for multiple consolidation ranges
- External data source references

#### 2.2 Performance Optimization
- Lazy loading of pivot cache data
- Efficient memory usage for large pivot tables
- Caching of parsed pivot table structures

### Phase 3: Full Feature Parity

#### 3.1 Complex Scenarios
- Pivot table styles and formatting
- Conditional formatting in pivot tables
- Pivot charts linked to pivot tables
- OLAP cube connections

#### 3.2 Validation and Error Handling
- Validation of pivot table integrity
- Graceful handling of corrupted pivot tables
- Clear error messages for unsupported features

## Technical Considerations

### XML Parsing Strategy
- Utilize the existing `quick_xml` infrastructure
- Follow the pattern established by table parsing in `read_table_metadata()`
- Handle namespaces correctly (pivot table XML uses multiple namespaces)

### Memory Management
- Consider streaming approach for large pivot caches
- Implement reference counting for shared pivot caches
- Option to load pivot table structure without loading all data

### API Design Principles
1. **Consistency**: Follow existing calamine API patterns
2. **Simplicity**: Make common operations easy
3. **Performance**: Don't load data unless requested
4. **Flexibility**: Allow access to both high-level and low-level data

## Example Usage

```rust
use calamine::{open_workbook, Reader, Xlsx};

let mut workbook: Xlsx<_> = open_workbook("sales_data.xlsx")?;

// List all pivot tables
let pivot_names = workbook.pivot_tables();

// Get a specific pivot table
let pivot = workbook.pivot_table("SalesSummary")?;

// Access the underlying data
let data = workbook.pivot_table_data("SalesSummary")?;

// Get pivot table layout information
println!("Row fields: {:?}", pivot.row_fields);
println!("Column fields: {:?}", pivot.column_fields);
println!("Data fields: {:?}", pivot.data_fields);
```

## Testing Strategy

### Unit Tests
- Test parsing of individual XML components
- Test data structure construction
- Test error handling for malformed XML

### Integration Tests
- Create test XLSX files with various pivot table configurations:
  - Simple pivot table with one data field
  - Multi-dimensional pivot table
  - Pivot table with filters
  - Pivot table with calculated fields
  - Pivot table referencing external data

### Compatibility Tests
- Test with files created in different Excel versions
- Test with files from other spreadsheet applications (LibreOffice Calc)
- Test edge cases and error conditions

## Implementation Timeline

### Week 1-2: Basic Infrastructure
- Implement data structures
- Basic XML parsing for pivot table definitions
- Simple test cases

### Week 3-4: Core Functionality
- Pivot cache parsing
- Integration with Reader trait
- Basic API methods

### Week 5-6: Testing and Refinement
- Comprehensive test suite
- Performance optimization
- Documentation

### Week 7-8: Advanced Features (if time permits)
- Calculated fields
- Grouped fields
- External data sources

## Open Questions

1. **API Design**: Should pivot tables be exposed as a separate type or integrated into the existing `Table` structure?
2. **Data Loading**: Should pivot cache data be loaded eagerly or lazily?
3. **Formatting**: How much pivot table formatting information should be preserved?
4. **Compatibility**: What minimum Excel version should we target for pivot table support?

## Dependencies

- No new external dependencies required
- Will use existing `quick_xml` for XML parsing
- Will leverage existing error handling infrastructure

## Success Criteria

1. Can successfully parse and read pivot table structure from XLSX files
2. Can access the underlying data referenced by pivot tables
3. Performance is acceptable for large pivot tables (< 1 second for 100k records)
4. API is intuitive and consistent with existing calamine patterns
5. Comprehensive test coverage (> 90%)
6. Clear documentation with examples

## Future Enhancements

1. Support for other formats (XLS, XLSB, ODS)
2. Pivot table modification and creation
3. Pivot table refresh/recalculation
4. Export pivot table to other formats 