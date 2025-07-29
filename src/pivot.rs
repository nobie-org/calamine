//! Pivot table structures and parsing functionality

use crate::Data;
use std::collections::HashMap;

/// Represents a pivot table in a spreadsheet
#[derive(Debug, Clone)]
pub struct PivotTable {
    /// The name of the pivot table
    pub name: String,
    /// The sheet containing the pivot table
    pub sheet_name: String,
    /// The location of the pivot table (top-left cell)
    pub location: (u32, u32),
    /// The source data range (if internal to the workbook)
    pub source_range: Option<String>,
    /// The source sheet name (if internal to the workbook)
    pub source_sheet: Option<String>,
    /// The ID of the pivot cache containing the data
    pub cache_id: u32,
    /// All fields defined in the pivot table
    pub fields: Vec<PivotField>,
    /// Indices of fields used as row fields
    pub row_fields: Vec<u32>,
    /// Indices of fields used as column fields
    pub column_fields: Vec<u32>,
    /// Data fields (values being aggregated)
    pub data_fields: Vec<PivotDataField>,
    /// Page/report filters
    pub filters: Vec<PivotFilter>,
}

/// Represents a field in a pivot table
#[derive(Debug, Clone)]
pub struct PivotField {
    /// Field name
    pub name: String,
    /// Field type
    pub field_type: PivotFieldType,
    /// Field items (unique values)
    pub items: Vec<String>,
    /// Field index in the cache
    pub cache_index: Option<u32>,
}

/// Type of pivot field
#[derive(Debug, Clone, PartialEq)]
pub enum PivotFieldType {
    /// Row field
    Row,
    /// Column field
    Column,
    /// Page/filter field
    Page,
    /// Data field
    Data,
    /// Hidden field
    Hidden,
}

/// Represents a data field in a pivot table (values being aggregated)
#[derive(Debug, Clone)]
pub struct PivotDataField {
    /// Name of the field
    pub name: String,
    /// Source field index
    pub field_index: u32,
    /// Aggregation function
    pub aggregation: AggregationFunction,
    /// Custom name for the data field
    pub display_name: Option<String>,
}

/// Aggregation functions for pivot table data fields
#[derive(Debug, Clone, PartialEq)]
pub enum AggregationFunction {
    /// Sum of values
    Sum,
    /// Count of values
    Count,
    /// Average of values
    Average,
    /// Maximum value
    Max,
    /// Minimum value
    Min,
    /// Product of values
    Product,
    /// Count of numeric values
    CountNums,
    /// Standard deviation (sample)
    StdDev,
    /// Standard deviation (population)
    StdDevP,
    /// Variance (sample)
    Var,
    /// Variance (population)
    VarP,
}

impl AggregationFunction {
    /// Parse aggregation function from string
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "sum" => Some(Self::Sum),
            "count" => Some(Self::Count),
            "average" | "avg" => Some(Self::Average),
            "max" => Some(Self::Max),
            "min" => Some(Self::Min),
            "product" => Some(Self::Product),
            "countNums" => Some(Self::CountNums),
            "stdDev" => Some(Self::StdDev),
            "stdDevp" => Some(Self::StdDevP),
            "var" => Some(Self::Var),
            "varp" => Some(Self::VarP),
            _ => None,
        }
    }
}

/// Represents a filter in a pivot table
#[derive(Debug, Clone)]
pub struct PivotFilter {
    /// Field index being filtered
    pub field_index: u32,
    /// Filter type
    pub filter_type: PivotFilterType,
    /// Filter values
    pub values: Vec<String>,
}

/// Type of pivot table filter
#[derive(Debug, Clone, PartialEq)]
pub enum PivotFilterType {
    /// Manual filter (specific values selected)
    Manual,
    /// Label filter
    Label,
    /// Value filter
    Value,
    /// Date filter
    Date,
}

/// Represents a pivot cache containing the source data
#[derive(Debug, Clone)]
pub struct PivotCache {
    /// Cache ID
    pub id: u32,
    /// Source type
    pub source_type: PivotSourceType,
    /// Source range (for worksheet sources)
    pub source_range: Option<String>,
    /// Source sheet name (for worksheet sources)
    pub source_sheet: Option<String>,
    /// Cache fields definition
    pub fields: Vec<PivotCacheField>,
    /// Whether the cache has records
    pub has_records: bool,
}

/// Type of pivot table data source
#[derive(Debug, Clone, PartialEq)]
pub enum PivotSourceType {
    /// Worksheet range
    Worksheet,
    /// External data source
    External,
    /// Consolidation
    Consolidation,
    /// Scenario
    Scenario,
}

/// Represents a field in the pivot cache
#[derive(Debug, Clone)]
pub struct PivotCacheField {
    /// Field name
    pub name: String,
    /// Field data type
    pub data_type: PivotFieldDataType,
    /// Shared items (unique values in the field)
    pub shared_items: Vec<Data>,
}

/// Data type of a pivot cache field
#[derive(Debug, Clone, PartialEq)]
pub enum PivotFieldDataType {
    /// String data type
    String,
    /// Numeric data type
    Number,
    /// Boolean data type
    Boolean,
    /// Date data type
    Date,
    /// Error value
    Error,
    /// Mixed data types
    Mixed,
}

/// Metadata about pivot tables in a workbook
#[derive(Debug, Default)]
pub struct PivotTableCollection {
    /// Map of sheet name to pivot table definitions
    pub tables_by_sheet: HashMap<String, Vec<PivotTableInfo>>,
    /// Map of cache ID to pivot cache
    pub caches: HashMap<u32, PivotCache>,
}

/// Basic information about a pivot table
#[derive(Debug, Clone)]
pub struct PivotTableInfo {
    /// Pivot table name
    pub name: String,
    /// File path to the pivot table definition
    pub path: String,
    /// Associated cache ID
    pub cache_id: Option<u32>,
}

impl PivotTableCollection {
    /// Create a new empty collection
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a pivot table to the collection
    pub fn add_table(&mut self, sheet_name: String, info: PivotTableInfo) {
        self.tables_by_sheet
            .entry(sheet_name)
            .or_insert_with(Vec::new)
            .push(info);
    }

    /// Add a pivot cache to the collection
    pub fn add_cache(&mut self, cache: PivotCache) {
        self.caches.insert(cache.id, cache);
    }

    /// Get all pivot table names
    pub fn table_names(&self) -> Vec<&str> {
        self.tables_by_sheet
            .values()
            .flat_map(|tables| tables.iter().map(|t| t.name.as_str()))
            .collect()
    }

    /// Get pivot tables in a specific sheet
    pub fn tables_in_sheet(&self, sheet_name: &str) -> Option<&[PivotTableInfo]> {
        self.tables_by_sheet.get(sheet_name).map(|v| v.as_slice())
    }
}
