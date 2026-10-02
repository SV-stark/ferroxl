//! Reading xlsx files (`openpyxl/reader/`).
//!
//! The reader is organised as one module per package part: `excel.rs` walks the archive,
//! `workbook.rs` handles workbook-level metadata, `worksheet.rs` parses sheets, `style.rs`
//! the shared styles and `strings.rs`/`comments.rs` the remaining parts.

pub mod archive;
pub mod comments;
pub mod excel;
pub mod strings;
pub mod style;
pub mod workbook;
pub mod worksheet;

pub use comments::read_comments;
pub use excel::{
    load_workbook, load_workbook_from_bytes, package_bytes, LoadOptions, WorkbookSource,
};
pub use strings::read_string_table;
pub use style::{read_style_table, StyleTable};
pub use workbook::{
    detect_worksheets, read_content_types, read_excel_base_date, read_named_ranges,
    read_properties_core, read_rels, read_sheets, read_workbook_settings, DetectedSheet,
};
pub use worksheet::{read_worksheet, WorksheetParseContext};
