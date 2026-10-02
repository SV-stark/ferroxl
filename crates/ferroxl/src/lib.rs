//! ferroxl — a Rust port of the `openpyxl` library for reading and writing Excel 2007
//! `.xlsx`/`.xlsm` files.
//!
//! The crate mirrors the Python package's module layout and semantics: cells
//! ([`cell`]), styles ([`styles`]), worksheets ([`worksheet`]), the workbook
//! ([`workbook`]), the reader ([`reader`]), the writer ([`writer`]) and the supporting
//! utilities ([`units`], [`date_time`], [`xml`]).
//!
//! # Reading
//!
//! ```no_run
//! use ferroxl::{load_workbook, LoadOptions};
//!
//! let workbook = load_workbook("report.xlsx", LoadOptions::default())?;
//!
//! for name in workbook.get_sheet_names() {
//!     println!("{name}");
//! }
//!
//! let sheet = workbook.active_sheet()?;
//! for row in sheet.range_values("A1:D10")? {
//!     for value in row {
//!         println!("{value:?}");
//!     }
//! }
//!
//! // One cell at a time, when a coordinate is all you have.
//! println!("{:?}", sheet.cell_value("B2"));
//! # Ok::<(), ferroxl::Error>(())
//! ```
//!
//! # Writing
//!
//! ```no_run
//! use ferroxl::{CellValue, Style, Workbook};
//!
//! let mut workbook = Workbook::new();
//!
//! // `Workbook::new()` already made `Sheet1`, and `create_sheet` appends without
//! // making the new sheet active, so take the index it returns.
//! let summary = workbook.create_sheet(Some("Summary"))?;
//!
//! let sheet = &mut workbook.worksheets[summary];
//! sheet.set("A1", CellValue::text("Item"))?;
//! sheet.set("B1", CellValue::text("Revenue"))?;
//! sheet.set("A2", CellValue::text("Widget"))?;
//! sheet.set("B2", CellValue::Number(1200.0))?;
//! sheet.set("C2", CellValue::Formula("=B2*1.2".to_string()))?;
//!
//! let mut header = Style::new();
//! header.font.bold = true;
//! header.fill.fill_type = Some("solid".to_string());
//! header.fill.start_color = ferroxl::Color::new("FFDDDDDD".to_string());
//! for cell in ["A1", "B1", "C1"] {
//!     sheet.set_style(cell, header.clone())?;
//! }
//! sheet.set_freeze_panes("A2");
//!
//! workbook.save("summary.xlsx")?;
//! # Ok::<(), ferroxl::Error>(())
//! ```
//!
//! # Feature parity
//!
//! Every public Python class and function has a Rust counterpart with the same behaviour.
//! Where Python's semantics are impossible or undesirable to reproduce exactly, the
//! deviation is documented at the call site and the closest faithful behaviour is
//! implemented instead — the places where this happens are the `HashableObject` equality
//! rules, tri-state booleans, floating-point formatting and PIL-dependent image handling.

#![deny(missing_docs)]

pub mod cell;
pub mod charts;
pub mod comments;
pub mod datavalidation;
pub mod date_time;
pub mod drawing;
pub mod exceptions;
pub mod formatting;
pub mod formula;
pub mod namedrange;
pub mod reader;
pub mod styles;
pub mod units;
pub mod workbook;
pub mod worksheet;
pub mod writer;
pub mod xml;

pub use cell::{
    absolute_coordinate, column_index_from_string, coordinate_from_string, get_column_letter, Cell,
    CellValue, DataType,
};
pub use comments::Comment;
pub use date_time::BaseDate;
pub use exceptions::{Error, Result};
pub use namedrange::{DefinedName, NamedRange, NamedRangeContainingValue};
pub use reader::{load_workbook, load_workbook_from_bytes, LoadOptions};
pub use styles::{Alignment, Border, Borders, Color, Fill, Font, NumberFormat, Protection, Style};
pub use workbook::{DocumentProperties, DocumentSecurity, Workbook};
pub use worksheet::{AutoFilter, HeaderFooter, PageMargins, PageSetup, SheetProtection, Worksheet};
pub use writer::{save_dump, save_workbook, save_workbook_to, DumpWorksheet, ExcelWriter};

/// The ferroxl version, matching the openpyxl release it ports.
pub const VERSION: &str = env!("CARGO_PKG_VERSION");

/// The upstream Python project this crate ports.
pub const UPSTREAM_PROJECT: &str = "openpyxl";

/// The MCP server crate name, for agents that discover tools by package.
pub const MCP_SERVER: &str = "ferroxl-mcp";

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn version_is_exposed() {
        // Compared against the manifest rather than a literal, so bumping the version does
        // not leave a test asserting a major that is two releases out of date.
        assert_eq!(VERSION, env!("CARGO_PKG_VERSION"));
        assert_eq!(UPSTREAM_PROJECT, "openpyxl");
        assert_eq!(MCP_SERVER, "ferroxl-mcp");
    }
}
