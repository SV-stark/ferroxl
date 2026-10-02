//! Worksheet parts: dimensions, filters, headers/footers, page setup and protection.

// `worksheet::worksheet` keeps the upstream file name so a change can be traced to the
// Python module it mirrors.
#![allow(clippy::module_inception)]

pub mod cell_range;
pub mod dependency;
pub mod dimensions;
pub mod filters;
pub mod header_footer;
pub mod iter_worksheet;
pub mod page;
pub mod password_hasher;
pub mod protection;
pub mod relationship;
pub mod table;
pub mod worksheet;

pub use cell_range::{CellRange, MultiCellRange};
pub use dependency::References;
pub use dimensions::{ColumnDimension, Dimension, RowDimension};
pub use filters::{AutoFilter, FilterColumn, SortCondition};
pub use header_footer::{HeaderFooter, HeaderFooterItem};
pub use iter_worksheet::{get_range_boundaries, RangeBounds, SheetDimensions};
pub use page::{PageMargins, PageSetup};
pub use password_hasher::hash_password;
pub use protection::SheetProtection;
pub use relationship::{Relationship, RelationshipType};
pub use table::{Table, TableColumn, TableList, TableStyleInfo};
pub use worksheet::Worksheet;
