//! Cell handling: values, types, coordinates, formulas and read-only cells.

// `cell::cell` and `worksheet::worksheet` keep the upstream file names so a change can be
// traced to the Python module it mirrors.
#![allow(clippy::module_inception)]

pub mod cell;
pub mod formula;
pub mod read_only;
pub mod utils;

pub use cell::{
    check_string, Cell, CellContext, CellValue, DataType, FormulaAttributes, ERROR_CODES,
    MAX_STRING_LENGTH,
};
pub use formula::{shift_references, FormulaStore, SharedFormula};
pub use read_only::{empty_cell, ReadOnlyCell, ReadOnlyTables};
pub use utils::{
    absolute_coordinate, column_index_from_string, coordinate_from_index, coordinate_from_string,
    get_column_letter, MAX_COLUMN_INDEX,
};
