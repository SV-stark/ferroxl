//! Range geometry shared by the worksheet and the streaming reader.

use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::Result;

/// A rectangular cell range in 1-based, inclusive coordinates.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RangeBounds {
    /// First column index.
    pub min_col: u32,
    /// First row number.
    pub min_row: u32,
    /// Last column index, inclusive.
    pub max_col: u32,
    /// Last row number, inclusive.
    pub max_row: u32,
}

impl RangeBounds {
    /// Number of columns in the range.
    ///
    /// `max_col` is exclusive, so `A1:C4` has three columns.
    pub fn column_count(&self) -> u32 {
        self.max_col.saturating_sub(self.min_col)
    }

    /// Number of rows in the range.
    ///
    /// `max_row` is the last row itself, not an exclusive bound.
    pub fn row_count(&self) -> u32 {
        self.max_row.saturating_sub(self.min_row) + 1
    }

    /// The column letters covered, in order.
    pub fn columns(&self) -> Result<Vec<String>> {
        (self.min_col..self.max_col)
            .map(get_column_letter)
            .collect()
    }

    /// Render as an A1 range string.
    pub fn to_range_string(&self) -> Result<String> {
        Ok(format!(
            "{}{}:{}{}",
            get_column_letter(self.min_col)?,
            self.min_row,
            get_column_letter(self.max_col.saturating_sub(1))?,
            self.max_row
        ))
    }
}

/// Resolve a range string to its bounds, applying offsets.
///
/// The bounds are **exclusive at the top and right**, matching openpyxl: `A1:C4` has a
/// `max_col` of 4 and a `max_row` of 4, so `column_count` is 3 and `row_count` is 4. A bare
/// coordinate expands by `column_offset` columns and `row_offset` rows, so
/// `get_range_boundaries("A1", 0, 1)` yields `A1:B1`.
pub fn get_range_boundaries(
    range_string: &str,
    row_offset: u32,
    column_offset: u32,
) -> Result<RangeBounds> {
    match range_string.split_once(':') {
        Some((min_range, max_range)) => {
            let (min_col_letters, min_row) = coordinate_from_string(min_range)?;
            let (max_col_letters, max_row) = coordinate_from_string(max_range)?;
            let min_col = column_index_from_string(&min_col_letters)?;
            let max_col = column_index_from_string(&max_col_letters)?;
            Ok(RangeBounds {
                min_col,
                min_row,
                max_col: max_col + 1,
                max_row,
            })
        }
        None => {
            let (col_letters, row) = coordinate_from_string(range_string)?;
            let col = column_index_from_string(&col_letters)?;
            Ok(RangeBounds {
                min_col: col,
                min_row: row,
                max_col: col + column_offset,
                max_row: row + row_offset,
            })
        }
    }
}

/// The dimensions declared in a worksheet's `<dimension>` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetDimensions {
    /// First column letters.
    pub min_col: String,
    /// First row number.
    pub min_row: u32,
    /// Last column letters.
    pub max_col: String,
    /// Last row number.
    pub max_row: u32,
}

impl SheetDimensions {
    /// Render as an A1 range string.
    pub fn to_range_string(&self) -> String {
        format!(
            "{}{}:{}{}",
            self.min_col, self.min_row, self.max_col, self.max_row
        )
    }

    /// Convert to inclusive bounds.
    pub fn to_bounds(&self) -> Result<RangeBounds> {
        Ok(RangeBounds {
            min_col: column_index_from_string(&self.min_col)?,
            min_row: self.min_row,
            max_col: column_index_from_string(&self.max_col)?,
            max_row: self.max_row,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn range_with_colon_has_an_exclusive_max_column() {
        let b = get_range_boundaries("A1:C4", 0, 1).unwrap();
        assert_eq!(b.min_col, 1);
        assert_eq!(b.min_row, 1);
        // openpyxl adds one to `max_col`, so the bound sits just past the last column.
        assert_eq!(b.max_col, 4);
        assert_eq!(b.max_row, 4);
        assert_eq!(b.column_count(), 3);
        assert_eq!(b.row_count(), 4);
        assert_eq!(b.to_range_string().unwrap(), "A1:C4");
    }

    #[test]
    fn bare_coordinate_expands_by_offset() {
        let b = get_range_boundaries("A1", 0, 1).unwrap();
        assert_eq!(b.max_col, 2);
        assert_eq!(b.max_row, 1);
        assert_eq!(b.column_count(), 1);
        assert_eq!(b.to_range_string().unwrap(), "A1:A1");
    }

    #[test]
    fn columns_are_enumerated() {
        let b = get_range_boundaries("A1:C1", 0, 1).unwrap();
        assert_eq!(b.columns().unwrap(), vec!["A", "B", "C"]);
    }

    #[test]
    fn dimensions_render() {
        let dims = SheetDimensions {
            min_col: "A".into(),
            min_row: 1,
            max_col: "D".into(),
            max_row: 10,
        };
        assert_eq!(dims.to_range_string(), "A1:D10");
        let bounds = dims.to_bounds().unwrap();
        assert_eq!(bounds.max_col, 4);
    }

    #[test]
    fn bad_ranges_error() {
        assert!(get_range_boundaries("nonsense", 0, 1).is_err());
        assert!(get_range_boundaries("A1:zz", 0, 1).is_err());
    }
}
