//! Rectangular ranges as values (`openpyxl/worksheet/cell_range.py`).
//!
//! A `CellRange` is a rectangle — the corners of it, not its contents. `A1:C3` is six
//! cells but it is one value, and everything in this module is about that value: does it
//! contain this coordinate, is it inside that other range, what is left after intersecting
//! two of them, where does it move to when a row is inserted above it.
//!
//! `RangeBounds` in [`crate::worksheet::iter_worksheet`] is the geometry — where the corners
//! are. This is the algebra on top: the operations you can perform on two rectangles. The
//! distinction matters because the geometry is needed to iterate a sheet, while the algebra
//! is needed to reason about which parts of it overlap.
//!
//! ```
//! use ferroxl::worksheet::cell_range::CellRange;
//!
//! let whole = CellRange::parse("A1:C10").unwrap();
//! let top = CellRange::parse("A1:B2").unwrap();
//!
//! assert_eq!(whole.intersection(&top).unwrap().to_string(), "A1:B2");
//! assert_eq!(top.shift(1, 0).unwrap().to_string(), "B1:C2");
//! assert!(top.is_subset(&whole));
//! ```

use std::collections::BTreeSet;
use std::fmt;

use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::{Error, Result};

/// The largest column Excel has.
const MAX_COLUMN: u32 = 16_384;
/// The largest row Excel has.
const MAX_ROW: u32 = 1_048_576;

/// A rectangular range of cells, stored as its corners.
///
/// Bounds are 1-based and inclusive at both ends, matching how a caller reads `A1:C3`.
/// This is the opposite of [`RangeBounds`](crate::worksheet::iter_worksheet::RangeBounds), whose `max_col` is
/// exclusive because it
/// comes out of a range string, and mixing the two is the single easiest way to get an
/// off-by-one here.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct CellRange {
    /// The leftmost column, 1-based.
    pub min_col: u32,
    /// The topmost row, 1-based.
    pub min_row: u32,
    /// The rightmost column, inclusive.
    pub max_col: u32,
    /// The bottommost row, inclusive.
    pub max_row: u32,
}

impl CellRange {
    /// Build a range from its four corners.
    ///
    /// Returns [`None`] if the range is empty or inverted. An inverted range is not
    /// normalised silently: `A5:B2` is a mistake, and quietly reading it as `B2:A5` would
    /// turn a typo into a wrong answer.
    pub fn new(min_col: u32, min_row: u32, max_col: u32, max_row: u32) -> Option<Self> {
        if min_col == 0 || min_row == 0 || max_col == 0 || max_row == 0 {
            return None;
        }
        if min_col > max_col || min_row > max_row {
            return None;
        }
        Some(CellRange {
            min_col,
            min_row,
            max_col,
            max_row,
        })
    }

    /// Parse an A1 range string such as `A1:C3` or a bare coordinate such as `B2`.
    ///
    /// A bare coordinate is a one-cell range, which is what `ws["A1"]` means.
    pub fn parse(range_string: &str) -> Result<Self> {
        let cleaned = range_string.trim().trim_start_matches('=');
        match cleaned.split_once(':') {
            Some((from, to)) => {
                let (min_col, min_row) = corners(from)?;
                let (max_col, max_row) = corners(to)?;
                CellRange::new(min_col, min_row, max_col, max_row)
                    .ok_or_else(|| Error::InsufficientCoordinates(range_string.to_string()))
            }
            None => {
                let (col, row) = corners(cleaned)?;
                CellRange::new(col, row, col, row)
                    .ok_or_else(|| Error::InsufficientCoordinates(range_string.to_string()))
            }
        }
    }

    /// The top row.
    pub fn top(&self) -> u32 {
        self.min_row
    }

    /// The bottom row.
    pub fn bottom(&self) -> u32 {
        self.max_row
    }

    /// The leftmost column, as a 1-based index.
    pub fn left(&self) -> u32 {
        self.min_col
    }

    /// The rightmost column, as a 1-based index.
    pub fn right(&self) -> u32 {
        self.max_col
    }

    /// How many cells the range covers.
    pub fn size(&self) -> u64 {
        (self.max_col - self.min_col + 1) as u64 * (self.max_row - self.min_row + 1) as u64
    }

    /// The coordinates the range covers, row by row.
    ///
    /// Bounded by [`ITERATION_LIMIT`] cells. A whole-column range would otherwise
    /// be sixteen thousand wide and a whole-sheet range over sixteen million, which is not a
    /// useful return value and would exhaust memory rather than fail.
    pub fn cells(&self) -> Result<Vec<String>> {
        if self.size() > ITERATION_LIMIT {
            return Err(Error::Value(format!(
                "the range covers {} cells; raise CellRange::ITERATION_LIMIT or narrow it",
                self.size()
            )));
        }
        let mut out = Vec::with_capacity(self.size() as usize);
        for row in self.min_row..=self.max_row {
            for column in self.min_col..=self.max_col {
                out.push(format!("{}{row}", get_column_letter(column)?));
            }
        }
        Ok(out)
    }

    /// The row numbers the range covers.
    pub fn rows(&self) -> Vec<u32> {
        (self.min_row..=self.max_row).collect()
    }

    /// The column indices the range covers.
    pub fn cols(&self) -> Vec<u32> {
        (self.min_col..=self.max_col).collect()
    }

    /// Whether `coordinate` is inside the range.
    ///
    /// A range contains itself, unlike a set.
    pub fn contains(&self, coordinate: &str) -> bool {
        let Ok((letters, row)) = coordinate_from_string(coordinate.trim()) else {
            return false;
        };
        let Ok(column) = column_index_from_string(&letters) else {
            return false;
        };
        (self.min_col..=self.max_col).contains(&column)
            && (self.min_row..=self.max_row).contains(&row)
    }

    /// Whether the two ranges share no cells.
    pub fn isdisjoint(&self, other: &CellRange) -> bool {
        self.intersection(other).is_none()
    }

    /// The overlapping part of two ranges, if any.
    pub fn intersection(&self, other: &CellRange) -> Option<CellRange> {
        CellRange::new(
            self.min_col.max(other.min_col),
            self.min_row.max(other.min_row),
            self.max_col.min(other.max_col),
            self.max_row.min(other.max_row),
        )
    }

    /// The smallest range covering both.
    ///
    /// Cannot fail: taking the minimum and maximum of two valid ranges is itself valid.
    /// Written out rather than going through `new` so the type says so.
    pub fn union(&self, other: &CellRange) -> CellRange {
        CellRange {
            min_col: self.min_col.min(other.min_col),
            min_row: self.min_row.min(other.min_row),
            max_col: self.max_col.max(other.max_col),
            max_row: self.max_row.max(other.max_row),
        }
    }

    /// Whether every cell of `self` is also in `other`.
    pub fn is_subset(&self, other: &CellRange) -> bool {
        other.contains_range(self)
    }

    /// Whether every cell of `other` is also in `self`.
    pub fn issuperset(&self, other: &CellRange) -> bool {
        self.contains_range(other)
    }

    fn contains_range(&self, other: &CellRange) -> bool {
        self.min_col <= other.min_col
            && self.min_row <= other.min_row
            && self.max_col >= other.max_col
            && self.max_row >= other.max_row
    }

    /// The range moved by the given offsets.
    ///
    /// A move that would leave the grid is refused rather than clamped. Clamping would
    /// silently produce a different range, and a shifted range that quietly stops covering
    /// what it covered is exactly the bug that inserts a row at the top of a sheet and
    /// corrupts a merged block.
    pub fn shift(&self, col_shift: i64, row_shift: i64) -> Result<CellRange> {
        let bounds = |value: u32, delta: i64, limit: u32, name: &str| -> Result<u32> {
            let moved = value as i64 + delta;
            if moved < 1 || moved > limit as i64 {
                return Err(Error::Value(format!(
                    "shifting {name} by {delta} from {value} leaves the grid"
                )));
            }
            Ok(moved as u32)
        };
        let min_col = bounds(self.min_col, col_shift, MAX_COLUMN, "the column")?;
        let max_col = bounds(self.max_col, col_shift, MAX_COLUMN, "the column")?;
        let min_row = bounds(self.min_row, row_shift, MAX_ROW, "the row")?;
        let max_row = bounds(self.max_row, row_shift, MAX_ROW, "the row")?;
        CellRange::new(min_col, min_row, max_col, max_row)
            .ok_or_else(|| Error::Value("the shift inverted the range".to_string()))
    }

    /// A range grown by the given amounts, one side at a time.
    pub fn expand(&self, right: u32, down: u32, left: u32, up: u32) -> Result<CellRange> {
        let min_col = self.min_col.saturating_sub(left);
        let min_row = self.min_row.saturating_sub(up);
        let max_col = (self.max_col + right).min(MAX_COLUMN);
        let max_row = (self.max_row + down).min(MAX_ROW);
        CellRange::new(min_col, min_row, max_col, max_row)
            .ok_or_else(|| Error::Value("expanding left or up past A1".to_string()))
    }

    /// A range reduced by the given amounts.
    ///
    /// A shrink that would consume the whole range is refused. Returning `None` would leave
    /// the caller with an unusable range and no signal about why; the error says which side
    /// was too large.
    pub fn shrink(&self, right: u32, bottom: u32, left: u32, top: u32) -> Result<CellRange> {
        let columns = self.max_col - self.min_col + 1;
        let rows = self.max_row - self.min_row + 1;
        if left >= columns || right >= columns || top >= rows || bottom >= rows {
            return Err(Error::Value(
                "the shrink would consume the whole range".to_string(),
            ));
        }
        CellRange::new(
            self.min_col + left,
            self.min_row + top,
            self.max_col - right,
            self.max_row - bottom,
        )
        .ok_or_else(|| Error::Value("the shrink inverted the range".to_string()))
    }

    /// The range's corners as `(min_col, min_row, max_col, max_row)`.
    pub fn bounds(&self) -> (u32, u32, u32, u32) {
        (self.min_col, self.min_row, self.max_col, self.max_row)
    }
}

/// How many cells [`CellRange::cells`] will materialise before giving up.
///
/// A range's whole point is that it is small, and a caller asking for the cells of
/// `A1:XFD1048576` is asking for sixteen million strings. Failing with an explanation beats
/// exhausting memory.
pub const ITERATION_LIMIT: u64 = 10_000_000;

impl fmt::Display for CellRange {
    /// Renders the way openpyxl's `coord` does, which collapses a one-cell range.
    ///
    /// `A1:A1` and `A1` name the same cell, but openpyxl prints the short form and a round
    /// trip through a string would otherwise rewrite every single-cell reference in the
    /// workbook into its longer equivalent.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let min_col = get_column_letter(self.min_col).map_err(|_| fmt::Error)?;
        if self.min_col == self.max_col && self.min_row == self.max_row {
            return write!(f, "{min_col}{}", self.min_row);
        }
        let max_col = get_column_letter(self.max_col).map_err(|_| fmt::Error)?;
        write!(f, "{min_col}{}:{max_col}{}", self.min_row, self.max_row)
    }
}

/// Read one corner of a range as `(column, row)`.
fn corners(coordinate: &str) -> Result<(u32, u32)> {
    let (letters, row) = coordinate_from_string(coordinate.trim())?;
    let column = column_index_from_string(&letters)?;
    if row == 0 {
        return Err(Error::CellCoordinates(format!("{coordinate} has no row")));
    }
    Ok((column, row))
}

/// An ordered, de-duplicated set of [`CellRange`]s.
///
/// openpyxl's `MultiCellRange`. A sheet can hold several separate ranges — conditional
/// formatting applies to a `sqref` like `A1:B2 D5:E6` — so this is a collection rather than
/// a single rectangle. Overlapping ranges are *not* merged: a caller asking "which rules
/// apply to A1" wants to know they asked twice, not that the answer was quietly deduplicated
/// on the way in.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MultiCellRange {
    ranges: BTreeSet<CellRange>,
}

impl MultiCellRange {
    /// An empty collection.
    pub fn new() -> Self {
        MultiCellRange::default()
    }

    /// Build from range strings.
    pub fn parse<'a>(ranges: impl IntoIterator<Item = &'a str>) -> Result<Self> {
        let mut out = MultiCellRange::new();
        for range in ranges {
            out.add_str(range)?;
        }
        Ok(out)
    }

    /// Add a range from its string form.
    pub fn add_str(&mut self, range: &str) -> Result<&mut Self> {
        let parsed = CellRange::parse(range)?;
        self.ranges.insert(parsed);
        Ok(self)
    }

    /// Add an already-parsed range.
    pub fn add(&mut self, range: CellRange) -> &mut Self {
        self.ranges.insert(range);
        self
    }

    /// Remove a range, by value or by the range covering the same cells.
    pub fn remove(&mut self, range: &str) -> Result<&mut Self> {
        let parsed = CellRange::parse(range)?;
        self.ranges.remove(&parsed);
        Ok(self)
    }

    /// Whether any range covers `coordinate`.
    pub fn contains(&self, coordinate: &str) -> bool {
        self.ranges.iter().any(|range| range.contains(coordinate))
    }

    /// The ranges covering `coordinate`.
    pub fn ranges_containing(&self, coordinate: &str) -> Vec<CellRange> {
        self.ranges
            .iter()
            .filter(|range| range.contains(coordinate))
            .copied()
            .collect()
    }

    /// How many ranges are held, counting overlaps separately.
    pub fn len(&self) -> usize {
        self.ranges.len()
    }

    /// Whether the collection is empty.
    pub fn is_empty(&self) -> bool {
        self.ranges.is_empty()
    }

    /// The ranges, in order.
    pub fn iter(&self) -> impl Iterator<Item = &CellRange> {
        self.ranges.iter()
    }
}

impl fmt::Display for MultiCellRange {
    /// Ranges separated by a space, as a `sqref` is written in the XML.
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        let rendered: Vec<String> = self.ranges.iter().map(|range| range.to_string()).collect();
        f.write_str(&rendered.join(" "))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn range(text: &str) -> CellRange {
        CellRange::parse(text).expect("a valid range")
    }

    #[test]
    fn a_range_parses_to_inclusive_corners() {
        let r = range("A1:C3");
        assert_eq!(r.bounds(), (1, 1, 3, 3));
        assert_eq!(r.left(), 1);
        assert_eq!(r.right(), 3);
        assert_eq!(r.top(), 1);
        assert_eq!(r.bottom(), 3);
        // Nine cells, not eight: the bounds are inclusive at both ends, which is the
        // opposite of RangeBounds and the easiest off-by-one in this file to get wrong.
        assert_eq!(r.size(), 9);
    }

    #[test]
    fn a_bare_coordinate_is_a_one_cell_range() {
        let r = range("B2");
        assert_eq!(r.bounds(), (2, 2, 2, 2));
        assert_eq!(r.size(), 1);
    }

    #[test]
    fn it_renders_back_to_the_input() {
        for text in ["A1", "A1:C3", "B2:D4", "AA10:AB12"] {
            assert_eq!(range(text).to_string(), text);
        }
    }

    #[test]
    fn a_one_cell_range_renders_short() {
        // `A1:A1` and `A1` name the same cell, but openpyxl's `coord` collapses it, and a
        // round trip through a string would otherwise lengthen every single-cell reference.
        assert_eq!(range("A1:A1").to_string(), "A1");
        assert_eq!(range("A1").to_string(), "A1");
    }

    #[test]
    fn an_inverted_range_is_refused_rather_than_normalised() {
        // Silently reading A5:B2 as B2:A5 would turn a typo into a wrong answer.
        assert!(CellRange::parse("A5:B2").is_err());
    }

    #[test]
    fn cells_come_out_row_major() {
        let cells = range("A1:B2").cells().expect("cells");
        assert_eq!(cells, ["A1", "B1", "A2", "B2"]);
    }

    #[test]
    fn containment_includes_its_own_corners() {
        let r = range("B2:D4");
        assert!(r.contains("B2"));
        assert!(r.contains("D4"));
        assert!(r.contains("C3"));
        assert!(!r.contains("A2"));
        assert!(!r.contains("B5"));
        // A malformed coordinate is simply not contained, rather than an error: containment
        // is a question, and a question with an unanswerable input is `false`.
        assert!(!r.contains("not a cell"));
    }

    #[test]
    fn intersection_and_union() {
        let a = range("A1:C10");
        let b = range("B2:D3");
        assert_eq!(a.intersection(&b).expect("overlap").to_string(), "B2:C3");
        assert_eq!(a.union(&b).to_string(), "A1:D10");
        // Touching at an edge is still an intersection: A1:A5 and A5:A9 share A5.
        assert_eq!(
            range("A1:A5")
                .intersection(&range("A5:A9"))
                .expect("share A5")
                .to_string(),
            "A5"
        );
    }

    #[test]
    fn disjoint_ranges_have_no_intersection() {
        let a = range("A1:B2");
        let b = range("C3:D4");
        assert!(a.isdisjoint(&b));
        assert!(a.intersection(&b).is_none());
        assert!(!a.isdisjoint(&range("B2:C3")));
    }

    #[test]
    fn subset_and_superset() {
        let whole = range("A1:D10");
        let part = range("B2:C3");
        assert!(part.is_subset(&whole));
        assert!(!whole.is_subset(&part));
        assert!(whole.issuperset(&part));
        assert!(part.issuperset(&part), "a range contains itself");
    }

    #[test]
    fn shifting_moves_both_corners() {
        assert_eq!(range("B2:C3").shift(1, 0).unwrap().to_string(), "C2:D3");
        assert_eq!(range("B2:C3").shift(0, 2).unwrap().to_string(), "B4:C5");
        assert_eq!(range("B2:C3").shift(-1, -1).unwrap().to_string(), "A1:B2");
    }

    #[test]
    fn a_shift_off_the_grid_is_refused() {
        // Clamping would return A1:B2, which looks like a valid range and covers the wrong
        // cells entirely -- the failure mode that corrupts a merged block after an insert.
        assert!(range("A1:B2").shift(0, -1).is_err());
        assert!(range("A1:B2").shift(-1, 0).is_err());
    }

    #[test]
    fn expand_and_shrink_are_inverses() {
        let start = range("C3:D4");
        let grown = start.expand(1, 1, 1, 1).unwrap();
        assert_eq!(grown.to_string(), "B2:E5");
        assert_eq!(grown.shrink(1, 1, 1, 1).unwrap(), start);
    }

    #[test]
    fn a_shrink_that_consumes_the_range_is_refused() {
        assert!(range("A1:B2").shrink(2, 2, 0, 0).is_err());
        assert!(range("A1:B2").shrink(0, 0, 2, 0).is_err());
    }

    #[test]
    fn a_range_too_large_to_iterate_is_refused_with_a_reason() {
        // A whole-sheet range is sixteen million cells. Returning an error naming the count
        // is more useful than exhausting memory.
        let huge = CellRange::new(1, 1, MAX_COLUMN, MAX_ROW).expect("a valid range");
        let err = huge.cells().expect_err("too many cells").to_string();
        assert!(err.contains("narrow it"), "{err}");
    }

    #[test]
    fn a_multi_range_keeps_overlaps() {
        let mut multi = MultiCellRange::new();
        multi.add_str("A1:B2").unwrap();
        multi.add_str("B2:C3").unwrap();
        assert_eq!(multi.len(), 2);
        assert_eq!(multi.to_string(), "A1:B2 B2:C3");
        // B2 is in both, and the collection says so rather than merging them away.
        assert_eq!(multi.ranges_containing("B2").len(), 2);
        assert!(multi.contains("A1"));
        assert!(!multi.contains("D4"));
    }

    #[test]
    fn a_multi_range_deduplicates_identical_ranges() {
        let mut multi = MultiCellRange::new();
        multi.add_str("A1:B2").unwrap();
        multi.add_str("A1:B2").unwrap();
        assert_eq!(multi.len(), 1);
    }

    #[test]
    fn a_multi_range_can_be_emptied() {
        let mut multi = MultiCellRange::parse(["A1:B2", "D5:E6"]).unwrap();
        assert!(!multi.is_empty());
        multi.remove("A1:B2").unwrap();
        assert_eq!(multi.to_string(), "D5:E6");
        multi.remove("D5:E6").unwrap();
        assert!(multi.is_empty());
        assert_eq!(multi.to_string(), "");
    }

    #[test]
    fn rows_and_cols_are_inclusive() {
        let r = range("B2:D4");
        assert_eq!(r.rows(), [2, 3, 4]);
        assert_eq!(r.cols(), [2, 3, 4]);
    }
}
