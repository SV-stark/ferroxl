//! Formula translation: what openpyxl calls `openpyxl.formula.translate.Translator`.
//!
//! A formula written for one cell is not automatically correct for another. `=A1*2` means
//! "the cell above this one, doubled" only because of where it sits, so moving it without
//! rewriting it silently changes what it computes. Excel resolves this with the dollar
//! sign: `$A1` fixes the column, `A$1` fixes the row, and a reference with neither moves
//! with the formula.
//!
//! This module applies those rules. It is the operation behind copy-paste, fill-right and
//! fill-down, and behind openpyxl's own `Translator`.
//!
//! What it does that [`shift_references`][shift] does not:
//!
//! - It takes an origin and a destination and works out the offsets, which is how a caller
//!   actually has the information.
//! - It translates whole-row and whole-column references. `3:4` and `A:BC` contain no cell
//!   to shift, so a scanner looking for `$A$1`-shaped tokens passes over them and they
//!   survive a copy unchanged — which is wrong, because `=SUM(3:4)` in `A10` becomes
//!   `=SUM(13:14)` in `A20`.
//!
//! It refuses rather than guesses in the same place `shift_references` does: a reference
//! that would leave the grid is a `#REF!` in Excel, and clamping it to the edge would
//! quietly change the formula's meaning.

use std::fmt;

use crate::cell::formula::shift_references;
use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::{Error, Result};

/// The largest row and column Excel has, as the ECMA-376 grid defines them.
const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// A formula could not be translated.
///
/// Separate from [`Error`] so a caller can catch the one failure this can raise and not
/// have to distinguish it from I/O or XML errors that share the type.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TranslatorError {
    /// The reference that could not be translated, as it appeared in the formula.
    pub reference: String,
    /// Why.
    pub reason: String,
}

impl fmt::Display for TranslatorError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "cannot translate {}: {}", self.reference, self.reason)
    }
}

impl std::error::Error for TranslatorError {}

impl From<TranslatorError> for Error {
    fn from(err: TranslatorError) -> Self {
        Error::Value(err.to_string())
    }
}

/// Translates a formula from the cell it was written for to the cell it is going to.
///
/// Construct it with the formula and the cell it currently belongs to, then ask for the
/// formula as it should read at its destination:
///
/// ```
/// use ferroxl::formula::Translator;
///
/// let t = Translator::new("=A1*2", "B10").unwrap();
/// assert_eq!(t.translate_formula(Some("C12")).unwrap(), "=B3*2");
/// ```
#[derive(Debug, Clone)]
pub struct Translator {
    formula: String,
    origin: (u32, u32),
}

impl Translator {
    /// Build a translator for `formula` as written at `origin`.
    ///
    /// A leading `=` is optional, since a cell's formula is stored without one.
    pub fn new(formula: &str, origin: &str) -> Result<Self> {
        Ok(Self {
            formula: formula.strip_prefix('=').unwrap_or(formula).to_string(),
            origin: coordinate_tuple(origin)?,
        })
    }

    /// The formula this translator was built from, without its leading `=`.
    pub fn formula(&self) -> &str {
        &self.formula
    }

    /// The cell the formula was written for, as `(row, column)`.
    pub fn origin(&self) -> (u32, u32) {
        self.origin
    }

    /// The formula as it should read at `dest`.
    ///
    /// Passing `None` is the same as passing the origin, which is how a caller checks that
    /// translation is the identity — worth asserting, because a translator that shifted
    /// something in place would be a very quiet bug.
    pub fn translate_formula(&self, dest: Option<&str>) -> Result<String> {
        let (row_delta, col_delta) = match dest {
            Some(dest) => {
                let (row, col) = coordinate_tuple(dest)?;
                (
                    row as i64 - self.origin.0 as i64,
                    col as i64 - self.origin.1 as i64,
                )
            }
            None => (0, 0),
        };
        self.translate_formula_by(row_delta, col_delta)
    }

    /// The formula as it should read `row_delta` rows and `col_delta` columns away.
    pub fn translate_formula_by(&self, row_delta: i64, col_delta: i64) -> Result<String> {
        if row_delta == 0 && col_delta == 0 {
            return Ok(format!("={}", self.formula));
        }
        self.check_references_stay_on_the_grid(row_delta, col_delta)?;
        // `shift_references` handles the `$A$1` case, including skipping string literals and
        // refusing to read `LOG10` as the cell `LOG10`. Whole-row and whole-column ranges
        // are not `$A$1`-shaped, so they are done here and the result substituted back in.
        let shifted = shift_references(&self.formula, row_delta, col_delta);
        Ok(format!(
            "={}",
            self.translate_row_and_column_ranges(&shifted, row_delta, col_delta)?
        ))
    }

    /// Refuse to translate a formula that would push a relative reference off the grid.
    ///
    /// This has to run against the formula as written rather than against
    /// [`shift_references`]'s output, because that function returns a `String` and handles
    /// an off-grid reference by leaving it where it was. The result would then be a formula
    /// reading `=A1` in a cell where `=A1` no longer means what the original did, which is
    /// worse than an error: nothing about it looks wrong.
    ///
    /// Only unanchored references are checked. `$A1` keeps its column and `A$1` keeps its
    /// row, so neither can be the reference that leaves the grid.
    fn check_references_stay_on_the_grid(&self, row_delta: i64, col_delta: i64) -> Result<()> {
        let formula = &self.formula;
        let bytes = formula.as_bytes();
        let mut i = 0usize;

        while i < bytes.len() {
            let ch = bytes[i] as char;
            if ch == '"' || ch == '\'' {
                i = closing_quote(formula, i, ch as u8);
                continue;
            }
            if i > 0 {
                let before = bytes[i - 1] as char;
                if before.is_ascii_alphanumeric() || before == '_' || before == '.' || before == '!'
                {
                    i += 1;
                    continue;
                }
            }

            // A `$`-prefixed column followed by digits is a cell reference, which is the one
            // shape `check` is here for. Anything else is left to the range translation.
            let column = match read_column_token(formula, i) {
                Some(token) if token.has_digits => token,
                _ => {
                    i += 1;
                    continue;
                }
            };
            let row = match read_row_token(formula, column.row_at) {
                Some(row) => row,
                None => {
                    i += 1;
                    continue;
                }
            };
            if !column.absolute {
                shift_column(column.index, col_delta, false, &column.text)?;
            }
            if !row.absolute {
                shift_row(row.value, row_delta, false, &row.text)?;
            }
            i += column.len + row.len;
        }
        Ok(())
    }

    /// Rewrite `3:4` and `A:BC` in an already-cell-shifted formula.
    ///
    /// Runs second, on the output of [`shift_references`], because the two cannot be done in
    /// one pass over the text: `SUM(A:A)` contains no cell reference and `SUM(A1:A3)`
    /// contains two, and deciding which is which needs the whole token.
    fn translate_row_and_column_ranges(
        &self,
        formula: &str,
        row_delta: i64,
        col_delta: i64,
    ) -> Result<String> {
        let bytes = formula.as_bytes();
        let mut out = String::with_capacity(formula.len());
        let mut i = 0usize;

        while i < bytes.len() {
            let ch = bytes[i] as char;

            if ch == '"' {
                let end = closing_quote(formula, i, b'"');
                out.push_str(&formula[i..end]);
                i = end;
                continue;
            }
            if ch == '\'' {
                let end = closing_quote(formula, i, b'\'');
                out.push_str(&formula[i..end]);
                i = end;
                continue;
            }
            // Anything that is not part of a range is copied verbatim.
            if !(ch.is_ascii_alphanumeric() || ch == '$') {
                out.push(ch);
                i += 1;
                continue;
            }
            // A range cannot continue an identifier, which is what keeps `LOG10` and
            // `myRange` from being read as one.
            if i > 0 {
                let before = bytes[i - 1] as char;
                if before.is_ascii_alphanumeric() || before == '_' || before == '.' || before == '!'
                {
                    out.push(ch);
                    i += 1;
                    continue;
                }
            }

            if let Some((len, text)) = self.translate_one_range(formula, i, row_delta, col_delta)? {
                out.push_str(&text);
                i += len;
                continue;
            }
            let mut step = 1;
            while i + step < bytes.len() && (bytes[i + step] & 0xC0) == 0x80 {
                step += 1;
            }
            out.push_str(&formula[i..i + step]);
            i += step;
        }
        Ok(out)
    }

    /// Translate the range starting at `start`, if there is one.
    ///
    /// Returns the text consumed and its translated form. Three shapes are recognised:
    /// `3:4`, `A:BC`, and anything already handled by `shift_references` (which this
    /// declines, leaving it to be copied through).
    fn translate_one_range(
        &self,
        formula: &str,
        start: usize,
        row_delta: i64,
        col_delta: i64,
    ) -> Result<Option<(usize, String)>> {
        let bytes = formula.as_bytes();

        // -- A column range: `A:BC`, optionally anchored with `$` on either end. ---------
        if let Some((len, text)) = self.column_range_at(formula, start, col_delta)? {
            return Ok(Some((len, text)));
        }
        // -- A row range: `3:4`, likewise anchored. --------------------------------------
        if let Some((len, text)) = self.row_range_at(formula, start, row_delta)? {
            return Ok(Some((len, text)));
        }
        // Not a whole-column or whole-row range. `$A$1` and `A1` are `shift_references`'s
        // business and are passed through untouched here.
        let _ = bytes;
        Ok(None)
    }

    /// Read a whole-column range at `start`, as in `A:BC` or `$A:$BC`.
    fn column_range_at(
        &self,
        formula: &str,
        start: usize,
        col_delta: i64,
    ) -> Result<Option<(usize, String)>> {
        let first = match read_column_token(formula, start) {
            Some(token) => token,
            None => return Ok(None),
        };
        let colon = start + first.len;
        if formula.as_bytes().get(colon) != Some(&b':') {
            return Ok(None);
        }
        // `A1:A9` is a cell range, not a column range: the first token carried digits.
        if first.has_digits {
            return Ok(None);
        }
        let second = match read_column_token(formula, colon + 1) {
            Some(token) => token,
            None => return Ok(None),
        };
        let end = colon + 1 + second.len;

        // A reference that moved off the grid is `#REF!` in Excel. Clamping it to the last
        // row or column would keep the formula loadable while quietly meaning something
        // else, which is worse than a visible error.
        let from = shift_column(first.index, col_delta, first.absolute, &first.text)?;
        let to = shift_column(second.index, col_delta, second.absolute, &second.text)?;
        let from_letter = get_column_letter(from).map_err(|_| out_of_bounds(&first.text))?;
        let to_letter = get_column_letter(to).map_err(|_| out_of_bounds(&second.text))?;
        let text = format!(
            "{}{}:{}{}",
            if first.absolute { "$" } else { "" },
            from_letter,
            if second.absolute { "$" } else { "" },
            to_letter
        );
        Ok(Some((end - start, text)))
    }

    /// Read a whole-row range at `start`, as in `3:4` or `$3:$4`.
    fn row_range_at(
        &self,
        formula: &str,
        start: usize,
        row_delta: i64,
    ) -> Result<Option<(usize, String)>> {
        let first = match read_row_token(formula, start) {
            Some(token) => token,
            None => return Ok(None),
        };
        let colon = start + first.len;
        if formula.as_bytes().get(colon) != Some(&b':') {
            return Ok(None);
        }
        let second = match read_row_token(formula, colon + 1) {
            Some(token) => token,
            None => return Ok(None),
        };
        let end = colon + 1 + second.len;

        let from = shift_row(first.value, row_delta, first.absolute, &first.text)?;
        let to = shift_row(second.value, row_delta, second.absolute, &second.text)?;
        Ok(Some((
            end - start,
            format!(
                "{}{}:{}{}",
                if first.absolute { "$" } else { "" },
                from,
                if second.absolute { "$" } else { "" },
                to
            ),
        )))
    }
}

/// A column token read off a formula.
struct ColumnToken {
    index: u32,
    absolute: bool,
    has_digits: bool,
    /// How much of the formula the token consumed, digits included.
    len: usize,
    /// Where the trailing digits start. `len` cannot be used for this, because it runs
    /// past them: for `A1` the letters end at 1 and `len` is 2.
    row_at: usize,
    text: String,
}

/// Read `$A` or `A` or `AB12` at `start`, reporting whether digits followed.
fn read_column_token(formula: &str, start: usize) -> Option<ColumnToken> {
    let bytes = formula.as_bytes();
    let mut i = start;
    let mut absolute = false;
    if bytes.get(i) == Some(&b'$') {
        absolute = true;
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_alphabetic() {
        i += 1;
    }
    if i == letters_start || i - letters_start > 3 {
        return None;
    }
    let text = formula[start..i].to_string();
    let index = column_index_from_string(&formula[letters_start..i].to_uppercase()).ok()?;
    let digits_start = i;
    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }
    while i < bytes.len() && (bytes[i] as char).is_ascii_digit() {
        i += 1;
    }
    Some(ColumnToken {
        index,
        absolute,
        has_digits: i > digits_start,
        len: i - start,
        row_at: start + (digits_start - start),
        text,
    })
}

/// A row token read off a formula.
struct RowToken {
    value: u32,
    absolute: bool,
    len: usize,
    text: String,
}

/// Read `$3` or `3` at `start`.
fn read_row_token(formula: &str, start: usize) -> Option<RowToken> {
    let bytes = formula.as_bytes();
    let mut i = start;
    let mut absolute = false;
    if bytes.get(i) == Some(&b'$') {
        absolute = true;
        i += 1;
    }
    let digits_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_digit() {
        i += 1;
    }
    if i == digits_start {
        return None;
    }
    let text = formula[start..i].to_string();
    let value: u32 = formula[digits_start..i].parse().ok()?;
    Some(RowToken {
        value,
        absolute,
        len: i - start,
        text,
    })
}

fn shift_column(
    index: u32,
    delta: i64,
    absolute: bool,
    original: &str,
) -> Result<u32, TranslatorError> {
    // `$A` is a fixed column: copying the formula does not move it, so it cannot be the
    // thing that goes off the edge either.
    if absolute {
        return Ok(index);
    }
    let moved = index as i64 + delta;
    if moved < 1 || moved > MAX_COLUMN as i64 {
        return Err(out_of_bounds(original));
    }
    Ok(moved as u32)
}

fn shift_row(
    value: u32,
    delta: i64,
    absolute: bool,
    original: &str,
) -> Result<u32, TranslatorError> {
    if absolute {
        return Ok(value);
    }
    let moved = value as i64 + delta;
    if moved < 1 || moved > MAX_ROW as i64 {
        return Err(out_of_bounds(original));
    }
    Ok(moved as u32)
}

fn out_of_bounds(reference: &str) -> TranslatorError {
    TranslatorError {
        reference: reference.to_string(),
        reason: "it moves outside the grid, which Excel reports as #REF!".to_string(),
    }
}

/// Index just past the closing quote of the literal starting at `start`.
fn closing_quote(formula: &str, start: usize, quote: u8) -> usize {
    let bytes = formula.as_bytes();
    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == quote {
            if i + 1 < bytes.len() && bytes[i + 1] == quote {
                i += 2;
                continue;
            }
            return i + 1;
        }
        i += 1;
    }
    formula.len()
}

/// Read `origin` as `(row, column)`, both 1-based.
fn coordinate_tuple(coordinate: &str) -> Result<(u32, u32)> {
    let cleaned = coordinate.trim().trim_start_matches('=');
    let (letters, row) = coordinate_from_string(cleaned)?;
    let column = column_index_from_string(&letters)?;
    if row == 0 {
        return Err(Error::CellCoordinates(format!("{coordinate} has no row")));
    }
    Ok((row, column))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn translated(formula: &str, origin: &str, dest: &str) -> String {
        Translator::new(formula, origin)
            .expect("translator")
            .translate_formula(Some(dest))
            .expect("translated")
    }

    #[test]
    fn a_relative_reference_moves_with_the_formula() {
        assert_eq!(translated("=A1*2", "B10", "C12"), "=B3*2");
    }

    #[test]
    fn a_fixed_column_stays_put_while_the_row_moves() {
        assert_eq!(translated("=$A1*2", "B10", "C12"), "=$A3*2");
    }

    #[test]
    fn a_fixed_row_stays_put_while_the_column_moves() {
        assert_eq!(translated("=A$1*2", "B10", "C12"), "=B$1*2");
    }

    #[test]
    fn a_fully_fixed_reference_does_not_move_at_all() {
        assert_eq!(translated("=$A$1*2", "B10", "C12"), "=$A$1*2");
    }

    #[test]
    fn both_dollars_on_only_one_end_of_a_range_are_honoured() {
        // The top-left corner is anchored in both directions and stays; the bottom-right
        // corner is anchored in neither and moves by the full offset.
        assert_eq!(translated("=SUM($A$1:B2)", "D4", "E5"), "=SUM($A$1:C3)");
    }

    #[test]
    fn a_whole_row_range_moves() {
        // `3:4` holds no cell, so a scanner looking for `$A$1` would step over it and the
        // range would survive the copy unchanged. In Excel, `SUM(3:4)` in A10 is `SUM(13:14)`
        // in A20.
        assert_eq!(translated("=SUM(3:4)", "A10", "A20"), "=SUM(13:14)");
    }

    #[test]
    fn a_whole_row_range_with_dollars_stays_put() {
        assert_eq!(translated("=SUM($3:$4)", "A10", "A20"), "=SUM($3:$4)");
    }

    #[test]
    fn a_whole_column_range_moves() {
        assert_eq!(translated("=SUM(A:B)", "A1", "C1"), "=SUM(C:D)");
    }

    #[test]
    fn a_whole_column_range_with_dollars_stays_put() {
        assert_eq!(translated("=SUM($A:$B)", "A1", "C1"), "=SUM($A:$B)");
    }

    #[test]
    fn a_cell_range_is_not_mistaken_for_a_whole_column_range() {
        // `A1:A9` starts with letters followed by digits, so it is a cell range and
        // `shift_references` owns it.
        assert_eq!(translated("=SUM(A1:A9)", "C3", "D4"), "=SUM(B2:B10)");
    }

    #[test]
    fn a_function_name_is_not_mistaken_for_a_column_range() {
        // `LOG10(` has no colon after it, and `LOG10` is not a column token in the first
        // place, so neither range reader claims it.
        assert_eq!(translated("=LOG10(A1)", "C3", "D4"), "=LOG10(B2)");
    }

    #[test]
    fn a_defined_name_is_left_alone() {
        assert_eq!(translated("=SUM(Totals)+A1", "C3", "D4"), "=SUM(Totals)+B2");
    }

    #[test]
    fn a_string_literal_is_not_translated() {
        assert_eq!(
            translated("=IF(A1=\"3:4\",\"A:B\",B1)", "C3", "D4"),
            "=IF(B2=\"3:4\",\"A:B\",C2)"
        );
    }

    #[test]
    fn a_sheet_qualified_reference_moves_its_cell_but_keeps_the_sheet() {
        assert_eq!(translated("=Data!B2", "A1", "C1"), "=Data!D2");
    }

    #[test]
    fn a_fully_anchored_reference_cannot_go_off_the_grid() {
        // It does not move, so there is nothing to leave the grid. Checked because the
        // off-grid case below would otherwise pass for the wrong reason if anchoring were
        // broken.
        assert_eq!(translated("=$A$1", "B10", "B5"), "=$A$1");
    }

    #[test]
    fn a_relative_reference_that_would_leave_the_grid_is_an_error() {
        // `A1` copied up five rows becomes row -4, which Excel reports as #REF!. Clamping
        // would keep the formula loadable while quietly changing what it means.
        let err = Translator::new("=A1", "B10")
            .expect("translator")
            .translate_formula(Some("B5"));
        let message = err.expect_err("off the top of the grid").to_string();
        assert!(message.contains("#REF!"), "{message}");
    }

    #[test]
    fn translating_to_the_origin_is_the_identity() {
        // Worth asserting rather than assuming: a translator that shifted something in place
        // would be a very quiet bug, because the formula would still parse.
        let original = "=SUM(A1:B2)+$C$3+SUM(4:5)";
        assert_eq!(
            Translator::new(original, "D7")
                .unwrap()
                .translate_formula(None)
                .unwrap(),
            original
        );
        assert_eq!(translated(original, "D7", "D7"), original);
    }

    #[test]
    fn a_leading_equals_sign_is_optional() {
        assert_eq!(
            Translator::new("A1+1", "B2")
                .unwrap()
                .translate_formula(Some("C3"))
                .unwrap(),
            "=B2+1"
        );
    }

    #[test]
    fn an_empty_formula_stays_empty() {
        assert_eq!(
            Translator::new("", "A1")
                .unwrap()
                .translate_formula(Some("B2"))
                .unwrap(),
            "="
        );
    }

    #[test]
    fn a_literal_alone_is_returned_unchanged() {
        // openpyxl returns the literal without a leading `=` when the whole formula is one.
        assert_eq!(translated("42", "A1", "B2"), "=42");
    }
}
