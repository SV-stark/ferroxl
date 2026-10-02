//! Handling shared and array formulas (`openpyxl/cell/formula.py`).
//!
//! The Python module is a stub containing a `SharedFormula` holder whose accessors
//! shadow the attributes (so they can never be read back). The rewrite keeps the useful
//! part: a store of shared formulas that expands an expression across the cells in a
//! shared range by translating relative references.

use std::collections::HashMap;

use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::Result;

/// A formula shared across a range of cells.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SharedFormula {
    /// Range of cells the formula applies to.
    pub range: String,
    /// Key identifying the shared formula group.
    pub key: String,
    /// The formula expression, including its leading `=`.
    pub expression: String,
}

impl SharedFormula {
    /// Build a shared formula record.
    pub fn new(
        range: impl Into<String>,
        key: impl Into<String>,
        expression: impl Into<String>,
    ) -> Self {
        SharedFormula {
            range: range.into(),
            key: key.into(),
            expression: expression.into(),
        }
    }
}

/// Per-worksheet store of shared formulas.
#[derive(Debug, Clone, Default)]
pub struct FormulaStore {
    entries: HashMap<String, SharedFormula>,
}

impl FormulaStore {
    /// An empty store.
    pub fn new() -> Self {
        FormulaStore::default()
    }

    /// Register a shared formula. Later registrations for the same key are ignored, which
    /// matches the Python `if key not in ws.formula_attributes` guard.
    pub fn add(&mut self, formula: SharedFormula) {
        self.entries.entry(formula.key.clone()).or_insert(formula);
    }

    /// Look up a shared formula by key.
    pub fn get(&self, key: &str) -> Option<&SharedFormula> {
        self.entries.get(key)
    }

    /// All registered keys.
    pub fn keys(&self) -> impl Iterator<Item = &String> {
        self.entries.keys()
    }

    /// The expression stored under `key`.
    pub fn expression(&self, key: &str) -> Option<&str> {
        self.entries.get(key).map(|f| f.expression.as_str())
    }

    /// Expand the formula registered under `key` to the cell at `coordinate`.
    ///
    /// The stored expression is anchored at the top-left cell of its range; relative
    /// references are shifted by the difference in row and column.
    pub fn expand(&self, key: &str, coordinate: &str) -> Result<Option<String>> {
        let Some(formula) = self.entries.get(key) else {
            return Ok(None);
        };
        let (start, _) = split_range(&formula.range);
        let (anchor_col, anchor_row) = coordinate_from_string(&start)?;
        let (target_col, target_row) = coordinate_from_string(coordinate)?;
        let anchor_col_idx = column_index_from_string(&anchor_col)?;
        let target_col_idx = column_index_from_string(&target_col)?;
        let row_delta = target_row as i64 - anchor_row as i64;
        let col_delta = target_col_idx as i64 - anchor_col_idx as i64;
        Ok(Some(shift_references(
            &formula.expression,
            row_delta,
            col_delta,
        )))
    }
}

/// Split `A1:B3` into its start and end coordinates.
pub fn split_range(range: &str) -> (String, String) {
    match range.split_once(':') {
        Some((start, end)) => (start.to_string(), end.to_string()),
        None => (range.to_string(), range.to_string()),
    }
}

/// Shift relative A1-style references in a formula by the given row/column deltas.
///
/// Absolute components (marked with `$`) are left alone, as are references that would move
/// off the sheet. Text inside quoted literals is copied verbatim, and a run of
/// letters/digits that is not at a token boundary (`LOG10`, `A1B2`) is not a reference.
pub fn shift_references(formula: &str, row_delta: i64, col_delta: i64) -> String {
    if row_delta == 0 && col_delta == 0 {
        return formula.to_string();
    }
    let bytes = formula.as_bytes();
    let mut out = String::with_capacity(formula.len());
    let mut i = 0usize;

    while i < bytes.len() {
        let ch = bytes[i] as char;
        if ch == '"' || ch == '\'' {
            let end = closing_quote(formula, i);
            out.push_str(&formula[i..end]);
            i = end;
            continue;
        }
        if ch == '$' || ch.is_ascii_alphabetic() {
            if let Some((len, shifted)) = shift_reference_at(formula, i, row_delta, col_delta) {
                out.push_str(&shifted);
                i += len;
                continue;
            }
        }
        // Advance by whole characters so multi-byte input is not split.
        let mut step = 1;
        while i + step < bytes.len() && (bytes[i + step] & 0xC0) == 0x80 {
            step += 1;
        }
        out.push_str(&formula[i..i + step]);
        i += step;
    }
    out
}

/// Index just past the closing quote of the literal starting at `start`.
fn closing_quote(formula: &str, start: usize) -> usize {
    let bytes = formula.as_bytes();
    let quote = bytes[start];
    let mut i = start + 1;
    while i < bytes.len() {
        if bytes[i] == quote {
            // A doubled quote is an escaped quote inside the literal.
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

/// Try to read a cell reference at `start`, returning its length and shifted form.
fn shift_reference_at(
    formula: &str,
    start: usize,
    row_delta: i64,
    col_delta: i64,
) -> Option<(usize, String)> {
    let bytes = formula.as_bytes();
    let mut i = start;
    // A reference never continues an identifier, so a preceding letter or digit rules it
    // out. This is what keeps `LOG10` from being read as the cell `LOG10`.
    if start > 0 {
        let before = bytes[start - 1] as char;
        if before.is_ascii_alphanumeric() || before == '_' || before == '.' {
            return None;
        }
    }
    let mut col_abs = false;
    if bytes[i] == b'$' {
        col_abs = true;
        i += 1;
    }
    let col_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_alphabetic() {
        i += 1;
    }
    if i == col_start || i - col_start > 3 {
        return None;
    }
    let col_text = formula[col_start..i].to_uppercase();
    // A column is written entirely in upper case or entirely in lower case; a mixed run
    // such as the `myA` in `myA1` is part of an identifier, not a reference.
    let raw = &formula[col_start..i];
    let all_upper = raw.chars().all(|c| c.is_ascii_uppercase());
    let all_lower = raw.chars().all(|c| c.is_ascii_lowercase());
    if !all_upper && !all_lower {
        return None;
    }
    let Ok(col_idx) = column_index_from_string(&col_text) else {
        return None;
    };
    // A trailing letter run means this is an identifier, not a reference.
    if i < bytes.len() && (bytes[i] as char).is_ascii_alphabetic() {
        return None;
    }
    let mut row_abs = false;
    if i < bytes.len() && bytes[i] == b'$' {
        row_abs = true;
        i += 1;
    }
    let digits_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_digit() {
        i += 1;
    }
    if i == digits_start {
        return None;
    }
    let row_text = &formula[digits_start..i];
    let Ok(row) = row_text.parse::<i64>() else {
        return None;
    };
    // The character after a reference must not continue an identifier.
    if i < bytes.len() {
        let next = bytes[i] as char;
        if next.is_ascii_alphanumeric() || next == '_' || next == '$' {
            return None;
        }
        // A reference is never immediately applied as a call, so `LOG10(` is a function
        // name rather than the cell `LOG10`. Excel's grammar has no reference-then-`(`.
        if next == '(' {
            return None;
        }
    }

    let mut new_col = col_idx as i64;
    let mut new_row = row;
    let mut changed = false;
    if !col_abs && col_delta != 0 {
        new_col += col_delta;
        if new_col < 1 || new_col > crate::cell::utils::MAX_COLUMN_INDEX as i64 {
            new_col = col_idx as i64;
        } else {
            changed = true;
        }
    }
    if !row_abs && row_delta != 0 {
        new_row += row_delta;
        if new_row < 1 {
            new_row = row;
        } else {
            changed = true;
        }
    }
    if !changed {
        return Some((i - start, formula[start..i].to_string()));
    }
    let letters = get_column_letter(new_col as u32).ok()?;
    let shifted = format!(
        "{}{}{}{}",
        if col_abs { "$" } else { "" },
        letters,
        if row_abs { "$" } else { "" },
        new_row
    );
    Some((i - start, shifted))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn registers_and_reads() {
        let mut store = FormulaStore::new();
        store.add(SharedFormula::new("A1:B3", "A1", "=A1+1"));
        assert_eq!(store.expression("A1"), Some("=A1+1"));
        assert_eq!(store.get("missing"), None);
        assert_eq!(store.keys().count(), 1);
    }

    #[test]
    fn expansion_shifts_relative_refs() {
        let mut store = FormulaStore::new();
        store.add(SharedFormula::new("A1:B3", "A1", "=SUM(A1:B1)"));
        assert_eq!(
            store.expand("A1", "B2").unwrap().as_deref(),
            Some("=SUM(B2:C2)")
        );
        // The anchor cell is unchanged.
        assert_eq!(
            store.expand("A1", "A1").unwrap().as_deref(),
            Some("=SUM(A1:B1)")
        );
    }

    #[test]
    fn absolute_refs_are_untouched() {
        let mut store = FormulaStore::new();
        store.add(SharedFormula::new("A1:B3", "A1", "=$A$1+A1"));
        assert_eq!(
            store.expand("A1", "C4").unwrap().as_deref(),
            Some("=$A$1+C4")
        );
    }

    #[test]
    fn shifting_past_the_sheet_is_ignored() {
        assert_eq!(shift_references("=A1", -5, 0), "=A1");
        assert_eq!(shift_references("=A1", 0, -5), "=A1");
        assert_eq!(shift_references("=ZZ1", 0, 1), "=AAA1");
    }

    #[test]
    fn ranges_split() {
        assert_eq!(split_range("A1:B3"), ("A1".to_string(), "B3".to_string()));
        assert_eq!(split_range("A1"), ("A1".to_string(), "A1".to_string()));
    }

    #[test]
    fn identifiers_are_not_mistaken_for_references() {
        assert_eq!(shift_references("=LOG10(A1)", 1, 0), "=LOG10(A2)");
        assert_eq!(shift_references("=myA1", 1, 0), "=myA1");
    }
}
