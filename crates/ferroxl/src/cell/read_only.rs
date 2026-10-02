//! Lazy, read-only cells (`openpyxl/cell/read_only.py`).
//!
//! `ReadOnlyCell` is a detached value: it holds a raw string plus the style id it was
//! written with, and only converts on access. The Python class keeps the string and
//! style tables in class-level globals set by `IterableWorksheet`; Rust threads the
//! tables through explicitly, which avoids global mutable state.

use std::collections::HashMap;
use std::sync::Arc;

use chrono::NaiveDateTime;

use crate::date_time::{from_excel, BaseDate, ExcelDateTime};
use crate::exceptions::{Error, Result};
use crate::styles::style::Style;
use crate::styles::Style as StyleAlias;

/// The tables needed to resolve read-only cells.
#[derive(Debug, Default)]
pub struct ReadOnlyTables {
    /// Shared string table, index to text.
    pub string_table: Arc<HashMap<usize, String>>,
    /// Style table, `cellXfs` index to style.
    pub style_table: Arc<Vec<StyleAlias>>,
    /// The workbook's date system.
    pub base_date: BaseDate,
}

impl Clone for ReadOnlyTables {
    fn clone(&self) -> Self {
        ReadOnlyTables {
            string_table: Arc::clone(&self.string_table),
            style_table: Arc::clone(&self.style_table),
            base_date: self.base_date,
        }
    }
}

/// A cell produced by the iterator-based reader.
#[derive(Debug, Clone, PartialEq)]
pub struct ReadOnlyCell {
    /// 1-based row number, `None` for the shared empty cell.
    pub row: Option<u32>,
    /// Column letters, `None` for the shared empty cell.
    pub column: Option<String>,
    /// The raw string value as stored in the file.
    raw: Option<String>,
    /// The inferred type.
    pub data_type: crate::cell::cell::DataType,
    /// Index into the loaded `cellXfs` table.
    pub style_id: Option<usize>,
}

impl ReadOnlyCell {
    /// Build a cell from a raw XML value.
    pub fn new(
        row: Option<u32>,
        column: Option<String>,
        value: Option<String>,
        data_type: crate::cell::cell::DataType,
        style_id: Option<usize>,
    ) -> Self {
        ReadOnlyCell {
            row,
            column,
            raw: value,
            data_type,
            style_id,
        }
    }

    /// The coordinate string; empty cells have none.
    pub fn coordinate(&self) -> Result<String> {
        match (self.row, &self.column) {
            // The column letters come first, as in an A1 reference.
            (Some(row), Some(column)) => Ok(format!("{column}{row}")),
            _ => Err(Error::Attribute(
                "Empty cells have no coordinates".to_string(),
            )),
        }
    }

    /// The raw value, never converted.
    pub fn internal_value(&self) -> Option<&str> {
        self.raw.as_deref()
    }

    /// The number format code from the style table.
    pub fn number_format<'t>(&self, tables: &'t ReadOnlyTables) -> Option<&'t str> {
        let id = self.style_id?;
        tables
            .style_table
            .get(id)
            .map(|style: &Style| style.number_format.format_code())
    }

    /// Whether this cell is a date-formatted numeric.
    pub fn is_date(&self, tables: &ReadOnlyTables) -> bool {
        self.data_type == crate::cell::cell::DataType::Numeric
            && crate::styles::numbers::is_date_format(self.number_format(tables))
    }

    /// The converted value.
    pub fn value(&self, tables: &ReadOnlyTables) -> Option<crate::cell::cell::CellValue> {
        use crate::cell::cell::{CellValue, DataType};
        let raw = self.raw.as_ref()?;
        Some(match self.data_type {
            DataType::Bool => CellValue::Bool(raw == "1"),
            DataType::InlineString | DataType::FormulaCacheString => CellValue::Text(raw.clone()),
            DataType::SharedString => {
                let index = raw.parse::<usize>().unwrap_or(usize::MAX);
                match tables.string_table.get(&index) {
                    Some(text) => CellValue::Text(text.clone()),
                    // A dangling index is not fatal; surface the raw text like the reader does.
                    None => CellValue::Text(raw.clone()),
                }
            }
            DataType::Error => CellValue::Error(raw.clone()),
            DataType::Formula => CellValue::Formula(format!("={raw}")),
            DataType::Numeric => {
                let number = raw.parse::<f64>().unwrap_or(f64::NAN);
                if self.is_date(tables) {
                    match from_excel(number, tables.base_date) {
                        ExcelDateTime::DateTime(dt) => CellValue::DateTime(dt),
                        ExcelDateTime::Date(d) => CellValue::Date(d),
                        ExcelDateTime::Time(t) => CellValue::Time(t),
                    }
                } else {
                    CellValue::Number(number)
                }
            }
        })
    }

    /// The value as a naive datetime, when this cell holds one.
    pub fn datetime(&self, tables: &ReadOnlyTables) -> Option<NaiveDateTime> {
        match self.value(tables)? {
            crate::cell::cell::CellValue::DateTime(dt) => Some(dt),
            _ => None,
        }
    }
}

/// A shared, immutable empty cell, matching the `EMPTY_CELL` singleton.
pub fn empty_cell() -> ReadOnlyCell {
    ReadOnlyCell::new(
        None,
        None,
        None,
        crate::cell::cell::DataType::SharedString,
        None,
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::cell::DataType;
    use crate::styles::numbers::NumberFormat;
    use crate::styles::Style;

    fn tables(base_date: BaseDate) -> ReadOnlyTables {
        let mut strings = HashMap::new();
        strings.insert(0usize, "hello".to_string());
        strings.insert(1usize, "world".to_string());

        let mut numeric = Style::new();
        numeric.number_format = NumberFormat::new();
        let mut dated = Style::new();
        dated.set_number_format_code("yyyy-mm-dd");

        ReadOnlyTables {
            string_table: Arc::new(strings),
            style_table: Arc::new(vec![numeric, dated]),
            base_date,
        }
    }

    #[test]
    fn shared_strings_resolve() {
        let t = tables(BaseDate::Windows1900);
        let cell = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("1".into()),
            DataType::SharedString,
            None,
        );
        assert_eq!(
            cell.value(&t),
            Some(crate::cell::cell::CellValue::text("world"))
        );
        assert_eq!(cell.coordinate().unwrap(), "A1");
    }

    #[test]
    fn booleans_and_numbers() {
        let t = tables(BaseDate::Windows1900);
        let yes = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("1".into()),
            DataType::Bool,
            None,
        );
        assert_eq!(
            yes.value(&t),
            Some(crate::cell::cell::CellValue::Bool(true))
        );
        let no = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("0".into()),
            DataType::Bool,
            None,
        );
        assert_eq!(
            no.value(&t),
            Some(crate::cell::cell::CellValue::Bool(false))
        );
        let num = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("12.5".into()),
            DataType::Numeric,
            Some(0),
        );
        assert_eq!(
            num.value(&t),
            Some(crate::cell::cell::CellValue::Number(12.5))
        );
    }

    #[test]
    fn dates_resolve_using_style() {
        let t = tables(BaseDate::Windows1900);
        let dated = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("40196".into()),
            DataType::Numeric,
            Some(1),
        );
        assert!(dated.is_date(&t));
        match dated.value(&t) {
            Some(crate::cell::cell::CellValue::DateTime(dt)) => {
                assert_eq!(dt.to_string(), "2010-01-18 00:00:00");
            }
            other => panic!("expected a datetime, got {other:?}"),
        }

        // Same serial, General format: stays a number.
        let plain = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("40196".into()),
            DataType::Numeric,
            Some(0),
        );
        assert!(!plain.is_date(&t));
        assert_eq!(
            plain.value(&t),
            Some(crate::cell::cell::CellValue::Number(40196.0))
        );
    }

    #[test]
    fn formulas_and_inline_strings() {
        let t = tables(BaseDate::Windows1900);
        let formula = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("SUM(B1:B2)".into()),
            DataType::Formula,
            None,
        );
        assert_eq!(
            formula.value(&t),
            Some(crate::cell::cell::CellValue::Formula("=SUM(B1:B2)".into()))
        );
        let inline = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("txt".into()),
            DataType::InlineString,
            None,
        );
        assert_eq!(
            inline.value(&t),
            Some(crate::cell::cell::CellValue::text("txt"))
        );
        let cached = ReadOnlyCell::new(
            Some(1),
            Some("A".into()),
            Some("txt".into()),
            DataType::FormulaCacheString,
            None,
        );
        assert_eq!(
            cached.value(&t),
            Some(crate::cell::cell::CellValue::text("txt"))
        );
    }

    #[test]
    fn empty_cell_has_no_coordinate() {
        let cell = empty_cell();
        assert!(cell.coordinate().is_err());
        assert!(cell.value(&tables(BaseDate::Windows1900)).is_none());
        assert_eq!(cell.internal_value(), None);
    }

    #[test]
    fn errors_surface_as_error_values() {
        let t = tables(BaseDate::Windows1900);
        let cell = ReadOnlyCell::new(
            Some(3),
            Some("C".into()),
            Some("#N/A".into()),
            DataType::Error,
            None,
        );
        assert_eq!(
            cell.value(&t),
            Some(crate::cell::cell::CellValue::Error("#N/A".into()))
        );
    }
}
