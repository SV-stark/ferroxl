//! Individual cells (`openpyxl/cell/cell.py`).
//!
//! The Python `Cell` reaches back to its parent worksheet for style lookups, the base
//! date and the guess-types flag. Rust cannot hold that reference without interior
//! mutability, so the context is passed explicitly to the methods that need it and the
//! worksheet owns style storage (see [`crate::worksheet::Worksheet`]).

// `from_str` is openpyxl's `classmethod from_str`, so the name is kept even though Rust
// would rather these implemented `FromStr`.
#![allow(clippy::should_implement_trait)]

use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use crate::date_time::{
    date_to_excel, from_excel, time_to_days, timedelta_to_days, to_excel, BaseDate, ExcelDateTime,
};
use crate::exceptions::{Error, Result};
use crate::styles::numbers::NumberFormat;
use regex::Regex;
use std::sync::OnceLock;

/// A cell's data type, mirroring `Cell.TYPE_*`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum DataType {
    /// Shared string (`t="s"`). Also the type used for empty cells.
    #[default]
    SharedString,
    /// Formula (`t="f"` is implied by the absence of a `t` attribute on formula cells).
    Formula,
    /// Numeric (`t="n"`).
    Numeric,
    /// Boolean (`t="b"`).
    Bool,
    /// Inline string (`t="inlineStr"`).
    InlineString,
    /// Error code (`t="e"`).
    Error,
    /// Cached string result of a formula (`t="str"`).
    FormulaCacheString,
}

impl DataType {
    /// The XML `t` attribute value for this type.
    ///
    /// `SharedString` doubles as the null type in openpyxl, hence the shared `"s"`.
    pub fn as_str(self) -> &'static str {
        match self {
            DataType::SharedString => "s",
            DataType::Formula => "f",
            DataType::Numeric => "n",
            DataType::Bool => "b",
            DataType::InlineString => "inlineStr",
            DataType::Error => "e",
            DataType::FormulaCacheString => "str",
        }
    }

    /// Parse a `t` attribute value.
    pub fn from_str(value: &str) -> Option<DataType> {
        Some(match value {
            "s" => DataType::SharedString,
            "f" => DataType::Formula,
            "n" | "" => DataType::Numeric,
            "b" => DataType::Bool,
            "inlineStr" => DataType::InlineString,
            "e" => DataType::Error,
            "str" => DataType::FormulaCacheString,
            _ => return None,
        })
    }

    /// All valid types, matching `Cell.VALID_TYPES`.
    pub const VALID_TYPES: [DataType; 8] = [
        DataType::SharedString,
        DataType::Formula,
        DataType::Numeric,
        DataType::Bool,
        DataType::SharedString,
        DataType::InlineString,
        DataType::Error,
        DataType::FormulaCacheString,
    ];
}

/// The seven error values Excel recognises (`Cell.ERROR_CODES`).
pub const ERROR_CODES: [&str; 7] = [
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
];

/// The maximum number of characters a cell string may hold.
pub const MAX_STRING_LENGTH: usize = 32767;

/// A value that can be stored in a cell.
///
/// Python distinguishes `int`, `float` and `Decimal` and keeps them apart when writing;
/// Rust collapses all three to `f64` and emits integral values without a decimal point,
/// which is byte-identical to what openpyxl's `repr`/`str` produced.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell.
    None,
    /// Boolean value.
    Bool(bool),
    /// Numeric value.
    Number(f64),
    /// Text value.
    Text(String),
    /// An Excel error code such as `#N/A`.
    Error(String),
    /// A formula, stored with its leading `=`.
    Formula(String),
    /// A datetime, stored as its Excel serial.
    DateTime(NaiveDateTime),
    /// A date, stored as its Excel serial.
    Date(NaiveDate),
    /// A time of day, stored as a fraction of a day.
    Time(NaiveTime),
    /// An elapsed duration, stored as a fraction of a day.
    Duration(Duration),
}

impl CellValue {
    /// Build a text value.
    pub fn text(value: impl Into<String>) -> Self {
        CellValue::Text(value.into())
    }

    /// Build a numeric value.
    pub fn number(value: f64) -> Self {
        CellValue::Number(value)
    }

    /// Build a formula value from an expression with or without a leading `=`.
    pub fn formula(expression: impl AsRef<str>) -> Self {
        let expression = expression.as_ref();
        if expression.starts_with('=') {
            CellValue::Formula(expression.to_string())
        } else {
            CellValue::Formula(format!("={expression}"))
        }
    }

    /// Whether this value is empty.
    pub fn is_empty(&self) -> bool {
        matches!(self, CellValue::None) || matches!(self, CellValue::Text(t) if t.is_empty())
    }

    /// The string form used by the writer for text-ish values.
    pub fn as_text(&self) -> Option<&str> {
        match self {
            CellValue::Text(t) | CellValue::Error(t) | CellValue::Formula(t) => Some(t),
            _ => None,
        }
    }

    /// The numeric form used by the writer.
    pub fn as_number(&self) -> Option<f64> {
        match self {
            CellValue::Number(v) => Some(*v),
            CellValue::Bool(v) => Some(if *v { 1.0 } else { 0.0 }),
            _ => None,
        }
    }

    /// Whether this is one of the date-bearing types.
    pub fn is_temporal(&self) -> bool {
        matches!(
            self,
            CellValue::DateTime(_)
                | CellValue::Date(_)
                | CellValue::Time(_)
                | CellValue::Duration(_)
        )
    }
}

impl From<bool> for CellValue {
    fn from(v: bool) -> Self {
        CellValue::Bool(v)
    }
}

impl From<f64> for CellValue {
    fn from(v: f64) -> Self {
        CellValue::Number(v)
    }
}

impl From<i64> for CellValue {
    fn from(v: i64) -> Self {
        CellValue::Number(v as f64)
    }
}

impl From<i32> for CellValue {
    fn from(v: i32) -> Self {
        CellValue::Number(v as f64)
    }
}

impl From<&str> for CellValue {
    fn from(v: &str) -> Self {
        CellValue::Text(v.to_string())
    }
}

impl From<String> for CellValue {
    fn from(v: String) -> Self {
        CellValue::Text(v)
    }
}

impl From<NaiveDateTime> for CellValue {
    fn from(v: NaiveDateTime) -> Self {
        CellValue::DateTime(v)
    }
}

impl From<NaiveDate> for CellValue {
    fn from(v: NaiveDate) -> Self {
        CellValue::Date(v)
    }
}

impl From<NaiveTime> for CellValue {
    fn from(v: NaiveTime) -> Self {
        CellValue::Time(v)
    }
}

impl From<Duration> for CellValue {
    fn from(v: Duration) -> Self {
        CellValue::Duration(v)
    }
}

impl<T> From<Option<T>> for CellValue
where
    CellValue: From<T>,
{
    fn from(v: Option<T>) -> Self {
        match v {
            Some(inner) => inner.into(),
            None => CellValue::None,
        }
    }
}

/// Ambient information a cell needs when binding a value (`Cell.parent` in Python).
#[derive(Debug, Clone, Copy)]
pub struct CellContext {
    /// The workbook's date system.
    pub base_date: BaseDate,
    /// Whether string values should be sniffed for numbers, percentages and times.
    pub guess_types: bool,
}

impl Default for CellContext {
    fn default() -> Self {
        CellContext {
            base_date: BaseDate::Windows1900,
            guess_types: false,
        }
    }
}

/// Describes cell-associated properties: style, type, value and address.
#[derive(Debug, Clone, PartialEq)]
pub struct Cell {
    /// Column letters, always upper case.
    pub column: String,
    /// 1-based row number.
    pub row: u32,
    /// The stored value. Date-like values are kept as Excel serials.
    value: CellValue,
    /// The raw value as supplied, before serialisation.
    ///
    /// openpyxl keeps dates in `_value` as serials and reconstructs them on read, losing
    /// whether the user passed a `date` or a `datetime`. Preserving the original keeps
    /// `Date` and `DateTime` distinguishable on a load/save round-trip.
    original: CellValue,
    /// The inferred or explicitly assigned type.
    pub data_type: DataType,
    /// Index into the style table.
    pub xf_index: usize,
    /// Whether the cell is part of a merge (and therefore blank).
    pub merged: bool,
    /// The relationship id of this cell's hyperlink, if any.
    pub hyperlink_rel_id: Option<String>,
    /// The hyperlink target, if any.
    pub hyperlink_target: Option<String>,
    /// The comment attached to this cell.
    pub comment: Option<crate::comments::Comment>,
    /// Number format code last applied to this cell.
    format_code: Option<String>,
    /// Formula attributes read from the file (shared formula bookkeeping).
    pub formula_attributes: Option<FormulaAttributes>,
}

/// Attributes recorded for a formula cell (`ws.formula_attributes[coord]`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FormulaAttributes {
    /// The `t` attribute, e.g. `shared`.
    pub formula_type: Option<String>,
    /// Shared-formula group index.
    pub si: Option<String>,
    /// Range for a shared formula.
    pub reference: Option<String>,
}

impl Cell {
    /// Create an empty cell at the given column and row.
    pub fn new(column: impl Into<String>, row: u32) -> Self {
        let column = column.into().to_uppercase();
        Cell {
            column,
            row,
            value: CellValue::None,
            original: CellValue::None,
            data_type: DataType::SharedString,
            xf_index: 0,
            merged: false,
            hyperlink_rel_id: None,
            hyperlink_target: None,
            comment: None,
            format_code: None,
            formula_attributes: None,
        }
    }

    /// The coordinate string for this cell, e.g. `B12`.
    pub fn coordinate(&self) -> String {
        format!("{}{}", self.column, self.row)
    }

    /// The 1-based column index.
    pub fn column_index(&self) -> Result<u32> {
        crate::cell::utils::column_index_from_string(&self.column)
    }

    /// The stored value, exactly as it will be written to XML.
    pub fn internal_value(&self) -> &CellValue {
        &self.value
    }

    /// The value as the user supplied it (dates still as date types).
    pub fn original_value(&self) -> &CellValue {
        &self.original
    }

    /// Whether the cell holds a value.
    pub fn has_value(&self) -> bool {
        !matches!(self.value, CellValue::None)
    }

    /// Set the value and infer type and display options.
    ///
    /// Returns the number format that must be applied to the cell's style, mirroring the
    /// `self.number_format = ...` side effects of the Python `_cast_*` helpers.
    pub fn set_value(&mut self, value: impl Into<CellValue>, ctx: CellContext) -> Option<String> {
        let value = value.into();
        self.data_type = Self::data_type_for_value(&value);

        if matches!(value, CellValue::None) {
            // An empty cell holds an empty string with the shared-string type, which is
            // what openpyxl stores and what makes the cell write as `<c/>`.
            self.store(CellValue::Text(String::new()), DataType::SharedString);
            return None;
        }

        if ctx.guess_types && self.data_type == DataType::SharedString {
            if let CellValue::Text(text) = &value {
                if self.cast_numeric(text) {
                    return None;
                }
                if self.cast_percentage(text) {
                    return Some(NumberFormat::FORMAT_PERCENTAGE.to_string());
                }
                if let Some(fmt) = self.cast_time(text) {
                    return Some(fmt);
                }
            }
        }

        if self.data_type == DataType::Numeric {
            if let Some(fmt) = self.cast_datetime(&value, ctx) {
                return Some(fmt);
            }
        }

        self.store(value.clone(), self.data_type);
        None
    }

    /// Coerce a value according to an explicit type (`Cell.set_explicit_value`).
    pub fn set_explicit_value(&mut self, value: CellValue, data_type: DataType) -> Result<()> {
        // Only strings and errors are sanitised; a formula keeps its own variant so that
        // `data_type == Formula` implies `CellValue::Formula`.
        let value = match value {
            CellValue::Text(text) | CellValue::Error(text) => CellValue::Text(check_string(&text)?),
            CellValue::Formula(text) => CellValue::Formula(check_string(&text)?),
            other => other,
        };
        self.store(value, data_type);
        Ok(())
    }

    fn store(&mut self, stored: CellValue, data_type: DataType) {
        self.original = stored.clone();
        // Temporal values live in the serial domain while the cell is stored.
        let serialised = match &stored {
            CellValue::DateTime(_)
            | CellValue::Date(_)
            | CellValue::Time(_)
            | CellValue::Duration(_) => CellValue::Number(serial_for(&stored)),
            _ => stored,
        };
        self.value = serialised;
        self.data_type = data_type;
    }

    /// Given a value, infer the correct data type.
    pub fn data_type_for_value(value: &CellValue) -> DataType {
        match value {
            CellValue::None => DataType::SharedString,
            CellValue::Bool(_) => DataType::Bool,
            CellValue::Number(_) => DataType::Numeric,
            CellValue::DateTime(_)
            | CellValue::Date(_)
            | CellValue::Time(_)
            | CellValue::Duration(_) => DataType::Numeric,
            CellValue::Text(text) => {
                if text.starts_with('=') {
                    DataType::Formula
                } else if ERROR_CODES.contains(&text.as_str()) {
                    DataType::Error
                } else {
                    DataType::SharedString
                }
            }
            CellValue::Error(_) => DataType::Error,
            CellValue::Formula(_) => DataType::Formula,
        }
    }

    /// Explicitly convert a string to a numeric value.
    pub fn cast_numeric(&mut self, text: &str) -> bool {
        if number_regex().is_match(text) {
            let parsed = text
                .parse::<f64>()
                .unwrap_or_else(|_| text.parse::<i64>().map(|v| v as f64).unwrap_or(0.0));
            self.store(CellValue::Number(parsed), DataType::Numeric);
            return true;
        }
        false
    }

    /// Convert a string to a number formatted as a percentage.
    pub fn cast_percentage(&mut self, text: &str) -> bool {
        let Some(caps) = percent_regex().captures(text) else {
            return false;
        };
        let number = caps.name("number").map(|m| m.as_str()).unwrap_or("0");
        let value: f64 = number.trim().parse().unwrap_or(0.0);
        self.store(CellValue::Number(value / 100.0), DataType::Numeric);
        true
    }

    /// Convert a string to a number formatted as a date or time.
    ///
    /// The Python module recognises two shapes: `HH:MM[:SS]` (with an odd two-or-three
    /// digit hour) and `MM:SS.ffffff`. They are matched by separate patterns here rather
    /// than one alternation, so the capture indices stay readable.
    pub fn cast_time(&mut self, text: &str) -> Option<String> {
        let (time, format) = match clock_regex().captures(text) {
            Some(caps) => {
                let hour = caps.name("hour")?.as_str().parse::<u32>().ok()?;
                let minute = caps.name("minute")?.as_str().parse::<u32>().ok()?;
                let second = caps
                    .name("second")
                    .map(|m| m.as_str().parse::<u32>().unwrap_or(0))
                    .unwrap_or(0);
                let time = NaiveTime::from_hms_opt(hour, minute, second)?;
                let format = if caps.name("second").is_some() {
                    NumberFormat::FORMAT_DATE_TIME6
                } else {
                    NumberFormat::FORMAT_DATE_TIME3
                };
                (time, format)
            }
            None => {
                let caps = subsecond_regex().captures(text)?;
                let minute = caps.name("minute")?.as_str().parse::<u32>().ok()?;
                let second = caps.name("second")?.as_str().parse::<u32>().ok()?;
                // Python truncates the string to 12 characters, capping at 6 fractional
                // digits.
                let micros = caps
                    .name("microsecond")?
                    .as_str()
                    .parse::<u32>()
                    .ok()?
                    .min(999_999);
                let time = NaiveTime::from_hms_micro_opt(0, minute, second, micros)?;
                (time, NumberFormat::FORMAT_DATE_TIME5)
            }
        };

        self.store(CellValue::Number(time_to_days(time)), DataType::Numeric);
        Some(format.to_string())
    }

    /// Convert a temporal value to its serial, returning the format to apply.
    ///
    /// The serial is what gets stored, but the original temporal value is kept so that
    /// reading the cell back reports a `date` rather than a `datetime` when that is what
    /// the caller wrote.
    fn cast_datetime(&mut self, value: &CellValue, ctx: CellContext) -> Option<String> {
        let (serial, format) = match value {
            CellValue::DateTime(dt) => (
                to_excel(*dt, ctx.base_date),
                // A datetime needs no extra format; the caller decides.
                None,
            ),
            CellValue::Date(date) => (
                date_to_excel(*date, ctx.base_date),
                Some(NumberFormat::FORMAT_DATE_YYYYMMDD2.to_string()),
            ),
            CellValue::Time(time) => (
                time_to_days(*time),
                Some(NumberFormat::FORMAT_DATE_TIME6.to_string()),
            ),
            CellValue::Duration(delta) => (
                timedelta_to_days(*delta),
                Some(NumberFormat::FORMAT_DATE_TIMEDELTA.to_string()),
            ),
            _ => return None,
        };
        self.store(CellValue::Number(serial), DataType::Numeric);
        self.original = value.clone();
        format
    }

    /// Reconstruct the display value from the stored serial.
    ///
    /// Requires the cell's number format, because only date-formatted numerics are
    /// converted back into datetimes, and the workbook's date system, because the same
    /// serial means different days in the 1900 and 1904 calendars -- a 1904 workbook read
    /// with the 1900 epoch is out by 1462 days.
    ///
    /// [`display_value_in`](Self::display_value_in) takes the context; this one is the
    /// 1900 default, kept because a caller that has no workbook in hand genuinely has no
    /// other option.
    pub fn display_value(&self, number_format: Option<&str>) -> CellValue {
        self.display_value_in(number_format, BaseDate::Windows1900)
    }

    /// [`display_value`](Self::display_value) against a known base date.
    pub fn display_value_in(&self, number_format: Option<&str>, base: BaseDate) -> CellValue {
        if self.is_date(number_format) {
            // A value written as a date keeps its original Rust type, so a `Date` stays a
            // `Date` and a `DateTime` stays a `DateTime`.
            if matches!(
                self.original,
                CellValue::Date(_) | CellValue::DateTime(_) | CellValue::Time(_)
            ) {
                return self.original.clone();
            }
            if let CellValue::Number(serial) = self.value {
                return match from_excel(serial, base) {
                    ExcelDateTime::DateTime(dt) => CellValue::DateTime(dt),
                    ExcelDateTime::Date(d) => CellValue::Date(d),
                    ExcelDateTime::Time(t) => CellValue::Time(t),
                };
            }
        }
        self.value.clone()
    }

    /// Whether the value is formatted as a date.
    pub fn is_date(&self, number_format: Option<&str>) -> bool {
        crate::styles::numbers::is_date_format(number_format)
            && self.data_type == DataType::Numeric
            && matches!(self.value, CellValue::Number(_))
    }

    /// The hyperlink target, or an empty string when there is none.
    pub fn hyperlink(&self) -> &str {
        self.hyperlink_target.as_deref().unwrap_or("")
    }

    /// Record a number format on this cell.
    ///
    /// Styles live on the worksheet, so this only records the most recently applied code;
    /// [`Cell::display_value`] takes the authoritative value from the caller.
    pub fn set_format_code(&mut self, code: &str) {
        self.format_code = Some(code.to_string());
    }

    /// The number format code last set on this cell.
    pub fn format_code(&self) -> Option<&str> {
        self.format_code.as_deref()
    }

    /// Set the hyperlink target and, when the cell is empty, its display value.
    pub fn set_hyperlink(&mut self, target: &str) {
        self.hyperlink_target = Some(target.to_string());
        if matches!(self.value, CellValue::None) {
            self.store(CellValue::Text(target.to_string()), DataType::SharedString);
        }
    }

    /// A relative cell location, given as `(row_offset, column_offset)`.
    pub fn offset(&self, row: i64, column: i64) -> Result<(String, u32)> {
        let base_col = self.column_index()? as i64;
        let new_col = base_col + column;
        if new_col < 1 {
            return Err(Error::ColumnStringIndex(format!(
                "Invalid column index {new_col}"
            )));
        }
        let new_row = self.row as i64 + row;
        if new_row < 1 {
            return Err(Error::CellCoordinates(format!("There is no row {new_row}")));
        }
        crate::cell::utils::coordinate_from_index(new_col as u32, new_row as u32).map(|c| {
            let (col, r) = crate::cell::utils::coordinate_from_string(&c)
                .expect("generated coordinate is valid");
            (col, r)
        })
    }
}

fn serial_for(value: &CellValue) -> f64 {
    match value {
        CellValue::DateTime(dt) => to_excel(*dt, BaseDate::Windows1900),
        CellValue::Date(d) => date_to_excel(*d, BaseDate::Windows1900),
        CellValue::Time(t) => time_to_days(*t),
        CellValue::Duration(d) => timedelta_to_days(*d),
        CellValue::Number(v) => *v,
        _ => 0.0,
    }
}

/// Check string coding, length and line-break characters.
///
/// Strings longer than 32,767 characters are truncated, `\r\n` is normalised to `\n`,
/// and control characters that Excel rejects raise an error.
pub fn check_string(value: &str) -> Result<String> {
    let mut value: String = value.chars().take(MAX_STRING_LENGTH).collect();
    for ch in value.chars() {
        let code = ch as u32;
        let illegal = matches!(code, 0..=8 | 11 | 12 | 14..=31);
        if illegal {
            return Err(Error::IllegalCharacter(ch));
        }
    }
    value = value.replace("\r\n", "\n");
    Ok(value)
}

fn number_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(r"^-?([\d]|[\d]+\.[\d]*|\.[\d]+|[1-9][\d]+\.?[\d]*)((E|e)-?[\d]+)?$")
            .expect("valid regex")
    })
}

fn percent_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r"^\-?(?P<number>[0-9]*\.?[0-9]*\s?)\%$").expect("valid regex"))
}

/// Matches `HH:MM` and `HH:MM:SS`.
///
/// The hour group is `[0-1]{0,1}[0-9]{2}`, reproducing the original pattern: it accepts
/// two or three digits, so `13:55` parses but `9:30` does not.
fn clock_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(
            r"^(?P<hour>[0-1]{0,1}[0-9]{2}):(?P<minute>[0-5][0-9]):?(?P<second>[0-5][0-9])?$",
        )
        .expect("valid regex")
    })
}

/// Matches `MM:SS.ffffff`.
fn subsecond_regex() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| {
        Regex::new(r"^(?P<minute>[0-5][0-9]):(?P<second>[0-5][0-9])\.(?P<microsecond>\d{1,6})$")
            .expect("valid regex")
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx() -> CellContext {
        CellContext::default()
    }

    #[test]
    fn coordinate_and_column_index() {
        let cell = Cell::new("b", 12);
        assert_eq!(cell.column, "B");
        assert_eq!(cell.coordinate(), "B12");
        assert_eq!(cell.column_index().unwrap(), 2);
    }

    #[test]
    fn type_inference() {
        assert_eq!(
            Cell::data_type_for_value(&CellValue::None),
            DataType::SharedString
        );
        assert_eq!(
            Cell::data_type_for_value(&CellValue::Bool(true)),
            DataType::Bool
        );
        assert_eq!(
            Cell::data_type_for_value(&CellValue::Number(1.0)),
            DataType::Numeric
        );
        assert_eq!(
            Cell::data_type_for_value(&CellValue::text("=SUM(A1:A2)")),
            DataType::Formula
        );
        assert_eq!(
            Cell::data_type_for_value(&CellValue::text("#N/A")),
            DataType::Error
        );
        assert_eq!(
            Cell::data_type_for_value(&CellValue::text("hello")),
            DataType::SharedString
        );
        let date = NaiveDate::from_ymd_opt(2010, 1, 18).unwrap();
        assert_eq!(
            Cell::data_type_for_value(&CellValue::Date(date)),
            DataType::Numeric
        );
    }

    #[test]
    fn setting_values_stores_serials_for_dates() {
        let mut cell = Cell::new("A", 1);
        let date = NaiveDate::from_ymd_opt(2010, 1, 18).unwrap();
        let fmt = cell.set_value(CellValue::Date(date), ctx());
        assert_eq!(fmt.as_deref(), Some("yyyy-mm-dd"));
        assert_eq!(cell.data_type, DataType::Numeric);
        assert_eq!(*cell.internal_value(), CellValue::Number(40196.0));
        assert_eq!(*cell.original_value(), CellValue::Date(date));
    }

    #[test]
    fn empty_value_clears_cell() {
        let mut cell = Cell::new("A", 1);
        cell.set_value(CellValue::text("x"), ctx());
        cell.set_value(CellValue::None, ctx());
        assert_eq!(*cell.internal_value(), CellValue::Text(String::new()));
        assert_eq!(cell.data_type, DataType::SharedString);
    }

    #[test]
    fn guess_types_casts_numbers_percentages_and_times() {
        let mut cell = Cell::new("A", 1);
        let gctx = CellContext {
            guess_types: true,
            ..CellContext::default()
        };

        cell.set_value(CellValue::text("12"), gctx);
        assert_eq!(cell.data_type, DataType::Numeric);
        assert_eq!(*cell.internal_value(), CellValue::Number(12.0));

        let fmt = cell.set_value(CellValue::text("50%"), gctx);
        assert_eq!(fmt.as_deref(), Some("0%"));
        assert_eq!(*cell.internal_value(), CellValue::Number(0.5));

        let fmt = cell.set_value(CellValue::text("13:55:12"), gctx);
        assert_eq!(fmt.as_deref(), Some("h:mm:ss"));
        assert!(matches!(cell.internal_value(), CellValue::Number(v) if *v > 0.5));

        let fmt = cell.set_value(CellValue::text("13:55"), gctx);
        assert_eq!(fmt.as_deref(), Some("h:mm"));

        // Without guessing, strings stay strings.
        let mut plain = Cell::new("A", 1);
        plain.set_value(CellValue::text("12"), ctx());
        assert_eq!(plain.data_type, DataType::SharedString);
    }

    #[test]
    fn illegal_characters_rejected_and_strings_truncated() {
        assert!(matches!(
            check_string("bad\u{0}char"),
            Err(Error::IllegalCharacter(_))
        ));
        let long = "x".repeat(40_000);
        assert_eq!(check_string(&long).unwrap().len(), MAX_STRING_LENGTH);
        assert_eq!(check_string("a\r\nb").unwrap(), "a\nb");
    }

    /// The same serial means different days in the 1900 and 1904 calendars -- 1462 days apart --
    /// so a 1904 workbook read against the 1900 default epoch returns dates four years early.
    /// Found by running `tests/data/genuine/mac_date.xlsx`, a real Mac Excel file, through
    /// openpyxl and this library and comparing.
    #[test]
    fn display_value_honours_the_1904_date_system() {
        let mut cell = Cell::new("A", 1);
        // 41184 is 2016-10-03 counted from 1904-01-01.
        cell.set_explicit_value(CellValue::Number(41184.0), DataType::Numeric)
            .unwrap();

        // openpyxl hands back a `datetime.datetime` for this cell, not a `date`, so the
        // reconstruction is a `DateTime` too.
        assert_eq!(
            cell.display_value_in(Some("dd/mm/yyyy"), BaseDate::Mac1904),
            CellValue::DateTime(
                chrono::NaiveDate::from_ymd_opt(2016, 10, 3)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
        // Against the 1900 epoch the same serial is four years earlier, which is the bug.
        assert_eq!(
            cell.display_value_in(Some("dd/mm/yyyy"), BaseDate::Windows1900),
            CellValue::DateTime(
                chrono::NaiveDate::from_ymd_opt(2012, 10, 2)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
        // And the 1900-default entry point is unchanged for a 1900 workbook.
        assert_eq!(
            cell.display_value(Some("dd/mm/yyyy")),
            CellValue::DateTime(
                chrono::NaiveDate::from_ymd_opt(2012, 10, 2)
                    .unwrap()
                    .and_hms_opt(0, 0, 0)
                    .unwrap()
            )
        );
        // A non-date format is untouched either way.
        assert_eq!(
            cell.display_value_in(Some("General"), BaseDate::Mac1904),
            CellValue::Number(41184.0)
        );
    }

    #[test]
    fn display_value_round_trips_dates() {
        let mut cell = Cell::new("A", 1);
        let dt = NaiveDate::from_ymd_opt(2010, 1, 18)
            .unwrap()
            .and_hms_opt(14, 15, 20)
            .unwrap();
        cell.set_value(CellValue::DateTime(dt), ctx());
        cell.set_format_code("yyyy-mm-dd");
        assert_eq!(
            cell.display_value(Some("yyyy-mm-dd")),
            CellValue::DateTime(dt)
        );
        // A General-formatted numeric is not a date.
        assert!(matches!(
            cell.display_value(Some("General")),
            CellValue::Number(_)
        ));
    }

    #[test]
    fn data_type_parsing_round_trips() {
        for dt in [
            DataType::SharedString,
            DataType::Formula,
            DataType::Numeric,
            DataType::Bool,
            DataType::InlineString,
            DataType::Error,
            DataType::FormulaCacheString,
        ] {
            assert_eq!(DataType::from_str(dt.as_str()), Some(dt));
        }
        assert_eq!(DataType::from_str("n"), Some(DataType::Numeric));
        assert_eq!(DataType::from_str(""), Some(DataType::Numeric));
        assert_eq!(DataType::from_str("bogus"), None);
    }

    #[test]
    fn offsets_and_hyperlinks() {
        let cell = Cell::new("B", 12);
        assert_eq!(cell.offset(0, 0).unwrap(), ("B".to_string(), 12));
        assert_eq!(cell.offset(1, 1).unwrap(), ("C".to_string(), 13));
        assert!(cell.offset(0, -2).is_err());

        let mut cell = Cell::new("A", 1);
        cell.set_hyperlink("http://example.com");
        assert_eq!(cell.hyperlink(), "http://example.com");
        assert_eq!(
            *cell.internal_value(),
            CellValue::text("http://example.com")
        );
    }

    #[test]
    fn formula_helper_adds_equals() {
        assert_eq!(
            CellValue::formula("SUM(A1)"),
            CellValue::Formula("=SUM(A1)".into())
        );
        assert_eq!(
            CellValue::formula("=SUM(A1)"),
            CellValue::Formula("=SUM(A1)".into())
        );
    }
}
