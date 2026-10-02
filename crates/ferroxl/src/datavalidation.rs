//! Data validation (`openpyxl/datavalidation.py`).
//!
//! Validation rules are stored per worksheet and serialised into `<dataValidations>`.
//! Cell addresses are collapsed into ranges, so that a column of validated cells becomes
//! a single `A1:A10` style `sqref`.

// `from_str` is openpyxl's `classmethod from_str`, so the name is kept even though Rust
// would rather these implemented `FromStr`.
#![allow(clippy::should_implement_trait)]

use std::collections::BTreeMap;

use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::Result;

/// The validation kinds (`ST_DataValidationType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ValidationType {
    /// No validation.
    None,
    /// Whole number.
    Whole,
    /// Decimal.
    Decimal,
    /// Value list.
    List,
    /// Date.
    Date,
    /// Time.
    Time,
    /// Text length.
    TextLength,
    /// Custom formula.
    Custom,
}

impl ValidationType {
    /// The XML token.
    pub fn as_str(self) -> &'static str {
        match self {
            ValidationType::None => "none",
            ValidationType::Whole => "whole",
            ValidationType::Decimal => "decimal",
            ValidationType::List => "list",
            ValidationType::Date => "date",
            ValidationType::Time => "time",
            ValidationType::TextLength => "textLength",
            ValidationType::Custom => "custom",
        }
    }

    /// Parse an XML token.
    pub fn from_str(value: &str) -> Option<Self> {
        Some(match value {
            "none" => ValidationType::None,
            "whole" => ValidationType::Whole,
            "decimal" => ValidationType::Decimal,
            "list" => ValidationType::List,
            "date" => ValidationType::Date,
            "time" => ValidationType::Time,
            "textLength" => ValidationType::TextLength,
            "custom" => ValidationType::Custom,
            _ => return None,
        })
    }
}

/// The comparison operators (`ST_DataValidationOperator`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ValidationOperator {
    /// Between.
    Between,
    /// Not between.
    NotBetween,
    /// Equal.
    Equal,
    /// Not equal.
    NotEqual,
    /// Less than.
    LessThan,
    /// Less than or equal.
    LessThanOrEqual,
    /// Greater than.
    GreaterThan,
    /// Greater than or equal.
    GreaterThanOrEqual,
}

impl ValidationOperator {
    /// The XML token.
    pub fn as_str(self) -> &'static str {
        match self {
            ValidationOperator::Between => "between",
            ValidationOperator::NotBetween => "notBetween",
            ValidationOperator::Equal => "equal",
            ValidationOperator::NotEqual => "notEqual",
            ValidationOperator::LessThan => "lessThan",
            ValidationOperator::LessThanOrEqual => "lessThanOrEqual",
            ValidationOperator::GreaterThan => "greaterThan",
            ValidationOperator::GreaterThanOrEqual => "greaterThanOrEqual",
        }
    }

    /// Parse an XML token.
    pub fn from_str(value: &str) -> Option<Self> {
        Some(match value {
            "between" => ValidationOperator::Between,
            "notBetween" => ValidationOperator::NotBetween,
            "equal" => ValidationOperator::Equal,
            "notEqual" => ValidationOperator::NotEqual,
            "lessThan" => ValidationOperator::LessThan,
            "lessThanOrEqual" => ValidationOperator::LessThanOrEqual,
            "greaterThan" => ValidationOperator::GreaterThan,
            "greaterThanOrEqual" => ValidationOperator::GreaterThanOrEqual,
            _ => return None,
        })
    }
}

/// The error presentation styles (`ST_DataValidationErrorStyle`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ValidationErrorStyle {
    /// Stop with a hard error.
    Stop,
    /// Show a warning.
    Warning,
    /// Show an informational message.
    Information,
}

impl ValidationErrorStyle {
    /// The XML token.
    pub fn as_str(self) -> &'static str {
        match self {
            ValidationErrorStyle::Stop => "stop",
            ValidationErrorStyle::Warning => "warning",
            ValidationErrorStyle::Information => "information",
        }
    }

    /// Parse an XML token.
    pub fn from_str(value: &str) -> Option<Self> {
        Some(match value {
            "stop" => ValidationErrorStyle::Stop,
            "warning" => ValidationErrorStyle::Warning,
            "information" => ValidationErrorStyle::Information,
            _ => return None,
        })
    }
}

/// The IME input modes (`ST_DataValidationImeMode`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImeMode {
    /// No control.
    NoControl,
    /// Off.
    Off,
    /// On.
    On,
    /// Disabled.
    Disabled,
    /// Hiragana.
    Hiragana,
    /// Full-width Katakana.
    FullKatakana,
    /// Half-width Katakana.
    HalfKatakana,
    /// Full-width alpha.
    FullAlpha,
    /// Half-width alpha.
    HalfAlpha,
    /// Full-width Hangul.
    FullHangul,
    /// Half-width Hangul.
    HalfHangul,
}

impl ImeMode {
    /// The XML token.
    pub fn as_str(self) -> &'static str {
        match self {
            ImeMode::NoControl => "noControl",
            ImeMode::Off => "off",
            ImeMode::On => "on",
            ImeMode::Disabled => "disabled",
            ImeMode::Hiragana => "hiragana",
            ImeMode::FullKatakana => "fullKatakana",
            ImeMode::HalfKatakana => "halfKatakana",
            ImeMode::FullAlpha => "fullAlpha",
            ImeMode::HalfAlpha => "halfAlpha",
            ImeMode::FullHangul => "fullHangul",
            ImeMode::HalfHangul => "halfHangul",
        }
    }

    /// Parse an XML token.
    pub fn from_str(value: &str) -> Option<Self> {
        Some(match value {
            "noControl" => ImeMode::NoControl,
            "off" => ImeMode::Off,
            "on" => ImeMode::On,
            "disabled" => ImeMode::Disabled,
            "hiragana" => ImeMode::Hiragana,
            "fullKatakana" => ImeMode::FullKatakana,
            "halfKatakana" => ImeMode::HalfKatakana,
            "fullAlpha" => ImeMode::FullAlpha,
            "halfAlpha" => ImeMode::HalfAlpha,
            "fullHangul" => ImeMode::FullHangul,
            "halfHangul" => ImeMode::HalfHangul,
            _ => return None,
        })
    }
}

/// A data-validation rule attached to one or more cells.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DataValidation {
    /// The kind of validation.
    pub validation_type: ValidationType,
    /// The comparison operator.
    pub operator: Option<ValidationOperator>,
    /// The first formula (`formula1`).
    pub formula1: String,
    /// The second formula (`formula2`).
    pub formula2: String,
    /// Whether empty cells are accepted.
    pub allow_blank: bool,
    /// Error presentation style.
    pub error_style: Option<ValidationErrorStyle>,
    /// IME input mode.
    pub ime_mode: Option<ImeMode>,
    /// Extra attributes, overriding defaults where present.
    pub attr_map: BTreeMap<String, String>,
    /// Individual cell coordinates covered by this rule.
    pub cells: Vec<String>,
    /// Pre-built ranges included verbatim in the `sqref`.
    pub ranges: Vec<String>,
}

impl Default for DataValidation {
    fn default() -> Self {
        DataValidation::new(ValidationType::None, None, None, None, false)
    }
}

impl DataValidation {
    /// Build a rule.
    pub fn new(
        validation_type: ValidationType,
        operator: Option<ValidationOperator>,
        formula1: Option<&str>,
        formula2: Option<&str>,
        allow_blank: bool,
    ) -> Self {
        let mut attr_map = BTreeMap::new();
        attr_map.insert("showInputMessage".to_string(), "1".to_string());
        attr_map.insert("showErrorMessage".to_string(), "1".to_string());
        DataValidation {
            validation_type,
            operator,
            formula1: formula1.unwrap_or_default().to_string(),
            formula2: formula2.unwrap_or_default().to_string(),
            allow_blank,
            error_style: None,
            ime_mode: None,
            attr_map,
            cells: Vec::new(),
            ranges: Vec::new(),
        }
    }

    /// Add a cell coordinate to the rule.
    pub fn add_cell(&mut self, coordinate: &str) {
        self.cells.push(coordinate.to_string());
    }

    /// Set a custom error message.
    pub fn set_error_message(&mut self, error: &str, error_title: &str) {
        self.attr_map
            .insert("errorTitle".to_string(), error_title.to_string());
        self.attr_map.insert("error".to_string(), error.to_string());
    }

    /// Set a custom prompt message.
    pub fn set_prompt_message(&mut self, prompt: &str, prompt_title: &str) {
        self.attr_map
            .insert("promptTitle".to_string(), prompt_title.to_string());
        self.attr_map
            .insert("prompt".to_string(), prompt.to_string());
    }

    /// The full attribute set written to `<dataValidation>`.
    pub fn attributes(&self) -> BTreeMap<String, String> {
        let mut attrs = self.attr_map.clone();
        attrs.insert(
            "type".to_string(),
            self.validation_type.as_str().to_string(),
        );
        attrs.insert(
            "allowBlank".to_string(),
            if self.allow_blank { "1" } else { "0" }.to_string(),
        );
        if let Some(operator) = self.operator {
            attrs.insert("operator".to_string(), operator.as_str().to_string());
        }
        if let Some(style) = self.error_style {
            attrs.insert("errorStyle".to_string(), style.as_str().to_string());
        }
        if let Some(mode) = self.ime_mode {
            attrs.insert("imeMode".to_string(), mode.as_str().to_string());
        }
        attrs.insert(
            "sqref".to_string(),
            collapse_cell_addresses(&self.cells, &self.ranges),
        );
        attrs
    }

    /// Read a `<dataValidation>` element back into a rule.
    ///
    /// Unknown attributes are kept in [`DataValidation::attr_map`] so a load/save cycle
    /// does not silently drop them.
    pub fn read(node: &crate::xml::functions::Element) -> Result<DataValidation> {
        let validation_type = node
            .get("type")
            .and_then(ValidationType::from_str)
            .unwrap_or(ValidationType::None);
        let operator = node.get("operator").and_then(ValidationOperator::from_str);
        let error_style = node
            .get("errorStyle")
            .and_then(ValidationErrorStyle::from_str);
        let ime_mode = node.get("imeMode").and_then(ImeMode::from_str);

        let mut formula1 = String::new();
        let mut formula2 = String::new();
        let mut cells = Vec::new();
        let mut ranges = Vec::new();
        for child in node.children() {
            let local = crate::xml::constants::local_name(&child.tag);
            let text = child.text.clone().unwrap_or_default();
            match local {
                "formula1" => formula1 = text,
                "formula2" => formula2 = text,
                "dataValidation" | "cell" => cells.push(text),
                _ => {}
            }
        }
        // `sqref` is the authoritative list; `cells` only holds what was spelled out.
        let sqref = node.get("sqref").unwrap_or_default();
        for part in sqref.split_whitespace() {
            if part.contains(':') {
                ranges.push(part.to_string());
            } else if !part.is_empty() {
                cells.push(part.to_string());
            }
        }

        let mut validation = DataValidation::new(
            validation_type,
            operator,
            Some(&formula1),
            Some(&formula2),
            node.get("allowBlank").map(|v| v == "1").unwrap_or(false),
        );
        validation.error_style = error_style;
        validation.ime_mode = ime_mode;
        validation.cells = cells;
        validation.ranges = ranges;
        for (key, value) in &node.attributes {
            let local = crate::xml::constants::local_name(key).to_string();
            if matches!(
                local.as_str(),
                "type"
                    | "operator"
                    | "allowBlank"
                    | "errorStyle"
                    | "imeMode"
                    | "sqref"
                    | "formula1"
                    | "formula2"
            ) {
                continue;
            }
            validation.attr_map.insert(local, value.clone());
        }
        Ok(validation)
    }
}

/// Collapse a collection of cell coordinates into an optimal set of ranges.
///
/// Cells are grouped by column, rows are grouped into contiguous runs, and each run
/// becomes either a single address or a `first:last` range. Only contiguous vertical runs
/// are collapsed, so `A1, A2, A3, B1, B2, B3` yields `A1:A3 B1:B3`.
pub fn collapse_cell_addresses(cells: &[String], input_ranges: &[String]) -> String {
    let mut by_column: BTreeMap<String, Vec<u32>> = BTreeMap::new();
    for cell in cells {
        if let Ok((column, row)) = coordinate_from_string(cell) {
            by_column.entry(column).or_default().push(row);
        }
    }
    let mut ranges: Vec<String> = input_ranges.to_vec();
    for (column, mut rows) in by_column {
        rows.sort_unstable();
        rows.dedup();
        let mut start = rows[0];
        let mut previous = rows[0];
        for &row in &rows[1..] {
            if row != previous + 1 {
                ranges.push(format_range(&column, start, previous));
                start = row;
            }
            previous = row;
        }
        ranges.push(format_range(&column, start, previous));
    }
    ranges.join(" ")
}

fn format_range(column: &str, first: u32, last: u32) -> String {
    if first == last {
        format!("{column}{first}")
    } else {
        format!("{column}{first}:{column}{last}")
    }
}

/// Collapse a contiguous range of column indices into a letter span, e.g. `A:D`.
pub fn column_span(min_col: u32, max_col: u32) -> Result<String> {
    if min_col == max_col {
        Ok(get_column_letter(min_col)?)
    } else {
        Ok(format!(
            "{}:{}",
            get_column_letter(min_col)?,
            get_column_letter(max_col)?
        ))
    }
}

/// Resolve a column name to its index, re-exported for callers building rules by hand.
pub fn resolve_column(name: &str) -> Result<u32> {
    column_index_from_string(name)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn coords(items: &[&str]) -> Vec<String> {
        items.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn single_cells_stay_uncollapsed() {
        assert_eq!(collapse_cell_addresses(&coords(&["A1"]), &[]), "A1");
    }

    #[test]
    fn vertical_runs_collapse_per_column() {
        assert_eq!(
            collapse_cell_addresses(&coords(&["A1", "A2", "A3", "B1", "B2", "B3"]), &[]),
            "A1:A3 B1:B3"
        );
    }

    #[test]
    fn gaps_split_runs() {
        assert_eq!(
            collapse_cell_addresses(&coords(&["A1", "A3", "A4", "A10"]), &[]),
            "A1 A3:A4 A10"
        );
    }

    #[test]
    fn input_ranges_are_preserved_first() {
        assert_eq!(
            collapse_cell_addresses(&coords(&["C1", "C2"]), &coords(&["A1:B5"])),
            "A1:B5 C1:C2"
        );
    }

    #[test]
    fn attributes_include_defaults_and_sqref() {
        let mut dv = DataValidation::new(
            ValidationType::Whole,
            Some(ValidationOperator::Between),
            Some("1"),
            Some("10"),
            true,
        );
        dv.add_cell("A1");
        dv.add_cell("A2");
        dv.set_error_message("bad", "Oops");
        let attrs = dv.attributes();
        assert_eq!(attrs.get("type").unwrap(), "whole");
        assert_eq!(attrs.get("operator").unwrap(), "between");
        assert_eq!(attrs.get("allowBlank").unwrap(), "1");
        assert_eq!(attrs.get("showInputMessage").unwrap(), "1");
        assert_eq!(attrs.get("sqref").unwrap(), "A1:A2");
        assert_eq!(attrs.get("errorTitle").unwrap(), "Oops");
    }

    #[test]
    fn token_round_trips() {
        for vt in [
            ValidationType::None,
            ValidationType::Whole,
            ValidationType::Decimal,
            ValidationType::List,
            ValidationType::Date,
            ValidationType::Time,
            ValidationType::TextLength,
            ValidationType::Custom,
        ] {
            assert_eq!(ValidationType::from_str(vt.as_str()), Some(vt));
        }
        for op in [
            ValidationOperator::Between,
            ValidationOperator::NotBetween,
            ValidationOperator::Equal,
            ValidationOperator::NotEqual,
            ValidationOperator::LessThan,
            ValidationOperator::LessThanOrEqual,
            ValidationOperator::GreaterThan,
            ValidationOperator::GreaterThanOrEqual,
        ] {
            assert_eq!(ValidationOperator::from_str(op.as_str()), Some(op));
        }
        for style in [
            ValidationErrorStyle::Stop,
            ValidationErrorStyle::Warning,
            ValidationErrorStyle::Information,
        ] {
            assert_eq!(ValidationErrorStyle::from_str(style.as_str()), Some(style));
        }
        for mode in [
            ImeMode::NoControl,
            ImeMode::Off,
            ImeMode::On,
            ImeMode::Disabled,
            ImeMode::Hiragana,
            ImeMode::FullKatakana,
            ImeMode::HalfKatakana,
            ImeMode::FullAlpha,
            ImeMode::HalfAlpha,
            ImeMode::FullHangul,
            ImeMode::HalfHangul,
        ] {
            assert_eq!(ImeMode::from_str(mode.as_str()), Some(mode));
        }
    }

    #[test]
    fn column_span_helper() {
        assert_eq!(column_span(1, 1).unwrap(), "A");
        assert_eq!(column_span(1, 4).unwrap(), "A:D");
    }
}
