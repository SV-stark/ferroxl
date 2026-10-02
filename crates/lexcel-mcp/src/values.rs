//! Moving cell values between JSON and [`lexcel::CellValue`].
//!
//! An agent's JSON is untyped, so a value can arrive as `42`, `"42"`, `true` or
//! `"=SUM(A1:A2)"`. This module makes those cases explicit rather than leaving each tool
//! to guess, and it renders values back out in a form a model can read at a glance.

use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use lexcel::CellValue;
use serde_json::{json, Value};

/// The error returned when a JSON value cannot become a cell value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValueError(pub String);

impl std::fmt::Display for ValueError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for ValueError {}

/// Convert a JSON value into a cell value.
///
/// The mapping is deliberately narrow: JSON's `string` becomes text, and the temporal
/// forms are recognised by shape (`YYYY-MM-DD` for a date, an ISO timestamp for a
/// datetime) so an agent can write a date without learning a special wrapper object.
pub fn from_json(value: &Value) -> Result<CellValue, ValueError> {
    Ok(match value {
        Value::Null => CellValue::None,
        Value::Bool(flag) => CellValue::Bool(*flag),
        Value::Number(number) => {
            let parsed = number.as_f64().ok_or_else(|| {
                ValueError(format!("{number} cannot be represented as a spreadsheet number"))
            })?;
            CellValue::Number(parsed)
        }
        Value::String(text) => from_text(text)?,
        other => {
            return Err(ValueError(format!(
                "a cell value must be a string, number, boolean or null, not {other}"
            )))
        }
    })
}

/// Interpret a string as a formula, a date, a time or plain text.
///
/// A leading `=` is a formula. A leading `%` is a fraction, matching the `"50%"` behaviour
/// of `Workbook.guess_types`. Anything that is not a recognised date or time is text, so
/// no input is ever rejected outright.
pub fn from_text(text: &str) -> Result<CellValue, ValueError> {
    let trimmed = text.trim();
    if let Some(formula) = trimmed.strip_prefix('=') {
        if formula.is_empty() {
            return Err(ValueError("a formula needs something after '='".to_string()));
        }
        return Ok(CellValue::Formula(format!("={formula}")));
    }
    if let Some(percent) = trimmed.strip_suffix('%') {
        if let Ok(fraction) = percent.trim().parse::<f64>() {
            return Ok(CellValue::Number(fraction / 100.0));
        }
    }
    if let Ok(date) = NaiveDate::parse_from_str(trimmed, "%Y-%m-%d") {
        return Ok(CellValue::Date(date));
    }
    if let Ok(datetime) = NaiveDateTime::parse_from_str(trimmed, "%Y-%m-%dT%H:%M:%S") {
        return Ok(CellValue::DateTime(datetime));
    }
    if let Ok(time) = NaiveTime::parse_from_str(trimmed, "%H:%M:%S") {
        return Ok(CellValue::Time(time));
    }
    Ok(CellValue::Text(text.to_string()))
}

/// Render a cell value as JSON.
///
/// Numbers are emitted as JSON numbers, and every other variant carries a `type` so the
/// agent can tell a formula from the text it evaluates to.
pub fn to_json(value: &CellValue) -> Value {
    match value {
        CellValue::None => Value::Null,
        CellValue::Bool(flag) => json!({ "type": "boolean", "value": flag }),
        CellValue::Number(number) => {
            if number.is_finite() {
                json!({ "type": "number", "value": number })
            } else {
                // JSON has no NaN or infinity, so a non-finite serial is reported as text.
                json!({ "type": "number", "value": number.to_string() })
            }
        }
        CellValue::Text(text) => json!({ "type": "string", "value": text }),
        CellValue::Formula(formula) => json!({ "type": "formula", "value": formula }),
        CellValue::Error(error) => json!({ "type": "error", "value": error }),
        CellValue::Date(date) => json!({ "type": "date", "value": date.to_string() }),
        CellValue::DateTime(datetime) => {
            json!({ "type": "datetime", "value": datetime.to_string() })
        }
        CellValue::Time(time) => json!({ "type": "time", "value": time.to_string() }),
        CellValue::Duration(delta) => {
            json!({ "type": "duration", "value": delta.to_string() })
        }
    }
}

/// Render a cell value as a single line of plain text, for a CSV-like view.
pub fn to_display(value: &CellValue) -> String {
    match value {
        CellValue::None => String::new(),
        CellValue::Bool(flag) => flag.to_string(),
        CellValue::Number(number) => {
            if number.is_finite() {
                // `{}` on an `f64` already avoids a trailing `.0` for whole numbers.
                format!("{number}")
            } else {
                number.to_string()
            }
        }
        CellValue::Text(text) => text.clone(),
        CellValue::Formula(formula) => formula.clone(),
        CellValue::Error(error) => error.clone(),
        CellValue::Date(date) => date.to_string(),
        CellValue::DateTime(datetime) => datetime.to_string(),
        CellValue::Time(time) => time.to_string(),
        CellValue::Duration(delta) => delta.to_string(),
    }
}

/// Quote a field for CSV output if it contains a separator, a quote or a newline.
pub fn csv_field(text: &str) -> String {
    if text.contains([',', '"', '\n', '\r']) {
        format!("\"{}\"", text.replace('"', "\"\""))
    } else {
        text.to_string()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn numbers_and_booleans_convert_directly() {
        assert_eq!(from_json(&json!(42.5)).unwrap(), CellValue::Number(42.5));
        assert_eq!(from_json(&json!(true)).unwrap(), CellValue::Bool(true));
        assert_eq!(from_json(&json!(null)).unwrap(), CellValue::None);
    }

    #[test]
    fn a_leading_equals_marks_a_formula() {
        assert_eq!(
            from_json(&json!("=SUM(A1:A2)")).unwrap(),
            CellValue::Formula("=SUM(A1:A2)".to_string())
        );
        // The equals sign is required, so an empty formula is an error rather than text.
        assert!(from_json(&json!("=")).is_err());
    }

    #[test]
    fn dates_datetimes_and_times_are_recognised_by_shape() {
        assert_eq!(
            from_json(&json!("2010-01-18")).unwrap(),
            CellValue::Date(NaiveDate::from_ymd_opt(2010, 1, 18).unwrap())
        );
        assert!(matches!(
            from_json(&json!("2010-01-18T14:15:20")).unwrap(),
            CellValue::DateTime(_)
        ));
        assert!(matches!(
            from_json(&json!("14:15:20")).unwrap(),
            CellValue::Time(_)
        ));
    }

    #[test]
    fn a_trailing_percent_becomes_a_fraction() {
        assert_eq!(
            from_json(&json!("50%")).unwrap(),
            CellValue::Number(0.5)
        );
        // A percent sign that is not a number stays text.
        assert!(matches!(
            from_json(&json!("50% off")).unwrap(),
            CellValue::Text(_)
        ));
    }

    #[test]
    fn anything_else_is_text() {
        assert_eq!(
            from_json(&json!("hello")).unwrap(),
            CellValue::Text("hello".to_string())
        );
    }

    #[test]
    fn structured_json_is_rejected_with_a_useful_message() {
        let error = from_json(&json!({"a": 1})).unwrap_err();
        assert!(error.0.contains("must be a string"), "{error}");
    }

    #[test]
    fn rendering_is_lossless_enough_to_round_trip() {
        for value in [
            CellValue::Number(1.0),
            CellValue::Text("x".to_string()),
            CellValue::Formula("=A1".to_string()),
            CellValue::Date(NaiveDate::from_ymd_opt(2010, 1, 18).unwrap()),
        ] {
            let rendered = to_json(&value);
            assert_eq!(from_json(rendered.get("value").unwrap()).unwrap(), value);
        }
    }

    #[test]
    fn non_finite_numbers_survive_rendering_as_text() {
        let rendered = to_json(&CellValue::Number(f64::NAN));
        assert_eq!(rendered["type"], json!("number"));
        assert!(rendered["value"].is_string());
    }

    #[test]
    fn csv_quoting_follows_rfc_4180() {
        assert_eq!(csv_field("plain"), "plain");
        assert_eq!(csv_field("a,b"), "\"a,b\"");
        assert_eq!(csv_field("say \"hi\""), "\"say \"\"hi\"\"\"");
        assert_eq!(csv_field("line\nbreak"), "\"line\nbreak\"");
    }
}
