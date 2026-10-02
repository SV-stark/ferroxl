//! Chart data references (`openpyxl/charts/reference.py`).

// `from_str` is openpyxl's `classmethod from_str`, so the name is kept even though Rust
// would rather these implemented `FromStr`.
#![allow(clippy::should_implement_trait)]

use crate::cell::utils::get_column_letter;
use crate::exceptions::{Error, Result};

/// Whether a reference holds numbers or strings, which decides the XML element names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ReferenceDataType {
    /// Numeric data.
    Numeric,
    /// String data.
    String,
}

impl ReferenceDataType {
    /// The XML token.
    pub fn as_str(self) -> &'static str {
        match self {
            ReferenceDataType::Numeric => "n",
            ReferenceDataType::String => "s",
        }
    }

    /// Parse an XML token.
    pub fn from_str(value: &str) -> Result<Self> {
        match value {
            "n" => Ok(ReferenceDataType::Numeric),
            "s" => Ok(ReferenceDataType::String),
            other => Err(Error::Value(format!(
                "References must be either numeric or strings, got {other}"
            ))),
        }
    }

    /// The `*Ref` element name for this data type.
    pub fn ref_element(self) -> &'static str {
        match self {
            ReferenceDataType::Numeric => "numRef",
            ReferenceDataType::String => "strRef",
        }
    }

    /// The `*Cache` element name for this data type.
    pub fn cache_element(self) -> &'static str {
        match self {
            ReferenceDataType::Numeric => "numCache",
            ReferenceDataType::String => "strCache",
        }
    }
}

/// A reference to a block of cells on a sheet, used as chart data.
///
/// The reference holds the sheet title and the 0-based `(row, column)` corners; values are
/// resolved by the caller (the worksheet) at write time, so a reference can be created
/// before the data exists.
#[derive(Debug, Clone, PartialEq)]
pub struct Reference {
    /// The sheet title.
    pub sheet_title: String,
    /// The top-left corner as 0-based `(row, column)`.
    pub pos1: (usize, usize),
    /// The bottom-right corner as 0-based `(row, column)`, exclusive on columns.
    pub pos2: Option<(usize, usize)>,
    /// Whether the referenced data is numeric or textual.
    pub data_type: Option<ReferenceDataType>,
    /// The number format applied to the cached values.
    pub number_format: Option<String>,
    /// The resolved values, if already read.
    values: Option<Vec<CellValue>>,
}

/// A resolved cell value as stored in a chart cache.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty.
    None,
    /// Text.
    Text(String),
    /// Number.
    Number(f64),
}

impl CellValue {
    /// The text written into `<c:v>`.
    pub fn to_cache_string(&self) -> String {
        match self {
            CellValue::None => String::new(),
            CellValue::Text(t) => t.clone(),
            CellValue::Number(v) => {
                if *v == v.trunc() && v.abs() < 1e16 {
                    format!("{}", *v as i64)
                } else {
                    format!("{v}")
                }
            }
        }
    }
}

impl Reference {
    /// Build a reference to a single cell or a range.
    ///
    /// Fails when `number_format` is not one of Excel's built-in format codes, which is
    /// what openpyxl's `number_format` setter checks.
    pub fn new(
        sheet_title: &str,
        pos1: (usize, usize),
        pos2: Option<(usize, usize)>,
        data_type: Option<ReferenceDataType>,
        number_format: Option<&str>,
    ) -> Result<Self> {
        // A chart cache is written with a `formatCode`, and Excel only understands the
        // built-in codes there, so openpyxl rejects anything else and so does this.
        if let Some(format) = number_format {
            if !crate::styles::numbers::is_builtin(format) {
                return Err(Error::Value(format!("Invalid number format: {format}")));
            }
        }
        Ok(Reference {
            sheet_title: sheet_title.to_string(),
            pos1,
            pos2,
            data_type,
            number_format: number_format.map(str::to_string),
            values: None,
        })
    }

    /// Set the resolved values, bypassing the worksheet lookup.
    pub fn set_values(&mut self, values: Vec<CellValue>, data_type: ReferenceDataType) {
        self.values = Some(values);
        self.data_type = Some(data_type);
    }

    /// The resolved values, if they have been set.
    pub fn values(&self) -> Option<&[CellValue]> {
        self.values.as_deref()
    }

    /// The data type, defaulting to numeric as the writer does.
    pub fn effective_data_type(&self) -> ReferenceDataType {
        self.data_type.unwrap_or(ReferenceDataType::Numeric)
    }

    /// The number format code, defaulting to `General`.
    pub fn effective_number_format(&self) -> &str {
        self.number_format.as_deref().unwrap_or("General")
    }

    /// Format as Excel reference notation (`'Sheet'!$A$1:$B$2`).
    ///
    /// `#REF!` means a region of a chart that should not be drawn, not an error.
    pub fn to_reference_string(&self) -> String {
        let (r1, c1) = self.pos1;
        match self.pos2 {
            Some((r2, c2)) => {
                let letters1 = get_column_letter(c1 as u32 + 1).unwrap_or_default();
                let letters2 = get_column_letter(c2 as u32 + 1).unwrap_or_default();
                format!(
                    "'{}'!${}${}:${}${}",
                    self.sheet_title,
                    letters1,
                    r1 + 1,
                    letters2,
                    r2 + 1
                )
            }
            None => {
                let letters = get_column_letter(c1 as u32 + 1).unwrap_or_default();
                format!("'{}'!${}${}", self.sheet_title, letters, r1 + 1)
            }
        }
    }

    /// The inclusive 1-based bounds of the reference.
    pub fn bounds(&self) -> Result<crate::worksheet::iter_worksheet::RangeBounds> {
        let (r1, c1) = self.pos1;
        match self.pos2 {
            Some((r2, c2)) => Ok(crate::worksheet::iter_worksheet::RangeBounds {
                min_col: c1 as u32 + 1,
                min_row: r1 as u32 + 1,
                max_col: c2 as u32 + 1,
                max_row: r2 as u32 + 1,
            }),
            None => Ok(crate::worksheet::iter_worksheet::RangeBounds {
                min_col: c1 as u32 + 1,
                min_row: r1 as u32 + 1,
                max_col: c1 as u32 + 1,
                max_row: r1 as u32 + 1,
            }),
        }
    }
}

impl std::fmt::Display for Reference {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.to_reference_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn single_cell_notation() {
        let reference = Reference::new("Sheet1", (0, 0), None, None, None).unwrap();
        assert_eq!(reference.to_reference_string(), "'Sheet1'!$A$1");
    }

    #[test]
    fn range_notation() {
        let reference = Reference::new("My Sheet", (1, 2), Some((4, 5)), None, None).unwrap();
        assert_eq!(reference.to_reference_string(), "'My Sheet'!$C$2:$F$5");
        assert_eq!(reference.to_string(), "'My Sheet'!$C$2:$F$5");
    }

    #[test]
    fn bounds_are_one_based() {
        let reference = Reference::new("S", (0, 0), Some((1, 1)), None, None).unwrap();
        let bounds = reference.bounds().unwrap();
        assert_eq!(bounds.min_col, 1);
        assert_eq!(bounds.min_row, 1);
        assert_eq!(bounds.max_col, 2);
        assert_eq!(bounds.max_row, 2);
    }

    #[test]
    fn data_type_drives_element_names() {
        assert_eq!(ReferenceDataType::Numeric.ref_element(), "numRef");
        assert_eq!(ReferenceDataType::Numeric.cache_element(), "numCache");
        assert_eq!(ReferenceDataType::String.ref_element(), "strRef");
        assert_eq!(ReferenceDataType::String.cache_element(), "strCache");
        assert!(ReferenceDataType::from_str("bogus").is_err());
    }

    #[test]
    fn defaults_are_numeric_and_general() {
        let reference = Reference::new("S", (0, 0), None, None, None).unwrap();
        assert_eq!(reference.effective_data_type(), ReferenceDataType::Numeric);
        assert_eq!(reference.effective_number_format(), "General");
    }

    #[test]
    fn only_a_built_in_number_format_is_accepted() {
        // A chart cache carries a `formatCode`, and Excel only reads the built-in codes
        // there, so a custom one is refused rather than written out and ignored.
        let reference = Reference::new("S", (0, 0), None, None, Some("0.00")).unwrap();
        assert_eq!(reference.effective_number_format(), "0.00");

        let error = Reference::new("S", (0, 0), None, None, Some("yyyy-mm-dd")).unwrap_err();
        assert!(
            error.to_string().contains("Invalid number format"),
            "{error}"
        );
    }

    #[test]
    fn cache_strings_match_python_formatting() {
        assert_eq!(CellValue::Number(1.0).to_cache_string(), "1");
        assert_eq!(CellValue::Number(1.5).to_cache_string(), "1.5");
        assert_eq!(CellValue::Text("x".into()).to_cache_string(), "x");
        assert_eq!(CellValue::None.to_cache_string(), "");
    }
}
