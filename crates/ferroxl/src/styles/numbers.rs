//! Number formats (`openpyxl/styles/numbers.py`).

use regex::Regex;
use std::sync::OnceLock;

/// Number formatting for use in styles.
///
/// The Python class overrides `__eq__` to compare `format_code` alone while `__hash__`
/// still folds in `_format_index`, which breaks the hash invariant. Rust cannot do that,
/// so both `Eq`/`Hash` here are based on the format code — the behaviour the override was
/// clearly reaching for, and what the style writer relies on to deduplicate formats.
#[derive(Debug, Clone)]
pub struct NumberFormat {
    format_code: String,
    format_index: Option<u32>,
}

impl PartialEq for NumberFormat {
    fn eq(&self, other: &Self) -> bool {
        self.format_code == other.format_code
    }
}

impl Eq for NumberFormat {}

impl std::hash::Hash for NumberFormat {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.format_code.hash(state);
    }
}

impl Default for NumberFormat {
    fn default() -> Self {
        NumberFormat {
            format_code: NumberFormat::FORMAT_GENERAL.to_string(),
            format_index: Some(0),
        }
    }
}

/// The built-in numeric format codes, keyed by their `numFmtId`.
pub const BUILTIN_FORMATS: [(u32, &str); 34] = [
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (5, "\"$\"#,##0_);(\"$\"#,##0)"),
    (6, "\"$\"#,##0_);[Red](\"$\"#,##0)"),
    (7, "\"$\"#,##0.00_);(\"$\"#,##0.00)"),
    (8, "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "mm-dd-yy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yy h:mm"),
    (37, "#,##0_);(#,##0)"),
    (38, "#,##0_);[Red](#,##0)"),
    (39, "#,##0.00_);(#,##0.00)"),
    (40, "#,##0.00_);[Red](#,##0.00)"),
    (41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)"),
    (
        42,
        "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
    ),
    (43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"),
    (
        44,
        "_(\"$\"* #,##0.00_)_(\"$\"* \\(#,##0.00\\)_(\"$\"* \"-\"??_)_(@_)",
    ),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mmss.0"),
];

impl NumberFormat {
    /// `General`.
    pub const FORMAT_GENERAL: &'static str = "General";
    /// Text (`@`).
    pub const FORMAT_TEXT: &'static str = "@";
    /// Integer (`0`).
    pub const FORMAT_NUMBER: &'static str = "0";
    /// Two decimals (`0.00`).
    pub const FORMAT_NUMBER_00: &'static str = "0.00";
    /// Comma-separated with two decimals.
    pub const FORMAT_NUMBER_COMMA_SEPARATED1: &'static str = "#,##0.00";
    /// Comma-separated accounting.
    pub const FORMAT_NUMBER_COMMA_SEPARATED2: &'static str = "#,##0.00_-";
    /// Whole-number percentage.
    pub const FORMAT_PERCENTAGE: &'static str = "0%";
    /// Percentage with two decimals.
    pub const FORMAT_PERCENTAGE_00: &'static str = "0.00%";
    /// ISO date.
    pub const FORMAT_DATE_YYYYMMDD2: &'static str = "yyyy-mm-dd";
    /// Short date.
    pub const FORMAT_DATE_YYYYMMDD: &'static str = "yy-mm-dd";
    /// `dd/mm/yy`.
    pub const FORMAT_DATE_DDMMYYYY: &'static str = "dd/mm/yy";
    /// `d/m/y`.
    pub const FORMAT_DATE_DMYSLASH: &'static str = "d/m/y";
    /// `d-m-y`.
    pub const FORMAT_DATE_DMYMINUS: &'static str = "d-m-y";
    /// `d-m`.
    pub const FORMAT_DATE_DMMINUS: &'static str = "d-m";
    /// `m-y`.
    pub const FORMAT_DATE_MYMINUS: &'static str = "m-y";
    /// `mm-dd-yy`.
    pub const FORMAT_DATE_XLSX14: &'static str = "mm-dd-yy";
    /// `d-mmm-yy`.
    pub const FORMAT_DATE_XLSX15: &'static str = "d-mmm-yy";
    /// `d-mmm`.
    pub const FORMAT_DATE_XLSX16: &'static str = "d-mmm";
    /// `mmm-yy`.
    pub const FORMAT_DATE_XLSX17: &'static str = "mmm-yy";
    /// `m/d/yy h:mm`.
    pub const FORMAT_DATE_XLSX22: &'static str = "m/d/yy h:mm";
    /// Locale date and time.
    pub const FORMAT_DATE_DATETIME: &'static str = "d/m/y h:mm";
    /// `h:mm AM/PM`.
    pub const FORMAT_DATE_TIME1: &'static str = "h:mm AM/PM";
    /// `h:mm:ss AM/PM`.
    pub const FORMAT_DATE_TIME2: &'static str = "h:mm:ss AM/PM";
    /// `h:mm`.
    pub const FORMAT_DATE_TIME3: &'static str = "h:mm";
    /// `h:mm:ss`.
    pub const FORMAT_DATE_TIME4: &'static str = "h:mm:ss";
    /// `mm:ss`.
    pub const FORMAT_DATE_TIME5: &'static str = "mm:ss";
    /// `h:mm:ss` (six-hour variant used by the time caster).
    pub const FORMAT_DATE_TIME6: &'static str = "h:mm:ss";
    /// `i:s.S`.
    pub const FORMAT_DATE_TIME7: &'static str = "i:s.S";
    /// `h:mm:ss@`.
    pub const FORMAT_DATE_TIME8: &'static str = "h:mm:ss@";
    /// Elapsed time.
    pub const FORMAT_DATE_TIMEDELTA: &'static str = "[hh]:mm:ss";
    /// `yy/mm/dd@`.
    pub const FORMAT_DATE_YYYYMMDDSLASH: &'static str = "yy/mm/dd@";
    /// Simple USD currency.
    pub const FORMAT_CURRENCY_USD_SIMPLE: &'static str = "\"$\"#,##0.00_-";
    /// USD currency.
    pub const FORMAT_CURRENCY_USD: &'static str = "$#,##0_-";
    /// EUR currency.
    pub const FORMAT_CURRENCY_EUR_SIMPLE: &'static str = "[$EUR ]#,##0.00_-";

    /// Characters whose presence hints at a date format.
    pub const DATE_INDICATORS: &'static str = "dmyhs";

    /// Build a `General` number format.
    pub fn new() -> Self {
        NumberFormat::default()
    }

    /// Build a format with an explicit code, resolving its builtin id.
    pub fn with_code(code: &str) -> Self {
        let index = builtin_format_id(code);
        NumberFormat {
            format_code: code.to_string(),
            format_index: index,
        }
    }

    /// The format code.
    pub fn format_code(&self) -> &str {
        &self.format_code
    }

    /// Replace the format code, recomputing the builtin id.
    pub fn set_format_code(&mut self, code: &str) {
        self.format_index = builtin_format_id(code);
        self.format_code = code.to_string();
    }

    /// The builtin `numFmtId`, when the code is a standard format.
    pub fn format_index(&self) -> Option<u32> {
        self.format_index
    }

    /// Override the builtin id (used when reading `cellXfs`).
    pub fn set_format_index(&mut self, index: Option<u32>) {
        self.format_index = index;
    }

    /// Return one of the standard format codes by index.
    pub fn builtin_format_code(index: u32) -> Option<&'static str> {
        BUILTIN_FORMATS
            .iter()
            .find(|(id, _)| *id == index)
            .map(|(_, code)| *code)
    }

    /// Whether this format code is a standard format code.
    pub fn is_builtin(&self) -> bool {
        is_builtin(&self.format_code)
    }

    /// The id of a standard style, when the code is builtin.
    pub fn builtin_id(&self) -> Option<u32> {
        builtin_format_id(&self.format_code)
    }

    /// Whether the number format is actually representing a date.
    pub fn is_date_format(&self) -> bool {
        is_date_format(Some(&self.format_code))
    }
}

/// Return the builtin id for a format code.
pub fn builtin_format_id(code: &str) -> Option<u32> {
    BUILTIN_FORMATS
        .iter()
        .find(|(_, builtin)| *builtin == code)
        .map(|(id, _)| *id)
}

/// Whether a format code is one of the standard formats.
pub fn is_builtin(code: &str) -> bool {
    BUILTIN_FORMATS.iter().any(|(_, builtin)| *builtin == code)
}

fn bad_date_re() -> &'static Regex {
    static RE: OnceLock<Regex> = OnceLock::new();
    RE.get_or_init(|| Regex::new(r#"(\[|").*[dmhys].*(\]|")"#).expect("valid regex"))
}

/// Whether a format code represents a date.
///
/// A code qualifies when it contains one of `dmyhs` and does **not** look like a quoted
/// literal containing those letters (e.g. `"USD"` in a currency format).
pub fn is_date_format(fmt: Option<&str>) -> bool {
    let Some(fmt) = fmt else { return false };
    if NumberFormat::DATE_INDICATORS
        .chars()
        .any(|indicator| fmt.contains(indicator))
    {
        return !bad_date_re().is_match(fmt);
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_to_general() {
        let nf = NumberFormat::new();
        assert_eq!(nf.format_code(), "General");
        assert_eq!(nf.format_index(), Some(0));
        assert!(nf.is_builtin());
    }

    #[test]
    fn builtin_lookup_round_trips() {
        for (id, code) in BUILTIN_FORMATS {
            assert_eq!(builtin_format_id(code), Some(id), "{code}");
            assert_eq!(NumberFormat::builtin_format_code(id), Some(code));
        }
        assert!(!is_builtin("yyyy-mm-dd;@"));
        assert_eq!(builtin_format_id("yyyy-mm-dd;@"), None);
    }

    #[test]
    fn setting_code_updates_index() {
        let mut nf = NumberFormat::new();
        nf.set_format_code(NumberFormat::FORMAT_PERCENTAGE_00);
        assert_eq!(nf.format_index(), Some(10));
        nf.set_format_code("0.000");
        assert_eq!(nf.format_index(), None);
        assert!(!nf.is_builtin());
    }

    #[test]
    fn date_detection() {
        assert!(is_date_format(Some("yyyy-mm-dd")));
        assert!(is_date_format(Some("h:mm:ss")));
        assert!(is_date_format(Some("d/m/yyyy")));
        assert!(!is_date_format(Some("General")));
        assert!(!is_date_format(Some("0.00")));
        assert!(!is_date_format(None));
        // Quoted literals are excluded even though they contain date letters.
        assert!(!is_date_format(Some("\"day\"")));
        // A date letter inside brackets is an elapsed-time or colour section, so it does
        // not count. This is openpyxl's rule, and it is why `[hh]:mm:ss` is not a date.
        assert!(!is_date_format(Some("[hh]:mm:ss")));
        assert!(!is_date_format(Some("[Red]")));
        let date_format = NumberFormat::with_code("mm-dd-yy");
        assert!(is_date_format(Some(date_format.format_code())));
        // Built-in id 14 is the same US short date format.
        assert_eq!(NumberFormat::builtin_format_code(14), Some("mm-dd-yy"));
    }

    #[test]
    fn number_format_equality_is_by_code() {
        let a = NumberFormat::with_code("0.00");
        let mut b = NumberFormat::new();
        b.set_format_code("0.00");
        assert_eq!(a, b);
        b.set_format_index(Some(999));
        // The Python __eq__ compares format_code only, so the Rust Hash must too.
        assert_eq!(a, b);
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};
        // Equality and hashing must agree, because the style writer dedupes on the hash.
        let hash_of = |value: &NumberFormat| {
            let mut hasher = DefaultHasher::new();
            value.hash(&mut hasher);
            hasher.finish()
        };
        assert_eq!(hash_of(&a), hash_of(&b));
    }
}
