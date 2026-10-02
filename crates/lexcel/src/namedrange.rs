//! Named groups of cells (`openpyxl/namedrange.py`).
//!
//! A named range maps a name to one or more `(worksheet, range)` destinations. Names may
//! also refer to a literal value rather than a range, which is what
//! `NamedRangeContainingValue` represents.

use crate::exceptions::{Error, Result};

/// A named group of cells.
///
/// `scope` is the index of the worksheet the name is local to, or `None` for workbook
/// scope. Rust uses an index rather than a worksheet reference to avoid a cyclic
/// borrow between workbook and named range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedRange {
    /// The name as written in the workbook.
    pub name: String,
    /// `(worksheet index, range)` destinations.
    pub destinations: Vec<(usize, String)>,
    /// The worksheet this name is local to, if any.
    pub scope: Option<usize>,
}

impl NamedRange {
    /// Build a named range.
    pub fn new(
        name: impl Into<String>,
        destinations: Vec<(usize, String)>,
        scope: Option<usize>,
    ) -> Self {
        NamedRange {
            name: name.into(),
            destinations,
            scope,
        }
    }

    /// Render as `Sheet!Range`, joined by commas for multiple destinations.
    ///
    /// Worksheet titles are passed in because the destinations only store indices.
    pub fn to_string_with_titles(&self, titles: &[String]) -> String {
        self.destinations
            .iter()
            .map(|(index, range)| {
                let title = titles.get(*index).cloned().unwrap_or_default();
                format!("{title}!{range}")
            })
            .collect::<Vec<_>>()
            .join(",")
    }
}

/// A name that refers to a constant value rather than a range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NamedRangeContainingValue {
    /// The name.
    pub name: String,
    /// The literal value the name refers to.
    pub value: String,
    /// The worksheet this name is local to, if any.
    pub scope: Option<usize>,
}

impl NamedRangeContainingValue {
    /// Build a value-bearing name.
    pub fn new(name: impl Into<String>, value: impl Into<String>, scope: Option<usize>) -> Self {
        NamedRangeContainingValue {
            name: name.into(),
            value: value.into(),
            scope,
        }
    }
}

/// Either kind of defined name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DefinedName {
    /// A range reference.
    Range(NamedRange),
    /// A literal value.
    Value(NamedRangeContainingValue),
}

impl DefinedName {
    /// The name.
    pub fn name(&self) -> &str {
        match self {
            DefinedName::Range(r) => &r.name,
            DefinedName::Value(v) => &v.name,
        }
    }

    /// The worksheet this name is local to.
    pub fn scope(&self) -> Option<usize> {
        match self {
            DefinedName::Range(r) => r.scope,
            DefinedName::Value(v) => v.scope,
        }
    }
}

/// Split a named-range string into its `(sheet, range)` destinations.
///
/// Quoted sheet names are unescaped and commas inside quotes do not split the string.
/// This is `split_named_range`.
pub fn split_named_range(range_string: &str) -> Result<Vec<(String, String)>> {
    let mut destinations = Vec::new();
    for part in split_outside_quotes(range_string) {
        let part = part.trim();
        if part.is_empty() {
            continue;
        }
        let (sheet, range) = parse_destination(part)?;
        destinations.push((sheet, range));
    }
    if destinations.is_empty() {
        return Err(Error::NamedRange(format!(
            "Invalid named range string: \"{range_string}\""
        )));
    }
    Ok(destinations)
}

/// Split on commas that are not inside single quotes.
fn split_outside_quotes(input: &str) -> Vec<String> {
    let mut parts = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = input.chars().peekable();
    while let Some(ch) = chars.next() {
        match ch {
            '\'' => {
                // A doubled quote is an escaped quote and never changes the quoting state,
                // so it must be consumed before the single-quote case is considered.
                if chars.peek() == Some(&'\'') {
                    current.push(chars.next().expect("peeked quote"));
                    current.push('\'');
                } else {
                    in_quotes = !in_quotes;
                    current.push(ch);
                }
            }
            ',' if !in_quotes => {
                parts.push(std::mem::take(&mut current));
            }
            _ => current.push(ch),
        }
    }
    parts.push(current);
    parts
}

/// Parse one `Sheet!Range` destination.
fn parse_destination(part: &str) -> Result<(String, String)> {
    let (sheet, range) = split_sheet_and_range(part)?;
    if !is_valid_range(&range) {
        return Err(Error::NamedRange(format!(
            "Invalid named range string: \"{part}\""
        )));
    }
    Ok((sheet, range))
}

fn split_sheet_and_range(part: &str) -> Result<(String, String)> {
    if let Some(rest) = part.strip_prefix('\'') {
        // Quoted sheet name: find the closing quote, honouring '' escapes.
        let bytes = rest.as_bytes();
        let mut i = 0usize;
        while i < bytes.len() {
            if bytes[i] == b'\'' {
                if i + 1 < bytes.len() && bytes[i + 1] == b'\'' {
                    i += 2;
                    continue;
                }
                let sheet = rest[..i].replace("''", "'");
                let after = &rest[i + 1..];
                let range = after.strip_prefix('!').ok_or_else(|| {
                    Error::NamedRange(format!("Invalid named range string: \"{part}\""))
                })?;
                return Ok((sheet, range.to_string()));
            }
            i += 1;
        }
        return Err(Error::NamedRange(format!(
            "Invalid named range string: \"{part}\""
        )));
    }
    match part.split_once('!') {
        Some((sheet, range)) => Ok((sheet.to_string(), range.to_string())),
        None => Err(Error::NamedRange(format!(
            "Invalid named range string: \"{part}\""
        ))),
    }
}

/// Whether the range half of a destination is a plausible A1-style range.
///
/// The Python regex is `^Sheet!$?([A-Za-z]+)?$?([0-9]+)?(:$?([A-Za-z]+)?$?([0-9]+)?)?`,
/// so a bare sheet name with an empty range is also accepted.
fn is_valid_range(range: &str) -> bool {
    if range.is_empty() {
        return true;
    }
    let (first, second) = match range.split_once(':') {
        Some((a, b)) => (a, Some(b)),
        None => (range, None),
    };
    if !is_cell_like(first) {
        return false;
    }
    match second {
        Some(s) => is_cell_like(s),
        None => true,
    }
}

fn is_cell_like(token: &str) -> bool {
    if token.is_empty() {
        return true;
    }
    // A reference is `[$]letters[$]digits`. The column anchor is the leading `$` and the row
    // anchor is the `$` between the letters and the digits; each is optional, and either
    // half may be empty (a whole-column or whole-row reference).
    let rest = token.strip_prefix('$').unwrap_or(token);
    match rest.split_once('$') {
        // The row anchor is present, so the halves are unambiguous.
        Some((letters, digits)) => {
            letters.chars().all(|c| c.is_ascii_alphabetic())
                && digits.chars().all(|c| c.is_ascii_digit())
        }
        // No row anchor: letters are whatever precedes the first digit.
        None => {
            let split = rest
                .char_indices()
                .find(|(_, c)| c.is_ascii_digit())
                .map(|(i, _)| i)
                .unwrap_or(rest.len());
            let (letters, digits) = rest.split_at(split);
            letters.chars().all(|c| c.is_ascii_alphabetic())
                && digits.chars().all(|c| c.is_ascii_digit())
        }
    }
}

/// Whether a string looks like a named-range reference (`refers_to_range`).
pub fn refers_to_range(range_string: &str) -> bool {
    if range_string.is_empty() {
        return false;
    }
    split_named_range(range_string).is_ok()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn splits_quoted_sheet() {
        assert_eq!(
            split_named_range("'My Sheet'!$D$8").unwrap(),
            vec![("My Sheet".to_string(), "$D$8".to_string())]
        );
    }

    #[test]
    fn splits_unquoted_sheet() {
        assert_eq!(
            split_named_range("HYPOTHESES!$B$3:$L$3").unwrap(),
            vec![("HYPOTHESES".to_string(), "$B$3:$L$3".to_string())]
        );
    }

    #[test]
    fn rejects_missing_bang() {
        assert!(matches!(
            split_named_range("HYPOTHESES$B$3"),
            Err(Error::NamedRange(_))
        ));
    }

    #[test]
    fn handles_commas_and_commas_in_quotes() {
        // A doubled `''` inside a quoted sheet name is an escaped quote, and a comma
        // between quotes does not split the destinations.
        let result = split_named_range(
            "'My Sheet with a , and '''!$U$16:$U$24,'My Sheet with a , and '''!$V$28:$V$36",
        )
        .unwrap();
        assert_eq!(result.len(), 2);
        assert_eq!(result[0].0, "My Sheet with a , and '");
        assert_eq!(result[0].1, "$U$16:$U$24");
        assert_eq!(result[1].0, "My Sheet with a , and '");
        assert_eq!(result[1].1, "$V$28:$V$36");
    }

    #[test]
    fn refers_to_range_detects_shape() {
        assert!(refers_to_range("'My Sheet'!$D$8"));
        assert!(refers_to_range("SHEET!$A$1:$B$2"));
        assert!(!refers_to_range("9.99"));
        assert!(!refers_to_range("NA()"));
        assert!(!refers_to_range(""));
    }

    #[test]
    fn bare_sheet_name_is_a_range() {
        assert_eq!(
            split_named_range("Sheet1!").unwrap(),
            vec![("Sheet1".to_string(), String::new())]
        );
    }

    #[test]
    fn rendering_uses_titles() {
        let titles = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        let range = NamedRange::new(
            "MyRef",
            vec![(0, "$A$1".to_string()), (1, "$B$2".to_string())],
            None,
        );
        assert_eq!(
            range.to_string_with_titles(&titles),
            "Sheet1!$A$1,Sheet2!$B$2"
        );
    }
}
