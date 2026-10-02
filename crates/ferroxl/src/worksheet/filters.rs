//! Auto filters and sort conditions (`openpyxl/worksheet/filters.rs`).

use std::collections::BTreeMap;

use crate::exceptions::Result;

/// Normalise a reference that may be a plain string or a range of cells.
///
/// Cell references always come back upper-cased with no `$` anchors stripped by the
/// caller; ranges are reduced to their bounding coordinates.
pub fn normalize_reference(reference: &str) -> Option<String> {
    let reference = reference.trim();
    if reference.is_empty() {
        None
    } else {
        Some(reference.to_uppercase())
    }
}

/// A filter applied to a single column.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FilterColumn {
    /// Zero-based column id; `0` is the first column.
    pub col_id: u32,
    /// The values to show.
    pub vals: Vec<String>,
    /// Whether rows with a blank cell are shown.
    pub blank: bool,
}

impl FilterColumn {
    /// Build a filter column.
    pub fn new(col_id: u32, vals: Vec<String>, blank: bool) -> Self {
        FilterColumn {
            col_id,
            vals,
            blank,
        }
    }

    /// Attributes for `<filters>`: only `blank` is emitted, and only when true.
    pub fn filter_attributes(&self) -> Vec<(String, String)> {
        if self.blank {
            vec![("blank".to_string(), "1".to_string())]
        } else {
            Vec::new()
        }
    }
}

/// A sort condition over a range of cells.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct SortCondition {
    /// The reference to sort.
    pub reference: String,
    /// Whether to sort descending.
    pub descending: bool,
}

impl SortCondition {
    /// Build a sort condition.
    pub fn new(reference: &str, descending: bool) -> Self {
        SortCondition {
            reference: normalize_reference(reference).unwrap_or_default(),
            descending,
        }
    }

    /// Attributes for `<sortCondtion>`.
    ///
    /// The misspelled element name is preserved: it is what the Python writer emits, and
    /// Excel tolerates it in files openpyxl produces.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = vec![("ref".to_string(), self.reference.clone())];
        if self.descending {
            attrs.push(("descending".to_string(), "1".to_string()));
        }
        attrs
    }
}

/// An auto filter over a sheet's data.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct AutoFilter {
    reference: Option<String>,
    filter_columns: BTreeMap<u32, FilterColumn>,
    sort_conditions: Vec<SortCondition>,
}

impl AutoFilter {
    /// An empty filter.
    pub fn new() -> Self {
        AutoFilter::default()
    }

    /// The reference this filter covers.
    pub fn reference(&self) -> Option<&str> {
        self.reference.as_deref()
    }

    /// Set the reference, normalising it to upper case.
    pub fn set_reference(&mut self, reference: &str) {
        self.reference = normalize_reference(reference);
    }

    /// The per-column filters, keyed by column id.
    pub fn filter_columns(&self) -> &BTreeMap<u32, FilterColumn> {
        &self.filter_columns
    }

    /// The sort conditions.
    pub fn sort_conditions(&self) -> &[SortCondition] {
        &self.sort_conditions
    }

    /// Add a filter for a column.
    ///
    /// `col_id` is zero-based, so `0` is the first column of the range.
    pub fn add_filter_column(
        &mut self,
        col_id: u32,
        vals: Vec<String>,
        blank: bool,
    ) -> FilterColumn {
        let column = FilterColumn::new(col_id, vals, blank);
        self.filter_columns.insert(col_id, column.clone());
        column
    }

    /// Add a sort condition.
    pub fn add_sort_condition(&mut self, reference: &str, descending: bool) -> SortCondition {
        let condition = SortCondition::new(reference, descending);
        self.sort_conditions.push(condition.clone());
        condition
    }

    /// Whether the filter needs a `<filterColumn>`/`<sortState>` body rather than a bare
    /// `ref` attribute.
    pub fn has_details(&self) -> bool {
        !self.filter_columns.is_empty() || !self.sort_conditions.is_empty()
    }
}

/// Parse the `colId`, `blank` and `descending` attributes as the reader does.
///
/// openpyxl passes these straight to `int()`/`bool()`; empty or malformed values become 0
/// or `false` rather than raising.
pub fn parse_u32_attr(value: Option<&str>) -> u32 {
    value
        .and_then(|v| v.trim().parse::<f64>().ok())
        .map(|v| v.trunc() as i64)
        .unwrap_or(0) as u32
}

/// Parse a truthy XML attribute the way Python's `bool(int(value))` would.
pub fn parse_bool_attr(value: Option<&str>) -> bool {
    parse_u32_attr(value) != 0
}

/// Resolve a reference to its bounding coordinates, for callers that need the range.
pub fn reference_bounds(reference: &str) -> Result<crate::worksheet::iter_worksheet::RangeBounds> {
    crate::worksheet::iter_worksheet::get_range_boundaries(reference, 0, 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reference_is_normalised() {
        assert_eq!(normalize_reference("a1:b3"), Some("A1:B3".to_string()));
        assert_eq!(normalize_reference(""), None);
        assert_eq!(normalize_reference("   "), None);
    }

    #[test]
    fn filters_and_sort_conditions_accumulate() {
        let mut filter = AutoFilter::new();
        filter.set_reference("a1:d10");
        filter.add_filter_column(1, vec!["x".into(), "y".into()], false);
        filter.add_sort_condition("a2:a10", true);
        assert_eq!(filter.reference(), Some("A1:D10"));
        assert_eq!(filter.filter_columns().len(), 1);
        assert_eq!(filter.filter_columns()[&1].vals.len(), 2);
        assert_eq!(filter.sort_conditions().len(), 1);
        assert!(filter.has_details());
    }

    #[test]
    fn blank_attribute_only_when_set() {
        let with_blank = FilterColumn::new(0, vec![], true);
        assert_eq!(
            with_blank.filter_attributes(),
            vec![("blank".to_string(), "1".to_string())]
        );
        let without = FilterColumn::new(0, vec![], false);
        assert!(without.filter_attributes().is_empty());
    }

    #[test]
    fn sort_condition_attributes() {
        let asc = SortCondition::new("a1:a5", false);
        assert_eq!(
            asc.attributes(),
            vec![("ref".to_string(), "A1:A5".to_string())]
        );
        let desc = SortCondition::new("a1:a5", true);
        assert_eq!(desc.attributes().len(), 2);
        assert_eq!(
            desc.attributes()[1],
            ("descending".to_string(), "1".to_string())
        );
    }

    #[test]
    fn attribute_parsing_matches_python_coercions() {
        assert_eq!(parse_u32_attr(Some("3")), 3);
        assert_eq!(parse_u32_attr(None), 0);
        assert_eq!(parse_u32_attr(Some("")), 0);
        assert_eq!(parse_u32_attr(Some("2.7")), 2);
        assert!(parse_bool_attr(Some("1")));
        assert!(!parse_bool_attr(None));
        assert!(!parse_bool_attr(Some("0")));
    }

    #[test]
    fn empty_filter_needs_no_body() {
        let mut filter = AutoFilter::new();
        assert!(!filter.has_details());
        filter.set_reference("A1:B2");
        assert!(!filter.has_details());
    }
}
