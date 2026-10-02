//! Chart series (`openpyxl/charts/series.py`).

use super::error_bar::ErrorBar;
use super::reference::{CellValue, Reference};
use crate::units::short_color;

/// The numeric values a series can be summarised over.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SeriesAttr {
    /// The y values.
    Values,
    /// The x values.
    XValues,
    /// The bubble sizes.
    BubbleSizes,
}

/// A series of data and possibly associated labels.
#[derive(Debug, Clone, PartialEq)]
pub struct Series {
    /// The reference holding the y values.
    pub reference: Option<Reference>,
    /// The reference holding the x values.
    pub x_reference: Option<Reference>,
    /// The reference holding the category labels.
    pub labels: Option<Reference>,
    /// The reference holding the bubble sizes, for a bubble chart.
    ///
    /// openpyxl calls this `zVal` on the series it builds from `SeriesFactory` and
    /// `bubbleSize` on the element; both names for the same third column of numbers.
    pub bubble_size: Option<Reference>,
    /// The series title.
    pub title: Option<String>,
    /// Marker style; [`Series::MARKER_NONE`] by default.
    pub marker: String,
    /// Series colour in short (`RRGGBB`) form.
    color: Option<String>,
    /// Error bars.
    pub error_bar: Option<ErrorBar>,
}

impl Default for Series {
    fn default() -> Self {
        Series {
            reference: None,
            x_reference: None,
            labels: None,
            bubble_size: None,
            title: None,
            marker: Series::MARKER_NONE.to_string(),
            color: None,
            error_bar: None,
        }
    }
}

impl Series {
    /// No markers.
    pub const MARKER_NONE: &'static str = "none";

    /// A series over the given reference.
    pub fn new(reference: Reference) -> Self {
        Series {
            reference: Some(reference),
            ..Series::default()
        }
    }

    /// Set the title.
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Set the x values.
    pub fn with_xvalues(mut self, reference: Reference) -> Self {
        self.x_reference = Some(reference);
        self
    }

    /// Set the category labels.
    pub fn with_labels(mut self, reference: Reference) -> Self {
        self.labels = Some(reference);
        self
    }

    /// Set the bubble sizes, for a bubble chart.
    pub fn with_bubble_size(mut self, reference: Reference) -> Self {
        self.bubble_size = Some(reference);
        self
    }

    /// Set the series colour; the alpha prefix is stripped.
    pub fn with_color(mut self, color: &str) -> Self {
        self.color = Some(short_color(color));
        self
    }

    /// Set error bars.
    pub fn with_error_bar(mut self, error_bar: ErrorBar) -> Self {
        self.error_bar = Some(error_bar);
        self
    }

    /// The series colour in short form.
    pub fn color(&self) -> Option<&str> {
        self.color.as_deref()
    }

    /// The values of the given attribute, if the reference has been resolved.
    pub fn values_for(&self, attr: SeriesAttr) -> &[CellValue] {
        let reference = match attr {
            SeriesAttr::Values => self.reference.as_ref(),
            SeriesAttr::XValues => self.x_reference.as_ref(),
            SeriesAttr::BubbleSizes => self.bubble_size.as_ref(),
        };
        reference.and_then(|r| r.values()).unwrap_or(&[])
    }

    /// The maximum numeric value, ignoring non-numeric and error-bar-extended values.
    pub fn max(&self, attr: SeriesAttr) -> Option<f64> {
        numeric_extremum(self.values_for(attr), true)
    }

    /// The minimum numeric value, ignoring non-numeric and error-bar-extended values.
    pub fn min(&self, attr: SeriesAttr) -> Option<f64> {
        numeric_extremum(self.values_for(attr), false)
    }

    /// Number of data points in the series.
    pub fn len(&self) -> usize {
        self.values_for(SeriesAttr::Values).len()
    }

    /// Whether the series holds no values.
    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

fn numeric_extremum(values: &[CellValue], want_max: bool) -> Option<f64> {
    let mut result: Option<f64> = None;
    for value in values {
        let CellValue::Number(number) = value else {
            continue;
        };
        result = Some(match result {
            None => *number,
            Some(current) if want_max => current.max(*number),
            Some(current) => current.min(*number),
        });
    }
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::reference::ReferenceDataType;

    fn reference_with(values: &[CellValue]) -> Reference {
        let mut reference = Reference::new("S", (0, 0), None, None, None).unwrap();
        reference.set_values(values.to_vec(), ReferenceDataType::Numeric);
        reference
    }

    #[test]
    fn defaults() {
        let series = Series::new(Reference::new("S", (0, 0), None, None, None).unwrap());
        assert_eq!(series.marker, Series::MARKER_NONE);
        assert!(series.title.is_none());
        assert!(series.color().is_none());
        assert!(series.error_bar.is_none());
    }

    #[test]
    fn builders_populate_fields() {
        let series = Series::new(Reference::new("S", (0, 0), None, None, None).unwrap())
            .with_title("Revenue")
            .with_color("FF112233")
            .with_xvalues(Reference::new("S", (0, 1), None, None, None).unwrap())
            .with_labels(Reference::new("S", (1, 0), None, None, None).unwrap());
        assert_eq!(series.title.as_deref(), Some("Revenue"));
        assert_eq!(series.color(), Some("112233"));
        assert!(series.x_reference.is_some());
        assert!(series.labels.is_some());
    }

    #[test]
    fn min_max_ignore_non_numeric() {
        let series = Series::new(reference_with(&[
            CellValue::Number(3.0),
            CellValue::Text("n/a".into()),
            CellValue::Number(10.0),
            CellValue::None,
        ]));
        assert_eq!(series.max(SeriesAttr::Values), Some(10.0));
        assert_eq!(series.min(SeriesAttr::Values), Some(3.0));
        assert_eq!(series.len(), 4);
        assert!(!series.is_empty());
    }

    #[test]
    fn empty_series_has_no_extremes() {
        let series = Series::default();
        assert!(series.is_empty());
        assert_eq!(series.max(SeriesAttr::Values), None);
        assert_eq!(series.min(SeriesAttr::Values), None);
    }

    #[test]
    fn negative_values_are_handled() {
        let series = Series::new(reference_with(&[
            CellValue::Number(-5.0),
            CellValue::Number(-1.0),
        ]));
        assert_eq!(series.max(SeriesAttr::Values), Some(-1.0));
        assert_eq!(series.min(SeriesAttr::Values), Some(-5.0));
    }
}
