//! Chart objects (`openpyxl/charts/`).
//!
//! [`Chart`] is the common base; [`GraphChart`] adds axes and the scaling computation that
//! the writer invokes; the concrete chart types only differ in their XML type name and
//! grouping.

use super::axis::{Axis, ValueAxis};
use super::legend::Legend;
use super::series::{Series, SeriesAttr};
use crate::drawing::{Drawing, Shape};

/// A chart without axes (pie charts).
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    /// The XML subchart element name, e.g. `pieChart`.
    pub chart_type: &'static str,
    /// The series grouping, written for bar and line charts.
    pub grouping: &'static str,
    /// The data series.
    pub series: Vec<Series>,
    /// The legend.
    pub legend: Legend,
    /// Whether the legend is drawn.
    pub show_legend: bool,
    /// The language tag written to `<c:lang>`.
    pub lang: String,
    /// The chart title.
    pub title: String,
    /// Print margins written to `<c:pageMargins>`.
    pub print_margins: Vec<(String, String)>,
    /// The drawing that anchors the chart on the sheet.
    pub drawing: Drawing,
    /// Plot area width as a fraction of the drawing.
    pub width: f64,
    /// Plot area height as a fraction of the drawing.
    pub height: f64,
    /// The axes, when this chart has them.
    ///
    /// A worksheet stores charts as the common [`Chart`] type, so a graph chart's axes are
    /// carried here for the writer. Pie charts leave this as `None`.
    pub axes: Option<(Axis, Axis)>,
    /// Base top margin as a fraction of the drawing.
    pub base_margin_top: f64,
    /// Base left margin as a fraction of the drawing.
    pub base_margin_left: f64,
    /// Shapes drawn inside the chart.
    pub shapes: Vec<Shape>,
}

impl Default for Chart {
    fn default() -> Self {
        let mut drawing = Drawing::new();
        drawing.left = 10;
        drawing.top = 400;
        drawing.set_height(400);
        drawing.set_width(800);
        Chart {
            chart_type: "",
            grouping: "standard",
            series: Vec::new(),
            legend: Legend::new(),
            show_legend: true,
            lang: "en-GB".to_string(),
            title: String::new(),
            print_margins: vec![
                ("b".to_string(), "0.75".to_string()),
                ("l".to_string(), "0.7".to_string()),
                ("r".to_string(), "0.7".to_string()),
                ("t".to_string(), "0.75".to_string()),
                ("header".to_string(), "0.3".to_string()),
                ("footer".to_string(), "0.3".to_string()),
            ],
            drawing,
            width: 0.6,
            height: 0.6,
            axes: None,
            base_margin_top: 1.0,
            base_margin_left: 0.0,
            shapes: Vec::new(),
        }
    }
}

impl Chart {
    /// A chart with no axes.
    pub fn new() -> Self {
        Chart::default()
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.series.push(series);
        self
    }

    /// Add a shape.
    pub fn add_shape(&mut self, shape: Shape) -> &mut Self {
        self.shapes.push(shape);
        self
    }

    /// Set the title.
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = title.into();
        self
    }

    /// The top margin as a fraction of the drawing, clamped so the plot area still fits.
    pub fn margin_top(&self) -> f64 {
        self.base_margin_top.min(self.max_margin_top())
    }

    /// Set the base top margin.
    pub fn set_margin_top(&mut self, value: f64) {
        self.base_margin_top = value;
    }

    /// The largest top margin that leaves room for the plot area and axis labels.
    pub fn max_margin_top(&self) -> f64 {
        let mb = (Shape::FONT_HEIGHT + Shape::MARGIN_BOTTOM) as f64;
        let plot_height = self.drawing.height() as f64 * self.height;
        (self.drawing.height() as f64 - plot_height - mb) / self.drawing.height() as f64
    }

    /// The left margin as a fraction of the drawing, at least wide enough for the labels.
    pub fn margin_left(&self) -> f64 {
        self.min_margin_left().max(self.base_margin_left)
    }

    /// Set the base left margin.
    pub fn set_margin_left(&mut self, value: f64) {
        self.base_margin_left = value;
    }

    /// The narrowest left margin that fits the y-axis labels.
    pub fn min_margin_left(&self) -> f64 {
        let ml = (self.y_chars() * Shape::FONT_WIDTH) + Shape::MARGIN_LEFT;
        ml as f64 / self.drawing.width() as f64
    }

    /// Estimate the number of characters in the y-axis labels.
    pub fn y_chars(&self) -> i64 {
        let max = self
            .series
            .iter()
            .filter_map(|s| s.max(SeriesAttr::Values))
            .fold(0.0, f64::max);
        format!("{}", max.trunc() as i64).len() as i64
    }
}

/// A chart with axes.
#[derive(Debug, Clone, PartialEq)]
pub struct GraphChart {
    /// The common chart fields.
    pub base: Chart,
    /// The x axis.
    pub x_axis: Axis,
    /// The y axis.
    pub y_axis: Axis,
    /// Whether the axes are computed from the data.
    pub auto_axis: bool,
}

impl GraphChart {
    /// Build a graph chart with the given XML type name.
    pub fn new(chart_type: &'static str) -> Self {
        GraphChart {
            base: Chart {
                chart_type,
                ..Chart::default()
            },
            x_axis: Axis::category(),
            y_axis: ValueAxis::new(),
            auto_axis: true,
        }
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.base.add_series(series);
        self
    }

    /// Compute the axis bounds from the series data.
    ///
    /// The x axis is only scaled when every series supplies x values.
    pub fn compute_axes(&mut self) {
        let (min, max) = extremes(&self.base.series, SeriesAttr::Values);
        self.y_axis.set_min(min);
        self.y_axis.set_max(max);
        self.y_axis.max_min();

        if self.base.series.iter().all(|s| !s.x_reference.is_none()) {
            let (min, max) = extremes(&self.base.series, SeriesAttr::XValues);
            self.x_axis.set_min(min);
            self.x_axis.set_max(max);
            self.x_axis.max_min();
        }
    }

    /// The width of one x-axis unit in pixels.
    pub fn x_units(&self) -> usize {
        self.base
            .series
            .iter()
            .map(|s| s.values_for(SeriesAttr::Values).len())
            .max()
            .unwrap_or(0)
    }

    /// The height of one y-axis unit in pixels.
    pub fn y_units(&self) -> f64 {
        let drawing_height = crate::units::pixels_to_emu(self.base.drawing.height() as f64);
        let plot_height = drawing_height as f64 * self.base.height;
        let max = self.y_axis.max();
        if max == 0.0 {
            return plot_height;
        }
        plot_height / max
    }
}

fn extremes(series: &[Series], attr: SeriesAttr) -> (f64, f64) {
    let mut min = 0.0f64;
    let mut max = 0.0f64;
    for item in series {
        if let Some(value) = item.max(attr) {
            max = max.max(value);
        }
        if let Some(value) = item.min(attr) {
            min = min.min(value);
        }
    }
    (min, max)
}

/// A clustered bar chart.
#[derive(Debug, Clone, PartialEq)]
pub struct BarChart(pub GraphChart);

impl BarChart {
    /// A clustered bar chart.
    pub fn new() -> Self {
        let mut chart = GraphChart::new("barChart");
        chart.base.grouping = "clustered";
        BarChart(chart)
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.0.add_series(series);
        self
    }

    /// Recompute the axis bounds from the series data.
    pub fn compute_axes(&mut self) {
        self.0.compute_axes();
    }

    /// The underlying chart, carrying its axes for storage on a worksheet.
    pub fn into_chart(self) -> Chart {
        let mut base = self.0.base;
        base.axes = Some((self.0.x_axis, self.0.y_axis));
        base
    }
}

impl Default for BarChart {
    fn default() -> Self {
        BarChart::new()
    }
}

/// A line chart.
#[derive(Debug, Clone, PartialEq)]
pub struct LineChart(pub GraphChart);

impl LineChart {
    /// A line chart.
    pub fn new() -> Self {
        LineChart(GraphChart::new("lineChart"))
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.0.add_series(series);
        self
    }

    /// Recompute the axis bounds from the series data.
    pub fn compute_axes(&mut self) {
        self.0.compute_axes();
    }

    /// The underlying chart, carrying its axes for storage on a worksheet.
    pub fn into_chart(self) -> Chart {
        let mut base = self.0.base;
        base.axes = Some((self.0.x_axis, self.0.y_axis));
        base
    }
}

impl Default for LineChart {
    fn default() -> Self {
        LineChart::new()
    }
}

/// A pie chart, which has no axes.
#[derive(Debug, Clone, PartialEq)]
pub struct PieChart(pub Chart);

impl PieChart {
    /// A pie chart.
    pub fn new() -> Self {
        PieChart(Chart {
            chart_type: "pieChart",
            ..Chart::default()
        })
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.0.add_series(series);
        self
    }

    /// The underlying chart.
    ///
    /// A pie chart has no axes, so unlike the other chart types nothing else has to be
    /// carried across.
    pub fn into_chart(self) -> Chart {
        self.0
    }
}

impl Default for PieChart {
    fn default() -> Self {
        PieChart::new()
    }
}

/// A scatter chart, whose x axis is a value axis.
#[derive(Debug, Clone, PartialEq)]
pub struct ScatterChart(pub GraphChart);

impl ScatterChart {
    /// A scatter chart.
    pub fn new() -> Self {
        let mut chart = GraphChart::new("scatterChart");
        chart.x_axis.axis_type = "valAx".to_string();
        chart.x_axis.cross_between = "midCat".to_string();
        chart.y_axis.cross_between = "midCat".to_string();
        ScatterChart(chart)
    }

    /// The underlying chart, carrying its axes for storage on a worksheet.
    pub fn into_chart(self) -> Chart {
        let mut base = self.0.base;
        base.axes = Some((self.0.x_axis, self.0.y_axis));
        base
    }

    /// Add a series.
    pub fn add_series(&mut self, series: Series) -> &mut Self {
        self.0.add_series(series);
        self
    }

    /// Recompute the axis bounds from the series data.
    pub fn compute_axes(&mut self) {
        self.0.compute_axes();
    }
}

impl Default for ScatterChart {
    fn default() -> Self {
        ScatterChart::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::reference::{CellValue, Reference, ReferenceDataType};

    fn series_with(values: &[f64]) -> Series {
        let cells: Vec<CellValue> = values.iter().map(|v| CellValue::Number(*v)).collect();
        let mut reference = Reference::new("S", (0, 0), None, None, None).unwrap();
        reference.set_values(cells, ReferenceDataType::Numeric);
        Series::new(reference)
    }

    #[test]
    fn chart_defaults_match_python() {
        let chart = Chart::new();
        assert!(chart.show_legend);
        assert_eq!(chart.lang, "en-GB");
        assert_eq!(chart.width, 0.6);
        assert_eq!(chart.height, 0.6);
        assert_eq!(chart.drawing.left, 10);
        assert_eq!(chart.drawing.top, 400);
        assert_eq!(chart.drawing.width(), 800);
        assert_eq!(chart.drawing.height(), 400);
        assert_eq!(chart.legend.position, "r");
        assert_eq!(chart.print_margins.len(), 6);
    }

    #[test]
    fn concrete_chart_types() {
        assert_eq!(BarChart::new().0.base.chart_type, "barChart");
        assert_eq!(BarChart::new().0.base.grouping, "clustered");
        assert_eq!(LineChart::new().0.base.chart_type, "lineChart");
        assert_eq!(PieChart::new().0.chart_type, "pieChart");
        assert_eq!(ScatterChart::new().0.base.chart_type, "scatterChart");
        assert_eq!(ScatterChart::new().0.x_axis.axis_type, "valAx");
    }

    #[test]
    fn axes_scale_to_data() {
        let mut chart = LineChart::new();
        chart.add_series(series_with(&[1.0, 5.0, 9.0]));
        chart.compute_axes();
        assert!(chart.0.y_axis.max() >= 9.0);
        assert!(chart.0.y_axis.min() <= 1.0);
        assert!(chart.0.y_axis.unit() > 0.0);
    }

    #[test]
    fn x_axis_skipped_without_x_values() {
        let mut chart = LineChart::new();
        chart.add_series(series_with(&[1.0, 2.0]));
        let before = chart.0.x_axis.raw_max;
        chart.compute_axes();
        assert_eq!(chart.0.x_axis.raw_max, before);
        assert_eq!(chart.0.x_units(), 2);
    }

    #[test]
    fn y_chars_estimate_label_width() {
        let mut chart = LineChart::new();
        chart.add_series(series_with(&[1.0, 9999.0]));
        assert_eq!(chart.0.base.y_chars(), 4);
    }

    #[test]
    fn margins_are_clamped() {
        let mut chart = LineChart::new();
        chart.0.base.set_margin_top(0.1);
        assert!(chart.0.base.margin_top() >= 0.1);
        chart.0.base.set_margin_top(2.0);
        assert!(chart.0.base.margin_top() <= 1.0);
        assert!(chart.0.base.margin_left() >= 0.0);
    }

    #[test]
    fn pie_chart_collects_series() {
        let mut chart = PieChart::new();
        chart.add_series(series_with(&[1.0]));
        assert_eq!(chart.0.series.len(), 1);
    }

    #[test]
    fn shapes_can_be_added() {
        let mut chart = LineChart::new();
        chart.0.base.add_shape(Shape::with_text("note"));
        assert_eq!(chart.0.base.shapes.len(), 1);
        assert_eq!(chart.0.base.shapes[0].text.as_deref(), Some("note"));
    }

    #[test]
    fn value_axis_used_for_y() {
        let chart = LineChart::new();
        assert_eq!(chart.0.y_axis.axis_type, "valAx");
        assert!(chart.0.y_axis.is_value_axis());
        assert!(!chart.0.x_axis.is_value_axis());
    }
}
