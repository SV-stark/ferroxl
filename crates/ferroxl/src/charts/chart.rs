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
    /// Options that only some chart types have.
    pub options: ChartOptions,
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
            options: ChartOptions::default(),
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

/// How a 3-D chart is oriented.
///
/// Only meaningful for the 3-D chart types. A 2-D chart that carries one is written with
/// it ignored rather than refused, because the field has nowhere to go and silently
/// dropping it would be harder to notice than ignoring it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View3D {
    /// Rotation about the x axis, in degrees. openpyxl allows -90 to 90.
    pub rot_x: i64,
    /// Rotation about the y axis, in degrees. openpyxl allows -90 to 90.
    pub rot_y: i64,
    /// Depth of the plot as a percentage of its width.
    pub depth_percent: i64,
    /// Whether the axes are drawn at right angles.
    pub right_angle_axes: bool,
}

impl View3D {
    /// A view with Excel's own defaults: no rotation, full depth, square axes.
    pub fn new() -> Self {
        View3D {
            rot_x: 15,
            rot_y: 20,
            depth_percent: 100,
            right_angle_axes: false,
        }
    }

    /// Set the x-axis rotation.
    pub fn with_rot_x(mut self, degrees: i64) -> Self {
        self.rot_x = degrees.clamp(-90, 90);
        self
    }

    /// Set the y-axis rotation.
    pub fn with_rot_y(mut self, degrees: i64) -> Self {
        self.rot_y = degrees.clamp(-90, 90);
        self
    }

    /// Set the plot depth as a percentage.
    pub fn with_depth_percent(mut self, percent: i64) -> Self {
        self.depth_percent = percent.clamp(1, 200);
        self
    }

    /// Draw the axes at right angles.
    pub fn with_right_angle_axes(mut self) -> Self {
        self.right_angle_axes = true;
        self
    }
}

impl Default for View3D {
    fn default() -> Self {
        View3D::new()
    }
}

/// Options that only some chart types have.
///
/// One field per option rather than a `HashMap`, because each is meaningful to a fixed set
/// of chart types and a missing one is a bug the writer can see. A map would make
/// `holeSize` on a bar chart a runtime surprise instead of a compile-time one.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartOptions {
    /// `radarStyle`: `standard`, `marker` or `filled`.
    pub radar_style: &'static str,
    /// `holeSize` as a percentage, 1 to 90. `None` leaves it to Excel's default.
    pub hole_size: Option<u16>,
    /// `firstSliceAng`, 0 to 360.
    pub first_slice_angle: Option<u16>,
    /// `bubble3D`.
    pub bubble_3d: bool,
    /// `bubbleScale`, 0 to 300.
    pub bubble_scale: Option<u16>,
    /// `showNegBubbles`.
    pub show_negative_bubbles: bool,
    /// `sizeRepresents`: `area` or `w`.
    pub size_represents: &'static str,
    /// `wireframe`, for surface charts.
    pub wireframe: bool,
    /// `ofPieType`: `pie` or `bar`, for a projected pie.
    pub of_pie_type: &'static str,
    /// The 3-D view, if this is a 3-D chart.
    pub view_3d: Option<View3D>,
}

impl Default for ChartOptions {
    fn default() -> Self {
        ChartOptions {
            radar_style: "standard",
            hole_size: None,
            first_slice_angle: None,
            bubble_3d: false,
            bubble_scale: None,
            show_negative_bubbles: false,
            size_represents: "area",
            wireframe: false,
            of_pie_type: "pie",
            view_3d: None,
        }
    }
}

impl ChartOptions {
    /// Set the radar style.
    pub fn with_radar_style(mut self, style: &'static str) -> Self {
        self.radar_style = style;
        self
    }

    /// Set the doughnut hole size as a percentage.
    pub fn with_hole_size(mut self, percent: u16) -> Self {
        self.hole_size = Some(percent.clamp(1, 90));
        self
    }

    /// Rotate the first slice, for pie and projected-pie charts.
    pub fn with_first_slice_angle(mut self, degrees: u16) -> Self {
        self.first_slice_angle = Some(degrees.min(360));
        self
    }

    /// Render bubbles as 3-D spheres.
    pub fn three_d_bubbles(mut self) -> Self {
        self.bubble_3d = true;
        self
    }

    /// Set the bubble scale, 0 to 300.
    pub fn with_bubble_scale(mut self, percent: u16) -> Self {
        self.bubble_scale = Some(percent.min(300));
        self
    }

    /// Show bubbles for negative values.
    pub fn with_negative_bubbles(mut self) -> Self {
        self.show_negative_bubbles = true;
        self
    }

    /// Whether bubble size represents `area` or `w` (width).
    pub fn with_size_represents(mut self, kind: &'static str) -> Self {
        self.size_represents = kind;
        self
    }

    /// Draw surface charts as a wireframe.
    pub fn wireframe(mut self) -> Self {
        self.wireframe = true;
        self
    }

    /// Whether a projected pie is drawn as `pie` or `bar`.
    pub fn with_of_pie_type(mut self, kind: &'static str) -> Self {
        self.of_pie_type = kind;
        self
    }

    /// Set the 3-D view.
    pub fn with_view_3d(mut self, view: View3D) -> Self {
        self.view_3d = Some(view);
        self
    }
}

/// Declares a chart type that is a graph chart: it has axes, and it can carry a 3-D view.
///
/// A macro rather than nine near-identical `impl` blocks. The bodies are genuinely
/// identical — build the inner `GraphChart`, add a series, recompute axes, carry the axes
/// across — and the part that differs is the type name and the tag, which is what the macro
/// is parameterised on. Spelling them out would be nine chances to mistype a tag.
macro_rules! graph_chart_type {
    ($(#[$meta:meta])* $name:ident, $tag:literal, $grouping:expr) => {
        $(#[$meta])*
        #[derive(Debug, Clone, PartialEq)]
        pub struct $name(pub GraphChart);

        impl $name {
            /// A new chart of this type.
            pub fn new() -> Self {
                let mut chart = GraphChart::new($tag);
                chart.base.grouping = $grouping;
                $name(chart)
            }

            /// A new chart of this type with the given options.
            pub fn with_options(options: ChartOptions) -> Self {
                let mut chart = GraphChart::new($tag);
                chart.base.grouping = $grouping;
                chart.base.options = options;
                $name(chart)
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

            /// The options this chart was built with.
            pub fn options(&self) -> &ChartOptions {
                &self.0.base.options
            }

            /// Replace the options.
            pub fn set_options(&mut self, options: ChartOptions) -> &mut Self {
                self.0.base.options = options;
                self
            }

            /// The underlying chart, carrying its axes for storage on a worksheet.
            pub fn into_chart(self) -> Chart {
                let mut base = self.0.base;
                base.axes = Some((self.0.x_axis, self.0.y_axis));
                base
            }
        }

        impl Default for $name {
            fn default() -> Self {
                $name::new()
            }
        }
    };
}

/// Declares a chart type that has no axes.
macro_rules! axeless_chart_type {
    ($(#[$meta:meta])* $name:ident, $tag:literal) => {
        $(#[$meta])*
        #[derive(Debug, Clone, PartialEq)]
        pub struct $name(pub Chart);

        impl $name {
            /// A new chart of this type.
            pub fn new() -> Self {
                $name(Chart {
                    chart_type: $tag,
                    ..Chart::default()
                })
            }

            /// A new chart of this type with the given options.
            pub fn with_options(options: ChartOptions) -> Self {
                let mut chart = Chart {
                    chart_type: $tag,
                    ..Chart::default()
                };
                chart.options = options;
                $name(chart)
            }

            /// Add a series.
            pub fn add_series(&mut self, series: Series) -> &mut Self {
                self.0.add_series(series);
                self
            }

            /// The options this chart was built with.
            pub fn options(&self) -> &ChartOptions {
                &self.0.options
            }

            /// Replace the options.
            pub fn set_options(&mut self, options: ChartOptions) -> &mut Self {
                self.0.options = options;
                self
            }

            /// The underlying chart.
            pub fn into_chart(self) -> Chart {
                self.0
            }
        }

        impl Default for $name {
            fn default() -> Self {
                $name::new()
            }
        }
    };
}

graph_chart_type!(
    /// A clustered area chart.
    AreaChart,
    "areaChart",
    "standard"
);
graph_chart_type!(
    /// A 3-D area chart.
    AreaChart3D,
    "area3DChart",
    "standard"
);
graph_chart_type!(
    /// A 3-D bar chart.
    BarChart3D,
    "bar3DChart",
    "clustered"
);
graph_chart_type!(
    /// A 3-D line chart.
    LineChart3D,
    "line3DChart",
    "standard"
);
graph_chart_type!(
    /// A radar chart. `radar_style` chooses between `standard`, `marker` and `filled`.
    RadarChart,
    "radarChart",
    "standard"
);
graph_chart_type!(
    /// A bubble chart. Each series needs x values, y values and a bubble size.
    BubbleChart,
    "bubbleChart",
    "standard"
);
graph_chart_type!(
    /// A stock chart: three or four series of open/high/low/close.
    ///
    /// A stock chart has no marker by default, which is why openpyxl sets the series marker
    /// to `none` unless told otherwise; a marker on a price series obscures the line.
    StockChart,
    "stockChart",
    "standard"
);
graph_chart_type!(
    /// A 3-D surface chart.
    SurfaceChart3D,
    "surface3DChart",
    "standard"
);
graph_chart_type!(
    /// A wireframe surface chart.
    SurfaceChart,
    "surfaceChart",
    "standard"
);

axeless_chart_type!(
    /// A 3-D pie chart.
    PieChart3D,
    "pie3DChart"
);
axeless_chart_type!(
    /// A doughnut chart: a pie with a hole in the middle.
    ///
    /// `hole_size` is the hole's diameter as a percentage of the whole. Excel's default is
    /// 10, and openpyxl's is 10 as well.
    DoughnutChart,
    "doughnutChart"
);
axeless_chart_type!(
    /// A projected pie: a pie plus a bar chart of the same data. `of_pie_type` chooses
    /// which of the two is the bar.
    ProjectedPieChart,
    "ofPieChart"
);

impl BarChart {
    /// A clustered bar chart with the given options.
    ///
    /// Present so every chart type takes options the same way. The 3-D and area variants
    /// have it from the macro that declares them; these predate it.
    pub fn with_options(options: ChartOptions) -> Self {
        let mut chart = GraphChart::new("barChart");
        chart.base.grouping = "clustered";
        chart.base.options = options;
        BarChart(chart)
    }
}

impl LineChart {
    /// A line chart with the given options.
    pub fn with_options(options: ChartOptions) -> Self {
        let mut chart = GraphChart::new("lineChart");
        chart.base.options = options;
        LineChart(chart)
    }
}

impl ScatterChart {
    /// A scatter chart with the given options.
    pub fn with_options(options: ChartOptions) -> Self {
        let mut chart = GraphChart::new("scatterChart");
        chart.x_axis.axis_type = "valAx".to_string();
        chart.x_axis.cross_between = "midCat".to_string();
        chart.y_axis.cross_between = "midCat".to_string();
        chart.base.options = options;
        ScatterChart(chart)
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
