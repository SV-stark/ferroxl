//! Charts (`openpyxl/charts/`).
//!
//! The chart object model here is a port of the Python classes: a `Chart` owns `Series`,
//! a `Legend`, a `Drawing` (for anchoring) and, for graph charts, two `Axis` objects. Axis
//! scaling is recomputed at write time from the series data, exactly as the writer does.

pub mod axis;
pub mod chart;
pub mod error_bar;
pub mod legend;
pub mod reference;
pub mod series;

pub use axis::{less_than_one, Axis};
pub use chart::{
    AreaChart, AreaChart3D, BarChart, BarChart3D, BubbleChart, Chart, ChartOptions, DoughnutChart,
    GraphChart, LineChart, LineChart3D, PieChart, PieChart3D, ProjectedPieChart, RadarChart,
    ScatterChart, StockChart, SurfaceChart, SurfaceChart3D, View3D,
};
pub use error_bar::{ErrorBar, ErrorBarType};
pub use legend::Legend;
pub use reference::Reference;
pub use series::Series;
