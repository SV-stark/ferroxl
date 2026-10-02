//! Chart axes (`openpyxl/charts/axis.py`).

use crate::styles::numbers::NumberFormat;

/// A category axis.
///
/// The Python classes `CategoryAxis` and `ValueAxis` differ from `Axis` only in their
/// class-level attribute defaults, so they are constructors here rather than distinct
/// types: a chart's axis is a single [`Axis`] whose defaults come from one of these.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct CategoryAxis;

impl CategoryAxis {
    /// A new axis with the category-axis defaults.
    // The name mirrors openpyxl's `CategoryAxis`, which is a class you instantiate even
    // though it builds an `Axis`.
    #[allow(clippy::new_ret_no_self)]
    pub fn new() -> Axis {
        Axis::category()
    }
}

/// A value axis.
///
/// See [`CategoryAxis`] for why this is a constructor rather than a separate type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct ValueAxis;

impl ValueAxis {
    /// A new axis with the value-axis defaults.
    #[allow(clippy::new_ret_no_self)]
    pub fn new() -> Axis {
        Axis::value()
    }
}

/// Axis scales.
///
/// [`Axis`] stores the raw `min`/`max` bounds set by the caller; the public accessors
/// recompute a padded, rounded range on read when `auto_axis` is set. That mirrors the
/// Python property behaviour, where `min`, `max` and `unit` all run `_max_min`.
#[derive(Debug, Clone, PartialEq)]
pub struct Axis {
    /// Whether the axis range is computed from the data.
    pub auto_axis: bool,
    /// The raw lower bound.
    pub raw_min: f64,
    /// The raw upper bound.
    pub raw_max: f64,
    /// The explicitly requested unit, or the computed one.
    pub raw_unit: Option<f64>,
    /// Axis title.
    pub title: String,
    /// Where the axis sits: `b` (bottom) or `l` (left).
    pub position: String,
    /// Where tick labels are drawn.
    pub tick_label_position: String,
    /// How the axis crosses the other one.
    pub crosses: String,
    /// Whether the axis is automatic.
    pub auto: bool,
    /// Label alignment.
    pub label_align: Option<String>,
    /// Label offset.
    pub label_offset: Option<i64>,
    /// Where the other axis crosses this one.
    pub cross_between: String,
    /// Axis orientation.
    pub orientation: String,
    /// The axis id written to `<axId>`.
    pub id: u32,
    /// The id of the axis this one crosses.
    pub cross: u32,
    /// Number format for the axis labels.
    pub number_format: NumberFormat,
    /// Whether to omit the axis.
    pub delete_axis: bool,
    /// Whether to draw major gridlines.
    pub major_gridlines: bool,
    /// The axis element name (`catAx` or `valAx`).
    pub axis_type: String,
}

impl Default for Axis {
    fn default() -> Self {
        Axis {
            auto_axis: true,
            raw_min: 0.0,
            raw_max: 0.0,
            raw_unit: None,
            title: String::new(),
            position: "b".to_string(),
            tick_label_position: "nextTo".to_string(),
            crosses: "autoZero".to_string(),
            auto: true,
            label_align: Some("ctr".to_string()),
            label_offset: Some(100),
            cross_between: "midCat".to_string(),
            orientation: "minMax".to_string(),
            id: 60_871_424,
            cross: 60_873_344,
            number_format: NumberFormat::default(),
            delete_axis: false,
            major_gridlines: false,
            axis_type: "catAx".to_string(),
        }
    }
}

impl Axis {
    /// Bottom of the plot area.
    pub const POSITION_BOTTOM: &'static str = "b";
    /// Left of the plot area.
    pub const POSITION_LEFT: &'static str = "l";
    /// Minimum-to-maximum orientation.
    pub const ORIENTATION_MIN_MAX: &'static str = "minMax";

    /// A category axis.
    pub fn category() -> Self {
        Axis::default()
    }

    /// Whether this axis is a value axis, which gets explicit bounds and gridlines.
    pub fn is_value_axis(&self) -> bool {
        self.axis_type == "valAx"
    }

    /// A value axis.
    pub fn value() -> Self {
        Axis {
            position: Axis::POSITION_LEFT.to_string(),
            major_gridlines: true,
            auto: false,
            cross_between: "between".to_string(),
            axis_type: "valAx".to_string(),
            id: 60_873_344,
            cross: 60_871_424,
            ..Axis::default()
        }
    }

    /// Whether `auto_axis` is set.
    pub fn new(auto_axis: bool) -> Self {
        Axis {
            auto_axis,
            ..Axis::default()
        }
    }

    /// The lower bound, recomputed when the axis is automatic.
    pub fn min(&self) -> f64 {
        if self.auto_axis {
            self.scaled_bounds().0
        } else {
            self.raw_min
        }
    }

    /// The upper bound, recomputed when the axis is automatic.
    pub fn max(&self) -> f64 {
        if self.auto_axis {
            self.scaled_bounds().1
        } else {
            self.raw_max
        }
    }

    /// The recomputed lower and upper bounds.
    pub fn scaled_bounds(&self) -> (f64, f64) {
        self.max_min().bounds
    }

    /// Set the raw lower bound.
    pub fn set_min(&mut self, value: f64) {
        self.raw_min = value;
    }

    /// Set the raw upper bound.
    pub fn set_max(&mut self, value: f64) {
        self.raw_max = value;
    }

    /// The tick interval.
    pub fn unit(&self) -> f64 {
        self.max_min().unit
    }

    /// Set the raw unit.
    pub fn set_unit(&mut self, value: f64) {
        self.raw_unit = Some(value);
    }

    /// Compute the padded, rounded bounds and tick interval for the current range.
    ///
    /// The range is scaled up by a power of ten when it is smaller than one, padded by
    /// 10%, snapped to a power-of-ten step, and finally limited to ten units.
    ///
    /// A flat series (where `max == min`) makes the Python original compute `0/0` and
    /// propagate `NaN` into every field. Rather than emit `NaN` into `<c:max>`/`<c:min>`,
    /// which Excel rejects, this returns the raw bounds with a unit of 1.
    pub fn max_min(&self) -> AxisScale {
        let length = self.raw_max - self.raw_min;
        if length == 0.0 {
            return AxisScale {
                bounds: (self.raw_min, self.raw_max),
                unit: 1.0,
            };
        }
        let sign = if length < 0.0 { -1.0 } else { 1.0 };
        let zoom = less_than_one(length).unwrap_or(1.0);
        let value = length * zoom;
        // Pad by 10%, round up, then restore the range's sign.
        let mut value = (value.abs() * 1.1).ceil();
        value *= sign;
        let log = value.abs().log10();
        let exp = log.trunc();
        let mant = log - exp;
        // The step is the smallest power of ten that divides the padded range, rounded up.
        let mut unit = (10f64.powf(mant).ceil() * 10f64.powf(exp - 1.0)).ceil();
        let value = (value / unit).ceil() * unit;
        unit /= zoom;
        if value / unit > 9.0 {
            unit *= 2.0;
        }
        let scale = value / length;
        let mini = (self.raw_min * scale).floor() / zoom;
        let maxi = (self.raw_max * scale).ceil() / zoom;
        AxisScale {
            bounds: (mini, maxi),
            unit,
        }
    }
}

/// The result of scaling an axis to its data: the padded bounds and the tick interval.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct AxisScale {
    /// The padded `(min, max)` pair.
    pub bounds: (f64, f64),
    /// The tick interval.
    pub unit: f64,
}

/// Rescale a value smaller than one by a power of ten so it exceeds one.
///
/// A zero value has no logarithm and therefore no rescaling.
pub fn less_than_one(value: f64) -> Option<f64> {
    let value = value.abs();
    if value < 1.0 && value > 0.0 {
        let exp = value.log10().abs().trunc();
        Some(10f64.powf(exp + 1.0))
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn category_axis_defaults() {
        let axis = Axis::category();
        assert_eq!(axis.position, "b");
        assert_eq!(axis.tick_label_position, "nextTo");
        assert_eq!(axis.crosses, "autoZero");
        assert!(axis.auto);
        assert_eq!(axis.label_align.as_deref(), Some("ctr"));
        assert_eq!(axis.label_offset, Some(100));
        assert_eq!(axis.cross_between, "midCat");
        assert_eq!(axis.axis_type, "catAx");
        assert_eq!(axis.id, 60_871_424);
    }

    #[test]
    fn value_axis_defaults() {
        let axis = Axis::value();
        assert_eq!(axis.position, "l");
        assert!(!axis.auto);
        assert_eq!(axis.cross_between, "between");
        assert_eq!(axis.axis_type, "valAx");
        assert_eq!(axis.id, 60_873_344);
        assert!(axis.major_gridlines);
    }

    #[test]
    fn manual_axis_returns_raw_bounds() {
        let mut axis = Axis::new(false);
        axis.set_min(5.0);
        axis.set_max(50.0);
        assert_eq!(axis.min(), 5.0);
        assert_eq!(axis.max(), 50.0);
    }

    #[test]
    fn auto_axis_pads_the_range() {
        let mut axis = Axis::new(true);
        axis.set_min(0.0);
        axis.set_max(100.0);
        let scale = axis.max_min();
        let (min, unit) = (scale.bounds.0, scale.unit);
        assert!(min <= 0.0);
        assert!(unit > 0.0);
        // At most ten units along the axis.
        let (lo, hi) = (axis.min(), axis.max());
        assert!(((hi - lo) / unit).ceil() <= 10.0);
    }

    #[test]
    fn auto_axis_pads_small_ranges() {
        let mut axis = Axis::new(true);
        axis.set_min(0.0);
        axis.set_max(0.001);
        let scale = axis.max_min();
        let unit = scale.unit;
        // The range is rescaled by 10^3, padded, then divided back down, giving ticks of
        // 0.0004 and a maximum just above the data.
        assert_eq!(scale.bounds.0, 0.0);
        assert!((unit - 0.0004).abs() < 1e-12, "unit was {unit}");
        assert!(axis.max() > 0.001, "the maximum must cover the data");
    }

    #[test]
    fn flat_series_gets_a_usable_unit() {
        let mut axis = Axis::new(true);
        axis.set_min(10.0);
        axis.set_max(10.0);
        let scale = axis.max_min();
        let (min, unit) = (scale.bounds.0, scale.unit);
        assert!(unit > 0.0);
        assert!(min.is_finite());
    }

    #[test]
    fn less_than_one_rescaling() {
        assert_eq!(less_than_one(0.5), Some(10.0));
        assert_eq!(less_than_one(0.05), Some(100.0));
        assert_eq!(less_than_one(5.0), None);
        assert_eq!(less_than_one(0.0), None);
        assert_eq!(less_than_one(-0.5), Some(10.0));
    }
}
