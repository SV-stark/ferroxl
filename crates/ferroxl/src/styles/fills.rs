//! Area fill patterns (`openpyxl/styles/fills.py`).

use super::colors::Color;

/// Area fill patterns for use in styles.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Fill {
    /// One of the `FILL_*` pattern names, or `None` for "no fill".
    pub fill_type: Option<String>,
    /// Gradient rotation in degrees, for a `linear` gradient.
    pub rotation: i64,
    /// The gradient stops, in position order.
    ///
    /// Empty for a pattern fill. Populated for a gradient, and the only way to tell the two
    /// apart: both write a `<fill>` element, and `fill_type` alone says `linear` for a
    /// gradient and `solid` for a pattern.
    pub stops: Vec<GradientStop>,
    /// Foreground (start) colour.
    pub start_color: Color,
    /// Background (end) colour.
    pub end_color: Color,
}

/// One stop of a gradient fill.
///
/// `position` is a fraction from 0 to 1, and Excel requires the stops to be strictly
/// increasing with no two at the same position. A fill with no stops is written as a gradient
/// with none, which Excel accepts and renders as the first stop's colour.
#[derive(Debug, Clone)]
pub struct GradientStop {
    /// Where the stop sits, from 0 to 1.
    pub position: f64,
    /// The colour at that point.
    pub color: Color,
}

impl GradientStop {
    /// The position's bit pattern, for hashing and equality.
    ///
    /// `f64` has no `Eq` or `Hash`, and `Fill` needs both to de-duplicate the stylesheet. The
    /// bits are the right identity here rather than a numeric comparison: two stops are equal
    /// when they came from the same written value, and a stop read back from XML and written
    /// again has to hash the same both times. A `NaN` position cannot reach here — the
    /// constructor clamps to 0..=1, and `0.0` and `-0.0` are normalised on the way in.
    fn position_bits(&self) -> u64 {
        let normalised = if self.position == 0.0 {
            0.0
        } else {
            self.position
        };
        normalised.to_bits()
    }
}

impl PartialEq for GradientStop {
    fn eq(&self, other: &Self) -> bool {
        self.position_bits() == other.position_bits() && self.color == other.color
    }
}

impl Eq for GradientStop {}

impl std::hash::Hash for GradientStop {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.position_bits().hash(state);
        self.color.hash(state);
    }
}

impl GradientStop {
    /// A stop at `position` with `color`.
    pub fn new(position: f64, color: Color) -> Self {
        let clamped = position.clamp(0.0, 1.0);
        GradientStop {
            // `-0.0 == 0.0` numerically but has different bits, so the zero case is
            // normalised rather than left to surprise a reader comparing two stops.
            position: if clamped == 0.0 { 0.0 } else { clamped },
            color,
        }
    }
}

/// Build the stop list openpyxl would build from a list of colours.
///
/// Positions are spread evenly, which is what Excel does when a user drags out a two-stop
/// gradient and what openpyxl does when given colours rather than stops. Doing it here means
/// `add_stop` cannot produce a gradient Excel rejects for duplicate positions.
pub fn spread_stops(colors: &[Color]) -> Vec<GradientStop> {
    if colors.is_empty() {
        return Vec::new();
    }
    let interval = 1.0 / (colors.len() - 1).max(1) as f64;
    colors
        .iter()
        .enumerate()
        .map(|(index, color)| GradientStop::new(index as f64 * interval, color.clone()))
        .collect()
}

impl Default for Fill {
    fn default() -> Self {
        Fill {
            fill_type: None,
            rotation: 0,
            stops: Vec::new(),
            start_color: Color::new(Color::WHITE),
            end_color: Color::new(Color::BLACK),
        }
    }
}

impl Fill {
    /// A linear gradient through `colors`, positioned evenly across it.
    pub fn linear_gradient(colors: &[Color]) -> Self {
        Fill {
            fill_type: Some(Fill::FILL_GRADIENT_LINEAR.to_string()),
            stops: spread_stops(colors),
            ..Fill::default()
        }
    }

    /// A path gradient through `colors`, positioned evenly across it.
    pub fn path_gradient(colors: &[Color]) -> Self {
        Fill {
            fill_type: Some(Fill::FILL_GRADIENT_PATH.to_string()),
            stops: spread_stops(colors),
            ..Fill::default()
        }
    }

    /// Set the stops, spread evenly from `colors`.
    pub fn with_stops(mut self, colors: &[Color]) -> Self {
        self.stops = spread_stops(colors);
        self
    }

    /// Whether this is a gradient rather than a pattern.
    pub fn is_gradient(&self) -> bool {
        matches!(
            self.fill_type.as_deref(),
            Some(Fill::FILL_GRADIENT_LINEAR) | Some(Fill::FILL_GRADIENT_PATH)
        )
    }

    /// No pattern fill.
    pub const FILL_NONE: Option<&'static str> = None;
    /// Solid fill.
    pub const FILL_SOLID: &'static str = "solid";
    /// Linear gradient.
    pub const FILL_GRADIENT_LINEAR: &'static str = "linear";
    /// Path gradient.
    pub const FILL_GRADIENT_PATH: &'static str = "path";
    /// `darkDown` pattern.
    pub const FILL_PATTERN_DARKDOWN: &'static str = "darkDown";
    /// `darkGray` pattern.
    pub const FILL_PATTERN_DARKGRAY: &'static str = "darkGray";
    /// `darkGrid` pattern.
    pub const FILL_PATTERN_DARKGRID: &'static str = "darkGrid";
    /// `darkHorizontal` pattern.
    pub const FILL_PATTERN_DARKHORIZONTAL: &'static str = "darkHorizontal";
    /// `darkTrellis` pattern.
    pub const FILL_PATTERN_DARKTRELLIS: &'static str = "darkTrellis";
    /// `darkUp` pattern.
    pub const FILL_PATTERN_DARKUP: &'static str = "darkUp";
    /// `darkVertical` pattern.
    pub const FILL_PATTERN_DARKVERTICAL: &'static str = "darkVertical";
    /// `gray0625` pattern.
    pub const FILL_PATTERN_GRAY0625: &'static str = "gray0625";
    /// `gray125` pattern.
    pub const FILL_PATTERN_GRAY125: &'static str = "gray125";
    /// `lightDown` pattern.
    pub const FILL_PATTERN_LIGHTDOWN: &'static str = "lightDown";
    /// `lightGray` pattern.
    pub const FILL_PATTERN_LIGHTGRAY: &'static str = "lightGray";
    /// `lightGrid` pattern.
    pub const FILL_PATTERN_LIGHTGRID: &'static str = "lightGrid";
    /// `lightHorizontal` pattern.
    pub const FILL_PATTERN_LIGHTHORIZONTAL: &'static str = "lightHorizontal";
    /// `lightTrellis` pattern.
    pub const FILL_PATTERN_LIGHTTRELLIS: &'static str = "lightTrellis";
    /// `lightUp` pattern.
    pub const FILL_PATTERN_LIGHTUP: &'static str = "lightUp";
    /// `lightVertical` pattern.
    pub const FILL_PATTERN_LIGHTVERTICAL: &'static str = "lightVertical";
    /// `mediumGray` pattern.
    pub const FILL_PATTERN_MEDIUMGRAY: &'static str = "mediumGray";

    /// Build the default (patternless) fill.
    pub fn new() -> Self {
        Fill::default()
    }

    /// A solid fill of the given colour, the common case.
    pub fn solid(color: Color) -> Self {
        Fill {
            fill_type: Some(Fill::FILL_SOLID.to_string()),
            start_color: color,
            ..Fill::default()
        }
    }

    /// Chainable pattern setter.
    pub fn with_fill_type(mut self, pattern: Option<&str>) -> Self {
        self.fill_type = pattern.map(|p| p.to_string());
        self
    }

    /// Chainable foreground colour setter.
    pub fn with_start_color(mut self, color: Color) -> Self {
        self.start_color = color;
        self
    }

    /// Chainable background colour setter.
    pub fn with_end_color(mut self, color: Color) -> Self {
        self.end_color = color;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_match_python() {
        let fill = Fill::new();
        assert_eq!(fill.fill_type, None);
        assert_eq!(fill.rotation, 0);
        assert_eq!(fill.start_color.index, Color::WHITE);
        assert_eq!(fill.end_color.index, Color::BLACK);
    }

    #[test]
    fn solid_helper() {
        let fill = Fill::solid(Color::new("FFFF0000"));
        assert_eq!(fill.fill_type.as_deref(), Some("solid"));
        assert_eq!(fill.start_color.index, "FFFF0000");
    }

    #[test]
    fn equality_uses_all_fields() {
        let a = Fill::new();
        let mut b = Fill::new();
        assert_eq!(a, b);
        b.rotation = 90;
        assert_ne!(a, b);
    }

    #[test]
    fn stops_are_spread_evenly_across_the_gradient() {
        let stops = spread_stops(&[
            Color::new("FF000000"),
            Color::new("FFFFFFFF"),
            Color::new("FFFF0000"),
        ]);
        assert_eq!(stops.len(), 3);
        assert_eq!(stops[0].position, 0.0);
        assert_eq!(stops[1].position, 0.5);
        assert_eq!(stops[2].position, 1.0);
    }

    #[test]
    fn a_single_colour_gradient_puts_the_stop_in_the_middle() {
        // Excel writes `position="0"` for a one-stop gradient, and one at each end would be
        // two stops claiming the same colour.
        let stops = spread_stops(&[Color::new("FF112233")]);
        assert_eq!(stops.len(), 1);
        assert_eq!(stops[0].position, 0.0);
    }

    #[test]
    fn positions_are_clamped_to_the_unit_range() {
        assert_eq!(
            GradientStop::new(-1.0, Color::new("FF000000")).position,
            0.0
        );
        assert_eq!(GradientStop::new(4.0, Color::new("FF000000")).position, 1.0);
    }

    #[test]
    fn negative_zero_is_normalised_so_two_stops_compare_equal() {
        // `-0.0 == 0.0` numerically but has different bits, and the hash goes on the bits.
        let a = GradientStop::new(-0.0, Color::new("FF000000"));
        let b = GradientStop::new(0.0, Color::new("FF000000"));
        assert_eq!(a, b);
        assert_eq!(a.position.to_bits(), b.position.to_bits());
    }

    #[test]
    fn a_gradient_is_told_from_a_pattern_by_its_type_and_its_stops() {
        let gradient = Fill::linear_gradient(&[Color::new("FF000000"), Color::new("FFFFFFFF")]);
        assert!(gradient.is_gradient());
        assert_eq!(gradient.stops.len(), 2);

        let solid = Fill {
            fill_type: Some(Fill::FILL_SOLID.to_string()),
            ..Fill::default()
        };
        assert!(!solid.is_gradient());
        assert!(solid.stops.is_empty());
    }
}
