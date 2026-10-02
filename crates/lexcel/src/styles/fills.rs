//! Area fill patterns (`openpyxl/styles/fills.py`).

use super::colors::Color;

/// Area fill patterns for use in styles.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Fill {
    /// One of the `FILL_*` pattern names, or `None` for "no fill".
    pub fill_type: Option<String>,
    /// Gradient rotation in degrees.
    pub rotation: i64,
    /// Foreground (start) colour.
    pub start_color: Color,
    /// Background (end) colour.
    pub end_color: Color,
}

impl Default for Fill {
    fn default() -> Self {
        Fill {
            fill_type: None,
            rotation: 0,
            start_color: Color::new(Color::WHITE),
            end_color: Color::new(Color::BLACK),
        }
    }
}

impl Fill {
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
}
