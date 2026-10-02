//! Border options (`openpyxl/styles/borders.py`).

use super::colors::Color;

/// A single border side.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Border {
    /// One of the `BORDER_*` style names, or `None` when the side is absent.
    pub border_style: Option<String>,
    /// Border colour.
    pub color: Color,
}

impl Default for Border {
    fn default() -> Self {
        Border {
            border_style: Some(Border::BORDER_NONE.to_string()),
            color: Color::new(Color::BLACK),
        }
    }
}

impl Border {
    /// No border.
    pub const BORDER_NONE: &'static str = "none";
    /// Dash-dot border.
    pub const BORDER_DASHDOT: &'static str = "dashDot";
    /// Dash-dot-dot border.
    pub const BORDER_DASHDOTDOT: &'static str = "dashDotDot";
    /// Dashed border.
    pub const BORDER_DASHED: &'static str = "dashed";
    /// Dotted border.
    pub const BORDER_DOTTED: &'static str = "dotted";
    /// Double border.
    pub const BORDER_DOUBLE: &'static str = "double";
    /// Hairline border.
    pub const BORDER_HAIR: &'static str = "hair";
    /// Medium border.
    pub const BORDER_MEDIUM: &'static str = "medium";
    /// Medium dash-dot border.
    pub const BORDER_MEDIUMDASHDOT: &'static str = "mediumDashDot";
    /// Medium dash-dot-dot border.
    pub const BORDER_MEDIUMDASHDOTDOT: &'static str = "mediumDashDotDot";
    /// Medium dashed border.
    pub const BORDER_MEDIUMDASHED: &'static str = "mediumDashed";
    /// Slant dash-dot border.
    pub const BORDER_SLANTDASHDOT: &'static str = "slantDashDot";
    /// Thick border.
    pub const BORDER_THICK: &'static str = "thick";
    /// Thin border.
    pub const BORDER_THIN: &'static str = "thin";

    /// Build a border with style `none` and a black colour.
    pub fn new() -> Self {
        Border::default()
    }

    /// A border with the given style and colour.
    pub fn styled(style: &str, color: Color) -> Self {
        Border {
            border_style: Some(style.to_string()),
            color,
        }
    }

    /// A border side that is not drawn at all (no `style` attribute in XML).
    pub fn absent() -> Self {
        Border {
            border_style: None,
            color: Color::new(Color::BLACK),
        }
    }

    /// Chainable style setter.
    pub fn with_style(mut self, style: Option<&str>) -> Self {
        self.border_style = style.map(|s| s.to_string());
        self
    }

    /// Chainable colour setter.
    pub fn with_color(mut self, color: Color) -> Self {
        self.color = color;
        self
    }

    /// Whether this side will be written out with a `style` attribute.
    pub fn has_style(&self) -> bool {
        !matches!(
            self.border_style.as_deref(),
            None | Some(Border::BORDER_NONE)
        )
    }
}

/// Border positioning for use in styles.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Borders {
    /// Left side.
    pub left: Border,
    /// Right side.
    pub right: Border,
    /// Top side.
    pub top: Border,
    /// Bottom side.
    pub bottom: Border,
    /// Diagonal side.
    pub diagonal: Border,
    /// One of the `DIAGONAL_*` constants.
    pub diagonal_direction: i64,
    /// `allBorders` edge used by conditional formatting.
    pub all_borders: Border,
    /// `outline` edge.
    pub outline: Border,
    /// `inside` edge.
    pub inside: Border,
    /// `vertical` edge.
    pub vertical: Border,
    /// `horizontal` edge.
    pub horizontal: Border,
}

impl Default for Borders {
    fn default() -> Self {
        Borders {
            left: Border::default(),
            right: Border::default(),
            top: Border::default(),
            bottom: Border::default(),
            diagonal: Border::default(),
            diagonal_direction: Borders::DIAGONAL_NONE,
            all_borders: Border::default(),
            outline: Border::default(),
            inside: Border::default(),
            vertical: Border::default(),
            horizontal: Border::default(),
        }
    }
}

impl Borders {
    /// No diagonal line.
    pub const DIAGONAL_NONE: i64 = 0;
    /// Diagonal runs bottom-left to top-right.
    pub const DIAGONAL_UP: i64 = 1;
    /// Diagonal runs top-left to bottom-right.
    pub const DIAGONAL_DOWN: i64 = 2;
    /// Both diagonals.
    pub const DIAGONAL_BOTH: i64 = 3;

    /// Build an all-`none` border set.
    pub fn new() -> Self {
        Borders::default()
    }

    /// Apply `style` to all four sides at once.
    pub fn all_sides(style: &str, color: Color) -> Self {
        let side = Border::styled(style, color);
        Borders {
            left: side.clone(),
            right: side.clone(),
            top: side.clone(),
            bottom: side.clone(),
            ..Borders::default()
        }
    }

    /// Chainable left-side setter.
    pub fn with_left(mut self, border: Border) -> Self {
        self.left = border;
        self
    }

    /// Chainable right-side setter.
    pub fn with_right(mut self, border: Border) -> Self {
        self.right = border;
        self
    }

    /// Chainable top-side setter.
    pub fn with_top(mut self, border: Border) -> Self {
        self.top = border;
        self
    }

    /// Chainable bottom-side setter.
    pub fn with_bottom(mut self, border: Border) -> Self {
        self.bottom = border;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_border_is_none_black() {
        let border = Border::new();
        assert_eq!(border.border_style.as_deref(), Some("none"));
        assert_eq!(border.color.index, Color::BLACK);
        assert!(!border.has_style());
    }

    #[test]
    fn absent_border_has_no_style() {
        let border = Border::absent();
        assert_eq!(border.border_style, None);
        assert!(!border.has_style());
    }

    #[test]
    fn all_sides_helper() {
        let borders = Borders::all_sides("thin", Color::new("FFFF0000"));
        assert_eq!(borders.left.border_style.as_deref(), Some("thin"));
        assert_eq!(borders.bottom.color.index, "FFFF0000");
        assert!(borders.left.has_style());
        assert_eq!(borders.diagonal_direction, Borders::DIAGONAL_NONE);
    }

    #[test]
    fn diagonal_direction_constants() {
        assert_eq!(Borders::DIAGONAL_BOTH, 3);
        assert_ne!(Borders::DIAGONAL_UP, Borders::DIAGONAL_DOWN);
    }
}
