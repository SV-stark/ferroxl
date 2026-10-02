//! Alignment options (`openpyxl/styles/alignment.py`).

/// Alignment options for use in styles.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Alignment {
    /// One of the `HORIZONTAL_*` constants.
    pub horizontal: String,
    /// One of the `VERTICAL_*` constants.
    pub vertical: String,
    /// Text rotation in degrees; Excel accepts `-90..90` and `90..180`.
    pub text_rotation: i64,
    /// Whether text wraps within the cell.
    pub wrap_text: bool,
    /// Whether text shrinks to fit the cell.
    pub shrink_to_fit: bool,
    /// Indent level.
    pub indent: i64,
}

impl Default for Alignment {
    fn default() -> Self {
        Alignment {
            horizontal: Alignment::HORIZONTAL_GENERAL.to_string(),
            vertical: Alignment::VERTICAL_BOTTOM.to_string(),
            text_rotation: 0,
            wrap_text: false,
            shrink_to_fit: false,
            indent: 0,
        }
    }
}

impl Alignment {
    /// `general` horizontal alignment.
    pub const HORIZONTAL_GENERAL: &'static str = "general";
    /// `left` horizontal alignment.
    pub const HORIZONTAL_LEFT: &'static str = "left";
    /// `right` horizontal alignment.
    pub const HORIZONTAL_RIGHT: &'static str = "right";
    /// `center` horizontal alignment.
    pub const HORIZONTAL_CENTER: &'static str = "center";
    /// `centerContinuous` horizontal alignment.
    pub const HORIZONTAL_CENTER_CONTINUOUS: &'static str = "centerContinuous";
    /// `justify` horizontal alignment.
    pub const HORIZONTAL_JUSTIFY: &'static str = "justify";
    /// `bottom` vertical alignment.
    pub const VERTICAL_BOTTOM: &'static str = "bottom";
    /// `top` vertical alignment.
    pub const VERTICAL_TOP: &'static str = "top";
    /// `center` vertical alignment.
    pub const VERTICAL_CENTER: &'static str = "center";
    /// `justify` vertical alignment.
    pub const VERTICAL_JUSTIFY: &'static str = "justify";

    /// Build the default alignment.
    pub fn new() -> Self {
        Alignment::default()
    }

    /// Chainable horizontal setter.
    pub fn with_horizontal(mut self, value: &str) -> Self {
        self.horizontal = value.to_string();
        self
    }

    /// Chainable vertical setter.
    pub fn with_vertical(mut self, value: &str) -> Self {
        self.vertical = value.to_string();
        self
    }

    /// Chainable rotation setter.
    pub fn with_text_rotation(mut self, value: i64) -> Self {
        self.text_rotation = value;
        self
    }

    /// Chainable wrap toggle.
    pub fn with_wrap_text(mut self, value: bool) -> Self {
        self.wrap_text = value;
        self
    }

    /// Chainable shrink-to-fit toggle.
    pub fn with_shrink_to_fit(mut self, value: bool) -> Self {
        self.shrink_to_fit = value;
        self
    }

    /// Chainable indent setter.
    pub fn with_indent(mut self, value: i64) -> Self {
        self.indent = value;
        self
    }

    /// Whether anything differs from the default alignment.
    pub fn is_default(&self) -> bool {
        self == &Alignment::default()
    }

    /// Attributes for the `<alignment/>` element, following the writer's rules.
    ///
    /// Negative rotations are stored as `90 - value`, and only deviations from the
    /// defaults are emitted.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = Vec::new();
        if self.horizontal != Alignment::HORIZONTAL_GENERAL {
            attrs.push(("horizontal".to_string(), self.horizontal.clone()));
        }
        if self.vertical != Alignment::VERTICAL_BOTTOM {
            attrs.push(("vertical".to_string(), self.vertical.clone()));
        }
        if self.wrap_text {
            attrs.push(("wrapText".to_string(), "1".to_string()));
        }
        if self.shrink_to_fit {
            attrs.push(("shrinkToFit".to_string(), "1".to_string()));
        }
        if self.indent > 0 {
            attrs.push(("indent".to_string(), self.indent.to_string()));
        }
        if self.text_rotation > 0 {
            attrs.push(("textRotation".to_string(), self.text_rotation.to_string()));
        } else if self.text_rotation < 0 {
            attrs.push((
                "textRotation".to_string(),
                (90 - self.text_rotation).to_string(),
            ));
        }
        attrs
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_match_python() {
        let a = Alignment::new();
        assert_eq!(a.horizontal, "general");
        assert_eq!(a.vertical, "bottom");
        assert_eq!(a.text_rotation, 0);
        assert!(a.is_default());
        assert!(a.attributes().is_empty());
    }

    #[test]
    fn negative_rotation_is_encoded() {
        let a = Alignment::new().with_text_rotation(-45);
        let attrs = a.attributes();
        assert!(attrs.contains(&("textRotation".to_string(), "135".to_string())));
    }

    #[test]
    fn deviations_only() {
        let a = Alignment::new()
            .with_horizontal(Alignment::HORIZONTAL_CENTER)
            .with_wrap_text(true)
            .with_indent(2)
            .with_text_rotation(90);
        let attrs = a.attributes();
        assert!(attrs.contains(&("horizontal".to_string(), "center".to_string())));
        assert!(attrs.contains(&("wrapText".to_string(), "1".to_string())));
        assert!(attrs.contains(&("indent".to_string(), "2".to_string())));
        assert!(attrs.contains(&("textRotation".to_string(), "90".to_string())));
        assert!(!attrs.iter().any(|(k, _)| k == "vertical"));
        assert!(!a.is_default());
    }
}
