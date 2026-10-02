//! Font options (`openpyxl/styles/fonts.py`).

use super::colors::Color;

/// Font options used in styles.
///
/// Equality and hashing treat `size` by its bit pattern, because the point size is an
/// `f64` and `f64` has neither `Eq` nor `Hash`. Two sizes are the same style only when
/// they are the same value, which is what the deduplication in the style writer needs.
#[derive(Debug, Clone)]
pub struct Font {
    /// Font family name.
    pub name: String,
    /// Point size.
    ///
    /// The Python reader assigns the raw XML attribute here, so a loaded font can hold a
    /// string such as `"11"`. Rust models the size as a number, which round-trips to the
    /// same XML text while keeping arithmetic meaningful.
    pub size: f64,
    /// Bold toggle.
    pub bold: bool,
    /// Italic toggle.
    pub italic: bool,
    /// Superscript toggle.
    pub superscript: bool,
    /// Subscript toggle.
    pub subscript: bool,
    /// Underline style; see the `UNDERLINE_*` constants.
    pub underline: String,
    /// Strikethrough toggle.
    pub strikethrough: bool,
    /// Font colour.
    pub color: Color,
}

impl PartialEq for Font {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
            && self.size.to_bits() == other.size.to_bits()
            && self.bold == other.bold
            && self.italic == other.italic
            && self.superscript == other.superscript
            && self.subscript == other.subscript
            && self.underline == other.underline
            && self.strikethrough == other.strikethrough
            && self.color == other.color
    }
}

impl Eq for Font {}

impl std::hash::Hash for Font {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.name.hash(state);
        self.size.to_bits().hash(state);
        self.bold.hash(state);
        self.italic.hash(state);
        self.superscript.hash(state);
        self.subscript.hash(state);
        self.underline.hash(state);
        self.strikethrough.hash(state);
        self.color.hash(state);
    }
}

impl Default for Font {
    fn default() -> Self {
        Font {
            name: "Calibri".to_string(),
            size: 11.0,
            bold: false,
            italic: false,
            superscript: false,
            subscript: false,
            underline: Font::UNDERLINE_NONE.to_string(),
            strikethrough: false,
            color: Color::new(Color::BLACK),
        }
    }
}

impl Font {
    /// No underline.
    pub const UNDERLINE_NONE: &'static str = "none";
    /// Double underline.
    pub const UNDERLINE_DOUBLE: &'static str = "double";
    /// Double accounting underline.
    pub const UNDERLINE_DOUBLE_ACCOUNTING: &'static str = "doubleAccounting";
    /// Single underline.
    pub const UNDERLINE_SINGLE: &'static str = "single";
    /// Single accounting underline.
    pub const UNDERLINE_SINGLE_ACCOUNTING: &'static str = "singleAccounting";

    /// Build the default font.
    pub fn new() -> Self {
        Font::default()
    }

    /// Chainable setter for the font name.
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = name.into();
        self
    }

    /// Chainable setter for the point size.
    pub fn with_size(mut self, size: f64) -> Self {
        self.size = size;
        self
    }

    /// Chainable bold setter.
    pub fn with_bold(mut self, bold: bool) -> Self {
        self.bold = bold;
        self
    }

    /// Chainable italic setter.
    pub fn with_italic(mut self, italic: bool) -> Self {
        self.italic = italic;
        self
    }

    /// Chainable underline setter.
    pub fn with_underline(mut self, style: &str) -> Self {
        self.underline = style.to_string();
        self
    }

    /// Chainable colour setter.
    pub fn with_color(mut self, color: Color) -> Self {
        self.color = color;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn defaults_match_python() {
        let font = Font::new();
        assert_eq!(font.name, "Calibri");
        assert_eq!(font.size, 11.0);
        assert!(!font.bold);
        assert_eq!(font.underline, "none");
        assert_eq!(font.color.index, Color::BLACK);
    }

    #[test]
    fn equality_uses_all_fields() {
        let mut a = Font::new();
        let b = Font::new();
        assert_eq!(a, b);
        a.bold = true;
        assert_ne!(a, b);
    }

    #[test]
    fn builders_compose() {
        let font = Font::new()
            .with_name("Arial")
            .with_size(14.0)
            .with_bold(true)
            .with_underline(Font::UNDERLINE_DOUBLE);
        assert_eq!(font.name, "Arial");
        assert_eq!(font.size, 14.0);
        assert_eq!(font.underline, "double");
    }
}
