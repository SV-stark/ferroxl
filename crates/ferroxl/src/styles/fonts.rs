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
    /// The character set, e.g. 1 for a system font. `-1` when unset.
    ///
    /// Read and written rather than dropped: a file using a symbol or far-east font keeps
    /// its appearance on load and loses it on save otherwise.
    pub charset: i64,
    /// The font family index, 0 to 14. `-1` when unset.
    pub family: i64,
    /// The theme font reference, e.g. `major` or `minor`. Empty when unset.
    ///
    /// A themed font follows the document's theme rather than naming a typeface, so a theme
    /// change is supposed to restyle the whole workbook. Dropping it turns a themed font into
    /// whatever the name said, which defeats the point of theming.
    pub scheme: String,
    /// Render as outline text.
    pub outline: Option<bool>,
    /// Render with a shadow.
    pub shadow: Option<bool>,
    /// Condense the character spacing.
    pub condense: Option<bool>,
    /// Extend the character spacing.
    pub extend: Option<bool>,
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
            charset: -1,
            family: -1,
            scheme: String::new(),
            outline: None,
            shadow: None,
            condense: None,
            extend: None,
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
    /// Set the character set index.
    pub fn with_charset(mut self, charset: i64) -> Self {
        self.charset = charset;
        self
    }

    /// Set the font family index, 0 to 14.
    pub fn with_family(mut self, family: i64) -> Self {
        self.family = family;
        self
    }

    /// Set the theme font reference, `major` or `minor`.
    pub fn with_scheme(mut self, scheme: impl Into<String>) -> Self {
        self.scheme = scheme.into();
        self
    }

    /// Render as outline text.
    pub fn with_outline(mut self, outline: bool) -> Self {
        self.outline = Some(outline);
        self
    }

    /// Render with a shadow.
    pub fn with_shadow(mut self, shadow: bool) -> Self {
        self.shadow = Some(shadow);
        self
    }

    /// Condense the character spacing.
    pub fn with_condense(mut self, condense: bool) -> Self {
        self.condense = Some(condense);
        self
    }

    /// Extend the character spacing.
    pub fn with_extend(mut self, extend: bool) -> Self {
        self.extend = Some(extend);
        self
    }

    /// Set the text colour.
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
