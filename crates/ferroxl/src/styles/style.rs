//! The aggregate style object (`openpyxl/styles/__init__.py`).

use super::alignment::Alignment;
use super::borders::Borders;
use super::fills::Fill;
use super::fonts::Font;
use super::numbers::NumberFormat;
use super::protection::Protection;

/// Style object containing all formatting details.
///
/// `static` mirrors the Python flag of the same name: styles loaded from a file are
/// marked `static`, and the worksheet copies them on first access so callers cannot
/// mutate the shared table entry.
#[derive(Debug, Clone, Default)]
pub struct Style {
    /// Font settings.
    pub font: Font,
    /// Fill settings.
    pub fill: Fill,
    /// Border settings.
    pub borders: Borders,
    /// Alignment settings.
    pub alignment: Alignment,
    /// Number format.
    pub number_format: NumberFormat,
    /// Protection settings.
    pub protection: Protection,
    /// Whether this style is shared with the loaded style table.
    ///
    /// This flag is bookkeeping, not part of the visual style: two styles differing only in
    /// `is_static` render identically, so equality and hashing ignore it.
    pub is_static: bool,
}

impl PartialEq for Style {
    fn eq(&self, other: &Self) -> bool {
        // `is_static` is bookkeeping rather than appearance, so it is excluded: a loaded
        // style and a private copy of it are the same style.
        self.font == other.font
            && self.fill == other.fill
            && self.borders == other.borders
            && self.alignment == other.alignment
            && self.number_format == other.number_format
            && self.protection == other.protection
    }
}

impl Eq for Style {}

impl std::hash::Hash for Style {
    fn hash<H: std::hash::Hasher>(&self, state: &mut H) {
        self.font.hash(state);
        self.fill.hash(state);
        self.borders.hash(state);
        self.alignment.hash(state);
        self.number_format.hash(state);
        self.protection.hash(state);
    }
}

impl Style {
    /// Build a fresh, mutable style.
    pub fn new() -> Self {
        Style::default()
    }

    /// Build a style that is marked `static` (as loaded styles are).
    pub fn static_style() -> Self {
        Style {
            is_static: true,
            ..Style::default()
        }
    }

    /// Deep copy, resetting the `static` flag exactly as `Style.copy()` does.
    pub fn copy_style(&self) -> Style {
        Style {
            font: self.font.clone(),
            fill: self.fill.clone(),
            borders: self.borders.clone(),
            alignment: self.alignment.clone(),
            number_format: self.number_format.clone(),
            protection: self.protection,
            is_static: false,
        }
    }

    /// The number format code applied to this style.
    pub fn number_format_code(&self) -> &str {
        self.number_format.format_code()
    }

    /// Replace the number format code.
    pub fn set_number_format_code(&mut self, code: &str) {
        self.number_format.set_format_code(code);
    }
}

/// The default style used for comparison against unstyled cells.
pub fn defaults() -> Style {
    Style::default()
}

impl Style {
    /// A canonical textual rendering of the style's visual fields.
    ///
    /// The style writer sorts by this so style indices are stable across runs and
    /// platforms. Python relies on hash-table iteration order here, which is not stable.
    pub fn sort_key(&self) -> String {
        let border = |side: &super::borders::Border| {
            format!(
                "{}:{}",
                side.border_style.clone().unwrap_or_default(),
                side.color.index
            )
        };
        format!(
            "{}|{}|{:?}|{:?}|{}|{:?}|{:?}|{}|{:?}|{}|{:?}|{:?}|{}",
            self.font.name,
            self.font.size,
            self.font.bold,
            self.font.italic,
            self.font.underline,
            self.font.color.index,
            self.fill.fill_type.clone().unwrap_or_default(),
            self.fill.start_color.index,
            self.fill.end_color.index,
            border(&self.borders.left),
            border(&self.borders.right),
            self.borders.diagonal_direction,
            self.number_format.format_code(),
        )
    }
}

/// Whether two styles differ only in the `static` flag, which is what
/// `openpyxl` means when it compares a cell's style against `DEFAULTS`.
pub fn same_visual_style(a: &Style, b: &Style) -> bool {
    let mut left = a.clone();
    let mut right = b.clone();
    left.is_static = false;
    right.is_static = false;
    left == right
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::styles::colors::Color;

    #[test]
    fn defaults_are_all_default() {
        let s = Style::new();
        assert_eq!(s.font, Font::default());
        assert_eq!(s.fill, Fill::default());
        assert_eq!(s.borders, Borders::default());
        assert_eq!(s.alignment, Alignment::default());
        assert_eq!(s.number_format, NumberFormat::default());
        assert_eq!(s.protection, Protection::default());
        assert!(!s.is_static);
    }

    #[test]
    fn copy_clears_static_flag() {
        let s = Style::static_style();
        let c = s.copy_style();
        assert!(s.is_static);
        assert!(!c.is_static);
        assert!(same_visual_style(&s, &c));
    }

    #[test]
    fn equality_covers_every_component() {
        let mut a = Style::new();
        let mut b = Style::new();
        assert_eq!(a, b);
        a.font.color = Color::new("FFFF0000");
        assert_ne!(a, b);
        b.font.color = Color::new("FFFF0000");
        assert_eq!(a, b);
        b.fill.fill_type = Some("solid".to_string());
        assert_ne!(a, b);
    }

    #[test]
    fn static_flag_is_not_part_of_identity() {
        // The writer deduplicates on visual identity, so a loaded style and its private
        // copy must compare and hash equal.
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        let loaded = Style::static_style();
        let copy = loaded.copy_style();
        assert!(loaded.is_static && !copy.is_static);
        assert_eq!(loaded, copy);
        assert!(same_visual_style(&loaded, &copy));

        let mut a = DefaultHasher::new();
        loaded.hash(&mut a);
        let mut b = DefaultHasher::new();
        copy.hash(&mut b);
        assert_eq!(a.finish(), b.finish());
    }

    #[test]
    fn sort_key_distinguishes_styles_and_ignores_static() {
        let loaded = Style::static_style();
        assert_eq!(loaded.sort_key(), Style::new().sort_key());

        let mut bold = Style::new();
        bold.font.bold = true;
        assert_ne!(bold.sort_key(), Style::new().sort_key());

        // The key must be stable across calls, which is what the writer's sort relies on.
        let a = bold.sort_key();
        let b = bold.sort_key();
        assert_eq!(a, b);
    }

    #[test]
    fn number_format_helpers() {
        let mut s = Style::new();
        assert_eq!(s.number_format_code(), "General");
        s.set_number_format_code("0.00%");
        assert_eq!(s.number_format_code(), "0.00%");
        assert!(!s.number_format.is_date_format());
        s.set_number_format_code("yyyy-mm-dd");
        assert!(s.number_format.is_date_format());
    }
}
