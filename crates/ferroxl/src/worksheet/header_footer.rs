//! Headers and footers (`openpyxl/worksheet/header_footer.py`).
//!
//! Excel encodes a header or footer as an ampersand-delimited string with `L`/`C`/`R`
//! section markers and formatting codes. The two-way conversion here follows the Python
//! implementation, including its quirks (see the notes on individual methods).

/// An individual left/centre/right header or footer item.
///
/// Ampersand codes understood by Excel:
///
/// | Code | Meaning |
/// |------|---------|
/// | `&A` | worksheet name |
/// | `&B` | toggle bold |
/// | `&D` / `&[Date]` | current date |
/// | `&E` | toggle double underline |
/// | `&F` / `&[File]` | workbook name |
/// | `&I` | toggle italic |
/// | `&N` / `&[Pages]` | page count |
/// | `&S` | toggle strikethrough |
/// | `&T` / `&[Time]` | current time |
/// | `&P` / `&[Page]` | page number |
/// | `&U` | toggle underline |
/// | `&X` | superscript |
/// | `&Y` | subscript |
/// | `&Z` / `&[Path]` | workbook path |
/// | `&&` | literal ampersand |
/// | `&"font"` | select font |
/// | `&nn` | font point size |
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct HeaderFooterItem {
    /// The section marker: `L`, `C` or `R`.
    pub item_type: String,
    /// The selected font name.
    pub font_name: String,
    /// The selected font size in points.
    pub font_size: Option<i64>,
    /// The selected font colour as six hex digits.
    pub font_color: String,
    /// The item text.
    pub text: Option<String>,
}

impl HeaderFooterItem {
    /// The left section marker.
    pub const LEFT: &'static str = "L";
    /// The centre section marker.
    pub const CENTER: &'static str = "C";
    /// The right section marker.
    pub const RIGHT: &'static str = "R";

    /// Expansions applied to the text before writing.
    pub const REPLACE_LIST: [(&str, &str); 9] = [
        ("\n", "_x000D_"),
        ("&[Page]", "&P"),
        ("&[Pages]", "&N"),
        ("&[Date]", "&D"),
        ("&[Time]", "&T"),
        ("&[Path]", "&Z"),
        ("&[File]", "&F"),
        ("&[Tab]", "&A"),
        ("&[Picture]", "&G"),
    ];

    /// Build an empty item of the given section type.
    pub fn new(item_type: &str) -> Self {
        HeaderFooterItem {
            item_type: item_type.to_string(),
            font_name: "Calibri,Regular".to_string(),
            font_size: None,
            font_color: "000000".to_string(),
            text: None,
        }
    }

    /// Whether the item carries text.
    pub fn has_text(&self) -> bool {
        self.text.as_deref().is_some_and(|t| !t.is_empty())
    }

    /// Serialise this section, including its formatting codes.
    pub fn to_header_string(&self) -> String {
        let Some(text) = self.text.as_deref().filter(|t| !t.is_empty()) else {
            return String::new();
        };
        let mut parts: Vec<String> = vec![
            format!("&{}", self.item_type),
            format!("&\"{}\"", self.font_name),
        ];
        if let Some(size) = self.font_size {
            parts.push(format!("&{size}"));
        }
        parts.push(format!("&K{}", self.font_color));
        let mut text = text.to_string();
        for (old, new) in HeaderFooterItem::REPLACE_LIST {
            text = text.replace(old, new);
        }
        parts.push(text);
        parts.concat()
    }

    /// Parse a section from its `&`-delimited form.
    ///
    /// The fields are `&`-delimited and start with the section's marker, so everything after the
    /// first re-joins with a leading `&` -- `&L` then `P` then ` report` is the text `&P report`.
    ///
    /// The first field is not always a bare marker. When the marker was glued to its text --
    /// `&Lleft`, which is the shape openpyxl and Excel both write -- the marker is a prefix of
    /// that field rather than the whole of it, and its remainder is the start of the text. That
    /// field is the only one whose text arrives without a separator in front of it, because the
    /// split it came from never put one there.
    pub fn set_from_header_string(&mut self, item_array: &[String]) {
        let mut text_array: Vec<String> = Vec::new();
        for (position, raw) in item_array.iter().enumerate() {
            // The marker's own field, with any text glued to it taken off the front.
            let (item, separated) = if position == 0 {
                match raw.strip_prefix(self.item_type.as_str()) {
                    Some(rest) => (rest.to_string(), false),
                    None => continue,
                }
            } else {
                (raw.clone(), true)
            };
            if item.is_empty() {
                continue;
            }

            // Text already started, so this field is a code and keeps its `&`.
            if !text_array.is_empty() {
                text_array.push(format!("&{item}"));
            } else if let Some(stripped) = item.strip_prefix('"') {
                self.font_name = stripped.replace('"', "");
            } else if let Some(color) = item.strip_prefix('K') {
                if color.len() >= 6 {
                    self.font_color = color[..6].to_string();
                    text_array.push(color[6..].to_string());
                } else {
                    text_array.push(format!("&{item}"));
                }
            } else if let Ok(size) = item.parse::<i64>() {
                self.font_size = Some(size);
            } else if separated {
                text_array.push(format!("&{item}"));
            } else {
                // Text that arrived glued to the marker, which never had a `&` in front of it.
                text_array.push(item);
            }
        }
        self.text = Some(text_array.concat());
    }
}

/// The header and footer configuration for a sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderFooter {
    /// Left header.
    pub left_header: HeaderFooterItem,
    /// Centre header.
    pub center_header: HeaderFooterItem,
    /// Right header.
    pub right_header: HeaderFooterItem,
    /// Left footer.
    pub left_footer: HeaderFooterItem,
    /// Centre footer.
    pub center_footer: HeaderFooterItem,
    /// Right footer.
    pub right_footer: HeaderFooterItem,
}

impl Default for HeaderFooter {
    fn default() -> Self {
        HeaderFooter {
            left_header: HeaderFooterItem::new(HeaderFooterItem::LEFT),
            center_header: HeaderFooterItem::new(HeaderFooterItem::CENTER),
            right_header: HeaderFooterItem::new(HeaderFooterItem::RIGHT),
            left_footer: HeaderFooterItem::new(HeaderFooterItem::LEFT),
            center_footer: HeaderFooterItem::new(HeaderFooterItem::CENTER),
            right_footer: HeaderFooterItem::new(HeaderFooterItem::RIGHT),
        }
    }
}

impl HeaderFooter {
    /// An empty header/footer pair.
    pub fn new() -> Self {
        HeaderFooter::default()
    }

    /// Whether any header section carries text.
    pub fn has_header(&self) -> bool {
        self.left_header.has_text() || self.center_header.has_text() || self.right_header.has_text()
    }

    /// Whether any footer section carries text.
    pub fn has_footer(&self) -> bool {
        self.left_footer.has_text() || self.center_footer.has_text() || self.right_footer.has_text()
    }

    /// Serialise the header.
    pub fn header_string(&self) -> String {
        let mut parts = Vec::new();
        for item in [&self.left_header, &self.center_header, &self.right_header] {
            if item.has_text() {
                parts.push(item.to_header_string());
            }
        }
        parts.concat()
    }

    /// Serialise the footer.
    pub fn footer_string(&self) -> String {
        let mut parts = Vec::new();
        for item in [&self.left_footer, &self.center_footer, &self.right_footer] {
            if item.has_text() {
                parts.push(item.to_header_string());
            }
        }
        parts.concat()
    }

    /// Parse a header string into its three sections.
    pub fn set_header(&mut self, item: &str) {
        let items = split_header(item);
        let (left, center, right) = section_bounds(&items);
        if let Some(fields) = left {
            self.left_header.set_from_header_string(&fields);
        }
        if let Some(fields) = center {
            self.center_header.set_from_header_string(&fields);
        }
        if let Some(fields) = right {
            self.right_header.set_from_header_string(&fields);
        }
    }

    /// Parse a footer string into its three sections.
    pub fn set_footer(&mut self, item: &str) {
        let items = split_header(item);
        let (left, center, right) = section_bounds(&items);
        if let Some(fields) = left {
            self.left_footer.set_from_header_string(&fields);
        }
        if let Some(fields) = center {
            self.center_footer.set_from_header_string(&fields);
        }
        if let Some(fields) = right {
            self.right_footer.set_from_header_string(&fields);
        }
    }
}

/// Split on `&` while protecting escaped ampersands.
///
/// The Python code round-trips `&&` through a sentinel so that an escaped ampersand is not
/// mistaken for a code marker.
fn split_header(item: &str) -> Vec<String> {
    const SENTINEL: &str = "#DOUBLEAMP#";
    let guarded = item.replace("&&", SENTINEL);
    guarded
        .split('&')
        .map(|part| part.replace(SENTINEL, "&&"))
        .collect()
}

/// The `&`-delimited fields of each of the `L`, `C` and `R` sections, in that order.
///
/// Each field list starts with the section's own marker, because
/// [`HeaderFooterItem::set_from_header_string`] skips the first entry as the marker and reads the
/// rest. A section that is not present is `None`.
pub type SectionSpans = (
    Option<Vec<String>>,
    Option<Vec<String>>,
    Option<Vec<String>>,
);

/// Split a header or footer string into the three sections' fields.
///
/// openpyxl's `_split_string` matches each marker with a regex that glues the section's text to it
/// -- `(&L(?P<left>.+?))?` -- so `&Lleft` is a left section carrying `left`, not a stray field named
/// `Lleft`. Comparing whole `&`-delimited fields gets that wrong, and it is not an edge case: this
/// is the shape openpyxl and Excel both *write*, so a header set through either of them read back
/// as nothing and was dropped on the next save.
fn section_bounds(items: &[String]) -> SectionSpans {
    // A marker is a field that *starts* with it; its text is what follows.
    let find = |marker: &str| items.iter().position(|item| item.starts_with(marker));
    let left = find(HeaderFooterItem::LEFT);
    let center = find(HeaderFooterItem::CENTER);
    let right = find(HeaderFooterItem::RIGHT);

    // The three positions in the order openpyxl's regex lays them out. A section runs from its
    // marker to the start of the next one, which is what the non-greedy `.+?` does.
    let order = [
        (HeaderFooterItem::LEFT, left),
        (HeaderFooterItem::CENTER, center),
        (HeaderFooterItem::RIGHT, right),
    ];

    let mut spans: SectionSpans = (None, None, None);
    for (position, (marker, at)) in order.into_iter().enumerate() {
        let Some(at) = at else {
            continue;
        };
        let next = order[position + 1..]
            .iter()
            .filter_map(|(_, other)| *other)
            .filter(|stop| *stop > at)
            .min()
            .unwrap_or(items.len());

        // The marker's own field is kept whole, so the parser can tell a bare marker from one with
        // its text glued to it. The fields after it carry the font name, colour and size codes.
        let mut fields = vec![items[at].clone()];
        if at + 1 < next {
            fields.extend_from_slice(&items[at + 1..next]);
        }
        debug_assert!(fields[0].starts_with(marker));

        let slot = match marker {
            HeaderFooterItem::LEFT => &mut spans.0,
            HeaderFooterItem::CENTER => &mut spans.1,
            _ => &mut spans.2,
        };
        *slot = Some(fields);
    }
    spans
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_header_footer_writes_nothing() {
        let hf = HeaderFooter::new();
        assert!(!hf.has_header());
        assert!(!hf.has_footer());
        assert_eq!(hf.header_string(), "");
        assert_eq!(hf.footer_string(), "");
    }

    #[test]
    fn round_trips_plain_text() {
        let mut hf = HeaderFooter::new();
        // Section markers are matched as standalone `&`-delimited fields, which is the
        // form Excel writes: `&L&"Calibri,Regular"&11&K000000Left text`.
        hf.set_header("&L&Left text");
        assert!(hf.has_header());
        assert_eq!(hf.left_header.text, Some("&Left text".to_string()));
        let rendered = hf.header_string();
        assert!(rendered.starts_with("&L&\"Calibri,Regular\""));
        assert!(rendered.contains("Left text"));
    }

    #[test]
    fn parses_sections() {
        let mut hf = HeaderFooter::new();
        hf.set_header("&L&left&C&centre&R&right");
        assert_eq!(hf.left_header.text, Some("&left".to_string()));
        assert_eq!(hf.center_header.text, Some("&centre".to_string()));
        assert_eq!(hf.right_header.text, Some("&right".to_string()));
        let rendered = hf.header_string();
        assert!(rendered.starts_with("&L"));
        assert!(rendered.contains("centre"));
        assert!(rendered.ends_with("&right"));
    }

    #[test]
    fn font_and_colour_codes_are_parsed() {
        let mut item = HeaderFooterItem::new(HeaderFooterItem::LEFT);
        item.set_from_header_string(&[
            "L".to_string(),
            "\"Arial\"".to_string(),
            "12".to_string(),
            "KFF0000".to_string(),
            "text".to_string(),
        ]);
        assert_eq!(item.font_name, "Arial");
        assert_eq!(item.font_size, Some(12));
        assert_eq!(item.font_color, "FF0000");
        assert_eq!(item.text, Some("&text".to_string()));
    }

    #[test]
    fn a_marker_glued_to_its_text_is_a_section() {
        // `&Lleft` splits into `["", "Lleft"]`, so the marker is the *start* of a field rather
        // than a field of its own. openpyxl's `_split_string` matches it that way -- the regex is
        // `(&L(?P<left>.+?))?` -- and this is the shape both openpyxl and Excel write, so treating
        // it as an unknown field dropped every header on the next save.
        let mut hf = HeaderFooter::new();
        hf.set_header("&Lleft");
        assert_eq!(hf.left_header.text.as_deref(), Some("left"));
        assert!(hf.has_header());
    }

    #[test]
    fn a_glued_marker_keeps_the_codes_that_follow_it() {
        // The shape the writer itself emits: marker, font name, colour, then the text.
        let mut hf = HeaderFooter::new();
        hf.set_header("&C&\"Calibri,Regular\"&K000000Regional report");
        assert_eq!(hf.center_header.text.as_deref(), Some("Regional report"));
        assert_eq!(hf.center_header.font_name, "Calibri,Regular");
        assert_eq!(hf.center_header.font_color, "000000");
    }

    #[test]
    fn three_glued_sections_are_all_found() {
        let mut hf = HeaderFooter::new();
        hf.set_header("&Lleft text&Cmiddle&Rright");
        assert_eq!(hf.left_header.text.as_deref(), Some("left text"));
        assert_eq!(hf.center_header.text.as_deref(), Some("middle"));
        assert_eq!(hf.right_header.text.as_deref(), Some("right"));
    }

    #[test]
    fn a_glued_footer_with_page_numbers_round_trips() {
        let mut hf = HeaderFooter::new();
        hf.set_footer("&RPage &P of &N");
        assert_eq!(hf.right_footer.text.as_deref(), Some("Page &P of &N"));
        assert!(hf.footer_string().contains("&P"));
        assert!(hf.footer_string().contains("&N"));
    }

    #[test]
    fn the_writers_own_output_reads_back() {
        // The writer emits marker, font name, colour, text. Parsing that back has to keep the
        // text, or a header survives one save and is gone by the second -- which is how the
        // entity-decoding defect that ate every `&` first showed itself.
        let rendered = HeaderFooterItem {
            text: Some("Regional report".to_string()),
            ..HeaderFooterItem::new(HeaderFooterItem::CENTER)
        }
        .to_header_string();
        let mut hf = HeaderFooter::new();
        hf.set_header(&rendered);
        assert_eq!(hf.center_header.text.as_deref(), Some("Regional report"));

        // And the second pass over that must be stable, which is the round trip that was failing.
        let mut again = HeaderFooter::new();
        again.set_header(&hf.header_string());
        assert_eq!(again.center_header.text.as_deref(), Some("Regional report"));
        assert!(again.has_header());
    }

    #[test]
    fn long_forms_are_expanded_on_output() {
        let mut item = HeaderFooterItem::new(HeaderFooterItem::RIGHT);
        item.text = Some("Page &[Page] of &[Pages]".to_string());
        let rendered = item.to_header_string();
        assert!(rendered.contains("&P"));
        assert!(rendered.contains("&N"));
        assert!(!rendered.contains("&[Page]"));
    }

    #[test]
    fn newlines_become_x000d() {
        let mut item = HeaderFooterItem::new(HeaderFooterItem::CENTER);
        item.text = Some("a\nb".to_string());
        assert!(item.to_header_string().contains("_x000D_"));
    }

    #[test]
    fn footer_uses_its_own_sections() {
        let mut hf = HeaderFooter::new();
        hf.set_footer("&L&only footer");
        assert!(hf.has_footer());
        assert!(!hf.has_header());
        assert_eq!(hf.left_footer.text, Some("&only footer".to_string()));
        assert!(hf.footer_string().contains("only footer"));
    }
}
