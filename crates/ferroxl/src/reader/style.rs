//! The shared style table (`openpyxl/reader/style.py`).
//!
//! `styles.xml` holds a font list, a fill list, a border list, optional differential styles
//! (`dxf`) used by conditional formatting, and the `cellXfs` records that cells index into.
//! The parser builds all of those and resolves each `xf` into a fully-populated [`Style`].

use std::collections::HashMap;

use crate::exceptions::{Error, Result};
use crate::formatting::DxfStyle;
use crate::styles::alignment::Alignment;
use crate::styles::borders::{Border, Borders};
use crate::styles::colors::{Color, COLOR_INDEX};
use crate::styles::fills::{Fill, GradientStop};
use crate::styles::fonts::Font;
use crate::styles::numbers::NumberFormat;
use crate::styles::protection::{Protection, ProtectionFlag};
use crate::styles::style::Style;
use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::{fromstring, Element};

/// The result of parsing `styles.xml`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct StyleTable {
    /// The `cellXfs` records, indexed by position.
    pub table: Vec<Style>,
    /// Differential styles referenced by conditional formatting.
    pub dxf_list: Vec<DxfStyle>,
    /// The indexed colour palette.
    pub color_index: Vec<String>,
}

/// Parse `styles.xml`.
pub fn read_style_table(xml_source: &[u8]) -> Result<StyleTable> {
    let root = fromstring(xml_source)?;
    let mut parser = StyleTableParser {
        root,
        style_prop: StyleTable {
            table: Vec::new(),
            dxf_list: Vec::new(),
            color_index: COLOR_INDEX.iter().map(|c| c.to_string()).collect(),
        },
    };
    parser.parse()?;
    Ok(parser.style_prop)
}

struct StyleTableParser {
    root: Element,
    style_prop: StyleTable,
}

impl StyleTableParser {
    fn parse(&mut self) -> Result<()> {
        let custom_formats = self.parse_custom_num_formats();
        self.parse_color_index();
        let font_list = self.parse_fonts();
        let fill_list = self.parse_fills();
        let border_list = self.parse_borders();
        self.parse_dxfs();
        self.parse_cell_xfs(&custom_formats, &font_list, &fill_list, &border_list)?;
        Ok(())
    }

    fn tag(&self, name: &str) -> String {
        format!("{{{SHEET_MAIN_NS}}}{name}")
    }

    fn parse_custom_num_formats(&self) -> HashMap<u32, String> {
        let mut formats = HashMap::new();
        let Some(num_fmts) = self.root.find(self.tag("numFmts")) else {
            return formats;
        };
        for node in num_fmts.find_all(self.tag("numFmt")) {
            let Some(id) = node
                .get("numFmtId")
                .and_then(|v| v.trim().parse::<u32>().ok())
            else {
                continue;
            };
            let Some(code) = node.get("formatCode") else {
                continue;
            };
            // openpyxl lower-cases custom format codes on read.
            formats.insert(id, code.to_lowercase());
        }
        formats
    }

    fn parse_color_index(&mut self) {
        let Some(colors) = self.root.find(self.tag("colors")) else {
            return;
        };
        let Some(indexed) = colors.find(self.tag("indexedColors")) else {
            return;
        };
        let entries: Vec<String> = indexed
            .find_all(self.tag("rgbColor"))
            .into_iter()
            .filter_map(|node| node.get("rgb").map(|v| v.to_string()))
            .collect();
        if !entries.is_empty() {
            self.style_prop.color_index = entries;
        }
    }

    fn relevant_color(&self, node: &Element) -> Option<String> {
        let index = node
            .get("indexed")
            .and_then(|v| v.trim().parse::<i64>().ok())
            .unwrap_or(-1);
        let theme = node.get("theme");
        let tint = node.get("tint");
        let rgb = node.get("rgb");
        if index >= 0 && (index as usize) < self.style_prop.color_index.len() {
            return Some(self.style_prop.color_index[index as usize].clone());
        }
        if let Some(theme) = theme {
            return Some(match tint {
                Some(tint) => format!("theme:{theme}:{tint}"),
                None => format!("theme:{theme}:"),
            });
        }
        rgb.map(|v| v.to_string())
    }

    fn parse_fonts(&self) -> Vec<Font> {
        let Some(fonts) = self.root.find(self.tag("fonts")) else {
            return Vec::new();
        };
        fonts
            .find_all(self.tag("font"))
            .into_iter()
            .map(|node| self.parse_font(node))
            .collect()
    }

    fn parse_font(&self, node: &Element) -> Font {
        let mut font = Font::new();
        if let Some(size) = node
            .find(self.tag("sz"))
            .and_then(|n| n.get("val"))
            .and_then(|v| v.trim().parse::<f64>().ok())
        {
            font.size = size;
        }
        if let Some(name) = node.find(self.tag("name")).and_then(|n| n.get("val")) {
            font.name = name.to_string();
        }
        // A toggle element's presence turns the flag on unless it carries `val="0"`,
        // which is how Excel writes an explicitly-off flag.
        if let Some(bold) = node.find(self.tag("b")) {
            font.bold = !matches!(bold.get("val"), Some("0") | Some("false"));
        }
        if let Some(italic) = node.find(self.tag("i")) {
            font.italic = !matches!(italic.get("val"), Some("0") | Some("false"));
        }
        if let Some(underline) = node.find(self.tag("u")) {
            font.underline = underline.get("val").unwrap_or("single").to_string();
        }
        if node.find(self.tag("strike")).is_some() {
            font.strikethrough = true;
        }
        // The remaining font attributes were read as nothing at all, so a themed font loaded
        // as its resolved name and was written back without the theme -- pinning it, which is
        // the opposite of what a theme is for.
        let integer = |name: &str| {
            node.find(self.tag(name))
                .and_then(|n| n.get("val"))
                .and_then(|v| v.trim().parse::<i64>().ok())
        };
        if let Some(charset) = integer("charset") {
            font.charset = charset;
        }
        if let Some(family) = integer("family") {
            font.family = family;
        }
        if let Some(scheme) = node.find(self.tag("scheme")).and_then(|n| n.get("val")) {
            font.scheme = scheme.to_string();
        }
        let toggle = |name: &str| {
            node.find(self.tag(name))
                .map(|n| !matches!(n.get("val"), Some("0") | Some("false")))
        };
        font.outline = toggle("outline");
        font.shadow = toggle("shadow");
        font.condense = toggle("condense");
        font.extend = toggle("extend");
        if let Some(vert) = node.find(self.tag("vertAlign")).and_then(|n| n.get("val")) {
            match vert {
                "superscript" => font.superscript = true,
                "subscript" => font.subscript = true,
                _ => {}
            }
        }
        if node.find(self.tag("strike")).is_some() {
            font.strikethrough = true;
        }
        if let Some(color) = node.find(self.tag("color")) {
            if let Some(value) = self.relevant_color(color) {
                font.color = Color::new(value);
            }
        }
        font
    }

    fn parse_fills(&self) -> Vec<Fill> {
        let Some(fills) = self.root.find(self.tag("fills")) else {
            return Vec::new();
        };
        fills
            .find_all(self.tag("fill"))
            .into_iter()
            .filter_map(|node| self.parse_fill(node))
            .collect()
    }

    /// Parse a `<fill>`, which is either a pattern or a gradient.
    ///
    /// The gradient branch was missing, so a workbook using one loaded with every cell
    /// falling back to a plain fill: silent, and it looked correct because a missing gradient
    /// is just a background colour.
    /// Parse a `<gradientFill>` into a [`Fill`].
    ///
    /// The element carries both `type` and a set of `degree`/direction attributes, and the
    /// `<stop>` children are the colours. A gradient with no stops is kept as a gradient with
    /// none rather than being dropped: Excel writes that for a single-colour gradient, and
    /// discarding it would turn a styled cell into an unstyled one.
    fn parse_gradient_fill(&self, node: &Element) -> Fill {
        let mut fill = Fill::new();
        fill.fill_type = Some(
            node.get("type")
                .unwrap_or(Fill::FILL_GRADIENT_LINEAR)
                .to_string(),
        );
        let number = |name: &str| node.get(name).and_then(|v| v.parse::<f64>().ok());
        fill.rotation = number("degree").unwrap_or(0.0) as i64;

        for stop in node.find_all(self.tag("stop")) {
            let position = stop
                .get("position")
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(0.0);
            // A stop's colour is an `<color>` child whose *name* selects the attribute: a
            // stop can carry rgb, theme or indexed and the reader has to look at all three.
            let color = stop
                .find(self.tag("color"))
                .and_then(|c| self.relevant_color(c))
                .unwrap_or(Color::WHITE.to_string());
            fill.stops
                .push(GradientStop::new(position, Color::new(color)));
        }
        fill
    }

    fn parse_fill(&self, node: &Element) -> Option<Fill> {
        if let Some(gradient) = node.find(self.tag("gradientFill")) {
            return Some(self.parse_gradient_fill(gradient));
        }
        let pattern = node.find(self.tag("patternFill"))?;
        let mut fill = Fill::new();
        fill.fill_type = pattern.get("patternType").map(|v| v.to_string());
        if let Some(fg) = pattern.find(self.tag("fgColor")) {
            if let Some(value) = self.relevant_color(fg) {
                fill.start_color = Color::new(value);
            }
        }
        if let Some(bg) = pattern.find(self.tag("bgColor")) {
            if let Some(value) = self.relevant_color(bg) {
                fill.end_color = Color::new(value);
            }
        }
        Some(fill)
    }

    fn parse_borders(&self) -> Vec<Borders> {
        let Some(borders) = self.root.find(self.tag("borders")) else {
            return Vec::new();
        };
        borders
            .find_all(self.tag("border"))
            .into_iter()
            .map(|node| self.parse_border(node))
            .collect()
    }

    fn parse_border(&self, node: &Element) -> Borders {
        let mut borders = Borders::new();
        // The attribute name is `diagonalUp`; the lookup is case-insensitive because the
        // XML in the wild uses both spellings.
        if xml_truthy(node.get("diagonalUp").or_else(|| node.get("diagonalup"))) {
            borders.diagonal_direction = Borders::DIAGONAL_UP;
        }
        if xml_truthy(node.get("diagonalDown")) {
            borders.diagonal_direction = if borders.diagonal_direction == Borders::DIAGONAL_UP {
                Borders::DIAGONAL_BOTH
            } else {
                Borders::DIAGONAL_DOWN
            };
        }
        for side in ["left", "right", "top", "bottom", "diagonal"] {
            let Some(side_node) = node.find(self.tag(side)) else {
                continue;
            };
            let mut border = match side {
                "left" => borders.left.clone(),
                "right" => borders.right.clone(),
                "top" => borders.top.clone(),
                "bottom" => borders.bottom.clone(),
                _ => borders.diagonal.clone(),
            };
            if let Some(style) = side_node.get("style") {
                border.border_style = Some(style.to_string());
            }
            if let Some(color) = side_node.find(self.tag("color")) {
                if let Some(value) = self.relevant_color(color) {
                    border.color = Color::new(value);
                }
            }
            match side {
                "left" => borders.left = border,
                "right" => borders.right = border,
                "top" => borders.top = border,
                "bottom" => borders.bottom = border,
                _ => borders.diagonal = border,
            }
        }
        borders
    }

    fn parse_dxfs(&mut self) {
        let Some(dxfs) = self.root.find(self.tag("dxfs")) else {
            return;
        };
        let mut list = Vec::new();
        for node in dxfs.find_all(self.tag("dxf")) {
            let mut dxf = DxfStyle::default();
            if let Some(font) = node.find(self.tag("font")) {
                dxf.font = Some(self.parse_font(font));
            }
            if let Some(fill) = node.find(self.tag("fill")) {
                dxf.fill = self.parse_fill(fill);
            }
            if let Some(border) = node.find(self.tag("border")) {
                dxf.border = Some(self.parse_border(border));
            }
            list.push(dxf);
        }
        self.style_prop.dxf_list = list;
    }

    fn parse_cell_xfs(
        &mut self,
        custom_formats: &HashMap<u32, String>,
        font_list: &[Font],
        fill_list: &[Fill],
        border_list: &[Borders],
    ) -> Result<()> {
        let Some(cell_xfs) = self.root.find(self.tag("cellXfs")) else {
            // Gnumeric and some other writers omit cellXfs entirely.
            return Ok(());
        };
        let mut table = Vec::new();
        for node in cell_xfs.find_all(self.tag("xf")) {
            let mut style = Style::static_style();
            let number_format_id = node
                .get("numFmtId")
                .and_then(|v| v.trim().parse::<u32>().ok())
                .ok_or_else(|| Error::Value("xf record has no numFmtId".to_string()))?;
            if number_format_id < 164 {
                let code = NumberFormat::builtin_format_code(number_format_id)
                    .unwrap_or(NumberFormat::FORMAT_GENERAL);
                style.number_format.set_format_code(code);
            } else {
                match custom_formats.get(&number_format_id) {
                    Some(code) => style.number_format.set_format_code(code),
                    None => return Err(Error::MissingNumberFormat(format!("{number_format_id}"))),
                }
            }
            // `quotePrefix` and `pivotButton` are per-`xf` attributes and were not read at
            // all. A quote prefix is the one that changes meaning: it makes Excel display a
            // leading apostrophe rather than treat it as an escape, so dropping it turns a
            // shown `'007` into a number on the next save.
            style.quote_prefix = xml_truthy(node.get("quotePrefix"));
            style.pivot_button = xml_truthy(node.get("pivotButton"));
            if xml_truthy(node.get("applyAlignment")) {
                if let Some(alignment) = node.find(self.tag("alignment")) {
                    if let Some(value) = alignment.get("horizontal") {
                        style.alignment.horizontal = value.to_string();
                    }
                    if let Some(value) = alignment.get("vertical") {
                        style.alignment.vertical = value.to_string();
                    }
                    if let Some(value) = alignment.get("indent") {
                        if let Ok(parsed) = value.parse::<i64>() {
                            style.alignment.indent = parsed;
                        }
                    }
                    style.alignment.wrap_text = xml_truthy(alignment.get("wrapText"));
                    style.alignment.shrink_to_fit = xml_truthy(alignment.get("shrinkToFit"));
                    if let Some(relative) = alignment
                        .get("relativeIndent")
                        .and_then(|v| v.trim().parse::<i32>().ok())
                    {
                        style.alignment.relative_indent = relative;
                    }
                    if alignment.get("justifyLastLine").is_some() {
                        style.alignment.justify_last_line =
                            Some(xml_truthy(alignment.get("justifyLastLine")));
                    }
                    if let Some(order) = alignment
                        .get("readingOrder")
                        .and_then(|v| v.trim().parse::<u32>().ok())
                    {
                        style.alignment.reading_order = order;
                    }
                    if let Some(rotation) = alignment
                        .get("textRotation")
                        .and_then(|v| v.trim().parse::<i64>().ok())
                    {
                        style.alignment.text_rotation = rotation;
                    }
                    // `relativeIndent` was read as absent, which turned a hanging indent into
                    // a plain one -- still an indent, so nothing looked wrong.
                    if let Some(relative) = alignment
                        .get("relativeIndent")
                        .and_then(|v| v.trim().parse::<i32>().ok())
                    {
                        style.alignment.relative_indent = relative;
                    }
                    if alignment.get("justifyLastLine").is_some() {
                        style.alignment.justify_last_line =
                            Some(xml_truthy(alignment.get("justifyLastLine")));
                    }
                    if let Some(order) = alignment
                        .get("readingOrder")
                        .and_then(|v| v.trim().parse::<u32>().ok())
                    {
                        style.alignment.reading_order = order;
                    }
                }
            }
            if xml_truthy(node.get("applyFont")) {
                if let Some(font) = node
                    .get("fontId")
                    .and_then(|v| v.trim().parse::<usize>().ok())
                    .and_then(|id| font_list.get(id))
                {
                    style.font = font.clone();
                }
            }
            if xml_truthy(node.get("applyFill")) {
                if let Some(fill) = node
                    .get("fillId")
                    .and_then(|v| v.trim().parse::<usize>().ok())
                    .and_then(|id| fill_list.get(id))
                {
                    style.fill = fill.clone();
                }
            }
            if xml_truthy(node.get("applyBorder")) {
                if let Some(borders) = node
                    .get("borderId")
                    .and_then(|v| v.trim().parse::<usize>().ok())
                    .and_then(|id| border_list.get(id))
                {
                    style.borders = borders.clone();
                }
            }
            if xml_truthy(node.get("applyProtection")) {
                if let Some(protection) = node.find(self.tag("protection")) {
                    style.protection.locked = ProtectionFlag::from_xml(protection.get("locked"));
                    style.protection.hidden = ProtectionFlag::from_xml(protection.get("hidden"));
                }
            }
            table.push(style);
        }
        self.style_prop.table = table;
        Ok(())
    }
}

/// ElementTree truthiness for an attribute: present, non-empty and not `"0"`.
fn xml_truthy(value: Option<&str>) -> bool {
    match value {
        None => false,
        Some("0") | Some("") | Some("false") | Some("False") => false,
        Some(_) => true,
    }
}

/// Helpers for tests and callers building styles from parts.
pub fn font_with_bold(mut font: Font) -> Font {
    font.bold = true;
    font
}

/// Helpers for tests and callers building styles from parts.
pub fn border_with_style(mut border: Border) -> Border {
    border.border_style = Some("thin".to_string());
    border
}

/// Helpers for tests and callers building styles from parts.
pub fn protection_defaults() -> Protection {
    Protection::default()
}

/// Helpers for tests and callers building styles from parts.
pub fn alignment_defaults() -> Alignment {
    Alignment::default()
}

#[cfg(test)]
mod tests {
    use super::*;

    const NS: &str = r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#;

    #[test]
    fn reads_a_minimal_stylesheet() {
        let xml = format!(
            r#"<styleSheet {NS}><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table.len(), 1);
        assert_eq!(table.table[0].number_format.format_code(), "General");
        assert!(table.table[0].is_static);
        assert_eq!(
            table.color_index.len(),
            56,
            "openpyxl ships a 56-entry palette"
        );
        assert!(table.dxf_list.is_empty());
    }

    #[test]
    fn builtin_number_formats_are_resolved() {
        let xml = format!(
            r#"<styleSheet {NS}><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table[0].number_format.format_code(), "General");
        assert_eq!(table.table[1].number_format.format_code(), "mm-dd-yy");
    }

    #[test]
    fn custom_number_formats_are_lower_cased() {
        let xml = format!(
            r#"<styleSheet {NS}><numFmts count="1"><numFmt numFmtId="165" formatCode="0.00&quot;K&quot;"/></numFmts><cellXfs count="1"><xf numFmtId="165"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table[0].number_format.format_code(), "0.00\"k\"");
    }

    #[test]
    fn missing_custom_number_format_errors() {
        let xml = format!(
            r#"<styleSheet {NS}><cellXfs count="1"><xf numFmtId="200"/></cellXfs></styleSheet>"#
        );
        assert!(matches!(
            read_style_table(xml.as_bytes()),
            Err(Error::MissingNumberFormat(_))
        ));
    }

    #[test]
    fn fonts_fills_and_borders_are_applied() {
        let xml = format!(
            r#"<styleSheet {NS}>
              <fonts count="2">
                <font><sz val="11"/><name val="Calibri"/></font>
                <font><b/><i/><u val="double"/><strike/><sz val="14"/><name val="Arial"/><color rgb="FFFF0000"/></font>
              </fonts>
              <fills count="2">
                <fill><patternFill patternType="none"/></fill>
                <fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor indexed="64"/></patternFill></fill>
              </fills>
              <borders count="2">
                <border><left/></border>
                <border diagonalUp="1"><left style="thin"><color rgb="FF0000FF"/></left></border>
              </borders>
              <cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs>
            </styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        let styled = &table.table[1];
        assert!(styled.font.bold);
        assert!(styled.font.italic);
        assert!(styled.font.strikethrough);
        assert_eq!(styled.font.underline, "double");
        assert_eq!(styled.font.size, 14.0);
        assert_eq!(styled.font.color.index, "FFFF0000");
        assert_eq!(styled.fill.fill_type.as_deref(), Some("solid"));
        assert_eq!(styled.fill.start_color.index, "FF00FF00");
        assert_eq!(styled.borders.left.border_style.as_deref(), Some("thin"));
        assert_eq!(styled.borders.diagonal_direction, Borders::DIAGONAL_UP);
    }

    #[test]
    fn alignment_and_protection_are_applied() {
        let xml = format!(
            r#"<styleSheet {NS}><cellXfs count="1"><xf numFmtId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="top" wrapText="1" shrinkToFit="1" indent="2" textRotation="45"/><protection locked="1" hidden="0"/></xf></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        let style = &table.table[0];
        assert_eq!(style.alignment.horizontal, "center");
        assert_eq!(style.alignment.vertical, "top");
        assert!(style.alignment.wrap_text);
        assert!(style.alignment.shrink_to_fit);
        assert_eq!(style.alignment.indent, 2);
        assert_eq!(style.alignment.text_rotation, 45);
        assert_eq!(style.protection.locked, ProtectionFlag::Protected);
        assert_eq!(style.protection.hidden, ProtectionFlag::Unprotected);
    }

    #[test]
    fn colour_index_is_overridden_when_present() {
        let xml = format!(
            r#"<styleSheet {NS}><colors><indexedColors><rgbColor rgb="FF111111"/><rgbColor rgb="FF222222"/></indexedColors></colors><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.color_index, vec!["FF111111", "FF222222"]);
    }

    #[test]
    fn theme_colours_are_encoded() {
        let xml = format!(
            r#"<styleSheet {NS}><fonts count="2"><font/><font><color theme="9"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="1" applyFont="1"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table[0].font.color.index, "theme:9:");

        let xml = format!(
            r#"<styleSheet {NS}><fonts count="2"><font/><font><color theme="9" tint="0.5"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="1" applyFont="1"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table[0].font.color.index, "theme:9:0.5");
        assert_eq!(table.table[0].font.color.theme_tint(), Some("0.5"));
    }

    #[test]
    fn indexed_colours_resolve_through_the_palette() {
        let xml = format!(
            r#"<styleSheet {NS}><colors><indexedColors><rgbColor rgb="FF111111"/><rgbColor rgb="FF222222"/></indexedColors></colors><fonts count="1"><font><color indexed="1"/></font></fonts><cellXfs count="1"><xf numFmtId="0" fontId="0" applyFont="1"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.table[0].font.color.index, "FF222222");
    }

    #[test]
    fn dxfs_are_parsed() {
        let xml = format!(
            r#"<styleSheet {NS}><dxfs count="1"><dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></dxf></dxfs><cellXfs count="1"><xf numFmtId="0"/></cellXfs></styleSheet>"#
        );
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert_eq!(table.dxf_list.len(), 1);
        assert!(table.dxf_list[0].font.as_ref().unwrap().bold);
        assert_eq!(
            table.dxf_list[0].fill.as_ref().unwrap().start_color.index,
            "FFFFFF00"
        );
    }

    #[test]
    fn missing_cell_xfs_is_tolerated() {
        let xml = format!(r#"<styleSheet {NS}><fonts count="1"><font/></fonts></styleSheet>"#);
        let table = read_style_table(xml.as_bytes()).unwrap();
        assert!(table.table.is_empty());
    }

    #[test]
    fn truthiness_rules() {
        assert!(xml_truthy(Some("1")));
        assert!(xml_truthy(Some("true")));
        assert!(!xml_truthy(Some("0")));
        assert!(!xml_truthy(None));
        assert!(!xml_truthy(Some("")));
    }

    #[test]
    fn helper_builders() {
        assert!(font_with_bold(Font::new()).bold);
        assert_eq!(
            border_with_style(Border::new()).border_style.as_deref(),
            Some("thin")
        );
        assert!(protection_defaults().is_default());
        assert!(alignment_defaults().is_default());
    }
}
