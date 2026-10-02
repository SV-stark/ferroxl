//! Writing the shared style table (`openpyxl/writer/styles.py`).
//!
//! The writer deduplicates the fonts, fills, borders and number formats referenced by the
//! workbook's cells, assigns each an index, and emits the `cellXfs` records that cells point
//! at. Style identity comes from the same field-based equality the Python `HashableObject`
//! uses.

use std::collections::HashMap;

use crate::styles::alignment::Alignment;
use crate::styles::borders::Border;
use crate::styles::fills::Fill;
use crate::styles::fonts::Font;
use crate::styles::numbers::NumberFormat;
use crate::styles::protection::{Protection, ProtectionFlag};
use crate::styles::style::{defaults, Style};
use crate::workbook::Workbook;
use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::Element;

/// The first custom `numFmtId`; Excel reserves everything below 164 for builtins.
pub const FIRST_CUSTOM_FORMAT_ID: u32 = 165;

/// The identifier assigned to a style in the shared tables.
///
/// The style table starts at 1 because index 0 is the default style, which Excel always
/// expects to exist.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct StyleId(pub usize);

impl StyleId {
    /// The identifier as a number.
    pub fn value(self) -> usize {
        self.0
    }
}

/// The deduplicated style tables built from a workbook.
#[derive(Debug, Clone, Default)]
pub struct StyleTables {
    /// Every distinct style used by the workbook, in `cellXfs` order.
    pub style_list: Vec<Style>,
    /// Style to `cellXfs` index.
    pub style_ids: HashMap<Style, StyleId>,
}

impl StyleTables {
    /// The index assigned to a style, if it is in the table.
    pub fn id_for(&self, style: &Style) -> Option<StyleId> {
        self.style_ids.get(style).copied()
    }

    /// The index of the style at a position in [`StyleTables::style_list`].
    pub fn id_at(&self, position: usize) -> Option<StyleId> {
        self.style_list
            .get(position)
            .and_then(|style| self.id_for(style))
    }

    /// The number of records, including the implicit default at index 0.
    pub fn len(&self) -> usize {
        self.style_list.len() + 1
    }

    /// Whether the workbook needs no records beyond the default.
    pub fn is_empty(&self) -> bool {
        self.style_list.is_empty()
    }
}

/// Build the deduplicated style tables for a workbook.
///
/// Every distinct style across all worksheets is collected, then sorted so the assignment
/// is deterministic rather than dependent on hash iteration order.
pub fn build_style_tables(workbook: &Workbook) -> StyleTables {
    let mut distinct: Vec<Style> = Vec::new();
    for sheet in &workbook.worksheets {
        // Row-level styles are keyed by the row number as a string.
        for style in sheet_style_values(sheet) {
            if !distinct.contains(&style) {
                distinct.push(style);
            }
        }
    }
    // Sorting by a canonical rendering gives a stable, platform-independent ordering. Python
    // relies on hash-table iteration order here, which is insertion-dependent; sorting simply
    // guarantees the output is reproducible across runs.
    distinct.sort_by_key(Style::sort_key);
    let style_ids = distinct
        .iter()
        .enumerate()
        .map(|(index, style)| (style.clone(), StyleId(index + 1)))
        .collect();
    StyleTables {
        style_list: distinct,
        style_ids,
    }
}

fn sheet_style_values(sheet: &crate::worksheet::Worksheet) -> Vec<Style> {
    let mut values = Vec::new();
    for style in sheet.style_values() {
        // Loaded styles carry a flag that is not part of the visual style.
        let mut normalised = style.clone();
        normalised.is_static = false;
        if !values.contains(&normalised) {
            values.push(normalised);
        }
    }
    values
}

/// Serialise `styles.xml`.
pub fn write_style_table(workbook: &Workbook) -> String {
    let tables = build_style_tables(workbook);
    let dxf_list = workbook
        .style_properties
        .as_ref()
        .map(|properties| properties.dxf_list.clone())
        .unwrap_or_default();

    let mut root = Element::new("styleSheet");
    root.set("xmlns", SHEET_MAIN_NS);

    write_number_formats(&mut root, &tables);
    write_fonts(&mut root, &tables);
    write_fills(&mut root, &tables);
    write_borders(&mut root, &tables);
    write_cell_style_xfs(&mut root);
    write_cell_xfs(&mut root, &tables);
    write_cell_styles(&mut root);
    write_dxfs(&mut root, &dxf_list);
    write_table_styles(&mut root);
    root.to_pretty_string()
}

/// Build the font table and its style-to-index map.
fn font_table(tables: &StyleTables) -> (Vec<String>, HashMap<Font, usize>) {
    let default = defaults();
    let mut xml = Vec::new();
    let mut indices = HashMap::new();

    // Excel requires the first font to be the default; index 0 is written below.
    xml.push(default_font_xml());
    let mut index = 1usize;
    for style in &tables.style_list {
        if style.font == default.font {
            continue;
        }
        if indices.contains_key(&style.font) {
            continue;
        }
        indices.insert(style.font.clone(), index);
        xml.push(font_xml(&style.font));
        index += 1;
    }
    (xml, indices)
}

fn default_font_xml() -> String {
    let mut font = Element::new("font");
    font.append(Element::with_attributes("sz", [("val", "11")]));
    font.append(Element::with_attributes("color", [("theme", "1")]));
    font.append(Element::with_attributes("name", [("val", "Calibri")]));
    font.append(Element::with_attributes("family", [("val", "2")]));
    font.append(Element::with_attributes("scheme", [("val", "minor")]));
    font.to_pretty_string()
}

/// A boolean as OOXML writes it: `1` and `0`, not `true` and `false`.
fn flag(value: bool) -> String {
    if value {
        "1".to_string()
    } else {
        "0".to_string()
    }
}

fn font_xml(font: &Font) -> String {
    let mut node = Element::new("font");
    node.append(Element::with_attributes(
        "sz",
        [("val", &format!("{}", font.size))],
    ));
    unpack_color(&mut node, "color", &font.color.index);
    node.append(Element::with_attributes("name", [("val", &font.name)]));
    // `family` was hard-coded to 2, which is what Excel's default font carries. A file using
    // any other family index loaded with the wrong one and saved it back, so the distinction
    // between "not set" and "set to 2" had to be kept.
    let family = if font.family < 0 {
        "2".to_string()
    } else {
        font.family.to_string()
    };
    node.append(Element::with_attributes(
        "family",
        [("val", family.as_str())],
    ));
    if font.bold {
        node.append(Element::new("b"));
    }
    if font.italic {
        node.append(Element::new("i"));
    }
    if font.strikethrough {
        node.append(Element::new("strike"));
    }
    // Every underline style, not just `single`: `doubleAccounting` is what Excel writes for a
    // double underline in an accounting format, and it was being dropped to a bare `<u/>`.
    if !font.underline.is_empty() && font.underline != Font::UNDERLINE_NONE {
        node.append(Element::with_attributes("u", [("val", &font.underline)]));
    }
    if let Some(outline) = font.outline {
        node.append(Element::with_attributes(
            "outline",
            [("val", &flag(outline))],
        ));
    }
    if let Some(shadow) = font.shadow {
        node.append(Element::with_attributes("shadow", [("val", &flag(shadow))]));
    }
    if let Some(condense) = font.condense {
        node.append(Element::with_attributes(
            "condense",
            [("val", &flag(condense))],
        ));
    }
    if let Some(extend) = font.extend {
        node.append(Element::with_attributes("extend", [("val", &flag(extend))]));
    }
    if font.charset >= 0 {
        node.append(Element::with_attributes(
            "charset",
            [("val", &font.charset.to_string())],
        ));
    }
    // A themed font names the theme slot rather than a typeface. Writing `name` as well would
    // be harmless, but omitting `scheme` would pin the font to that name and defeat the theme.
    if !font.scheme.is_empty() {
        node.append(Element::with_attributes("scheme", [("val", &font.scheme)]));
    }
    if font.superscript {
        node.append(Element::with_attributes(
            "vertAlign",
            [("val", "superscript")],
        ));
    } else if font.subscript {
        node.append(Element::with_attributes(
            "vertAlign",
            [("val", "subscript")],
        ));
    }
    node.to_pretty_string()
}

/// Build the fill table and its style-to-index map.
fn fill_table(tables: &StyleTables) -> (Vec<String>, HashMap<Fill, usize>) {
    let default = defaults();
    let mut xml = Vec::new();
    let mut indices = HashMap::new();

    xml.push(plain_fill_xml("none"));
    xml.push(plain_fill_xml("gray125"));
    let mut index = 2usize;
    for style in &tables.style_list {
        if style.fill == default.fill {
            continue;
        }
        if indices.contains_key(&style.fill) {
            continue;
        }
        indices.insert(style.fill.clone(), index);
        xml.push(fill_xml(&style.fill, &default.fill));
        index += 1;
    }
    (xml, indices)
}

/// A `<gradientFill>` element.
///
/// `degree` is only meaningful for a linear gradient; a path gradient uses the direction
/// attributes instead, and Excel ignores `degree` on one. The stops go inside in position
/// order, which is the order they were created in.
fn gradient_xml(fill: &Fill) -> Element {
    let gradient_type = fill
        .fill_type
        .clone()
        .unwrap_or_else(|| Fill::FILL_GRADIENT_LINEAR.to_string());
    let mut attributes: Vec<(&str, String)> = vec![("type", gradient_type.clone())];
    if gradient_type == Fill::FILL_GRADIENT_LINEAR && fill.rotation != 0 {
        attributes.push(("degree", fill.rotation.to_string()));
    }
    let mut node = Element::with_attributes("gradientFill", attributes);
    for stop in &fill.stops {
        let mut child = Element::with_attributes(
            "stop",
            [(
                "position",
                crate::xml::functions::safe_string(stop.position),
            )],
        );
        unpack_color(&mut child, "color", &stop.color.index);
        node.append(child);
    }
    node
}

fn plain_fill_xml(pattern: &str) -> String {
    let mut fill = Element::new("fill");
    fill.append(Element::with_attributes(
        "patternFill",
        [("patternType", pattern)],
    ));
    fill.to_pretty_string()
}

fn fill_xml(fill: &Fill, default: &Fill) -> String {
    let mut node = Element::new("fill");
    if fill.is_gradient() {
        node.append(gradient_xml(fill));
        return node.to_pretty_string();
    }
    if fill.fill_type.as_deref() == default.fill_type.as_deref() {
        return node.to_pretty_string();
    }
    // A fill with no pattern type is serialised as an empty `patternFill` element.
    let pattern = fill.fill_type.clone().unwrap_or_default();
    let mut pattern_node = Element::with_attributes("patternFill", [("patternType", &pattern)]);
    if fill.start_color != default.start_color {
        unpack_color(&mut pattern_node, "fgColor", &fill.start_color.index);
    }
    if fill.end_color != default.end_color {
        unpack_color(&mut pattern_node, "bgColor", &fill.end_color.index);
    }
    node.append(pattern_node);
    node.to_pretty_string()
}

/// Build the border table and its style-to-index map.
fn border_table(
    tables: &StyleTables,
) -> (Vec<String>, HashMap<crate::styles::borders::Borders, usize>) {
    let default = defaults();
    let mut xml = Vec::new();
    let mut indices = HashMap::new();

    xml.push(default_borders_xml());
    let mut index = 1usize;
    for style in &tables.style_list {
        if style.borders == default.borders {
            continue;
        }
        if indices.contains_key(&style.borders) {
            continue;
        }
        indices.insert(style.borders.clone(), index);
        xml.push(borders_xml(&style.borders));
        index += 1;
    }
    (xml, indices)
}

fn default_borders_xml() -> String {
    let mut border = Element::new("border");
    for side in ["left", "right", "top", "bottom", "diagonal"] {
        border.append(Element::new(side));
    }
    border.to_pretty_string()
}

fn borders_xml(borders: &crate::styles::borders::Borders) -> String {
    let mut node = Element::new("border");
    for (side, border) in [
        ("left", &borders.left),
        ("right", &borders.right),
        ("top", &borders.top),
        ("bottom", &borders.bottom),
        ("diagonal", &borders.diagonal),
    ] {
        let mut side_node = Element::new(side);
        match border.border_style.as_deref() {
            None | Some(Border::BORDER_NONE) => {}
            Some(style) => {
                side_node.set("style", style);
                unpack_color(&mut side_node, "color", &border.color.index);
            }
        }
        node.append(side_node);
    }
    node.to_pretty_string()
}

/// Split a colour token into `rgb`, or `theme`/`tint` when it is theme-prefixed.
fn unpack_color(parent: &mut Element, key: &str, color: &str) {
    let mut parts = color.splitn(3, ':');
    match (parts.next(), parts.next(), parts.next()) {
        (Some("theme"), Some(theme), tint) => {
            let mut node = Element::new(key);
            node.set("theme", theme);
            if let Some(tint) = tint {
                if !tint.is_empty() {
                    node.set("tint", tint);
                }
            }
            parent.append(node);
        }
        _ => {
            let mut node = Element::new(key);
            node.set("rgb", color);
            parent.append(node);
        }
    }
}

/// Build the number-format table, assigning custom ids from 165 upwards.
fn number_format_table(tables: &StyleTables) -> (HashMap<NumberFormat, u32>, Vec<(u32, String)>) {
    let mut ids = HashMap::new();
    let mut customs = Vec::new();
    let mut next_id = FIRST_CUSTOM_FORMAT_ID;
    for style in &tables.style_list {
        if ids.contains_key(&style.number_format) {
            continue;
        }
        let code = style.number_format.format_code().to_string();
        match style.number_format.builtin_id() {
            Some(builtin) => {
                ids.insert(style.number_format.clone(), builtin);
            }
            None => {
                ids.insert(style.number_format.clone(), next_id);
                customs.push((next_id, code));
                next_id += 1;
            }
        }
    }
    (ids, customs)
}

fn write_number_formats(root: &mut Element, tables: &StyleTables) {
    let (_, customs) = number_format_table(tables);
    let mut node = Element::new("numFmts");
    node.set("count", customs.len().to_string());
    for (id, code) in customs {
        node.append(Element::with_attributes(
            "numFmt",
            [("numFmtId", &id.to_string()), ("formatCode", &code)],
        ));
    }
    root.append(node);
}

fn write_fonts(root: &mut Element, tables: &StyleTables) {
    let (xml, _) = font_table(tables);
    let mut node = Element::new("fonts");
    node.set("count", xml.len().to_string());
    for entry in xml {
        node.append(crate::xml::functions::fromstring(entry.as_bytes()).unwrap_or_default());
    }
    root.append(node);
}

fn write_fills(root: &mut Element, tables: &StyleTables) {
    let (xml, _) = fill_table(tables);
    let mut node = Element::new("fills");
    node.set("count", xml.len().to_string());
    for entry in xml {
        node.append(crate::xml::functions::fromstring(entry.as_bytes()).unwrap_or_default());
    }
    root.append(node);
}

fn write_borders(root: &mut Element, tables: &StyleTables) {
    let (xml, _) = border_table(tables);
    let mut node = Element::new("borders");
    node.set("count", xml.len().to_string());
    for entry in xml {
        node.append(crate::xml::functions::fromstring(entry.as_bytes()).unwrap_or_default());
    }
    root.append(node);
}

fn write_cell_style_xfs(root: &mut Element) {
    let mut xfs = Element::new("cellStyleXfs");
    xfs.set("count", "1");
    xfs.append(Element::with_attributes(
        "xf",
        [
            ("numFmtId", "0"),
            ("fontId", "0"),
            ("fillId", "0"),
            ("borderId", "0"),
        ],
    ));
    root.append(xfs);
}

fn write_cell_xfs(root: &mut Element, tables: &StyleTables) {
    let default = defaults();
    let (number_formats, _) = number_format_table(tables);
    let (_, fonts) = font_table(tables);
    let (_, fills) = fill_table(tables);
    let (_, borders) = border_table(tables);

    let mut xfs = Element::new("cellXfs");
    xfs.set("count", tables.len().to_string());
    // Index 0 is always the default.
    xfs.append(Element::with_attributes(
        "xf",
        [
            ("numFmtId", "0"),
            ("fontId", "0"),
            ("fillId", "0"),
            ("xfId", "0"),
            ("borderId", "0"),
        ],
    ));
    for style in &tables.style_list {
        let mut node = Element::new("xf");
        node.set("numFmtId", "0");
        node.set("fontId", "0");
        node.set("fillId", "0");
        node.set("xfId", "0");
        node.set("borderId", "0");
        // The number format's id is always written, but the flag that says to apply it was
        // not. Excel reads the id regardless, which is why nothing looked wrong -- and a
        // stricter reader would show a cell with the wrong format.
        if style.number_format != default.number_format {
            node.set("applyNumberFormat", "1");
        }
        if style.font != default.font {
            node.set("fontId", fonts.get(&style.font).unwrap_or(&0).to_string());
            node.set("applyFont", "1");
        }
        if style.borders != default.borders {
            node.set(
                "borderId",
                borders.get(&style.borders).unwrap_or(&0).to_string(),
            );
            node.set("applyBorder", "1");
        }
        // A quote prefix is a *style* property, not a cell value: it makes Excel treat a
        // leading apostrophe as part of the display rather than as text that was escaped.
        // It has to survive a round trip or a numeric-looking string changes meaning.
        if style.quote_prefix {
            node.set("quotePrefix", "1");
        }
        if style.pivot_button {
            node.set("pivotButton", "1");
        }
        if style.fill != default.fill {
            node.set("fillId", fills.get(&style.fill).unwrap_or(&0).to_string());
            node.set("applyFill", "1");
        }
        if style.number_format != default.number_format {
            let id = number_formats
                .get(&style.number_format)
                .copied()
                .unwrap_or(0);
            node.set("numFmtId", id.to_string());
            node.set("applyNumberFormat", "1");
        }
        if style.alignment != default.alignment {
            node.set("applyAlignment", "1");
            let attributes = alignment_attributes(&style.alignment);
            if !attributes.is_empty() {
                node.append(Element::with_attributes("alignment", attributes));
            }
        }
        if style.protection != default.protection {
            node.set("applyProtection", "1");
            let attributes = style.protection.attributes();
            if !attributes.is_empty() {
                node.append(Element::with_attributes("protection", attributes));
            }
        }
        xfs.append(node);
    }
    root.append(xfs);
}

/// The `<alignment/>` attributes for a style that deviates from the default.
fn alignment_attributes(alignment: &Alignment) -> Vec<(String, String)> {
    alignment.attributes()
}

fn write_cell_styles(root: &mut Element) {
    let mut styles = Element::new("cellStyles");
    styles.set("count", "1");
    styles.append(Element::with_attributes(
        "cellStyle",
        [("name", "Normal"), ("xfId", "0"), ("builtinId", "0")],
    ));
    root.append(styles);
}

fn write_dxfs(root: &mut Element, dxf_list: &[crate::formatting::DxfStyle]) {
    let mut dxfs = Element::new("dxfs");
    dxfs.set("count", dxf_list.len().to_string());
    for dxf in dxf_list {
        let mut node = Element::new("dxf");
        if let Some(font) = &dxf.font {
            let mut font_node = Element::new("font");
            unpack_color(&mut font_node, "color", &font.color.index);
            if font.bold {
                font_node.append(Element::with_attributes("b", [("val", "1")]));
            }
            if font.italic {
                font_node.append(Element::with_attributes("i", [("val", "1")]));
            }
            if font.underline != Font::UNDERLINE_NONE {
                font_node.append(Element::with_attributes("u", [("val", &font.underline)]));
            }
            if font.strikethrough {
                font_node.append(Element::new("strike"));
            }
            node.append(font_node);
        }
        if let Some(fill) = &dxf.fill {
            let default = defaults();
            let mut fill_node = Element::new("fill");
            let mut pattern = match &fill.fill_type {
                Some(pattern) => {
                    Element::with_attributes("patternFill", [("patternType", pattern)])
                }
                None => Element::new("patternFill"),
            };
            if fill.start_color != default.fill.start_color {
                unpack_color(&mut pattern, "fgColor", &fill.start_color.index);
            }
            if fill.end_color != default.fill.end_color {
                unpack_color(&mut pattern, "bgColor", &fill.end_color.index);
            }
            fill_node.append(pattern);
            node.append(fill_node);
        }
        if let Some(borders) = &dxf.border {
            let mut border_node = Element::new("border");
            for (side, border) in [
                ("left", &borders.left),
                ("right", &borders.right),
                ("top", &borders.top),
                ("bottom", &borders.bottom),
            ] {
                let mut side_node = Element::new(side);
                match border.border_style.as_deref() {
                    None | Some(Border::BORDER_NONE) => {}
                    Some(style) => {
                        side_node.set("style", style);
                        unpack_color(&mut side_node, "color", &border.color.index);
                    }
                }
                border_node.append(side_node);
            }
            node.append(border_node);
        }
        dxfs.append(node);
    }
    root.append(dxfs);
}

fn write_table_styles(root: &mut Element) {
    root.append(Element::with_attributes(
        "tableStyles",
        [
            ("count", "0"),
            ("defaultTableStyle", "TableStyleMedium9"),
            ("defaultPivotStyle", "PivotStyleLight16"),
        ],
    ));
}

/// Serialise the protection attributes for a style.
pub fn protection_attributes(protection: &Protection) -> Vec<(String, String)> {
    if protection.locked == ProtectionFlag::Protected
        && protection.hidden == ProtectionFlag::Protected
    {
        vec![
            ("locked".to_string(), "1".to_string()),
            ("hidden".to_string(), "1".to_string()),
        ]
    } else {
        protection.attributes()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::styles::colors::Color;
    use crate::styles::style::Style;
    use crate::xml::functions::fromstring;

    fn workbook_with_styles() -> Workbook {
        let mut workbook = Workbook::new();
        let context = workbook.cell_context();
        {
            let sheet = &mut workbook.worksheets[0];
            let mut bold = Style::new();
            bold.font.bold = true;
            bold.font.name = "Arial".to_string();
            bold.font.size = 14.0;
            bold.set_number_format_code("0.00%");
            let mut filled = Style::new();
            filled.fill = crate::styles::Fill::solid(Color::new("FFFF0000"));
            filled.alignment = crate::styles::Alignment::new()
                .with_horizontal("center")
                .with_wrap_text(true);
            sheet.context = context;
            sheet.set("A1", CellValue::text("a")).unwrap();
            sheet.set_style("A1", bold).unwrap();
            sheet.set("A2", CellValue::text("b")).unwrap();
            sheet.set_style("A2", filled).unwrap();
            sheet.set("B1", CellValue::text("c")).unwrap();
        }
        workbook
    }

    #[test]
    fn deduplicates_styles() {
        let workbook = workbook_with_styles();
        let tables = build_style_tables(&workbook);
        assert_eq!(tables.style_list.len(), 2);
        assert_eq!(tables.len(), 3, "index 0 is the implicit default");
    }

    #[test]
    fn identical_styles_share_an_index() {
        let mut workbook = Workbook::new();
        let style = Style::new();
        {
            let sheet = &mut workbook.worksheets[0];
            sheet.set("A1", CellValue::text("a")).unwrap();
            sheet.set("A2", CellValue::text("b")).unwrap();
            sheet.set_style("A1", style.clone()).unwrap();
            sheet.set_style("A2", style).unwrap();
        }
        let tables = build_style_tables(&workbook);
        assert_eq!(tables.style_list.len(), 1);
    }

    #[test]
    fn custom_number_formats_start_at_165() {
        let mut workbook = Workbook::new();
        {
            let sheet = &mut workbook.worksheets[0];
            let mut style = Style::new();
            style.set_number_format_code("0.000");
            sheet.set("A1", CellValue::text("a")).unwrap();
            sheet.set_style("A1", style).unwrap();
        }
        let tables = build_style_tables(&workbook);
        let (ids, customs) = number_format_table(&tables);
        assert_eq!(customs.len(), 1);
        assert_eq!(customs[0].0, FIRST_CUSTOM_FORMAT_ID);
        let format = NumberFormat::with_code("0.000");
        assert_eq!(ids.get(&format), Some(&165));
    }

    #[test]
    fn builtin_number_formats_reuse_their_id() {
        let mut workbook = Workbook::new();
        {
            let sheet = &mut workbook.worksheets[0];
            let mut style = Style::new();
            style.set_number_format_code("0.00%");
            sheet.set("A1", CellValue::text("a")).unwrap();
            sheet.set_style("A1", style).unwrap();
        }
        let tables = build_style_tables(&workbook);
        let (ids, customs) = number_format_table(&tables);
        assert!(customs.is_empty());
        assert_eq!(ids.get(&NumberFormat::with_code("0.00%")), Some(&10));
    }

    #[test]
    fn stylesheet_is_well_formed_and_complete() {
        let workbook = workbook_with_styles();
        let xml = write_style_table(&workbook);
        let root = fromstring(xml.as_bytes()).expect("styles.xml must parse");
        for section in [
            "numFmts",
            "fonts",
            "fills",
            "borders",
            "cellStyleXfs",
            "cellXfs",
            "cellStyles",
            "dxfs",
            "tableStyles",
        ] {
            assert!(root.find(section).is_some(), "missing {section}");
        }
    }

    #[test]
    fn cell_xfs_count_matches_the_tables() {
        let workbook = workbook_with_styles();
        let xml = write_style_table(&workbook);
        let root = fromstring(xml.as_bytes()).unwrap();
        let xfs = root.find("cellXfs").unwrap();
        assert_eq!(xfs.get("count"), Some("3"));
        assert_eq!(xfs.children().len(), 3);
    }

    #[test]
    fn non_default_styles_carry_apply_flags() {
        let workbook = workbook_with_styles();
        let root = fromstring(write_style_table(&workbook).as_bytes()).unwrap();
        let xfs = root.find("cellXfs").unwrap();
        let applied: Vec<&str> = xfs
            .children()
            .iter()
            .skip(1)
            .filter_map(|node| node.get("applyFont").or_else(|| node.get("applyFill")))
            .collect();
        assert!(!applied.is_empty());
    }

    #[test]
    fn theme_colours_are_split_into_theme_and_tint() {
        let mut node = Element::new("font");
        unpack_color(&mut node, "color", "theme:9:0.5");
        let child = node.children()[0].clone();
        assert_eq!(child.get("theme"), Some("9"));
        assert_eq!(child.get("tint"), Some("0.5"));
        assert_eq!(child.get("rgb"), None);

        let mut node = Element::new("font");
        unpack_color(&mut node, "color", "theme:9:");
        let child = node.children()[0].clone();
        assert_eq!(child.get("theme"), Some("9"));
        assert_eq!(child.get("tint"), None);

        let mut node = Element::new("font");
        unpack_color(&mut node, "color", "FFFF0000");
        let child = node.children()[0].clone();
        assert_eq!(child.get("rgb"), Some("FFFF0000"));
        assert_eq!(child.get("theme"), None);
    }

    #[test]
    fn default_fills_are_always_written() {
        let root = fromstring(write_style_table(&Workbook::new()).as_bytes()).unwrap();
        let fills = root.find("fills").unwrap();
        assert_eq!(fills.get("count"), Some("2"));
        let patterns: Vec<Option<&str>> = fills
            .children()
            .iter()
            .map(|fill| fill.find("patternFill").and_then(|p| p.get("patternType")))
            .collect();
        assert_eq!(patterns, vec![Some("none"), Some("gray125")]);
    }

    #[test]
    fn dxf_styles_are_emitted_when_present() {
        let mut workbook = workbook_with_styles();
        let mut properties = workbook.style_properties.clone().unwrap_or_default();
        let mut dxf = crate::formatting::DxfStyle::default();
        let mut font = Font::new();
        font.bold = true;
        font.color = Color::new("FFFF0000");
        dxf.font = Some(font);
        properties.dxf_list.push(dxf);
        workbook.style_properties = Some(properties);

        let root = fromstring(write_style_table(&workbook).as_bytes()).unwrap();
        let dxfs = root.find("dxfs").unwrap();
        assert_eq!(dxfs.get("count"), Some("1"));
        let font_node = dxfs.children()[0].find("font").unwrap();
        assert!(font_node.find("b").is_some());
    }

    #[test]
    fn empty_dxf_list_is_written_as_zero() {
        let root = fromstring(write_style_table(&Workbook::new()).as_bytes()).unwrap();
        assert_eq!(root.find("dxfs").unwrap().get("count"), Some("0"));
    }

    #[test]
    fn style_lookup_by_identity() {
        let workbook = workbook_with_styles();
        let tables = build_style_tables(&workbook);
        let style = &tables.style_list[0];
        assert_eq!(tables.id_for(style), Some(StyleId(1)));
        assert_eq!(tables.id_at(0), Some(StyleId(1)));
        assert_eq!(tables.id_at(99), None);
    }

    #[test]
    fn protection_attributes_match_python_rules() {
        let both = Protection {
            locked: ProtectionFlag::Protected,
            hidden: ProtectionFlag::Protected,
        };
        assert_eq!(
            protection_attributes(&both),
            vec![
                ("locked".to_string(), "1".to_string()),
                ("hidden".to_string(), "1".to_string())
            ]
        );
        let mixed = Protection {
            locked: ProtectionFlag::Unprotected,
            hidden: ProtectionFlag::Inherit,
        };
        assert_eq!(
            protection_attributes(&mixed),
            vec![("locked".to_string(), "0".to_string())]
        );
    }

    #[test]
    fn negative_text_rotation_is_encoded_in_alignment() {
        let alignment = Alignment::new().with_text_rotation(-45);
        let attrs = alignment_attributes(&alignment);
        assert!(attrs.iter().any(|(k, v)| k == "textRotation" && v == "135"));
    }

    #[test]
    fn a_gradient_fill_survives_a_package_round_trip() {
        // This is the test the missing gradient branch failed: a workbook using one loaded
        // with every cell falling back to a plain fill, silently, and it looked correct
        // because a missing gradient is just a background colour.
        use crate::cell::cell::CellValue;
        use crate::styles::colors::Color;
        use crate::styles::fills::Fill;
        use crate::workbook::Workbook;

        let gradient = Fill::linear_gradient(&[Color::new("FF102030"), Color::new("FFA0B0C0")]);
        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut().expect("sheet");
        sheet.set("A1", CellValue::text("x")).expect("cell");
        sheet
            .set_style(
                "A1",
                crate::styles::style::Style {
                    fill: gradient.clone(),
                    ..crate::styles::style::Style::default()
                },
            )
            .expect("styled");
        let bytes = workbook.to_bytes().expect("saved");

        let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
            .expect("loaded");
        let style = loaded.worksheets[0].get_style("A1");

        assert!(style.fill.is_gradient(), "the fill became {:?}", style.fill);
        assert_eq!(
            style.fill.fill_type.as_deref(),
            Some(Fill::FILL_GRADIENT_LINEAR)
        );
        assert_eq!(style.fill.stops.len(), 2);
        assert_eq!(style.fill.stops[0].color.index, "FF102030");
        assert_eq!(style.fill.stops[1].color.index, "FFA0B0C0");
    }
}
