//! Writing worksheets to XML (`openpyxl/writer/worksheet.py`).
//!
//! The output is deliberately unindented and written with a streaming writer, matching what
//! the Python implementation produces via `XMLGenerator`. Element order follows the CT_Worksheet
//! schema, which Excel validates on load.

use std::collections::BTreeMap;

use crate::cell::cell::{Cell, CellValue, DataType};
use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::Result;
use crate::formatting::RULE_ATTRIBUTES;
use crate::worksheet::filters::SortCondition;
use crate::worksheet::Worksheet;
use crate::writer::styles::{build_style_tables, StyleId, StyleTables};
use crate::xml::constants::{COMMENTS_NS, PKG_REL_NS, REL_NS, SHEET_MAIN_NS, VML_NS};
use crate::xml::functions::{fromstring, Element, XmlWriter};

/// Serialise a worksheet to its XML part.
///
/// `string_table` maps cell text to its shared-string index, and `style_tables` maps styles
/// to their `cellXfs` index.
///
/// This is the in-memory writer: it builds the whole part as a `String`. `save_dump` is the
/// streaming counterpart, for a sheet large enough that holding its XML matters. The two
/// produce identical bytes, which a test asserts.
pub fn write_worksheet(
    worksheet: &Worksheet,
    string_table: &crate::writer::strings::StringTable,
    style_tables: &StyleTables,
) -> Result<String> {
    let mut doc = XmlWriter::new();
    write_worksheet_head(&mut doc, worksheet, style_tables)?;
    write_sheet_rows(&mut doc, worksheet, string_table, style_tables)?;
    write_worksheet_tail(&mut doc, worksheet)?;
    Ok(doc.into_string())
}

/// The head of a worksheet part as a string, for the streaming writer.
///
/// A thin wrapper over `write_worksheet_head`. The head is bounded by the sheet's
/// configuration rather than by its cell count, so buffering it costs nothing that matters.
pub fn write_worksheet_head_to_string(
    worksheet: &Worksheet,
    style_tables: &StyleTables,
) -> Result<String> {
    let mut doc = XmlWriter::new();
    write_worksheet_head(&mut doc, worksheet, style_tables)?;
    Ok(doc.into_string())
}

/// The tail of a worksheet part as a string, from `</sheetData>` to the end.
pub fn write_worksheet_tail_to_string(worksheet: &Worksheet) -> Result<String> {
    let mut doc = XmlWriter::new();
    write_worksheet_tail(&mut doc, worksheet)?;
    Ok(doc.into_string())
}

/// Write everything up to and including the opening `<sheetData>`.
///
/// The part is split there and after the rows so the streaming writer can reuse both halves
/// unchanged. Neither grows with the size of the sheet's data: the head holds the
/// dimensions, the view and the column definitions, and the tail holds the merges,
/// validations and conditional formats.
fn write_worksheet_head(
    doc: &mut XmlWriter,
    worksheet: &Worksheet,
    style_tables: &StyleTables,
) -> Result<()> {
    doc.start_tag(
        "worksheet",
        [
            attr("xmlns", SHEET_MAIN_NS),
            attr(
                format!("xmlns:{}", crate::xml::constants::REL_PREFIX),
                REL_NS,
            ),
        ],
    );

    let vba_root = match &worksheet.xml_source {
        Some(source) => fromstring(source).ok(),
        None => None,
    };

    // sheetPr: a VBA workbook carries the sheet's code name, which VBA macros reference.
    match vba_root
        .as_ref()
        .and_then(|root| root.find(format!("{{{SHEET_MAIN_NS}}}sheetPr")))
        .and_then(|el| el.get("codeName"))
    {
        Some(code_name) => doc.start_tag("sheetPr", [("codeName", code_name)]),
        None if vba_root.is_some() => doc.start_tag("sheetPr", [("codeName", worksheet.title())]),
        None => doc.start_tag("sheetPr", [] as [(&str, &str); 0]),
    };
    doc.tag(
        "outlinePr",
        [
            ("summaryBelow", bool_str(worksheet.show_summary_below)),
            ("summaryRight", bool_str(worksheet.show_summary_right)),
        ],
        None,
    );
    if worksheet.page_setup.fit_to_page_enabled() {
        doc.tag("pageSetUpPr", [("fitToPage", "1")], None);
    }
    doc.end_tag("sheetPr");

    let dimension = worksheet.calculate_dimension()?;
    doc.tag("dimension", [("ref", &dimension)], None);
    write_sheet_views(doc, worksheet);
    doc.tag("sheetFormatPr", [("defaultRowHeight", "15")], None);
    write_cols(doc, worksheet, style_tables)?;
    doc.start_tag("sheetData", [] as [(&str, &str); 0]);
    Ok(())
}

/// Write everything after `</sheetData>`, and close the worksheet element.
///
/// `vba_root` is the sheet's original XML, kept from the head because a VBA workbook needs
/// the same parse for the code name and for the legacy drawing reference. Parsing the part
/// twice to avoid passing one value between two functions would cost more than it saves.
fn write_worksheet_tail(doc: &mut XmlWriter, worksheet: &Worksheet) -> Result<()> {
    doc.end_tag("sheetData");
    let vba_root = match &worksheet.xml_source {
        Some(source) => fromstring(source).ok(),
        None => None,
    };
    if worksheet.protection.enabled {
        let mut attributes = vec![("objects", "1"), ("scenarios", "1"), ("sheet", "1")];
        if !worksheet.protection.password().is_empty() {
            attributes.push(("password", worksheet.protection.password()));
        }
        doc.tag("sheetProtection", attributes, None);
    }

    write_auto_filter(doc, worksheet);
    write_merge_cells(doc, worksheet);
    write_data_validations(doc, worksheet);
    write_hyperlinks(doc, worksheet);
    write_conditional_formatting(doc, worksheet);

    let options = worksheet.page_setup.option_attributes();
    if !options.is_empty() {
        doc.tag("printOptions", attribute_refs(&options), None);
    }
    let margins = worksheet.page_margins.margin_attributes();
    if !margins.is_empty() {
        doc.tag("pageMargins", attribute_refs(&margins), None);
    }
    let setup = worksheet.page_setup.setup_attributes();
    if !setup.is_empty() {
        doc.tag("pageSetup", attribute_refs(&setup), None);
    }
    if worksheet.header_footer.has_header() || worksheet.header_footer.has_footer() {
        doc.start_tag("headerFooter", [] as [(&str, &str); 0]);
        if worksheet.header_footer.has_header() {
            doc.tag(
                "oddHeader",
                [] as [(&str, &str); 0],
                Some(&worksheet.header_footer.header_string()),
            );
        }
        if worksheet.header_footer.has_footer() {
            doc.tag(
                "oddFooter",
                [] as [(&str, &str); 0],
                Some(&worksheet.header_footer.footer_string()),
            );
        }
        doc.end_tag("headerFooter");
    }
    if !worksheet.charts.is_empty() || !worksheet.images.is_empty() {
        doc.tag(
            "drawing",
            [(format!("{}:id", crate::xml::constants::REL_PREFIX), "rId1")],
            None,
        );
    }
    let table_ids: Vec<String> = worksheet.table_relationship_ids();
    if let Some(parts) = crate::writer::table::write_table_parts(&table_ids) {
        doc.raw(&parts);
    }
    if let Some(root) = &vba_root {
        if let Some(legacy) = root.find(format!("{{{SHEET_MAIN_NS}}}legacyDrawing")) {
            if let Some(r_id) = legacy.get(format!("{{{REL_NS}}}id")) {
                doc.tag(
                    "legacyDrawing",
                    [(format!("{}:id", crate::xml::constants::REL_PREFIX), r_id)],
                    None,
                );
            }
        }
    }
    if !worksheet.page_breaks.is_empty() {
        let count = worksheet.page_breaks.len().to_string();
        doc.start_tag(
            "rowBreaks",
            [("count", &count), ("manualBreakCount", &count)],
        );
        for brk in &worksheet.page_breaks {
            doc.tag(
                "brk",
                [
                    attr("id", brk.to_string()),
                    attr("man", "true"),
                    attr("max", "16383"),
                    attr("min", "0"),
                ],
                None,
            );
        }
        doc.end_tag("rowBreaks");
    }
    if worksheet.comment_count() > 0 {
        doc.tag(
            "legacyDrawing",
            [(
                format!("{}:id", crate::xml::constants::REL_PREFIX),
                "commentsvml",
            )],
            None,
        );
    }
    doc.end_tag("worksheet");
    doc.end_document();
    Ok(())
}

fn bool_str(value: bool) -> &'static str {
    if value {
        "1"
    } else {
        "0"
    }
}

/// An owned attribute, so a literal and a computed value can share one array.
fn attr(key: impl Into<String>, value: impl Into<String>) -> (String, String) {
    (key.into(), value.into())
}

fn attribute_refs(attributes: &[(String, String)]) -> Vec<(&str, &str)> {
    attributes
        .iter()
        .map(|(k, v)| (k.as_str(), v.as_str()))
        .collect()
}

/// Borrow a `BTreeMap`'s entries as attribute pairs.
fn map_attributes(attributes: &std::collections::BTreeMap<String, String>) -> Vec<(&str, &str)> {
    attributes
        .iter()
        .map(|(k, v)| (k.as_str(), v.as_str()))
        .collect()
}

fn write_sheet_views(doc: &mut XmlWriter, worksheet: &Worksheet) {
    doc.start_tag("sheetViews", [] as [(&str, &str); 0]);
    doc.start_tag("sheetView", [("workbookViewId", "0")]);
    let mut selection: Vec<(String, String)> = Vec::new();
    if let Some(top_left) = worksheet.freeze_panes.clone() {
        if let Ok((column_letters, row)) = coordinate_from_string(&top_left) {
            if let Ok(column) = column_index_from_string(&column_letters) {
                let mut pane = "topRight".to_string();
                let mut attributes: Vec<(String, String)> = Vec::new();
                if column > 1 {
                    attributes.push(("xSplit".to_string(), (column - 1).to_string()));
                }
                if row > 1 {
                    attributes.push(("ySplit".to_string(), (row - 1).to_string()));
                    pane = "bottomLeft".to_string();
                    if column > 1 {
                        pane = "bottomRight".to_string();
                    }
                }
                attributes.push(("topLeftCell".to_string(), top_left.clone()));
                attributes.push(("activePane".to_string(), pane.clone()));
                attributes.push(("state".to_string(), "frozen".to_string()));
                doc.tag("pane", attribute_refs(&attributes), None);
                selection.push(("pane".to_string(), pane.clone()));
                if row > 1 && column > 1 {
                    doc.tag("selection", [("pane", "topRight")], None);
                    doc.tag("selection", [("pane", "bottomLeft")], None);
                }
            }
        }
    }
    selection.push(("activeCell".to_string(), worksheet.active_cell.clone()));
    selection.push(("sqref".to_string(), worksheet.selected_cell.clone()));
    doc.tag("selection", attribute_refs(&selection), None);
    doc.end_tag("sheetView");
    doc.end_tag("sheetViews");
}

fn write_cols(
    doc: &mut XmlWriter,
    worksheet: &Worksheet,
    style_tables: &StyleTables,
) -> Result<()> {
    if worksheet.column_dimensions.is_empty() {
        return Ok(());
    }
    doc.start_tag("cols", [] as [(&str, &str); 0]);
    for (letters, dimension) in &worksheet.column_dimensions {
        let col_index = column_index_from_string(letters)?;
        let mut attributes: Vec<(String, String)> = vec![
            ("min".to_string(), col_index.to_string()),
            ("max".to_string(), col_index.to_string()),
        ];
        if dimension.width != -1.0 {
            attributes.push(("customWidth".to_string(), "1".to_string()));
        }
        if !dimension.visible {
            attributes.push(("hidden".to_string(), "true".to_string()));
        }
        if dimension.outline_level > 0 {
            attributes.push((
                "outlineLevel".to_string(),
                dimension.outline_level.to_string(),
            ));
        }
        if dimension.collapsed {
            attributes.push(("collapsed".to_string(), "true".to_string()));
        }
        if dimension.auto_size {
            attributes.push(("bestFit".to_string(), "true".to_string()));
        }
        let style = worksheet.get_style(letters);
        if worksheet.has_style(letters) {
            if let Some(id) = style_tables.id_for(&style) {
                attributes.push(("style".to_string(), id.0.to_string()));
            }
        }
        if dimension.width > 0.0 {
            attributes.push(("width".to_string(), format_number(dimension.width)));
        } else {
            attributes.push(("width".to_string(), "9.10".to_string()));
        }
        doc.tag("col", attribute_refs(&attributes), None);
    }
    doc.end_tag("cols");
    Ok(())
}

/// Format a float for XML: integral values lose the decimal point, as Python's `str` does.
fn format_number(value: f64) -> String {
    if value == value.trunc() {
        format!("{}", value as i64)
    } else {
        format!("{value}")
    }
}

/// Group a sheet's cells into rows, in column order within each row.
///
/// A worksheet stores cells by coordinate, and `"A10"` sorts before `"A2"`, so iteration
/// order is not row order. The rows come out ascending and the cells within each row
/// left to right, which is the order the schema wants.
pub fn group_rows(worksheet: &Worksheet) -> Vec<(u32, Vec<&Cell>)> {
    let mut by_row: BTreeMap<u32, Vec<&Cell>> = BTreeMap::new();
    for cell in worksheet.cells() {
        by_row.entry(cell.row).or_default().push(cell);
    }
    for cells in by_row.values_mut() {
        cells.sort_by_key(|cell| column_index_from_string(&cell.column).unwrap_or(0));
    }
    by_row.into_iter().collect()
}

fn write_sheet_rows(
    doc: &mut XmlWriter,
    worksheet: &Worksheet,
    string_table: &crate::writer::strings::StringTable,
    style_tables: &StyleTables,
) -> Result<()> {
    let max_column = worksheet.highest_column();
    for (row_index, cells) in group_rows(worksheet) {
        write_row(
            doc,
            row_index,
            &cells,
            worksheet,
            string_table,
            style_tables,
            max_column,
        )?;
    }
    Ok(())
}

/// Write one `<row>` element and the cells in it.
pub fn write_row(
    doc: &mut XmlWriter,
    row_index: u32,
    cells: &[&Cell],
    worksheet: &Worksheet,
    string_table: &crate::writer::strings::StringTable,
    style_tables: &StyleTables,
    max_column: u32,
) -> Result<()> {
    let row_dimension = worksheet
        .row_dimensions
        .get(&row_index)
        .cloned()
        .unwrap_or_else(|| crate::worksheet::RowDimension::new(row_index));
    let mut attributes: Vec<(String, String)> = vec![
        ("r".to_string(), row_index.to_string()),
        ("spans".to_string(), format!("1:{max_column}")),
    ];
    if !row_dimension.visible {
        attributes.push(("hidden".to_string(), "1".to_string()));
    }
    if row_dimension.height > 0.0 {
        attributes.push(("ht".to_string(), format_number(row_dimension.height)));
        attributes.push(("customHeight".to_string(), "1".to_string()));
    }
    let row_key = row_index.to_string();
    if worksheet.has_style(&row_key) {
        if let Some(id) = style_tables.id_for(&worksheet.get_style(&row_key)) {
            attributes.push(("s".to_string(), id.0.to_string()));
            attributes.push(("customFormat".to_string(), "1".to_string()));
        }
    }
    doc.start_tag("row", attribute_refs(&attributes));

    for cell in cells {
        let coordinate = cell.coordinate();
        let value = cell.internal_value();
        let mut cell_attributes: Vec<(String, String)> =
            vec![("r".to_string(), coordinate.clone())];
        if cell.data_type != DataType::Formula {
            cell_attributes.push(("t".to_string(), cell.data_type.as_str().to_string()));
        }
        if worksheet.has_style(&coordinate) {
            if let Some(id) = style_tables.id_for(&worksheet.get_style(&coordinate)) {
                cell_attributes.push(("s".to_string(), id.0.to_string()));
            }
        }

        if value.is_empty() {
            doc.tag("c", attribute_refs(&cell_attributes), None);
            continue;
        }

        doc.start_tag("c", attribute_refs(&cell_attributes));
        match cell.data_type {
            DataType::SharedString => {
                let text = value.as_text().unwrap_or_default().to_string();
                let index = string_table.get(&text).copied().unwrap_or(0);
                doc.tag("v", [] as [(&str, &str); 0], Some(&index.to_string()));
            }
            DataType::Formula => {
                if let Some(attributes) = &cell.formula_attributes {
                    let mut formula_attributes: Vec<(String, String)> = Vec::new();
                    if let Some(formula_type) = &attributes.formula_type {
                        formula_attributes.push(("t".to_string(), formula_type.clone()));
                    }
                    if let Some(si) = &attributes.si {
                        formula_attributes.push(("si".to_string(), si.clone()));
                    }
                    // A shared formula with no range is written as an empty `<f/>`:
                    // the followers inherit it from the group's master cell.
                    let bodyless = attributes.formula_type.as_deref() == Some("shared")
                        && attributes.reference.is_none();
                    if bodyless {
                        doc.tag("f", attribute_refs(&formula_attributes), None);
                    } else {
                        let body = value.as_text().unwrap_or_default();
                        let body = body.strip_prefix('=').unwrap_or(body).to_string();
                        doc.tag("f", attribute_refs(&formula_attributes), Some(&body));
                    }
                } else {
                    let body = value.as_text().unwrap_or_default();
                    let body = body.strip_prefix('=').unwrap_or(body).to_string();
                    doc.tag("f", [] as [(&str, &str); 0], Some(&body));
                }
                doc.tag("v", [] as [(&str, &str); 0], None);
            }
            DataType::Numeric => match value {
                CellValue::Number(number) => {
                    doc.tag(
                        "v",
                        [] as [(&str, &str); 0],
                        Some(&crate::xml::functions::repr_float(*number)),
                    );
                }
                _ => {
                    doc.tag("v", [] as [(&str, &str); 0], value.as_text());
                }
            },
            DataType::Bool => {
                let text = match value {
                    CellValue::Bool(flag) => if *flag { "1" } else { "0" }.to_string(),
                    _ => "0".to_string(),
                };
                doc.tag("v", [] as [(&str, &str); 0], Some(&text));
            }
            _ => {
                let text = value.as_text().unwrap_or_default().to_string();
                doc.tag("v", [] as [(&str, &str); 0], Some(&text));
            }
        }
        doc.end_tag("c");
    }
    doc.end_tag("row");
    Ok(())
}

fn write_auto_filter(doc: &mut XmlWriter, worksheet: &Worksheet) {
    let filter = &worksheet.auto_filter;
    if filter.filter_columns().is_empty() && filter.sort_conditions().is_empty() {
        if let Some(reference) = filter.reference() {
            doc.tag("autoFilter", [("ref", reference)], None);
        }
        return;
    }
    let reference = filter.reference().unwrap_or_default().to_string();
    doc.start_tag("autoFilter", [("ref", &reference)]);
    for (col_id, column) in filter.filter_columns() {
        doc.start_tag("filterColumn", [("colId", &col_id.to_string())]);
        doc.start_tag("filters", attribute_refs(&column.filter_attributes()));
        for value in &column.vals {
            doc.tag("filter", [("val", value)], None);
        }
        doc.end_tag("filters");
        doc.end_tag("filterColumn");
    }
    if !filter.sort_conditions().is_empty() {
        doc.start_tag("sortState", [("ref", &reference)]);
        for condition in filter.sort_conditions() {
            // The element name is misspelled in the Python writer and Excel accepts it.
            doc.tag(
                "sortCondtion",
                attribute_refs(&sort_attributes(condition)),
                None,
            );
        }
        doc.end_tag("sortState");
    }
    doc.end_tag("autoFilter");
}

fn sort_attributes(condition: &SortCondition) -> Vec<(String, String)> {
    condition.attributes()
}

fn write_merge_cells(doc: &mut XmlWriter, worksheet: &Worksheet) {
    let merged = worksheet.merged_cells();
    if merged.is_empty() {
        return;
    }
    doc.start_tag("mergeCells", [("count", &merged.len().to_string())]);
    for range in merged {
        doc.tag("mergeCell", [("ref", range)], None);
    }
    doc.end_tag("mergeCells");
}

fn write_data_validations(doc: &mut XmlWriter, worksheet: &Worksheet) {
    let required: Vec<_> = worksheet
        .data_validations
        .iter()
        .filter(|validation| !validation.cells.is_empty() || !validation.ranges.is_empty())
        .collect();
    if required.is_empty() {
        return;
    }
    doc.start_tag("dataValidations", [("count", &required.len().to_string())]);
    for validation in required {
        doc.start_tag("dataValidation", map_attributes(&validation.attributes()));
        if !validation.formula1.is_empty() {
            doc.tag(
                "formula1",
                [] as [(&str, &str); 0],
                Some(&validation.formula1),
            );
        }
        if !validation.formula2.is_empty() {
            doc.tag(
                "formula2",
                [] as [(&str, &str); 0],
                Some(&validation.formula2),
            );
        }
        doc.end_tag("dataValidation");
    }
    doc.end_tag("dataValidations");
}

fn write_hyperlinks(doc: &mut XmlWriter, worksheet: &Worksheet) {
    let linked: Vec<&Cell> = worksheet.cells_with_hyperlinks().collect();
    if linked.is_empty() {
        return;
    }
    doc.start_tag("hyperlinks", [] as [(&str, &str); 0]);
    for cell in linked {
        let mut attributes: Vec<(String, String)> = vec![
            ("display".to_string(), cell.hyperlink().to_string()),
            ("ref".to_string(), cell.coordinate()),
            (
                format!("{}:id", crate::xml::constants::REL_PREFIX),
                cell.hyperlink_rel_id.clone().unwrap_or_default(),
            ),
        ];
        attributes.retain(|(_, value)| !value.is_empty());
        doc.tag("hyperlink", attribute_refs(&attributes), None);
    }
    doc.end_tag("hyperlinks");
}

fn write_conditional_formatting(doc: &mut XmlWriter, worksheet: &Worksheet) {
    for (range, rules) in &worksheet.conditional_formatting.cf_rules {
        if rules.is_empty() {
            continue;
        }
        doc.start_tag("conditionalFormatting", [("sqref", range)]);
        for rule in rules {
            let mut attributes: Vec<(String, String)> =
                vec![("type".to_string(), rule.rule_type.clone())];
            for name in RULE_ATTRIBUTES {
                if let Some(value) = rule.attributes.get(name) {
                    attributes.push((name.to_string(), value.clone()));
                }
            }
            doc.start_tag("cfRule", attribute_refs(&attributes));
            for formula in &rule.formula {
                doc.tag("formula", [] as [(&str, &str); 0], Some(formula));
            }
            if let Some(scale) = &rule.color_scale {
                doc.start_tag("colorScale", [] as [(&str, &str); 0]);
                for cfvo in &scale.cfvo {
                    doc.tag("cfvo", attribute_refs(&cfvo.attributes()), None);
                }
                for color in &scale.color {
                    write_scale_color(doc, &color.index);
                }
                doc.end_tag("colorScale");
            }
            if let Some(icon_set) = &rule.icon_set {
                let mut attributes: Vec<(String, String)> = Vec::new();
                if let Some(name) = &icon_set.icon_set {
                    attributes.push(("iconSet".to_string(), name.clone()));
                }
                if let Some(value) = &icon_set.show_value {
                    attributes.push(("showValue".to_string(), value.clone()));
                }
                if let Some(value) = &icon_set.reverse {
                    attributes.push(("reverse".to_string(), value.clone()));
                }
                if let Some(value) = &icon_set.percent {
                    attributes.push(("percent".to_string(), value.clone()));
                }
                doc.start_tag("iconSet", attribute_refs(&attributes));
                for cfvo in &icon_set.cfvo {
                    doc.tag("cfvo", attribute_refs(&cfvo.attributes()), None);
                }
                doc.end_tag("iconSet");
            }
            if let Some(bar) = &rule.data_bar {
                // The thresholds come before the colour: the schema fixes that order.
                let mut attributes: Vec<(String, String)> = Vec::new();
                let mut put = |name: &str, value: Option<String>| {
                    if let Some(value) = value {
                        attributes.push((name.to_string(), value));
                    }
                };
                put(
                    "showValue",
                    bar.show_value
                        .map(|v| if v { "1".into() } else { "0".into() }),
                );
                put("minLength", bar.min_length.map(|v| v.to_string()));
                put("maxLength", bar.max_length.map(|v| v.to_string()));
                doc.start_tag("dataBar", attribute_refs(&attributes));
                for cfvo in &bar.cfvo {
                    doc.tag("cfvo", attribute_refs(&cfvo.attributes()), None);
                }
                if !bar.color.is_empty() {
                    write_scale_color(doc, &bar.color);
                }
                doc.end_tag("dataBar");
            }
            doc.end_tag("cfRule");
        }
        doc.end_tag("conditionalFormatting");
    }
}

fn write_scale_color(doc: &mut XmlWriter, color: &str) {
    let parts: Vec<&str> = color.split(':').collect();
    if parts.first() == Some(&"theme") && parts.len() >= 2 {
        let mut attributes: Vec<(String, String)> =
            vec![("theme".to_string(), parts[1].to_string())];
        if let Some(tint) = parts.get(2) {
            if !tint.is_empty() {
                attributes.push(("tint".to_string(), (*tint).to_string()));
            }
        }
        doc.tag("color", attribute_refs(&attributes), None);
    } else {
        doc.tag("color", [("rgb", color)], None);
    }
}

/// Write the relationships part for a worksheet.
pub fn write_worksheet_rels(
    worksheet: &Worksheet,
    drawing_id: u32,
    comments_id: u32,
    tables: &[(String, u32)],
) -> String {
    let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
    for relationship in &worksheet.relationships {
        let mut node = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        if let Some(id) = &relationship.id {
            node.set("Id", id.clone());
        }
        node.set("Type", relationship.type_uri());
        if let Some(target) = &relationship.target {
            node.set("Target", target.clone());
        }
        if let Some(mode) = &relationship.target_mode {
            node.set("TargetMode", mode.clone());
        }
        root.append(node);
    }
    if !worksheet.charts.is_empty() || !worksheet.images.is_empty() {
        let mut node = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        node.set("Id", "rId1");
        node.set("Type", format!("{REL_NS}/drawing"));
        node.set("Target", format!("../drawings/drawing{drawing_id}.xml"));
        root.append(node);
    }
    // One relationship per table part. The id is what the sheet's `<tableParts>` refers to,
    // so it has to be the same string in both places.
    for (id, table_id) in tables {
        let mut node = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        node.set("Id", id.clone());
        node.set("Type", crate::writer::workbook::TABLE_REL_TYPE.to_string());
        node.set("Target", format!("../tables/table{table_id}.xml"));
        root.append(node);
    }
    if worksheet.comment_count() > 0 {
        let mut node = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        node.set("Id", "comments");
        node.set("Type", COMMENTS_NS);
        node.set("Target", format!("../comments{comments_id}.xml"));
        root.append(node);

        let mut node = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        node.set("Id", "commentsvml");
        node.set("Type", VML_NS);
        node.set(
            "Target",
            format!("../drawings/commentsDrawing{comments_id}.vml"),
        );
        root.append(node);
    }
    root.to_pretty_string()
}

/// Build the style tables for a workbook, exposed for callers assembling parts directly.
pub fn style_tables_for(workbook: &crate::workbook::Workbook) -> StyleTables {
    build_style_tables(workbook)
}

/// The style id at a table position, exposed for callers assembling parts directly.
pub fn style_id_at(tables: &StyleTables, position: usize) -> Option<StyleId> {
    tables.id_at(position)
}

/// Helper: the column letters for a cell's column, for tests and callers.
pub fn cell_column_letters(cell: &Cell) -> Result<String> {
    let (letters, _) = coordinate_from_string(&cell.coordinate())?;
    get_column_letter(column_index_from_string(&letters)?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::cell::CellContext;
    use crate::styles::style::Style;
    use crate::writer::strings::create_string_table;
    use crate::xml::functions::fromstring;

    fn write_one(worksheet: &Worksheet) -> String {
        // Build the shared tables from this worksheet alone so the test does not need a
        // whole workbook.
        let tables = tables_from_sheets(std::slice::from_ref(worksheet));
        let string_table = create_string_table(std::slice::from_ref(worksheet));
        write_worksheet(worksheet, &string_table, &tables).unwrap()
    }

    fn tables_from_sheets(sheets: &[Worksheet]) -> StyleTables {
        let mut workbook = crate::workbook::Workbook::new();
        workbook.worksheets.clear();
        workbook.worksheets.extend_from_slice(sheets);
        build_style_tables(&workbook)
    }

    fn new_sheet() -> Worksheet {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        sheet.context = CellContext::default();
        sheet
    }

    #[test]
    fn writes_a_minimal_worksheet() {
        let sheet = new_sheet();
        let xml = write_one(&sheet);
        assert!(xml.starts_with("<worksheet"));
        assert!(xml.contains("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\""));
        assert!(xml.ends_with("</worksheet>"));
        assert!(xml.contains("<dimension ref=\"A1:A1\"></dimension>"));
        assert!(xml.contains("<sheetFormatPr defaultRowHeight=\"15\"></sheetFormatPr>"));
    }

    #[test]
    fn writes_cell_values_by_type() {
        let mut sheet = new_sheet();
        sheet.set("A1", "hello").unwrap();
        sheet.set("B1", 42.0).unwrap();
        sheet.set("C1", true).unwrap();
        let xml = write_one(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        let row = root
            .find("sheetData")
            .and_then(|data| data.find("row"))
            .expect("row");
        let cells = row.find_all("c");
        assert_eq!(cells.len(), 3);
        assert_eq!(cells[0].get("t"), Some("s"));
        assert_eq!(cells[0].find_text("v", ""), "0");
        assert_eq!(cells[1].get("t"), Some("n"));
        assert_eq!(cells[1].find_text("v", ""), "42");
        assert_eq!(cells[2].get("t"), Some("b"));
        assert_eq!(cells[2].find_text("v", ""), "1");
    }

    #[test]
    fn formula_cells_have_no_type_and_an_empty_v() {
        let mut sheet = new_sheet();
        sheet.set("A1", "=SUM(B1:B2)").unwrap();
        let xml = write_one(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        let cell = root
            .find("sheetData")
            .and_then(|d| d.find("row"))
            .and_then(|r| r.find("c"))
            .expect("cell");
        assert_eq!(cell.get("t"), None);
        assert_eq!(cell.find_text("f", ""), "SUM(B1:B2)");
        assert!(cell.find("v").is_some());
    }

    #[test]
    fn shared_formulas_without_a_range_write_an_empty_f() {
        let mut sheet = new_sheet();
        sheet.set("A1", "=B1*2").unwrap();
        let cell = sheet.cell_mut("A1").unwrap();
        cell.formula_attributes = Some(crate::cell::cell::FormulaAttributes {
            formula_type: Some("shared".to_string()),
            si: Some("0".to_string()),
            reference: None,
        });
        let xml = write_one(&sheet);
        assert!(xml.contains("si=\"0\""));
        assert!(xml.contains("<f t=\"shared\" si=\"0\"></f>"));
    }

    #[test]
    fn merged_cells_and_dimension() {
        let mut sheet = new_sheet();
        sheet.set("A1", "x").unwrap();
        sheet.merge_cells("A1:B2").unwrap();
        sheet.set("C5", 1.0).unwrap();
        let xml = write_one(&sheet);
        assert!(xml.contains("<mergeCells count=\"1\">"));
        assert!(xml.contains("<mergeCell ref=\"A1:B2\"></mergeCell>"));
        assert!(xml.contains("ref=\"A1:C5\""));
    }

    #[test]
    fn freeze_panes_produce_a_pane_element() {
        let mut sheet = new_sheet();
        sheet.set_freeze_panes("B2");
        let xml = write_one(&sheet);
        assert!(xml.contains("<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"></pane>"));
        assert!(xml.contains("pane=\"bottomRight\""));
    }

    #[test]
    fn column_dimensions_are_written() {
        let mut sheet = new_sheet();
        let mut dimension = crate::worksheet::ColumnDimension::new("A");
        dimension.width = 20.0;
        sheet.column_dimensions.insert("A".to_string(), dimension);
        let xml = write_one(&sheet);
        assert!(xml.contains("<col min=\"1\" max=\"1\" customWidth=\"1\" width=\"20\"></col>"));
    }

    #[test]
    fn auto_filter_variants() {
        let mut sheet = new_sheet();
        sheet.set("A1", "x").unwrap();
        sheet.auto_filter.set_reference("A1:C3");
        let xml = write_one(&sheet);
        assert!(xml.contains("<autoFilter ref=\"A1:C3\"></autoFilter>"));

        let mut sheet = new_sheet();
        sheet.auto_filter.set_reference("A1:C3");
        sheet
            .auto_filter
            .add_filter_column(0, vec!["a".to_string()], true);
        sheet.auto_filter.add_sort_condition("A2:A3", true);
        let xml = write_one(&sheet);
        assert!(xml.contains("<filterColumn colId=\"0\">"));
        assert!(xml.contains("<filters blank=\"1\">"));
        assert!(xml.contains("<filter val=\"a\"></filter>"));
        assert!(xml.contains("<sortCondtion ref=\"A2:A3\" descending=\"1\"></sortCondtion>"));
    }

    #[test]
    fn data_validations_are_written() {
        let mut sheet = new_sheet();
        let mut validation = crate::datavalidation::DataValidation::new(
            crate::datavalidation::ValidationType::Whole,
            Some(crate::datavalidation::ValidationOperator::Between),
            Some("1"),
            Some("10"),
            true,
        );
        validation.add_cell("A1");
        validation.add_cell("A2");
        sheet.add_data_validation(validation);
        let xml = write_one(&sheet);
        assert!(xml.contains("<dataValidations count=\"1\">"));
        assert!(xml.contains("type=\"whole\""));
        assert!(xml.contains("sqref=\"A1:A2\""));
        assert!(xml.contains("<formula1>1</formula1>"));
        assert!(xml.contains("<formula2>10</formula2>"));
    }

    #[test]
    fn hyperlinks_are_written_with_relationship_ids() {
        let mut sheet = new_sheet();
        sheet.set_hyperlink("A1", "http://example.com").unwrap();
        let xml = write_one(&sheet);
        assert!(xml.contains("<hyperlinks>"));
        assert!(xml.contains("r:id=\"rId1\""));
        assert!(xml.contains("ref=\"A1\""));
    }

    #[test]
    fn page_setup_and_margins() {
        let mut sheet = new_sheet();
        sheet
            .set_printer_settings("9", Worksheet::ORIENTATION_LANDSCAPE)
            .unwrap();
        sheet.page_margins = crate::worksheet::PageMargins::new().with_defaults();
        let xml = write_one(&sheet);
        assert!(xml.contains("<pageMargins left=\"0.70\""));
        assert!(xml.contains("<pageSetup orientation=\"landscape\" paperSize=\"9\"></pageSetup>"));
    }

    #[test]
    fn header_footer_round_trip() {
        let mut sheet = new_sheet();
        sheet.header_footer.set_header("&L&report");
        let xml = write_one(&sheet);
        assert!(xml.contains("<oddHeader>&amp;L"));
    }

    #[test]
    fn protection_is_written_when_enabled() {
        let mut sheet = new_sheet();
        sheet.protection.enable();
        let xml = write_one(&sheet);
        assert!(xml.contains(
            "<sheetProtection objects=\"1\" scenarios=\"1\" sheet=\"1\"></sheetProtection>"
        ));
    }

    #[test]
    fn conditional_formatting_is_written() {
        let mut sheet = new_sheet();
        let rule = crate::formatting::rules::CellIsRule::new(Some(">"), Some("5"), true).to_rule();
        sheet.conditional_formatting.add("A1:A4", rule);
        let xml = write_one(&sheet);
        assert!(xml.contains("<conditionalFormatting sqref=\"A1:A4\">"));
        assert!(xml.contains("type=\"cellIs\""));
        assert!(xml.contains("operator=\"greaterThan\""));
        assert!(xml.contains("stopIfTrue=\"1\""));
    }

    #[test]
    fn colour_scales_split_theme_and_tint() {
        let mut sheet = new_sheet();
        let rule = crate::formatting::rules::ColorScaleRule::new(
            Some("min"),
            None,
            Some(crate::styles::Color::theme("9", None)),
            None,
            None,
            None,
            Some("max"),
            None,
            Some(crate::styles::Color::new("FF4F81BD")),
        )
        .to_rule();
        sheet.conditional_formatting.add("A1:A4", rule);
        let xml = write_one(&sheet);
        assert!(xml.contains("<colorScale>"));
        assert!(xml.contains("theme=\"9\""));
        assert!(xml.contains("<color rgb=\"FF4F81BD\"></color>"));
        // The theme slot with an empty tint must not emit a tint attribute.
        assert!(!xml.contains("tint="));
    }

    #[test]
    fn styled_cells_carry_the_style_index() {
        let mut sheet = new_sheet();
        sheet.set("A1", "x").unwrap();
        let mut style = Style::new();
        style.font.bold = true;
        sheet.set_style("A1", style).unwrap();
        // The worksheet carries only the index into the shared style tables; the font itself
        // lives in `xl/styles.xml`.
        let root = fromstring(write_one(&sheet).as_bytes()).unwrap();
        let cell = root
            .find("sheetData")
            .and_then(|d| d.find("row"))
            .and_then(|r| r.find("c"))
            .unwrap();
        let style_id = cell
            .get("s")
            .and_then(|id| id.parse::<usize>().ok())
            .unwrap();
        assert_eq!(style_id, 1, "the bold style is not the default");
    }

    #[test]
    fn row_ordering_is_by_row_then_column() {
        let mut sheet = new_sheet();
        sheet.set("B2", 1.0).unwrap();
        sheet.set("A2", 2.0).unwrap();
        sheet.set("A1", 3.0).unwrap();
        let xml = write_one(&sheet);
        let root = fromstring(xml.as_bytes()).unwrap();
        let rows = root
            .find("sheetData")
            .unwrap()
            .find_all("row")
            .iter()
            .map(|row| row.get("r").unwrap_or("").to_string())
            .collect::<Vec<_>>();
        assert_eq!(rows, vec!["1", "2"]);
        let second_row = root.find("sheetData").unwrap().find_all("row")[1].clone();
        let refs: Vec<String> = second_row
            .find_all("c")
            .iter()
            .map(|c| c.get("r").unwrap_or("").to_string())
            .collect();
        assert_eq!(refs, vec!["A2", "B2"]);
    }

    #[test]
    fn relationships_include_drawings_and_comments() {
        let mut sheet = new_sheet();
        sheet.set("A1", "x").unwrap();
        sheet
            .set_comment("A1", Some(crate::comments::Comment::new("note", "me")))
            .unwrap();
        let rels = write_worksheet_rels(&sheet, 1, 1, &[]);
        assert!(rels.contains("Id=\"comments\""));
        assert!(rels.contains("../comments1.xml"));
        assert!(rels.contains("commentsDrawing1.vml"));
    }

    #[test]
    fn empty_cells_are_written_as_self_closing() {
        let mut sheet = new_sheet();
        sheet.set("A1", CellValue::None).unwrap();
        let xml = write_one(&sheet);
        assert!(xml.contains("<c r=\"A1\" t=\"s\"></c>"));
    }

    #[test]
    fn float_formatting_drops_trailing_zeros() {
        assert_eq!(format_number(20.0), "20");
        assert_eq!(format_number(9.1), "9.1");
    }
}
