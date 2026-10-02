//! Writing the table part (`xl/tables/tableN.xml`).
//!
//! A table is a separate package part, not part of the sheet. The sheet names it in a
//! `<tableParts>` element, a relationship in `sheetN.xml.rels` points at the part, and
//! `[Content_Types].xml` declares the type. All four have to agree or Excel reports the file
//! as corrupt — and it reports it by refusing to open, not by pointing at the table.
//!
//! Element order inside `<table>` is fixed by the schema: `autoFilter`, `sortState`,
//! `tableColumns`, `tableStyleInfo`.

use crate::exceptions::Result;
use crate::worksheet::table::{Table, TableColumn};
use crate::xml::constants::SHEET_MAIN_NS;
use crate::xml::functions::XmlWriter;

/// Serialise a table to its own part.
pub fn write_table(table: &Table) -> String {
    // `XmlWriter` takes its attributes when the element opens, so the whole set is collected
    // first rather than added afterwards.
    let mut attributes: Vec<(String, String)> = vec![
        ("xmlns".to_string(), SHEET_MAIN_NS.to_string()),
        ("id".to_string(), table.id.to_string()),
        ("name".to_string(), table.formula_name().to_string()),
        ("displayName".to_string(), table.display_name.clone()),
        ("ref".to_string(), table.reference.clone()),
    ];
    let mut put = |name: &str, value: String| attributes.push((name.to_string(), value));

    if let Some(comment) = &table.comment {
        put("comment", comment.clone());
    }
    if let Some(table_type) = &table.table_type {
        put("tableType", table_type.clone());
    }
    if let Some(count) = table.header_row_count {
        put("headerRowCount", count.to_string());
    }
    if let Some(insert) = table.insert_row {
        put("insertRow", xml_flag(insert));
    }
    if let Some(count) = table.totals_row_count {
        put("totalsRowCount", count.to_string());
    }
    if let Some(shown) = table.totals_row_shown {
        put("totalsRowShown", xml_flag(shown));
    }
    for (name, value) in [
        ("headerRowDxfId", table.header_row_dxf_id),
        ("dataDxfId", table.data_dxf_id),
        ("totalsRowDxfId", table.totals_row_dxf_id),
    ] {
        if let Some(id) = value {
            put(name, id.to_string());
        }
    }
    for (name, value) in [
        ("headerRowCellStyle", &table.header_row_cell_style),
        ("dataCellStyle", &table.data_cell_style),
        ("totalsRowCellStyle", &table.totals_row_cell_style),
    ] {
        if let Some(style) = value {
            put(name, style.clone());
        }
    }

    let mut doc = XmlWriter::new();
    doc.start_tag("table", attributes);

    // The element sequence. A table with no autofilter still needs its `tableColumns`.
    write_auto_filter(&mut doc, table);
    doc.start_tag("tableColumns", [("count", table.columns.len().to_string())]);
    for column in &table.columns {
        write_column(&mut doc, column);
    }
    doc.end_tag("tableColumns");

    write_style_info(&mut doc, table);
    doc.end_tag("table");
    doc.into_string()
}

/// A boolean as OOXML spells it.
///
/// `1` and `0` rather than `true` and `false`: both are accepted by the schema's lexical
/// space, but Excel writes the numeric form and a strict reader may not.
fn xml_flag(value: bool) -> String {
    if value {
        "1".to_string()
    } else {
        "0".to_string()
    }
}

/// The `<autoFilter>` over the table's range.
///
/// Written whenever the table has a header row and no filter of its own, matching what Excel
/// expects: a table with a header row has filter dropdowns unless told otherwise.
fn write_auto_filter(doc: &mut XmlWriter, table: &Table) {
    if table.header_row_count.unwrap_or(1) == 0 {
        return;
    }
    doc.tag("autoFilter", [("ref", table.reference.clone())], None);
}

fn write_column(doc: &mut XmlWriter, column: &TableColumn) {
    let mut attributes: Vec<(String, String)> = vec![
        ("id".to_string(), column.id.to_string()),
        ("name".to_string(), column.name.clone()),
    ];
    if let Some(function) = &column.totals_row_function {
        attributes.push(("totalsRowFunction".to_string(), function.clone()));
    }
    if let Some(label) = &column.totals_row_label {
        attributes.push(("totalsRowLabel".to_string(), label.clone()));
    }
    for (name, value) in [
        ("headerRowDxfId", column.header_row_dxf_id),
        ("dataDxfId", column.data_dxf_id),
        ("totalsRowDxfId", column.totals_row_dxf_id),
    ] {
        if let Some(id) = value {
            attributes.push((name.to_string(), id.to_string()));
        }
    }
    for (name, value) in [
        ("headerRowCellStyle", &column.header_row_cell_style),
        ("dataCellStyle", &column.data_cell_style),
        ("totalsRowCellStyle", &column.totals_row_cell_style),
    ] {
        if let Some(style) = value {
            attributes.push((name.to_string(), style.clone()));
        }
    }
    doc.start_tag("tableColumn", attributes);

    // `calculatedColumnFormula` before `totalsRowFormula`: the schema fixes the order, and a
    // reversed pair is a file Excel rejects rather than one it repairs.
    if let Some(formula) = &column.calculated_column_formula {
        write_formula(doc, "calculatedColumnFormula", formula);
    }
    if let Some(formula) = &column.totals_row_formula {
        write_formula(doc, "totalsRowFormula", formula);
    }
    doc.end_tag("tableColumn");
}

fn write_formula(doc: &mut XmlWriter, tag: &str, formula: &crate::worksheet::table::TableFormula) {
    match formula.array {
        Some(array) => {
            doc.tag(tag, [("array", xml_flag(array))], Some(&formula.text));
        }
        None => {
            doc.tag(tag, [] as [(&str, &str); 0], Some(&formula.text));
        }
    }
}

fn write_style_info(doc: &mut XmlWriter, table: &Table) {
    let style = &table.style_info;
    let mut attributes: Vec<(String, String)> = Vec::new();
    if let Some(name) = &style.name {
        attributes.push(("name".to_string(), name.clone()));
    }
    for (name, value) in [
        ("showFirstColumn", style.show_first_column),
        ("showLastColumn", style.show_last_column),
        ("showRowStripes", style.show_row_stripes),
        ("showColumnStripes", style.show_column_stripes),
    ] {
        if let Some(value) = value {
            attributes.push((name.to_string(), xml_flag(value)));
        }
    }
    if attributes.is_empty() {
        return;
    }
    doc.tag("tableStyleInfo", attributes, None);
}

/// Undo the attribute escaping [`XmlWriter`] applied, for reading a column name back.
///
/// The writer escapes on the way out, so the reader is the only place this is needed; a name
/// written and read back has to come out identical or a structured reference built on it
/// points at a column that no longer has that name.
pub fn unescape_name(name: &str) -> String {
    name.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&amp;", "&")
}

/// The `<tableParts>` element for the sheet, naming each table's relationship.
///
/// `None` when the sheet has no tables, because an empty `<tableParts count="0"/>` is noise
/// and Excel rewrites it away.
pub fn write_table_parts(relationship_ids: &[String]) -> Option<String> {
    if relationship_ids.is_empty() {
        return None;
    }
    let mut doc = XmlWriter::new();
    doc.start_tag(
        "tableParts",
        [("count", relationship_ids.len().to_string())],
    );
    for id in relationship_ids {
        doc.tag(
            "tablePart",
            [(
                format!("{}:id", crate::xml::constants::REL_PREFIX),
                id.as_str(),
            )],
            None,
        );
    }
    doc.end_tag("tableParts");
    Some(doc.into_string())
}

/// Read a table part back into a [`Table`].
///
/// A parse failure is reported rather than producing a half-built table: a table read from a
/// file and written back out has to be the same table, and silently dropping a column would
/// make a structured reference point at nothing.
pub fn read_table(xml: &str) -> Result<Table> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut table = Table::default();
    let mut current: Option<TableColumn> = None;
    let mut formula_target: Option<(String, u32)> = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e)) => {
                handle_start(
                    &mut table,
                    &mut current,
                    &mut formula_target,
                    e.clone(),
                    false,
                );
            }
            Ok(Event::Empty(ref e)) => {
                handle_start(
                    &mut table,
                    &mut current,
                    &mut formula_target,
                    e.clone(),
                    true,
                );
            }
            Ok(Event::Text(t)) => {
                if let Some((tag, id)) = &formula_target {
                    let body = t.as_ref().to_string();
                    if let Some(column) =
                        table.columns.iter_mut().find(|c| c.id == *id).or_else(|| {
                            // The column has not been pushed yet, so patch the in-progress one.
                            current.as_mut().filter(|c| c.id == *id)
                        })
                    {
                        let slot = if tag == "calculatedColumnFormula" {
                            &mut column.calculated_column_formula
                        } else {
                            &mut column.totals_row_formula
                        };
                        if let Some(formula) = slot {
                            formula.text = body;
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                let name = local_name(e.name().as_ref().as_bytes());
                match name.as_str() {
                    "tableColumn" => {
                        if let Some(column) = current.take() {
                            table.columns.push(column);
                        }
                    }
                    "calculatedColumnFormula" | "totalsRowFormula" => formula_target = None,
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(crate::exceptions::Error::Xml(e.to_string())),
            _ => {}
        }
    }

    if table.display_name.is_empty() {
        return Err(crate::exceptions::Error::Xml(
            "the table part has no displayName".to_string(),
        ));
    }
    Ok(table)
}

/// Handle an opening element, for both the `Start` and `Empty` cases.
///
/// A self-closing `<tableColumn/>` has no `End` event to collect it on, and that is how a
/// column with neither a formula nor a totals row is written, so the emptiness has to reach
/// the body rather than being inferred from a later event.
fn handle_start(
    table: &mut Table,
    current: &mut Option<TableColumn>,
    formula_target: &mut Option<(String, u32)>,
    e: quick_xml::events::BytesStart,
    empty: bool,
) {
    let name = local_name(e.name().as_ref().as_bytes());
    let attributes = attributes(&e);
    match name.as_str() {
        "table" => {
            table.id = number(&attributes, "id").unwrap_or(1);
            table.display_name = text(&attributes, "displayName").unwrap_or_default();
            table.name = text(&attributes, "name");
            table.reference = text(&attributes, "ref").unwrap_or_default();
            table.comment = text(&attributes, "comment");
            table.table_type = text(&attributes, "tableType");
            table.header_row_count = number(&attributes, "headerRowCount");
            table.insert_row = flag(&attributes, "insertRow");
            table.totals_row_count = number(&attributes, "totalsRowCount");
            table.totals_row_shown = flag(&attributes, "totalsRowShown");
            table.header_row_dxf_id = number(&attributes, "headerRowDxfId");
            table.data_dxf_id = number(&attributes, "dataDxfId");
            table.totals_row_dxf_id = number(&attributes, "totalsRowDxfId");
            table.header_row_cell_style = text(&attributes, "headerRowCellStyle");
            table.data_cell_style = text(&attributes, "dataCellStyle");
            table.totals_row_cell_style = text(&attributes, "totalsRowCellStyle");
        }
        "tableStyleInfo" => {
            table.style_info.name = text(&attributes, "name");
            table.style_info.show_first_column = flag(&attributes, "showFirstColumn");
            table.style_info.show_last_column = flag(&attributes, "showLastColumn");
            table.style_info.show_row_stripes = flag(&attributes, "showRowStripes");
            table.style_info.show_column_stripes = flag(&attributes, "showColumnStripes");
        }
        "tableColumn" => {
            let column = TableColumn {
                id: number(&attributes, "id").unwrap_or(1),
                name: unescape_name(&text(&attributes, "name").unwrap_or_default()),
                totals_row_function: text(&attributes, "totalsRowFunction"),
                totals_row_label: text(&attributes, "totalsRowLabel"),
                header_row_dxf_id: number(&attributes, "headerRowDxfId"),
                data_dxf_id: number(&attributes, "dataDxfId"),
                totals_row_dxf_id: number(&attributes, "totalsRowDxfId"),
                header_row_cell_style: text(&attributes, "headerRowCellStyle"),
                data_cell_style: text(&attributes, "dataCellStyle"),
                totals_row_cell_style: text(&attributes, "totalsRowCellStyle"),
                ..TableColumn::default()
            };
            if empty {
                table.columns.push(column);
            } else {
                *current = Some(column);
            }
        }
        "calculatedColumnFormula" | "totalsRowFormula" => {
            if let Some(column) = current.as_mut() {
                let id = column.id;
                let slot = if name == "calculatedColumnFormula" {
                    &mut column.calculated_column_formula
                } else {
                    &mut column.totals_row_formula
                };
                *slot = Some(crate::worksheet::table::TableFormula {
                    array: flag(&attributes, "array"),
                    text: String::new(),
                });
                *formula_target = Some((name.clone(), id));
            }
        }
        _ => {}
    }
}

fn local_name(raw: &[u8]) -> String {
    let text = String::from_utf8_lossy(raw);
    match text.rsplit_once(':') {
        Some((_, local)) => local.to_string(),
        None => text.to_string(),
    }
}

fn attributes(e: &quick_xml::events::BytesStart) -> Vec<(String, String)> {
    e.attributes()
        .flatten()
        .map(|a| (local_name(a.key.as_ref().as_bytes()), a.value.to_string()))
        .collect()
}

fn text(attributes: &[(String, String)], key: &str) -> Option<String> {
    attributes
        .iter()
        .find(|(name, _)| name == key)
        .map(|(_, value)| value.clone())
}

fn number(attributes: &[(String, String)], key: &str) -> Option<u32> {
    text(attributes, key).and_then(|v| v.parse().ok())
}

fn flag(attributes: &[(String, String)], key: &str) -> Option<bool> {
    text(attributes, key).map(|v| v == "1" || v == "true")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::worksheet::table::{TableFormula, TableStyleInfo};

    #[test]
    fn a_table_writes_its_identity_and_range() {
        let xml = write_table(&Table::new("Sales", "A1:C10").expect("a table"));
        assert!(xml.contains("<table "), "{xml}");
        assert!(xml.contains(r#"name="Sales""#), "{xml}");
        assert!(xml.contains(r#"displayName="Sales""#), "{xml}");
        assert!(xml.contains(r#"ref="A1:C10""#), "{xml}");
        // A header row means filter dropdowns, which openpyxl supplies too.
        assert!(xml.contains(r#"<autoFilter ref="A1:C10">"#), "{xml}");
    }

    #[test]
    fn columns_come_out_inside_table_columns() {
        let table = Table::new("Sales", "A1:C10")
            .expect("a table")
            .with_columns([
                TableColumn::new(1, "Product"),
                TableColumn::new(2, "Region"),
                TableColumn::new(3, "Sales"),
            ]);
        let xml = write_table(&table);
        assert!(xml.contains(r#"<tableColumns count="3">"#), "{xml}");
        for name in ["Product", "Region", "Sales"] {
            assert!(xml.contains(&format!(r#"name="{name}""#)), "{name}: {xml}");
        }
        assert!(xml.contains("</tableColumns>"), "{xml}");
    }

    #[test]
    fn the_element_order_follows_the_schema() {
        // autoFilter, sortState, tableColumns, tableStyleInfo. Excel rejects a rearranged
        // sequence rather than repairing it, so this is asserted rather than assumed.
        let table = Table::new("Sales", "A1:C10")
            .expect("a table")
            .with_style(TableStyleInfo::banded());
        let xml = write_table(&table);
        let filter = xml.find("<autoFilter").expect("autoFilter");
        let columns = xml.find("<tableColumns").expect("tableColumns");
        let style = xml.find("<tableStyleInfo").expect("tableStyleInfo");
        assert!(filter < columns, "{xml}");
        assert!(columns < style, "{xml}");
    }

    #[test]
    fn a_calculated_formula_comes_before_the_totals_formula() {
        let column = TableColumn {
            calculated_column_formula: Some(TableFormula::new("[@Qty]*2")),
            totals_row_formula: Some(TableFormula::new("SUM([@Qty])")),
            ..TableColumn::new(1, "Total")
        };
        let xml = write_table(
            &Table::new("T", "A1:A5")
                .expect("a table")
                .with_column(column),
        );
        let calculated = xml.find("<calculatedColumnFormula").expect("calculated");
        let totals = xml.find("<totalsRowFormula").expect("totals");
        assert!(calculated < totals, "{xml}");
        assert!(!xml.contains("=[@Qty]"), "no leading =: {xml}");
    }

    #[test]
    fn a_column_name_is_escaped() {
        // Excel rejects a name containing a quote or bracket, and openpyxl escapes the same
        // set, so applying it here is what makes a header cell round-trip.
        let xml = write_table(
            &Table::new("T", "A1:A5")
                .expect("a table")
                .with_column(TableColumn::new(1, r#"Q"1" & <b>"#)),
        );
        assert!(xml.contains("&quot;"), "{xml}");
        // The writer escapes attributes itself, so escaping here as well would have produced
        // `&amp;quot;` and a name Excel reads literally.
        assert!(!xml.contains(r#"name="Q"1""#), "{xml}");
        assert_eq!(
            unescape_name("Q&quot;1&quot; &amp; &lt;b&gt;"),
            r#"Q"1" & <b>"#
        );
    }

    #[test]
    fn a_table_with_no_header_row_has_no_autofilter() {
        let mut table = Table::new("T", "A1:C5").expect("a table");
        table.set_header_row(false);
        let xml = write_table(&table);
        assert!(!xml.contains("autoFilter"), "{xml}");
        assert!(xml.contains(r#"headerRowCount="0""#), "{xml}");
    }

    #[test]
    fn table_parts_is_absent_when_there_are_no_tables() {
        assert!(write_table_parts(&[]).is_none());
        let xml = write_table_parts(&["rId1".to_string(), "rId2".to_string()]).expect("parts");
        assert!(xml.contains(r#"<tableParts count="2">"#), "{xml}");
        assert!(xml.contains("rId1"), "{xml}");
        assert!(xml.contains("</tableParts>"), "{xml}");
    }

    #[test]
    fn a_table_round_trips_through_its_own_part() {
        let table = Table::new("Sales", "A1:C10")
            .expect("a table")
            .with_id(3)
            .with_comment("quarterly")
            .with_style(TableStyleInfo::banded().with_row_stripes(false))
            .with_columns([
                TableColumn::new(1, "Product").with_calculated_formula("=[@Qty]*2"),
                TableColumn::new(2, "Region"),
                TableColumn::new(3, "Sales"),
            ]);
        let xml = write_table(&table);
        let read = read_table(&xml).expect("read back");

        assert_eq!(read.display_name, "Sales");
        assert_eq!(read.formula_name(), "Sales");
        assert_eq!(read.reference, "A1:C10");
        assert_eq!(read.comment.as_deref(), Some("quarterly"));
        assert_eq!(read.column_names(), ["Product", "Region", "Sales"]);
        assert_eq!(
            read.columns[0]
                .calculated_column_formula
                .as_ref()
                .expect("a formula")
                .text,
            "[@Qty]*2"
        );
        assert_eq!(read.style_info.name.as_deref(), Some("TableStyleMedium9"));
        assert_eq!(read.style_info.show_row_stripes, Some(false));
        // Re-writing the table it read produces the same part, which is the property that
        // matters: a load/save round trip must not drift.
        assert_eq!(write_table(&read), xml);
    }

    #[test]
    fn an_escaped_column_name_comes_back_unescaped() {
        let table = Table::new("T", "A1:A5")
            .expect("a table")
            .with_column(TableColumn::new(1, r#"Q"1 & <b>"#));
        let read = read_table(&write_table(&table)).expect("read back");
        assert_eq!(read.columns[0].name, r#"Q"1 & <b>"#);
    }

    #[test]
    fn a_table_survives_a_package_round_trip() {
        // The point of all four pieces -- the part, the relationship, the content type and
        // `<tableParts>` -- is that they agree. Checking each in isolation would pass while
        // Excel still refused the file, so this goes through a real save and load.
        use crate::cell::cell::CellValue;
        use crate::workbook::Workbook;

        let mut workbook = Workbook::new();
        let sheet = workbook.active_sheet_mut().expect("the default sheet");
        for (coordinate, text) in [("A1", "Product"), ("B1", "Region"), ("C1", "Sales")] {
            sheet
                .set(coordinate, CellValue::text(text))
                .expect("header");
        }
        sheet
            .add_table(Table::new("Sales", "A1:C10").expect("a table"))
            .expect("added");
        let bytes = workbook.to_bytes().expect("saved");

        let loaded = crate::reader::excel::load_workbook_from_bytes(bytes, Default::default())
            .expect("loaded");
        let sheet = &loaded.worksheets[0];
        let table = sheet.table("Sales").expect("the table survived");

        assert_eq!(table.reference, "A1:C10");
        assert_eq!(table.column_names(), ["Product", "Region", "Sales"]);
    }

    #[test]
    fn a_sheet_with_no_tables_carries_no_table_parts() {
        assert!(write_table_parts(&[]).is_none());
    }

    #[test]
    fn a_part_with_no_display_name_is_rejected() {
        assert!(read_table(
            r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
        )
        .is_err());
    }
}
