//! Workbook-level metadata (`openpyxl/reader/workbook.py`).
//!
//! Reads core properties, the active tab, the date system, the content types, the workbook
//! relationships and the defined names.

use std::collections::HashMap;

use chrono::NaiveDateTime;

use crate::date_time::{w3cdtf_to_datetime, BaseDate};
use crate::exceptions::Result;
use crate::namedrange::{
    refers_to_range, split_named_range, DefinedName, NamedRange, NamedRangeContainingValue,
};
use crate::workbook::DocumentProperties;
use crate::xml::constants::{
    CONTYPES_NS, COREPROPS_NS, DCORE_NS, DCTERMS_NS, PKG_REL_NS, REL_NS, SHEET_MAIN_NS,
};
use crate::xml::functions::{fromstring, Element};

/// Names that are dropped when reading, because their definitions are unreliable.
pub const BUGGY_NAMED_RANGES: [&str; 2] = ["NA()", "#REF!"];

/// Name fragments that cause a defined name to be discarded.
pub const DISCARDED_RANGES: [&str; 2] = ["Excel_BuiltIn", "Print_Area"];

/// The content type of a worksheet part.
pub const VALID_WORKSHEET: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

/// Read the core document properties.
pub fn read_properties_core(xml_source: &[u8]) -> Result<DocumentProperties> {
    let root = fromstring(xml_source)?;
    let mut properties = DocumentProperties::new();
    properties.creator = root.find_text(format!("{{{DCORE_NS}}}creator"), "");
    properties.last_modified_by = root.find_text(format!("{{{COREPROPS_NS}}}lastModifiedBy"), "");
    let created = root.find(format!("{{{DCTERMS_NS}}}created"));
    properties.created = match created.and_then(|n| n.text.clone()) {
        Some(text) => w3cdtf_to_datetime(&text).unwrap_or_default(),
        None => NaiveDateTime::default(),
    };
    let modified = root.find(format!("{{{DCTERMS_NS}}}modified"));
    properties.modified = match modified.and_then(|n| n.text.clone()) {
        Some(text) => w3cdtf_to_datetime(&text).unwrap_or(properties.created),
        None => properties.created,
    };
    Ok(properties)
}

/// Read the `date1904` workbook property.
pub fn read_excel_base_date(xml_source: &[u8]) -> Result<BaseDate> {
    let root = fromstring(xml_source)?;
    let base = root
        .find(format!("{{{SHEET_MAIN_NS}}}workbookPr"))
        .and_then(|node| node.get("date1904"))
        .map(|flag| BaseDate::from_date1904_flag(Some(flag)))
        .unwrap_or(BaseDate::Windows1900);
    Ok(base)
}

/// Read the active tab from the workbook view.
/// Read `<calcPr>` back into `CalcProperties`.
///
/// Absent entirely from the reader before this, so a workbook set to manual calculation loaded
/// as automatic: and ferroxl writes formulas without cached results, so "automatic" is the
/// difference between a workbook that shows values when Excel opens it and one that does not.
pub fn read_calc_properties(xml_source: &[u8]) -> crate::workbook::CalcProperties {
    let Ok(root) = crate::xml::functions::fromstring(xml_source) else {
        return crate::workbook::CalcProperties::default();
    };
    let Some(node) = root.find(format!("{{{SHEET_MAIN_NS}}}calcPr")) else {
        return crate::workbook::CalcProperties::default();
    };
    use crate::workbook::CalcProperties;
    let number = |name: &str| node.get(name).and_then(|v| v.trim().parse::<u32>().ok());
    let boolean = |name: &str| node.get(name).map(|v| v != "0" && v != "false");
    CalcProperties {
        calc_id: number("calcId").unwrap_or(124_519),
        calc_mode: node.get("calcMode").map(|v| v.to_string()),
        full_calc_on_load: boolean("fullCalcOnLoad"),
        ref_mode: node.get("refMode").map(|v| v.to_string()),
        iterate: boolean("iterate"),
        iterate_count: number("iterateCount"),
        iterate_delta: node
            .get("iterateDelta")
            .and_then(|v| v.trim().parse::<f64>().ok()),
        full_precision: boolean("fullPrecision"),
        calc_completed: boolean("calcCompleted"),
        calc_on_save: boolean("calcOnSave"),
        concurrent_calc: boolean("concurrentCalc"),
        concurrent_manual_count: number("concurrentManualCount"),
        force_full_calc: boolean("forceFullCalc"),
    }
}

/// The workbook's active tab, as an index into the sheets.
pub fn read_workbook_settings(xml_source: &[u8]) -> Result<Option<usize>> {
    let root = fromstring(xml_source)?;
    let Some(view) = root.find(format!("*/{{{SHEET_MAIN_NS}}}workbookView")) else {
        return Ok(None);
    };
    let Some(active) = view.get("activeTab") else {
        return Ok(None);
    };
    Ok(active.trim().parse::<usize>().ok())
}

/// A `(PartName, ContentType)` pair from `[Content_Types].xml`.
pub type ContentTypeEntry = (String, String);

/// Read the content-type overrides.
pub fn read_content_types(xml_source: &[u8]) -> Result<Vec<ContentTypeEntry>> {
    let root = fromstring(xml_source)?;
    Ok(root
        .find_all(format!("{{{CONTYPES_NS}}}Override"))
        .into_iter()
        .filter_map(|node| {
            Some((
                node.get("PartName")?.to_string(),
                node.get("ContentType")?.to_string(),
            ))
        })
        .collect())
}

/// Read the workbook relationships as `rId → Target`.
pub fn read_rels(xml_source: &[u8]) -> Result<HashMap<usize, String>> {
    let root = fromstring(xml_source)?;
    let mut rels = HashMap::new();
    for node in root.find_all(format!("{{{PKG_REL_NS}}}Relationship")) {
        let Some(id) = node
            .get("Id")
            .and_then(|v| v.strip_prefix("rId"))
            .and_then(|v| v.parse::<usize>().ok())
        else {
            continue;
        };
        if let Some(target) = node.get("Target") {
            rels.insert(id, target.to_string());
        }
    }
    Ok(rels)
}

/// Read `(sheet title, rId)` pairs from the workbook part.
pub fn read_sheets(xml_source: &[u8]) -> Result<Vec<(String, usize)>> {
    let root = fromstring(xml_source)?;
    // The `<sheet>` entries are nested inside `<sheets>`.
    let sheets_node = root.find(format!("{{{SHEET_MAIN_NS}}}sheets"));
    let mut sheets = Vec::new();
    let Some(sheets_node) = sheets_node else {
        return Ok(sheets);
    };
    for node in sheets_node.find_all(format!("{{{SHEET_MAIN_NS}}}sheet")) {
        let title = node.get("name").unwrap_or_default().to_string();
        let r_id = node
            .get(format!("{{{REL_NS}}}id"))
            .and_then(|v| v.strip_prefix("rId"))
            .and_then(|v| v.parse::<usize>().ok());
        if let Some(r_id) = r_id {
            sheets.push((title, r_id));
        }
    }
    Ok(sheets)
}

/// A worksheet discovered in the archive.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DetectedSheet {
    /// The sheet title.
    pub title: String,
    /// The path of the sheet part relative to `xl/`.
    pub path: String,
}

/// Resolve the worksheet parts by cross-referencing relationships and content types.
pub fn detect_worksheets(
    content_types: &[ContentTypeEntry],
    rels: &HashMap<usize, String>,
    sheets: &[(String, usize)],
) -> Vec<DetectedSheet> {
    let mut titles: HashMap<usize, String> = HashMap::new();
    for (title, r_id) in sheets {
        titles.insert(*r_id, title.clone());
    }
    let mut ids: Vec<usize> = rels.keys().copied().collect();
    ids.sort_unstable();
    let mut out = Vec::new();
    for r_id in ids {
        let path = &rels[&r_id];
        let full_path = format!("xl/{path}");
        let Some((_, content_type)) = content_types
            .iter()
            .find(|(part, _)| part.trim_start_matches('/') == full_path)
        else {
            continue;
        };
        if content_type != VALID_WORKSHEET {
            continue;
        }
        out.push(DetectedSheet {
            title: titles.get(&r_id).cloned().unwrap_or_default(),
            path: path.clone(),
        });
    }
    out
}

/// Read the defined names, dropping hidden, discarded and malformed entries.
///
/// `resolve` maps a sheet title to its index; unresolved titles are dropped from a range's
/// destinations, matching the Python behaviour of skipping unknown sheets.
pub fn read_named_ranges(
    xml_source: &[u8],
    resolve: &dyn Fn(&str) -> Option<usize>,
) -> Result<Vec<DefinedName>> {
    let root = fromstring(xml_source)?;
    let mut out = Vec::new();
    let Some(names) = root.find(format!("{{{SHEET_MAIN_NS}}}definedNames")) else {
        return Ok(out);
    };
    for node in names.children() {
        let range_name = node.get("name").unwrap_or_default().to_string();
        let node_text = node.text.clone().unwrap_or_default();
        if node.get("hidden") == Some("1") {
            continue;
        }
        if DISCARDED_RANGES.iter().any(|d| range_name.contains(d)) {
            continue;
        }
        if BUGGY_NAMED_RANGES.iter().any(|b| node_text.contains(b)) {
            continue;
        }
        let scope = node
            .get("localSheetId")
            .and_then(|v| v.trim().parse::<usize>().ok());
        if refers_to_range(&node_text) {
            let destinations = split_named_range(&node_text)?
                .into_iter()
                .filter_map(|(title, range)| resolve(&title).map(|index| (index, range)))
                .collect();
            out.push(DefinedName::Range(NamedRange::new(
                range_name,
                destinations,
                scope,
            )));
        } else {
            out.push(DefinedName::Value(NamedRangeContainingValue::new(
                range_name, node_text, scope,
            )));
        }
    }
    Ok(out)
}

/// Resolve a sheet title to its index from a list of titles.
pub fn title_resolver(titles: &[String]) -> impl Fn(&str) -> Option<usize> + '_ {
    move |title: &str| titles.iter().position(|t| t == title)
}

/// Helper used by [`read_properties_core`] for callers with an already-parsed tree.
pub fn properties_from_element(root: &Element) -> Result<DocumentProperties> {
    read_properties_core(&crate::xml::functions::serialize(root))
}
