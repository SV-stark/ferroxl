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

/// The directory the workbook's relationships are resolved against.
pub const WORKBOOK_PART_DIR: &str = "xl";

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
    /// The sheet part's path in the archive, already resolved against `xl/`.
    ///
    /// Prefixed again by a caller this becomes `xl/xl/worksheets/sheet1.xml` and the sheet
    /// disappears, so treat it as the final name rather than as a still-relative target.
    pub path: String,
}

/// Resolve a relationship target to a part name inside the package.
///
/// A `Target` is relative to the part holding the `.rels` file -- `xl/` for the workbook -- and
/// generators disagree about how to say so. All of these name `xl/worksheets/sheet1.xml`:
///
/// - `worksheets/sheet1.xml` (the usual form Excel writes)
/// - `/xl/worksheets/sheet1.xml` (absolute from the package root)
/// - `../xl/worksheets/sheet1.xml` (relative, with a hop back out of `xl/` first)
/// - `/xl/./worksheets/../worksheets/sheet1.xml` (absolute, with detours)
///
/// Assuming the first form and nothing else is how a sheet goes missing: the lookup against
/// `[Content_Types].xml` fails, `detect_worksheets` returns nothing, and the workbook opens
/// with no sheets at all. No error, no warning -- an empty workbook.
pub fn resolve_part(base_dir: &str, target: &str) -> String {
    // A leading slash makes the target absolute from the package root, so `base_dir` no longer
    // applies; an empty `base_dir` leaves nothing to prepend. Both mean "use the target alone".
    // Splitting on '/' and dropping empty segments below then handles the leading slash, any
    // `//`, and any `.` along the way.
    let joined = if target.starts_with('/') || base_dir.is_empty() {
        target.to_string()
    } else {
        format!("{base_dir}/{target}")
    };

    // Resolve `.` and `..` segment by segment. Doing this with string surgery rather than a
    // canonicalise crate keeps the dependency list short and the failure mode obvious: a `..`
    // that escapes the root is dropped, matching how a zip reader treats such a path.
    let mut parts: Vec<&str> = Vec::new();
    for segment in joined.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
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
        let target = &rels[&r_id];
        let full_path = resolve_part(WORKBOOK_PART_DIR, target);
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
            // The resolved path, not the raw target: the caller looks the part up in the zip,
            // and the zip is keyed by the resolved name.
            path: full_path,
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

#[cfg(test)]
mod target_tests {
    use super::*;
    use std::collections::HashMap;

    fn content_types() -> Vec<ContentTypeEntry> {
        vec![
            ("/xl/workbook.xml".to_string(), "workbook".to_string()),
            (
                "/xl/worksheets/sheet1.xml".to_string(),
                VALID_WORKSHEET.to_string(),
            ),
            (
                "/xl/worksheets/sheet2.xml".to_string(),
                VALID_WORKSHEET.to_string(),
            ),
        ]
    }

    fn rels(targets: &[&str]) -> HashMap<usize, String> {
        targets
            .iter()
            .enumerate()
            .map(|(index, target)| (index + 1, target.to_string()))
            .collect()
    }

    fn titles() -> Vec<(String, usize)> {
        vec![("One".to_string(), 1), ("Two".to_string(), 2)]
    }

    #[test]
    fn every_spelling_of_the_same_part_resolves_to_one_name() {
        // The whole point: a generator that writes any of these must still be detected, or the
        // workbook opens with no sheets and nothing says why.
        let spellings = [
            "worksheets/sheet1.xml",
            "/xl/worksheets/sheet1.xml",
            "../xl/worksheets/sheet1.xml",
            "/xl/./worksheets/../worksheets/sheet1.xml",
            "./worksheets/sheet1.xml",
            "worksheets//sheet1.xml",
        ];
        for spelling in spellings {
            assert_eq!(
                resolve_part(WORKBOOK_PART_DIR, spelling),
                "xl/worksheets/sheet1.xml",
                "{spelling}"
            );
        }
    }

    #[test]
    fn an_absolute_target_does_not_become_a_double_prefix() {
        let found = detect_worksheets(
            &content_types(),
            &rels(&["/xl/worksheets/sheet1.xml", "/xl/worksheets/sheet2.xml"]),
            &titles(),
        );
        assert_eq!(found.len(), 2, "both absolute targets were found");
        assert_eq!(found[0].path, "xl/worksheets/sheet1.xml");
        assert_eq!(found[0].title, "One");
        assert_eq!(found[1].path, "xl/worksheets/sheet2.xml");
        assert_eq!(found[1].title, "Two");
    }

    #[test]
    fn the_usual_relative_form_still_works() {
        let found = detect_worksheets(
            &content_types(),
            &rels(&["worksheets/sheet1.xml", "worksheets/sheet2.xml"]),
            &titles(),
        );
        assert_eq!(found.len(), 2);
        assert_eq!(found[0].path, "xl/worksheets/sheet1.xml");
    }

    #[test]
    fn a_target_that_climbs_out_of_xl_lands_at_the_package_root() {
        // `../customXml/item1.xml` from `xl/` is `customXml/item1.xml`, not `xl/../customXml/...`
        // and certainly not `xl/customXml/...`.
        assert_eq!(
            resolve_part(WORKBOOK_PART_DIR, "../customXml/item1.xml"),
            "customXml/item1.xml"
        );
    }

    #[test]
    fn a_target_escaping_the_package_root_loses_the_escape_rather_than_the_path() {
        // A `..` past the root is meaningless. Dropping it keeps the part findable, which beats
        // rejecting the file outright for a path no writer should have produced.
        assert_eq!(
            resolve_part(WORKBOOK_PART_DIR, "../../xl/worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn a_target_with_no_extension_such_as_a_directory_still_resolves() {
        assert_eq!(resolve_part(WORKBOOK_PART_DIR, "../docProps"), "docProps");
    }

    #[test]
    fn a_non_worksheet_relationship_is_still_excluded() {
        let types = vec![
            (
                "/xl/worksheets/sheet1.xml".to_string(),
                VALID_WORKSHEET.to_string(),
            ),
            ("/xl/styles.xml".to_string(), "styles".to_string()),
        ];
        let found = detect_worksheets(
            &types,
            &rels(&["worksheets/sheet1.xml", "styles.xml"]),
            &titles(),
        );
        assert_eq!(found.len(), 1, "styles.xml is not a sheet");
        assert_eq!(found[0].title, "One");
    }

    #[test]
    fn a_target_naming_no_part_at_all_is_skipped_rather_than_failing() {
        let found = detect_worksheets(
            &content_types(),
            &rels(&["worksheets/missing.xml"]),
            &titles(),
        );
        assert!(found.is_empty());
    }
}
