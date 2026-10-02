//! Writing the workbook and package-level parts (`openpyxl/writer/workbook.py`).

use std::collections::BTreeSet;

use crate::cell::utils::absolute_coordinate;
use crate::date_time::datetime_to_w3cdtf;
use crate::namedrange::DefinedName;
use crate::workbook::DocumentProperties;
use crate::workbook::Workbook;
use crate::xml::constants::*;
use crate::xml::functions::Element;

/// The static part of `[Content_Types].xml`.
///
/// The tuple is `(kind, part, content type)`, where `kind` is `Override` or `Default`.
pub const STATIC_CONTENT_TYPES: [(&str, &str, &str); 10] = [
    (
        "Override",
        ARC_THEME,
        "application/vnd.openxmlformats-officedocument.theme+xml",
    ),
    (
        "Override",
        ARC_STYLE,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
    ),
    (
        "Default",
        "rels",
        "application/vnd.openxmlformats-package.relationships+xml",
    ),
    ("Default", "xml", "application/xml"),
    ("Default", "png", "image/png"),
    (
        "Default",
        "vml",
        "application/vnd.openxmlformats-officedocument.vmlDrawing",
    ),
    (
        "Override",
        ARC_WORKBOOK,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    ),
    (
        "Override",
        ARC_APP,
        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    ),
    (
        "Override",
        ARC_CORE,
        "application/vnd.openxmlformats-package.core-properties+xml",
    ),
    (
        "Override",
        ARC_SHARED_STRINGS,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
    ),
];

/// An owned attribute.
///
/// Attribute arrays must be homogeneous, so literals and computed values are both widened
/// to `String` before being collected.
fn attr(key: impl Into<String>, value: impl Into<String>) -> (String, String) {
    (key.into(), value.into())
}

/// Build a leaf element holding text.
fn text_element(tag: impl Into<String>, text: impl Into<String>) -> Element {
    let mut node = Element::new(tag);
    node.set_text(text);
    node
}

/// Serialise `docProps/core.xml`.
pub fn write_properties_core(properties: &DocumentProperties) -> String {
    let mut root = Element::new(format!("{{{COREPROPS_NS}}}coreProperties"));
    root.append(text_element(
        format!("{{{DCORE_NS}}}creator"),
        properties.creator.clone(),
    ));
    root.append(text_element(
        format!("{{{COREPROPS_NS}}}lastModifiedBy"),
        properties.last_modified_by.clone(),
    ));
    let xsi_type = format!("{{{XSI_NS}}}type");
    let w3cdtf = format!("{DCTERMS_PREFIX}:W3CDTF");
    let mut created = Element::with_attributes(
        format!("{{{DCTERMS_NS}}}created"),
        [(xsi_type.clone(), w3cdtf.clone())],
    );
    created.set_text(datetime_to_w3cdtf(properties.created));
    root.append(created);
    let mut modified =
        Element::with_attributes(format!("{{{DCTERMS_NS}}}modified"), [(xsi_type, w3cdtf)]);
    modified.set_text(datetime_to_w3cdtf(properties.modified));
    root.append(modified);

    root.append(text_element(
        format!("{{{DCORE_NS}}}title"),
        properties.title.clone(),
    ));
    root.append(text_element(
        format!("{{{DCORE_NS}}}description"),
        properties.description.clone(),
    ));
    root.append(text_element(
        format!("{{{DCORE_NS}}}subject"),
        properties.subject.clone(),
    ));
    root.append(text_element(
        format!("{{{COREPROPS_NS}}}keywords"),
        properties.keywords.clone(),
    ));
    root.append(text_element(
        format!("{{{COREPROPS_NS}}}category"),
        properties.category.clone(),
    ));
    root.to_pretty_string()
}

/// Serialise `[Content_Types].xml`.
///
/// Parts already declared by a preserved VBA archive are not re-declared, and per-sheet
/// parts (worksheets, drawings, charts, comments) are added with running ids.
/// The content type of a table part.
///
/// Its own part type rather than a variant of the worksheet's, which is why a table needs an
/// Override in `[Content_Types].xml` even though it is not a package root.
const TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";

/// The relationship type of a table part.
pub const TABLE_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

/// Serialise `[Content_Types].xml`.
pub fn write_content_types(workbook: &Workbook) -> String {
    let mut seen_parts: BTreeSet<String> = BTreeSet::new();
    let mut seen_extensions: BTreeSet<String> = BTreeSet::new();
    if workbook.vba_archive.is_some() {
        // The VBA archive carries its own content types; record what it already covers.
        for (part, extension) in VBA_CONTENT_TYPE_PARTS {
            seen_parts.insert(part.to_string());
            seen_extensions.insert(extension.to_string());
        }
    }

    let mut root = Element::new(format!("{{{CONTYPES_NS}}}Types"));
    for (kind, name, content_type) in STATIC_CONTENT_TYPES {
        if kind == "Override" {
            if seen_parts.contains(&format!("/{name}")) {
                continue;
            }
            root.append(Element::with_attributes(
                format!("{{{CONTYPES_NS}}}Override"),
                [
                    attr("ContentType".to_string(), content_type.to_string()),
                    attr("PartName".to_string(), format!("/{name}")),
                ],
            ));
        } else if !seen_extensions.contains(name) {
            root.append(Element::with_attributes(
                format!("{{{CONTYPES_NS}}}Default"),
                [
                    attr("ContentType".to_string(), content_type.to_string()),
                    attr("Extension".to_string(), name.to_string()),
                ],
            ));
        }
    }

    let mut drawing_id = 1u32;
    let mut chart_id = 1u32;
    let mut comments_id = 1u32;
    for (index, sheet) in workbook.worksheets.iter().enumerate() {
        let part = format!("/xl/worksheets/sheet{}.xml", index + 1);
        if !seen_parts.contains(&part) {
            root.append(Element::with_attributes(
                format!("{{{CONTYPES_NS}}}Override"),
                [
                    attr("PartName".to_string(), part),
                    attr(
                        "ContentType".to_string(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                            .to_string(),
                    ),
                ],
            ));
        }
        for table in sheet.tables.iter() {
            let part = format!("/xl/tables/table{}.xml", table.id);
            if seen_parts.contains(&part) {
                continue;
            }
            root.append(Element::with_attributes(
                format!("{{{CONTYPES_NS}}}Override"),
                [
                    attr("PartName".to_string(), part),
                    attr("ContentType".to_string(), TABLE_CONTENT_TYPE.to_string()),
                ],
            ));
        }
        if !sheet.charts.is_empty() || !sheet.images.is_empty() {
            let part = format!("/xl/drawings/drawing{drawing_id}.xml");
            if !seen_parts.contains(&part) {
                root.append(Element::with_attributes(
                    format!("{{{CONTYPES_NS}}}Override"),
                    [
                        attr("PartName".to_string(), part),
                        attr(
                            "ContentType".to_string(),
                            "application/vnd.openxmlformats-officedocument.drawing+xml".to_string(),
                        ),
                    ],
                ));
            }
            drawing_id += 1;
            for chart in &sheet.charts {
                let part = format!("/xl/charts/chart{chart_id}.xml");
                if !seen_parts.contains(&part) {
                    root.append(Element::with_attributes(
                        format!("{{{CONTYPES_NS}}}Override"),
                        [
                            attr("PartName".to_string(), part),
                            attr(
                                "ContentType".to_string(),
                                "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
                                    .to_string(),
                            ),
                        ],
                    ));
                }
                chart_id += 1;
                if !chart.shapes.is_empty() {
                    let part = format!("/xl/drawings/drawing{drawing_id}.xml");
                    if !seen_parts.contains(&part) {
                        root.append(Element::with_attributes(
                            format!("{{{CONTYPES_NS}}}Override"),
                            [
                                attr("PartName".to_string(), part),
                                attr(
                                    "ContentType".to_string(),
                                    "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
                                        .to_string(),
                                ),
                            ],
                        ));
                    }
                    drawing_id += 1;
                }
            }
        }
        if sheet.comment_count() > 0 {
            root.append(Element::with_attributes(
                format!("{{{CONTYPES_NS}}}Override"),
                [
                    attr(
                        "PartName".to_string(),
                        format!("/xl/comments{comments_id}.xml"),
                    ),
                    attr(
                        "ContentType".to_string(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
                            .to_string(),
                    ),
                ],
            ));
            comments_id += 1;
        }
    }
    root.to_pretty_string()
}

/// Parts and extensions a VBA archive declares, so they are not duplicated.
const VBA_CONTENT_TYPE_PARTS: [(&str, &str); 8] = [
    ("/xl/workbook.xml", "xml"),
    ("/xl/styles.xml", "xml"),
    ("/xl/sharedStrings.xml", "xml"),
    ("/xl/theme/theme1.xml", "xml"),
    ("/docProps/core.xml", "xml"),
    ("/docProps/app.xml", "xml"),
    ("xl/media/image", "png"),
    ("xl/drawings/vmlDrawing", "vml"),
];

/// Serialise `docProps/app.xml`.
pub fn write_properties_app(workbook: &Workbook) -> String {
    let count = workbook.worksheets.len();
    let mut root = Element::new(format!("{{{XPROPS_NS}}}Properties"));
    root.append(text_element(
        format!("{{{XPROPS_NS}}}Application"),
        "Microsoft Excel",
    ));
    root.append(text_element(format!("{{{XPROPS_NS}}}DocSecurity"), "0"));
    root.append(text_element(format!("{{{XPROPS_NS}}}ScaleCrop"), "false"));
    root.append(Element::new(format!("{{{XPROPS_NS}}}Company")));
    root.append(text_element(
        format!("{{{XPROPS_NS}}}LinksUpToDate"),
        "false",
    ));
    root.append(text_element(format!("{{{XPROPS_NS}}}SharedDoc"), "false"));
    root.append(text_element(
        format!("{{{XPROPS_NS}}}HyperlinksChanged"),
        "false",
    ));
    root.append(text_element(
        format!("{{{XPROPS_NS}}}AppVersion"),
        "12.0000",
    ));

    let mut heading_pairs = Element::new(format!("{{{XPROPS_NS}}}HeadingPairs"));
    let mut vector = Element::with_attributes(
        format!("{{{VTYPES_NS}}}vector"),
        [("size", "2"), ("baseType", "variant")],
    );
    vector.append(text_element(format!("{{{VTYPES_NS}}}lpstr"), "Worksheets"));
    let mut variant = Element::new(format!("{{{VTYPES_NS}}}variant"));
    variant.append(text_element(
        format!("{{{VTYPES_NS}}}i4"),
        count.to_string(),
    ));
    vector.append(variant);
    heading_pairs.append(vector);
    root.append(heading_pairs);

    let mut title_of_parts = Element::new(format!("{{{XPROPS_NS}}}TitlesOfParts"));
    let mut vector = Element::with_attributes(
        format!("{{{VTYPES_NS}}}vector"),
        [attr("size", count.to_string()), attr("baseType", "lpstr")],
    );
    for sheet in &workbook.worksheets {
        vector.append(text_element(
            format!("{{{VTYPES_NS}}}lpstr"),
            sheet.title().to_string(),
        ));
    }
    title_of_parts.append(vector);
    root.append(title_of_parts);
    root.to_pretty_string()
}

/// Serialise `_rels/.rels`.
///
/// The package-level relationships are the same for every workbook — the office document,
/// the core properties and the extended properties — so nothing here depends on
/// `workbook`. The parameter is kept so the call site reads like the other `write_*`
/// functions.
pub fn write_root_rels(_workbook: &Workbook) -> String {
    let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
    for (id, target, rel_type) in [
        ("rId1", ARC_WORKBOOK, format!("{REL_NS}/officeDocument")),
        (
            "rId2",
            ARC_CORE,
            format!("{PKG_REL_NS}/metadata/core-properties"),
        ),
        ("rId3", ARC_APP, format!("{REL_NS}/extended-properties")),
    ] {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                attr("Id", id),
                attr("Target", target),
                attr("Type", rel_type),
            ],
        ));
    }
    root.to_pretty_string()
}

/// Serialise `xl/workbook.xml`.
pub fn write_workbook(workbook: &Workbook) -> String {
    let mut root = Element::new(format!("{{{SHEET_MAIN_NS}}}workbook"));
    root.append(Element::with_attributes(
        format!("{{{SHEET_MAIN_NS}}}fileVersion"),
        [
            ("appName", "xl"),
            ("lastEdited", "4"),
            ("lowestEdited", "4"),
            ("rupBuild", "4505"),
        ],
    ));
    let mut properties = Element::with_attributes(
        format!("{{{SHEET_MAIN_NS}}}workbookPr"),
        [
            ("defaultThemeVersion", "124226"),
            ("codeName", "ThisWorkbook"),
        ],
    );
    if workbook.properties.excel_base_date == crate::date_time::BaseDate::Mac1904 {
        properties.set("date1904", "1");
    }
    root.append(properties);

    let mut book_views = Element::new(format!("{{{SHEET_MAIN_NS}}}bookViews"));
    book_views.append(Element::with_attributes(
        format!("{{{SHEET_MAIN_NS}}}workbookView"),
        [
            attr("activeTab", workbook.active().to_string()),
            attr("autoFilterDateGrouping", "1"),
            attr("firstSheet", "0"),
            attr("minimized", "0"),
            attr("showHorizontalScroll", "1"),
            attr("showSheetTabs", "1"),
            attr("showVerticalScroll", "1"),
            attr("tabRatio", "600"),
            attr("visibility", "visible"),
        ],
    ));
    root.append(book_views);

    let mut sheets = Element::new(format!("{{{SHEET_MAIN_NS}}}sheets"));
    for (index, sheet) in workbook.worksheets.iter().enumerate() {
        let mut node = Element::with_attributes(
            format!("{{{SHEET_MAIN_NS}}}sheet"),
            [
                attr("name".to_string(), sheet.title().to_string()),
                attr("sheetId".to_string(), (index + 1).to_string()),
                attr(format!("{{{REL_NS}}}id"), format!("rId{}", index + 1)),
            ],
        );
        if sheet.sheet_state != crate::worksheet::Worksheet::SHEETSTATE_VISIBLE {
            node.set("state", sheet.sheet_state.clone());
        }
        sheets.append(node);
    }
    root.append(sheets);

    let mut defined_names = Element::new(format!("{{{SHEET_MAIN_NS}}}definedNames"));
    let titles = workbook.get_sheet_names();
    for entry in workbook.get_named_ranges() {
        let mut node = Element::with_attributes(
            format!("{{{SHEET_MAIN_NS}}}definedName"),
            [attr("name".to_string(), entry.name().to_string())],
        );
        if let Some(scope) = entry.scope() {
            node.set("localSheetId", scope.to_string());
        }
        let body = match entry {
            DefinedName::Range(range) => range
                .destinations
                .iter()
                .map(|(index, cell_range)| {
                    let title = titles.get(*index).cloned().unwrap_or_default();
                    format!(
                        "'{}'!{}",
                        title.replace('\'', "''"),
                        absolute_coordinate(cell_range)
                    )
                })
                .collect::<Vec<String>>()
                .join(","),
            DefinedName::Value(value) => value.value.clone(),
        };
        node.set_text(body);
        defined_names.append(node);
    }
    // Excel stores the autofilter range as a hidden, sheet-scoped defined name.
    for (index, sheet) in workbook.worksheets.iter().enumerate() {
        let Some(reference) = sheet.auto_filter.reference() else {
            continue;
        };
        let mut node = Element::with_attributes(
            format!("{{{SHEET_MAIN_NS}}}definedName"),
            [
                attr("name".to_string(), "_xlnm._FilterDatabase".to_string()),
                attr("localSheetId".to_string(), index.to_string()),
                attr("hidden".to_string(), "1".to_string()),
            ],
        );
        node.set_text(format!(
            "'{}'!{}",
            sheet.title().replace('\'', "''"),
            absolute_coordinate(reference)
        ));
        defined_names.append(node);
    }
    root.append(defined_names);

    // The calculation properties were three fixed attributes. `fullCalcOnLoad` is the one that
    // matters: ferroxl writes formulas with no cached result, so this is what makes Excel
    // calculate them on open rather than showing blanks.
    root.append(Element::with_attributes(
        format!("{{{SHEET_MAIN_NS}}}calcPr"),
        workbook.calculation.attributes(),
    ));
    root.to_pretty_string()
}

/// Serialise `xl/_rels/workbook.xml.rels`.
pub fn write_workbook_rels(workbook: &Workbook) -> String {
    let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
    for index in 1..=workbook.worksheets.len() {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                attr("Id".to_string(), format!("rId{index}")),
                attr("Target".to_string(), format!("worksheets/sheet{index}.xml")),
                attr("Type".to_string(), format!("{REL_NS}/worksheet")),
            ],
        ));
    }
    let base = workbook.worksheets.len() + 1;
    for (offset, (target, suffix)) in [
        ("sharedStrings.xml", "sharedStrings"),
        ("styles.xml", "styles"),
        ("theme/theme1.xml", "theme"),
    ]
    .iter()
    .enumerate()
    {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                attr("Id".to_string(), format!("rId{}", base + offset)),
                attr("Target".to_string(), (*target).to_string()),
                attr("Type".to_string(), format!("{REL_NS}/{suffix}")),
            ],
        ));
    }
    if workbook.vba_archive.is_some() {
        root.append(Element::with_attributes(
            format!("{{{PKG_REL_NS}}}Relationship"),
            [
                attr("Id".to_string(), format!("rId{}", base + 3)),
                attr("Target".to_string(), "vbaProject.bin".to_string()),
                attr(
                    "Type".to_string(),
                    "http://schemas.microsoft.com/office/2006/relationships/vbaProject".to_string(),
                ),
            ],
        ));
    }
    root.to_pretty_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::CellValue;
    use crate::namedrange::NamedRange;
    use crate::xml::functions::fromstring;

    fn two_sheet_workbook() -> Workbook {
        let mut workbook = Workbook::new();
        workbook.create_sheet(Some("Data")).unwrap();
        workbook.worksheets[0]
            .set("A1", CellValue::text("hello"))
            .unwrap();
        workbook
    }

    #[test]
    fn core_properties_are_written() {
        let mut workbook = Workbook::new();
        workbook.properties.creator = "eric".to_string();
        workbook.properties.title = "Report".to_string();
        let xml = write_properties_core(&workbook.properties);
        let root = fromstring(xml.as_bytes()).unwrap();
        assert_eq!(root.tag, format!("{{{COREPROPS_NS}}}coreProperties"));
        assert_eq!(root.find_text(format!("{{{DCORE_NS}}}creator"), ""), "eric");
        assert_eq!(root.find_text(format!("{{{DCORE_NS}}}title"), ""), "Report");
        assert!(xml.contains("W3CDTF"));
    }

    #[test]
    fn app_properties_count_sheets() {
        let workbook = two_sheet_workbook();
        let xml = write_properties_app(&workbook);
        let root = fromstring(xml.as_bytes()).unwrap();
        let vector = root
            .find(format!("{{{XPROPS_NS}}}TitlesOfParts"))
            .and_then(|el| el.find(format!("{{{VTYPES_NS}}}vector")))
            .expect("vector");
        assert_eq!(vector.get("size"), Some("2"));
        let titles: Vec<String> = vector
            .find_all(format!("{{{VTYPES_NS}}}lpstr"))
            .into_iter()
            .map(|node| node.text.clone().unwrap_or_default())
            .collect();
        assert_eq!(titles, vec!["Sheet1".to_string(), "Data".to_string()]);
    }

    #[test]
    fn content_types_cover_every_part() {
        let mut workbook = two_sheet_workbook();
        workbook.worksheets[0]
            .set_comment("A1", Some(crate::comments::Comment::new("hi", "me")))
            .unwrap();
        let xml = write_content_types(&workbook);
        let root = fromstring(xml.as_bytes()).unwrap();
        let parts: Vec<String> = root
            .find_all(format!("{{{CONTYPES_NS}}}Override"))
            .into_iter()
            .filter_map(|node| node.get("PartName").map(|v| v.to_string()))
            .collect();
        for expected in [
            "/xl/workbook.xml",
            "/xl/styles.xml",
            "/xl/sharedStrings.xml",
            "/xl/theme/theme1.xml",
            "/docProps/core.xml",
            "/docProps/app.xml",
            "/xl/worksheets/sheet1.xml",
            "/xl/worksheets/sheet2.xml",
            "/xl/comments1.xml",
        ] {
            assert!(parts.contains(&expected.to_string()), "missing {expected}");
        }
    }

    #[test]
    fn content_types_declare_image_and_vml_defaults() {
        let root = fromstring(write_content_types(&Workbook::new()).as_bytes()).unwrap();
        let extensions: Vec<String> = root
            .find_all(format!("{{{CONTYPES_NS}}}Default"))
            .into_iter()
            .filter_map(|node| node.get("Extension").map(|v| v.to_string()))
            .collect();
        assert!(extensions.contains(&"png".to_string()));
        assert!(extensions.contains(&"vml".to_string()));
        assert!(extensions.contains(&"rels".to_string()));
    }

    #[test]
    fn root_relationships_are_static() {
        let xml = write_root_rels(&Workbook::new());
        let root = fromstring(xml.as_bytes()).unwrap();
        let relationships = root.find_all(format!("{{{PKG_REL_NS}}}Relationship"));
        assert_eq!(relationships.len(), 3);
        assert_eq!(relationships[0].get("Target"), Some(ARC_WORKBOOK));
    }

    #[test]
    fn workbook_lists_sheets_with_relationship_ids() {
        let workbook = two_sheet_workbook();
        let root = fromstring(write_workbook(&workbook).as_bytes()).unwrap();
        let sheets = root
            .find(format!("{{{SHEET_MAIN_NS}}}sheets"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}sheet"));
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].get("name"), Some("Sheet1"));
        assert_eq!(sheets[0].get(format!("{{{REL_NS}}}id")), Some("rId1"));
        assert_eq!(sheets[1].get(format!("{{{REL_NS}}}id")), Some("rId2"));
    }

    #[test]
    fn hidden_sheets_carry_state() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].sheet_state = "hidden".to_string();
        let root = fromstring(write_workbook(&workbook).as_bytes()).unwrap();
        let sheet = root
            .find(format!("{{{SHEET_MAIN_NS}}}sheets"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}sheet"))[0]
            .clone();
        assert_eq!(sheet.get("state"), Some("hidden"));
    }

    #[test]
    fn mac_workbooks_record_date1904() {
        let mut workbook = Workbook::new();
        workbook.properties.excel_base_date = crate::date_time::BaseDate::Mac1904;
        let xml = write_workbook(&workbook);
        assert!(xml.contains("date1904=\"1\""));
    }

    #[test]
    fn named_ranges_are_absolutised_and_quoted() {
        let mut workbook = Workbook::new();
        workbook.create_sheet(Some("My Sheet")).unwrap();
        workbook.add_named_range(NamedRange::new(
            "MyRef",
            vec![(1, "a1:b3".to_string())],
            None,
        ));
        let root = fromstring(write_workbook(&workbook).as_bytes()).unwrap();
        let defined = root
            .find(format!("{{{SHEET_MAIN_NS}}}definedNames"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}definedName"));
        assert_eq!(defined[0].get("name"), Some("MyRef"));
        assert_eq!(defined[0].text.as_deref(), Some("'My Sheet'!$A$1:$B$3"));
    }

    #[test]
    fn autofilters_become_hidden_defined_names() {
        let mut workbook = Workbook::new();
        workbook.worksheets[0].auto_filter.set_reference("a1:c5");
        let root = fromstring(write_workbook(&workbook).as_bytes()).unwrap();
        let defined = root
            .find(format!("{{{SHEET_MAIN_NS}}}definedNames"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}definedName"));
        assert_eq!(defined[0].get("name"), Some("_xlnm._FilterDatabase"));
        assert_eq!(defined[0].get("hidden"), Some("1"));
        assert_eq!(defined[0].text.as_deref(), Some("'Sheet1'!$A$1:$C$5"));
    }

    #[test]
    fn value_names_are_written_verbatim() {
        let mut workbook = Workbook::new();
        workbook.add_named_value("MyValue", "9.99", Some(0));
        let root = fromstring(write_workbook(&workbook).as_bytes()).unwrap();
        let defined = root
            .find(format!("{{{SHEET_MAIN_NS}}}definedNames"))
            .unwrap()
            .find_all(format!("{{{SHEET_MAIN_NS}}}definedName"))[0]
            .clone();
        assert_eq!(defined.text.as_deref(), Some("9.99"));
        assert_eq!(defined.get("localSheetId"), Some("0"));
    }

    #[test]
    fn workbook_relationships_cover_every_part() {
        let workbook = two_sheet_workbook();
        let root = fromstring(write_workbook_rels(&workbook).as_bytes()).unwrap();
        let relationships = root.find_all(format!("{{{PKG_REL_NS}}}Relationship"));
        assert_eq!(relationships.len(), 5);
        assert_eq!(
            relationships[0].get("Target"),
            Some("worksheets/sheet1.xml")
        );
        assert_eq!(relationships[2].get("Target"), Some("sharedStrings.xml"));
        assert_eq!(relationships[3].get("Target"), Some("styles.xml"));
        assert_eq!(relationships[4].get("Target"), Some("theme/theme1.xml"));
    }

    #[test]
    fn vba_workbooks_declare_the_vba_project() {
        let mut workbook = Workbook::new();
        workbook.vba_archive = Some(Vec::new());
        let root = fromstring(write_workbook_rels(&workbook).as_bytes()).unwrap();
        let relationships = root.find_all(format!("{{{PKG_REL_NS}}}Relationship"));
        assert_eq!(relationships.len(), 5);
        assert_eq!(relationships[4].get("Target"), Some("vbaProject.bin"));
    }

    #[test]
    fn calc_pr_forces_recalculation_on_load() {
        let xml = write_workbook(&Workbook::new());
        assert!(xml.contains("fullCalcOnLoad=\"1\""));
    }
}
