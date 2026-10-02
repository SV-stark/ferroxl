//! Capturing the parts of a package that the writer does not model.
//!
//! Run once at load time, over the archive as it stands. It reads every part, keeps the ones the
//! writer will not emit, and notes the content types and relationships that make those parts
//! reachable. See [`crate::workbook::preserved`] for why all three are needed together.

use crate::workbook::preserved::{PreservedContentType, PreservedParts, PreservedRelationship};
use crate::xml::functions::fromstring;
use std::io::{Cursor, Read};

/// `[Content_Types].xml`.
const CONTENT_TYPES: &str = "[Content_Types].xml";
/// `xl/workbook.xml`.
const WORKBOOK: &str = "xl/workbook.xml";

/// Paths the writer always produces, so a copy is never also preserved.
const ALWAYS_WRITER_OWNED: [&str; 8] = [
    CONTENT_TYPES,
    "_rels/.rels",
    "xl/_rels/workbook.xml.rels",
    WORKBOOK,
    "xl/styles.xml",
    "xl/sharedStrings.xml",
    "xl/theme/theme1.xml",
    "docProps/app.xml",
];

/// Whether `path` is `<directory>/<noun><index>.<extension>`, ignoring a `_rels/` step.
///
/// The `_rels` case matters: `xl/worksheets/_rels/sheet1.xml.rels` is the same part's
/// relationships, written by the writer, and without stripping that directory it looks like a
/// filename the writer has never heard of and gets preserved as an orphan.
fn is_indexed(path: &str, directory: &str, noun: &str) -> bool {
    let Some(rest) = path.strip_prefix(directory) else {
        return false;
    };
    let rest = rest.strip_prefix("_rels/").unwrap_or(rest);
    let Some((stem, extension)) = rest.rsplit_once('.') else {
        return false;
    };
    let stem = stem.strip_suffix(".xml").unwrap_or(stem);
    let digits = stem
        .strip_prefix(noun)
        .filter(|rest| !rest.is_empty() && rest.chars().all(|c| c.is_ascii_digit()));
    digits.is_some() && matches!(extension, "xml" | "vml" | "rels")
}

/// Whether the writer produces `path`, and so must not also have it preserved.
///
/// The rules mirror `ExcelWriter`. Paths it names by index -- worksheets, tables, charts,
/// drawings, comments -- are recognised by shape rather than by a list that would drift as the
/// writer gains parts. Everything else, including `vmlDrawing`, `pivotTables`, `slicers`,
/// `threadedComments`, `ctrlProps`, `queryTables` and `customXml`, is preserved.
pub fn is_writer_owned(path: &str) -> bool {
    if ALWAYS_WRITER_OWNED.contains(&path) || path == "docProps/core.xml" {
        return true;
    }
    if path.starts_with("xl/media/") {
        return true;
    }
    for (directory, noun) in [
        ("xl/worksheets/", "sheet"),
        ("xl/tables/", "table"),
        ("xl/charts/", "chart"),
        ("xl/drawings/", "drawing"),
        ("xl/", "comments"),
    ] {
        if is_indexed(path, directory, noun) {
            return true;
        }
    }
    // A chart's colour and style parts sit beside it under their own names, and a comment's VML
    // sits beside the drawings.
    if path.starts_with("xl/charts/")
        && (path.contains("colors") || path.contains("style") || path.contains("userShapes"))
    {
        return true;
    }
    is_indexed(path, "xl/drawings/", "commentsDrawing")
}

/// Read every part of `bytes` and keep what the writer will not produce.
///
/// `worksheet_paths` are the parts backing the sheets, in workbook order, so preserved worksheet
/// children are recorded against the sheet index that will find them again.
pub fn capture(bytes: &[u8], worksheet_paths: &[String]) -> PreservedParts {
    let mut preserved = PreservedParts::default();
    let Ok(mut archive) = zip::ZipArchive::new(Cursor::new(bytes.to_vec())) else {
        return preserved;
    };

    let names: Vec<String> = (0..archive.len())
        .filter_map(|index| {
            archive
                .by_index(index)
                .ok()
                .map(|entry| entry.name().to_string())
        })
        .collect();

    // The content types first: they say what the other parts are, and a source missing an entry
    // is exactly the case `PreservedParts::untyped_part_paths` exists to catch.
    if let Some(data) = read(&mut archive, CONTENT_TYPES) {
        for entry in parse_content_types(&data) {
            preserved.add_content_type(entry);
        }
    }

    for name in &names {
        if name == CONTENT_TYPES {
            continue;
        }
        let Some(data) = read(&mut archive, name) else {
            continue;
        };

        if name == WORKBOOK {
            if let Ok(root) = fromstring(&data) {
                preserved.capture_workbook_children(&root);
            }
            continue;
        }
        if let Some(at) = worksheet_paths.iter().position(|path| path == name) {
            if let Ok(root) = fromstring(&data) {
                preserved.capture_worksheet_children(at, &root);
            }
            continue;
        }
        if is_writer_owned(name) {
            continue;
        }
        if name.ends_with(".rels") {
            if let Some(relationships) = parse_relationships(&data) {
                preserved.add_relationships(name.clone(), relationships);
            }
            continue;
        }
        if name == "xl/vbaProject.bin" {
            // Recorded so the output keeps a macro-enabled content type. The bytes travel with
            // `vba_archive`, which already exists for exactly this.
            preserved.set_macro_enabled(Some("xlsm".to_string()));
            continue;
        }
        preserved.add_part(name.clone(), data);
    }

    preserved
}

fn read(archive: &mut zip::ZipArchive<Cursor<Vec<u8>>>, name: &str) -> Option<Vec<u8>> {
    let mut entry = archive.by_name(name).ok()?;
    let mut data = Vec::with_capacity(entry.size() as usize);
    entry.read_to_end(&mut data).ok()?;
    Some(data)
}

/// Every `Default` and `Override` in a `[Content_Types].xml`.
pub fn parse_content_types(data: &[u8]) -> Vec<PreservedContentType> {
    let Ok(root) = fromstring(data) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    for child in root.children() {
        match local_name(&child.tag) {
            Some("Default") => {
                let (Some(extension), Some(content_type)) =
                    (child.get("Extension"), child.get("ContentType"))
                else {
                    continue;
                };
                out.push(PreservedContentType {
                    kind: "Default".to_string(),
                    name: extension.to_string(),
                    content_type: content_type.to_string(),
                });
            }
            Some("Override") => {
                let (Some(part), Some(content_type)) =
                    (child.get("PartName"), child.get("ContentType"))
                else {
                    continue;
                };
                out.push(PreservedContentType {
                    kind: "Override".to_string(),
                    name: part.to_string(),
                    content_type: content_type.to_string(),
                });
            }
            _ => {}
        }
    }
    out
}

/// Every relationship in a `.rels` part.
pub fn parse_relationships(data: &[u8]) -> Option<Vec<PreservedRelationship>> {
    let root = fromstring(data).ok()?;
    let mut out = Vec::new();
    for child in root.children() {
        if local_name(&child.tag) != Some("Relationship") {
            continue;
        }
        let (Some(id), Some(kind), Some(target)) =
            (child.get("Id"), child.get("Type"), child.get("Target"))
        else {
            continue;
        };
        out.push(PreservedRelationship {
            id: id.to_string(),
            relationship_type: kind.to_string(),
            target: target.to_string(),
            target_mode: child.get("TargetMode").map(|mode| mode.to_string()),
        });
    }
    Some(out)
}

/// The sheet's `.rels` path, from its part path.
pub fn rels_path_for(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// The part a `.rels` path belongs to.
pub fn part_for_rels(rels: &str) -> String {
    let Some((directory, name)) = rels.rsplit_once("/_rels/") else {
        return rels.trim_end_matches(".rels").to_string();
    };
    let name = name.trim_end_matches(".rels");
    if directory.is_empty() {
        name.to_string()
    } else {
        format!("{directory}/{name}")
    }
}

fn local_name(tag: &str) -> Option<&str> {
    tag.rsplit_once('}').map(|(_, local)| local)
}

#[cfg(test)]
mod tests {
    use super::*;

    const CONTENT_TYPES_XML: &str = r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>"#;

    const RELS_XML: &str = r#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
                Target="../pivotCache/pivotCacheDefinition1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                Target="https://example.com" TargetMode="External"/>
</Relationships>"#;

    #[test]
    fn both_content_type_spellings_are_read() {
        let entries = parse_content_types(CONTENT_TYPES_XML.as_bytes());
        assert_eq!(entries.len(), 3);
        assert_eq!(entries[0].kind, "Default");
        assert_eq!(entries[0].name, "rels");
        assert_eq!(entries[1].name, "bin");
        assert_eq!(entries[2].kind, "Override");
        assert_eq!(entries[2].name, "/xl/pivotTables/pivotTable1.xml");
    }

    #[test]
    fn relationships_are_read_with_and_without_a_target_mode() {
        let relationships = parse_relationships(RELS_XML.as_bytes()).expect("parses");
        assert_eq!(relationships.len(), 2);
        assert_eq!(relationships[0].id, "rId1");
        assert_eq!(
            relationships[0].target,
            "../pivotCache/pivotCacheDefinition1.xml"
        );
        assert_eq!(relationships[0].target_mode, None);
        assert_eq!(relationships[1].target_mode.as_deref(), Some("External"));
    }

    #[test]
    fn a_rels_path_round_trips_through_its_part() {
        for part in [
            "xl/workbook.xml",
            "xl/worksheets/sheet1.xml",
            "xl/drawings/drawing1.xml",
        ] {
            assert_eq!(part_for_rels(&rels_path_for(part)), part);
        }
    }

    #[test]
    fn an_indexed_part_is_recognised_by_shape_not_by_list() {
        // A workbook with nine sheets must not need nine entries here.
        for path in [
            "xl/worksheets/sheet12.xml",
            "xl/worksheets/_rels/sheet12.xml.rels",
            "xl/tables/table3.xml",
            "xl/charts/chart7.xml",
            "xl/drawings/drawing2.xml",
            "xl/drawings/_rels/drawing2.xml.rels",
            "xl/drawings/commentsDrawing4.vml",
            "xl/comments3.xml",
            "xl/media/image1.png",
        ] {
            assert!(is_writer_owned(path), "{path} should be the writer's");
        }
    }

    #[test]
    fn an_unmodelled_part_is_not_mistaken_for_a_writer_owned_one() {
        for path in [
            "xl/pivotTables/pivotTable1.xml",
            "xl/pivotCache/pivotCacheDefinition1.xml",
            "xl/slicers/slicer1.xml",
            "xl/slicerCaches/slicerCache1.xml",
            "xl/threadedComments/threadedComment1.xml",
            "xl/queryTables/queryTable1.xml",
            "xl/connections.xml",
            "customXml/item1.xml",
            "xl/ctrlProps/ctrlProp1.xml",
            // A legacy drawing is *not* the writer's: it belongs to an ActiveX control, and
            // preserving it is the only thing keeping that control alive.
            "xl/drawings/vmlDrawing1.vml",
        ] {
            assert!(!is_writer_owned(path), "{path} should be preserved");
        }
    }

    #[test]
    fn a_part_merely_starting_like_a_writer_owned_one_is_not_one() {
        for path in [
            "xl/charts/chart.xml",
            "xl/tables/tables.xml",
            "xl/worksheets/sheet.xml",
            "xl/comments.xml",
            "xl/drawings/drawing.xml",
        ] {
            assert!(
                !is_writer_owned(path),
                "{path} has no index, so it is preserved"
            );
        }
    }
}
