//! Writing back the parts ferroxl does not model.
//!
//! Everything here merges into XML the writer has already produced rather than rebuilding it,
//! because the generated parts are the writer's contract with Excel and re-deriving them would
//! be a second implementation to keep in step.
//!
//! Merging is textual at three points, and each is deliberate:
//!
//! - `[Content_Types].xml` and the `.rels` parts are parsed, merged and re-serialised. That is
//!   cheap, and it means the merge works on the element model rather than on string matching.
//! - `<workbook>` and `<worksheet>` children are appended before the closing tag. That closing
//!   tag was written by this module a moment earlier, so finding it is not a guess.
//! - Relationship ids that collided are remapped, and the `r:id` in the preserved element that
//!   named them is rewritten to match. Skipping that step produces a file that opens and has
//!   quietly lost whatever the reference pointed at.

use crate::workbook::preserved::{
    merge_relationships, rewrite_relationship_ids, GeneratedRelationship, PreservedParts,
};
use crate::xml::functions::{fromstring, Element};
use std::collections::BTreeMap;
use std::io::Write;

/// The MIME type used for a preserved part that no content type declares.
///
/// `application/xml` is what the OPC default for `xml` already says, so reaching this means the
/// source declared the part by an extension we did not recognise. Guessing here beats dropping
/// the part, because a part with no declared type is a package Excel refuses outright.
const FALLBACK_XML_TYPE: &str = "application/xml";

/// The content types namespace, which differs from the relationships one.
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
/// The relationships namespace.
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn local_name(tag: &str) -> Option<&str> {
    tag.rsplit_once('}').map(|(_, local)| local)
}

/// Merge the preserved content types into generated `[Content_Types].xml`.
///
/// Returns `generated` unchanged when there is nothing to add, so the common case costs one
/// `is_empty` check.
pub fn merge_content_types(generated: &str, preserved: &PreservedParts) -> String {
    if preserved.content_types().is_empty() && preserved.untyped_part_paths().is_empty() {
        return generated.to_string();
    }
    let Ok(mut root) = fromstring(generated.as_bytes()) else {
        return generated.to_string();
    };

    let mut have_extensions: Vec<String> = Vec::new();
    let mut have_parts: Vec<String> = Vec::new();
    for child in root.children() {
        match local_name(&child.tag) {
            Some("Default") => {
                if let Some(extension) = child.get("Extension") {
                    have_extensions.push(extension.to_lowercase());
                }
            }
            Some("Override") => {
                if let Some(part) = child.get("PartName") {
                    have_parts.push(part.trim_start_matches('/').to_string());
                }
            }
            _ => {}
        }
    }

    // The source's own declarations first, so a package that declared `bin` as a VBA project
    // keeps saying so rather than falling back to a sniffed type.
    for entry in preserved.content_types() {
        if entry.kind == "Default" {
            if !have_extensions.contains(&entry.name.to_lowercase()) {
                have_extensions.push(entry.name.to_lowercase());
                root.append(content_type_element(
                    "Default",
                    "Extension",
                    entry.name.clone(),
                    entry.content_type.clone(),
                ));
            }
        } else {
            let part = entry.name.trim_start_matches('/').to_string();
            if !have_parts.contains(&part) {
                have_parts.push(part);
                root.append(content_type_element(
                    "Override",
                    "PartName",
                    entry.name.clone(),
                    entry.content_type.clone(),
                ));
            }
        }
    }

    // Then anything preserved that still has no declaration.
    for path in preserved.parts().keys() {
        if have_parts.iter().any(|part| part == path) {
            continue;
        }
        let extension = path.rsplit('.').next().unwrap_or_default().to_lowercase();
        if have_extensions.contains(&extension) {
            continue;
        }
        have_parts.push(path.clone());
        root.append(content_type_element(
            "Override",
            "PartName",
            format!("/{path}"),
            FALLBACK_XML_TYPE.to_string(),
        ));
    }

    root.to_pretty_string()
}

fn content_type_element(
    element: &str,
    key_attribute: &str,
    name: String,
    content_type: String,
) -> Element {
    Element::with_attributes(
        format!("{{{CONTENT_TYPES_NS}}}{element}"),
        [("ContentType", content_type), (key_attribute, name)],
    )
}

/// Merge preserved relationships into a generated `.rels` part.
///
/// Returns the merged XML and the map from source ids to written ids, so a preserved element
/// naming a relationship that moved can be corrected. An empty map means nothing moved and the
/// references are already right.
pub fn merge_rels(
    generated: &str,
    preserved: &PreservedParts,
    rels_path: &str,
) -> (String, BTreeMap<String, String>) {
    let Some(preserved_rels) = preserved.relationships(rels_path) else {
        return (generated.to_string(), BTreeMap::new());
    };
    let Ok(root) = fromstring(generated.as_bytes()) else {
        return (generated.to_string(), BTreeMap::new());
    };

    let mut own: Vec<GeneratedRelationship> = Vec::new();
    for child in root.children() {
        if local_name(&child.tag) != Some("Relationship") {
            continue;
        }
        let (Some(id), Some(kind), Some(target)) =
            (child.get("Id"), child.get("Type"), child.get("Target"))
        else {
            continue;
        };
        own.push(GeneratedRelationship {
            id: id.to_string(),
            relationship_type: kind.to_string(),
            target: target.to_string(),
            target_mode: child.get("TargetMode").map(|mode| mode.to_string()),
        });
    }

    let merged = merge_relationships(&own, preserved_rels);
    (rels_document(&merged.relationships), merged.id_map)
}

/// A `.rels` part holding only preserved relationships.
///
/// Used for a sheet the writer writes nothing else for: it skips the `.rels` part entirely when a
/// sheet has no charts, images, comments or tables, which is exactly the case for a loaded sheet
/// whose only relationship is to a pivot table. Without this the preserved relationship has
/// nowhere to live.
pub fn rels_from_preserved(preserved: &PreservedParts, rels_path: &str) -> Option<String> {
    let preserved_rels = preserved.relationships(rels_path)?;
    if preserved_rels.is_empty() {
        return None;
    }
    let merged = merge_relationships(&[], preserved_rels);
    Some(rels_document(&merged.relationships))
}

fn rels_document(relationships: &[GeneratedRelationship]) -> String {
    let mut root = Element::new(format!("{{{RELATIONSHIPS_NS}}}Relationships"));
    for entry in relationships {
        root.append(Element::with_attributes(
            format!("{{{RELATIONSHIPS_NS}}}Relationship"),
            entry.to_attributes(),
        ));
    }
    root.to_pretty_string()
}

/// The offset of the last closing tag whose local name is `element`, or `None`.
///
/// Matched by local name rather than by literal string because the two writers do not agree on
/// how to close an element. `write_worksheet` streams, so it emits `</worksheet>`;
/// `write_workbook` serialises an `Element` tree, which renders the same element as
/// `</ns0:workbook>`. A literal search finds one and silently misses the other, and a missed
/// match means the children are dropped without a word.
fn closing_tag_offset(xml: &str, element: &str) -> Option<usize> {
    let mut found = None;
    let mut at = 0usize;
    while let Some(offset) = xml[at..].find("</") {
        let start = at + offset;
        let end = xml[start..].find('>')? + start + 1;
        // A closing tag may carry a namespace prefix (`ns0:workbook`), a braced namespace
        // (`{ns}workbook`) or neither, depending on which writer produced the part.
        let name = &xml[start + 2..end - 1];
        let name = name
            .rsplit_once(':')
            .map(|(_, local)| local)
            .unwrap_or(name);
        let name = name
            .rsplit_once('}')
            .map(|(_, local)| local)
            .unwrap_or(name);
        if name == element {
            found = Some(start);
        }
        at = end;
    }
    found
}

/// Append preserved children to a generated part, before its closing tag.
///
/// `element` is the local name of the part's root, e.g. `worksheet` or `workbook`.
pub fn append_children(
    generated: &str,
    element: &str,
    children: &[Element],
    id_map: &BTreeMap<String, String>,
) -> String {
    if children.is_empty() {
        return generated.to_string();
    }
    let Some(at) = closing_tag_offset(generated, element) else {
        // No closing tag means something upstream changed the shape of the part. Returning the
        // generated XML unchanged loses the children, but writing them into an unexpected place
        // would produce a file Excel rejects -- and a rejected file is worse than a lossy one.
        return generated.to_string();
    };
    // Children are separated by newlines rather than terminated by one, so the output has no
    // trailing blank line before the closing tag.
    let rendered: Vec<String> = children
        .iter()
        .map(|child| {
            let mut owned = child.clone();
            rewrite_relationship_ids(&mut owned, id_map);
            owned.to_pretty_string()
        })
        .collect();
    format!(
        "{}{}{}",
        &generated[..at],
        rendered.join("\n"),
        &generated[at..]
    )
}

/// Write every preserved part that is not a `.rels` part.
///
/// `.rels` parts are merged into the ones the writer produces rather than written as they stand,
/// because a relationship id the writer has reused has to move.
pub fn write_parts(
    archive: &mut zip::ZipWriter<std::io::Cursor<Vec<u8>>>,
    preserved: &PreservedParts,
) -> crate::exceptions::Result<()> {
    for (path, data) in preserved.parts() {
        if path.ends_with(".rels") {
            continue;
        }
        write_entry(archive, path, data)?;
    }
    write_preserved_rels(archive, preserved)
}

/// Write the relationships of every preserved part that needs one.
///
/// A preserved part's `.rels` is parsed at load time rather than kept verbatim, so it has to be
/// rendered here. Only parts that are themselves preserved get one: the `.rels` the writer owns
/// are merged into its own output instead. Without this a preserved drawing is a set of anchors
/// pointing at nothing -- the part survives, and nothing says which chart or image it holds.
fn write_preserved_rels(
    archive: &mut zip::ZipWriter<std::io::Cursor<Vec<u8>>>,
    preserved: &PreservedParts,
) -> crate::exceptions::Result<()> {
    for rels_path in preserved.all_relationships().keys() {
        let owner = crate::reader::preserved::part_for_rels(rels_path);
        if preserved.part(&owner).is_none() {
            continue;
        }
        let Some(document) = rels_from_preserved(preserved, rels_path) else {
            continue;
        };
        write_entry(archive, rels_path, document.as_bytes())?;
    }
    Ok(())
}

fn write_entry(
    archive: &mut zip::ZipWriter<std::io::Cursor<Vec<u8>>>,
    path: &str,
    data: &[u8],
) -> crate::exceptions::Result<()> {
    use crate::exceptions::Error;
    archive
        .start_file(
            path,
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated),
        )
        .map_err(|error| Error::Io(error.to_string()))?;
    archive
        .write_all(data)
        .map_err(|error| Error::Io(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::workbook::preserved::{PreservedContentType, PreservedRelationship};

    /// A stand-in for what `write_content_types` produces.
    ///
    /// It includes the Overrides the real one has for `xl/styles.xml` and `xl/theme/theme1.xml`,
    /// because the tests below are about what happens when a source declares a part the writer
    /// has *already* declared -- which is only a conflict if the writer really declared it.
    const GENERATED_TYPES: &str = r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.ms-excel.styles"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.ms-theme"/>
</Types>"#;

    fn preserved_with_a_pivot_table() -> PreservedParts {
        let mut preserved = PreservedParts::default();
        preserved.add_content_type(PreservedContentType {
            kind: "Override".to_string(),
            name: "/xl/pivotTables/pivotTable1.xml".to_string(),
            content_type: "application/vnd.ms-pivotTable".to_string(),
        });
        preserved.add_part("xl/pivotTables/pivotTable1.xml", b"<p/>".to_vec());
        preserved
    }

    /// Whether a merged content-types document declares `part` with `content_type`.
    ///
    /// Namespace-agnostic on purpose: re-serialising the element model renames the default
    /// namespace to an `ns0:` prefix, which is semantically identical and is what the rest of
    /// the writer already emits. Asserting on the spelling would make this test fail for a
    /// reason that does not matter.
    fn declares(merged: &str, part: &str, content_type: &str) -> bool {
        let Ok(root) = fromstring(merged.as_bytes()) else {
            return false;
        };
        root.children().iter().any(|child| {
            let name = child
                .get("PartName")
                .or_else(|| child.get("Extension"))
                .unwrap_or_default();
            name.trim_start_matches('/') == part.trim_start_matches('/')
                && child.get("ContentType") == Some(content_type)
        })
    }

    #[test]
    fn a_preserved_part_gets_a_content_type_it_otherwise_would_not_have() {
        let merged = merge_content_types(GENERATED_TYPES, &preserved_with_a_pivot_table());
        assert!(
            declares(
                &merged,
                "/xl/pivotTables/pivotTable1.xml",
                "application/vnd.ms-pivotTable"
            ),
            "the pivot table is now declared: {merged}"
        );
    }

    #[test]
    fn the_generated_declarations_are_left_alone() {
        // An Override for a part the writer emits would be a second, conflicting declaration.
        let mut preserved = PreservedParts::default();
        preserved.add_content_type(PreservedContentType {
            kind: "Override".to_string(),
            name: "/xl/styles.xml".to_string(),
            content_type: "application/wrong".to_string(),
        });
        let merged = merge_content_types(GENERATED_TYPES, &preserved);
        assert!(
            !declares(&merged, "/xl/styles.xml", "application/wrong"),
            "a second, conflicting declaration was added: {merged}"
        );
    }

    #[test]
    fn a_part_with_no_declaration_anywhere_still_gets_one() {
        // `.dat` rather than `.xml`: the generated document already declares a `Default` for
        // `xml`, so an `.xml` part is covered and correctly left alone. The case that needs
        // handling is an extension nothing declares.
        let mut preserved = PreservedParts::default();
        preserved.add_part("xl/queryTables/cache1.dat", b"<c/>".to_vec());
        let merged = merge_content_types(GENERATED_TYPES, &preserved);
        assert!(
            declares(&merged, "/xl/queryTables/cache1.dat", FALLBACK_XML_TYPE),
            "a part with no declared type is a package Excel refuses: {merged}"
        );
    }

    #[test]
    fn a_colliding_relationship_id_is_moved_and_the_reference_rewritten() {
        let mut preserved = PreservedParts::default();
        preserved.add_relationships(
            "xl/_rels/workbook.xml.rels",
            vec![PreservedRelationship::internal(
                "rId1",
                "t/pivotCacheDefinition",
                "pivotCache/pivotCacheDefinition1.xml",
            )],
        );
        let generated = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="t/sheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#;
        let (merged, id_map) = merge_rels(generated, &preserved, "xl/_rels/workbook.xml.rels");

        assert_eq!(id_map.get("rId1").map(String::as_str), Some("rId2"));
        assert!(
            merged.contains("pivotCache/pivotCacheDefinition1.xml"),
            "{merged}"
        );
        assert_eq!(
            merged.matches("Id=\"rId1\"").count(),
            1,
            "one rId1, not two: {merged}"
        );
    }

    #[test]
    fn a_preserved_child_is_appended_before_a_closing_tag_with_a_prefix() {
        // `write_workbook` serialises an `Element` tree and closes as `</ns0:workbook>`;
        // `write_worksheet` streams and closes as `</worksheet>`. A literal search finds one and
        // silently misses the other, which is how the workbook's `<pivotCaches>` went missing
        // while the tests passed.
        let generated = "<ns0:workbook xmlns:ns0=\"urn:x\"><ns0:sheets/></ns0:workbook>";
        let child = Element::new("pivotCaches");
        let appended = append_children(
            generated,
            "workbook",
            std::slice::from_ref(&child),
            &BTreeMap::new(),
        );
        assert!(
            appended.contains("pivotCaches") && appended.trim_end().ends_with("</ns0:workbook>"),
            "{appended}"
        );
    }

    #[test]
    fn a_preserved_child_is_appended_before_the_closing_tag() {
        let generated = "<worksheet><sheetData/></worksheet>";
        let child = Element::new("pivotTableParts");
        let appended = append_children(
            generated,
            "worksheet",
            std::slice::from_ref(&child),
            &BTreeMap::new(),
        );
        assert_eq!(
            appended.trim(),
            "<worksheet><sheetData/><pivotTableParts/></worksheet>",
            "and it lands inside the element, not after it"
        );
    }

    #[test]
    fn a_generated_part_without_its_closing_tag_is_returned_unchanged() {
        // A rejected file is worse than a lossy one, so an unexpected shape gives up the
        // children rather than writing them somewhere arbitrary.
        let generated = "<worksheet><sheetData/>";
        let child = Element::new("pivotTableParts");
        let appended = append_children(
            generated,
            "</worksheet>",
            std::slice::from_ref(&child),
            &BTreeMap::new(),
        );
        assert_eq!(appended, generated);
    }

    #[test]
    fn a_sheet_whose_only_relationship_is_preserved_still_gets_a_rels_part() {
        let mut preserved = PreservedParts::default();
        preserved.add_relationships(
            "xl/worksheets/_rels/sheet1.xml.rels",
            vec![PreservedRelationship::internal(
                "rId1",
                "t/pivotTable",
                "../pivotTables/pivotTable1.xml",
            )],
        );
        let rels = rels_from_preserved(&preserved, "xl/worksheets/_rels/sheet1.xml.rels")
            .expect("the relationship still gets a part");
        assert!(rels.contains("../pivotTables/pivotTable1.xml"), "{rels}");
    }

    #[test]
    fn nothing_preserved_means_nothing_changes() {
        // Byte-identical, not merely equivalent: with nothing to add there is no reason to
        // re-serialise a part the writer produced.
        let preserved = PreservedParts::default();
        assert_eq!(
            merge_content_types(GENERATED_TYPES, &preserved),
            GENERATED_TYPES
        );
        let (merged, map) = merge_rels(GENERATED_TYPES, &preserved, "xl/_rels/workbook.xml.rels");
        assert_eq!(merged, GENERATED_TYPES);
        assert!(map.is_empty());
    }
}
