//! Parts of a package ferroxl does not model, kept so a round trip does not lose them.
//!
//! # Why this is not just "copy the bytes"
//!
//! An enterprise workbook is mostly parts this library has never heard of: pivot tables and
//! their caches, slicers, query tables, connections, external links, threaded comments,
//! ActiveX controls, `vbaProject.bin`. Preserving those files is the easy half. The hard half
//! is that a part nobody can reach is inert, and in a package the three ways of reaching a part
//! have to travel together:
//!
//! 1. **The part itself.** Its bytes, written back unchanged.
//! 2. **Its content type.** `[Content_Types].xml` has to declare it, or Excel reports the
//!    package as corrupt before it looks at anything else.
//! 3. **A relationship to it.** `<pivotCache r:id="rId4"/>` in `xl/workbook.xml` resolves
//!    through `xl/_rels/workbook.xml.rels`.
//!
//! Drop (3) and keep (1) and you get inert data: a file that opens, shows the right cells, and
//! quietly has no pivot table. So the relationship is preserved too, and its id remapped when
//! the writer has already used that id for something of its own -- with the `r:id` reference in
//! the preserved element rewritten to match.
//!
//! # What still does not survive
//!
//! - **Unknown *attributes*.** Children are carried over; attributes on `<worksheet>`,
//!   `<sheetPr>` and the like are not. An attribute ferroxl does not understand is lost.
//! - **Anything outside a `<workbook>` or `<worksheet>` root child.** Shared formulas' master
//!   cells, for instance, live on `<c>` elements that are rebuilt from the cell model.
//! - **Byte-identical zip entries.** The preserved bytes are re-compressed, so an entry is
//!   content-identical rather than byte-identical. Reproducing the original deflate stream would
//!   tie the output to one compressor version.
//! - **Parts the writer emits.** Those are the writer's to produce, and a preserved copy would
//!   collide; `reader::preserved::is_writer_owned` decides which those are.

use crate::xml::functions::Element;
use std::collections::{BTreeMap, BTreeSet};

/// Children of `<workbook>` the writer does not emit.
const PRESERVED_WORKBOOK_CHILDREN: [&str; 11] = [
    "fileVersion",
    "functionGroups",
    "externalReferences",
    "oleSize",
    "customWorkbookViews",
    "pivotCaches",
    "smartTagPr",
    "smartTagTypes",
    "webPublishing",
    "fileRecoveryPr",
    "extLst",
];

/// Children of `<worksheet>` that may be preserved, because the writer does not produce them.
///
/// A superset, not the final answer. `<drawing>`, `<legacyDrawing>` and `<tableParts>` are in it
/// because a *loaded* sheet's are preserved -- ferroxl does not read charts or legacy drawings
/// back, so its writer produces none for a loaded workbook and the originals are all that is
/// left of them. The writer drops the ones it does produce for a sheet it is building, because a
/// duplicated `<drawing>` or `<tableParts>` is a file Excel reports as corrupt.
///
/// The elements genuinely absent from the writer are preserved unconditionally. Several are on
/// the parity roadmap; when one is implemented it comes out of this list in the same change, or
/// it will be written twice.
const CANDIDATE_WORKSHEET_CHILDREN: [&str; 23] = [
    "pivotTableParts",
    "extLst",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "customSheetViews",
    "sheetProtection",
    "phoneticPr",
    "smartTags",
    "protectedRanges",
    "scenarios",
    "cellWatches",
    "dataConsolidate",
    "drawing",
    "tableParts",
    "ignoredErrors",
    "customProperties",
    "oleSize",
    "sheetCalcPr",
];

/// A relationship carried over from a loaded file.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedRelationship {
    /// The relationship id, as written in the source. May be remapped on the way out.
    pub id: String,
    /// The relationship type URI.
    pub relationship_type: String,
    /// The target, relative to the part holding the `.rels`.
    pub target: String,
    /// `TargetMode`, which is `External` for a link rather than an embedded part.
    pub target_mode: Option<String>,
}

impl PreservedRelationship {
    /// A relationship to a part inside the package.
    pub fn internal(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        PreservedRelationship {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target: target.into(),
            target_mode: None,
        }
    }

    /// A relationship to something outside the package.
    pub fn external(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        PreservedRelationship {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target: target.into(),
            target_mode: Some("External".to_string()),
        }
    }
}

/// A relationship the writer is producing itself.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GeneratedRelationship {
    /// The id the writer assigned.
    pub id: String,
    /// The relationship type URI.
    pub relationship_type: String,
    /// The target, relative to the part holding the `.rels`.
    pub target: String,
    /// `TargetMode`, when the target is outside the package.
    pub target_mode: Option<String>,
}

impl GeneratedRelationship {
    /// A relationship to a part inside the package.
    pub fn internal(
        id: impl Into<String>,
        relationship_type: impl Into<String>,
        target: impl Into<String>,
    ) -> Self {
        GeneratedRelationship {
            id: id.into(),
            relationship_type: relationship_type.into(),
            target: target.into(),
            target_mode: None,
        }
    }

    /// This relationship as XML attributes, ready for a `<Relationship>` element.
    pub fn to_attributes(&self) -> Vec<(&'static str, String)> {
        let mut attributes = vec![
            ("Id", self.id.clone()),
            ("Type", self.relationship_type.clone()),
            ("Target", self.target.clone()),
        ];
        if let Some(mode) = &self.target_mode {
            attributes.push(("TargetMode", mode.clone()));
        }
        attributes
    }
}

/// The result of merging generated relationships with preserved ones.
#[derive(Debug, Default)]
pub struct MergedRelationships {
    /// Every relationship, in the order they should be written.
    pub relationships: Vec<GeneratedRelationship>,
    /// Preserved ids that had to move, from the source id to the id actually written.
    ///
    /// Needed because the elements referencing them are written too, and a
    /// `<pivotCache r:id="rId4"/>` pointing at a relationship that is now `rId9` resolves to
    /// nothing -- silently, which is the failure this whole module exists to avoid.
    pub id_map: BTreeMap<String, String>,
}

/// A `[Content_Types].xml` entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PreservedContentType {
    /// `Default` (keyed by extension) or `Override` (keyed by part name).
    pub kind: String,
    /// The extension or the part name, depending on `kind`.
    pub name: String,
    /// The MIME type.
    pub content_type: String,
}

impl PreservedContentType {
    /// The attribute name this entry uses.
    ///
    /// The two spellings differ, and getting them the wrong way round produces a package Excel
    /// silently refuses.
    pub fn key_attribute(&self) -> &'static str {
        if self.kind == "Default" {
            "Extension"
        } else {
            "PartName"
        }
    }

    /// Whether this entry is an `Override` naming this exact part.
    pub fn covers_part(&self, path: &str) -> bool {
        self.kind != "Default" && self.name.trim_start_matches('/') == path
    }

    /// Whether this entry is a `Default` covering a part's extension.
    pub fn covers_extension(&self, extension: &str) -> bool {
        self.kind == "Default" && self.name.eq_ignore_ascii_case(extension)
    }
}

/// Everything a loaded workbook carried that the writer does not produce itself.
///
/// `Default` for an empty value, so a workbook built from scratch has none of it and pays
/// nothing.
#[derive(Debug, Clone, Default)]
pub struct PreservedParts {
    parts: BTreeMap<String, Vec<u8>>,
    content_types: Vec<PreservedContentType>,
    relationships: BTreeMap<String, Vec<PreservedRelationship>>,
    workbook_children: Vec<Element>,
    worksheet_children: BTreeMap<usize, Vec<Element>>,
    macro_extension: Option<String>,
    macro_enabled: bool,
}

impl PreservedParts {
    /// Whether anything at all was preserved.
    pub fn is_empty(&self) -> bool {
        self.parts.is_empty()
            && self.relationships.is_empty()
            && self.workbook_children.is_empty()
            && self.worksheet_children.is_empty()
    }

    /// Record a part's bytes.
    pub fn add_part(&mut self, path: impl Into<String>, data: Vec<u8>) {
        self.parts.insert(path.into(), data);
    }

    /// Record a content type entry.
    pub fn add_content_type(&mut self, entry: PreservedContentType) {
        self.content_types.push(entry);
    }

    /// Record the relationships held in a `.rels` part.
    pub fn add_relationships(
        &mut self,
        rels_path: impl Into<String>,
        relationships: Vec<PreservedRelationship>,
    ) {
        if !relationships.is_empty() {
            self.relationships.insert(rels_path.into(), relationships);
        }
    }

    /// Record a `<workbook>` child that the writer does not emit.
    pub fn add_workbook_child(&mut self, child: Element) {
        self.workbook_children.push(child);
    }

    /// Record the `<workbook>` children the writer does not emit.
    pub fn capture_workbook_children(&mut self, root: &Element) {
        self.workbook_children = root
            .children()
            .into_iter()
            .filter(|child| {
                local_name(&child.tag)
                    .is_some_and(|name| PRESERVED_WORKBOOK_CHILDREN.contains(&name))
            })
            .cloned()
            .collect();
    }

    /// Record the `<worksheet>` children that may need preserving, for a sheet.
    ///
    /// Which of these actually get written is the writer's call: it knows whether this sheet has
    /// charts, images, tables or comments, and so whether it will produce a `<drawing>`,
    /// `<legacyDrawing>` or `<tableParts>` of its own.
    pub fn capture_worksheet_children(&mut self, sheet: usize, root: &Element) {
        let children: Vec<Element> = root
            .children()
            .into_iter()
            .filter(|child| {
                local_name(&child.tag)
                    .is_some_and(|name| CANDIDATE_WORKSHEET_CHILDREN.contains(&name))
            })
            .cloned()
            .collect();
        if !children.is_empty() {
            self.worksheet_children.insert(sheet, children);
        }
    }

    /// Note that the source workbook was macro-enabled, and of what kind.
    pub fn set_macro_enabled(&mut self, extension: Option<String>) {
        self.macro_enabled = true;
        self.macro_extension = extension;
    }

    /// Whether the source workbook was macro-enabled.
    pub fn is_macro_enabled(&self) -> bool {
        self.macro_enabled
    }

    /// The source's macro-enabled extension, if it was one.
    pub fn macro_extension(&self) -> Option<&str> {
        self.macro_extension.as_deref()
    }

    /// The preserved part bytes, by path.
    pub fn parts(&self) -> &BTreeMap<String, Vec<u8>> {
        &self.parts
    }

    /// One preserved part's bytes.
    pub fn part(&self, path: &str) -> Option<&[u8]> {
        self.parts.get(path).map(|data| data.as_slice())
    }

    /// The preserved content type entries.
    pub fn content_types(&self) -> &[PreservedContentType] {
        &self.content_types
    }

    /// The preserved relationships held in one `.rels` part.
    pub fn relationships(&self, rels_path: &str) -> Option<&[PreservedRelationship]> {
        self.relationships
            .get(rels_path)
            .map(|list| list.as_slice())
    }

    /// Every `.rels` part that has preserved relationships, with them.
    pub fn all_relationships(&self) -> &BTreeMap<String, Vec<PreservedRelationship>> {
        &self.relationships
    }

    /// The preserved `<workbook>` children.
    pub fn workbook_children(&self) -> &[Element] {
        &self.workbook_children
    }

    /// The preserved `<worksheet>` children for a sheet, by zero-based index.
    pub fn worksheet_children(&self, sheet: usize) -> &[Element] {
        self.worksheet_children
            .get(&sheet)
            .map(|list| list.as_slice())
            .unwrap_or_default()
    }

    /// The preserved children for a sheet, minus any the writer is producing itself.
    ///
    /// `produced` names the child tags the writer has already emitted for this sheet. Without
    /// that argument a caller would have to know the writer's mind, and get it wrong by writing
    /// the same element twice.
    pub fn worksheet_children_excluding(&self, sheet: usize, produced: &[&str]) -> Vec<&Element> {
        self.worksheet_children(sheet)
            .iter()
            .filter(|child| local_name(&child.tag).is_none_or(|name| !produced.contains(&name)))
            .collect()
    }

    /// Whether this sheet index has preserved children worth writing.
    pub fn has_worksheet_children(&self, sheet: usize) -> bool {
        self.worksheet_children
            .get(&sheet)
            .is_some_and(|list| !list.is_empty())
    }

    /// The paths of preserved parts that no content type yet declares.
    ///
    /// A part with no declared type is a package Excel refuses, so this is the net for a source
    /// that was itself incomplete.
    pub fn untyped_part_paths(&self) -> Vec<&str> {
        self.parts
            .keys()
            .filter(|path| {
                let extension = path.rsplit('.').next().unwrap_or_default();
                !self
                    .content_types
                    .iter()
                    .any(|entry| entry.covers_part(path) || entry.covers_extension(extension))
            })
            .map(|path| path.as_str())
            .collect()
    }
}

/// The local part of a `{namespace}local` tag.
fn local_name(tag: &str) -> Option<&str> {
    tag.rsplit_once('}').map(|(_, local)| local)
}

/// Which relationship ids are already spoken for.
///
/// The writer numbers its own from `rId1` in every `.rels` part it writes, and a loaded
/// workbook's ids are usually also `rId1..rIdN` for sheets, styles and theme. So a preserved
/// `rId4` is very likely to collide, and emitting both would leave two parts claiming one id.
#[derive(Debug, Default)]
pub struct IdAllocator {
    used: BTreeSet<String>,
}

impl IdAllocator {
    /// Reserve the ids a writer has already emitted.
    pub fn reserve(&mut self, ids: impl IntoIterator<Item = String>) {
        self.used.extend(ids);
    }

    /// Map a source id to a free one, recording the choice.
    ///
    /// An id that is already free is kept, so a workbook with nothing to remap comes out with
    /// the ids it went in with -- which keeps `r:id` references in preserved elements valid
    /// without a rewrite.
    pub fn allocate(&mut self, wanted: &str) -> String {
        if !self.used.contains(wanted) {
            self.used.insert(wanted.to_string());
            return wanted.to_string();
        }
        // `rId` with a number is the overwhelmingly common form, so the search starts there
        // rather than at a hash that would produce an id no tool expects to see.
        let prefix = match wanted.strip_prefix("rId") {
            Some(rest) if !rest.is_empty() && rest.chars().all(|c| c.is_ascii_digit()) => {
                "rId".to_string()
            }
            _ => format!("{wanted}_"),
        };
        let mut counter = 1usize;
        loop {
            let candidate = format!("{prefix}{counter}");
            if !self.used.contains(&candidate) {
                self.used.insert(candidate.clone());
                return candidate;
            }
            counter += 1;
        }
    }

    /// Whether `id` is already taken.
    pub fn is_used(&self, id: &str) -> bool {
        self.used.contains(id)
    }
}

/// Resolve a relationship `target` to the package-absolute path it names.
///
/// A target is spelled three ways in the wild, and all three reach this code:
///
/// - absolute from the package root, `/xl/worksheets/sheet1.xml`, which is what openpyxl
///   writes;
/// - relative to the declaring part's directory, `worksheets/sheet1.xml`, which is what this
///   writer emits and what Excel writes;
/// - relative with a step up, `../drawings/drawing1.xml`, which is what a sheet's `.rels` uses.
///
/// `base_directory` is the directory of the part holding the `.rels`, so `xl/_rels/workbook.xml.rels`
/// gives `xl` and `xl/worksheets/_rels/sheet1.xml.rels` gives `xl/worksheets`. Comparing the raw
/// strings instead is the defect this exists to remove: `worksheets/sheet1.xml` and
/// `/xl/worksheets/sheet1.xml` are one part, and treating them as two leaves a duplicate
/// relationship behind every save.
pub fn resolve_target(base_directory: &str, target: &str) -> String {
    // A leading `/` means the target is already package-absolute, so `base_directory` does not
    // apply. Reading it as relative is the easy mistake here, and it produces a path that names
    // no part at all -- `xl/xl/styles.xml` -- which then matches nothing.
    let absolute = target.starts_with('/');
    let trimmed = target.trim_start_matches('/');
    let mut segments: Vec<&str> = if absolute || base_directory.is_empty() {
        Vec::new()
    } else {
        base_directory
            .split('/')
            .filter(|s| !s.is_empty())
            .collect()
    };
    for piece in trimmed.split('/') {
        match piece {
            // A `Target` may carry a fragment or a query; neither is part of the part's name.
            "" | "." => {}
            ".." => {
                segments.pop();
            }
            piece => segments.push(piece.split(['?', '#']).next().unwrap_or(piece)),
        }
    }
    segments.join("/")
}

/// The directory a `.rels` part's targets are relative to.
///
/// `_rels/.rels` declares the package itself rather than a part, so its targets resolve from the
/// root. It has to be named rather than derived: there is no owning part to take a directory
/// from, and reading one out of the path yields `_rels`, which would resolve every target into a
/// directory that does not exist and so match nothing.
pub fn rels_base_directory(rels_path: &str) -> String {
    if rels_path == "_rels/.rels" || rels_path == ".rels" {
        return String::new();
    }
    let owner = crate::reader::preserved::part_for_rels(rels_path);
    match owner.rsplit_once('/') {
        Some((directory, _)) => directory.to_string(),
        None => String::new(),
    }
}

/// Merge the writer's relationships with a source's, moving any preserved id that would collide.
///
/// `rels_path` is the `.rels` being written, which is what says what the targets are relative to.
///
/// A preserved relationship is dropped when it is redundant or stale, which between them is every
/// relationship into a part the writer owns:
///
/// - **redundant** -- its target is one the writer emits, so two relationships name one part.
///   That is not what the source said, and Excel resolves whichever id an element happens to name.
/// - **stale** -- its target is in a family the writer renumbers from scratch, such as
///   `xl/worksheets/sheetN.xml`. The writer is the only authority on which of those exist, so a
///   preserved one naming a part that is not there would leave the package with a relationship
///   pointing at nothing -- which is a file Excel refuses to open.
///
/// Relationships into a family the writer does *not* own are kept, and that is the whole reason
/// this function exists: a preserved pivot cache, drawing, chart or VBA project is reached only
/// through a relationship the source wrote and nothing here would write it again.
pub fn merge_relationships(
    generated: &[GeneratedRelationship],
    preserved: &[PreservedRelationship],
    rels_path: &str,
) -> MergedRelationships {
    let base = rels_base_directory(rels_path);
    let mut allocator = IdAllocator::default();
    allocator.reserve(generated.iter().map(|entry| entry.id.clone()));

    let generated_targets: BTreeSet<String> = generated
        .iter()
        .map(|entry| resolve_target(&base, &entry.target))
        .collect();

    let mut merged = MergedRelationships {
        relationships: generated.to_vec(),
        id_map: BTreeMap::new(),
    };
    for entry in preserved {
        if entry.target_mode.is_none() {
            let resolved = resolve_target(&base, &entry.target);
            if generated_targets.contains(&resolved)
                || crate::reader::preserved::is_writer_owned(&resolved)
            {
                continue;
            }
        }
        let id = allocator.allocate(&entry.id);
        if id != entry.id {
            merged.id_map.insert(entry.id.clone(), id.clone());
        }
        merged.relationships.push(GeneratedRelationship {
            id,
            relationship_type: entry.relationship_type.clone(),
            target: entry.target.clone(),
            target_mode: entry.target_mode.clone(),
        });
    }
    merged
}

/// Rewrite every `r:id` in an element subtree through `id_map`.
///
/// A no-op for an id that did not move, which is the common case.
pub fn rewrite_relationship_ids(element: &mut Element, id_map: &BTreeMap<String, String>) {
    if id_map.is_empty() {
        return;
    }
    for (key, value) in element.attributes.iter_mut() {
        let is_relationship_id = key == "r:id" || (key.starts_with('{') && key.ends_with("}id"));
        if is_relationship_id {
            if let Some(moved) = id_map.get(value.as_str()) {
                *value = moved.clone();
            }
        }
    }
    for child in element.children.iter_mut() {
        rewrite_relationship_ids(child, id_map);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::functions::fromstring;

    #[test]
    fn an_id_that_is_free_is_kept_so_references_stay_valid() {
        let mut allocator = IdAllocator::default();
        allocator.reserve(["rId1".to_string(), "rId2".to_string()]);
        assert_eq!(allocator.allocate("rId7"), "rId7");
    }

    #[test]
    fn a_colliding_id_is_moved_rather_than_duplicated() {
        let mut allocator = IdAllocator::default();
        allocator.reserve(["rId1".to_string(), "rId2".to_string()]);
        // The source's `rId2` pointed at a pivot cache. Emitting it as `rId2` too would leave two
        // parts claiming one id.
        assert_eq!(allocator.allocate("rId2"), "rId3");
        assert_eq!(allocator.allocate("rId3"), "rId4");
    }

    #[test]
    fn a_non_numeric_id_gets_a_suffixed_replacement() {
        let mut allocator = IdAllocator::default();
        allocator.reserve(["pkg1".to_string()]);
        assert_eq!(allocator.allocate("pkg1"), "pkg1_1");
    }

    #[test]
    fn nothing_is_reported_preserved_for_an_empty_value() {
        assert!(PreservedParts::default().is_empty());
    }

    #[test]
    fn a_part_without_a_content_type_is_reported() {
        let mut parts = PreservedParts::default();
        parts.add_part("xl/pivotTables/pivotTable1.xml", vec![1]);
        parts.add_part("xl/media/image9.png", vec![2]);
        parts.add_content_type(PreservedContentType {
            kind: "Default".to_string(),
            name: "png".to_string(),
            content_type: "image/png".to_string(),
        });
        assert_eq!(
            parts.untyped_part_paths(),
            vec!["xl/pivotTables/pivotTable1.xml"]
        );
    }

    #[test]
    fn an_override_satisfies_a_part_that_its_extension_would_not() {
        let mut parts = PreservedParts::default();
        parts.add_part("customXml/item1.xml", vec![1]);
        parts.add_content_type(PreservedContentType {
            kind: "Override".to_string(),
            name: "/customXml/item1.xml".to_string(),
            content_type: "application/xml".to_string(),
        });
        assert!(parts.untyped_part_paths().is_empty());
    }

    #[test]
    fn a_relationship_renders_the_two_attribute_spellings_correctly() {
        let embedded = GeneratedRelationship::internal("rId1", "type/a", "pivot1.xml");
        let attributes: BTreeMap<&str, String> = embedded.to_attributes().into_iter().collect();
        assert_eq!(
            attributes.get("Target").map(String::as_str),
            Some("pivot1.xml")
        );
        assert!(!attributes.contains_key("TargetMode"));

        let linked = GeneratedRelationship {
            target_mode: Some("External".to_string()),
            ..GeneratedRelationship::internal("rId2", "type/b", "https://example.com")
        };
        let attributes: BTreeMap<&str, String> = linked.to_attributes().into_iter().collect();
        assert_eq!(
            attributes.get("TargetMode").map(String::as_str),
            Some("External")
        );
    }

    #[test]
    fn a_content_type_reports_which_attribute_keys_it() {
        let by_extension = PreservedContentType {
            kind: "Default".to_string(),
            name: "bin".to_string(),
            content_type: "application/vnd.ms-office.vbaProject".to_string(),
        };
        let by_part = PreservedContentType {
            kind: "Override".to_string(),
            name: "/xl/workbook.xml".to_string(),
            content_type: "application/xml".to_string(),
        };
        assert_eq!(by_extension.key_attribute(), "Extension");
        assert_eq!(by_part.key_attribute(), "PartName");
    }

    #[test]
    fn a_colliding_preserved_id_is_remapped_and_the_reference_follows_it() {
        let generated = vec![
            GeneratedRelationship::internal("rId1", "t/sheet", "worksheets/sheet1.xml"),
            GeneratedRelationship::internal("rId2", "t/styles", "styles.xml"),
        ];
        let preserved = vec![PreservedRelationship::internal(
            "rId2",
            "t/pivotCacheDefinition",
            "pivotCache/pivotCacheDefinition1.xml",
        )];
        let merged = merge_relationships(&generated, &preserved, "xl/_rels/workbook.xml.rels");

        assert_eq!(merged.relationships.len(), 3);
        let moved = merged
            .relationships
            .last()
            .expect("the preserved one is last");
        assert_eq!(moved.id, "rId3", "it had to move: rId2 was taken");
        assert_eq!(
            merged.id_map.get("rId2").map(String::as_str),
            Some("rId3"),
            "and the reference has to be told where it went"
        );

        const REL_ID: &str =
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id";
        let mut element = Element::with_attributes("pivotCache", [(REL_ID, "rId2".to_string())]);
        rewrite_relationship_ids(&mut element, &merged.id_map);
        assert_eq!(element.get(REL_ID), Some("rId3"));
    }

    #[test]
    fn a_preserved_relationship_the_writer_also_produces_is_dropped() {
        // Two relationships to one part is not what the source said, and Excel would resolve
        // whichever id an element happened to name.
        let generated = vec![GeneratedRelationship::internal(
            "rId1",
            "t/theme",
            "theme/theme1.xml",
        )];
        let preserved = vec![PreservedRelationship::internal(
            "rId9",
            "t/theme",
            "theme/theme1.xml",
        )];
        let merged = merge_relationships(&generated, &preserved, "xl/_rels/workbook.xml.rels");
        assert_eq!(merged.relationships.len(), 1);
        assert!(
            merged.id_map.is_empty(),
            "nothing moved, so nothing needs rewriting"
        );
    }

    /// The defect this all came from: a target the source spelled absolutely and the writer
    /// spelled relatively is still the same part.
    ///
    /// openpyxl writes `/xl/worksheets/sheet1.xml`; this writer emits
    /// `worksheets/sheet1.xml`. Comparing the strings found two different targets, kept both,
    /// and a `remove_sheet` followed by any other edit then duplicated the remaining sheet into
    /// the file -- `list_sheets` reported a `Sheet2` that was not there.
    #[test]
    fn a_target_spelled_absolutely_is_recognised_as_the_same_part() {
        let generated = vec![GeneratedRelationship::internal(
            "rId1",
            "t/sheet",
            "worksheets/sheet1.xml",
        )];
        let preserved = vec![PreservedRelationship::internal(
            "rId5",
            "t/sheet",
            "/xl/worksheets/sheet1.xml",
        )];
        let merged = merge_relationships(&generated, &preserved, "xl/_rels/workbook.xml.rels");
        assert_eq!(
            merged.relationships.len(),
            1,
            "one part, one relationship: {:?}",
            merged.relationships
        );
    }

    /// A preserved relationship into a part the writer renumbers is stale, not redundant.
    ///
    /// Removing a sheet from a three-sheet workbook leaves the source's relationship to
    /// `sheet3.xml` naming a part the writer never emitted. A package with a relationship
    /// pointing at nothing is a file Excel reports as needing repair, which is the same failure
    /// as the drawing bug in 0.1.10 reached by a different door.
    #[test]
    fn a_preserved_relationship_to_a_removed_part_is_dropped() {
        let generated = vec![
            GeneratedRelationship::internal("rId1", "t/sheet", "worksheets/sheet1.xml"),
            GeneratedRelationship::internal("rId2", "t/sheet", "worksheets/sheet2.xml"),
        ];
        let preserved = vec![
            PreservedRelationship::internal("rId1", "t/sheet", "/xl/worksheets/sheet1.xml"),
            PreservedRelationship::internal("rId2", "t/sheet", "/xl/worksheets/sheet2.xml"),
            PreservedRelationship::internal("rId3", "t/sheet", "/xl/worksheets/sheet3.xml"),
        ];
        let merged = merge_relationships(&generated, &preserved, "xl/_rels/workbook.xml.rels");
        assert_eq!(
            merged.relationships.len(),
            2,
            "the writer said which sheets exist; sheet3.xml is not one of them: {:?}",
            merged.relationships
        );
    }

    /// A relationship into a family the writer does not own is exactly what is being kept.
    ///
    /// A pivot cache, a chart, a drawing and a VBA project are reached only through a
    /// relationship the source wrote. Dropping these would lose the parts along with it.
    #[test]
    fn a_relationship_to_a_preserved_part_survives() {
        let generated = vec![GeneratedRelationship::internal(
            "rId1",
            "t/sheet",
            "worksheets/sheet1.xml",
        )];
        let preserved = vec![
            PreservedRelationship::internal(
                "rId2",
                "t/pivotCacheDefinition",
                "pivotCache/pivotCacheDefinition1.xml",
            ),
            PreservedRelationship::internal("rId3", "t/vbaProject", "vbaProject.bin"),
        ];
        // A sheet's relationships, which is where a drawing is reached from.
        let merged = merge_relationships(
            &generated,
            &preserved,
            "xl/worksheets/_rels/sheet1.xml.rels",
        );
        assert_eq!(merged.relationships.len(), 3, "{:?}", merged.relationships);

        let drawing = [PreservedRelationship::internal(
            "rId4",
            "t/drawing",
            "../drawings/drawing1.xml",
        )];
        let merged =
            merge_relationships(&generated, &drawing, "xl/worksheets/_rels/sheet1.xml.rels");
        assert_eq!(
            merged.relationships.len(),
            2,
            "the relationship to the drawing the writer cannot reproduce: {:?}",
            merged.relationships
        );
    }

    #[test]
    fn a_target_resolves_through_a_step_up_and_to_the_package_root() {
        // `..` has to be resolved before the comparison, or `../tables/table1.xml` from a
        // sheet and `xl/tables/table1.xml` from the workbook look like different parts.
        assert_eq!(
            resolve_target("xl/worksheets", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
        assert_eq!(resolve_target("xl", "/xl/styles.xml"), "xl/styles.xml");
        assert_eq!(
            resolve_target("xl/worksheets", "../../docProps/app.xml"),
            "docProps/app.xml"
        );
        // `_rels/.rels` declares the package, so its targets resolve from the root.
        assert_eq!(rels_base_directory("_rels/.rels"), "");
        assert_eq!(rels_base_directory("xl/_rels/workbook.xml.rels"), "xl");
        assert_eq!(
            rels_base_directory("xl/worksheets/_rels/sheet1.xml.rels"),
            "xl/worksheets"
        );
    }

    #[test]
    fn an_external_relationship_is_kept_even_when_the_target_looks_internal() {
        let generated = vec![GeneratedRelationship::internal(
            "rId1",
            "t/hyperlink",
            "https://example.com",
        )];
        let preserved = vec![PreservedRelationship::external(
            "rId5",
            "t/hyperlink",
            "https://example.com",
        )];
        let merged = merge_relationships(&generated, &preserved, "xl/_rels/workbook.xml.rels");
        assert_eq!(
            merged.relationships.len(),
            2,
            "a link is not the same as an embedded part"
        );
    }

    #[test]
    fn the_writer_produced_children_are_excluded_from_the_preserved_set() {
        let root = fromstring(
            br#"<worksheet xmlns="x">
                 <drawing r:id="rId1"/>
                 <pivotTableParts count="1"><pivotTablePart r:id="rId2"/></pivotTableParts>
               </worksheet>"#,
        )
        .expect("parses");
        let mut parts = PreservedParts::default();
        parts.capture_worksheet_children(0, &root);
        assert_eq!(parts.worksheet_children(0).len(), 2);
        assert_eq!(parts.worksheet_children_excluding(0, &["drawing"]).len(), 1);
    }

    #[test]
    fn a_workbook_child_the_writer_will_emit_is_not_preserved() {
        // `<sheets>` is written from the model. Preserving the source's would duplicate every
        // sheet.
        let root = fromstring(
            br#"<workbook xmlns="x"><sheets><sheet name="a"/></sheets><pivotCaches/></workbook>"#,
        )
        .expect("parses");
        let mut parts = PreservedParts::default();
        parts.capture_workbook_children(&root);
        assert_eq!(parts.workbook_children().len(), 1);
        // The tag keeps its namespace prefix, which is what makes re-emitting it valid; the
        // filtering is by local name because that is the part that identifies the element.
        assert!(
            parts.workbook_children()[0].tag.ends_with("pivotCaches"),
            "got {:?}",
            parts.workbook_children()[0].tag
        );
    }
}
