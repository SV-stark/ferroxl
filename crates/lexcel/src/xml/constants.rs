//! Namespace URIs, package paths and content types (`openpyxl/xml/constants.py`).

/// Lowest valid row index (0-based).
pub const MIN_ROW: u32 = 0;
/// Lowest valid column index (0-based).
pub const MIN_COLUMN: u32 = 0;
/// Maximum number of columns in a worksheet.
pub const MAX_COLUMN: u32 = 16384;
/// Maximum number of rows in a worksheet.
pub const MAX_ROW: u32 = 1048576;
/// Lowest valid column index in A1 notation.
pub const MIN_ROW_1: u32 = 1;

/// The `docProps` folder.
pub const PACKAGE_PROPS: &str = "docProps";
/// The `xl` folder.
pub const PACKAGE_XL: &str = "xl";
/// The `_rels` folder.
pub const PACKAGE_RELS: &str = "_rels";
/// The theme part folder.
pub const PACKAGE_THEME: &str = "xl/theme";
/// The worksheets folder.
pub const PACKAGE_WORKSHEETS: &str = "xl/worksheets";
/// The drawings folder.
pub const PACKAGE_DRAWINGS: &str = "xl/drawings";
/// The charts folder.
pub const PACKAGE_CHARTS: &str = "xl/charts";
/// The embedded media folder.
pub const PACKAGE_IMAGES: &str = "xl/media";
/// The worksheet relationships folder.
pub const PACKAGE_WORKSHEET_RELS: &str = "xl/worksheets/_rels";

/// `[Content_Types].xml`.
pub const ARC_CONTENT_TYPES: &str = "[Content_Types].xml";
/// The package root relationships.
pub const ARC_ROOT_RELS: &str = "_rels/.rels";
/// The workbook relationships part.
pub const ARC_WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
/// The core properties part.
pub const ARC_CORE: &str = "docProps/core.xml";
/// The extended properties part.
pub const ARC_APP: &str = "docProps/app.xml";
/// The workbook part.
pub const ARC_WORKBOOK: &str = "xl/workbook.xml";
/// The shared styles part.
pub const ARC_STYLE: &str = "xl/styles.xml";
/// The theme part.
pub const ARC_THEME: &str = "xl/theme/theme1.xml";
/// The shared strings part.
pub const ARC_SHARED_STRINGS: &str = "xl/sharedStrings.xml";
/// The custom UI part.
pub const ARC_CUSTOM_UI: &str = "customUI/customUI.xml";

/// Folder prefixes copied verbatim when preserving a VBA archive.
pub const ARC_VBA: [&str; 10] = [
    "xl/vba",
    "xl/activeX",
    "xl/drawings",
    "xl/media",
    "xl/ctrlProps",
    "xl/worksheets/_rels",
    "customUI",
    "xl/printerSettings",
    "xl/charts",
    "xl/theme",
];

/// DrawingML chart namespace.
pub const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/// DrawingML main namespace.
pub const DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Comments relationship type.
pub const COMMENTS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
/// VML drawing relationship type.
pub const VML_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
/// SpreadsheetDrawing namespace.
pub const SHEET_DRAWING_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
/// ChartDrawing namespace.
pub const CHART_DRAWING_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
/// Office relationship namespace.
pub const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
/// Package relationship namespace.
pub const PKG_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
/// Variant types namespace.
pub const VTYPES_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
/// Extended properties namespace.
pub const XPROPS_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
/// Core properties namespace.
pub const COREPROPS_NS: &str =
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
/// Content types namespace.
pub const CONTYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
/// Dublin Core elements namespace.
pub const DCORE_NS: &str = "http://purl.org/dc/elements/1.1/";
/// Dublin Core terms namespace.
pub const DCTERMS_NS: &str = "http://purl.org/dc/terms/";
/// Prefix used for the Dublin Core terms namespace.
pub const DCTERMS_PREFIX: &str = "dcterms";
/// XML Schema instance namespace.
pub const XSI_NS: &str = "http://www.w3.org/2001/XMLSchema-instance";
/// Reserved `xml:` namespace.
pub const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
/// SpreadsheetML main namespace.
pub const SHEET_MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
/// Custom UI extensibility namespace.
pub const CUSTOMUI_NS: &str =
    "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility";

/// The prefix → namespace mapping registered by `openpyxl.xml.functions`.
pub const NAMESPACES: [(&str, &str); 9] = [
    ("cp", COREPROPS_NS),
    ("dc", DCORE_NS),
    (DCTERMS_PREFIX, DCTERMS_NS),
    ("dcmitype", "http://purl.org/dc/dcmitype/"),
    ("xsi", XSI_NS),
    ("vt", VTYPES_NS),
    ("xml", XML_NS),
    ("main", SHEET_MAIN_NS),
    ("c", CHART_NS),
];

/// The prefix used for SpreadsheetML elements when writing.
pub const SHEET_MAIN_PREFIX: &str = "s";
/// The prefix used for relationship-qualified attributes when writing.
pub const REL_PREFIX: &str = "r";

/// Helper that formats a `{namespace}tag` selector like ElementTree expects.
pub fn qname(ns: &str, tag: &str) -> String {
    format!("{{{ns}}}{tag}")
}

/// Reverse of [`qname`]: strips the `{ns}` prefix from a qualified name.
pub fn local_name(tag: &str) -> &str {
    match tag.rfind('}') {
        Some(idx) => &tag[idx + 1..],
        None => tag,
    }
}

/// Returns the namespace portion of a qualified name, if any.
pub fn namespace_of(tag: &str) -> Option<&str> {
    tag.strip_prefix('{')
        .and_then(|rest| rest.split_once('}').map(|(ns, _)| ns))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn qname_round_trip() {
        let q = qname(SHEET_MAIN_NS, "sheetData");
        assert_eq!(namespace_of(&q), Some(SHEET_MAIN_NS));
        assert_eq!(local_name(&q), "sheetData");
        assert_eq!(namespace_of("plain"), None);
        assert_eq!(local_name("plain"), "plain");
    }

    #[test]
    fn namespace_constants_are_distinct() {
        assert_ne!(SHEET_MAIN_NS, DRAWING_NS);
        assert_ne!(CHART_NS, CHART_DRAWING_NS);
        assert!(ARC_VBA.contains(&"xl/vba"));
    }
}
