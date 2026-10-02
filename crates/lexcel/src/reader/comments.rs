//! Reading comments (`openpyxl/reader/comments.py`).

use crate::comments::Comment;
use crate::exceptions::Result;
use crate::xml::constants::{COMMENTS_NS, PACKAGE_WORKSHEETS};
use crate::xml::functions::fromstring;

/// The comments part path for a worksheet part name.
///
/// Returns `None` when the worksheet has no relationships part.
pub fn comments_file_path(
    worksheet_part: &str,
    rels_source: &[u8],
    valid_files: &[String],
) -> Option<String> {
    // Only a worksheet can have comments, so anything outside `xl/worksheets/` is
    // rejected before the relationships are consulted.
    if !worksheet_part.starts_with(PACKAGE_WORKSHEETS) {
        return None;
    }
    let root = fromstring(rels_source).ok()?;
    for node in root.children() {
        if node.get("Type") == Some(COMMENTS_NS) {
            let Some(target) = node.get("Target") else {
                continue;
            };
            // Resolve `../comments1.xml` relative to `xl/worksheets/`.
            let normalised = normalise(&format!("{PACKAGE_WORKSHEETS}/{target}"));
            if valid_files.iter().any(|name| name == &normalised) {
                return Some(normalised);
            }
        }
    }
    None
}

/// Normalise `..` segments in a part path, as `os.path.normpath` would.
fn normalise(path: &str) -> String {
    let mut parts: Vec<&str> = Vec::new();
    for part in path.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            other => parts.push(other),
        }
    }
    parts.join("/")
}

/// Read the author list from a comments part.
pub fn get_author_list(xml_source: &[u8]) -> Result<Vec<String>> {
    let root = fromstring(xml_source)?;
    let authors_tag = format!("{{{}}}authors", crate::xml::constants::SHEET_MAIN_NS);
    let Some(authors) = root.find(&authors_tag) else {
        return Ok(Vec::new());
    };
    Ok(authors
        .children()
        .iter()
        .map(|node| node.text.clone().unwrap_or_default())
        .collect())
}

/// A comment read from a comments part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedComment {
    /// The cell the comment is attached to.
    pub coordinate: String,
    /// The author index.
    pub author_id: usize,
    /// The comment text.
    pub text: String,
}

/// Parse a comments part.
pub fn parse_comments(xml_source: &[u8]) -> Result<Vec<ParsedComment>> {
    let root = fromstring(xml_source)?;
    let main = crate::xml::constants::SHEET_MAIN_NS;
    let comment_tag = format!("{{{main}}}comment");
    // The comments live under `<commentList>`, so the list is searched rather than the root.
    let list_tag = format!("{{{main}}}commentList");
    let text_tag = format!("{{{main}}}text");
    let run_tag = format!("{{{main}}}r");
    let value_tag = format!("{{{main}}}t");

    let mut out = Vec::new();
    let Some(list) = root.find(&list_tag) else {
        return Ok(out);
    };
    for node in list.find_all(&comment_tag) {
        let Some(coordinate) = node.get("ref") else {
            continue;
        };
        let author_id = node
            .get("authorId")
            .and_then(|v| v.trim().parse::<usize>().ok())
            .unwrap_or(0);
        let mut pieces = Vec::new();
        if let Some(text_node) = node.find(&text_tag) {
            for run in text_node.find_all(&run_tag) {
                for value in run.find_all(&value_tag) {
                    pieces.push(value.text.clone().unwrap_or_default());
                }
            }
        }
        out.push(ParsedComment {
            coordinate: coordinate.to_string(),
            author_id,
            text: pieces.concat(),
        });
    }
    Ok(out)
}

/// Assign the comments from a comments part to the worksheet's cells.
pub fn read_comments(
    worksheet: &mut crate::worksheet::Worksheet,
    xml_source: &[u8],
) -> Result<usize> {
    let authors = get_author_list(xml_source)?;
    let comments = parse_comments(xml_source)?;
    let mut count = 0usize;
    for parsed in comments {
        let author = authors.get(parsed.author_id).cloned().unwrap_or_default();
        worksheet.set_comment(&parsed.coordinate, Some(Comment::new(parsed.text, author)))?;
        count += 1;
    }
    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::worksheet::Worksheet;

    const MAIN: &str = r#"xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#;

    #[test]
    fn reads_authors_and_comments() {
        let xml = format!(
            r#"<comments {MAIN}><authors><author>Alice</author><author>Bob</author></authors><commentList><comment ref="A1" authorId="0"><text><r><t>Hello</t></r></text></comment><comment ref="C3" authorId="1"><text><r><t>World</t></r><r><t>!</t></r></text></comment></commentList></comments>"#
        );
        let authors = get_author_list(xml.as_bytes()).unwrap();
        assert_eq!(authors, vec!["Alice".to_string(), "Bob".to_string()]);
        let comments = parse_comments(xml.as_bytes()).unwrap();
        assert_eq!(comments.len(), 2);
        assert_eq!(comments[0].coordinate, "A1");
        assert_eq!(comments[0].text, "Hello");
        assert_eq!(comments[1].text, "World!");
    }

    #[test]
    fn assigns_comments_to_cells() {
        let xml = format!(
            r#"<comments {MAIN}><authors><author>Alice</author></authors><commentList><comment ref="B2" authorId="0"><text><r><t>Note</t></r></text></comment></commentList></comments>"#
        );
        let mut worksheet = Worksheet::new("Sheet1").unwrap();
        let count = read_comments(&mut worksheet, xml.as_bytes()).unwrap();
        assert_eq!(count, 1);
        assert_eq!(worksheet.comment_count(), 1);
        let comment = worksheet.get_cell("B2").unwrap().comment.as_ref().unwrap();
        assert_eq!(comment.text(), "Note");
        assert_eq!(comment.author(), "Alice");
    }

    #[test]
    fn resolves_the_comments_part_path() {
        let rels = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{COMMENTS_NS}" Target="../comments1.xml"/></Relationships>"#
        );
        let valid = vec!["xl/comments1.xml".to_string()];
        assert_eq!(
            comments_file_path("xl/worksheets/sheet1.xml", rels.as_bytes(), &valid),
            Some("xl/comments1.xml".to_string())
        );
        assert_eq!(
            comments_file_path("xl/worksheets/sheet1.xml", rels.as_bytes(), &[]),
            None
        );
    }

    #[test]
    fn unrelated_relationship_types_are_ignored() {
        let rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://example.com/other" Target="../comments1.xml"/></Relationships>"#;
        let valid = vec!["xl/comments1.xml".to_string()];
        assert_eq!(
            comments_file_path("xl/worksheets/sheet1.xml", rels.as_bytes(), &valid),
            None
        );
    }

    #[test]
    fn missing_authors_yields_an_empty_list() {
        assert!(get_author_list(b"<comments/>").unwrap().is_empty());
        assert!(parse_comments(b"<comments/>").unwrap().is_empty());
    }
}
