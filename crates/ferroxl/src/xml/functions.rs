//! Shared XML tools (`openpyxl/xml/functions.py`, `namespace.py`).
//!
//! The Python original leans on ElementTree both for building output and for parsing
//! input, exposing a handful of helpers (`get_document_content`, `pretty_indent`,
//! `start_tag`, `tag`, `safe_iterator`, `ConditionalElement`). This module provides the
//! same surface: a small DOM for building parts, a compact streaming writer, and a
//! namespace-aware parser that maps to ElementTree-style `{ns}tag` selectors.

use std::collections::BTreeMap;
use std::io::Read;

use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::{Reader, Writer as XmlWriterInner};

use crate::exceptions::{Error, Result};
use crate::xml::constants::*;

/// Prefixes registered by `openpyxl.xml.functions.register_namespace`.
///
/// ElementTree looks these up when serialising, so a `{SHEET_MAIN_NS}workbook` tag
/// becomes `s:workbook`.
pub const REGISTERED_PREFIXES: [(&str, &str); 10] = [
    (DCTERMS_PREFIX, DCTERMS_NS),
    ("dcmitype", "http://purl.org/dc/dcmitype/"),
    ("cp", COREPROPS_NS),
    ("c", CHART_NS),
    ("a", DRAWING_NS),
    ("s", SHEET_MAIN_NS),
    ("r", REL_NS),
    ("vt", VTYPES_NS),
    ("xdr", SHEET_DRAWING_NS),
    ("cdr", CHART_DRAWING_NS),
];

/// A namespace-qualified name, stored ElementTree-style as `{ns}local`.
///
/// Names without a namespace are stored bare (`"dimension"`).
pub type QName = String;

/// An in-memory XML element, mirroring `xml.etree.ElementTree.Element`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Element {
    /// Element tag, either `{ns}local` or bare.
    pub tag: QName,
    /// Attributes in insertion order; keys may be `{ns}local` or bare.
    pub attributes: Vec<(String, String)>,
    /// Text content directly under this element.
    pub text: Option<String>,
    /// Text following this element (used for pretty-printing).
    pub tail: Option<String>,
    /// Child elements.
    pub children: Vec<Element>,
}

impl Element {
    /// Create an element with the given qualified or bare tag.
    pub fn new(tag: impl Into<QName>) -> Self {
        Element {
            tag: tag.into(),
            ..Default::default()
        }
    }

    /// Create an element with attributes, preserving their order.
    ///
    /// The tag and both attribute fields are anything that renders to a string, so callers
    /// can pass computed values without converting them by hand.
    pub fn with_attributes<T, I, K, V>(tag: T, attrs: I) -> Self
    where
        T: std::fmt::Display,
        I: IntoIterator<Item = (K, V)>,
        K: std::fmt::Display,
        V: std::fmt::Display,
    {
        let mut el = Element::new(tag.to_string());
        for (k, v) in attrs {
            el.set(k.to_string(), v.to_string());
        }
        el
    }

    /// Set an attribute, replacing any existing value for that key.
    pub fn set(&mut self, key: impl Into<String>, value: impl Into<String>) -> &mut Self {
        let key = key.into();
        let value = value.into();
        match self.attributes.iter_mut().find(|(k, _)| *k == key) {
            Some(slot) => slot.1 = value,
            None => self.attributes.push((key, value)),
        }
        self
    }

    /// Get an attribute value by qualified or bare key.
    pub fn get(&self, key: impl AsRef<str>) -> Option<&str> {
        let key = key.as_ref();
        self.attributes
            .iter()
            .find(|(k, _)| k == key)
            .map(|(_, v)| v.as_str())
    }

    /// Set the text content.
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Self {
        self.text = Some(text.into());
        self
    }

    /// Append a child element.
    pub fn append(&mut self, child: Element) -> &mut Self {
        self.children.push(child);
        self
    }

    /// Insert a child at `index`.
    pub fn insert(&mut self, index: usize, child: Element) -> &mut Self {
        self.children.insert(index, child);
        self
    }

    /// First direct child matching `path`, where `path` may be a bare tag name or an
    /// ElementTree path such as `*/{ns}workbookView`.
    pub fn find(&self, path: impl AsRef<str>) -> Option<&Element> {
        let path = path.as_ref();
        let (ancestor, tag) = match split_path(path) {
            Some((a, t)) => (Some(a), t),
            None => (None, path),
        };
        if let Some(ancestor) = ancestor {
            if !self.matches_step(ancestor) {
                return None;
            }
        }
        self.children.iter().find(|child| child.matches_tag(tag))
    }

    /// Mutable variant of [`Element::find`].
    pub fn find_mut(&mut self, path: impl AsRef<str>) -> Option<&mut Element> {
        let path = path.as_ref();
        let (ancestor, tag) = match split_path(path) {
            Some((a, t)) => (Some(a), t),
            None => (None, path),
        };
        if let Some(ancestor) = ancestor {
            if !self.matches_step(ancestor) {
                return None;
            }
        }
        self.children
            .iter_mut()
            .find(|child| child.matches_tag(tag))
    }

    fn matches_step(&self, step: &str) -> bool {
        step == "*" || self.matches_tag(step)
    }

    /// Whether this element's tag answers to `tag`.
    ///
    /// An unqualified query matches by local name. ElementTree would demand the `{ns}`
    /// form, but the writer's parts are all in one namespace and its own tests read far
    /// more clearly without the prefix repeated on every lookup.
    fn matches_tag(&self, tag: &str) -> bool {
        self.tag == tag
            || (super::constants::namespace_of(tag).is_none()
                && super::constants::local_name(&self.tag) == tag)
    }

    /// All direct children matching `tag`.
    pub fn find_all(&self, tag: impl AsRef<str>) -> Vec<&Element> {
        let tag = tag.as_ref();
        self.children
            .iter()
            .filter(|c| c.matches_tag(tag))
            .collect()
    }

    /// Mutable variant of [`Element::find_all`].
    pub fn find_all_mut(&mut self, tag: impl AsRef<str>) -> Vec<&mut Element> {
        let tag = tag.as_ref();
        self.children
            .iter_mut()
            .filter(|c| c.matches_tag(tag))
            .collect()
    }

    /// All direct children, in document order.
    pub fn children(&self) -> Vec<&Element> {
        self.children.iter().collect()
    }

    /// Mutable variant of [`Element::children`].
    pub fn children_mut(&mut self) -> Vec<&mut Element> {
        self.children.iter_mut().collect()
    }

    /// Text of the first descendant matching `path`, or `default` when it has none.
    pub fn find_text(&self, path: impl AsRef<str>, default: &str) -> String {
        match self.find(path.as_ref()) {
            Some(node) => node.text.clone().unwrap_or_else(|| default.to_string()),
            None => default.to_string(),
        }
    }

    /// Text of the first descendant matching `path`, if both exist.
    pub fn find_text_opt(&self, path: impl AsRef<str>) -> Option<String> {
        self.find(path.as_ref()).and_then(|node| node.text.clone())
    }

    /// Alias of [`Element::find_all`], named for parity with the Python helper.
    pub fn iter_tag(&self, tag: impl AsRef<str>) -> Vec<&Element> {
        self.find_all(tag)
    }

    /// Serialise the tree to compact XML bytes.
    pub fn to_string_bytes(&self) -> Vec<u8> {
        serialize(self)
    }

    /// Serialise the tree to indented XML.
    ///
    /// The indentation mirrors `xml.etree.ElementTree.indent`, which openpyxl applies to
    /// the parts it builds with an ElementTree rather than the streaming writer.
    pub fn to_pretty_string(&self) -> String {
        let mut namespaces: Vec<String> = Vec::new();
        collect_namespaces(self, &mut namespaces);
        let mut prefixes: BTreeMap<String, String> = BTreeMap::new();
        let mut declarations: Vec<(String, String)> = Vec::new();
        for ns in &namespaces {
            if ns == super::constants::XML_NS {
                // `xml:` is bound by the XML specification, so it is not declared.
                prefixes.insert(ns.clone(), "xml".to_string());
                continue;
            }
            let prefix = REGISTERED_PREFIXES
                .iter()
                .find(|(_, uri)| *uri == ns)
                .map(|(prefix, _)| (*prefix).to_string())
                .unwrap_or_else(|| format!("ns{}", prefixes.len()));
            declarations.push((prefix.clone(), ns.clone()));
            prefixes.insert(ns.clone(), prefix);
        }
        let mut out = String::new();
        write_pretty(&mut out, self, &prefixes, &declarations, 0, &mut false);
        out
    }
}

/// Collect every namespace URI used by the tree, in first-seen order.
fn collect_namespaces(el: &Element, out: &mut Vec<String>) {
    if let Some(ns) = namespace_of(&el.tag) {
        if !out.iter().any(|x| x == ns) {
            out.push(ns.to_string());
        }
    }
    for (key, _) in &el.attributes {
        if let Some(ns) = namespace_of(key) {
            if !out.iter().any(|x| x == ns) {
                out.push(ns.to_string());
            }
        }
    }
    for child in &el.children {
        collect_namespaces(child, out);
    }
}

/// Write one element and its subtree with indentation.
///
/// `first_child_of_line` tracks whether the next write needs to open a new line, which is
/// how ElementTree decides between `<a><b/></a>` and `<a>\n  <b/>\n</a>`.
fn write_pretty(
    out: &mut String,
    el: &Element,
    prefixes: &BTreeMap<String, String>,
    declarations: &[(String, String)],
    depth: usize,
    _is_root: &mut bool,
) {
    const INDENT: &str = "    ";
    for _ in 0..depth {
        out.push_str(INDENT);
    }
    out.push('<');
    out.push_str(&render_qname(&el.tag, prefixes, true));
    if depth == 0 {
        for (prefix, uri) in declarations {
            out.push_str(&format!(" xmlns:{prefix}=\"{}\"", escape_attribute(uri)));
        }
    }
    for (key, value) in &el.attributes {
        out.push_str(&format!(
            " {}=\"{}\"",
            render_qname(key, prefixes, false),
            escape_attribute(value)
        ));
    }
    let has_children = !el.children.is_empty();
    let text = el.text.as_deref().filter(|text| !text.is_empty());
    match (has_children, text) {
        (false, None) => out.push_str("/>"),
        (false, Some(text)) => {
            out.push('>');
            out.push_str(&escape_text(text));
            out.push_str("</");
            out.push_str(&render_qname(&el.tag, prefixes, true));
            out.push('>');
        }
        (true, text) => {
            out.push('>');
            if let Some(text) = text {
                out.push_str(&escape_text(text));
            }
            for child in &el.children {
                out.push('\n');
                write_pretty(out, child, prefixes, declarations, depth + 1, _is_root);
            }
            out.push('\n');
            for _ in 0..depth {
                out.push_str(INDENT);
            }
            out.push_str("</");
            out.push_str(&render_qname(&el.tag, prefixes, true));
            out.push('>');
        }
    }
}

fn render_qname(name: &str, prefixes: &BTreeMap<String, String>, element: bool) -> String {
    match namespace_of(name) {
        Some(ns) => {
            let local = super::constants::local_name(name);
            match prefixes.get(ns) {
                Some(prefix) => {
                    if element {
                        format!("{prefix}:{local}")
                    } else {
                        // Unprefixed attributes are not in the default namespace; an
                        // OOXML `r:id` must carry its prefix explicitly.
                        format!("{prefix}:{local}")
                    }
                }
                None => local.to_string(),
            }
        }
        None => name.to_string(),
    }
}

/// Escape text content the way `xml.sax.saxutils.escape` does.
pub fn escape_text(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}

/// Escape an attribute value the way `quoteattr` does.
pub fn escape_attribute(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\n' => out.push_str("&#10;"),
            '\r' => out.push_str("&#13;"),
            '\t' => out.push_str("&#9;"),
            _ => out.push(ch),
        }
    }
    out
}

/// A compact streaming XML writer, equivalent to `xml.sax.saxutils.XMLGenerator`.
///
/// openpyxl uses this for worksheets and the shared string table, where the output is
/// deliberately unindented. The API mirrors the Python `start_tag`/`end_tag`/`tag`
/// wrappers.
#[derive(Debug, Default)]
pub struct XmlWriter {
    out: String,
}

impl XmlWriter {
    /// Create an empty writer.
    pub fn new() -> Self {
        XmlWriter { out: String::new() }
    }

    /// Consume the writer and return the accumulated XML.
    pub fn into_string(self) -> String {
        self.out
    }

    /// Borrow the accumulated XML.
    pub fn as_str(&self) -> &str {
        &self.out
    }

    /// Append raw XML text.
    pub fn raw(&mut self, text: &str) -> &mut Self {
        self.out.push_str(text);
        self
    }

    /// Open an element.
    ///
    /// Attribute names and values are anything that renders to a string, so a caller can
    /// mix literals with computed values in one array.
    pub fn start_tag<I, K, V>(&mut self, name: &str, attrs: I) -> &mut Self
    where
        I: IntoIterator<Item = (K, V)>,
        K: std::fmt::Display,
        V: std::fmt::Display,
    {
        self.out.push('<');
        self.out.push_str(name);
        for (k, v) in attrs {
            self.out.push(' ');
            self.out.push_str(&k.to_string());
            self.out.push_str("=\"");
            self.out.push_str(&escape_attribute(&v.to_string()));
            self.out.push('"');
        }
        self.out.push('>');
        self
    }

    /// Close an element.
    pub fn end_tag(&mut self, name: &str) -> &mut Self {
        self.out.push_str("</");
        self.out.push_str(name);
        self.out.push('>');
        self
    }

    /// Write an element containing only text (`tag(doc, name, attr, body)`).
    pub fn tag<I, K, V>(&mut self, name: &str, attrs: I, body: Option<&str>) -> &mut Self
    where
        I: IntoIterator<Item = (K, V)>,
        K: std::fmt::Display,
        V: std::fmt::Display,
    {
        self.start_tag(name, attrs);
        if let Some(body) = body {
            if !body.is_empty() {
                self.out.push_str(&escape_text(body));
            }
        }
        self.end_tag(name)
    }

    /// Close the document.
    pub fn end_document(&mut self) -> &mut Self {
        self
    }
}

/// `start_tag(doc, name, attr, body)` as a free function, kept for parity with the
/// Python module API.
pub fn start_tag<I, K, V>(doc: &mut XmlWriter, name: &str, attrs: I, body: Option<&str>)
where
    I: IntoIterator<Item = (K, V)>,
    K: std::fmt::Display,
    V: std::fmt::Display,
{
    doc.start_tag(name, attrs);
    if let Some(body) = body {
        if !body.is_empty() {
            doc.raw(&escape_text(body));
        }
    }
}

/// `end_tag(doc, name)` as a free function.
pub fn end_tag(doc: &mut XmlWriter, name: &str) {
    doc.end_tag(name);
}

/// `tag(doc, name, attr, body)` as a free function.
pub fn tag<I, K, V>(doc: &mut XmlWriter, name: &str, attrs: I, body: Option<&str>)
where
    I: IntoIterator<Item = (K, V)>,
    K: std::fmt::Display,
    V: std::fmt::Display,
{
    doc.tag(name, attrs, body);
}

/// `ConditionalElement`: append a child only when `condition` holds.
///
/// When `attr` is a string, it becomes a `name="1"` attribute; otherwise no attributes
/// are emitted.
pub fn conditional_element<'a>(
    parent: &'a mut Element,
    tag: &str,
    condition: bool,
    attr: Option<&str>,
) -> Option<&'a mut Element> {
    if !condition {
        return None;
    }
    let mut child = Element::new(tag);
    if let Some(name) = attr {
        child.set(name, "1");
    }
    parent.append(child);
    parent.children.last_mut()
}

/// Parse an XML document into an [`Element`] tree (`fromstring`).
///
/// Element and attribute names are normalised to ElementTree's `{ns}local` form so the
/// same lookup helpers work for both.
/// Split an ElementTree path into its last ancestor step and its final tag.
///
/// A namespace URI inside a `{ns}tag` step contains slashes, so only a slash outside
/// braces separates two steps.
fn split_path(path: &str) -> Option<(&str, &str)> {
    let mut depth = 0i32;
    for (index, ch) in path.char_indices() {
        match ch {
            '{' => depth += 1,
            '}' => depth -= 1,
            '/' if depth == 0 => return Some((&path[..index], &path[index + 1..])),
            _ => {}
        }
    }
    None
}
/// Parse an XML document into an [`Element`] tree (`fromstring`).
///
/// Element and attribute names are normalised to ElementTree's `{ns}local` form so the
/// same lookup helpers work for both. Namespace declarations are resolved as the tree
/// is built, because `quick_xml::Reader` does not apply them itself.
pub fn fromstring(data: &[u8]) -> Result<Element> {
    let text = std::str::from_utf8(data)
        .map_err(|e| Error::Xml(format!("input is not valid UTF-8: {e}")))?;
    let mut reader = Reader::from_str(text);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<Element> = vec![Element::new("__document__")];
    // `quick_xml::Reader` does not apply namespace declarations itself, so the bindings are
    // tracked here: `bindings` holds every declaration in scope, and `scope_marks` records
    // where each open element's own declarations begin so they can be dropped on close.
    let mut bindings: Vec<(String, String)> = Vec::new();
    let mut scope_marks: Vec<usize> = vec![0];
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) => {
                scope_marks.push(bindings.len());
                let el = start_element(start, &mut bindings)?;
                stack.push(el);
            }
            Ok(Event::Empty(start)) => {
                let el = start_element(start, &mut bindings)?;
                let parent = stack.last_mut().expect("document root");
                parent.children.push(el);
            }
            Ok(Event::End(BytesEnd { .. })) => {
                if stack.len() > 1 {
                    if let Some(mark) = scope_marks.pop() {
                        bindings.truncate(mark);
                    }
                    let el = stack.pop().expect("element underflow");
                    let parent = stack.last_mut().expect("document root");
                    parent.children.push(el);
                }
            }
            Ok(Event::Text(text)) => {
                if let Some(top) = stack.last_mut() {
                    // The content is still entity-encoded, so the escapes are resolved.
                    let decoded = quick_xml::escape::unescape(&text)
                        .map_err(|e| Error::Xml(e.to_string()))?;
                    match top.text.as_mut() {
                        Some(existing) => existing.push_str(&decoded),
                        None => top.text = Some(decoded.into_owned()),
                    }
                }
            }
            Ok(Event::CData(cdata)) => {
                if let Some(top) = stack.last_mut() {
                    // CDATA content is literal, so it is copied verbatim.
                    let decoded = cdata
                        .escape()
                        .map_err(|e| Error::Xml(e.to_string()))?
                        .to_string();
                    match top.text.as_mut() {
                        Some(existing) => existing.push_str(&decoded),
                        None => top.text = Some(decoded),
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(e) => return Err(Error::Xml(e.to_string())),
        }
        buffer.clear();
    }
    // The stack holds a sentinel plus the document element; the sentinel is dropped and
    // the real root returned. Input with no element at all is malformed, not empty.
    if stack.len() == 1 {
        let mut document = stack.pop().expect("document sentinel");
        return match document.children.pop() {
            Some(root) => Ok(root),
            None => Err(Error::Xml("document has no root element".into())),
        };
    }
    Err(Error::Xml("no root element found".into()))
}

/// Build an element from a start tag, extending `bindings` with its declarations.
fn start_element(start: BytesStart<'_>, bindings: &mut Vec<(String, String)>) -> Result<Element> {
    // The raw attributes are read first because the element's own declarations must be in
    // scope before its name can be resolved.
    let mut raw: Vec<(String, String)> = Vec::new();
    for attr in start.attributes().with_checks(false) {
        let attr = attr.map_err(|e| Error::Xml(e.to_string()))?;
        let value = attr
            .normalized_value(quick_xml::XmlVersion::Implicit1_0)?
            .into_owned();
        raw.push((attr.key.as_ref().to_string(), value));
    }
    for (key, value) in &raw {
        if key == "xmlns" {
            bindings.push((String::new(), value.clone()));
        } else if let Some(prefix) = key.strip_prefix("xmlns:") {
            bindings.push((prefix.to_string(), value.clone()));
        }
    }

    let name = start.name();
    let tag = qualify(name.as_ref(), bindings, false);
    let mut el = Element::new(tag);
    for (key, value) in raw {
        if key == "xmlns" || key.starts_with("xmlns:") {
            continue;
        }
        el.attributes.push((qualify(&key, bindings, true), value));
    }
    Ok(el)
}

/// Resolve a qualified name against the bindings in scope.
///
/// An unprefixed name takes the default namespace, which is what lets a SpreadsheetML
/// document with `xmlns=".../main"` produce `{.../main}sheet` tags. An unprefixed
/// *attribute*, however, is never in the default namespace, so `is_attribute` suppresses
/// that lookup.
fn qualify(name: &str, bindings: &[(String, String)], is_attribute: bool) -> String {
    let (prefix, local) = match name.split_once(':') {
        Some((prefix, local)) => (prefix, local),
        None => ("", name),
    };
    let resolved = if is_attribute && prefix.is_empty() {
        None
    } else {
        lookup(bindings, prefix)
    };
    match resolved {
        Some(uri) => format!("{{{uri}}}{local}"),
        None => local.to_string(),
    }
}

/// Find the most recent binding for `prefix`.
fn lookup<'a>(bindings: &'a [(String, String)], prefix: &str) -> Option<&'a str> {
    if prefix == "xml" {
        return Some(super::constants::XML_NS);
    }
    bindings
        .iter()
        .rev()
        .find(|(declared, _)| declared == prefix)
        .map(|(_, uri)| uri.as_str())
}

/// Serialise an element tree to bytes without pretty-printing.
///
/// Namespace prefixes are declared on the root, matching the ElementTree behaviour the
/// reader's helpers expect.
pub fn serialize(root: &Element) -> Vec<u8> {
    let mut writer = XmlWriterInner::new(Vec::new());
    write_element_with(&mut writer, root);
    writer.into_inner()
}

fn write_element_with<W: std::io::Write>(writer: &mut XmlWriterInner<W>, el: &Element) {
    let mut start = BytesStart::new(super::constants::local_name(&el.tag).to_string());
    if let Some(ns) = namespace_of(&el.tag) {
        start = start.with_attributes([("xmlns", ns)]);
    }
    for (key, value) in &el.attributes {
        let attr_name = if let Some(ns) = namespace_of(key) {
            BytesStart::new(format!(
                "xmlns:{}",
                REGISTERED_PREFIXES
                    .iter()
                    .find(|(_, u)| *u == ns)
                    .map(|(p, _)| *p)
                    .unwrap_or("ns")
            ))
        } else {
            BytesStart::new(key.clone())
        };
        let _ = attr_name;
        start = start.with_attributes([(key.as_str(), value.as_str())]);
    }
    if el.children.is_empty() && el.text.is_none() {
        let _ = writer.write_event(Event::Empty(start));
    } else {
        let _ = writer.write_event(Event::Start(start));
        if let Some(text) = &el.text {
            // Escaping happens in `serialize`; writing the raw text avoids double-encoding.
            let _ = writer.write_event(Event::Text(BytesText::new(text)));
        }
        for child in &el.children {
            write_element_with(writer, child);
        }
        let _ = writer.write_event(Event::End(BytesEnd::new(
            super::constants::local_name(&el.tag).to_string(),
        )));
    }
}

/// Read all bytes from a reader (helper for file-like inputs).
pub fn read_all<R: Read>(mut reader: R) -> Result<Vec<u8>> {
    let mut buf = Vec::new();
    reader.read_to_end(&mut buf)?;
    Ok(buf)
}

/// Format a float the way Python's `repr()` does for the values written to XML.
///
/// Python writes the shortest representation that round-trips; Rust's default `{}` for
/// `f64` does the same but always keeps a decimal point for integral values, so the
/// integral case is special-cased to emit an integer-looking literal like `1` rather
/// than `1.0` — matching openpyxl's `repr`/`str` output for floats.
pub fn repr_float(value: f64) -> String {
    if value == value.trunc() && value.abs() < 1e16 {
        format!("{}", value as i64)
    } else {
        let mut s = format!("{value}");
        if s.contains('e') || s.contains("inf") || s.contains("NaN") {
            s = format!("{value:?}");
        }
        s
    }
}

/// Python's `"%.15g" % value`, used by `safe_string` for numeric XML payloads.
pub fn safe_string(value: f64) -> String {
    if value == 0.0 {
        return "0".to_string();
    }
    let exp = value.abs().log10().floor() as i32;
    if !(-4..15).contains(&exp) {
        let mut s = format!("{:.14e}", value);
        // Python renders exponents with at least two digits and no zero padding on
        // the mantissa beyond what is needed.
        if let Some(pos) = s.find('e') {
            let (mantissa, exponent) = s.split_at(pos);
            let mut mantissa = mantissa.trim_end_matches('0').to_string();
            if mantissa.ends_with('.') {
                mantissa.pop();
            }
            let exponent = &exponent[1..];
            let (sign, digits) = match exponent.strip_prefix('-') {
                Some(d) => ('-', d),
                None => ('+', exponent),
            };
            s = format!("{mantissa}e{sign}{:0>2}", digits);
        }
        s
    } else {
        let decimals = (14 - exp).max(0) as usize;
        let mut s = format!("{:.*}", decimals, value);
        if s.contains('.') {
            s = s.trim_end_matches('0').trim_end_matches('.').to_string();
        }
        s
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_and_find() {
        let xml =
            br#"<?xml version="1.0"?><root xmlns="urn:x"><child a="1">hi</child><child/></root>"#;
        let root = fromstring(xml).unwrap();
        assert_eq!(root.tag, "{urn:x}root");
        assert_eq!(root.find_text("{urn:x}child", ""), "hi");
        assert_eq!(root.find("{urn:x}child").unwrap().get("a"), Some("1"));
        assert_eq!(root.find_all("{urn:x}child").len(), 2);
        assert!(root.find("{urn:x}missing").is_none());
    }

    #[test]
    fn parse_namespaced_attributes() {
        let xml =
            br#"<w:workbook xmlns:w="urn:w" xmlns:r="urn:r"><w:sheet r:id="rId1"/></w:workbook>"#;
        let root = fromstring(xml).unwrap();
        let sheet = root.find("{urn:w}sheet").unwrap();
        assert_eq!(sheet.get("{urn:r}id"), Some("rId1"));
    }

    #[test]
    fn serialise_declares_namespaces() {
        let mut el = Element::new(format!("{{{SHEET_MAIN_NS}}}workbook"));
        el.append(Element::new(format!("{{{SHEET_MAIN_NS}}}sheets")));
        let out = el.to_pretty_string();
        // Every namespace in the tree is declared on the root with a generated prefix.
        assert!(out.starts_with("<s:workbook xmlns:s="));
        assert!(out.contains(SHEET_MAIN_NS));
        assert!(out.contains("<s:sheets/>"));
        assert!(out.trim_end().ends_with("</s:workbook>"));
    }

    #[test]
    fn streaming_writer_matches_python_shape() {
        let mut doc = XmlWriter::new();
        doc.start_tag("worksheet", [("xmlns", SHEET_MAIN_NS)]);
        doc.tag("dimension", [("ref", "A1:B2")], None);
        doc.start_tag("row", [("r", "1")]);
        doc.tag("c", [("r", "A1"), ("t", "s")], Some("0"));
        doc.end_tag("row");
        doc.end_tag("worksheet");
        // Python's `XMLGenerator` never self-closes an element, so `tag` writes an explicit
        // close tag rather than `/>`.
        assert_eq!(
            doc.as_str(),
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><dimension ref=\"A1:B2\"></dimension><row r=\"1\"><c r=\"A1\" t=\"s\">0</c></row></worksheet>"
        );
    }

    #[test]
    fn escaping_is_applied() {
        let mut doc = XmlWriter::new();
        doc.tag("t", [("a", "x\"&<>")], Some("a<b&c>d"));
        // Python's `quoteattr` switches to single quotes when a value contains a double
        // quote; ferroxl always uses double quotes and escapes the quote itself, which is
        // equivalent XML.
        assert_eq!(
            doc.as_str(),
            "<t a=\"x&quot;&amp;&lt;&gt;\">a&lt;b&amp;c&gt;d</t>"
        );
    }

    #[test]
    fn safe_string_matches_python_g_format() {
        assert_eq!(safe_string(0.0), "0");
        assert_eq!(safe_string(1.0), "1");
        assert_eq!(safe_string(1.5), "1.5");
        assert_eq!(safe_string(1234567890123456.0), "1.23456789012346e+15");
        assert_eq!(safe_string(0.0001), "0.0001");
    }

    #[test]
    fn repr_float_drops_trailing_dot() {
        assert_eq!(repr_float(1.0), "1");
        assert_eq!(repr_float(1.5), "1.5");
    }
}
