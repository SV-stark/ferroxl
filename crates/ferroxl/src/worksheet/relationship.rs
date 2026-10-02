//! Worksheet relationships (`openpyxl/worksheet/relationship.py`).

// `from_str` is openpyxl's `classmethod from_str`, so the name is kept even though Rust
// would rather these implemented `FromStr`.
#![allow(clippy::should_implement_trait)]

use crate::exceptions::{Error, Result};
use crate::xml::constants::{PKG_REL_NS, REL_NS};
use crate::xml::functions::Element;

/// The relationship kinds a worksheet can hold.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RelationshipType {
    /// An external hyperlink.
    Hyperlink,
    /// A drawing part.
    Drawing,
    /// An embedded image.
    Image,
}

impl RelationshipType {
    /// The short name used by `Relationship.TYPES`.
    pub fn as_str(self) -> &'static str {
        match self {
            RelationshipType::Hyperlink => "hyperlink",
            RelationshipType::Drawing => "drawing",
            RelationshipType::Image => "image",
        }
    }

    /// The full relationship type URI.
    pub fn uri(self) -> String {
        format!("{REL_NS}/{}", self.as_str())
    }

    /// Parse a short name.
    pub fn from_str(value: &str) -> Result<Self> {
        Ok(match value {
            "hyperlink" => RelationshipType::Hyperlink,
            "drawing" => RelationshipType::Drawing,
            "image" => RelationshipType::Image,
            other => return Err(Error::Value(format!("Invalid relationship type {other}"))),
        })
    }

    /// Whether the type URI belongs to a worksheet relationship.
    pub fn from_uri(uri: &str) -> Option<Self> {
        let short = uri.strip_prefix(REL_NS)?.strip_prefix('/')?;
        RelationshipType::from_str(short).ok()
    }
}

/// A single relationship from a worksheet part to another part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Relationship {
    /// The relationship kind.
    pub relationship_type: RelationshipType,
    /// The target path or URL.
    pub target: Option<String>,
    /// `External` for external targets.
    pub target_mode: Option<String>,
    /// The relationship id, e.g. `rId1`.
    pub id: Option<String>,
}

impl Relationship {
    /// Build a relationship.
    pub fn new(relationship_type: RelationshipType, target: Option<&str>) -> Self {
        Relationship {
            relationship_type,
            target: target.map(|t| t.to_string()),
            target_mode: None,
            id: None,
        }
    }

    /// The `Type` attribute value.
    pub fn type_uri(&self) -> String {
        self.relationship_type.uri()
    }

    /// Build a single-relationship `<Relationships>` document, matching
    /// `Relationship.__repr__`.
    pub fn to_document(&self) -> Element {
        let mut root = Element::new(format!("{{{PKG_REL_NS}}}Relationships"));
        let mut body = Element::new(format!("{{{PKG_REL_NS}}}Relationship"));
        if let Some(id) = &self.id {
            body.set("Id", id.clone());
        }
        body.set("Type", self.type_uri());
        if let Some(target) = &self.target {
            body.set("Target", target.clone());
        }
        if let Some(mode) = &self.target_mode {
            body.set("TargetMode", mode.clone());
        }
        root.append(body);
        root
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn valid_types_round_trip() {
        for t in [
            RelationshipType::Hyperlink,
            RelationshipType::Drawing,
            RelationshipType::Image,
        ] {
            assert_eq!(RelationshipType::from_str(t.as_str()).unwrap(), t);
            assert_eq!(RelationshipType::from_uri(&t.uri()), Some(t));
        }
        assert!(RelationshipType::from_str("bogus").is_err());
        assert_eq!(RelationshipType::from_uri("http://example.com/x"), None);
    }

    #[test]
    fn uri_shape() {
        assert_eq!(
            RelationshipType::Hyperlink.uri(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
        );
    }

    #[test]
    fn document_includes_attributes() {
        let mut rel = Relationship::new(RelationshipType::Hyperlink, Some("http://x"));
        rel.id = Some("rId3".to_string());
        rel.target_mode = Some("External".to_string());
        let doc = rel.to_document();
        let xml = doc.to_pretty_string();
        assert!(xml.contains("Id=\"rId3\""));
        assert!(xml.contains("Target=\"http://x\""));
        assert!(xml.contains("TargetMode=\"External\""));
        assert!(xml.contains("relationships/hyperlink"));
    }
}
