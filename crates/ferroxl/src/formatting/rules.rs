//! Conditional formatting rule constructors (`openpyxl/formatting/rules.py`).

use std::collections::BTreeMap;

use crate::formatting::{ColorScale, DataBar, DxfStyle, IconSet};
use crate::styles::borders::Borders;
use crate::styles::colors::Color;
use crate::styles::fills::Fill;
use crate::styles::fonts::Font;

/// A conditional formatting rule.
///
/// The Python class is a dictionary wrapper restricted to a known key set. Rust models it
/// as a struct with an ordered attribute map plus typed payloads, which serialises the
/// same XML while keeping the data inspectable.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Rule {
    /// The rule type: `expression`, `cellIs`, `colorScale`, `iconSet`, …
    pub rule_type: String,
    /// Rule attributes written to `<cfRule>`, in insertion order.
    pub attributes: BTreeMap<String, String>,
    /// Formula expressions.
    pub formula: Vec<String>,
    /// Colour-scale payload.
    pub color_scale: Option<ColorScale>,
    /// Icon-set payload.
    pub icon_set: Option<IconSet>,
    /// Data-bar payload.
    pub data_bar: Option<DataBar>,
    /// Differential style, replaced by `dxfId` once written.
    pub dxf: Option<DxfStyle>,
}

impl Rule {
    /// Build a rule of the given type.
    pub fn new(rule_type: &str) -> Self {
        Rule {
            rule_type: rule_type.to_string(),
            ..Rule::default()
        }
    }

    /// Set a rule attribute.
    pub fn with_attribute(mut self, key: &str, value: &str) -> Self {
        self.attributes.insert(key.to_string(), value.to_string());
        self
    }

    /// Set the priority attribute.
    pub fn with_priority(self, priority: u32) -> Self {
        self.with_attribute("priority", &priority.to_string())
    }

    /// Set the differential style.
    pub fn with_dxf(mut self, dxf: DxfStyle) -> Self {
        self.dxf = Some(dxf);
        self
    }

    /// Whether this is a data-bar rule.
    ///
    /// Kept because it names the one payload that is not a colour scale or an icon set, and a
    /// caller walking a rule list wants to ask that without matching on a string.
    pub fn is_data_bar(&self) -> bool {
        self.rule_type == "dataBar"
    }
}

/// A conditional formatting rule based on a colour scale.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ColorScaleRule {
    /// Start threshold type.
    pub start_type: Option<String>,
    /// Start threshold value.
    pub start_value: Option<String>,
    /// Start colour.
    pub start_color: Option<Color>,
    /// Mid threshold type.
    pub mid_type: Option<String>,
    /// Mid threshold value.
    pub mid_value: Option<String>,
    /// Mid colour.
    pub mid_color: Option<Color>,
    /// End threshold type.
    pub end_type: Option<String>,
    /// End threshold value.
    pub end_value: Option<String>,
    /// End colour.
    pub end_color: Option<Color>,
}

/// The threshold types a colour-scale rule accepts.
pub const COLOR_SCALE_VALID_TYPES: [&str; 6] =
    ["min", "max", "num", "percent", "percentile", "formula"];

impl ColorScaleRule {
    /// Build a rule from its nine components.
    ///
    /// The three thresholds are the rule's whole shape, so they are passed positionally
    /// rather than as a struct, matching openpyxl's `ColorScaleRule` signature.
    #[allow(clippy::too_many_arguments)]
    pub fn new(
        start_type: Option<&str>,
        start_value: Option<&str>,
        start_color: Option<Color>,
        mid_type: Option<&str>,
        mid_value: Option<&str>,
        mid_color: Option<Color>,
        end_type: Option<&str>,
        end_value: Option<&str>,
        end_color: Option<Color>,
    ) -> Self {
        ColorScaleRule {
            start_type: start_type.map(|t| t.to_string()),
            start_value: start_value.map(|v| v.to_string()),
            start_color,
            mid_type: mid_type.map(|t| t.to_string()),
            mid_value: mid_value.map(|v| v.to_string()),
            mid_color,
            end_type: end_type.map(|t| t.to_string()),
            end_value: end_value.map(|v| v.to_string()),
            end_color,
        }
    }

    /// The threshold stops, skipping unset positions.
    pub fn cfvo(&self) -> Vec<crate::formatting::Cfvo> {
        let mut vals = Vec::new();
        for (t, v) in [
            (&self.start_type, &self.start_value),
            (&self.mid_type, &self.mid_value),
            (&self.end_type, &self.end_value),
        ] {
            let Some(t) = t else { continue };
            vals.push(match v {
                Some(v) => crate::formatting::Cfvo::new(t, v),
                None => crate::formatting::Cfvo::with_type(t),
            });
        }
        vals
    }

    /// The start, mid and end colours that are set.
    pub fn colors(&self) -> Vec<Color> {
        [
            self.start_color.clone(),
            self.mid_color.clone(),
            self.end_color.clone(),
        ]
        .into_iter()
        .flatten()
        .collect()
    }

    /// Convert to the serialisable rule form.
    pub fn to_rule(&self) -> Rule {
        Rule {
            rule_type: "colorScale".to_string(),
            color_scale: Some(ColorScale {
                cfvo: self.cfvo(),
                color: self.colors(),
            }),
            ..Rule::default()
        }
    }
}

impl Default for ColorScaleRule {
    fn default() -> Self {
        ColorScaleRule::new(None, None, None, None, None, None, None, None, None)
    }
}

/// A conditional formatting rule based on a formula.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct FormulaRule {
    /// The formula(s) to evaluate.
    pub formula: Option<String>,
    /// Whether later rules are skipped when this one matches.
    pub stop_if_true: bool,
    /// Differential font.
    pub font: Option<Font>,
    /// Differential border.
    pub border: Option<Borders>,
    /// Differential fill.
    pub fill: Option<Fill>,
}

impl FormulaRule {
    /// Build a formula rule.
    pub fn new(formula: Option<&str>, stop_if_true: bool) -> Self {
        FormulaRule {
            formula: formula.map(|f| f.to_string()),
            stop_if_true,
            ..FormulaRule::default()
        }
    }

    /// Attach a differential font.
    pub fn with_font(mut self, font: Font) -> Self {
        self.font = Some(font);
        self
    }

    /// Attach a differential fill.
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Attach a differential border.
    pub fn with_border(mut self, border: Borders) -> Self {
        self.border = Some(border);
        self
    }

    /// Convert to the serialisable rule form.
    pub fn to_rule(&self) -> Rule {
        let mut rule = Rule::new("expression");
        rule.formula = self.formula.iter().cloned().collect();
        rule.dxf = Some(DxfStyle {
            font: self.font.clone(),
            border: self.border.clone(),
            fill: self.fill.clone(),
        });
        if self.stop_if_true {
            rule.attributes
                .insert("stopIfTrue".to_string(), "1".to_string());
        }
        rule
    }
}

/// A conditional formatting rule based on cell contents.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CellIsRule {
    /// The comparison operator, in OOXML form.
    pub operator: Option<String>,
    /// The comparison formula.
    pub formula: Option<String>,
    /// Whether later rules are skipped when this one matches.
    pub stop_if_true: bool,
    /// Differential font.
    pub font: Option<Font>,
    /// Differential border.
    pub border: Option<Borders>,
    /// Differential fill.
    pub fill: Option<Fill>,
}

impl CellIsRule {
    /// Build a cell-is rule.
    ///
    /// Symbolic operators (`>`, `>=`, `<`, `<=`, `=`, `==`, `!=`) are expanded to their
    /// OOXML equivalents, as `CellIsRule.operator` does.
    pub fn new(operator: Option<&str>, formula: Option<&str>, stop_if_true: bool) -> Self {
        CellIsRule {
            operator: operator.map(expand_operator).map(|o| o.to_string()),
            formula: formula.map(|f| f.to_string()),
            stop_if_true,
            ..CellIsRule::default()
        }
    }

    /// Attach a differential font.
    pub fn with_font(mut self, font: Font) -> Self {
        self.font = Some(font);
        self
    }

    /// Attach a differential fill.
    pub fn with_fill(mut self, fill: Fill) -> Self {
        self.fill = Some(fill);
        self
    }

    /// Attach a differential border.
    pub fn with_border(mut self, border: Borders) -> Self {
        self.border = Some(border);
        self
    }

    /// Convert to the serialisable rule form.
    pub fn to_rule(&self) -> Rule {
        let mut rule = Rule::new("cellIs");
        if let Some(operator) = &self.operator {
            rule.attributes
                .insert("operator".to_string(), operator.clone());
        }
        rule.formula = self.formula.iter().cloned().collect();
        rule.dxf = Some(DxfStyle {
            font: self.font.clone(),
            border: self.border.clone(),
            fill: self.fill.clone(),
        });
        if self.stop_if_true {
            rule.attributes
                .insert("stopIfTrue".to_string(), "1".to_string());
        }
        rule
    }
}

/// The symbolic → OOXML operator map used by `CellIsRule`.
pub const OPERATOR_EXPANSION: [(&str, &str); 7] = [
    (">", "greaterThan"),
    (">=", "greaterThanOrEqual"),
    ("<", "lessThan"),
    ("<=", "lessThanOrEqual"),
    ("=", "equal"),
    ("==", "equal"),
    ("!=", "notEqual"),
];

fn expand_operator(operator: &str) -> &str {
    OPERATOR_EXPANSION
        .iter()
        .find(|(symbol, _)| *symbol == operator)
        .map(|(_, expanded)| *expanded)
        .unwrap_or(operator)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::styles::colors::Color;

    #[test]
    fn color_scale_rule_shape() {
        let rule = ColorScaleRule::new(
            Some("min"),
            None,
            Some(Color::new("FFFF7128")),
            Some("percentile"),
            Some("50"),
            Some(Color::new("FFFFEF9C")),
            Some("max"),
            None,
            Some(Color::new("FF4F81BD")),
        );
        let cfvo = rule.cfvo();
        assert_eq!(cfvo.len(), 3);
        assert_eq!(cfvo[0].cfvo_type.as_deref(), Some("min"));
        assert_eq!(cfvo[0].val, None);
        assert_eq!(cfvo[1].cfvo_type.as_deref(), Some("percentile"));
        assert_eq!(cfvo[1].val.as_deref(), Some("50"));
        assert_eq!(cfvo[2].cfvo_type.as_deref(), Some("max"));
        assert_eq!(rule.colors().len(), 3);

        let serialised = rule.to_rule();
        assert_eq!(serialised.rule_type, "colorScale");
        let scale = serialised.color_scale.unwrap();
        assert_eq!(scale.cfvo.len(), 3);
        assert_eq!(scale.color.len(), 3);
    }

    #[test]
    fn colour_scale_skips_unset_midpoint() {
        let rule = ColorScaleRule::new(
            Some("min"),
            None,
            Some(Color::new("FFFF7128")),
            None,
            None,
            None,
            Some("max"),
            None,
            Some(Color::new("FF4F81BD")),
        );
        assert_eq!(rule.cfvo().len(), 2);
        assert_eq!(rule.colors().len(), 2);
    }

    #[test]
    fn formula_rule_shape() {
        let rule = FormulaRule::new(Some("A1>5"), true)
            .with_font(Font::new().with_bold(true))
            .to_rule();
        assert_eq!(rule.rule_type, "expression");
        assert_eq!(rule.formula, vec!["A1>5".to_string()]);
        assert_eq!(rule.attributes.get("stopIfTrue").unwrap(), "1");
        assert!(rule.dxf.unwrap().font.is_some());
    }

    #[test]
    fn cell_is_rule_expands_operators() {
        for (symbol, expected) in OPERATOR_EXPANSION {
            let rule = CellIsRule::new(Some(symbol), Some("5"), false).to_rule();
            assert_eq!(rule.attributes.get("operator").unwrap(), expected);
        }
        let passthrough = CellIsRule::new(Some("between"), Some("1"), false).to_rule();
        assert_eq!(passthrough.attributes.get("operator").unwrap(), "between");
    }

    #[test]
    fn rule_defaults_and_helpers() {
        let rule = Rule::new("cellIs")
            .with_priority(3)
            .with_attribute("operator", "equal");
        assert_eq!(rule.rule_type, "cellIs");
        assert_eq!(rule.attributes.get("priority").unwrap(), "3");
        assert!(!rule.is_data_bar());
        assert!(Rule::new("dataBar").is_data_bar());
    }
}
