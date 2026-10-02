//! Conditional formatting (`openpyxl/formatting/`).
//!
//! Rules are grouped by the range string they apply to. Priorities are global and
//! normalised on load; differential styles (`dxf`) referenced by rules are collected into
//! the workbook's style properties so they can be written to `styles.xml`.

pub mod rules;

use std::collections::BTreeMap;

use crate::styles::borders::Borders;
use crate::styles::colors::Color;
use crate::styles::fills::Fill;
use crate::styles::fonts::Font;

pub use rules::{CellIsRule, ColorScaleRule, FormulaRule, Rule};

/// The `cfRule` attributes read and written verbatim.
pub const RULE_ATTRIBUTES: [&str; 11] = [
    "aboveAverage",
    "bottom",
    "dxfId",
    "equalAverage",
    "operator",
    "percent",
    "priority",
    "rank",
    "stdDev",
    "stopIfTrue",
    "text",
];

/// The `iconSet` attributes read and written verbatim.
pub const ICON_ATTRIBUTES: [&str; 3] = ["iconSet", "showValue", "reverse"];

/// A single colour-scale stop.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Cfvo {
    /// The threshold type: `min`, `max`, `num`, `percent`, `percentile` or `formula`.
    pub cfvo_type: Option<String>,
    /// The threshold value.
    pub val: Option<String>,
}

impl Cfvo {
    /// A stop with only a type.
    pub fn with_type(cfvo_type: &str) -> Self {
        Cfvo {
            cfvo_type: Some(cfvo_type.to_string()),
            val: None,
        }
    }

    /// A stop with a type and value.
    pub fn new(cfvo_type: &str, val: &str) -> Self {
        Cfvo {
            cfvo_type: Some(cfvo_type.to_string()),
            val: Some(val.to_string()),
        }
    }

    /// Attributes for the `<cfvo/>` element.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = Vec::new();
        if let Some(t) = &self.cfvo_type {
            attrs.push(("type".to_string(), t.clone()));
        }
        if let Some(v) = &self.val {
            attrs.push(("val".to_string(), v.clone()));
        }
        attrs
    }
}

/// The colour-scale portion of a rule.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ColorScale {
    /// The threshold values.
    pub cfvo: Vec<Cfvo>,
    /// The colours for each threshold.
    pub color: Vec<Color>,
}

/// The icon-set portion of a rule.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct IconSet {
    /// The threshold values.
    pub cfvo: Vec<Cfvo>,
    /// The icon set name.
    pub icon_set: Option<String>,
    /// Whether to show values.
    pub show_value: Option<String>,
    /// Whether to reverse the icon order.
    pub reverse: Option<String>,
}

/// A differential style used by a rule (`dxf`).
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DxfStyle {
    /// Font override.
    pub font: Option<Font>,
    /// Fill override.
    pub fill: Option<Fill>,
    /// Border override.
    pub border: Option<Borders>,
}

/// The rules applied to one range string.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ConditionalFormatting {
    /// Range string → rules, preserving insertion order.
    pub cf_rules: BTreeMap<String, Vec<Rule>>,
    /// The highest priority assigned so far.
    pub max_priority: u32,
    /// Rules collected during a parse, before priority normalisation.
    pub parse_rules: BTreeMap<String, Vec<Rule>>,
}

impl ConditionalFormatting {
    /// An empty rule set.
    pub fn new() -> Self {
        ConditionalFormatting::default()
    }

    /// Add a rule to a range, assigning the next priority.
    pub fn add(&mut self, range_string: &str, mut rule: Rule) {
        self.max_priority += 1;
        rule.attributes
            .insert("priority".to_string(), self.max_priority.to_string());
        self.cf_rules
            .entry(range_string.to_string())
            .or_default()
            .push(rule);
    }

    /// Replace the rules from a dictionary, as used when loading a document.
    ///
    /// Priorities are then renumbered globally in ascending order.
    pub fn update(&mut self, cf_rules: BTreeMap<String, Vec<Rule>>) {
        for (range_string, rules) in cf_rules {
            let entry = self.cf_rules.entry(range_string).or_default();
            entry.extend(rules);
        }
        self.max_priority = 0;
        let mut priorities: Vec<u32> = self
            .cf_rules
            .values()
            .flatten()
            .filter_map(|rule| rule.attributes.get("priority").and_then(|p| p.parse().ok()))
            .collect();
        priorities.sort_unstable();
        for rules in self.cf_rules.values_mut() {
            for rule in rules.iter_mut() {
                let Some(priority) = rule
                    .attributes
                    .get("priority")
                    .and_then(|p| p.parse::<u32>().ok())
                else {
                    continue;
                };
                let index = priorities.iter().position(|p| *p == priority).unwrap_or(0);
                let new_priority = (index + 1) as u32;
                rule.attributes
                    .insert("priority".to_string(), new_priority.to_string());
                if new_priority > self.max_priority {
                    self.max_priority = new_priority;
                }
            }
        }
    }

    /// Collect the differential styles referenced by the rules into `dxf_list`.
    ///
    /// The styles are appended in rule order and each rule's `dxfId` is rewritten to the
    /// assigned index. Returns the collected styles so the caller can store them on the
    /// workbook.
    pub fn collect_dxf_styles(&mut self, dxf_list: &mut Vec<DxfStyle>) {
        for rules in self.cf_rules.values_mut() {
            for rule in rules.iter_mut() {
                let Some(dxf) = rule.dxf.take() else {
                    continue;
                };
                let mut filtered = DxfStyle::default();
                if dxf.font.is_some() {
                    filtered.font = dxf.font;
                }
                if dxf.border.is_some() {
                    filtered.border = dxf.border;
                }
                if dxf.fill.is_some() {
                    filtered.fill = dxf.fill;
                }
                dxf_list.push(filtered);
                let index = dxf_list.len() - 1;
                rule.attributes
                    .insert("dxfId".to_string(), index.to_string());
            }
        }
    }

    /// All rules across all ranges.
    pub fn all_rules(&self) -> impl Iterator<Item = &Rule> {
        self.cf_rules.values().flatten()
    }
}

/// The style properties that the reader produces and the writer consumes.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct StyleProperties {
    /// Indexed colour palette from `styles.xml`.
    pub color_index: Vec<String>,
    /// Differential styles referenced by conditional formatting.
    pub dxf_list: Vec<DxfStyle>,
}
