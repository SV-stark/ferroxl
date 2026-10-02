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
pub const RULE_ATTRIBUTES: [&str; 12] = [
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
    // `timePeriod` is what makes a rule relative to today rather than to the data: without it
    // "yesterday" and "last week" rules cannot be expressed at all.
    "timePeriod",
    "text",
];

/// The `iconSet` attributes read and written verbatim.
/// The `iconSet` attributes read and written verbatim.
pub const ICON_ATTRIBUTES: [&str; 4] = ["iconSet", "showValue", "reverse", "percent"];

/// A data bar: a horizontal bar drawn behind each cell, scaled to the range.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DataBar {
    /// The two thresholds, low then high.
    pub cfvo: Vec<Cfvo>,
    /// The bar's colour.
    pub color: String,
    /// Whether the cell's value is shown alongside the bar.
    pub show_value: Option<bool>,
    /// The shortest bar as a percentage of the cell width.
    pub min_length: Option<u32>,
    /// The longest bar as a percentage of the cell width.
    pub max_length: Option<u32>,
}

impl DataBar {
    /// A data bar between two thresholds, in `color`.
    pub fn new(start: Cfvo, end: Cfvo, color: &str) -> Self {
        DataBar {
            cfvo: vec![start, end],
            color: color.to_string(),
            show_value: None,
            min_length: None,
            max_length: None,
        }
    }

    /// Hide the cell's value, leaving only the bar.
    pub fn without_value(mut self) -> Self {
        self.show_value = Some(false);
        self
    }

    /// Set the bar lengths as percentages of the cell width.
    pub fn with_lengths(mut self, min: u32, max: u32) -> Self {
        self.min_length = Some(min);
        self.max_length = Some(max);
        self
    }

    /// The attributes for the `<dataBar/>` element.
    ///
    /// `color` is a child element rather than an attribute, which is the one place the data
    /// bar differs from an icon set's attribute handling.
    pub fn attributes(&self) -> Vec<(String, String)> {
        let mut attrs = Vec::new();
        if let Some(show) = self.show_value {
            attrs.push(("showValue".to_string(), show.to_string()));
        }
        if let Some(min) = self.min_length {
            attrs.push(("minLength".to_string(), min.to_string()));
        }
        if let Some(max) = self.max_length {
            attrs.push(("maxLength".to_string(), max.to_string()));
        }
        attrs
    }
}

impl Rule {
    /// An icon-set rule: three icons, or more, lit according to a value's position.
    ///
    /// `values` holds one `(type, value)` pair per icon, from the lowest threshold up.
    pub fn icon_set(
        icon_style: &str,
        values: &[(&str, &str)],
        show_value: Option<bool>,
        percent: Option<bool>,
        reverse: Option<bool>,
    ) -> Rule {
        let flag = |value: Option<bool>| value.map(|v| v.to_string());
        Rule {
            rule_type: "iconSet".to_string(),
            icon_set: Some(IconSet {
                cfvo: values
                    .iter()
                    .map(|(kind, value)| Cfvo::new(kind, value))
                    .collect(),
                icon_set: Some(icon_style.to_string()),
                show_value: flag(show_value),
                reverse: flag(reverse),
                percent: flag(percent),
            }),
            ..Rule::default()
        }
    }

    /// A data-bar rule: a bar behind each cell, scaled between two thresholds.
    pub fn data_bar(
        start: (&str, &str),
        end: (&str, &str),
        color: &str,
        show_value: Option<bool>,
        min_length: Option<u32>,
        max_length: Option<u32>,
    ) -> Rule {
        Rule {
            rule_type: "dataBar".to_string(),
            data_bar: Some(DataBar {
                cfvo: vec![Cfvo::new(start.0, start.1), Cfvo::new(end.0, end.1)],
                color: color.to_string(),
                show_value,
                min_length,
                max_length,
            }),
            ..Rule::default()
        }
    }
}

/// A single colour-scale stop.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Cfvo {
    /// The threshold type: `min`, `max`, `num`, `percent`, `percentile` or `formula`.
    pub cfvo_type: Option<String>,
    /// The threshold value.
    pub val: Option<String>,
    /// Whether the threshold is `>=` rather than `>`.
    ///
    /// Defaults to true in the schema, and it is the difference between an icon set that
    /// includes the boundary value and one that does not. Omitting it when false changes
    /// which rows light up.
    pub gte: Option<bool>,
}

impl Cfvo {
    /// A stop with only a type.
    pub fn with_type(cfvo_type: &str) -> Self {
        Cfvo {
            cfvo_type: Some(cfvo_type.to_string()),
            val: None,
            gte: None,
        }
    }

    /// A stop with a type and value.
    pub fn new(cfvo_type: &str, val: &str) -> Self {
        Cfvo {
            cfvo_type: Some(cfvo_type.to_string()),
            val: Some(val.to_string()),
            gte: None,
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
        // Only written when false: the schema's default is true, so an absent `gte` and an
        // explicit `gte="0"` mean different things to Excel.
        if self.gte == Some(false) {
            attrs.push(("gte".to_string(), "0".to_string()));
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
    /// Whether the thresholds are percentages rather than counts.
    ///
    /// Absent means false, which is why an icon set written with percentages came back with
    /// its thresholds reinterpreted against the row count.
    pub percent: Option<String>,
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_data_bar_rule_carries_its_thresholds_and_colour() {
        let rule = Rule::data_bar(
            ("min", "0"),
            ("max", "0"),
            "FF638EC6",
            Some(true),
            Some(10),
            Some(90),
        );
        assert_eq!(rule.rule_type, "dataBar");
        assert!(rule.is_data_bar());
        let bar = rule.data_bar.expect("a data bar");
        assert_eq!(bar.cfvo.len(), 2);
        assert_eq!(bar.cfvo[0].cfvo_type.as_deref(), Some("min"));
        assert_eq!(bar.cfvo[1].cfvo_type.as_deref(), Some("max"));
        assert_eq!(bar.color, "FF638EC6");
        assert_eq!(bar.min_length, Some(10));
        assert_eq!(bar.max_length, Some(90));
    }

    #[test]
    fn an_icon_set_rule_carries_every_threshold() {
        let rule = Rule::icon_set(
            "3TrafficLights1",
            &[("percent", "0"), ("percent", "33"), ("percent", "67")],
            Some(true),
            Some(true),
            Some(false),
        );
        assert_eq!(rule.rule_type, "iconSet");
        let set = rule.icon_set.expect("an icon set");
        assert_eq!(set.cfvo.len(), 3);
        assert_eq!(set.icon_set.as_deref(), Some("3TrafficLights1"));
        assert_eq!(set.percent.as_deref(), Some("true"));
        assert_eq!(set.reverse.as_deref(), Some("false"));
    }

    #[test]
    fn gte_is_only_written_when_it_is_false() {
        // The schema defaults `gte` to true, so an absent attribute and an explicit "0" mean
        // different things to Excel. Writing `gte="1"` everywhere would be correct but noisy;
        // writing it when true would be worse, because it would pin a default.
        let inclusive = Cfvo::new("percent", "0");
        assert!(inclusive.attributes().iter().all(|(k, _)| k != "gte"));

        let exclusive = Cfvo {
            gte: Some(false),
            ..Cfvo::new("percent", "50")
        };
        assert!(exclusive
            .attributes()
            .contains(&("gte".to_string(), "0".to_string())));
    }

    #[test]
    fn time_period_is_a_rule_attribute() {
        // Without it a "yesterday" or "last week" rule cannot be expressed at all.
        assert!(RULE_ATTRIBUTES.contains(&"timePeriod"));
    }

    #[test]
    fn percent_is_an_icon_set_attribute() {
        assert!(ICON_ATTRIBUTES.contains(&"percent"));
    }
}
