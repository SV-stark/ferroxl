//! Read-only tools: everything an agent needs to understand a file before editing it.

use ferroxl::CellValue;
use serde_json::{json, Value};

use super::{cells_in_range, err, open, sheet_index, unknown_sheet, used_range, Handled};
use crate::tools::Args;
use crate::values;
use crate::workspace::Workspace;

/// List every sheet with a short description of what it holds.
pub fn list_sheets(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let active = workbook.active();
    let sheets: Vec<Value> = workbook
        .worksheets
        .iter()
        .map(|sheet| {
            json!({
                "index": sheet.index,
                "name": sheet.title,
                "active": sheet.index == active,
                "state": sheet.sheet_state,
                "dimension": sheet.calculate_dimension().unwrap_or_default(),
                "highest_row": sheet.highest_row(),
                "highest_column": sheet.highest_column(),
                "cells": sheet.cell_count(),
                "formulas": count_kind(&sheet_cells(sheet), |value| {
                    matches!(value, CellValue::Formula(_))
                }),
                "merged_ranges": sheet.merged_cells(),
            })
        })
        .collect();
    let summary = format!(
        "{} has {} sheet{}: {}",
        args.require_str("path").unwrap_or_default(),
        sheets.len(),
        if sheets.len() == 1 { "" } else { "s" },
        sheets
            .iter()
            .map(|sheet| sheet["name"].as_str().unwrap_or("?").to_string())
            .collect::<Vec<_>>()
            .join(", ")
    );
    Ok((summary, json!({ "sheets": sheets })))
}

/// Describe one sheet in enough detail to plan an edit.
pub fn describe_sheet(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let cells = sheet_cells(sheet);
    let numeric = count_kind(&cells, |value| matches!(value, CellValue::Number(_)));
    let text = count_kind(&cells, |value| matches!(value, CellValue::Text(_)));
    let boolean = count_kind(&cells, |value| matches!(value, CellValue::Bool(_)));
    let temporal = count_kind(&cells, |value| {
        matches!(
            value,
            CellValue::Date(_)
                | CellValue::DateTime(_)
                | CellValue::Time(_)
                | CellValue::Duration(_)
        )
    });
    let formulas = count_kind(&cells, |value| matches!(value, CellValue::Formula(_)));
    let hyperlinks = sheet.cells_with_hyperlinks().count();

    let validations: Vec<Value> = sheet
        .data_validations
        .iter()
        .map(|rule| {
            json!({
                "type": rule.validation_type.as_str(),
                "operator": rule.operator.map(|o| o.as_str()),
                "formula1": rule.formula1,
                "formula2": rule.formula2,
                "cells": rule.cells,
                "ranges": rule.ranges,
                "allow_blank": rule.allow_blank,
            })
        })
        .collect();
    let conditional: Vec<Value> = sheet
        .conditional_formatting
        .cf_rules
        .iter()
        .map(|(range, rules)| {
            json!({
                "range": range,
                "rules": rules
                    .iter()
                    .map(|rule| json!({ "type": rule.rule_type, "priority": rule.attributes.get("priority") }))
                    .collect::<Vec<_>>(),
            })
        })
        .collect();

    let payload = json!({
        "name": sheet.title,
        "index": sheet.index,
        "state": sheet.sheet_state,
        "dimension": sheet.calculate_dimension().unwrap_or_default(),
        "highest_row": sheet.highest_row(),
        "highest_column": sheet.highest_column(),
        "populated_cells": cells.len(),
        "by_type": {
            "number": numeric,
            "text": text,
            "boolean": boolean,
            "date_or_time": temporal,
            "formula": formulas,
        },
        "hyperlinks": hyperlinks,
        "merged_ranges": sheet.merged_cells(),
        "freeze_panes": sheet.freeze_panes,
        "autofilter": sheet.auto_filter.reference(),
        "row_dimensions": sheet.row_dimensions.len(),
        "column_dimensions": sheet.column_dimensions.len(),
        "data_validations": validations,
        "conditional_formats": conditional,
        "comments": sheet.comment_count(),
        "charts": sheet.charts.len(),
        "images": sheet.images.len(),
    });
    let summary = format!(
        "{:?} is a {} sheet with {} populated cells over {}",
        sheet.title,
        sheet.calculate_dimension().unwrap_or_default(),
        cells.len(),
        cells.len()
    );
    Ok((summary, payload))
}

/// Read a rectangle of cells.
pub fn read_cells(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let range = used_range(sheet, args)?;
    let include_empty = args.opt_bool("include_empty", true);
    let cells = cells_in_range(sheet, &range)?;

    let mut rows: Vec<Value> = Vec::new();
    let mut included = 0usize;
    for (coordinate, value) in cells {
        let empty = matches!(value, CellValue::None);
        if empty && !include_empty {
            continue;
        }
        if !empty {
            included += 1;
        }
        rows.push(json!({
            "cell": coordinate,
            "value": values::to_json(&value),
            "display": values::to_display(&value),
            "number_format": sheet.number_format(&coordinate),
        }));
    }
    let summary = format!(
        "{} in {}: {included} populated cell{}",
        range,
        sheet.title,
        if included == 1 { "" } else { "s" }
    );
    Ok((
        summary,
        json!({
            "sheet": sheet.title,
            "range": range,
            "cells": rows,
        }),
    ))
}

/// List the formula cells in a range.
pub fn read_formulas(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let range = used_range(sheet, args)?;
    let mut formulas = Vec::new();
    for (coordinate, value) in cells_in_range(sheet, &range)? {
        if let CellValue::Formula(formula) = value {
            formulas.push(json!({ "cell": coordinate, "formula": formula }));
        }
    }
    let summary = format!(
        "{} formula{} in {}",
        formulas.len(),
        if formulas.len() == 1 { "" } else { "s" },
        sheet.title
    );
    Ok((
        summary,
        json!({ "sheet": sheet.title, "range": range, "formulas": formulas }),
    ))
}

/// What feeds a cell, in the order a recalculation would visit it.
pub fn trace_precedents(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let cell = args.require_str("cell")?;
    let precedents = sheet.trace_precedents(&cell).map_err(|e| e.to_string())?;
    let summary = if precedents.is_empty() {
        format!("{} has no precedents in {}", cell, sheet.title)
    } else {
        format!(
            "{} in {} is fed by {} cell{}: {}",
            cell,
            sheet.title,
            precedents.len(),
            if precedents.len() == 1 { "" } else { "s" },
            join_capped(&precedents)
        )
    };
    Ok((
        summary,
        json!({ "sheet": sheet.title, "cell": cell, "precedents": precedents }),
    ))
}

/// Every formula that would go stale if a cell changed.
pub fn trace_dependents(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let cell = args.require_str("cell")?;
    let dependents = sheet.trace_dependents(&cell).map_err(|e| e.to_string())?;
    let summary = if dependents.is_empty() {
        format!("Nothing in {} depends on {}", sheet.title, cell)
    } else {
        format!(
            "Changing {} in {} affects {} formula{}: {}",
            cell,
            sheet.title,
            dependents.len(),
            if dependents.len() == 1 { "" } else { "s" },
            join_capped(&dependents)
        )
    };
    Ok((
        summary,
        json!({ "sheet": sheet.title, "cell": cell, "dependents": dependents }),
    ))
}

/// Every circular reference in a sheet, as closed paths.
pub fn check_circular_references(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let cycles = sheet.circular_references().map_err(|e| e.to_string())?;
    let summary = if cycles.is_empty() {
        format!("No circular references in {}", sheet.title)
    } else {
        format!(
            "{} circular reference{} in {}: {}",
            cycles.len(),
            if cycles.len() == 1 { "" } else { "s" },
            sheet.title,
            cycles
                .iter()
                .map(|cycle| cycle.join(" -> "))
                .collect::<Vec<_>>()
                .join("; ")
        )
    };
    Ok((summary, json!({ "sheet": sheet.title, "cycles": cycles })))
}

/// Join coordinates for a summary line without letting a large graph fill the context.
fn join_capped(cells: &[String]) -> String {
    const SHOWN: usize = 12;
    if cells.len() <= SHOWN {
        return cells.join(", ");
    }
    format!(
        "{} and {} more",
        cells[..SHOWN].join(", "),
        cells.len() - SHOWN
    )
}

/// Search a sheet for text.
pub fn search_values(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let query = args.require_str("query")?;
    let case_sensitive = args.opt_bool("case_sensitive", false);
    let match_formulas = args.opt_bool("match_formulas", true);
    let limit = args
        .opt_usize("max_results")
        .unwrap_or(100)
        .clamp(1, 10_000);
    let range = used_range(sheet, args)?;

    let needle = if case_sensitive {
        query.clone()
    } else {
        query.to_lowercase()
    };
    let mut matches = Vec::new();
    let mut truncated = false;
    for (coordinate, value) in cells_in_range(sheet, &range)? {
        let haystacks: Vec<String> = match &value {
            CellValue::Formula(formula) if match_formulas => vec![formula.clone()],
            other => vec![values::to_display(other)],
        };
        let hit = haystacks.iter().any(|text| {
            if case_sensitive {
                text.contains(&needle)
            } else {
                text.to_lowercase().contains(&needle)
            }
        });
        if !hit {
            continue;
        }
        if matches.len() == limit {
            truncated = true;
            break;
        }
        matches.push(json!({
            "cell": coordinate,
            "value": values::to_json(&value),
        }));
    }
    let summary = format!(
        "{} match{} for {query:?} in {}",
        matches.len(),
        if matches.len() == 1 { "" } else { "es" },
        sheet.title
    );
    Ok((
        summary,
        json!({
            "sheet": sheet.title,
            "range": range,
            "query": query,
            "truncated": truncated,
            "matches": matches,
        }),
    ))
}

/// Count and measure a range.
pub fn summarize_range(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let range = used_range(sheet, args)?;
    let cells = cells_in_range(sheet, &range)?;

    let mut numeric = Vec::new();
    let mut blanks = 0usize;
    let mut text = 0usize;
    let mut booleans = 0usize;
    let mut temporals = 0usize;
    for (_, value) in &cells {
        match value {
            CellValue::Number(number) if number.is_finite() => numeric.push(*number),
            CellValue::None => blanks += 1,
            // An empty string is a blank, not a piece of text.
            CellValue::Text(text) if text.is_empty() => blanks += 1,
            CellValue::Text(_) => text += 1,
            CellValue::Bool(_) => booleans += 1,
            CellValue::Date(_)
            | CellValue::DateTime(_)
            | CellValue::Time(_)
            | CellValue::Duration(_) => temporals += 1,
            // A non-finite serial has no useful aggregate, so it is left out entirely.
            _ => {}
        }
    }
    // A range with no numbers in it has nothing to aggregate, so every aggregate is
    // reported as null rather than as a sum of zero — a zero would read as "the total is
    // zero" when the truth is "there are no numbers here".
    let sum = if numeric.is_empty() {
        None
    } else {
        Some(numeric.iter().sum::<f64>())
    };
    let mean = sum.map(|sum| sum / numeric.len() as f64);
    let payload = json!({
        "sheet": sheet.title,
        "range": range,
        "cells": cells.len(),
        "blanks": blanks,
        "text": text,
        "booleans": booleans,
        "dates_or_times": temporals,
        "numeric": numeric.len(),
        "sum": sum,
        "mean": mean,
        "min": numeric.iter().copied().reduce(f64::min),
        "max": numeric.iter().copied().reduce(f64::max),
    });
    let summary = match sum {
        Some(sum) => format!("{range}: {} numeric, sum {sum}", numeric.len()),
        None => format!("{range}: {} cells, none of them numeric", cells.len()),
    };
    Ok((summary, payload))
}

/// List the cell comments in a sheet.
pub fn list_comments(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let mut comments: Vec<(u32, Value)> = sheet
        .cells()
        .filter_map(|cell| {
            let comment = cell.comment.as_ref()?;
            let column = cell.column_index().ok()?;
            Some((
                column,
                json!({
                    "cell": cell.coordinate(),
                    "author": comment.author(),
                    "text": comment.text(),
                }),
            ))
        })
        .collect();
    comments.sort_by_key(|(column, _)| *column);
    let rows: Vec<Value> = comments.into_iter().map(|(_, entry)| entry).collect();
    let summary = format!(
        "{} comment{} in {}",
        rows.len(),
        if rows.len() == 1 { "" } else { "s" },
        sheet.title
    );
    Ok((summary, json!({ "sheet": sheet.title, "comments": rows })))
}

/// List the workbook's defined names.
pub fn list_named_ranges(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let titles = workbook.get_sheet_names();
    let names: Vec<Value> = workbook
        .get_named_ranges()
        .iter()
        .map(|defined| match defined {
            ferroxl::DefinedName::Range(range) => {
                let destinations: Vec<Value> = range
                    .destinations
                    .iter()
                    .map(|(index, reference)| {
                        json!({
                            "sheet": titles.get(*index).cloned().unwrap_or_default(),
                            "range": reference,
                        })
                    })
                    .collect();
                json!({
                    "name": range.name,
                    "kind": "range",
                    "scope": range
                        .scope
                        .and_then(|index| titles.get(index).cloned()),
                    "destinations": destinations,
                })
            }
            ferroxl::DefinedName::Value(value) => json!({
                "name": value.name,
                "kind": "value",
                "scope": value.scope.and_then(|index| titles.get(index).cloned()),
                "value": value.value,
            }),
        })
        .collect();
    let summary = format!("{} defined name(s)", names.len());
    Ok((summary, json!({ "names": names })))
}

/// Render a range as CSV.
pub fn export_csv(workspace: &Workspace, args: &Args) -> Handled {
    let workbook = open(workspace, args)?;
    let index = sheet_index(&workbook, args)?;
    let sheet = &workbook.worksheets[index];
    let range = used_range(sheet, args)?;
    let rows = sheet.range_values(&range).map_err(err)?;
    let mut csv = String::new();
    for (number, row) in rows.iter().enumerate() {
        if number > 0 {
            csv.push_str("\r\n");
        }
        let line: Vec<String> = row
            .iter()
            .map(|value| values::csv_field(&values::to_display(value)))
            .collect();
        csv.push_str(&line.join(","));
    }
    let summary = format!(
        "{} row{} of {range} as CSV",
        rows.len(),
        if rows.len() == 1 { "" } else { "s" }
    );
    Ok((
        summary,
        json!({
            "sheet": sheet.title,
            "range": range,
            "row_count": rows.len(),
            "csv": csv,
        }),
    ))
}

/// The values of a sheet's populated cells, in a stable order.
fn sheet_cells(sheet: &ferroxl::Worksheet) -> Vec<CellValue> {
    let mut cells: Vec<CellValue> = sheet
        .cells()
        .map(|cell| {
            sheet
                .cell_value(&cell.coordinate())
                .unwrap_or(CellValue::None)
        })
        // A cell exists but carries nothing — a merged cell's follower, for instance — and
        // that is a blank rather than an empty piece of text.
        .filter(|value| !is_blank(value))
        .collect();
    // `cells()` iterates a hash map, so the counts are the same either way but a stable
    // order keeps the tool output reproducible.
    cells.sort_by_key(|value| match value {
        CellValue::Formula(formula) => formula.clone(),
        other => values::to_display(other),
    });
    cells
}

/// Count the values that satisfy `predicate`.
fn count_kind(cells: &[CellValue], predicate: impl Fn(&CellValue) -> bool) -> usize {
    cells.iter().filter(|value| predicate(value)).count()
}

/// Whether a cell exists but holds nothing worth counting.
fn is_blank(value: &CellValue) -> bool {
    match value {
        CellValue::None => true,
        CellValue::Text(text) => text.is_empty(),
        _ => false,
    }
}

/// Report an unreadable sheet in a form the model can act on.
#[allow(dead_code)]
fn sheet_error(workbook: &ferroxl::Workbook, name: &str) -> String {
    unknown_sheet(workbook, name)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::testing;
    use crate::tools;

    #[test]
    fn list_sheets_reports_names_and_dimensions() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "sheet": "Numbers",
        }));
        let (summary, payload) = list_sheets(&workspace, &args).unwrap();
        assert!(summary.contains("2 sheets"), "{summary}");
        assert_eq!(payload["sheets"][1]["name"], json!("Numbers"));
        assert_eq!(payload["sheets"][1]["dimension"], json!("A1:B2"));
        assert_eq!(payload["sheets"][1]["cells"], json!(4));
    }

    #[test]
    fn describe_sheet_counts_each_kind_of_value() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx" }));
        let (_, payload) = describe_sheet(&workspace, &args).unwrap();
        assert_eq!(payload["by_type"]["number"], json!(3));
        assert_eq!(payload["by_type"]["text"], json!(2));
        assert_eq!(payload["by_type"]["formula"], json!(1));
        assert_eq!(payload["merged_ranges"], json!(["E1:F1"]));
    }

    #[test]
    fn trace_precedents_returns_what_feeds_a_cell() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "chain.xlsx", "cell": "A5" }));
        let (summary, payload) = trace_precedents(&workspace, &args).unwrap();
        // A5 reads A4, A4 reads A1:A3, so the numbers come before the formula that used them.
        assert_eq!(payload["precedents"], json!(["A1", "A2", "A3", "A4"]));
        assert!(summary.contains("A1, A2, A3, A4"), "{summary}");
        // The cell asked about is not one of its own precedents, which the payload above
        // already shows; the summary names it only as the subject of the question.
    }

    #[test]
    fn trace_precedents_says_so_when_there_are_none() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "chain.xlsx", "cell": "A1" }));
        let (summary, payload) = trace_precedents(&workspace, &args).unwrap();
        assert_eq!(payload["precedents"], json!([]));
        assert!(summary.contains("no precedents"), "{summary}");
    }

    #[test]
    fn trace_dependents_covers_the_whole_transitive_closure() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "chain.xlsx", "cell": "A1" }));
        let (summary, payload) = trace_dependents(&workspace, &args).unwrap();
        let affected: Vec<&str> = payload["dependents"]
            .as_array()
            .unwrap()
            .iter()
            .map(|c| c.as_str().unwrap())
            .collect();
        // A5 does not read A1, but it reads A4, which does. Everything downstream breaks.
        assert_eq!(affected, ["A4", "B1", "A5"]);
        assert!(summary.contains("3 formulas"), "{summary}");
    }

    #[test]
    fn a_cycle_is_reported_as_a_closed_path() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "cycle.xlsx" }));
        let (summary, payload) = check_circular_references(&workspace, &args).unwrap();
        let cycles = payload["cycles"].as_array().unwrap();
        assert_eq!(cycles.len(), 1, "{payload}");
        let cells: Vec<&str> = cycles[0]
            .as_array()
            .unwrap()
            .iter()
            .map(|c| c.as_str().unwrap())
            .collect();
        assert_eq!(cells.len(), 4, "three cells plus the return: {cells:?}");
        assert_eq!(cells.first(), cells.last(), "the path closes: {cells:?}");
        assert!(summary.contains("->"), "{summary}");
    }

    #[test]
    fn a_clean_sheet_reports_no_cycles() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "chain.xlsx" }));
        let (summary, payload) = check_circular_references(&workspace, &args).unwrap();
        assert_eq!(payload["cycles"], json!([]));
        assert!(summary.contains("No circular"), "{summary}");
    }

    #[test]
    fn the_new_tools_are_advertised_with_a_schema() {
        for name in [
            "trace_precedents",
            "trace_dependents",
            "check_circular_references",
        ] {
            let spec =
                tools::find(name).unwrap_or_else(|| panic!("{name} is not in the catalogue"));
            assert!(
                spec.input_schema["properties"]["path"].is_object(),
                "{name}"
            );
        }
        assert!(
            tools::find("trace_precedents").unwrap().input_schema["properties"]["cell"].is_object()
        );
    }

    #[test]
    fn read_cells_can_skip_the_blanks() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "range": "A1:D2",
            "include_empty": false,
        }));
        let (summary, payload) = read_cells(&workspace, &args).unwrap();
        assert!(summary.contains('5'), "{summary}");
        assert_eq!(payload["cells"].as_array().unwrap().len(), 5);
    }

    #[test]
    fn read_cells_reports_the_number_format_alongside_the_value() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "range": "A4:A4",
        }));
        let (_, payload) = read_cells(&workspace, &args).unwrap();
        assert_eq!(payload["cells"][0]["number_format"], json!("0.00%"));
        assert_eq!(payload["cells"][0]["value"]["value"], json!(0.125));
    }

    #[test]
    fn read_formulas_finds_only_formulas() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx" }));
        let (_, payload) = read_formulas(&workspace, &args).unwrap();
        assert_eq!(payload["formulas"].as_array().unwrap().len(), 1);
        assert_eq!(payload["formulas"][0]["cell"], json!("B2"));
    }

    #[test]
    fn search_finds_values_case_insensitively_by_default() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "query": "alpha",
        }));
        let (_, payload) = search_values(&workspace, &args).unwrap();
        assert_eq!(payload["matches"].as_array().unwrap().len(), 1);
        assert_eq!(payload["matches"][0]["cell"], json!("A1"));
    }

    #[test]
    fn search_can_be_limited_and_marks_truncation() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "query": "a",
            "max_results": 1,
        }));
        let (_, payload) = search_values(&workspace, &args).unwrap();
        assert_eq!(payload["matches"].as_array().unwrap().len(), 1);
        assert_eq!(payload["truncated"], json!(true));
    }

    #[test]
    fn summarise_measures_the_numbers_and_ignores_the_rest() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "range": "A1:C1",
        }));
        let (_, payload) = summarize_range(&workspace, &args).unwrap();
        assert_eq!(payload["numeric"], json!(2));
        assert_eq!(payload["sum"], json!(30.0));
        assert_eq!(payload["min"], json!(10.0));
        assert_eq!(payload["max"], json!(20.0));
        assert_eq!(payload["mean"], json!(15.0));
    }

    #[test]
    fn a_range_with_no_numbers_reports_no_aggregates() {
        // A sum of zero would say "the total is zero"; the truth is "there are no numbers",
        // so every aggregate is null.
        let workspace = testing::workspace();
        let args = Args::new(&json!({
            "path": "report.xlsx",
            "range": "B2:B2",
        }));
        let (summary, payload) = summarize_range(&workspace, &args).unwrap();
        assert_eq!(payload["numeric"], json!(0));
        assert_eq!(payload["sum"], Value::Null);
        assert_eq!(payload["mean"], Value::Null);
        assert_eq!(payload["min"], Value::Null);
        assert_eq!(payload["max"], Value::Null);
        assert!(summary.contains("none of them numeric"), "{summary}");
    }

    #[test]
    fn csv_export_quotes_only_what_needs_it() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx", "range": "A1:C1" }));
        let (_, payload) = export_csv(&workspace, &args).unwrap();
        assert_eq!(payload["csv"], json!("alpha,10,20"));
    }

    #[test]
    fn named_ranges_are_listed_with_their_destinations() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "names.xlsx" }));
        let (_, payload) = list_named_ranges(&workspace, &args).unwrap();
        let names = payload["names"].as_array().unwrap();
        assert_eq!(names[0]["name"], json!("Totals"));
        assert_eq!(
            names[0]["destinations"][0]["sheet"],
            json!("Data"),
            "{names:?}"
        );
    }

    #[test]
    fn a_missing_sheet_is_reported_with_the_names_that_exist() {
        let workspace = testing::workspace();
        let args = Args::new(&json!({ "path": "report.xlsx", "sheet": "Ghost" }));
        let error = read_cells(&workspace, &args).unwrap_err();
        assert!(error.contains("Ghost"), "{error}");
        assert!(error.contains("Data"), "{error}");
    }
}
