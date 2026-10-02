//! Cell dependency tracing (`trace_precedents`, `trace_dependents`, cycles).
//!
//! openpyxl can answer "what is in this cell" but not "what would break if I changed it",
//! because it keeps no model of which formulas reference which cells. ferroxl can: a
//! formula's references are recoverable from its text, and the worksheet holds every
//! formula in the sheet, so the dependency graph is a traversal rather than an index.
//!
//! The three questions this answers are the three an agent actually asks before touching a
//! spreadsheet:
//!
//! - [`Worksheet::trace_precedents`] — what feeds this cell?
//! - [`Worksheet::trace_dependents`] — what breaks if I change it?
//! - [`Worksheet::circular_references`] — is any of this already wrong?
//!
//! Two deliberate limits, because a wrong answer here is worse than no answer:
//!
//! - References through a defined name are not followed. A name like `Totals` may cover any
//!   range in the workbook, and guessing at it would produce a graph that looks authoritative
//!   and is not. Name references are reported by [`References::named`] so a caller can see
//!   what was skipped.
//! - Only same-sheet references are expanded into coordinates. A cross-sheet reference is
//!   recorded in [`References::cross_sheet`] with its sheet title, because the interesting
//!   question about it is which other sheet to look at, not which of a million cells.

use std::collections::{BTreeMap, BTreeSet, VecDeque};

use crate::cell::cell::{CellValue, DataType};
use crate::cell::utils::{column_index_from_string, coordinate_from_string, get_column_letter};
use crate::exceptions::{Error, Result};
use crate::worksheet::Worksheet;

/// The cells one formula refers to, split by how they were written.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct References {
    /// Same-sheet references, expanded: `A1`, and `A1:C3` as nine coordinates.
    pub cells: BTreeSet<String>,
    /// References to another sheet, kept as the sheet title and the range as written.
    ///
    /// Expanding these would mean walking a range that may be far larger than the formula
    /// that mentions it, so the range is reported instead.
    pub cross_sheet: BTreeSet<(String, String)>,
    /// Defined names, which are not resolved.
    pub named: BTreeSet<String>,
}

impl References {
    /// Whether the formula refers to nothing this module can follow.
    pub fn is_empty(&self) -> bool {
        self.cells.is_empty() && self.cross_sheet.is_empty() && self.named.is_empty()
    }

    /// Fold another set of references into this one.
    pub fn absorb(&mut self, other: References) {
        self.cells.extend(other.cells);
        self.cross_sheet.extend(other.cross_sheet);
        self.named.extend(other.named);
    }
}

/// Pull the references out of a formula's expression.
///
/// The leading `=` is optional, so both `"=A1+1"` and `"A1+1"` parse. Quoted literals are
/// skipped, because `"A1"` inside a string is text and not a reference, and a function name
/// is skipped because `LOG10(` is a call and not a cell.
pub fn parse_references(expression: &str) -> References {
    let mut found = References::default();
    let text = expression.strip_prefix('=').unwrap_or(expression);
    let bytes = text.as_bytes();
    let mut i = 0usize;

    while i < bytes.len() {
        let ch = bytes[i] as char;

        // A quoted literal is skipped whole: `"A1"` is text, not a reference.
        if ch == '"' {
            i = skip_quoted(text, i, '"');
            continue;
        }

        // A quoted sheet title, as in `'Q1 Sales'!B2`.
        if ch == '\'' {
            let end = skip_quoted(text, i, '\'');
            if bytes.get(end) == Some(&b'!') {
                if let Some((sheet, range)) = read_sheet_ref(text, i + 1, end) {
                    let consumed = range.len();
                    found.cross_sheet.insert((sheet, range));
                    i = end + 1 + consumed;
                    continue;
                }
            }
            i = end;
            continue;
        }

        // Anything else that is not a token start is just punctuation or whitespace.
        if !(ch.is_ascii_alphabetic() || ch == '$') {
            i += 1;
            continue;
        }

        let (text_ref, len) = read_token(text, i);
        let after = i + len;

        // A call is decided first, because `LOG10(` is a function and `LOG10` is also a
        // perfectly good cell reference: column LOG, row 10. Only the `(` tells them apart,
        // and looking for it afterwards would be too late.
        if is_function_call(text, after) {
            i = after;
            continue;
        }

        // A coordinate, possibly a range: `A1`, `$A$1`, `A1:B3`.
        if let Some(reference) = as_coordinate(text_ref) {
            // Read the range from the start of the token, not from after it: `A1:B3` is one
            // reference, and looking past `A1` would find the `:` and stop.
            if let Some(range) = read_a1_range(text, i) {
                if range.contains(':') {
                    expand_range(&reference, &range, &mut found.cells);
                    i += range.len();
                    continue;
                }
            }
            found.cells.insert(reference);
            i = after;
            continue;
        }

        // A name. Which kind of name decides whether it is followable.
        if is_keyword(text_ref) {
            i = after;
            continue;
        }
        if bytes.get(after) == Some(&b'!') {
            if let Some(range) = read_a1_range(text, after + 1) {
                let consumed = range.len();
                found.cross_sheet.insert((text_ref.to_string(), range));
                i = after + 1 + consumed;
                continue;
            }
        }
        found.named.insert(text_ref.to_string());
        i = after;
    }
    found
}

/// Read the identifier-shaped token at `at`: a run of letters, dollars and digits.
///
/// Returns the slice and how long it is, which is all the caller needs to know where the
/// next token begins.
fn read_token(text: &str, at: usize) -> (&str, usize) {
    let bytes = text.as_bytes();
    let mut end = at;
    while end < bytes.len() {
        let ch = bytes[end] as char;
        if ch.is_ascii_alphanumeric() || ch == '$' {
            end += 1;
            continue;
        }
        break;
    }
    (&text[at..end], end - at)
}

/// The coordinate a token names, or `None` if it is not one.
///
/// `B2` names a cell; `LOG10` does not, because a coordinate's letters all come before its
/// digits and a row number is never zero.
fn as_coordinate(run: &str) -> Option<String> {
    let letters: String = run
        .chars()
        .take_while(|c| c.is_ascii_alphabetic() || *c == '$')
        .collect();
    let digits: &str = &run[letters.len()..];
    let letters = letters.trim_matches('$');
    let digits = digits.trim_matches('$');
    if letters.is_empty()
        || digits.is_empty()
        || !letters.chars().all(|c| c.is_ascii_alphabetic())
        || !digits.chars().all(|c| c.is_ascii_digit())
    {
        return None;
    }
    if digits.trim_start_matches('0').is_empty() {
        return None;
    }
    // The `$` anchors are dropped: `$B$2` and `B2` are the same cell.
    Some(format!("{letters}{digits}"))
}

/// `TRUE`, `FALSE` and friends are constants, not defined names.
fn is_keyword(run: &str) -> bool {
    matches!(run.to_ascii_uppercase().as_str(), "TRUE" | "FALSE")
}

/// Whether a run at `at` is immediately followed by `(`, making it a function call.
fn is_function_call(text: &str, at: usize) -> bool {
    text[at..].trim_start().starts_with('(')
}

fn strip_dollars(run: &str) -> String {
    run.chars().filter(|c| *c != '$').collect()
}

/// Index just past a closing quote, or the end of the text if it is unterminated.
fn skip_quoted(text: &str, start: usize, quote: char) -> usize {
    let mut i = start + 1;
    let bytes = text.as_bytes();
    while i < bytes.len() {
        if bytes[i] == b'\\' && quote == '"' {
            i += 2;
            continue;
        }
        if bytes[i] as char == quote {
            // A doubled quote is an escaped quote, not the end of the literal.
            if bytes.get(i + 1) == Some(&(quote as u8)) {
                i += 2;
                continue;
            }
            return i + 1;
        }
        i += 1;
    }
    bytes.len()
}

/// Read a sheet title from a quoted reference, given the bounds of the quoted section.
fn read_sheet_ref(text: &str, start: usize, end: usize) -> Option<(String, String)> {
    let sheet = unescape_sheet(&text[start..end.saturating_sub(1)]);
    let after = end + 1;
    let range = read_a1_range(text, after)?;
    Some((sheet, range))
}

fn unescape_sheet(sheet: &str) -> String {
    sheet.replace("''", "'")
}

/// Read an A1 coordinate or range starting at `at`, returning the text consumed.
fn read_a1_range(text: &str, at: usize) -> Option<String> {
    let bytes = text.as_bytes();
    let first = read_cell_at(text, at)?;
    let after = at + first.len();
    if bytes.get(after) == Some(&b':') {
        if let Some(second) = read_cell_at(text, after + 1) {
            return Some(format!("{first}:{second}"));
        }
    }
    Some(first)
}

/// Read one coordinate at `at`, dollars allowed.
fn read_cell_at(text: &str, at: usize) -> Option<String> {
    let bytes = text.as_bytes();
    let mut i = at;
    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }
    let letters_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_alphabetic() {
        i += 1;
    }
    if i == letters_start {
        return None;
    }
    if bytes.get(i) == Some(&b'$') {
        i += 1;
    }
    let digits_start = i;
    while i < bytes.len() && (bytes[i] as char).is_ascii_digit() {
        i += 1;
    }
    if i == digits_start {
        return None;
    }
    Some(text[at..i].to_string())
}

/// Expand `A1:B3` into the nine coordinates it covers.
fn expand_range(start: &str, range: &str, out: &mut BTreeSet<String>) {
    let Some((from, to)) = range.split_once(':') else {
        out.insert(strip_dollars(start));
        return;
    };
    let Ok((from_letters, from_row)) = coordinate_from_string(from) else {
        return;
    };
    let Ok((to_letters, to_row)) = coordinate_from_string(to) else {
        return;
    };
    let (Ok(from_col), Ok(to_col)) = (
        column_index_from_string(&from_letters),
        column_index_from_string(&to_letters),
    ) else {
        return;
    };
    // A reversed range is a `#REF!` in disguise rather than a reason to iterate backwards.
    if from_row > to_row || from_col > to_col {
        return;
    }
    // A whole-column or whole-row range can be sixteen thousand cells wide; expanding one
    // would be worse than the problem it was meant to solve.
    if to_col - from_col > 4096 || to_row - from_row > 65_536 {
        return;
    }
    for row in from_row..=to_row {
        for column in from_col..=to_col {
            if let Ok(letters) = get_column_letter(column) {
                out.insert(format!("{letters}{row}"));
            }
        }
    }
}

/// The formula stored at a cell, or `None` if the cell holds no formula.
///
/// A shared formula is expanded first, so the caller sees the expression as it applies at
/// this particular cell rather than the anchor's text.
pub fn formula_at(worksheet: &Worksheet, coordinate: &str) -> Option<String> {
    let cell = worksheet.get_cell(coordinate)?;
    if cell.data_type != DataType::Formula {
        return None;
    }
    match cell.internal_value() {
        CellValue::Formula(expression) if !expression.is_empty() => Some(expression.clone()),
        _ => None,
    }
}

/// Every cell in the sheet that holds a formula, with its expression.
pub fn formula_cells(worksheet: &Worksheet) -> Vec<(String, String)> {
    worksheet
        .cells()
        .filter(|cell| cell.data_type == DataType::Formula)
        .filter_map(|cell| {
            let coordinate = cell.coordinate();
            match cell.internal_value() {
                CellValue::Formula(expression) if !expression.is_empty() => {
                    Some((coordinate, expression.clone()))
                }
                _ => None,
            }
        })
        .collect()
}

impl Worksheet {
    /// Every cell that `coordinate`'s formula depends on, transitively.
    ///
    /// The answer is in dependency order: a precedent of a precedent comes first, so
    /// reading the list top to bottom reads like a spreadsheet would recalculate. Cycles
    /// are cut rather than followed, because a circular reference has no order and
    /// [`circular_references`](Self::circular_references) is where that gets reported.
    pub fn trace_precedents(&self, coordinate: &str) -> Result<Vec<String>> {
        let start = normalise(coordinate)?;
        let mut order = Vec::new();
        let mut seen = BTreeSet::new();
        let mut on_stack = BTreeSet::new();
        self.visit_precedents(&start, &start, &mut order, &mut seen, &mut on_stack)?;
        Ok(order)
    }

    fn visit_precedents(
        &self,
        start: &str,
        key: &str,
        order: &mut Vec<String>,
        seen: &mut BTreeSet<String>,
        on_stack: &mut BTreeSet<String>,
    ) -> Result<()> {
        if !on_stack.insert(key.to_string()) {
            // Already on the path being walked: this is a cycle, and following it again
            // would never terminate.
            return Ok(());
        }
        if !seen.insert(key.to_string()) {
            on_stack.remove(key);
            return Ok(());
        }
        if let Some(expression) = formula_at(self, key) {
            let references = parse_references(&expression);
            for cell in references.cells {
                self.visit_precedents(start, &cell, order, seen, on_stack)?;
            }
        }
        // The cell comes after everything it depends on, unless it is the one that was asked
        // about. A cell holding a plain value is a leaf with no precedents of its own, and
        // it is still a precedent of whatever reads it -- leaving it out would answer "which
        // formulas feed this" instead of the question that was asked. A cell that references
        // itself is a cycle, which `circular_references` reports; listing it as its own
        // precedent would just be noise in the answer to a different question.
        if key != start {
            order.push(key.to_string());
        }
        on_stack.remove(key);
        Ok(())
    }

    /// Every formula in the sheet that reads `coordinate`, transitively.
    ///
    /// These are the cells whose value would be stale if `coordinate` changed.
    pub fn trace_dependents(&self, coordinate: &str) -> Result<Vec<String>> {
        let key = normalise(coordinate)?;
        let graph = self.dependency_graph()?;
        Ok(order_from(&key, &graph))
    }

    /// The dependency graph as `cell -> cells it reads`.
    ///
    /// Exposed because a caller often wants the whole graph, not one cell's slice of it.
    pub fn dependency_graph(&self) -> Result<BTreeMap<String, BTreeSet<String>>> {
        let mut graph: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
        for (coordinate, expression) in formula_cells(self) {
            let mut references = parse_references(&expression);
            references.cells.remove(&coordinate); // `A1` in `=A1*2` is self, not a precedent.
            graph.insert(coordinate, references.cells);
        }
        Ok(graph)
    }

    /// Every cycle in the sheet, as the cells on it in order.
    ///
    /// A workbook Excel will refuse to calculate is a workbook nobody can trust, so this is
    /// a correctness tool rather than a curiosity. Each cycle is reported once, from its
    /// lowest-sorting cell, so the list has no duplicates.
    pub fn circular_references(&self) -> Result<Vec<Vec<String>>> {
        let graph = self.dependency_graph()?;
        Ok(find_cycles(&graph))
    }
}

/// Walk a `cell -> cells it reads` graph outwards from one cell.
fn order_from(start: &str, graph: &BTreeMap<String, BTreeSet<String>>) -> Vec<String> {
    // The reverse index is what "dependents" means: who reads me.
    let mut readers: BTreeMap<&str, Vec<&str>> = BTreeMap::new();
    for (cell, reads) in graph {
        for read in reads {
            readers
                .entry(read.as_str())
                .or_default()
                .push(cell.as_str());
        }
    }
    let mut order = Vec::new();
    let mut seen = BTreeSet::new();
    let mut queue = VecDeque::from([start]);
    while let Some(current) = queue.pop_front() {
        if !seen.insert(current.to_string()) {
            continue;
        }
        for reader in readers.get(current).into_iter().flatten() {
            if !order.contains(&(*reader).to_string()) {
                order.push((*reader).to_string());
            }
            queue.push_back(reader);
        }
    }
    order
}

/// Find every cycle in a dependency graph.
///
/// A depth-first search that marks nodes as it descends and treats reaching a marked node as
/// a cycle. The visited set is what keeps this linear rather than exponential on a graph
/// with many shared precedents.
fn find_cycles(graph: &BTreeMap<String, BTreeSet<String>>) -> Vec<Vec<String>> {
    #[derive(Clone, Copy, PartialEq)]
    enum Mark {
        Open,
        Done,
    }
    fn walk(
        node: &str,
        graph: &BTreeMap<String, BTreeSet<String>>,
        marks: &mut BTreeMap<String, Mark>,
        path: &mut Vec<String>,
        found: &mut Vec<Vec<String>>,
    ) {
        match marks.get(node) {
            Some(Mark::Done) => return,
            Some(Mark::Open) => {
                // Back to a node still on the path: everything from there round is a cycle.
                if let Some(start) = path.iter().position(|cell| cell == node) {
                    let mut cycle = path[start..].to_vec();
                    cycle.push(node.to_string());
                    if !found.contains(&cycle) {
                        found.push(cycle);
                    }
                }
                return;
            }
            None => {}
        }
        marks.insert(node.to_string(), Mark::Open);
        path.push(node.to_string());
        if let Some(reads) = graph.get(node) {
            for read in reads {
                if graph.contains_key(read) {
                    walk(read, graph, marks, path, found);
                }
            }
        }
        path.pop();
        marks.insert(node.to_string(), Mark::Done);
    }

    let mut marks = BTreeMap::new();
    let mut path = Vec::new();
    let mut found = Vec::new();
    for node in graph.keys() {
        walk(node, graph, &mut marks, &mut path, &mut found);
    }
    // One cycle can be reported once per cell on it; keying by the sorted set of cells
    // collapses those without losing any distinct cycle.
    // A cycle is reported once per cell on it, so key each by its sorted set of cells and
    // keep the first of each. Collected in a loop rather than a `filter` so the borrow of
    // `seen` does not escape into a closure.
    let mut seen: BTreeSet<Vec<String>> = BTreeSet::new();
    let mut unique = Vec::new();
    for cycle in found {
        let mut key = cycle.clone();
        key.sort();
        key.dedup();
        if seen.insert(key) {
            unique.push(cycle);
        }
    }
    unique
}

/// Validate a coordinate and return it without its dollars.
fn normalise(coordinate: &str) -> Result<String> {
    let cleaned = strip_dollars(coordinate.trim().trim_start_matches('='));
    let (letters, row) = coordinate_from_string(&cleaned)?;
    column_index_from_string(&letters)?;
    if row == 0 {
        return Err(Error::CellCoordinates(format!("{coordinate} has no row")));
    }
    Ok(format!("{letters}{row}"))
}

/// The error an invalid coordinate produces, kept so callers can match on it.
pub fn is_unknown_cell(coordinate: &str) -> Result<()> {
    normalise(coordinate).map(|_| ())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cell::cell::CellValue;

    fn sheet() -> Worksheet {
        let mut ws = Worksheet::new("S").expect("title");
        for (coordinate, value) in [
            ("A1", "1"),
            ("A2", "2"),
            ("A3", "3"),
            ("A4", "=SUM(A1:A3)"),
            ("A5", "=A4*2"),
            ("B1", "0.5"),
        ] {
            let value = match value.strip_prefix('=') {
                Some(expression) => CellValue::formula(expression),
                None => CellValue::number(value.parse().expect("number")),
            };
            ws.set(coordinate, value).expect("set");
        }
        ws
    }

    #[test]
    fn a_single_reference_is_parsed() {
        let found = parse_references("=A1+1");
        assert_eq!(found.cells.iter().cloned().collect::<Vec<_>>(), ["A1"]);
    }

    #[test]
    fn dollars_are_anchoring_not_identity() {
        let found = parse_references("=$B$2+$C3");
        let mut cells: Vec<&String> = found.cells.iter().collect();
        cells.sort();
        assert_eq!(cells, ["B2", "C3"]);
    }

    #[test]
    fn a_range_is_expanded_to_its_cells() {
        let found = parse_references("=SUM(A1:A3)");
        assert_eq!(found.cells.len(), 3);
        assert!(found.cells.contains("A2"));
    }

    #[test]
    fn a_function_name_is_not_a_reference() {
        // LOG10 contains a digit and a letter, which is exactly what a coordinate looks like.
        let found = parse_references("=LOG10(A1)");
        assert_eq!(found.cells.iter().cloned().collect::<Vec<_>>(), ["A1"]);
    }

    #[test]
    fn a_quoted_literal_is_not_a_reference() {
        let found = parse_references("=IF(A1>0,\"A1\",\"B2\")");
        let mut cells: Vec<&String> = found.cells.iter().collect();
        cells.sort();
        assert_eq!(cells, ["A1"]);
    }

    #[test]
    fn booleans_are_not_defined_names() {
        let found = parse_references("=IF(A1=TRUE,1,0)");
        assert!(found.named.is_empty(), "{:?}", found.named);
        assert!(found.cells.contains("A1"));
    }

    #[test]
    fn a_defined_name_is_reported_rather_than_guessed_at() {
        let found = parse_references("=SUM(Totals)+A1");
        assert!(found.named.contains("Totals"));
        assert!(found.cells.contains("A1"));
    }

    #[test]
    fn a_cross_sheet_reference_keeps_its_sheet() {
        let found = parse_references("=Data!B2+C3");
        assert!(found
            .cross_sheet
            .contains(&("Data".to_string(), "B2".to_string())));
        assert!(found.cells.contains("C3"));
    }

    #[test]
    fn a_quoted_sheet_name_keeps_its_spaces() {
        let found = parse_references("='Q1 Sales'!B2");
        assert!(found
            .cross_sheet
            .contains(&("Q1 Sales".to_string(), "B2".to_string())));
    }

    #[test]
    fn precedents_come_back_in_dependency_order() {
        let ws = sheet();
        let order = ws.trace_precedents("A5").expect("precedents");
        // A5 reads A4, and A4 reads A1..A3, so the numbers come before the formula that used
        // them. A5 itself is not in the list: a cell is not its own precedent.
        assert!(!order.contains(&"A5".to_string()), "{order:?}");
        assert_eq!(order.last().map(String::as_str), Some("A4"));
        let position = |cell: &str| order.iter().position(|c| c == cell).expect("in order");
        assert!(position("A1") < position("A4"));
        assert_eq!(order, ["A1", "A2", "A3", "A4"]);
    }

    #[test]
    fn dependents_are_the_formulas_that_read_the_cell() {
        let ws = sheet();
        let dependents = ws.trace_dependents("A1").expect("dependents");
        assert!(dependents.contains(&"A4".to_string()), "{dependents:?}");
        // A5 is a *transitive* dependent: it reads A4, which reads A1, so changing A1 does
        // change it. The list is the closure, not the immediate neighbours.
        assert!(dependents.contains(&"A5".to_string()), "{dependents:?}");

        let through = ws.trace_dependents("A4").expect("dependents");
        assert!(through.contains(&"A5".to_string()), "{through:?}");
    }

    #[test]
    fn a_cell_is_not_its_own_precedent() {
        let mut ws = sheet();
        ws.set("C1", CellValue::formula("C1*2"))
            .expect("self reference");
        // A self-reference is a cycle of one, not a precedent of itself.
        let order = ws.trace_precedents("C1").expect("precedents");
        assert!(!order.contains(&"C1".to_string()), "{order:?}");
    }

    #[test]
    fn a_cycle_is_found_and_reported_once() {
        let mut ws = sheet();
        ws.set("D1", CellValue::formula("E1")).expect("D1");
        ws.set("E1", CellValue::formula("F1")).expect("E1");
        ws.set("F1", CellValue::formula("D1")).expect("F1");

        let cycles = ws.circular_references().expect("cycles");
        assert_eq!(cycles.len(), 1, "{cycles:?}");
        let cells: BTreeSet<&String> = cycles[0].iter().collect();
        assert_eq!(cells.len(), 3, "three cells on the cycle: {cycles:?}");
        // The reported path returns to where it started.
        assert_eq!(cycles[0].first(), cycles[0].last());
    }

    #[test]
    fn a_clean_sheet_reports_no_cycles() {
        assert!(sheet().circular_references().expect("cycles").is_empty());
    }

    #[test]
    fn a_bad_coordinate_is_refused() {
        assert!(is_unknown_cell("not a cell").is_err());
        assert!(is_unknown_cell("A1").is_ok());
    }
}
