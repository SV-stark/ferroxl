//! Evaluating formulas, so a saved workbook carries values rather than blanks.
//!
//! # Why this is careful
//!
//! A wrong cached value is worse than a missing one. A missing `<v>` is visible -- a reader
//! knows it has to recalculate. A wrong one is silently believed, and the mistake surfaces days
//! later in a report nobody traces back to this library. So three rules hold throughout:
//!
//! - **Nothing is invented.** A formula this module cannot evaluate gets no `<v>` written at
//!   all, and is named in [`Recalculation::unresolved`] with the reason.
//! - **The caller can check first.** [`Recalculation`] is returned rather than side effects
//!   only, and [`supports`] answers "will you handle this?" without running anything.
//! - **Excel recomputes on open.** `recalculate` sets `calcPr/@fullCalcOnLoad`, so these values
//!   are a convenience for readers that are not Excel. A value this module gets wrong cannot
//!   survive being opened and saved by a human.
//!
//! # What is deliberately absent
//!
//! The functions that are hard to get right are missing rather than approximated:
//! `VLOOKUP`'s approximate match against unsorted data, `INDEX` returning a reference rather
//! than a value, `TEXT`'s format codes, date arithmetic across the 1900 leap-year bug, and
//! `SUMIF`'s wildcards. Each is a day of edge cases and a place to be subtly wrong, so a
//! formula using one is reported unresolved instead of being given a plausible number.

use crate::cell::cell::CellValue;
use std::collections::BTreeMap;

/// Every function this module can evaluate.
///
/// Kept in step with the dispatch table in `call_function`, and asserted against it by a test,
/// because a function that is listed but unimplemented would answer `supports` with a `yes` and
/// then produce nothing.
pub const SUPPORTED_FUNCTIONS: &[&str] = &[
    "ABS",
    "AVERAGE",
    "AND",
    "CONCAT",
    "CONCATENATE",
    "COUNT",
    "COUNTA",
    "FALSE",
    "IF",
    "IFERROR",
    "INT",
    "LEFT",
    "LEN",
    "LOWER",
    "MAX",
    "MIN",
    "MOD",
    "NOT",
    "OR",
    "POWER",
    "PRODUCT",
    "RIGHT",
    "ROUND",
    "ROUNDDOWN",
    "ROUNDUP",
    "SIGN",
    "SQRT",
    "SUM",
    "TRIM",
    "TRUE",
    "UPPER",
];

/// Whether this module can evaluate a formula naming `name`.
///
/// The check is on the name alone, so it cannot tell you a formula is safe -- only that no part
/// of it is unrecognised. A formula can still come back unresolved for a reference this module
/// could not read.
///
/// ```
/// assert!(ferroxl::formula::supports("SUM"));
/// assert!(!ferroxl::formula::supports("VLOOKUP"));
/// ```
pub fn supports(name: &str) -> bool {
    let wanted = name.trim().to_uppercase();
    SUPPORTED_FUNCTIONS.contains(&wanted.as_str())
}

/// A formula that could not be evaluated, and why.
///
/// The reason is part of the contract: "unsupported function" and "circular reference" call for
/// different responses from the caller, and collapsing them into a failure count would throw
/// that away.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Unresolved {
    /// The cell holding the formula.
    pub sheet: String,
    /// Its coordinate.
    pub coordinate: String,
    /// The formula as written.
    pub formula: String,
    /// Why no value could be produced.
    pub reason: String,
}

/// What a recalculation produced.
#[derive(Debug, Clone, Default)]
pub struct Recalculation {
    /// Values written, keyed by sheet title then coordinate.
    ///
    /// Only cells that evaluated successfully appear, so the length of this map is the number of
    /// cached values in the saved file.
    pub computed: BTreeMap<String, BTreeMap<String, CellValue>>,
    /// Formulas left without a value, with the reason for each.
    pub unresolved: Vec<Unresolved>,
}

impl Recalculation {
    /// How many cells got a value.
    pub fn computed_count(&self) -> usize {
        self.computed.values().map(BTreeMap::len).sum()
    }

    /// Whether every formula in the workbook was evaluated.
    pub fn is_complete(&self) -> bool {
        self.unresolved.is_empty()
    }

    /// The reasons the evaluation stopped, deduplicated with a count for each.
    ///
    /// The useful summary when a workbook has four hundred formulas and one function is
    /// missing: four hundred messages saying the same thing is noise.
    pub fn reasons(&self) -> Vec<(String, usize)> {
        let mut counts: BTreeMap<&str, usize> = BTreeMap::new();
        for entry in &self.unresolved {
            *counts.entry(entry.reason.as_str()).or_default() += 1;
        }
        counts
            .into_iter()
            .map(|(reason, count)| (reason.to_string(), count))
            .collect()
    }
}

/// How a reference resolved.
enum Resolved {
    /// A single cell's value.
    Single(CellValue),
    /// A range's values, in reading order.
    Range(Vec<CellValue>),
}

/// Reads cells for the evaluator.
pub trait ValueSource {
    /// The value of one cell, or blank if it is empty or absent.
    fn cell(&self, sheet: &str, coordinate: &str) -> CellValue;
}

/// Evaluates `formula` as it would be evaluated in `sheet`.
///
/// `Ok(Some(_))` is a computed value. `Err` carries the reason nothing could be produced --
/// a malformed formula, an unrecognised function, an unreadable range -- and never a guess.
/// There is deliberately no `Ok(None)`: "no value" and "here is the value" should not be
/// adjacent options to mix up.
pub fn evaluate(
    formula: &str,
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, String> {
    let expression = formula.trim().trim_start_matches('=');
    let tokens = tokenize(expression).map_err(|reason| format!("cannot parse: {reason}"))?;
    let mut parser = Parser {
        tokens,
        position: 0,
    };
    let node = parser
        .parse_expression()
        .ok_or_else(|| "cannot parse: unexpected end of formula".to_string())?;
    if !parser.at_end() {
        return Err("cannot parse: trailing text after the formula".to_string());
    }
    eval(&node, sheet, source)
}

// ---------------------------------------------------------------------------------------
// Tokens
// ---------------------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
enum Token {
    Number(f64),
    Text(String),
    /// `TRUE`, `FALSE`, or `#DIV/0!` and the rest.
    Word(String),
    /// A bare name that is not a literal: a function name or a reference.
    Name(String),
    Operator(String),
    OpenParen,
    CloseParen,
    Comma,
    Colon,
    Bang,
}

fn tokenize(input: &str) -> std::result::Result<Vec<Token>, String> {
    let chars: Vec<char> = input.chars().collect();
    let mut tokens = Vec::new();
    let mut at = 0usize;

    while at < chars.len() {
        let ch = chars[at];
        match ch {
            ' ' | '\t' | '\n' | '\r' => at += 1,
            '(' => {
                tokens.push(Token::OpenParen);
                at += 1;
            }
            ')' => {
                tokens.push(Token::CloseParen);
                at += 1;
            }
            ',' | ';' => {
                // Excel accepts either as an argument separator depending on locale.
                tokens.push(Token::Comma);
                at += 1;
            }
            ':' => {
                tokens.push(Token::Colon);
                at += 1;
            }
            '!' => {
                tokens.push(Token::Bang);
                at += 1;
            }
            '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<' | '>' | '%' => {
                let start = at;
                at += 1;
                // `<>` and `<=` are one operator; `>=` likewise. `==` is *not* -- it is two
                // comparisons, so the second `=` is left for the next pass.
                let is_pair = match ch {
                    '<' | '>' => matches!(chars.get(at), Some('>') | Some('=')),
                    _ => false,
                };
                if is_pair {
                    at += 1;
                }
                tokens.push(Token::Operator(chars[start..at].iter().collect()));
            }
            '"' => {
                let mut text = String::new();
                at += 1;
                loop {
                    match chars.get(at) {
                        None => return Err("unterminated text".to_string()),
                        Some('"') => {
                            // A doubled quote is one literal quote, which is how Excel escapes it.
                            if chars.get(at + 1) == Some(&'"') {
                                text.push('"');
                                at += 2;
                            } else {
                                at += 1;
                                break;
                            }
                        }
                        Some(other) => {
                            text.push(*other);
                            at += 1;
                        }
                    }
                }
                tokens.push(Token::Text(text));
            }
            '#' => {
                let start = at;
                while at < chars.len() && (chars[at].is_ascii_alphanumeric() || chars[at] == '!') {
                    at += 1;
                }
                let word: String = chars[start..at].iter().collect();
                if ERROR_CODES.contains(&word.to_uppercase().as_str()) {
                    tokens.push(Token::Word(word));
                } else {
                    return Err(format!("unknown error literal {word:?}"));
                }
            }
            _ if ch.is_ascii_digit()
                || (ch == '.' && chars.get(at + 1).is_some_and(|c| c.is_ascii_digit())) =>
            {
                let start = at;
                while at < chars.len() && (chars[at].is_ascii_digit() || chars[at] == '.') {
                    at += 1;
                }
                // Scientific notation, which Excel writes for very large and small numbers.
                if matches!(chars.get(at), Some('E') | Some('e')) {
                    let mut ahead = at + 1;
                    if matches!(chars.get(ahead), Some('+') | Some('-')) {
                        ahead += 1;
                    }
                    if chars.get(ahead).is_some_and(char::is_ascii_digit) {
                        at = ahead;
                        while at < chars.len() && chars[at].is_ascii_digit() {
                            at += 1;
                        }
                    }
                }
                let text: String = chars[start..at].iter().collect();
                let value = text
                    .parse::<f64>()
                    .map_err(|_| format!("cannot read the number {text:?}"))?;
                tokens.push(Token::Number(value));
            }
            _ if is_name_start(ch) => {
                let start = at;
                while at < chars.len() && is_name_part(chars[at]) {
                    at += 1;
                }
                let word: String = chars[start..at].iter().collect();
                // A `$` may lead or trail a reference, so it has to be consumed here rather
                // than being rejected as an unknown character.
                while at < chars.len() && chars[at] == '$' {
                    at += 1;
                }
                let upper = word.to_uppercase();
                if matches!(upper.as_str(), "TRUE" | "FALSE") {
                    tokens.push(Token::Word(upper));
                } else if crate::styles::named_style::builtin(&word).is_some()
                    && !looks_like_reference(&word)
                {
                    // `Good`, `Bad`, `Title` and the rest: a name, not a reference. They are not
                    // values, so evaluation reports them rather than resolving them to a style.
                    tokens.push(Token::Name(word));
                } else {
                    tokens.push(Token::Name(word));
                }
            }
            _ => return Err(format!("unexpected character {ch:?}")),
        }
    }
    Ok(tokens)
}

/// The seven values Excel treats as errors, so the tokenizer can read `#DIV/0!` as one.
const ERROR_CODES: [&str; 7] = [
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
];

fn is_name_start(ch: char) -> bool {
    ch.is_ascii_alphabetic() || ch == '_' || ch == '\\' || ch == '$'
}

fn is_name_part(ch: char) -> bool {
    ch.is_ascii_alphanumeric() || ch == '_' || ch == '.' || ch == '\\' || ch == '$'
}

/// Whether `word` reads as `A1` rather than as a name.
///
/// `$` is stripped first so `$A$1` is a reference, and the digit test rejects `Good` while
/// accepting `XFD1048576`.
fn looks_like_reference(word: &str) -> bool {
    let letters: String = word.chars().filter(|c| *c != '$').collect();
    let mut chars = letters.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !first.is_ascii_alphabetic() {
        return false;
    }
    if chars.clone().all(|c| c.is_ascii_alphabetic()) {
        return false;
    }
    chars.all(|c| c.is_ascii_alphanumeric())
}

// ---------------------------------------------------------------------------------------
// Syntax tree
// ---------------------------------------------------------------------------------------

#[derive(Debug, Clone, PartialEq)]
enum Node {
    Literal(CellValue),
    /// A cell or range reference, already split into an optional sheet and a body.
    Reference(Option<String>, String),
    Function(String, Vec<Node>),
    Unary(String, Box<Node>),
    Binary(String, Box<Node>, Box<Node>),
}

struct Parser {
    tokens: Vec<Token>,
    position: usize,
}

impl Parser {
    fn at_end(&self) -> bool {
        self.position >= self.tokens.len()
    }

    fn peek(&self) -> Option<&Token> {
        self.tokens.get(self.position)
    }

    fn take(&mut self) -> Option<Token> {
        let token = self.tokens.get(self.position).cloned();
        if token.is_some() {
            self.position += 1;
        }
        token
    }

    fn eat_operator(&mut self, wanted: &str) -> bool {
        if matches!(self.peek(), Some(Token::Operator(o)) if o == wanted) {
            self.position += 1;
            true
        } else {
            false
        }
    }

    /// Comparison, the loosest binding.
    fn parse_expression(&mut self) -> Option<Node> {
        let mut left = self.parse_concat()?;
        loop {
            let operator = match self.peek() {
                Some(Token::Operator(o))
                    if matches!(o.as_str(), "=" | "<>" | "<" | "<=" | ">" | ">=") =>
                {
                    o.clone()
                }
                _ => return Some(left),
            };
            self.position += 1;
            let right = self.parse_concat()?;
            left = Node::Binary(operator, Box::new(left), Box::new(right));
        }
    }

    fn parse_concat(&mut self) -> Option<Node> {
        let mut left = self.parse_additive()?;
        while self.eat_operator("&") {
            let right = self.parse_additive()?;
            left = Node::Binary("&".to_string(), Box::new(left), Box::new(right));
        }
        Some(left)
    }

    fn parse_additive(&mut self) -> Option<Node> {
        let mut left = self.parse_multiplicative()?;
        loop {
            let operator = match self.peek() {
                Some(Token::Operator(o)) if o == "+" || o == "-" => o.clone(),
                _ => return Some(left),
            };
            self.position += 1;
            let right = self.parse_multiplicative()?;
            left = Node::Binary(operator, Box::new(left), Box::new(right));
        }
    }

    fn parse_multiplicative(&mut self) -> Option<Node> {
        let mut left = self.parse_exponent()?;
        loop {
            let operator = match self.peek() {
                Some(Token::Operator(o)) if o == "*" || o == "/" => o.clone(),
                _ => return Some(left),
            };
            self.position += 1;
            let right = self.parse_exponent()?;
            left = Node::Binary(operator, Box::new(left), Box::new(right));
        }
    }

    fn parse_exponent(&mut self) -> Option<Node> {
        let mut left = self.parse_unary()?;
        // Left-associative, as Excel has it: `2^3^2` is 64, not 512.
        while self.eat_operator("^") {
            let right = self.parse_unary()?;
            left = Node::Binary("^".to_string(), Box::new(left), Box::new(right));
        }
        Some(left)
    }

    /// Unary minus binds *tighter* than `^` in Excel, which is why `-2^2` is 4 and not -4.
    fn parse_unary(&mut self) -> Option<Node> {
        if self.eat_operator("-") {
            let inner = self.parse_unary()?;
            return Some(Node::Unary("-".to_string(), Box::new(inner)));
        }
        if self.eat_operator("+") {
            return self.parse_unary();
        }
        self.parse_postfix()
    }

    fn parse_postfix(&mut self) -> Option<Node> {
        let mut node = self.parse_primary()?;
        // The percent sign is a postfix operator on whatever precedes it.
        while self.eat_operator("%") {
            node = Node::Unary("%".to_string(), Box::new(node));
        }
        Some(node)
    }

    fn parse_primary(&mut self) -> Option<Node> {
        match self.peek().cloned() {
            Some(Token::Number(value)) => {
                self.position += 1;
                Some(Node::Literal(CellValue::Number(value)))
            }
            Some(Token::Text(text)) => {
                self.position += 1;
                Some(Node::Literal(CellValue::Text(text)))
            }
            Some(Token::Word(word)) => {
                self.position += 1;
                match word.as_str() {
                    "TRUE" => Some(Node::Literal(CellValue::Bool(true))),
                    "FALSE" => Some(Node::Literal(CellValue::Bool(false))),
                    other => Some(Node::Literal(CellValue::Error(other.to_string()))),
                }
            }
            Some(Token::OpenParen) => {
                self.position += 1;
                let inner = self.parse_expression()?;
                if !matches!(self.take(), Some(Token::CloseParen)) {
                    return None;
                }
                Some(inner)
            }
            Some(Token::Name(name)) => {
                self.position += 1;
                self.parse_name(name)
            }
            _ => None,
        }
    }

    /// A name: a function call, a reference, or a sheet-qualified reference.
    fn parse_name(&mut self, name: String) -> Option<Node> {
        // A `(` right after the name makes it a call. Anything else is a reference.
        if matches!(self.peek(), Some(Token::OpenParen)) {
            self.position += 1;
            let mut arguments = Vec::new();
            if !matches!(self.peek(), Some(Token::CloseParen)) {
                loop {
                    arguments.push(self.parse_expression()?);
                    match self.peek() {
                        Some(Token::Comma) => {
                            self.position += 1;
                        }
                        _ => break,
                    }
                }
            }
            if !matches!(self.take(), Some(Token::CloseParen)) {
                return None;
            }
            return Some(Node::Function(name.to_uppercase(), arguments));
        }

        // `Sheet1!A1`, where the sheet may itself be quoted or reached through a path.
        let mut sheet = None;
        let mut body = name;
        while matches!(self.peek(), Some(Token::Bang)) {
            self.position += 1;
            let next = match self.take() {
                Some(Token::Name(second)) => second,
                Some(Token::Word(second)) => second,
                _ => return None,
            };
            sheet = Some(match sheet {
                Some(previous) => format!("{previous}!{body}"),
                None => body,
            });
            body = next;
        }

        // `A1:B2`, and the sheet-qualified form of the same.
        if matches!(self.peek(), Some(Token::Colon)) {
            self.position += 1;
            let second = match self.take() {
                Some(Token::Name(second)) => second,
                _ => return None,
            };
            let mut range = format!("{body}:{second}");
            // A whole-row or whole-column range names cells this module does not enumerate, so
            // it is reported rather than silently reading one row.
            if range.contains(':') && range.split(':').any(|side| !looks_like_reference(side)) {
                range.push(' ');
            }
            return Some(Node::Reference(sheet, range));
        }

        Some(Node::Reference(sheet, body))
    }
}

// ---------------------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------------------

/// Why a formula could not produce a value.
type Reason = String;

fn error_value(code: &str) -> CellValue {
    CellValue::Error(code.to_string())
}

/// What a referenced cell contributes as a *value*.
///
/// A cell holding a formula reads as blank. This module has no recalculation order, so
/// evaluating its dependency would mean inventing one; and returning the formula's text would be
/// worse still, since `=A1+1` would then try to add one to the string `"=1+1"`. Blank is
/// visible, and [`Workbook::recalculate`][crate::Workbook::recalculate] puts the value there once the order is known.
fn readable(value: CellValue) -> CellValue {
    match value {
        CellValue::Formula(_) => CellValue::None,
        other => other,
    }
}

fn is_error(value: &CellValue) -> bool {
    matches!(value, CellValue::Error(_))
}

fn eval(
    node: &Node,
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, Reason> {
    // An error anywhere propagates, and does so before the rest of the node is even looked at,
    // which is what makes `1/0+SUM(A1:A2)` yield `#DIV/0!` rather than something surprising.
    match node {
        Node::Literal(value) => Ok(value.clone()),
        Node::Reference(reference_sheet, body) => {
            let resolved = resolve(reference_sheet.as_deref().unwrap_or(sheet), body, source)?;
            match resolved {
                Resolved::Single(value) => Ok(readable(value)),
                // A range where a single value is needed takes the first cell, which is what
                // Excel's implicit intersection does for a one-row range.
                Resolved::Range(mut values) => Ok(if values.is_empty() {
                    CellValue::None
                } else {
                    values.remove(0)
                }),
            }
        }
        Node::Unary(operator, inner) => {
            let value = eval(inner, sheet, source)?;
            unary(operator, value)
        }
        Node::Binary(operator, left, right) => {
            let a = eval(left, sheet, source)?;
            let b = eval(right, sheet, source)?;
            binary(operator, a, b)
        }
        Node::Function(name, arguments) => {
            if !supports(name) {
                return Err(format!("unsupported function {name}()"));
            }
            call_function(name, arguments, sheet, source)
        }
    }
}

/// Resolve a reference body against a sheet.
fn resolve(
    sheet: &str,
    body: &str,
    source: &dyn ValueSource,
) -> std::result::Result<Resolved, Reason> {
    let body = body.trim();
    let (start, end) = match body.split_once(':') {
        Some((start, end)) => (start, Some(end)),
        None => (body, None),
    };
    let start = normalise_reference(start)?;
    match end {
        None => Ok(Resolved::Single(readable(source.cell(sheet, &start)))),
        Some(end) => {
            let range = crate::worksheet::cell_range::CellRange::parse(&format!(
                "{start}:{}",
                normalise_reference(end)?
            ))
            .map_err(|error| format!("cannot read the range {body:?}: {error}"))?;
            let mut values = Vec::new();
            for coordinate in range.cells().map_err(|error| error.to_string())? {
                values.push(readable(source.cell(sheet, &coordinate)));
            }
            Ok(Resolved::Range(values))
        }
    }
}

/// Turn `a1` or `$A$1` into `A1`, the form the source is keyed by.
fn normalise_reference(text: &str) -> std::result::Result<String, Reason> {
    let cleaned: String = text.chars().filter(|c| *c != '$').collect();
    if !looks_like_reference(&cleaned) {
        return Err(format!("{text:?} is not a cell reference"));
    }
    Ok(cleaned.to_uppercase())
}

fn unary(operator: &str, value: CellValue) -> std::result::Result<CellValue, Reason> {
    if is_error(&value) {
        return Ok(value);
    }
    match operator {
        "%" => to_number(&value).map(|n| CellValue::Number(n / 100.0)),
        "-" => to_number(&value).map(|n| CellValue::Number(-n)),
        other => Err(format!("unknown operator {other}")),
    }
}

fn binary(
    operator: &str,
    left: CellValue,
    right: CellValue,
) -> std::result::Result<CellValue, Reason> {
    if let CellValue::Error(code) = &left {
        return Ok(error_value(code));
    }
    if let CellValue::Error(code) = &right {
        return Ok(error_value(code));
    }
    match operator {
        "&" => {
            // Concatenation renders a number the way Excel shows it, not as a raw float:
            // `1/3&""` is `0.333333333333333`, and a trailing `.0` never appears.
            Ok(CellValue::Text(format!(
                "{}{}",
                as_display(&left),
                as_display(&right)
            )))
        }
        "=" | "<>" | "<" | "<=" | ">" | ">=" => {
            let ordering = compare(&left, &right);
            let holds = match operator {
                "=" => ordering == std::cmp::Ordering::Equal,
                "<>" => ordering != std::cmp::Ordering::Equal,
                "<" => ordering == std::cmp::Ordering::Less,
                "<=" => ordering != std::cmp::Ordering::Greater,
                ">" => ordering == std::cmp::Ordering::Greater,
                _ => ordering != std::cmp::Ordering::Less,
            };
            Ok(CellValue::Bool(holds))
        }
        _ => {
            let a = to_number(&left)?;
            let b = to_number(&right)?;
            let value = match operator {
                "+" => a + b,
                "-" => a - b,
                "*" => a * b,
                "/" => {
                    if b == 0.0 {
                        return Ok(error_value("#DIV/0!"));
                    }
                    a / b
                }
                "^" => {
                    let result = a.powf(b);
                    if result.is_nan() {
                        return Ok(error_value("#NUM!"));
                    }
                    result
                }
                other => return Err(format!("unknown operator {other}")),
            };
            if value.is_infinite() {
                return Ok(error_value("#NUM!"));
            }
            Ok(CellValue::Number(value))
        }
    }
}

/// Excel's ordering across types: numbers before text before booleans.
fn rank(value: &CellValue) -> u8 {
    match value {
        CellValue::None => 0,
        CellValue::Number(_) => 1,
        CellValue::Text(_) => 2,
        CellValue::Bool(_) => 3,
        _ => 4,
    }
}

fn compare(left: &CellValue, right: &CellValue) -> std::cmp::Ordering {
    // A blank compares as zero against a number and as an empty string against text, which is
    // why `=A1=""` is true for an empty cell.
    let left = coerce_for_comparison(left);
    let right = coerce_for_comparison(right);
    match (&left, &right) {
        (CellValue::Number(a), CellValue::Number(b)) => {
            a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal)
        }
        (CellValue::Text(a), CellValue::Text(b)) => {
            // Excel compares text case-insensitively.
            a.to_lowercase().cmp(&b.to_lowercase())
        }
        (CellValue::Bool(a), CellValue::Bool(b)) => a.cmp(b),
        _ => rank(&left).cmp(&rank(&right)),
    }
}

fn coerce_for_comparison(value: &CellValue) -> CellValue {
    match value {
        CellValue::None => CellValue::Number(0.0),
        other => other.clone(),
    }
}

/// Excel's coercion to a number, and the `#VALUE!` it produces when there is none.
fn to_number(value: &CellValue) -> std::result::Result<f64, Reason> {
    match value {
        CellValue::Number(v) => Ok(*v),
        CellValue::Bool(v) => Ok(if *v { 1.0 } else { 0.0 }),
        // A blank is zero in arithmetic, which is what makes `=A1+1` work on an empty cell.
        CellValue::None => Ok(0.0),
        CellValue::Text(text) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                return Ok(0.0);
            }
            let cleaned = trimmed.replace(['$', ','], "");
            // Excel also reads a percentage and a date literal here; those are left out rather
            // than guessed at, and the caller sees `#VALUE!` as it would from a bad entry.
            cleaned
                .parse::<f64>()
                .map_err(|_| format!("#VALUE! cannot read {text:?} as a number"))
        }
        CellValue::DateTime(moment) => Ok(crate::date_time::to_excel(
            *moment,
            crate::date_time::BaseDate::Windows1900,
        )),
        CellValue::Date(day) => Ok(crate::date_time::date_to_excel(
            *day,
            crate::date_time::BaseDate::Windows1900,
        )),
        CellValue::Time(clock) => Ok(crate::date_time::time_to_days(*clock)),
        CellValue::Duration(span) => Ok(crate::date_time::timedelta_to_days(*span)),
        CellValue::Error(code) => Err(code.clone()),
        CellValue::Formula(_) => Err("#VALUE! a formula used as a value".to_string()),
    }
}

fn to_number_value(value: &CellValue) -> CellValue {
    match to_number(value) {
        Ok(number) => CellValue::Number(number),
        Err(code) => error_value(code.trim_start_matches("#VALUE! ")),
    }
}

/// The text a value renders as, for `&` and the text functions.
fn as_display(value: &CellValue) -> String {
    match value {
        CellValue::Text(text) => text.clone(),
        CellValue::Bool(v) => if *v { "TRUE" } else { "FALSE" }.to_string(),
        CellValue::None => String::new(),
        CellValue::Number(v) => format_number(*v),
        CellValue::DateTime(moment) => moment.format("%Y-%m-%d %H:%M:%S").to_string(),
        other => format!("{other:?}"),
    }
}

/// Render a number the way a spreadsheet cell shows it: no trailing `.0`, no float noise.
fn format_number(value: f64) -> String {
    if value == value.trunc() && value.abs() < 1e15 {
        return format!("{}", value as i64);
    }
    let mut text = format!("{value}");
    if text.contains('e') {
        text = format!("{value:.10}");
    }
    text
}

// ---------------------------------------------------------------------------------------
// Functions
// ---------------------------------------------------------------------------------------

/// The values a function argument evaluates to, ranges flattened and blanks kept.
///
/// Keeping blanks matters: `COUNT` and `SUM` treat a blank differently from a zero, and only one
/// of them can be right for any given function.
fn argument_values(
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<Vec<CellValue>, Reason> {
    let mut values = Vec::new();
    for argument in arguments {
        match argument {
            Node::Reference(reference_sheet, body) => {
                match resolve(reference_sheet.as_deref().unwrap_or(sheet), body, source)? {
                    Resolved::Single(value) => values.push(value),
                    Resolved::Range(range) => values.extend(range),
                }
            }
            other => values.push(eval(other, sheet, source)?),
        }
    }
    Ok(values)
}

/// The numbers an aggregate should add, skipping blanks and text inside ranges.
///
/// `SUM(A1:A3)` ignores a text cell in the range but `SUM("x")` is `#VALUE!`, because a literal
/// was clearly meant as a number and a cell's contents were not.
fn aggregate_numbers(
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<Vec<f64>, Reason> {
    let mut numbers = Vec::new();
    for argument in arguments {
        match argument {
            Node::Reference(reference_sheet, body) => {
                let sheet = reference_sheet.as_deref().unwrap_or(sheet);
                match resolve(sheet, body, source)? {
                    Resolved::Single(value) => {
                        // A lone cell reference follows the range rule: blank and text are
                        // skipped rather than being an error.
                        if let CellValue::Number(number) = value {
                            numbers.push(number);
                        }
                    }
                    Resolved::Range(values) => {
                        for value in values {
                            if let CellValue::Number(number) = value {
                                numbers.push(number);
                            }
                        }
                    }
                }
            }
            other => {
                let value = eval(other, sheet, source)?;
                numbers.push(to_number(&value)?);
            }
        }
    }
    Ok(numbers)
}

fn call_function(
    name: &str,
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, Reason> {
    // The functions that need their arguments unevaluated -- IF and IFERROR -- are dispatched
    // before anything flattens them, so `IF(A1=0,"",1/A1)` does not divide by zero on the
    // branch it does not take.
    match name {
        "IF" => return call_if(arguments, sheet, source),
        "IFERROR" => return call_iferror(arguments, sheet, source),
        "AND" | "OR" => return call_and_or(name, arguments, sheet, source),
        _ => {}
    }

    match name {
        "SUM" => {
            let numbers = aggregate_numbers(arguments, sheet, source)?;
            Ok(CellValue::Number(numbers.iter().sum()))
        }
        "PRODUCT" => {
            let numbers = aggregate_numbers(arguments, sheet, source)?;
            Ok(CellValue::Number(numbers.iter().product()))
        }
        "AVERAGE" => {
            let numbers = aggregate_numbers(arguments, sheet, source)?;
            if numbers.is_empty() {
                // Excel's answer for an empty average, which is not zero.
                return Ok(error_value("#DIV/0!"));
            }
            Ok(CellValue::Number(
                numbers.iter().sum::<f64>() / numbers.len() as f64,
            ))
        }
        "MIN" | "MAX" => {
            let numbers = aggregate_numbers(arguments, sheet, source)?;
            if numbers.is_empty() {
                return Ok(CellValue::Number(0.0));
            }
            let picked = if name == "MIN" {
                numbers.iter().copied().fold(f64::INFINITY, f64::min)
            } else {
                numbers.iter().copied().fold(f64::NEG_INFINITY, f64::max)
            };
            Ok(CellValue::Number(picked))
        }
        "COUNT" => {
            let values = argument_values(arguments, sheet, source)?;
            Ok(CellValue::Number(
                values
                    .iter()
                    .filter(|value| {
                        matches!(
                            value,
                            CellValue::Number(_) | CellValue::Bool(_) | CellValue::DateTime(_)
                        )
                    })
                    .count() as f64,
            ))
        }
        "COUNTA" => {
            let values = argument_values(arguments, sheet, source)?;
            Ok(CellValue::Number(
                values.iter().filter(|value| !value.is_empty()).count() as f64,
            ))
        }
        "ABS" | "INT" | "SIGN" | "SQRT" | "ROUND" | "ROUNDUP" | "ROUNDDOWN" | "MOD" | "POWER" => {
            let values = argument_values(arguments, sheet, source)?;
            if values.is_empty() {
                return Err(format!("{name}() needs an argument"));
            }
            let first = to_number_value(&values[0]);
            if is_error(&first) {
                return Ok(first);
            }
            let number = match first {
                CellValue::Number(n) => n,
                _ => 0.0,
            };
            match name {
                "ABS" => Ok(CellValue::Number(number.abs())),
                "INT" => Ok(CellValue::Number(number.floor())),
                "SIGN" => Ok(CellValue::Number(number.signum())),
                "SQRT" => {
                    if number < 0.0 {
                        return Ok(error_value("#NUM!"));
                    }
                    Ok(CellValue::Number(number.sqrt()))
                }
                "ROUND" | "ROUNDUP" | "ROUNDDOWN" => {
                    let places = match values.get(1) {
                        Some(value) => to_number(value)?,
                        // Excel's default is zero places, and rounding to zero places is the
                        // common case by far.
                        None => 0.0,
                    };
                    if is_error(&to_number_value(values.get(1).unwrap_or(&CellValue::None))) {
                        return Ok(to_number_value(values.get(1).unwrap_or(&CellValue::None)));
                    }
                    Ok(CellValue::Number(round(number, places, name)))
                }
                "MOD" => {
                    let divisor = match values.get(1) {
                        Some(value) => to_number(value)?,
                        None => return Err("MOD() needs a divisor".to_string()),
                    };
                    if divisor == 0.0 {
                        return Ok(error_value("#DIV/0!"));
                    }
                    // Excel's MOD follows the divisor's sign, unlike Rust's `%`.
                    Ok(CellValue::Number(
                        number - divisor * (number / divisor).floor(),
                    ))
                }
                "POWER" => {
                    let exponent = match values.get(1) {
                        Some(value) => to_number(value)?,
                        None => return Err("POWER() needs an exponent".to_string()),
                    };
                    let result = number.powf(exponent);
                    if result.is_nan() {
                        return Ok(error_value("#NUM!"));
                    }
                    Ok(CellValue::Number(result))
                }
                _ => unreachable!("the match above is exhaustive"),
            }
        }
        "CONCAT" | "CONCATENATE" => {
            let values = argument_values(arguments, sheet, source)?;
            if let Some(bad) = values.iter().find(|value| is_error(value)) {
                return Ok(bad.clone());
            }
            let mut text = String::new();
            for value in values {
                match value {
                    // A blank contributes nothing rather than a zero, which is what makes
                    // `CONCAT(A1:A3)` usable on a sparse column.
                    CellValue::None => {}
                    other => text.push_str(&as_display(&other)),
                }
            }
            Ok(CellValue::Text(text))
        }
        "LEFT" | "RIGHT" => {
            let values = argument_values(arguments, sheet, source)?;
            let text = as_display(values.first().unwrap_or(&CellValue::None));
            let count = match values.get(1) {
                Some(value) => to_number(value)?.max(0.0) as usize,
                None => 1,
            };
            let characters: Vec<char> = text.chars().collect();
            let taken: String = if name == "LEFT" {
                characters.iter().take(count).collect()
            } else {
                characters
                    .iter()
                    .skip(characters.len().saturating_sub(count))
                    .collect()
            };
            Ok(CellValue::Text(taken))
        }
        "LEN" => {
            let values = argument_values(arguments, sheet, source)?;
            let text = as_display(values.first().unwrap_or(&CellValue::None));
            // Excel counts characters, not bytes, so an emoji is one character.
            Ok(CellValue::Number(text.chars().count() as f64))
        }
        "TRIM" | "UPPER" | "LOWER" => {
            let values = argument_values(arguments, sheet, source)?;
            let text = as_display(values.first().unwrap_or(&CellValue::None));
            Ok(CellValue::Text(match name {
                // Excel's TRIM also collapses internal runs of spaces, not just the ends.
                "TRIM" => text.split_whitespace().collect::<Vec<_>>().join(" "),
                "UPPER" => text.to_uppercase(),
                _ => text.to_lowercase(),
            }))
        }
        "NOT" => {
            let values = argument_values(arguments, sheet, source)?;
            match to_bool(values.first().unwrap_or(&CellValue::None))? {
                Some(flag) => Ok(CellValue::Bool(!flag)),
                None => Ok(CellValue::Number(0.0)),
            }
        }
        "TRUE" => Ok(CellValue::Bool(true)),
        "FALSE" => Ok(CellValue::Bool(false)),
        other => Err(format!("unsupported function {other}()")),
    }
}

fn call_if(
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, Reason> {
    let Some(condition) = arguments.first() else {
        return Err("IF() needs a condition".to_string());
    };
    let value = eval(condition, sheet, source)?;
    if is_error(&value) {
        // An error in the condition propagates; it is not a false.
        return Ok(value);
    }
    let flag = to_bool(&value)?;
    match flag {
        Some(true) => match arguments.get(1) {
            Some(branch) => eval(branch, sheet, source),
            // `IF(A1,)` is a zero-length string, not an error.
            None => Ok(CellValue::Text(String::new())),
        },
        Some(false) => match arguments.get(2) {
            Some(branch) => eval(branch, sheet, source),
            // Excel returns FALSE for a missing false-branch, not zero and not an empty string.
            None => Ok(CellValue::Bool(false)),
        },
        None => Ok(CellValue::Number(0.0)),
    }
}

fn call_iferror(
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, Reason> {
    let Some(value_node) = arguments.first() else {
        return Err("IFERROR() needs a value".to_string());
    };
    match eval(value_node, sheet, source) {
        Ok(value) if is_error(&value) => match arguments.get(1) {
            Some(fallback) => eval(fallback, sheet, source),
            None => Ok(CellValue::Text(String::new())),
        },
        other => other,
    }
}

fn call_and_or(
    name: &str,
    arguments: &[Node],
    sheet: &str,
    source: &dyn ValueSource,
) -> std::result::Result<CellValue, Reason> {
    if arguments.is_empty() {
        return Err(format!("{name}() needs an argument"));
    }
    // AND needs "every value is true" and OR needs "some value is true", so both are tracked.
    // Folding them into one flag gets one of the two wrong, which is what happened first time.
    let mut any_true = false;
    let mut all_true = true;
    let mut seen = false;
    for argument in arguments {
        // A range contributes each of its values, so `AND(A1:A3)` checks three cells.
        let values = match argument {
            Node::Reference(reference_sheet, body) => {
                match resolve(reference_sheet.as_deref().unwrap_or(sheet), body, source)? {
                    Resolved::Single(value) => vec![value],
                    Resolved::Range(values) => values,
                }
            }
            other => vec![eval(other, sheet, source)?],
        };
        for value in values {
            if is_error(&value) {
                return Ok(value);
            }
            if let CellValue::Text(text) = &value {
                // Text in a reference is ignored by AND and OR, matching Excel; a literal
                // `AND("x")` is still `#VALUE!`.
                if matches!(argument, Node::Reference(_, _)) {
                    continue;
                }
                return Err(format!("#VALUE! cannot read {text:?} as a condition"));
            }
            // A blank is not a condition: it neither satisfies nor breaks either, which is
            // what `seen` tracks.
            if let Some(flag) = to_bool(&value)? {
                seen = true;
                any_true |= flag;
                all_true &= flag;
            }
        }
    }
    if !seen {
        // Excel's AND() over nothing but blanks is TRUE, and OR() is FALSE.
        return Ok(CellValue::Bool(name == "AND"));
    }
    Ok(CellValue::Bool(if name == "AND" {
        all_true
    } else {
        any_true
    }))
}

/// Excel's coercion to a condition: `None` for a blank, which the callers treat as zero.
fn to_bool(value: &CellValue) -> std::result::Result<Option<bool>, Reason> {
    match value {
        CellValue::Bool(flag) => Ok(Some(*flag)),
        CellValue::Number(number) => Ok(Some(*number != 0.0)),
        CellValue::None => Ok(None),
        CellValue::Text(text) => Err(format!("#VALUE! cannot read {text:?} as a condition")),
        CellValue::Error(code) => Err(code.clone()),
        other => Err(format!("#VALUE! cannot use {other:?} as a condition")),
    }
}

/// Excel rounds half away from zero, which Rust's `round` does not do for negatives.
fn round(value: f64, places: f64, mode: &str) -> f64 {
    let factor = 10f64.powi(places as i32);
    let scaled = value * factor;
    let rounded = match mode {
        "ROUNDUP" => {
            if scaled >= 0.0 {
                scaled.ceil()
            } else {
                scaled.floor()
            }
        }
        "ROUNDDOWN" => {
            if scaled >= 0.0 {
                scaled.floor()
            } else {
                scaled.ceil()
            }
        }
        _ => {
            if scaled >= 0.0 {
                (scaled + 0.5).floor()
            } else {
                // Away from zero, so -2.5 rounds to -3 rather than -2.
                (scaled - 0.5).ceil()
            }
        }
    };
    rounded / factor
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::BTreeMap;

    /// A source backed by a map, so a test can write `A1=1` and read what the formula made of it.
    struct Map(BTreeMap<String, CellValue>);

    impl Map {
        fn new() -> Self {
            let mut cells = BTreeMap::new();
            for (coordinate, value) in [
                ("A1", CellValue::Number(1.0)),
                ("A2", CellValue::Number(2.0)),
                ("A3", CellValue::Number(3.0)),
                ("B1", CellValue::Text("hello".to_string())),
                ("B2", CellValue::Text("  padded  ".to_string())),
                ("C1", CellValue::Bool(true)),
                ("C2", CellValue::Bool(false)),
                ("A4", CellValue::Number(0.0)),
            ] {
                cells.insert(coordinate.to_string(), value);
            }
            Map(cells)
        }

        fn set(&mut self, coordinate: &str, value: CellValue) {
            self.0.insert(coordinate.to_string(), value);
        }
    }

    impl ValueSource for Map {
        fn cell(&self, _sheet: &str, coordinate: &str) -> CellValue {
            self.0.get(coordinate).cloned().unwrap_or(CellValue::None)
        }
    }

    fn value(formula: &str) -> CellValue {
        let source = Map::new();
        evaluate(formula, "Sheet1", &source).unwrap_or_else(|reason| panic!("{formula}: {reason}"))
    }

    /// The reason a formula could not be evaluated, if it could not be.
    fn reason(formula: &str) -> Option<String> {
        let source = Map::new();
        evaluate(formula, "Sheet1", &source).err()
    }

    fn number(formula: &str) -> f64 {
        match value(formula) {
            CellValue::Number(v) => v,
            other => panic!("{formula} gave {other:?}, not a number"),
        }
    }

    fn text(formula: &str) -> String {
        match value(formula) {
            CellValue::Text(v) => v,
            other => panic!("{formula} gave {other:?}, not text"),
        }
    }

    #[test]
    fn the_supported_list_has_no_names_the_dispatcher_lacks() {
        // A name listed but not implemented would answer `supports` with a yes and then
        // produce nothing, which is the one failure mode this list exists to prevent.
        for name in SUPPORTED_FUNCTIONS {
            let node = Node::Function(
                (*name).to_string(),
                vec![Node::Literal(CellValue::Number(1.0))],
            );
            let source = Map::new();
            let result = eval(&node, "Sheet1", &source);
            assert!(
                !matches!(result, Err(reason) if reason.starts_with("unsupported function")),
                "{name} is listed but the dispatcher refuses it"
            );
        }
        assert_eq!(SUPPORTED_FUNCTIONS.len(), 31, "the list grew or shrank");
    }

    #[test]
    fn arithmetic_follows_the_usual_precedence() {
        assert_eq!(number("=1+2*3"), 7.0);
        assert_eq!(number("=(1+2)*3"), 9.0);
        assert_eq!(number("=2^3^2"), 64.0, "Excel's ^ is left-associative");
    }

    #[test]
    fn unary_minus_binds_tighter_than_the_power() {
        // The classic Excel surprise: -2^2 is 4, because the minus applies to the 2 first.
        assert_eq!(number("=-2^2"), 4.0);
        assert_eq!(number("=-(2^2)"), -4.0);
    }

    #[test]
    fn percentages_and_blanks_behave_as_they_do_in_excel() {
        assert_eq!(number("=50%"), 0.5);
        assert_eq!(number("=E1+1"), 1.0, "a blank is zero in arithmetic");
        assert_eq!(number("=SUM(E1:E3)"), 0.0);
    }

    #[test]
    fn division_by_zero_is_an_error_rather_than_an_infinity() {
        assert!(matches!(value("=1/0"), CellValue::Error(code) if code == "#DIV/0!"));
    }

    #[test]
    fn an_error_anywhere_propagates() {
        assert!(matches!(value("=1/0+SUM(A1:A3)"), CellValue::Error(code) if code == "#DIV/0!"));
        assert!(matches!(value("=SUM(A1:A3)+1/0"), CellValue::Error(code) if code == "#DIV/0!"));
    }

    #[test]
    fn ranges_aggregate_over_their_cells() {
        assert_eq!(number("=SUM(A1:A3)"), 6.0);
        assert_eq!(number("=AVERAGE(A1:A3)"), 2.0);
        assert_eq!(number("=MIN(A1:A3)"), 1.0);
        assert_eq!(number("=MAX(A1:A3)"), 3.0);
        assert_eq!(number("=COUNT(A1:B3)"), 3.0, "text is not counted");
        // A1:B3 holds three numbers, two strings, and an empty B3.
        assert_eq!(
            number("=COUNTA(A1:B3)"),
            5.0,
            "COUNTA counts the text but not the blank"
        );
    }

    #[test]
    fn aggregating_an_empty_range_follows_excel() {
        // Not zero: Excel's AVERAGE of nothing is #DIV/0!.
        assert!(matches!(value("=AVERAGE(E1:E3)"), CellValue::Error(code) if code == "#DIV/0!"));
        assert_eq!(number("=SUM(E1:E3)"), 0.0);
        assert_eq!(number("=MIN(E1:E3)"), 0.0);
    }

    #[test]
    fn only_the_branch_if_takes_is_evaluated() {
        // Evaluating both branches is how an engine produces #DIV/0! for a formula Excel is
        // perfectly happy with. A3 is the cell holding the zero, and the division by it only
        // appears on the branch the formula does not take.
        assert_eq!(
            number("=IF(A4=0,0,1/A4)"),
            0.0,
            "the true branch avoids the division"
        );
        // The other way round: the false branch is the one that divides, and here it *is*
        // taken, so the error has to come through rather than being swallowed by laziness.
        assert!(
            matches!(value("=IF(A4<>0,0,1/A4)"), CellValue::Error(code) if code == "#DIV/0!"),
            "an error on the branch that is taken still propagates"
        );
        assert!(
            !matches!(value("=IF(A4=0,0,1/A4)"), CellValue::Error(_)),
            "the skipped branch must not raise #DIV/0!"
        );
        assert_eq!(
            number("=IF(A1=0,0,1/A1)"),
            1.0,
            "the other branch is taken normally"
        );
    }

    #[test]
    fn if_returns_false_when_the_else_branch_is_missing() {
        assert!(matches!(value("=IF(FALSE,1)"), CellValue::Bool(false)));
    }

    #[test]
    fn comparison_orders_types_with_numbers_first() {
        assert!(matches!(value("=1<\"a\""), CellValue::Bool(true)));
        assert!(matches!(value("=\"a\"<TRUE"), CellValue::Bool(true)));
        // Text compares case-insensitively.
        assert!(matches!(value("=\"ABC\"=\"abc\""), CellValue::Bool(true)));
        assert!(matches!(value("=\"a\"=\"b\""), CellValue::Bool(false)));
    }

    #[test]
    fn a_blank_compares_as_zero_against_a_number() {
        assert!(matches!(value("=E1=0"), CellValue::Bool(true)));
    }

    #[test]
    fn text_functions_behave() {
        assert_eq!(text("=UPPER(B1)"), "HELLO");
        assert_eq!(text("=LOWER(B1)"), "hello");
        assert_eq!(text("=LEFT(B1,2)"), "he");
        assert_eq!(text("=RIGHT(B1,3)"), "llo");
        assert_eq!(text("=LEFT(B1)"), "h");
        assert_eq!(number("=LEN(B1)"), 5.0);
        // TRIM collapses internal runs too, not just the ends.
        assert_eq!(text("=TRIM(B2)"), "padded");
    }

    #[test]
    fn concatenation_renders_numbers_the_way_a_cell_shows_them() {
        // A raw float would give `2.5` correctly but `2` as `2.0`, which is visible.
        assert_eq!(text("=1+1&\"\""), "2");
        assert_eq!(text("=\"n=\"&1&\"\" "), "n=1");
        assert_eq!(text("=A1&B1"), "1hello");
    }

    #[test]
    fn rounding_goes_away_from_zero_on_a_half() {
        assert_eq!(number("=ROUND(2.5)"), 3.0);
        assert_eq!(number("=ROUND(-2.5)"), -3.0, "Rust's round would say -2");
        assert_eq!(number("=ROUND(2.345,2)"), 2.35);
        assert_eq!(number("=ROUNDUP(2.1)"), 3.0);
        assert_eq!(number("=ROUNDDOWN(2.9)"), 2.0);
    }

    #[test]
    fn mod_follows_the_divisors_sign() {
        assert_eq!(number("=MOD(-3,2)"), 1.0);
        assert_eq!(number("=MOD(3,-2)"), -1.0);
        assert!(matches!(value("=MOD(1,0)"), CellValue::Error(code) if code == "#DIV/0!"));
    }

    #[test]
    fn iferror_replaces_an_error_and_leaves_everything_else_alone() {
        assert_eq!(number("=IFERROR(1/0,7)"), 7.0);
        assert_eq!(number("=IFERROR(1/2,7)"), 0.5);
    }

    #[test]
    fn and_or_and_not_read_conditions() {
        assert!(matches!(value("=AND(TRUE,TRUE)"), CellValue::Bool(true)));
        assert!(matches!(value("=AND(TRUE,FALSE)"), CellValue::Bool(false)));
        assert!(matches!(value("=OR(FALSE,TRUE)"), CellValue::Bool(true)));
        assert!(matches!(value("=NOT(TRUE)"), CellValue::Bool(false)));
        // A range contributes each cell.
        assert!(matches!(value("=AND(C1:C2)"), CellValue::Bool(false)));
    }

    #[test]
    fn an_unsupported_function_is_reported_rather_than_guessed_at() {
        // The whole reason this module can be trusted: a formula it does not understand gets
        // no value, not a plausible one.
        let source = Map::new();
        let result = evaluate("=VLOOKUP(A1,B1:C3,2,FALSE)", "Sheet1", &source);
        assert_eq!(
            result,
            Err("unsupported function VLOOKUP()".to_string()),
            "VLOOKUP must not produce a value"
        );
    }

    #[test]
    fn a_sheet_qualified_reference_is_read() {
        let source = Map::new();
        let value = evaluate("=Other!A1", "Other", &source).expect("A1 on that sheet");
        assert_eq!(value, CellValue::Number(1.0));
    }

    #[test]
    fn absolute_references_are_accepted() {
        let source = Map::new();
        let value = evaluate("=SUM($A$1:$A$3)", "Sheet1", &source).expect("the range");
        assert_eq!(value, CellValue::Number(6.0));
    }

    #[test]
    fn text_is_parsed_with_doubled_quotes_inside_it() {
        assert_eq!(text("=\"say \"\"hi\"\"\""), "say \"hi\"");
    }

    #[test]
    fn a_formula_that_cannot_be_parsed_is_an_error_not_a_value() {
        let source = Map::new();
        assert!(evaluate("=1+", "Sheet1", &source).is_err());
        assert!(evaluate("=SUM(", "Sheet1", &source).is_err());
        assert!(evaluate("=1 1", "Sheet1", &source).is_err());
    }

    #[test]
    fn the_formula_may_be_written_with_or_without_its_leading_equals() {
        let source = Map::new();
        assert_eq!(
            evaluate("1+1", "Sheet1", &source).expect("no leading equals"),
            CellValue::Number(2.0)
        );
        assert_eq!(
            evaluate("=1+1", "Sheet1", &source).expect("a leading equals"),
            CellValue::Number(2.0)
        );
    }

    #[test]
    fn a_cell_reference_to_a_cell_holding_a_formula_is_not_followed() {
        // Reading it would need the recalculation order, which this module does not have. A
        // missing value is visible; a stale one is not.
        let mut source = Map::new();
        source.set("D1", CellValue::Formula("=1+1".to_string()));
        assert_eq!(
            evaluate("=D1", "Sheet1", &source).expect("the reference itself is fine"),
            CellValue::None,
            "a cell holding a formula reads as blank, not as its unevaluated text"
        );
    }

    #[test]
    fn scientific_notation_is_read() {
        assert_eq!(number("=1.5E2"), 150.0);
    }

    #[test]
    fn functions_and_references_are_told_apart_by_the_bracket() {
        // `SUM` with no bracket is a name, and a name is not a value.
        let source = Map::new();
        assert!(
            evaluate("=SUM", "Sheet1", &source).is_err(),
            "a bare name is not a value"
        );
    }

    #[test]
    fn a_whole_column_reference_is_reported_rather_than_partly_read() {
        // Reading `A:A` as one row would produce a number for a formula about a million rows.
        let source = Map::new();
        assert!(
            evaluate("=SUM(A:A)", "Sheet1", &source).is_err(),
            "a whole-column reference is reported, not read as one row"
        );
        assert!(reason("=SUM(A:A)").is_some());
    }

    #[test]
    fn recalculation_counts_and_reasons_are_reported() {
        let mut sheet = BTreeMap::new();
        sheet.insert("A1".to_string(), CellValue::Number(1.0));
        let mut computed = BTreeMap::new();
        computed.insert("Sheet1".to_string(), sheet);
        let report = Recalculation {
            computed,
            unresolved: vec![
                Unresolved {
                    sheet: "Sheet1".into(),
                    coordinate: "A2".into(),
                    formula: "=VLOOKUP(1,A:B,2)".into(),
                    reason: "unsupported function VLOOKUP()".into(),
                },
                Unresolved {
                    sheet: "Sheet1".into(),
                    coordinate: "A3".into(),
                    formula: "=XLOOKUP(1,A:B)".into(),
                    reason: "unsupported function VLOOKUP()".into(),
                },
            ],
        };
        assert_eq!(report.computed_count(), 1);
        assert!(!report.is_complete());
        assert_eq!(
            report.reasons(),
            vec![("unsupported function VLOOKUP()".to_string(), 2)]
        );
    }
}
