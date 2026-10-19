//! Checking values against a sheet's data validation rules: what Excel does
//! when a value is typed into a validated cell, and its Data - Data
//! Validation - Circle Invalid Data.
//!
//! The rules themselves are [`DataValidation`] on each sheet; this module
//! only reads them. Ported from the validation checks of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use crate::coordinate::{CellRef, shift_references};
use crate::formula::{Engine, Origin, Value};
use crate::model::{CellValue, DataValidation, Spreadsheet, ValidationOperator, ValidationType};
use crate::style::format::GENERAL;

/// The rule that governs `at`, if any. Where rules overlap, the first one
/// listed wins, as Excel shows only one.
#[must_use]
pub fn rule_at(book: &Spreadsheet, sheet: usize, at: CellRef) -> Option<&DataValidation> {
    book.sheet(sheet)?
        .data_validations
        .iter()
        .find(|rule| rule.sqref.iter().any(|r| r.contains(at)))
}

/// Whether `value`, entered at `at`, breaks the rule on that cell: `None`
/// when it is accepted (or there is no rule), the rule it breaks otherwise.
/// The rule's `error_style`, `error_title` and `error` say how Excel would
/// report it.
///
/// The rule's formulas are read as written for the first cell of its range
/// and moved to `at`, as Excel does, and a custom rule sees `value` in its
/// own cell. A rule whose `show_error_message` is off still answers: whether
/// to stop the entry is the caller's choice.
#[must_use]
pub fn check<'b>(
    book: &'b Spreadsheet,
    sheet: usize,
    at: CellRef,
    value: &CellValue,
) -> Option<&'b DataValidation> {
    let rule = rule_at(book, sheet, at)?;
    let stored = book
        .sheet(sheet)
        .and_then(|ws| ws.get(at))
        .map(|c| &c.value);
    // A custom rule reads the cell itself, so it must see the new value
    // there. Other rules only read their bounds and keep the book as it is.
    let entered;
    let view = if rule.kind == ValidationType::Custom && stored != Some(value) {
        let mut copy = book.clone();
        if let Some(ws) = copy.sheet_mut(sheet) {
            ws.set(at, value.clone());
        }
        entered = copy;
        &entered
    } else {
        book
    };
    (!accepts(view, sheet, at, rule, value)).then_some(rule)
}

/// Every cell of the sheet whose value breaks its rule: Excel's Circle
/// Invalid Data. Empty cells are not reported, in address order.
#[must_use]
pub fn invalid_cells(book: &Spreadsheet, sheet: usize) -> Vec<CellRef> {
    let Some(ws) = book.sheet(sheet) else {
        return Vec::new();
    };
    if ws.data_validations.is_empty() {
        return Vec::new();
    }
    ws.iter()
        .filter(|(_, cell)| !cell.value.result().is_empty())
        .filter(|(at, cell)| {
            rule_at(book, sheet, *at)
                .is_some_and(|rule| !accepts(book, sheet, *at, rule, &cell.value))
        })
        .map(|(at, _)| at)
        .collect()
}

/// The choices a list rule offers, as its drop-down shows them: the items
/// of a literal list (`"Yes,No"`), or the values of the range or name it
/// points at, empty ones left out. Empty for any other kind of rule.
#[must_use]
pub fn list_items(book: &Spreadsheet, sheet: usize, rule: &DataValidation) -> Vec<String> {
    if rule.kind != ValidationType::List {
        return Vec::new();
    }
    let source = rule.formula1.trim().trim_start_matches('=');
    let literal = |text: &str| -> Vec<String> {
        text.split(',')
            .map(|s| s.trim().to_owned())
            .filter(|s| !s.is_empty())
            .collect()
    };
    if let Some(inner) = source.strip_prefix('"').and_then(|s| s.strip_suffix('"')) {
        return literal(&inner.replace("\"\"", "\""));
    }
    let at = rule.sqref.first().map(|r| r.start).unwrap_or_default();
    let mut engine = Engine::new(book);
    match engine.eval(Origin::new(sheet, at), source) {
        Value::Array(rows) => rows
            .iter()
            .flatten()
            .filter(|v| !matches!(v, Value::Blank))
            .filter_map(|v| v.text().ok())
            .filter(|s| !s.is_empty())
            .collect(),
        // Something that is no formula at all: a list written without the
        // quotes some files leave off.
        Value::Error(_) => literal(source),
        v => v
            .text()
            .ok()
            .filter(|s| !s.is_empty())
            .into_iter()
            .collect(),
    }
}

/// Whether `rule` accepts `value` at `at`; `book` already holds it there
/// when the rule needs to see it.
fn accepts(
    book: &Spreadsheet,
    sheet: usize,
    at: CellRef,
    rule: &DataValidation,
    value: &CellValue,
) -> bool {
    let value = value.result();
    let blank = match value {
        CellValue::Empty => true,
        CellValue::Text(t) => t.is_empty(),
        _ => false,
    };
    if rule.kind == ValidationType::None {
        return true;
    }
    if blank {
        return rule.allow_blank;
    }
    let first = rule.sqref.first().map_or(at, |r| r.start);
    let moved = |formula: &str| {
        let formula = formula.trim().trim_start_matches('=');
        shift_references(
            formula,
            i64::from(at.col.index()) - i64::from(first.col.index()),
            i64::from(at.row.index()) - i64::from(first.row.index()),
        )
    };
    let mut engine = Engine::new(book);
    let mut bound = |formula: &str| -> Option<f64> {
        if formula.trim().is_empty() {
            return None;
        }
        engine
            .eval(Origin::new(sheet, at), &moved(formula))
            .number()
            .ok()
    };
    match rule.kind {
        ValidationType::None => true,
        ValidationType::List => {
            let shown = value.display(GENERAL, book.epoch).to_lowercase();
            list_items(book, sheet, rule)
                .iter()
                .any(|item| item.to_lowercase() == shown)
        }
        ValidationType::Custom => {
            let formula = moved(&rule.formula1);
            Engine::new(book)
                .eval(Origin::new(sheet, at), &formula)
                .boolean()
                .unwrap_or(false)
        }
        ValidationType::TextLength => {
            let length = value.display(GENERAL, book.epoch).chars().count();
            #[allow(clippy::cast_precision_loss, reason = "a cell holds 32 767 characters")]
            let length = length as f64;
            let low = bound(&rule.formula1);
            let high = bound(&rule.formula2);
            compare(rule.operator, length, low, high)
        }
        ValidationType::Whole
        | ValidationType::Decimal
        | ValidationType::Date
        | ValidationType::Time => {
            let CellValue::Number(n) = *value else {
                return false;
            };
            if rule.kind == ValidationType::Whole && n.fract() != 0.0 {
                return false;
            }
            let low = bound(&rule.formula1);
            let high = bound(&rule.formula2);
            compare(rule.operator, n, low, high)
        }
    }
}

/// `n` against the rule's bounds. A missing first bound, or a missing second
/// one for the two-sided operators, accepts nothing.
#[allow(clippy::float_cmp, reason = "Excel's equal is exact equality")]
fn compare(op: ValidationOperator, n: f64, low: Option<f64>, high: Option<f64>) -> bool {
    let Some(a) = low else {
        return false;
    };
    match op {
        ValidationOperator::Between => high.is_some_and(|b| n >= a.min(b) && n <= a.max(b)),
        ValidationOperator::NotBetween => high.is_some_and(|b| n < a.min(b) || n > a.max(b)),
        ValidationOperator::Equal => n == a,
        ValidationOperator::NotEqual => n != a,
        ValidationOperator::GreaterThan => n > a,
        ValidationOperator::LessThan => n < a,
        ValidationOperator::GreaterThanOrEqual => n >= a,
        ValidationOperator::LessThanOrEqual => n <= a,
    }
}
