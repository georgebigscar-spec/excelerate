//! What-If Analysis: Goal Seek, the Scenario Manager and data tables.
//!
//! The root finder was ported from the `data.goalSeek` command of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use crate::coordinate::{CellRef, Col, Range, Row};
use crate::error::{Error, Result};
use crate::formula::eval::{Engine, recalculate_from_with};
use crate::formula::{Origin, Value};
use crate::model::scenario::{Scenario, ScenarioInput};
use crate::model::{CellValue, Spreadsheet};
use crate::progress::Options;

/// Excel's defaults for Maximum Iterations and Maximum Change, which Goal
/// Seek reads from the workbook's calculation settings.
const DEFAULT_ITERATIONS: u32 = 100;
const DEFAULT_TOLERANCE: f64 = 0.001;

/// How Goal Seek ended.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Sought {
    /// The value left in the changing cell: the best one tried.
    pub value: f64,
    /// What the set cell computes with it. `NaN` when it computes an error
    /// or something that is not a number.
    pub result: f64,
    /// How many times the set cell was computed.
    pub iterations: u32,
    /// Whether `result` is within the tolerance of the target.
    pub converged: bool,
}

/// Limits of a search. [`Default`] is Excel's: 100 steps, within 0.001.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SeekLimits {
    /// The most times the set cell is computed.
    pub max_iterations: u32,
    /// How close to the target counts as reached.
    pub tolerance: f64,
}

impl Default for SeekLimits {
    fn default() -> Self {
        Self {
            max_iterations: DEFAULT_ITERATIONS,
            tolerance: DEFAULT_TOLERANCE,
        }
    }
}

impl SeekLimits {
    /// The limits the workbook's own calculation settings give
    /// (`iterateCount` and `iterateDelta` of `<calcPr>`), as Excel's Goal
    /// Seek uses them; Excel's defaults where the file says nothing.
    #[must_use]
    pub fn of(book: &Spreadsheet) -> Self {
        let read = |key: &str| {
            book.calculation_properties
                .iter()
                .find(|(k, _)| k == key)
                .map(|(_, v)| v.as_str())
        };
        Self {
            max_iterations: read("iterateCount")
                .and_then(|v| v.parse().ok())
                .filter(|&n| n > 0)
                .unwrap_or(DEFAULT_ITERATIONS),
            tolerance: read("iterateDelta")
                .and_then(|v| v.parse().ok())
                .filter(|&t: &f64| t.is_finite() && t > 0.0)
                .unwrap_or(DEFAULT_TOLERANCE),
        }
    }
}

/// Finds the value of the `changing` cell that makes the formula in `set`
/// compute `target`: Excel's Data - What-If Analysis - Goal Seek.
///
/// Both cells are `(sheet, cell)`. The search starts from what `changing`
/// holds now and stops after the workbook's Maximum Iterations or once the
/// result is within its Maximum Change ([`SeekLimits::of`]). Whether or not
/// it converged, the best value found is left in `changing` (its style is
/// kept) and every formula that reads it, `set` included, is recomputed.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a sheet the workbook does not have,
/// and [`Error::GoalSeek`] if `set` holds no formula, `changing` holds one,
/// or `target` is not a finite number. The workbook is unchanged then.
pub fn goal_seek(
    book: &mut Spreadsheet,
    set: (usize, CellRef),
    target: f64,
    changing: (usize, CellRef),
) -> Result<Sought> {
    let limits = SeekLimits::of(book);
    goal_seek_with(book, set, target, changing, limits, &Options::default())
}

/// [`goal_seek`] with limits of your own, and the caller's functions in
/// `options` for formulas that call them.
///
/// # Errors
/// As [`goal_seek`].
pub fn goal_seek_with(
    book: &mut Spreadsheet,
    set: (usize, CellRef),
    target: f64,
    changing: (usize, CellRef),
    limits: SeekLimits,
    options: &Options<'_>,
) -> Result<Sought> {
    if !target.is_finite() {
        return Err(Error::GoalSeek("the target must be a finite number".into()));
    }
    let value_at = |book: &Spreadsheet, (sheet, at): (usize, CellRef)| {
        book.sheet(sheet)
            .map(|ws| ws.get(at).map(|c| c.value.clone()).unwrap_or_default())
            .ok_or(Error::SheetIndexOutOfRange(sheet))
    };
    let CellValue::Formula { formula, .. } = value_at(book, set)? else {
        return Err(Error::GoalSeek(format!("{} must contain a formula", set.1)));
    };
    let start = match value_at(book, changing)? {
        CellValue::Formula { .. } => {
            return Err(Error::GoalSeek(format!(
                "{} must contain a value, not a formula",
                changing.1
            )));
        }
        other => other.as_number().unwrap_or(0.0),
    };
    let formula = crate::formula::parse(&formula)
        .map_err(|e| Error::InvalidFormula(format!("{}: {e}", set.1)))?;

    let mut trial = |x: f64| -> f64 {
        if let Some(ws) = book.sheet_mut(changing.0) {
            ws.set(changing.1, x);
        }
        let mut engine = match options.functions() {
            Some(custom) => Engine::with_functions(book, custom),
            None => Engine::new(book),
        };
        match engine
            .eval_tree(Origin::new(set.0, set.1), &formula)
            .scalar()
        {
            Value::Number(n) => n - target,
            _ => f64::NAN,
        }
    };
    let found = seek(start, &mut trial, limits.max_iterations, limits.tolerance);

    if let Some(ws) = book.sheet_mut(changing.0) {
        ws.set(changing.1, found.x);
    }
    recalculate_from_with(book, &[changing], options);
    Ok(Sought {
        value: found.x,
        result: found.fx + target,
        iterations: found.iterations,
        converged: found.converged,
    })
}

/// Where a search ended: the best `x`, `f(x)` there, the steps taken and
/// whether `|f(x)|` got within the tolerance.
#[derive(Debug, Clone, Copy, PartialEq)]
struct Seek {
    x: f64,
    fx: f64,
    iterations: u32,
    converged: bool,
}

/// Finds `x` with `|f(x)| <= tol`: secant steps, growing while nothing is
/// bracketed, then secant-or-bisection inside a sign change. `f` answers
/// `NaN` where the model errors. Returns the best point seen.
fn seek(x0: f64, f: &mut dyn FnMut(f64) -> f64, max_iter: u32, tol: f64) -> Seek {
    let tol = if tol.is_finite() && tol > 0.0 {
        tol
    } else {
        DEFAULT_TOLERANCE
    };
    let max_iter = max_iter.clamp(1, 32_767);
    let x0 = if x0.is_finite() { x0 } else { 0.0 };
    let f0 = f(x0);
    let mut best = (x0, f0);
    let consider = |best: &mut (f64, f64), x: f64, fx: f64| {
        if fx.is_finite() && (!best.1.is_finite() || fx.abs() < best.1.abs()) {
            *best = (x, fx);
        }
    };
    let done = |best: (f64, f64), iterations: u32| Seek {
        x: best.0,
        fx: best.1,
        iterations,
        converged: best.1.is_finite() && best.1.abs() <= tol,
    };
    if f0.is_finite() && f0.abs() <= tol {
        return done(best, 1);
    }
    let step0 = if x0 == 0.0 { 0.01 } else { x0.abs() * 0.01 };
    let (mut xa, mut fa) = (x0, f0);
    let mut xb = x0 + step0;
    let mut fb = f(xb);
    let mut it = 2u32;
    consider(&mut best, xb, fb);
    // (lo, f(lo), hi, f(hi)) across a sign change.
    let mut bracket: Option<(f64, f64, f64, f64)> = None;
    let mut slow = 0u32;
    while it < max_iter {
        if fb.is_finite() && fb.abs() <= tol {
            break;
        }
        if bracket.is_none() && fa.is_finite() && fb.is_finite() && (fa < 0.0) != (fb < 0.0) {
            bracket = Some((xa, fa, xb, fb));
        }
        #[allow(clippy::float_cmp, reason = "a flat secant is exactly equal ends")]
        let secant = if fa.is_finite() && fb.is_finite() && fb != fa {
            xb - fb * (xb - xa) / (fb - fa)
        } else {
            f64::NAN
        };
        let next = if let Some((lo, _, hi, _)) = bracket {
            let (a, b) = (lo.min(hi), lo.max(hi));
            if slow >= 2 || !(secant > a && secant < b) {
                f64::midpoint(lo, hi)
            } else {
                secant
            }
        } else {
            let span = (xb - xa).abs().max(f64::MIN_POSITIVE);
            if !fb.is_finite() {
                // The model errors here: step back towards the last good point.
                f64::midpoint(xa, xb)
            } else if !secant.is_finite() {
                2.0f64.mul_add(xb - xa, xb)
            } else if (secant - xb).abs() > 100.0f64.mul_add(span, 1.0) {
                (secant - xb)
                    .signum()
                    .mul_add(100.0f64.mul_add(span, 1.0), xb)
            } else {
                secant
            }
        };
        if !next.is_finite() {
            break;
        }
        let fnext = f(next);
        it += 1;
        consider(&mut best, next, fnext);
        if let Some((lo, flo, hi, fhi)) = bracket {
            let width = (hi - lo).abs();
            let narrowed = if fnext.is_finite() {
                if (fnext < 0.0) == (flo < 0.0) {
                    (next, fnext, hi, fhi)
                } else {
                    (lo, flo, next, fnext)
                }
            } else if flo.abs() < fhi.abs() {
                // An error inside the bracket: shrink towards the better end.
                (lo, flo, next, fhi)
            } else {
                (next, flo, hi, fhi)
            };
            bracket = Some(narrowed);
            let (l, _, h, _) = narrowed;
            let new_width = (h - l).abs();
            slow = if new_width > width * 0.5 { slow + 1 } else { 0 };
            if new_width <= f64::EPSILON * (1.0 + l.abs().max(h.abs())) {
                break;
            }
        }
        xa = xb;
        fa = fb;
        xb = next;
        fb = fnext;
    }
    done(best, it)
}

// ---------------------------------------------------------------- scenarios

/// The text a scenario keeps for a value: a number as its shortest digits.
fn scenario_text(value: &CellValue) -> String {
    match value.result() {
        CellValue::Number(n) => n.to_string(),
        CellValue::Bool(b) => if *b { "TRUE" } else { "FALSE" }.to_owned(),
        other => other.display(
            crate::style::format::GENERAL,
            crate::shared::date::Epoch::default(),
        ),
    }
}

fn sheet_of(book: &Spreadsheet, sheet: usize) -> Result<&crate::model::Worksheet> {
    book.sheet(sheet).ok_or(Error::SheetIndexOutOfRange(sheet))
}

/// Adds a scenario to a sheet with the values its input cells hold now, as
/// the Scenario Manager's Add does; a scenario of the same name is
/// replaced. Returns its index.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::GoalSeek`] for no input cell, more than Excel's 32, or one that
/// holds a formula.
pub fn add_scenario(
    book: &mut Spreadsheet,
    sheet: usize,
    name: &str,
    inputs: &[CellRef],
) -> Result<usize> {
    let ws = sheet_of(book, sheet)?;
    if inputs.is_empty() || inputs.len() > 32 {
        return Err(Error::GoalSeek("a scenario changes 1 to 32 cells".into()));
    }
    let mut cells = Vec::new();
    for &cell in inputs {
        let value = ws.get(cell).map(|c| c.value.clone()).unwrap_or_default();
        if matches!(value, CellValue::Formula { .. }) {
            return Err(Error::GoalSeek(format!(
                "{cell} holds a formula; a scenario changes values"
            )));
        }
        cells.push(ScenarioInput {
            cell,
            value: scenario_text(&value),
            ..ScenarioInput::default()
        });
    }
    let scenario = Scenario {
        name: name.to_owned(),
        inputs: cells,
        ..Scenario::default()
    };
    let list = &mut book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?
        .scenarios
        .list;
    if let Some(at) = list.iter().position(|s| s.name.eq_ignore_ascii_case(name)) {
        list[at] = scenario;
        Ok(at)
    } else {
        list.push(scenario);
        Ok(list.len() - 1)
    }
}

/// Puts a scenario's values into its input cells and recomputes what reads
/// them, as the Scenario Manager's Show does. Returns how many cells were
/// set.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::GoalSeek`] for a scenario the sheet does not have.
pub fn show_scenario(book: &mut Spreadsheet, sheet: usize, name: &str) -> Result<usize> {
    let ws = sheet_of(book, sheet)?;
    let (index, scenario) = ws
        .scenarios
        .list
        .iter()
        .enumerate()
        .find(|(_, s)| s.name.eq_ignore_ascii_case(name))
        .ok_or_else(|| Error::GoalSeek(format!("the sheet has no scenario {name:?}")))?;
    let inputs: Vec<(CellRef, String)> = scenario
        .inputs
        .iter()
        .filter(|i| !i.deleted)
        .map(|i| (i.cell, i.value.clone()))
        .collect();
    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    for (cell, value) in &inputs {
        ws.set(*cell, crate::reader::csv::value_of(value));
    }
    ws.scenarios.current = u32::try_from(index).ok();
    let changed: Vec<(usize, CellRef)> = inputs.iter().map(|(c, _)| (sheet, *c)).collect();
    recalculate_from_with(book, &changed, &Options::default());
    Ok(inputs.len())
}

/// Writes the Scenario Manager's Summary report on a new sheet, "Scenario
/// Summary" (numbered when taken): for each scenario, the values of its
/// changing cells and what each of `results` computes with them, beside the
/// current values. The sheet it summarises is left as it was. Returns the
/// new sheet's index.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::GoalSeek`] when the sheet has no scenario.
pub fn scenario_summary(
    book: &mut Spreadsheet,
    sheet: usize,
    results: &[CellRef],
) -> Result<usize> {
    let ws = sheet_of(book, sheet)?;
    if ws.scenarios.list.is_empty() {
        return Err(Error::GoalSeek(
            "the sheet has no scenario to summarise".into(),
        ));
    }
    let scenarios = ws.scenarios.list.clone();
    // Every changing cell any scenario has, in order of first appearance.
    let mut changing: Vec<CellRef> = Vec::new();
    for input in scenarios.iter().flat_map(|s| &s.inputs) {
        if !changing.contains(&input.cell) {
            changing.push(input.cell);
        }
    }
    let read = |book: &Spreadsheet, cell: CellRef| {
        book.sheet(sheet)
            .and_then(|ws| ws.get(cell))
            .map(|c| c.value.result().clone())
            .unwrap_or_default()
    };
    // One column per state: the current values, then each scenario shown on
    // a copy of the workbook.
    let mut columns: Vec<(String, Vec<CellValue>)> = Vec::new();
    let current: Vec<CellValue> = changing
        .iter()
        .chain(results)
        .map(|&c| read(book, c))
        .collect();
    columns.push(("Current Values:".into(), current));
    for scenario in &scenarios {
        let mut copy = book.clone();
        show_scenario(&mut copy, sheet, &scenario.name)?;
        let values = changing
            .iter()
            .chain(results)
            .map(|&c| read(&copy, c))
            .collect();
        columns.push((scenario.name.clone(), values));
    }

    let title = {
        let taken = |t: &str| {
            book.sheets()
                .iter()
                .any(|s| s.title().eq_ignore_ascii_case(t))
        };
        let mut title = "Scenario Summary".to_owned();
        let mut n = 2;
        while taken(&title) {
            title = format!("Scenario Summary {n}");
            n += 1;
        }
        title
    };
    let mut report = crate::model::Worksheet::new(title)?;
    // Rows and columns counted from 1 here, as the report reads.
    let mut put = |row: u32, col: u32, value: CellValue| {
        if let (Some(r), Some(c)) = (Row::new(row - 1), Col::new(col - 1)) {
            report.set(CellRef::new(c, r), value);
        }
    };
    put(1, 1, CellValue::text("Scenario Summary"));
    for (i, (name, _)) in (3u32..).zip(&columns) {
        put(2, i, CellValue::text(name.as_str()));
    }
    put(3, 1, CellValue::text("Changing Cells:"));
    let mut row = 4;
    let mut lines: Vec<(u32, CellRef)> = Vec::new();
    for &cell in &changing {
        lines.push((row, cell));
        row += 1;
    }
    if !results.is_empty() {
        put(row, 1, CellValue::text("Result Cells:"));
        row += 1;
        for &cell in results {
            lines.push((row, cell));
            row += 1;
        }
    }
    for (i, &(r, cell)) in lines.iter().enumerate() {
        put(r, 2, CellValue::text(cell.to_string().as_str()));
        for (c, (_, values)) in (3u32..).zip(&columns) {
            put(r, c, values[i].clone());
        }
    }
    put(
        row + 1,
        1,
        CellValue::text("Notes: Current Values column represents values of changing cells at"),
    );
    put(
        row + 2,
        1,
        CellValue::text("time Scenario Summary Report was created."),
    );
    book.add_sheet(report)
}

// --------------------------------------------------------------- data tables

/// Fills a data table, as What-If Analysis - Data Table does, with values.
///
/// `area` is the whole table, its top row and left column included.
/// - With only `column_input`: the inputs run down the left column and the
///   formulas across the top row; each body cell is its column's formula
///   with `column_input` set to its row's input.
/// - With only `row_input`: the other way round.
/// - With both: the formula is the corner cell, the row inputs across the
///   top and the column inputs down the left.
///
/// Excel writes `{=TABLE(...)}` there, which recalculates; this crate does
/// not model that formula, so the body gets the values it computes, and the
/// table must be filled again when the model changes. The input cells keep
/// their values. Returns how many cells were filled.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::GoalSeek`] when neither input is given, an input lies inside
/// the table, the table is smaller than 2 by 2, or a cell that should hold a
/// formula does not.
pub fn data_table(
    book: &mut Spreadsheet,
    sheet: usize,
    area: Range,
    row_input: Option<CellRef>,
    column_input: Option<CellRef>,
) -> Result<usize> {
    sheet_of(book, sheet)?;
    if row_input.is_none() && column_input.is_none() {
        return Err(Error::GoalSeek(
            "a data table needs a row or a column input cell".into(),
        ));
    }
    if area.width() < 2 || area.height() < 2 {
        return Err(Error::GoalSeek(format!(
            "{area} is too small for a data table"
        )));
    }
    if [row_input, column_input]
        .into_iter()
        .flatten()
        .any(|c| area.contains(c))
    {
        return Err(Error::GoalSeek(
            "an input cell may not lie inside the table".into(),
        ));
    }
    let at = |dr: u32, dc: u32| {
        CellRef::new(
            Col::new(area.start.col.index() + dc).unwrap_or(area.start.col),
            Row::new(area.start.row.index() + dr).unwrap_or(area.start.row),
        )
    };
    let value = |book: &Spreadsheet, cell: CellRef| {
        book.sheet(sheet)
            .and_then(|ws| ws.get(cell))
            .map(|c| c.value.result().clone())
            .unwrap_or_default()
    };
    let formula_at = |book: &Spreadsheet, cell: CellRef| -> Result<crate::formula::Expr> {
        match book
            .sheet(sheet)
            .and_then(|ws| ws.get(cell))
            .map(|c| &c.value)
        {
            Some(CellValue::Formula { formula, .. }) => crate::formula::parse(formula)
                .map_err(|e| Error::InvalidFormula(format!("{cell}: {e}"))),
            _ => Err(Error::GoalSeek(format!(
                "{cell} must hold the formula the table computes"
            ))),
        }
    };
    let (rows, cols) = (area.height(), area.width());
    // (body cell, formula cell, row input value, column input value)
    let mut jobs: Vec<(CellRef, CellRef, Option<CellValue>, Option<CellValue>)> = Vec::new();
    for r in 1..rows {
        for c in 1..cols {
            let formula_cell = match (row_input, column_input) {
                (Some(_), Some(_)) => at(0, 0),
                (None, Some(_)) => at(0, c),
                _ => at(r, 0),
            };
            let row_value = row_input.map(|_| value(book, at(0, c)));
            let col_value = column_input.map(|_| value(book, at(r, 0)));
            jobs.push((at(r, c), formula_cell, row_value, col_value));
        }
    }
    let mut work = book.clone();
    let mut out = Vec::with_capacity(jobs.len());
    for (body, formula_cell, row_value, col_value) in jobs {
        let expr = formula_at(book, formula_cell)?;
        if let Some(ws) = work.sheet_mut(sheet) {
            if let (Some(cell), Some(v)) = (row_input, row_value) {
                ws.set(cell, v);
            }
            if let (Some(cell), Some(v)) = (column_input, col_value) {
                ws.set(cell, v);
            }
        }
        let result = Engine::new(&work).eval_tree(Origin::new(sheet, formula_cell), &expr);
        out.push((body, crate::formula::eval::stored(result.scalar())));
    }
    let count = out.len();
    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    for (cell, value) in out {
        ws.set(cell, value);
    }
    Ok(count)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn finds_a_square_root() {
        let r = seek(1.0, &mut |x| x * x - 2.0, 100, 1e-9);
        assert!(r.converged);
        assert!((r.x - 2f64.sqrt()).abs() < 1e-6);
    }

    #[test]
    fn steps_out_of_an_error() {
        // LN errors below zero; the search must still land on e.
        let r = seek(
            5.0,
            &mut |x| if x > 0.0 { x.ln() - 1.0 } else { f64::NAN },
            100,
            1e-9,
        );
        assert!(r.converged);
        assert!((r.x - std::f64::consts::E).abs() < 1e-6);
    }

    #[test]
    fn reports_the_best_point_when_there_is_no_root() {
        let r = seek(3.0, &mut |x| x * x + 1.0, 50, 1e-9);
        assert!(!r.converged);
        assert!(r.fx >= 1.0);
    }
}
