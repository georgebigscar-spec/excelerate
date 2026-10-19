//! What-If Analysis: Goal Seek.
//!
//! The root finder was ported from the `data.goalSeek` command of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use crate::coordinate::CellRef;
use crate::error::{Error, Result};
use crate::formula::eval::{Engine, recalculate_from_with};
use crate::formula::{Origin, Value};
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
