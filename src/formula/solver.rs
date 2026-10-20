//! Solver: values for a sheet's variable cells that maximise, minimise or
//! hit a value in an objective cell, subject to constraints, as Excel's Data
//! - Solver does.
//!
//! The workbook is the model: every trial sets the variable cells on a copy
//! and recomputes what reads them, so any formula works. Two methods:
//!
//! - **Simplex LP**: the model is linearised from `n + 1` trials and checked
//!   at one more point; a dense two-phase simplex (Dantzig's rule, Bland's
//!   when it stalls) then solves it exactly. Integer and binary constraints
//!   are met by depth-first branch and bound.
//! - **GRG Nonlinear**, as Excel's dialog names it: an augmented-Lagrangian
//!   method whose inner problems are solved by projected gradient descent
//!   with Barzilai-Borwein steps and an Armijo line search, on central
//!   differences. It finds a local optimum of a smooth model.
//!
//! A model can be kept in the workbook as the hidden names Excel's Solver
//! add-in uses (`solver_opt`, `solver_adj`, `solver_lhs1`...), so a model set
//! up in Excel can be solved here and the other way round.
//!
//! The numerical core was ported from the Solver of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use crate::coordinate::{CellRef, Range};
use crate::error::{Error, Result};
use crate::formula::eval::Dependencies;
use crate::model::{CellValue, DefinedName, Spreadsheet};

/// Most variable cells, as in Excel.
const MAX_VARS: usize = 200;
/// Trials a solve may take unless told otherwise.
const DEFAULT_TRIALS: u64 = 20_000;
/// Work units of one LP relaxation.
const LP_WORK: u64 = 2_000_000_000;
const MAX_NODES_LP: usize = 10_000;
const MAX_NODES_NLP: usize = 200;
/// Excel's default constraint precision.
const PRECISION: f64 = 1e-6;
const EPS: f64 = 1e-9;

/// What the objective should do.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum SolverGoal {
    /// As large as it gets.
    Max,
    /// As small as it gets.
    Min,
    /// Equal to this.
    Value(f64),
}

/// How the model is solved.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SolverMethod {
    /// For smooth nonlinear models: a local optimum.
    #[default]
    GrgNonlinear,
    /// For linear models: the exact optimum. A model found not to be linear
    /// is reported as [`SolverStatus::NotLinear`].
    SimplexLp,
}

/// How a constraint binds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Relation {
    /// `<=`.
    LessOrEqual,
    /// `>=`.
    GreaterOrEqual,
    /// `=`.
    Equal,
    /// Whole numbers (`int`); only on variable cells.
    Integer,
    /// 0 or 1 (`bin`); only on variable cells.
    Binary,
}

/// The right-hand side of a constraint.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Bound {
    /// A number.
    Value(f64),
    /// Cells, matched to the left side's one by one (or one for all).
    Cells(Range),
    /// None: for [`Relation::Integer`] and [`Relation::Binary`].
    None,
}

/// One constraint: each cell of `cells` against `bound`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SolverConstraint {
    /// The left side.
    pub cells: Range,
    /// How it binds.
    pub relation: Relation,
    /// The right side.
    pub bound: Bound,
}

/// A Solver model on one sheet.
#[derive(Debug, Clone, PartialEq)]
pub struct SolverModel {
    /// The objective cell, which holds a formula.
    pub objective: CellRef,
    /// What it should do.
    pub goal: SolverGoal,
    /// The variable cells ("By Changing Variable Cells").
    pub variables: Vec<Range>,
    /// The constraints.
    pub constraints: Vec<SolverConstraint>,
    /// "Make Unconstrained Variables Non-Negative", on by default as in Excel.
    pub non_negative: bool,
    /// The method.
    pub method: SolverMethod,
    /// The most trials before giving up.
    pub max_trials: u64,
}

impl SolverModel {
    /// A model with Excel's defaults: no constraints, non-negative
    /// variables, GRG Nonlinear.
    #[must_use]
    pub fn new(objective: CellRef, goal: SolverGoal, variables: Vec<Range>) -> Self {
        Self {
            objective,
            goal,
            variables,
            constraints: Vec::new(),
            non_negative: true,
            method: SolverMethod::default(),
            max_trials: DEFAULT_TRIALS,
        }
    }
}

/// How a solve ended.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SolverStatus {
    /// "Solver found a solution. All constraints and optimality conditions
    /// are satisfied."
    Optimal,
    /// A solution meeting the constraints, found before the trials ran out.
    Feasible,
    /// No point meets every constraint.
    Infeasible,
    /// The objective grows without limit.
    Unbounded,
    /// Simplex LP was asked to solve a model that is not linear.
    NotLinear,
}

/// What [`solve`] found.
#[derive(Debug, Clone, PartialEq)]
pub struct Solved {
    /// How it ended.
    pub status: SolverStatus,
    /// The objective's value at the solution, or where the cells were left.
    pub objective: f64,
    /// The variable cells and the values they were given.
    pub values: Vec<(CellRef, f64)>,
    /// How many trials it took.
    pub trials: u64,
}

// ------------------------------------------------------------------- model

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Sense {
    Le,
    Ge,
    Eq,
}

/// One scalar constraint: `lhs (sense) rhs`.
#[derive(Debug, Clone)]
struct Row {
    lhs: CellRef,
    sense: Sense,
    rhs: Option<CellRef>,
    rhs_const: f64,
}

/// The objective and `lhs - rhs` per row at a point.
#[derive(Debug, Clone)]
struct Point {
    f: f64,
    g: Vec<f64>,
}

/// Runs trials on a copy of the workbook.
struct Evaluator {
    work: Spreadsheet,
    deps: Dependencies,
    sheet: usize,
    vars: Vec<CellRef>,
    objective: CellRef,
    rows: Vec<Row>,
    trials: u64,
    budget: u64,
}

impl Evaluator {
    const fn exhausted(&self) -> bool {
        self.trials >= self.budget
    }

    fn read(&self, at: CellRef) -> f64 {
        match self
            .work
            .sheet(self.sheet)
            .and_then(|ws| ws.get(at))
            .map(|c| c.value.result())
        {
            Some(CellValue::Number(n)) => *n,
            Some(CellValue::Bool(b)) => f64::from(u8::from(*b)),
            None | Some(CellValue::Empty) => 0.0,
            Some(_) => f64::NAN,
        }
    }

    /// The model at `x`; `None` once the trials are spent.
    fn eval(&mut self, x: &[f64]) -> Option<Point> {
        if self.exhausted() {
            return None;
        }
        self.trials += 1;
        let ws = self.work.sheet_mut(self.sheet)?;
        for (cell, v) in self.vars.iter().zip(x) {
            ws.set(*cell, *v);
        }
        let changed: Vec<(usize, CellRef)> = self.vars.iter().map(|c| (self.sheet, *c)).collect();
        self.deps.recalculate_from(&mut self.work, &changed);
        let f = self.read(self.objective);
        let g = self
            .rows
            .iter()
            .map(|r| self.read(r.lhs) - r.rhs.map_or(r.rhs_const, |c| self.read(c)))
            .collect();
        Some(Point { f, g })
    }
}

struct Problem {
    lo: Vec<f64>,
    hi: Vec<f64>,
    integer: Vec<bool>,
    rows: Vec<Row>,
}

/// Splits the constraints into variable bounds and general rows.
fn build_problem(model: &SolverModel, vars: &[CellRef]) -> Result<Problem> {
    let n = vars.len();
    let mut lo = vec![f64::NEG_INFINITY; n];
    let mut hi = vec![f64::INFINITY; n];
    let mut has_lower = vec![false; n];
    let mut integer = vec![false; n];
    let mut rows = Vec::new();
    for c in &model.constraints {
        let rhs_cells: Vec<CellRef> = match c.bound {
            Bound::Cells(r) => r.cells().collect(),
            _ => Vec::new(),
        };
        let lhs_cells: Vec<CellRef> = c.cells.cells().collect();
        if let Bound::Cells(_) = c.bound
            && rhs_cells.len() != 1
            && rhs_cells.len() != lhs_cells.len()
        {
            return Err(Error::GoalSeek(format!(
                "the two sides of the constraint on {} differ in size",
                c.cells
            )));
        }
        for (k, cell) in lhs_cells.iter().enumerate() {
            let j = vars.iter().position(|v| v == cell);
            match c.relation {
                Relation::Integer | Relation::Binary => {
                    let Some(j) = j else {
                        return Err(Error::GoalSeek(format!(
                            "{cell} is not a variable cell; int and bin bind variables only"
                        )));
                    };
                    integer[j] = true;
                    if c.relation == Relation::Binary {
                        lo[j] = lo[j].max(0.0);
                        hi[j] = hi[j].min(1.0);
                        has_lower[j] = true;
                    }
                }
                relation => {
                    let sense = match relation {
                        Relation::LessOrEqual => Sense::Le,
                        Relation::GreaterOrEqual => Sense::Ge,
                        _ => Sense::Eq,
                    };
                    let (rhs, rhs_const) = match c.bound {
                        Bound::Value(v) => (None, v),
                        Bound::Cells(_) => {
                            (rhs_cells.get(k).or_else(|| rhs_cells.first()).copied(), 0.0)
                        }
                        Bound::None => (None, 0.0),
                    };
                    match (j, rhs) {
                        // A constant bound on a variable is a box bound.
                        (Some(j), None) => {
                            if sense != Sense::Ge {
                                hi[j] = hi[j].min(rhs_const);
                            }
                            if sense != Sense::Le {
                                lo[j] = lo[j].max(rhs_const);
                                has_lower[j] = true;
                            }
                        }
                        _ => rows.push(Row {
                            lhs: *cell,
                            sense,
                            rhs,
                            rhs_const,
                        }),
                    }
                }
            }
        }
    }
    if model.non_negative {
        for (l, has) in lo.iter_mut().zip(&has_lower) {
            if !has {
                *l = l.max(0.0);
            }
        }
    }
    Ok(Problem {
        lo,
        hi,
        integer,
        rows,
    })
}

/// What a relaxation, or the whole solve, ended with.
#[derive(Debug, Clone)]
enum Outcome {
    /// x, the objective to minimise, and whether it converged.
    Solved(Vec<f64>, f64, bool),
    Infeasible,
    Unbounded,
    Limit,
}

// ---------------------------------------------------------- branch and bound

/// Depth-first branch and bound over a relaxation: the best integer point.
fn branch_and_bound(
    lo: &[f64],
    hi: &[f64],
    integer: &[bool],
    max_nodes: usize,
    relax: &mut dyn FnMut(&[f64], &[f64]) -> Outcome,
) -> Outcome {
    let mut stack: Vec<(Vec<f64>, Vec<f64>)> = vec![(lo.to_vec(), hi.to_vec())];
    let mut best: Option<(Vec<f64>, f64)> = None;
    let mut nodes = 0usize;
    let mut limited = false;
    let mut root = true;
    while let Some((l, h)) = stack.pop() {
        if nodes >= max_nodes {
            limited = true;
            break;
        }
        nodes += 1;
        let out = relax(&l, &h);
        let was_root = std::mem::replace(&mut root, false);
        let (x, obj, converged) = match out {
            Outcome::Solved(x, obj, c) => (x, obj, c),
            Outcome::Unbounded if was_root => return Outcome::Unbounded,
            Outcome::Limit => {
                limited = true;
                break;
            }
            _ => continue,
        };
        if !converged {
            limited = true;
        }
        if let Some((_, b)) = &best
            && obj >= b - 1e-9 * (1.0 + b.abs())
        {
            continue;
        }
        // The most fractional integer variable.
        let frac = x
            .iter()
            .zip(integer)
            .enumerate()
            .filter(|(_, (_, int))| **int)
            .map(|(j, (v, _))| (j, (v - v.round()).abs()))
            .filter(|(_, f)| *f > 1e-6)
            .max_by(|a, b| a.1.total_cmp(&b.1));
        if let Some((j, _)) = frac {
            let v = x[j];
            let (mut down_h, mut up_l) = (h.clone(), l.clone());
            down_h[j] = down_h[j].min(v.floor());
            up_l[j] = up_l[j].max(v.ceil());
            // The nearer side first: pushed last.
            if v - v.floor() < 0.5 {
                stack.push((up_l, h));
                stack.push((l, down_h));
            } else {
                stack.push((l, down_h));
                stack.push((up_l, h));
            }
        } else {
            let snapped = x
                .iter()
                .zip(integer)
                .map(|(v, int)| if *int { v.round() } else { *v })
                .collect();
            best = Some((snapped, obj));
        }
    }
    match best {
        Some((x, obj)) => Outcome::Solved(x, obj, !limited),
        None if limited => Outcome::Limit,
        None => Outcome::Infeasible,
    }
}

// ------------------------------------------------------------------ simplex

/// Pivots the tableau on `(r, j)`, the objective row too.
#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn pivot(t: &mut [Vec<f64>], obj: &mut [f64], r: usize, j: usize) {
    let pv = t[r][j];
    if pv == 0.0 {
        return;
    }
    for v in &mut t[r] {
        *v /= pv;
    }
    let prow = t[r].clone();
    for (i, row) in t.iter_mut().enumerate() {
        let f = row[j];
        if i != r && f != 0.0 {
            for (v, p) in row.iter_mut().zip(&prow) {
                *v -= f * p;
            }
        }
    }
    let f = obj[j];
    if f != 0.0 {
        for (v, p) in obj.iter_mut().zip(&prow) {
            *v -= f * p;
        }
    }
}

enum Step {
    Optimal,
    Unbounded,
    Limit,
}

/// Minimises with the reduced costs in `obj` (its last entry: minus the
/// objective). Columns not `allowed` never enter.
#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn simplex(
    t: &mut [Vec<f64>],
    basis: &mut [usize],
    obj: &mut [f64],
    allowed: &[bool],
    budget: &mut u64,
) -> Step {
    let n = obj.len().saturating_sub(1);
    let cost = (t.len() as u64 + 1).saturating_mul(n as u64 + 1).max(1);
    let mut stalled = 0u32;
    loop {
        if *budget < cost {
            return Step::Limit;
        }
        *budget -= cost;
        let bland = stalled > 50;
        let mut enter = None;
        let mut best = -EPS;
        for (j, d) in obj.iter().take(n).enumerate() {
            if !allowed[j] || *d >= -EPS {
                continue;
            }
            if bland {
                enter = Some(j);
                break;
            }
            if *d < best {
                best = *d;
                enter = Some(j);
            }
        }
        let Some(j) = enter else { return Step::Optimal };
        let mut leave: Option<(usize, f64)> = None;
        for (i, row) in t.iter().enumerate() {
            let a = row[j];
            if a <= EPS {
                continue;
            }
            let ratio = row[n] / a;
            let better = leave.is_none_or(|(li, lr)| {
                ratio < lr - 1e-12 || (ratio <= lr + 1e-12 && basis[i] < basis[li])
            });
            if better {
                leave = Some((i, ratio));
            }
        }
        let Some((r, ratio)) = leave else {
            return Step::Unbounded;
        };
        stalled = if ratio <= EPS { stalled + 1 } else { 0 };
        pivot(t, obj, r, j);
        basis[r] = j;
    }
}

/// Minimises `c . x` subject to `rows` and `lo <= x <= hi`.
#[allow(clippy::too_many_lines, reason = "the two phases of one tableau")]
#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn solve_lp(
    c: &[f64],
    rows: &[(Vec<f64>, Sense, f64)],
    lo: &[f64],
    hi: &[f64],
    budget: &mut u64,
) -> Outcome {
    struct Map {
        var: usize,
        sign: f64,
    }
    let n = c.len();
    if lo.iter().zip(hi).any(|(l, h)| l > h) {
        return Outcome::Infeasible;
    }
    // x_j = offset + sum of sign * y_k, with y >= 0.
    let mut cols: Vec<Map> = Vec::new();
    let mut offset = vec![0.0; n];
    let mut extra: Vec<(usize, f64)> = Vec::new();
    for j in 0..n {
        let (l, h) = (lo[j], hi[j]);
        if l.is_finite() {
            offset[j] = l;
            cols.push(Map { var: j, sign: 1.0 });
            if h.is_finite() {
                extra.push((cols.len() - 1, h - l));
            }
        } else if h.is_finite() {
            offset[j] = h;
            cols.push(Map { var: j, sign: -1.0 });
        } else {
            cols.push(Map { var: j, sign: 1.0 });
            cols.push(Map { var: j, sign: -1.0 });
        }
    }
    let ny = cols.len();
    let mut yrows: Vec<(Vec<f64>, Sense, f64)> = Vec::with_capacity(rows.len() + extra.len());
    for (a, sense, b) in rows {
        let coef: Vec<f64> = cols.iter().map(|m| a[m.var] * m.sign).collect();
        let shift: f64 = a.iter().zip(&offset).map(|(x, o)| x * o).sum();
        yrows.push((coef, *sense, b - shift));
    }
    for (k, b) in extra {
        let mut coef = vec![0.0; ny];
        coef[k] = 1.0;
        yrows.push((coef, Sense::Le, b));
    }
    let m = yrows.len();
    let n_slack = yrows.iter().filter(|r| r.1 != Sense::Eq).count();
    // Every right side made non-negative.
    let normalized: Vec<(Vec<f64>, Sense, f64)> = yrows
        .into_iter()
        .map(|(a, s, b)| {
            if b < 0.0 {
                let s = match s {
                    Sense::Le => Sense::Ge,
                    Sense::Ge => Sense::Le,
                    Sense::Eq => Sense::Eq,
                };
                (a.iter().map(|v| -v).collect(), s, -b)
            } else {
                (a, s, b)
            }
        })
        .collect();
    let n_art = normalized.iter().filter(|r| r.1 != Sense::Le).count();
    let width = ny + n_slack + n_art;
    let mut t: Vec<Vec<f64>> = Vec::with_capacity(m);
    let mut basis = vec![0usize; m];
    let (mut si, mut ai) = (ny, ny + n_slack);
    for (i, (a, s, b)) in normalized.iter().enumerate() {
        let mut row = vec![0.0; width + 1];
        row[..a.len()].copy_from_slice(a);
        match s {
            Sense::Le => {
                row[si] = 1.0;
                basis[i] = si;
                si += 1;
            }
            Sense::Ge => {
                row[si] = -1.0;
                si += 1;
                row[ai] = 1.0;
                basis[i] = ai;
                ai += 1;
            }
            Sense::Eq => {
                row[ai] = 1.0;
                basis[i] = ai;
                ai += 1;
            }
        }
        row[width] = *b;
        t.push(row);
    }
    let is_art = |j: usize| j >= ny + n_slack && j < width;
    // Phase 1: the sum of the artificials, to zero.
    let mut obj = vec![0.0; width + 1];
    for v in &mut obj[ny + n_slack..width] {
        *v = 1.0;
    }
    for (i, row) in t.iter().enumerate() {
        if is_art(basis[i]) {
            for (v, x) in obj.iter_mut().zip(row) {
                *v -= x;
            }
        }
    }
    if n_art > 0 {
        if let Step::Limit = simplex(&mut t, &mut basis, &mut obj, &vec![true; width], budget) {
            return Outcome::Limit;
        }
        let scale = normalized.iter().map(|r| r.2.abs()).fold(1.0f64, f64::max);
        if -obj[width] > 1e-7 * scale {
            return Outcome::Infeasible;
        }
        // Zero artificials out of the basis where possible.
        for i in 0..m {
            if !is_art(basis[i]) {
                continue;
            }
            if let Some(j) = (0..ny + n_slack).find(|&j| t[i][j].abs() > EPS) {
                pivot(&mut t, &mut obj, i, j);
                basis[i] = j;
            }
        }
    }
    // Phase 2: the real costs; artificials may not come back.
    let mut cost = vec![0.0; width + 1];
    for (k, mp) in cols.iter().enumerate() {
        cost[k] = c[mp.var] * mp.sign;
    }
    let mut obj = cost.clone();
    for (i, row) in t.iter().enumerate() {
        let cb = cost[basis[i]];
        if cb != 0.0 {
            for (v, x) in obj.iter_mut().zip(row) {
                *v -= cb * x;
            }
        }
    }
    let allowed: Vec<bool> = (0..width).map(|j| !is_art(j)).collect();
    match simplex(&mut t, &mut basis, &mut obj, &allowed, budget) {
        Step::Limit => return Outcome::Limit,
        Step::Unbounded => return Outcome::Unbounded,
        Step::Optimal => {}
    }
    let mut y = vec![0.0; width];
    for (i, row) in t.iter().enumerate() {
        y[basis[i]] = row[width];
    }
    let mut x = offset;
    for (k, mp) in cols.iter().enumerate() {
        x[mp.var] += mp.sign * y[k];
    }
    let value = c.iter().zip(&x).map(|(a, b)| a * b).sum();
    Outcome::Solved(x, value, true)
}

/// The objective and each row as linear functions of the variables, from
/// `n + 1` trials and a check at one more point.
type Linear = ((Vec<f64>, f64), Vec<(Vec<f64>, f64)>);

#[derive(Debug, PartialEq, Eq)]
enum Unlinear {
    Budget,
    Error,
    Nonlinear,
}

#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn linearize(ev: &mut Evaluator, x0: &[f64]) -> std::result::Result<Linear, Unlinear> {
    let n = x0.len();
    let p0 = ev.eval(x0).ok_or(Unlinear::Budget)?;
    if !p0.f.is_finite() || p0.g.iter().any(|v| !v.is_finite()) {
        return Err(Unlinear::Error);
    }
    let mut cf = vec![0.0; n];
    let mut ga = vec![vec![0.0; n]; p0.g.len()];
    for j in 0..n {
        let mut x = x0.to_vec();
        x[j] += 1.0;
        let p = ev.eval(&x).ok_or(Unlinear::Budget)?;
        cf[j] = p.f - p0.f;
        for (row, (gi, g0)) in ga.iter_mut().zip(p.g.iter().zip(&p0.g)) {
            row[j] = gi - g0;
        }
    }
    let dot = |a: &[f64], x: &[f64]| a.iter().zip(x).map(|(p, q)| p * q).sum::<f64>();
    let f_const = p0.f - dot(&cf, x0);
    let rows: Vec<(Vec<f64>, f64)> = ga
        .into_iter()
        .zip(&p0.g)
        .map(|(a, g0)| {
            let k = g0 - dot(&a, x0);
            (a, k)
        })
        .collect();
    // Checked at a point spread out from the first.
    #[allow(clippy::cast_precision_loss, reason = "small offsets by index")]
    let xt: Vec<f64> = x0
        .iter()
        .enumerate()
        .map(|(j, v)| {
            v + if j % 2 == 0 {
                0.3f64.mul_add((j % 5) as f64, 1.7)
            } else {
                -0.2f64.mul_add((j % 3) as f64, 0.6)
            }
        })
        .collect();
    let pt = ev.eval(&xt).ok_or(Unlinear::Budget)?;
    let close = |actual: f64, predicted: f64, terms: f64| {
        actual.is_finite() && (actual - predicted).abs() <= 1e-6 * (1.0 + actual.abs() + terms)
    };
    let terms = |a: &[f64]| a.iter().zip(&xt).map(|(p, q)| (p * q).abs()).sum::<f64>();
    if !close(pt.f, f_const + dot(&cf, &xt), terms(&cf)) {
        return Err(Unlinear::Nonlinear);
    }
    for ((a, k), g) in rows.iter().zip(&pt.g) {
        if !close(*g, k + dot(a, &xt), terms(a)) {
            return Err(Unlinear::Nonlinear);
        }
    }
    Ok(((cf, f_const), rows))
}

// ---------------------------------------------------------------- nonlinear

/// The augmented Lagrangian at a point.
fn aug(phi: f64, cons: &[f64], senses: &[Sense], lam: &[f64], mu: f64) -> f64 {
    let mut v = phi;
    for ((c, s), l) in cons.iter().zip(senses).zip(lam) {
        v += if *s == Sense::Eq {
            (0.5 * mu * c).mul_add(*c, l * c)
        } else {
            let t = mu.mul_add(*c, *l).max(0.0);
            t.mul_add(t, -l * l) / (2.0 * mu)
        };
    }
    v
}

/// The objective to minimise and the constraints as `<= 0` or `= 0`.
fn shape(p: &Point, rows: &[Row], goal: SolverGoal) -> (f64, Vec<f64>) {
    let phi = match goal {
        SolverGoal::Max => -p.f,
        SolverGoal::Min => p.f,
        SolverGoal::Value(v) => (p.f - v) * (p.f - v),
    };
    let cons =
        p.g.iter()
            .zip(rows)
            .map(|(g, r)| if r.sense == Sense::Ge { -g } else { *g })
            .collect();
    (phi, cons)
}

fn violation(cons: &[f64], senses: &[Sense]) -> f64 {
    cons.iter()
        .zip(senses)
        .map(|(c, s)| if *s == Sense::Eq { c.abs() } else { c.max(0.0) })
        .fold(0.0, f64::max)
}

fn project(x: &mut [f64], lo: &[f64], hi: &[f64]) {
    for ((v, l), h) in x.iter_mut().zip(lo).zip(hi) {
        *v = v.max(*l).min(*h);
    }
}

/// The gradient of the augmented Lagrangian by central differences, kept
/// inside the bounds.
#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn gradient(
    ev: &mut Evaluator,
    x: &[f64],
    (lo, hi): (&[f64], &[f64]),
    rows: &[Row],
    senses: &[Sense],
    goal: SolverGoal,
    (lam, mu): (&[f64], f64),
) -> Option<Vec<f64>> {
    let mut gr = vec![0.0; x.len()];
    for j in 0..x.len() {
        let xj = x[j];
        let h = 1e-6 * xj.abs().max(1.0);
        let a = if xj - h >= lo[j] { xj - h } else { xj };
        let b = if xj + h <= hi[j] { xj + h } else { xj };
        if b <= a {
            continue;
        }
        let mut xa = x.to_vec();
        let mut xb = x.to_vec();
        xa[j] = a;
        xb[j] = b;
        let (pa, ca) = shape(&ev.eval(&xa)?, rows, goal);
        let (pb, cb) = shape(&ev.eval(&xb)?, rows, goal);
        let d = (aug(pb, &cb, senses, lam, mu) - aug(pa, &ca, senses, lam, mu)) / (b - a);
        gr[j] = if d.is_finite() { d } else { 0.0 };
    }
    Some(gr)
}

/// A local minimum of the model within `lo..hi`, from `x0`.
#[allow(
    clippy::too_many_lines,
    reason = "the outer and inner loops of one method"
)]
#[allow(
    clippy::many_single_char_names,
    reason = "the textbook names of the method"
)]
fn nlp(ev: &mut Evaluator, x0: &[f64], lo: &[f64], hi: &[f64], goal: SolverGoal) -> Outcome {
    let rows = ev.rows.clone();
    let senses: Vec<Sense> = rows.iter().map(|r| r.sense).collect();
    let mut x = x0.to_vec();
    project(&mut x, lo, hi);
    let mut lam = vec![0.0; rows.len()];
    let mut mu = 10.0;
    let Some(p) = ev.eval(&x) else {
        return Outcome::Limit;
    };
    let (mut phi, mut cons) = shape(&p, &rows, goal);
    if !phi.is_finite() {
        return Outcome::Infeasible;
    }
    let phi0 = phi;
    let mut prev_viol = violation(&cons, &senses);
    let mut converged = false;
    let mut out_of_budget = false;
    let mut carried: Option<f64> = None;
    let scale = rows.iter().map(|r| r.rhs_const.abs()).fold(1.0, f64::max);
    for _outer in 0..40 {
        let mut l_cur = aug(phi, &cons, &senses, &lam, mu);
        let Some(mut g) = gradient(ev, &x, (lo, hi), &rows, &senses, goal, (&lam, mu)) else {
            out_of_budget = true;
            break;
        };
        let gnorm = g.iter().map(|v| v.abs()).fold(0.0, f64::max);
        let mut alpha = carried.unwrap_or(if gnorm > 0.0 { 1.0 / gnorm } else { 1.0 });
        let mut inner_done = false;
        let mut quiet = 0;
        for _inner in 0..500 {
            let mut accepted = None;
            let mut a = alpha;
            for _ in 0..40 {
                let mut xn: Vec<f64> = x.iter().zip(&g).map(|(v, d)| v - a * d).collect();
                project(&mut xn, lo, hi);
                let Some(pn) = ev.eval(&xn) else { break };
                let (phn, cn) = shape(&pn, &rows, goal);
                let ln = aug(phn, &cn, &senses, &lam, mu);
                let decrease: f64 = g
                    .iter()
                    .zip(x.iter().zip(&xn))
                    .map(|(d, (u, v))| d * (u - v))
                    .sum();
                if ln.is_finite() && ln <= 1e-4f64.mul_add(-decrease, l_cur) {
                    accepted = Some((xn, phn, cn, ln));
                    break;
                }
                a *= 0.25;
            }
            if ev.exhausted() {
                out_of_budget = true;
            }
            let Some((xn, phn, cn, ln)) = accepted else {
                inner_done = true;
                break;
            };
            // Running off towards infinity: unbounded.
            if phn.abs() > 1e15 * (1.0 + phi0.abs())
                || xn
                    .iter()
                    .zip(x0)
                    .any(|(v, s)| v.abs() > 1e15 * (1.0 + s.abs()))
            {
                return Outcome::Unbounded;
            }
            let s: Vec<f64> = xn.iter().zip(&x).map(|(a, b)| a - b).collect();
            let step = s.iter().map(|v| v.abs()).fold(0.0, f64::max);
            let xmax = xn.iter().map(|v| v.abs()).fold(1.0, f64::max);
            let small = step <= 1e-10 * xmax || (l_cur - ln).abs() <= 1e-13 * (1.0 + l_cur.abs());
            quiet = if small { quiet + 1 } else { 0 };
            x = xn;
            phi = phn;
            cons = cn;
            l_cur = ln;
            if quiet >= 3 {
                inner_done = true;
                break;
            }
            let Some(gn) = gradient(ev, &x, (lo, hi), &rows, &senses, goal, (&lam, mu)) else {
                out_of_budget = true;
                break;
            };
            // Barzilai-Borwein.
            let yv: Vec<f64> = gn.iter().zip(&g).map(|(a, b)| a - b).collect();
            let sy: f64 = s.iter().zip(&yv).map(|(a, b)| a * b).sum();
            let ss: f64 = s.iter().map(|v| v * v).sum();
            alpha = if sy > 0.0 {
                (ss / sy).clamp(1e-12, 1e12)
            } else {
                (a * 4.0).min(1e12)
            };
            g = gn;
            if g.iter().all(|v| *v == 0.0) {
                inner_done = true;
                break;
            }
        }
        carried = Some(alpha);
        if out_of_budget {
            break;
        }
        let viol = violation(&cons, &senses);
        if viol <= PRECISION * scale && inner_done {
            converged = true;
            if rows.is_empty() {
                break;
            }
        }
        for ((l, c), s) in lam.iter_mut().zip(&cons).zip(&senses) {
            *l = match s {
                Sense::Eq => mu.mul_add(*c, *l),
                _ => mu.mul_add(*c, *l).max(0.0),
            };
        }
        if viol > 0.25 * prev_viol {
            mu = (mu * 10.0).min(1e10);
        }
        if converged && viol <= PRECISION * scale * 1e-2 {
            break;
        }
        prev_viol = viol;
        converged = false;
    }
    let viol = violation(&cons, &senses);
    if viol > PRECISION * scale * 10.0 {
        return if out_of_budget {
            Outcome::Limit
        } else {
            Outcome::Infeasible
        };
    }
    Outcome::Solved(
        x,
        phi,
        (converged || viol <= PRECISION * scale) && !out_of_budget,
    )
}

// -------------------------------------------------------------------- solve

/// Runs the model's method from `x0`: the outcome, and whether Simplex LP
/// found the model not to be linear.
fn optimise(
    ev: &mut Evaluator,
    model: &SolverModel,
    problem: &Problem,
    x0: &[f64],
) -> (Outcome, bool) {
    let any_int = problem.integer.iter().any(|b| *b);
    let (lo, hi) = (&problem.lo, &problem.hi);
    let mut not_linear = false;
    let outcome = if lo.iter().zip(hi).any(|(l, h)| l > h) {
        Outcome::Infeasible
    } else {
        match model.method {
            SolverMethod::SimplexLp => match linearize(ev, x0) {
                Err(Unlinear::Budget) => Outcome::Limit,
                Err(_) => {
                    not_linear = true;
                    Outcome::Infeasible
                }
                Ok(((cf, f_const), lin)) => {
                    let mut rows: Vec<(Vec<f64>, Sense, f64)> = lin
                        .iter()
                        .zip(&ev.rows)
                        .map(|((a, k), r)| (a.clone(), r.sense, -k))
                        .collect();
                    let c: Vec<f64> = match model.goal {
                        SolverGoal::Max => cf.iter().map(|v| -v).collect(),
                        SolverGoal::Min => cf,
                        SolverGoal::Value(v) => {
                            let width = cf.len();
                            rows.push((cf, Sense::Eq, v - f_const));
                            vec![0.0; width]
                        }
                    };
                    let mut relax = |l: &[f64], h: &[f64]| {
                        let mut work = LP_WORK;
                        solve_lp(&c, &rows, l, h, &mut work)
                    };
                    if any_int {
                        branch_and_bound(lo, hi, &problem.integer, MAX_NODES_LP, &mut relax)
                    } else {
                        relax(lo, hi)
                    }
                }
            },
            SolverMethod::GrgNonlinear => {
                if any_int {
                    let mut relax = |l: &[f64], h: &[f64]| {
                        let mut from = x0.to_vec();
                        project(&mut from, l, h);
                        nlp(ev, &from, l, h, model.goal)
                    };
                    branch_and_bound(lo, hi, &problem.integer, MAX_NODES_NLP, &mut relax)
                } else {
                    nlp(ev, x0, lo, hi, model.goal)
                }
            }
        }
    };
    (outcome, not_linear)
}

/// Leaves `values` in the variable cells, recomputes what reads them, and
/// reads the objective and the variables back.
fn write_back(
    book: &mut Spreadsheet,
    sheet: usize,
    vars: &[CellRef],
    values: &[f64],
    objective: CellRef,
) -> Result<(f64, Vec<(CellRef, f64)>)> {
    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    for (cell, v) in vars.iter().zip(values) {
        // Integers found by branch and bound are whole; clean float dust off.
        let v = if (v - v.round()).abs() < 1e-9 {
            v.round()
        } else {
            *v
        };
        ws.set(*cell, v);
    }
    let changed: Vec<(usize, CellRef)> = vars.iter().map(|c| (sheet, *c)).collect();
    crate::formula::eval::recalculate_from(book, &changed);
    let read = |at: CellRef| {
        book.sheet(sheet)
            .and_then(|ws| ws.get(at))
            .and_then(|cell| cell.value.as_number())
    };
    let value = read(objective).unwrap_or(f64::NAN);
    let cells = vars.iter().map(|&c| (c, read(c).unwrap_or(0.0))).collect();
    Ok((value, cells))
}

/// Solves `model` on `sheet`: Excel's Data - Solver, choosing "Keep Solver
/// Solution". When a solution is found (`Optimal` or `Feasible`) the
/// variable cells keep it and the workbook is recomputed; otherwise they are
/// put back as they were.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::GoalSeek`] when the objective holds no formula, there are no
/// variable cells or more than 200, a variable holds a formula, or an
/// integer constraint is on a cell that is not a variable.
pub fn solve(book: &mut Spreadsheet, sheet: usize, model: &SolverModel) -> Result<Solved> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    if !matches!(
        ws.get(model.objective).map(|c| &c.value),
        Some(CellValue::Formula { .. })
    ) {
        return Err(Error::GoalSeek(format!(
            "the objective {} must hold a formula",
            model.objective
        )));
    }
    let mut vars: Vec<CellRef> = Vec::new();
    for cell in model.variables.iter().flat_map(Range::cells) {
        if !vars.contains(&cell) {
            vars.push(cell);
        }
    }
    if vars.is_empty() || vars.len() > MAX_VARS {
        return Err(Error::GoalSeek(format!(
            "Solver changes 1 to {MAX_VARS} cells"
        )));
    }
    let start: Vec<f64> = vars
        .iter()
        .map(|&c| match ws.get(c).map(|cell| &cell.value) {
            Some(CellValue::Formula { .. }) => {
                Err(Error::GoalSeek(format!("the variable {c} holds a formula")))
            }
            Some(v) => Ok(v.as_number().unwrap_or(0.0)),
            None => Ok(0.0),
        })
        .collect::<Result<_>>()?;
    let problem = build_problem(model, &vars)?;
    let mut work = book.clone();
    crate::formula::eval::recalculate(&mut work, None, &crate::progress::Options::default());
    let deps = Dependencies::of(&work);
    let mut ev = Evaluator {
        work,
        deps,
        sheet,
        vars: vars.clone(),
        objective: model.objective,
        rows: problem.rows.clone(),
        trials: 0,
        budget: model.max_trials.max(1),
    };
    let mut x0 = start.clone();
    project(&mut x0, &problem.lo, &problem.hi);
    let (outcome, not_linear) = optimise(&mut ev, model, &problem, &x0);
    let trials = ev.trials;
    let (status, keep) = match outcome {
        Outcome::Solved(x, _, converged) => (
            if converged {
                SolverStatus::Optimal
            } else {
                SolverStatus::Feasible
            },
            Some(x),
        ),
        Outcome::Unbounded => (SolverStatus::Unbounded, None),
        Outcome::Infeasible if not_linear => (SolverStatus::NotLinear, None),
        Outcome::Infeasible | Outcome::Limit => (SolverStatus::Infeasible, None),
    };
    let values = keep.unwrap_or(start);
    let (objective, values) = write_back(book, sheet, &vars, &values, model.objective)?;
    Ok(Solved {
        status,
        objective,
        values,
        trials,
    })
}

// ------------------------------------------------- the model in the workbook

fn absolute(range: Range) -> String {
    let cell = |c: CellRef| format!("${}${}", c.col.to_letters(), c.row.one_based());
    if range.start == range.end {
        cell(range.start)
    } else {
        format!("{}:{}", cell(range.start), cell(range.end))
    }
}

fn sheet_ref(book: &Spreadsheet, sheet: usize, range: Range) -> String {
    let title = book.sheet(sheet).map_or("Sheet1", |ws| ws.title());
    let quoted = if title.chars().all(|c| c.is_alphanumeric() || c == '_') {
        title.to_owned()
    } else {
        format!("'{}'", title.replace('\'', "''"))
    };
    format!("{quoted}!{}", absolute(range))
}

fn parse_ref(formula: &str) -> Option<Range> {
    let text = formula.trim().trim_start_matches('=');
    let text = text.rsplit_once('!').map_or(text, |(_, r)| r);
    Range::parse(&text.replace('$', "")).ok()
}

/// Keeps `model` in the workbook as the hidden names Excel's Solver add-in
/// writes for `sheet` (`solver_opt`, `solver_typ`, `solver_adj`,
/// `solver_lhs1`, `solver_rel1`, `solver_rhs1`...), replacing any it had, so
/// Excel opens the file with the model set up.
pub fn save_model(book: &mut Spreadsheet, sheet: usize, model: &SolverModel) {
    book.defined_names
        .retain(|d| !(d.sheet == Some(sheet) && d.name.to_lowercase().starts_with("solver_")));
    let mut names: Vec<(String, String)> = Vec::new();
    let objective = Range::new(model.objective, model.objective);
    names.push(("solver_opt".into(), sheet_ref(book, sheet, objective)));
    let (typ, val) = match model.goal {
        SolverGoal::Max => (1, 0.0),
        SolverGoal::Min => (2, 0.0),
        SolverGoal::Value(v) => (3, v),
    };
    names.push(("solver_typ".into(), typ.to_string()));
    names.push(("solver_val".into(), val.to_string()));
    let adj: Vec<String> = model
        .variables
        .iter()
        .map(|r| sheet_ref(book, sheet, *r))
        .collect();
    names.push(("solver_adj".into(), adj.join(",")));
    names.push((
        "solver_neg".into(),
        if model.non_negative { "1" } else { "2" }.into(),
    ));
    names.push((
        "solver_eng".into(),
        match model.method {
            SolverMethod::GrgNonlinear => "1",
            SolverMethod::SimplexLp => "2",
        }
        .into(),
    ));
    names.push(("solver_num".into(), model.constraints.len().to_string()));
    for (i, c) in (1..).zip(&model.constraints) {
        names.push((format!("solver_lhs{i}"), sheet_ref(book, sheet, c.cells)));
        let rel = match c.relation {
            Relation::LessOrEqual => 1,
            Relation::Equal => 2,
            Relation::GreaterOrEqual => 3,
            Relation::Integer => 4,
            Relation::Binary => 5,
        };
        names.push((format!("solver_rel{i}"), rel.to_string()));
        let rhs = match c.bound {
            Bound::Value(v) => v.to_string(),
            Bound::Cells(r) => sheet_ref(book, sheet, r),
            Bound::None => match c.relation {
                Relation::Integer => "\"integer\"".into(),
                _ => "\"binary\"".into(),
            },
        };
        names.push((format!("solver_rhs{i}"), rhs));
    }
    names.push(("solver_ver".into(), "3".into()));
    for (name, formula) in names {
        book.defined_names.push(DefinedName {
            name,
            sheet: Some(sheet),
            formula,
            hidden: true,
        });
    }
}

/// A code of the Solver add-in's names, as written: a small whole number.
fn code_of(v: f64) -> i64 {
    #[allow(
        clippy::cast_possible_truncation,
        reason = "clamped to a small range first"
    )]
    let code = v.round().clamp(-1e6, 1e6) as i64;
    code
}

/// The Solver model kept in the workbook for `sheet`, as Excel's add-in or
/// [`save_model`] wrote it; `None` when there is none. The Evolutionary
/// method, which this crate does not run, reads as GRG Nonlinear.
#[must_use]
pub fn saved_model(book: &Spreadsheet, sheet: usize) -> Option<SolverModel> {
    let name = |n: &str| {
        book.defined_names
            .iter()
            .find(|d| d.sheet == Some(sheet) && d.name.eq_ignore_ascii_case(n))
            .map(|d| d.formula.trim().trim_start_matches('=').to_owned())
    };
    let number = |n: &str| name(n).and_then(|v| v.trim_matches('"').parse::<f64>().ok());
    // The add-in's small codes: 1 for max, 2 for <=, and so on.
    let code = |n: &str| number(n).map(code_of);
    let objective = parse_ref(&name("solver_opt")?)?.start;
    let goal = match code("solver_typ").unwrap_or(1) {
        2 => SolverGoal::Min,
        3 => SolverGoal::Value(number("solver_val").unwrap_or(0.0)),
        _ => SolverGoal::Max,
    };
    let variables: Vec<Range> = name("solver_adj")?
        .split(',')
        .filter_map(parse_ref)
        .collect();
    let mut model = SolverModel::new(objective, goal, variables);
    model.non_negative = code("solver_neg").is_none_or(|v| v == 1);
    model.method = match code("solver_eng").unwrap_or(1) {
        2 => SolverMethod::SimplexLp,
        _ => SolverMethod::GrgNonlinear,
    };
    #[allow(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "a small count"
    )]
    let count = number("solver_num").unwrap_or(0.0).max(0.0) as usize;
    for i in 1..=count {
        let Some(cells) = name(&format!("solver_lhs{i}"))
            .as_deref()
            .and_then(parse_ref)
        else {
            continue;
        };
        let relation = match code(&format!("solver_rel{i}")).unwrap_or(1) {
            2 => Relation::Equal,
            3 => Relation::GreaterOrEqual,
            4 => Relation::Integer,
            5 => Relation::Binary,
            _ => Relation::LessOrEqual,
        };
        let rhs = name(&format!("solver_rhs{i}")).unwrap_or_default();
        let bound = match relation {
            Relation::Integer | Relation::Binary => Bound::None,
            _ => rhs
                .trim()
                .parse::<f64>()
                .map(Bound::Value)
                .ok()
                .or_else(|| parse_ref(&rhs).map(Bound::Cells))
                .unwrap_or(Bound::Value(0.0)),
        };
        model.constraints.push(SolverConstraint {
            cells,
            relation,
            bound,
        });
    }
    Some(model)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_small_lp() {
        // Maximise 3x + 5y with x <= 4, 2y <= 12, 3x + 2y <= 18: x = 2, y = 6.
        let rows = vec![
            (vec![1.0, 0.0], Sense::Le, 4.0),
            (vec![0.0, 2.0], Sense::Le, 12.0),
            (vec![3.0, 2.0], Sense::Le, 18.0),
        ];
        let mut work = LP_WORK;
        let Outcome::Solved(x, value, true) = solve_lp(
            &[-3.0, -5.0],
            &rows,
            &[0.0, 0.0],
            &[f64::INFINITY; 2],
            &mut work,
        ) else {
            panic!("not solved");
        };
        assert!(
            (x[0] - 2.0).abs() < 1e-9 && (x[1] - 6.0).abs() < 1e-9,
            "{x:?}"
        );
        assert!((value + 36.0).abs() < 1e-9);
    }

    #[test]
    fn infeasible_and_unbounded() {
        let mut work = LP_WORK;
        let rows = vec![(vec![1.0], Sense::Ge, 5.0), (vec![1.0], Sense::Le, 3.0)];
        assert!(matches!(
            solve_lp(&[1.0], &rows, &[0.0], &[f64::INFINITY], &mut work),
            Outcome::Infeasible
        ));
        let mut work = LP_WORK;
        assert!(matches!(
            solve_lp(&[-1.0], &[], &[0.0], &[f64::INFINITY], &mut work),
            Outcome::Unbounded
        ));
    }
}
