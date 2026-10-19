//! The formula engine: evaluating, recalculating, and functions of the
//! caller's own.

use super::convert::at_index;
use super::{Book, js, reporter};
use crate::coordinate::CellRef;
use crate::error::CellError;
use crate::formula::eval::{Dependencies, Engine, Origin};
use crate::formula::value::Value;
use wasm_bindgen::JsCast;
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
impl Book {
    /// Registers a function of your own, callable from any formula in the
    /// workbook by that name.
    ///
    /// The function is handed the arguments already computed - a range arrives
    /// as an array of arrays - and whatever it returns becomes the value. A
    /// name a built-in claims stays the built-in's: a workbook where `SUM`
    /// means something else is a workbook nobody else can read.
    ///
    /// ```js
    /// book.registerFunction("MYRATE", (base) => base * 1.5 + 0.5);
    /// book.evaluate(0, "A1", "=MYRATE(10)"); // 15.5
    /// ```
    #[wasm_bindgen(js_name = registerFunction)]
    pub fn register_function(&mut self, name: &str, function: &js_sys::Function) {
        let function = function.clone();
        self.custom.register(name, move |args| {
            let js_args: js_sys::Array = args.iter().map(value_to_js_deep).collect();
            // A function that throws is an error value rather than a panic
            // crossing the boundary: a formula may not take the process down.
            match js_sys::Reflect::apply(&function, &JsValue::NULL, &js_args) {
                Ok(value) => js_to_value(&value),
                Err(_) => Value::Error(CellError::Value),
            }
        });
        // The index remembers which names were unknown, so it is rebuilt.
        self.deps = None;
    }

    /// Forgets a function registered earlier; answers whether there was one.
    #[wasm_bindgen(js_name = unregisterFunction)]
    pub fn unregister_function(&mut self, name: &str) -> bool {
        self.deps = None;
        self.custom.remove(name)
    }

    /// The names registered, in no particular order.
    #[wasm_bindgen(js_name = registeredFunctions)]
    #[must_use]
    pub fn registered_functions(&self) -> Vec<String> {
        self.custom.names().map(str::to_owned).collect()
    }

    /// Finds the value of `changing` that makes the formula in `set` compute
    /// `target`, as Excel's Goal Seek does, both cells on `sheet`. The best
    /// value found is left in `changing` and what reads it is recomputed.
    /// `options.maxIterations` and `options.tolerance` default to the
    /// workbook's own iteration settings (100 and 0.001 unless it says
    /// otherwise). Returns `{ value, result, iterations, converged }`.
    ///
    /// ```js
    /// const { value, converged } = book.goalSeek(0, "B4", 300, "B1");
    /// ```
    #[wasm_bindgen(js_name = goalSeek, unchecked_return_type = "GoalSeekResult")]
    pub fn goal_seek(
        &mut self,
        sheet: usize,
        set: &str,
        target: f64,
        changing: &str,
        #[wasm_bindgen(unchecked_optional_param_type = "GoalSeekOptions")] options: &JsValue,
    ) -> Result<JsValue, JsError> {
        use super::convert::{field, object};
        use crate::formula::what_if::{SeekLimits, goal_seek_with};
        let set = CellRef::parse(set).map_err(js)?;
        let changing = CellRef::parse(changing).map_err(js)?;
        let mut limits = SeekLimits::of(&self.workbook);
        if let Some(n) = field(options, "maxIterations").and_then(|v| v.as_f64()) {
            if !(n.fract() == 0.0 && n >= 1.0 && n <= f64::from(u32::MAX)) {
                return Err(JsError::new("maxIterations is a whole number from 1"));
            }
            #[expect(
                clippy::cast_possible_truncation,
                clippy::cast_sign_loss,
                reason = "checked whole and in range just above"
            )]
            let n = n as u32;
            limits.max_iterations = n;
        }
        if let Some(t) = field(options, "tolerance").and_then(|v| v.as_f64()) {
            if !(t.is_finite() && t > 0.0) {
                return Err(JsError::new("tolerance is a number above 0"));
            }
            limits.tolerance = t;
        }
        let functions = crate::progress::Options::new().with_functions(&self.custom);
        let found = goal_seek_with(
            &mut self.workbook,
            (sheet, set),
            target,
            (sheet, changing),
            limits,
            &functions,
        )
        .map_err(js)?;
        let number = |n: f64| {
            if n.is_finite() {
                JsValue::from_f64(n)
            } else {
                JsValue::NULL
            }
        };
        Ok(object(&[
            ("value", JsValue::from_f64(found.value)),
            ("result", number(found.result)),
            ("iterations", JsValue::from_f64(f64::from(found.iterations))),
            ("converged", JsValue::from_bool(found.converged)),
        ]))
    }

    /// Whether `value` would be accepted in a cell under its data validation
    /// rule: `null` when it would (or there is no rule), otherwise what Excel
    /// would report, `{ style, title, message, sqref }`.
    ///
    /// ```js
    /// const problem = book.checkValidation(0, "B5", 150);
    /// if (problem) alert(problem.message || "Not allowed");
    /// ```
    #[wasm_bindgen(js_name = checkValidation, unchecked_return_type = "ValidationProblem | null")]
    pub fn check_validation(
        &self,
        sheet: usize,
        address: &str,
        #[wasm_bindgen(unchecked_param_type = "CellValue")] value: &JsValue,
    ) -> Result<JsValue, JsError> {
        use super::convert::{js_to_cell, object, ranges_to_js};
        let at = CellRef::parse(address).map_err(js)?;
        self.sheet_of(sheet)?;
        let value = js_to_cell(value)?;
        Ok(
            match crate::formula::validation::check(&self.workbook, sheet, at, &value) {
                None => JsValue::NULL,
                Some(rule) => object(&[
                    ("style", JsValue::from_str(rule.error_style.as_str())),
                    ("title", JsValue::from_str(&rule.error_title)),
                    ("message", JsValue::from_str(&rule.error)),
                    ("sqref", ranges_to_js(&rule.sqref)),
                ]),
            },
        )
    }

    /// The cells of a sheet whose values break their data validation rule,
    /// as Excel's Circle Invalid Data finds them: `["B3", "B9"]`.
    #[wasm_bindgen(js_name = invalidCells)]
    pub fn invalid_cells(&self, sheet: usize) -> Result<Vec<String>, JsError> {
        self.sheet_of(sheet)?;
        Ok(
            crate::formula::validation::invalid_cells(&self.workbook, sheet)
                .into_iter()
                .map(|at| at.to_string())
                .collect(),
        )
    }

    /// The choices of the drop-down list on a cell, as Excel shows them; an
    /// empty array when the cell has no list rule.
    #[wasm_bindgen(js_name = validationList)]
    pub fn validation_list(&self, sheet: usize, address: &str) -> Result<Vec<String>, JsError> {
        use crate::formula::validation::{list_items, rule_at};
        let at = CellRef::parse(address).map_err(js)?;
        self.sheet_of(sheet)?;
        Ok(rule_at(&self.workbook, sheet, at)
            .map(|rule| list_items(&self.workbook, sheet, rule))
            .unwrap_or_default())
    }

    /// Evaluates a formula (with or without the leading `=`) as if it sat in
    /// `address` of `sheet`, and returns its result.
    #[wasm_bindgen(unchecked_return_type = "CellValue")]
    pub fn evaluate(&self, sheet: usize, address: &str, formula: &str) -> Result<JsValue, JsError> {
        let at = CellRef::parse(address).map_err(js)?;
        self.sheet_of(sheet)?;
        let mut engine = Engine::with_functions(&self.workbook, &self.custom);
        let got = engine.eval(
            Origin::new(sheet, at),
            formula.strip_prefix('=').unwrap_or(formula),
        );
        Ok(value_to_js(&got))
    }

    /// Recomputes every formula and stores the results beside them, so that
    /// `get` and `toCsv` show this engine's answers rather than the cache the
    /// file was saved with. Returns how many formulas were computed.
    ///
    /// With no argument the whole workbook is recomputed; with a sheet index,
    /// only that sheet's formulas are - though they still read the whole
    /// workbook, as a cross-sheet reference must.
    #[expect(
        clippy::needless_pass_by_value,
        reason = "a callback crosses the wasm boundary owned"
    )]
    pub fn recalculate(
        &mut self,
        sheet: Option<usize>,
        on_progress: Option<js_sys::Function>,
    ) -> usize {
        let report = on_progress.as_ref().map(reporter);
        let mut options = crate::progress::Options::new().with_functions(&self.custom);
        if let Some(report) = &report {
            options = options.reporting(report);
        }
        crate::formula::eval::recalculate(&mut self.workbook, sheet, &options)
    }

    /// Recomputes the formula in one cell and stores its result. Returns
    /// whether there was a formula there.
    ///
    /// Not the same as `recalculateFrom`, which recomputes the formulas
    /// *reading* this cell and leaves the cell itself alone.
    #[wasm_bindgen(js_name = recalculateCell)]
    pub fn recalculate_cell(&mut self, sheet: usize, address: &str) -> Result<bool, JsError> {
        Ok(self.recalculate_one(sheet, CellRef::parse(address).map_err(js)?))
    }

    /// The same by 1-based row and column.
    #[wasm_bindgen(js_name = recalculateCellAt)]
    pub fn recalculate_cell_at(
        &mut self,
        sheet: usize,
        row: u32,
        column: u32,
    ) -> Result<bool, JsError> {
        Ok(self.recalculate_one(sheet, at_index(row, column)?))
    }

    /// Recomputes only what depends on one cell that changed - the formulas
    /// reading it, the formulas reading those, and so on. Returns how many were
    /// computed. This is the pass to run after `set`.
    ///
    /// For a batch of edits use `recalculateFromMany`, which answers them all
    /// in one pass.
    #[wasm_bindgen(js_name = recalculateFrom)]
    pub fn recalculate_from(&mut self, sheet: usize, address: &str) -> Result<usize, JsError> {
        self.recalculate_changed(sheet, &[CellRef::parse(address).map_err(js)?])
    }

    /// The same by 1-based row and column.
    #[wasm_bindgen(js_name = recalculateFromAt)]
    pub fn recalculate_from_at(
        &mut self,
        sheet: usize,
        row: u32,
        column: u32,
    ) -> Result<usize, JsError> {
        self.recalculate_changed(sheet, &[at_index(row, column)?])
    }

    /// The same for a batch of cells, all on sheet `sheet`, answered in one
    /// pass: editing a hundred cells costs one walk of the workbook, not a
    /// hundred.
    ///
    /// The index of what reads what is built on the first call and kept, so a
    /// run of edits parses the formulas only once. `set` keeps it in step.
    #[wasm_bindgen(js_name = recalculateFromMany)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an array of strings crosses the wasm boundary owned"
    )]
    pub fn recalculate_from_many(
        &mut self,
        sheet: usize,
        addresses: Vec<String>,
    ) -> Result<usize, JsError> {
        let cells = addresses
            .iter()
            .map(|address| CellRef::parse(address).map_err(js))
            .collect::<Result<Vec<_>, _>>()?;
        self.recalculate_changed(sheet, &cells)
    }
}

impl Book {
    fn recalculate_one(&mut self, sheet: usize, at: CellRef) -> bool {
        let options = crate::progress::Options::new().with_functions(&self.custom);
        crate::formula::eval::recalculate_cell_with(&mut self.workbook, sheet, at, &options)
    }

    fn recalculate_changed(&mut self, sheet: usize, cells: &[CellRef]) -> Result<usize, JsError> {
        self.sheet_of(sheet)?;
        let changed: Vec<_> = cells.iter().map(|&at| (sheet, at)).collect();
        let deps = self
            .deps
            .get_or_insert_with(|| Dependencies::of(&self.workbook));
        Ok(deps.recalculate_from(&mut self.workbook, &changed))
    }
}

/// What a registered function answered, as the engine sees it.
///
/// The mirror of [`value_to_js`]: numbers, text and booleans come back as
/// themselves, an array of arrays as an array value, and anything else - a
/// promise, an object, `undefined` - as blank, because a formula has to end
/// with a value and there is nothing else to make of it. An error is spelled
/// the way a cell spells it, so returning `"#N/A"` gives `#N/A` rather than
/// the text.
fn js_to_value(value: &JsValue) -> Value {
    if let Some(n) = value.as_f64() {
        return Value::Number(n);
    }
    if let Some(b) = value.as_bool() {
        return Value::Bool(b);
    }
    if let Some(text) = value.as_string() {
        return match CellError::parse(&text) {
            Some(e) => Value::Error(e),
            None => Value::Text(text),
        };
    }
    if let Ok(rows) = value.clone().dyn_into::<js_sys::Array>() {
        let grid: Vec<Vec<Value>> = rows
            .iter()
            .map(|row| match row.dyn_into::<js_sys::Array>() {
                Ok(cells) => cells.iter().map(|c| js_to_value(&c)).collect(),
                // A flat array is one row, which is how JS callers write one.
                Err(one) => vec![js_to_value(&one)],
            })
            .collect();
        return Value::array(grid);
    }
    Value::Blank
}

/// An argument as a registered function sees it: a range arrives whole, as an
/// array of rows, rather than collapsed the way a cell would show it.
fn value_to_js_deep(value: &Value) -> JsValue {
    match value {
        Value::Array(rows) => rows
            .iter()
            .map(|row| row.iter().map(value_to_js_deep).collect::<js_sys::Array>())
            .collect::<js_sys::Array>()
            .into(),
        other => value_to_js(other),
    }
}

/// An engine result as JS; an array comes back as its top-left value, the way
/// a single cell shows one. A function is not something JS can be handed, and
/// never reaches a cell anyway, so it is `null`.
fn value_to_js(value: &Value) -> JsValue {
    match value {
        Value::Blank | Value::Lambda(_) => JsValue::NULL,
        Value::Number(n) => JsValue::from_f64(*n),
        Value::Text(t) => JsValue::from_str(t),
        Value::Bool(b) => JsValue::from_bool(*b),
        Value::Error(e) => JsValue::from_str(e.as_str()),
        Value::Array(rows) => rows
            .first()
            .and_then(|r| r.first())
            .map_or(JsValue::NULL, value_to_js),
    }
}
