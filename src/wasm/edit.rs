//! Editing the grid: rows, columns and cells inserted and removed, blocks
//! copied, moved, sorted and filled, sheets moved and removed. Every one of
//! these rewrites formulas across the workbook, so each drops the dependency
//! index when it succeeds.

use super::convert::{array, col_at, row_at};
use super::{Book, js};
use crate::coordinate::{CellRef, Col, Range, Row};
use crate::edit::{Axis, CopyOrigin, SortBy, SortKey, SortOptions, SplitBy};
use wasm_bindgen::prelude::*;

#[wasm_bindgen]
impl Book {
    /// Inserts `count` rows above row `at`, moving everything below down.
    /// `copyOrigin` - `"before"` (the row above), `"after"` (the row below)
    /// or `"none"` - says whose formatting the new rows take. Left out, it is
    /// `"before"`, as in Excel.
    ///
    /// Formulas across the whole workbook follow the cells they read, and so
    /// do merges, links, validations, tables and drawings. `at` is the number
    /// a user sees.
    #[wasm_bindgen(js_name = insertRows)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an optional string crosses the wasm boundary owned"
    )]
    pub fn insert_rows(
        &mut self,
        sheet: usize,
        at: u32,
        count: u32,
        #[wasm_bindgen(
            js_name = "copyOrigin",
            unchecked_optional_param_type = "\"before\" | \"after\" | \"none\""
        )]
        copy_origin: Option<String>,
    ) -> Result<(), JsError> {
        let row = Row::from_one_based(u64::from(at)).map_err(js)?;
        let origin = copy_origin_of(copy_origin.as_deref())?;
        let done = crate::edit::insert_rows_with(&mut self.workbook, sheet, row, count, origin);
        self.edited(done)
    }

    /// Removes `count` rows from row `at` down. A formula that read a removed
    /// cell reads `#REF!` afterwards.
    #[wasm_bindgen(js_name = removeRows)]
    pub fn remove_rows(&mut self, sheet: usize, at: u32, count: u32) -> Result<(), JsError> {
        let row = Row::from_one_based(u64::from(at)).map_err(js)?;
        let done = crate::edit::remove_rows(&mut self.workbook, sheet, row, count);
        self.edited(done)
    }

    /// Inserts `count` columns to the left of column `at`, 1-based.
    /// `copyOrigin` works as in `insertRows`, with left and right.
    #[wasm_bindgen(js_name = insertColumns)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an optional string crosses the wasm boundary owned"
    )]
    pub fn insert_columns(
        &mut self,
        sheet: usize,
        at: u32,
        count: u32,
        #[wasm_bindgen(
            js_name = "copyOrigin",
            unchecked_optional_param_type = "\"before\" | \"after\" | \"none\""
        )]
        copy_origin: Option<String>,
    ) -> Result<(), JsError> {
        let col = Col::from_one_based(u64::from(at)).map_err(js)?;
        let origin = copy_origin_of(copy_origin.as_deref())?;
        let done = crate::edit::insert_columns_with(&mut self.workbook, sheet, col, count, origin);
        self.edited(done)
    }

    /// Removes `count` columns from column `at`, 1-based.
    #[wasm_bindgen(js_name = removeColumns)]
    pub fn remove_columns(&mut self, sheet: usize, at: u32, count: u32) -> Result<(), JsError> {
        let col = Col::from_one_based(u64::from(at)).map_err(js)?;
        let done = crate::edit::remove_columns(&mut self.workbook, sheet, col, count);
        self.edited(done)
    }

    /// Inserts rows at many places in one pass: `places` is `[at, count]`
    /// pairs, numbered as the sheet is before the call, in any order.
    /// `book.insertRowsMany(0, [[5, 1], [9, 1]])` puts a row above each of
    /// rows 5 and 9 - the old row 9, now row 10. Each new block takes its
    /// formatting from its own neighbour, as `copyOrigin` says.
    ///
    /// The result is that of the single inserts made from the bottom up, but
    /// the workbook is walked once rather than once per place: four thousand
    /// subtotal rows into forty thousand take milliseconds, not a minute.
    #[wasm_bindgen(js_name = insertRowsMany)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an optional string crosses the wasm boundary owned"
    )]
    pub fn insert_rows_many(
        &mut self,
        sheet: usize,
        #[wasm_bindgen(unchecked_param_type = "[number, number][]")] places: &JsValue,
        #[wasm_bindgen(
            js_name = "copyOrigin",
            unchecked_optional_param_type = "\"before\" | \"after\" | \"none\""
        )]
        copy_origin: Option<String>,
    ) -> Result<(), JsError> {
        let at = places_of(places, row_at)?;
        let origin = copy_origin_of(copy_origin.as_deref())?;
        let done = crate::edit::insert_rows_many(&mut self.workbook, sheet, &at, origin);
        self.edited(done)
    }

    /// Removes rows at many places in one pass: `[at, count]` pairs, numbered
    /// as the sheet is before the call. Blocks that overlap are merged.
    #[wasm_bindgen(js_name = removeRowsMany)]
    pub fn remove_rows_many(
        &mut self,
        sheet: usize,
        #[wasm_bindgen(unchecked_param_type = "[number, number][]")] places: &JsValue,
    ) -> Result<(), JsError> {
        let spans = places_of(places, row_at)?;
        let done = crate::edit::remove_rows_many(&mut self.workbook, sheet, &spans);
        self.edited(done)
    }

    /// `insertRowsMany` for columns.
    #[wasm_bindgen(js_name = insertColumnsMany)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an optional string crosses the wasm boundary owned"
    )]
    pub fn insert_columns_many(
        &mut self,
        sheet: usize,
        #[wasm_bindgen(unchecked_param_type = "[number, number][]")] places: &JsValue,
        #[wasm_bindgen(
            js_name = "copyOrigin",
            unchecked_optional_param_type = "\"before\" | \"after\" | \"none\""
        )]
        copy_origin: Option<String>,
    ) -> Result<(), JsError> {
        let at = places_of(places, col_at)?;
        let origin = copy_origin_of(copy_origin.as_deref())?;
        let done = crate::edit::insert_columns_many(&mut self.workbook, sheet, &at, origin);
        self.edited(done)
    }

    /// `removeRowsMany` for columns.
    #[wasm_bindgen(js_name = removeColumnsMany)]
    pub fn remove_columns_many(
        &mut self,
        sheet: usize,
        #[wasm_bindgen(unchecked_param_type = "[number, number][]")] places: &JsValue,
    ) -> Result<(), JsError> {
        let spans = places_of(places, col_at)?;
        let done = crate::edit::remove_columns_many(&mut self.workbook, sheet, &spans);
        self.edited(done)
    }

    /// Inserts blank cells over a range, pushing what was there `"down"` or
    /// `"right"` - Excel's "Insert Cells", which moves part of a row rather
    /// than the whole of it. `copyOrigin` works as in `insertRows`.
    #[wasm_bindgen(js_name = insertCells)]
    #[expect(
        clippy::needless_pass_by_value,
        reason = "an optional string crosses the wasm boundary owned"
    )]
    pub fn insert_cells(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_param_type = "\"down\" | \"right\"")] shift: &str,
        #[wasm_bindgen(
            js_name = "copyOrigin",
            unchecked_optional_param_type = "\"before\" | \"after\" | \"none\""
        )]
        copy_origin: Option<String>,
    ) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let origin = copy_origin_of(copy_origin.as_deref())?;
        let done = crate::edit::insert_cells_with(
            &mut self.workbook,
            sheet,
            area,
            axis_of(shift)?,
            origin,
        );
        self.edited(done)
    }

    /// Removes the cells of a range, pulling the rest back over the hole from
    /// below (`"up"`) or from the right (`"left"`).
    #[wasm_bindgen(js_name = removeCells)]
    pub fn remove_cells(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_param_type = "\"up\" | \"left\"")] shift: &str,
    ) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let done = crate::edit::remove_cells(&mut self.workbook, sheet, area, axis_of(shift)?);
        self.edited(done)
    }

    /// Copies a rectangle of cells so its top left corner lands on `to`,
    /// rewriting the formulas as Excel rewrites a copied formula: `=A1` one
    /// row down reads `=A2`, `$A$1` stays put.
    ///
    /// `toSheet` defaults to the sheet copied from. A cell the source leaves
    /// empty empties the target, the way pasting a block does.
    #[wasm_bindgen(js_name = copyRange)]
    pub fn copy_range(
        &mut self,
        sheet: usize,
        range: &str,
        to: &str,
        to_sheet: Option<usize>,
    ) -> Result<(), JsError> {
        let (area, at) = (
            Range::parse(range).map_err(js)?,
            CellRef::parse(to).map_err(js)?,
        );
        let done = crate::edit::copy_range(
            &mut self.workbook,
            sheet,
            area,
            to_sheet.unwrap_or(sheet),
            at,
        );
        self.edited(done)
    }

    /// Moves a rectangle of cells, leaving the source empty.
    ///
    /// A moved formula keeps reading the cells it read; the formulas elsewhere
    /// that read *these* cells follow them instead, as they do when Excel cuts
    /// and pastes.
    #[wasm_bindgen(js_name = moveRange)]
    pub fn move_range(
        &mut self,
        sheet: usize,
        range: &str,
        to: &str,
        to_sheet: Option<usize>,
    ) -> Result<(), JsError> {
        let (area, at) = (
            Range::parse(range).map_err(js)?,
            CellRef::parse(to).map_err(js)?,
        );
        let done = crate::edit::move_range(
            &mut self.workbook,
            sheet,
            area,
            to_sheet.unwrap_or(sheet),
            at,
        );
        self.edited(done)
    }

    /// Sorts the rows of a range, or its columns. A key is a column number of
    /// the sheet, 1-based (a row number when sorting columns), or the text of
    /// a header when `options.header` says the first line is one; a minus
    /// sorts it largest first: `[2, -4]`, `["Region", "-Amount"]`. The first
    /// key decides first.
    ///
    /// ```js
    /// book.sortRange(0, "A1:D100", ["Region", "-Amount"], { header: true });
    /// ```
    #[wasm_bindgen(js_name = sortRange)]
    pub fn sort_range(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_param_type = "SortKeys")] keys: &JsValue,
        #[wasm_bindgen(unchecked_optional_param_type = "SortRangeOptions")] options: &JsValue,
    ) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let flag = |name: &str| super::convert::field(options, name).is_some_and(|v| v.is_truthy());
        let options = SortOptions {
            header: flag("header"),
            orientation: if flag("byColumns") {
                Axis::Columns
            } else {
                Axis::Rows
            },
        };
        let keys = sort_keys_of(keys)?;
        let done = crate::edit::sort_range_with(&mut self.workbook, sheet, area, &keys, options);
        self.edited(done)
    }

    /// Sorts the data rows of a table, found by name anywhere in the book;
    /// its header and totals rows stay put. Keys name its columns, as in
    /// `sortRange`: `book.sortTable("Sales", ["-Amount"])`.
    #[wasm_bindgen(js_name = sortTable)]
    pub fn sort_table(
        &mut self,
        name: &str,
        #[wasm_bindgen(unchecked_param_type = "SortKeys")] keys: &JsValue,
    ) -> Result<(), JsError> {
        let keys = sort_keys_of(keys)?;
        let done = crate::edit::sort_table(&mut self.workbook, name, &keys);
        self.edited(done)
    }

    /// Continues what the first cells of a range start, as dragging the fill
    /// handle does: `1, 2` goes on `3, 4`, `Кв1` to `Кв2`, `Jan` to `Feb`, a
    /// date by a day; anything else repeats. `"down"` fills each column,
    /// `"right"` each row; `"up"` and `"left"` run the series back from the
    /// last cells.
    #[wasm_bindgen(js_name = fillSeries)]
    pub fn fill_series(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_param_type = "\"down\" | \"right\" | \"up\" | \"left\"")]
        direction: &str,
    ) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let fill = match direction {
            "down" | "right" => crate::edit::fill_series,
            "up" | "left" => crate::edit::fill_series_back,
            _ => {
                return Err(JsError::new(
                    r#"a fill goes "down", "right", "up" or "left""#,
                ));
            }
        };
        let done = fill(&mut self.workbook, sheet, area, axis_of(direction)?);
        self.edited(done)
    }

    /// Copies the first row of a range into the rows below it, as Ctrl+D.
    #[wasm_bindgen(js_name = fillDown)]
    pub fn fill_down(&mut self, sheet: usize, range: &str) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let done = crate::edit::fill(&mut self.workbook, sheet, area, Axis::Rows);
        self.edited(done)
    }

    /// Copies the first column of a range into the columns to its right, as
    /// Ctrl+R.
    #[wasm_bindgen(js_name = fillRight)]
    pub fn fill_right(&mut self, sheet: usize, range: &str) -> Result<(), JsError> {
        let area = Range::parse(range).map_err(js)?;
        let done = crate::edit::fill(&mut self.workbook, sheet, area, Axis::Columns);
        self.edited(done)
    }

    /// Moves a sheet along the tab bar. Formulas are untouched - a sheet is
    /// named, not numbered - but every index moves, this one included, so read
    /// `sheetIndex` again afterwards.
    #[wasm_bindgen(js_name = moveSheet)]
    pub fn move_sheet(&mut self, from: usize, to: usize) -> Result<(), JsError> {
        crate::edit::move_sheet(&mut self.workbook, from, to).map_err(js)
    }

    /// Removes a sheet. References to it from anywhere in the workbook become
    /// `#REF!`, the way Excel writes them, and the sheets after it shift down
    /// by one index.
    #[wasm_bindgen(js_name = removeSheet)]
    pub fn remove_sheet(&mut self, sheet: usize) -> Result<(), JsError> {
        let done = crate::edit::remove_sheet(&mut self.workbook, sheet);
        self.edited(done)
    }
    /// Removes the rows of a range that repeat a row above them, as Excel's
    /// Remove Duplicates does: cells are compared by the text they show,
    /// without regard to case. `options.columns` are the 1-based sheet
    /// columns compared (all of the range when left out), and
    /// `options.header` keeps the first row out of it. Returns
    /// `{ removed, remaining }`.
    ///
    /// ```js
    /// book.removeDuplicates(0, "A1:C500", { columns: [1, 3], header: true });
    /// ```
    #[wasm_bindgen(js_name = removeDuplicates, unchecked_return_type = "DuplicatesRemoved")]
    pub fn remove_duplicates(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_optional_param_type = "RemoveDuplicatesOptions")]
        options: &JsValue,
    ) -> Result<JsValue, JsError> {
        let area = Range::parse(range).map_err(js)?;
        let header = super::convert::field(options, "header").is_some_and(|v| v.is_truthy());
        let columns = match super::convert::field(options, "columns") {
            Some(columns) => array(&columns, "columns")?
                .iter()
                .map(|c| {
                    c.as_f64()
                        .filter(|n| n.fract() == 0.0 && *n >= 1.0 && *n <= f64::from(u32::MAX))
                        .ok_or_else(|| JsError::new("a column counts from 1"))
                        .and_then(|n| {
                            #[expect(
                                clippy::cast_possible_truncation,
                                clippy::cast_sign_loss,
                                reason = "checked whole and in range just above"
                            )]
                            col_at(n as u32)
                        })
                })
                .collect::<Result<Vec<_>, _>>()?,
            None => Vec::new(),
        };
        let done =
            crate::edit::remove_duplicates(&mut self.workbook, sheet, area, &columns, header)
                .map_err(js)?;
        self.forget_dependencies();
        Ok(super::convert::object(&[
            ("removed", super::convert::count(done.removed)),
            ("remaining", super::convert::count(done.remaining)),
        ]))
    }

    /// Splits the text of each cell in the first column of `range` across
    /// the columns to its right, as Excel's Text to Columns does. Pieces that
    /// read as numbers, booleans or errors become them. Returns how many
    /// columns the widest row took.
    ///
    /// `options.delimiters` is a string of delimiter characters (`";"`,
    /// `",\t"`), or `options.widths` the character positions to cut before
    /// for fixed width; `mergeConsecutive` treats a run of delimiters as one,
    /// `qualifier` is the quote (`"` unless given; `null` for none), and
    /// `destination` is the first cell written (the range's own when left
    /// out).
    ///
    /// ```js
    /// book.textToColumns(0, "A2:A100", { delimiters: ";" });
    /// book.textToColumns(0, "A2:A100", { widths: [4, 10], destination: "C2" });
    /// ```
    #[wasm_bindgen(js_name = textToColumns)]
    pub fn text_to_columns(
        &mut self,
        sheet: usize,
        range: &str,
        #[wasm_bindgen(unchecked_param_type = "TextToColumnsOptions")] options: &JsValue,
    ) -> Result<u32, JsError> {
        use super::convert::field;
        let area = Range::parse(range).map_err(js)?;
        let destination = match field(options, "destination").and_then(|v| v.as_string()) {
            Some(address) => CellRef::parse(&address).map_err(js)?,
            None => area.start,
        };
        let split = if let Some(widths) = field(options, "widths") {
            let widths = array(&widths, "widths")?
                .iter()
                .map(|w| {
                    w.as_f64()
                        .filter(|n| n.fract() == 0.0 && *n >= 0.0 && *n <= f64::from(u32::MAX))
                        .ok_or_else(|| JsError::new("a width is a whole number of characters"))
                        .map(|n| {
                            #[expect(
                                clippy::cast_possible_truncation,
                                clippy::cast_sign_loss,
                                reason = "checked whole and in range just above"
                            )]
                            let n = n as usize;
                            n
                        })
                })
                .collect::<Result<Vec<_>, _>>()?;
            SplitBy::FixedWidth(widths)
        } else {
            let delimiters: Vec<char> = field(options, "delimiters")
                .and_then(|v| v.as_string())
                .map_or_else(|| vec!['\t'], |d| d.chars().collect());
            if delimiters.is_empty() {
                return Err(JsError::new("textToColumns needs delimiters or widths"));
            }
            let qualifier = match field(options, "qualifier") {
                None => Some('"'),
                Some(q) if q.is_null() => None,
                Some(q) => Some(super::convert::one_char(&q)?),
            };
            SplitBy::Delimited {
                delimiters,
                merge_consecutive: field(options, "mergeConsecutive")
                    .is_some_and(|v| v.is_truthy()),
                qualifier,
            }
        };
        let width =
            crate::edit::text_to_columns(&mut self.workbook, sheet, area, destination, &split)
                .map_err(js)?;
        self.forget_dependencies();
        Ok(width)
    }

    /// Adds a data validation rule to a sheet, as Data - Data Validation
    /// does: `type` is `"whole"`, `"decimal"`, `"list"`, `"date"`, `"time"`,
    /// `"textLength"` or `"custom"`, `operator` one of `"between"` (the
    /// default), `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`,
    /// `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`. Formulas are
    /// written for the first cell of `sqref`; a list is `"\"Yes,No\""` or a
    /// range such as `"$E$1:$E$5"`.
    ///
    /// ```js
    /// book.addDataValidation(0, { sqref: "B2:B100", type: "list", formula1: "\"Yes,No\"" });
    /// ```
    #[wasm_bindgen(js_name = addDataValidation)]
    pub fn add_data_validation(
        &mut self,
        sheet: usize,
        #[wasm_bindgen(unchecked_param_type = "NewDataValidation")] rule: &JsValue,
    ) -> Result<(), JsError> {
        use super::convert::field;
        use crate::model::{
            DataValidation, ValidationErrorStyle, ValidationOperator, ValidationType,
        };
        let text = |name: &str| field(rule, name).and_then(|v| v.as_string());
        let flag = |name: &str, default: bool| field(rule, name).map_or(default, |v| v.is_truthy());
        let sqref = text("sqref")
            .ok_or_else(|| JsError::new("a rule needs its sqref, such as \"A1:A10\""))?
            .split_whitespace()
            .map(Range::parse)
            .collect::<crate::error::Result<Vec<_>>>()
            .map_err(js)?;
        let kind = ValidationType::parse(&text("type").unwrap_or_default());
        if kind == ValidationType::None {
            return Err(JsError::new(
                "type is whole, decimal, list, date, time, textLength or custom",
            ));
        }
        let formula1 = text("formula1").unwrap_or_default();
        let formula1 = formula1.strip_prefix('=').unwrap_or(&formula1).to_owned();
        let formula2 = text("formula2").unwrap_or_default();
        let formula2 = formula2.strip_prefix('=').unwrap_or(&formula2).to_owned();
        let operator = match text("operator") {
            None => ValidationOperator::Between,
            Some(name) => {
                let operator = ValidationOperator::parse(&name);
                if operator.as_str() != name {
                    return Err(JsError::new(&format!(
                        "{name:?} is not a validation operator"
                    )));
                }
                operator
            }
        };
        let validation = DataValidation {
            sqref,
            kind,
            operator,
            formula1,
            formula2,
            allow_blank: flag("allowBlank", true),
            hide_drop_down: !flag("showDropDown", true),
            show_input_message: flag("showInputMessage", true),
            show_error_message: flag("showErrorMessage", true),
            error_title: text("errorTitle").unwrap_or_default(),
            error: text("error").unwrap_or_default(),
            prompt_title: text("promptTitle").unwrap_or_default(),
            prompt: text("prompt").unwrap_or_default(),
            error_style: ValidationErrorStyle::parse(&text("errorStyle").unwrap_or_default()),
        };
        self.sheet_mut(sheet)?.data_validations.push(validation);
        Ok(())
    }
}

impl Book {
    /// The end of every edit: its error for JS, or the dependency index
    /// dropped, since every address in it may have moved.
    fn edited(&mut self, done: crate::error::Result<()>) -> Result<(), JsError> {
        done.map_err(js)?;
        self.forget_dependencies();
        Ok(())
    }
}

/// Which way a shift or a fill runs.
fn axis_of(direction: &str) -> Result<Axis, JsError> {
    match direction {
        "down" | "up" => Ok(Axis::Rows),
        "right" | "left" => Ok(Axis::Columns),
        _ => Err(JsError::new(
            r#"a shift is "down", "up", "right" or "left""#,
        )),
    }
}

/// Reads the `[at, count]` pairs of the batch edits, `at` 1-based.
fn places_of<T>(
    places: &JsValue,
    line: fn(u32) -> Result<T, JsError>,
) -> Result<Vec<(T, u32)>, JsError> {
    let number = |v: JsValue| {
        v.as_f64()
            .filter(|n| n.fract() == 0.0 && (0.0..=f64::from(u32::MAX)).contains(n))
            .ok_or_else(|| JsError::new("a place is [at, count], two whole numbers"))
    };
    array(places, "places")?
        .iter()
        .map(|pair| {
            let pair = array(&pair, "a place")?;
            #[expect(
                clippy::cast_possible_truncation,
                clippy::cast_sign_loss,
                reason = "checked whole and in range just above"
            )]
            let (at, count) = (number(pair.get(0))? as u32, number(pair.get(1))? as u32);
            Ok((line(at)?, count))
        })
        .collect()
}

/// Reads the `copyOrigin` argument of the inserts.
fn copy_origin_of(name: Option<&str>) -> Result<CopyOrigin, JsError> {
    match name {
        Some("none") => Ok(CopyOrigin::Blank),
        None | Some("before") => Ok(CopyOrigin::Before),
        Some("after") => Ok(CopyOrigin::After),
        Some(other) => Err(JsError::new(&format!(
            "copyOrigin must be \"before\", \"after\" or \"none\", not {other:?}"
        ))),
    }
}

/// Reads the keys of `sortRange`: a number is a 1-based line of the sheet, a
/// string a header; a minus in front sorts largest first.
fn sort_keys_of(keys: &JsValue) -> Result<Vec<SortKey>, JsError> {
    array(keys, "sort keys")?
        .iter()
        .map(|key| {
            if let Some(n) = key.as_f64() {
                #[expect(
                    clippy::cast_possible_truncation,
                    clippy::cast_sign_loss,
                    reason = "a line number of the sheet, saturated if it is not one"
                )]
                let line = (n.abs() as u32)
                    .checked_sub(1)
                    .ok_or_else(|| JsError::new("a sort key counts from 1"))?;
                return Ok(SortKey {
                    by: SortBy::Line(line),
                    descending: n < 0.0,
                });
            }
            let text = key
                .as_string()
                .ok_or_else(|| JsError::new("a sort key is a number or a header"))?;
            Ok(match text.strip_prefix('-') {
                Some(name) => SortKey::header(name).descending(),
                None => SortKey::header(text),
            })
        })
        .collect()
}
