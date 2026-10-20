//! The Data tab's range tools that rewrite cells in place: Remove Duplicates
//! and Text to Columns.
//!
//! Both follow the same rules as [`sort_range`](super::sort_range): cells
//! travel with their formatting, a formula that moves is rewritten as if
//! copied to its new place, and formulas elsewhere are not retargeted.
//!
//! The behaviour was ported from the data commands of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use std::collections::HashSet;

use crate::coordinate::{CellRef, Col, Range, Row, shift_references};
use crate::error::{Error, Result};
use crate::model::{Cell, CellValue, Spreadsheet};

/// What [`remove_duplicates`] did.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Deduplicated {
    /// Rows dropped as repeats of a row above them.
    pub removed: usize,
    /// Data rows left, the header not counted.
    pub remaining: usize,
}

/// Removes the rows of `area` that repeat a row above them: Excel's
/// Data - Remove Duplicates.
///
/// Two rows are the same when every column in `columns` *shows* the same text,
/// without regard to case: as in Excel, a date in two formats is not a
/// duplicate, while `1.01` and `1.02` both shown as `1.0` are. An empty
/// `columns` compares every column of `area`; a column is counted on the
/// sheet, not inside the range. With `header`, the first row stays where it
/// is and takes no part.
///
/// The first of each set of equal rows is kept, the rows below close up, and
/// the rows freed at the bottom of `area` are emptied. Only the cells inside
/// `area` move: the rest of the sheet does not. The part of `area` past the
/// last cell holding anything is not looked at, so a whole column costs no
/// more than the data in it.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::InvalidRange`] if a column lies outside `area` or a merged range
/// crosses it.
pub fn remove_duplicates(
    book: &mut Spreadsheet,
    sheet: usize,
    area: Range,
    columns: &[Col],
    header: bool,
) -> Result<Deduplicated> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    if let Some(column) = columns
        .iter()
        .find(|c| **c < area.start.col || **c > area.end.col)
    {
        return Err(Error::InvalidRange(format!(
            "column {} is outside {area}",
            column.to_letters()
        )));
    }
    if ws.merges.iter().any(|m| m.intersects(&area)) {
        return Err(Error::InvalidRange(format!(
            "a merged range crosses {area}; unmerge it to remove duplicates"
        )));
    }
    let Some(used) = ws.dimension() else {
        return Ok(Deduplicated::default());
    };
    let first = area.start.row.index() + u32::from(header);
    let last = area.end.row.index().min(used.end.row.index());
    if first > last {
        return Ok(Deduplicated::default());
    }
    let keys: Vec<Col> = if columns.is_empty() {
        (area.start.col.index()..=area.end.col.index())
            .filter_map(Col::new)
            .collect()
    } else {
        columns.to_vec()
    };

    let mut seen = HashSet::new();
    let mut keep = Vec::new();
    for index in first..=last {
        let Some(row) = Row::new(index) else { continue };
        let key: Vec<String> = keys
            .iter()
            .map(|&col| book.formatted(sheet, CellRef::new(col, row)).to_lowercase())
            .collect();
        if seen.insert(key) {
            keep.push(index);
        }
    }
    let total = usize::try_from(last - first + 1).unwrap_or(usize::MAX);
    let report = Deduplicated {
        removed: total - keep.len(),
        remaining: keep.len(),
    };
    if report.removed == 0 {
        return Ok(report);
    }

    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let cross: Vec<Col> = (area.start.col.index()..=area.end.col.index())
        .filter_map(Col::new)
        .collect();
    // Every row is taken out first, so a row never lands on one not yet read.
    let taken: Vec<(u32, Vec<(Col, Cell)>)> = (first..=last)
        .filter_map(|index| {
            let row = Row::new(index)?;
            let cells = cross
                .iter()
                .filter_map(|&col| Some((col, ws.remove(CellRef::new(col, row))?)))
                .collect();
            Some((index, cells))
        })
        .collect();
    let mut taken: std::collections::HashMap<u32, Vec<(Col, Cell)>> = taken.into_iter().collect();
    for (to, from) in (first..).zip(&keep) {
        let (Some(row), Some(cells)) = (Row::new(to), taken.remove(from)) else {
            continue;
        };
        let delta = i64::from(to) - i64::from(*from);
        for (col, mut cell) in cells {
            if delta != 0
                && let CellValue::Formula { formula, .. } = &cell.value
            {
                cell.value = CellValue::Formula {
                    formula: shift_references(formula, 0, delta),
                    cached: None,
                };
            }
            *ws.entry(CellRef::new(col, row)) = cell;
        }
    }
    Ok(report)
}

/// How [`text_to_columns`] cuts a cell's text.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SplitBy {
    /// At any of these characters, as a CSV line is cut.
    Delimited {
        /// The characters that separate fields: `'\t'`, `';'`, `','`, `' '`
        /// or any other.
        delimiters: Vec<char>,
        /// Whether a run of delimiters counts as one ("Treat consecutive
        /// delimiters as one").
        merge_consecutive: bool,
        /// The character text is quoted with, inside which a delimiter is
        /// text; doubled, it stands for itself. `None` quotes nothing.
        qualifier: Option<char>,
    },
    /// Before each of these character positions, counted from 0. Each piece
    /// is trimmed of the spaces that padded it.
    FixedWidth(Vec<usize>),
}

impl SplitBy {
    /// Fields separated by `delimiter`, quoted with `"`.
    #[must_use]
    pub fn delimiter(delimiter: char) -> Self {
        Self::Delimited {
            delimiters: vec![delimiter],
            merge_consecutive: false,
            qualifier: Some('"'),
        }
    }

    /// Cuts `text` into its pieces.
    #[must_use]
    pub fn split(&self, text: &str) -> Vec<String> {
        match self {
            Self::Delimited {
                delimiters,
                merge_consecutive,
                qualifier,
            } => split_delimited(text, delimiters, *merge_consecutive, *qualifier),
            Self::FixedWidth(breaks) => {
                let chars: Vec<char> = text.chars().collect();
                let mut cuts = vec![0];
                let mut breaks: Vec<usize> = breaks
                    .iter()
                    .copied()
                    .filter(|&b| b > 0 && b < chars.len())
                    .collect();
                breaks.sort_unstable();
                breaks.dedup();
                cuts.extend(breaks);
                cuts.push(chars.len());
                cuts.windows(2)
                    .map(|w| {
                        chars[w[0]..w[1]]
                            .iter()
                            .collect::<String>()
                            .trim()
                            .to_owned()
                    })
                    .collect()
            }
        }
    }
}

/// Splits the text of each cell in the first column of `source` across the
/// columns from `destination` rightwards: Excel's Data - Text to Columns,
/// with every column in the General format.
///
/// The text cut is what the cell shows. A piece that reads as a number,
/// `TRUE`, `FALSE` or an error becomes one, as a CSV field does; the rest is
/// text. An empty piece empties its cell. Cells written over keep their
/// style. `destination` may be `source`'s own first cell, which is what Excel
/// does unless told otherwise. Returns how many columns the widest row took.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a row splits into more columns than are
/// left right of `destination`; nothing is written then.
pub fn text_to_columns(
    book: &mut Spreadsheet,
    sheet: usize,
    source: Range,
    destination: CellRef,
    split: &SplitBy,
) -> Result<u32> {
    if book.sheet(sheet).is_none() {
        return Err(Error::SheetIndexOutOfRange(sheet));
    }
    let column = source.start.col;
    let rows: Vec<Vec<String>> = (source.start.row.index()..=source.end.row.index())
        .map(|index| {
            Row::new(index).map_or_else(Vec::new, |row| {
                let at = CellRef::new(column, row);
                match book.sheet(sheet).and_then(|ws| ws.get(at)) {
                    Some(cell) if !cell.value.result().is_empty() => {
                        split.split(&book.formatted(sheet, at))
                    }
                    _ => Vec::new(),
                }
            })
        })
        .collect();
    let widest = rows.iter().map(Vec::len).max().unwrap_or(0);
    let widest = u32::try_from(widest).map_err(|_| Error::WouldPushOffSheet)?;
    if widest == 0 {
        return Ok(0);
    }
    if Col::new(destination.col.index() + widest - 1).is_none()
        || Row::new(destination.row.index() + source.height() - 1).is_none()
    {
        return Err(Error::WouldPushOffSheet);
    }
    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    for (i, pieces) in (0u32..).zip(rows) {
        // A row with nothing in its source cell is left alone, as in Excel.
        if pieces.is_empty() {
            continue;
        }
        let Some(row) = Row::new(destination.row.index() + i) else {
            continue;
        };
        for (j, piece) in (0u32..).zip(pieces) {
            let Some(col) = Col::new(destination.col.index() + j) else {
                continue;
            };
            let at = CellRef::new(col, row);
            if piece.is_empty() {
                match ws.get(at) {
                    Some(cell) if cell.style != crate::style::StyleId::default() => {
                        ws.entry(at).value = CellValue::Empty;
                    }
                    Some(_) => {
                        ws.remove(at);
                    }
                    None => {}
                }
            } else {
                ws.set(at, crate::reader::csv::value_of(&piece));
            }
        }
    }
    Ok(widest)
}

/// Cuts `text` at `delimiters`, honouring a text qualifier.
fn split_delimited(
    text: &str,
    delimiters: &[char],
    merge_consecutive: bool,
    qualifier: Option<char>,
) -> Vec<String> {
    let mut out = Vec::new();
    let mut current = String::new();
    let mut quoted = false;
    let mut after_delimiter = false;
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        if Some(c) == qualifier {
            if quoted && chars.peek() == Some(&c) {
                current.push(c);
                chars.next();
            } else {
                quoted = !quoted;
            }
            after_delimiter = false;
            continue;
        }
        if !quoted && delimiters.contains(&c) {
            if !(merge_consecutive && after_delimiter) {
                out.push(std::mem::take(&mut current));
            }
            after_delimiter = true;
            continue;
        }
        after_delimiter = false;
        current.push(c);
    }
    out.push(current);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn delimited_with_quotes_and_runs() {
        let by = SplitBy::Delimited {
            delimiters: vec![',', ' '],
            merge_consecutive: true,
            qualifier: Some('"'),
        };
        assert_eq!(
            by.split(r#"a,  "b, c" ,"say ""hi""""#),
            ["a", "b, c", "say \"hi\""].map(String::from)
        );
        assert_eq!(
            SplitBy::delimiter(';').split("1;;x"),
            ["1", "", "x"].map(String::from)
        );
        assert_eq!(
            SplitBy::delimiter(',').split(r#""a,b",c"#),
            ["a,b", "c"].map(String::from)
        );
    }

    #[test]
    fn fixed_width_trims_and_ignores_breaks_past_the_end() {
        let by = SplitBy::FixedWidth(vec![3, 8, 40]);
        assert_eq!(
            by.split("ab cdef gh"),
            ["ab", "cdef", "gh"].map(String::from)
        );
        assert_eq!(by.split("ab"), ["ab"].map(String::from));
    }
}

// ---------------------------------------------------------- advanced filter

/// Where [`advanced_filter`] puts what passes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterOutput {
    /// Hide the rows of the list that do not pass ("Filter the list,
    /// in-place").
    InPlace,
    /// Copy the header and the rows that pass to here ("Copy to another
    /// location").
    CopyTo(usize, CellRef),
}

/// Whether `text` matches an Excel wildcard pattern, without regard to case:
/// `*` any run, `?` one character, `~` before either stands for it.
fn wildcard(pattern: &str, text: &str) -> bool {
    let pattern: Vec<char> = pattern.to_lowercase().chars().collect();
    let text: Vec<char> = text.to_lowercase().chars().collect();
    wildcard_chars(&pattern, &text)
}

fn wildcard_chars(p: &[char], t: &[char]) -> bool {
    match p.first() {
        None => t.is_empty(),
        Some('*') => (0..=t.len()).any(|i| wildcard_chars(&p[1..], &t[i..])),
        Some('?') => !t.is_empty() && wildcard_chars(&p[1..], &t[1..]),
        Some('~') if p.len() > 1 => t.first() == Some(&p[1]) && wildcard_chars(&p[2..], &t[1..]),
        Some(c) => t.first() == Some(c) && wildcard_chars(&p[1..], &t[1..]),
    }
}

/// Whether a cell passes one criteria cell, as Excel reads them: `>5`,
/// `<>East`, `=` for an empty cell, `=East` for exactly East, and plain text
/// for anything that begins with it; `*` and `?` work in text.
fn passes(criterion: &str, shown: &str, number: Option<f64>) -> bool {
    let criterion = criterion.trim();
    if criterion.is_empty() {
        return true;
    }
    let (op, operand) = ["<>", ">=", "<=", "=", ">", "<"]
        .iter()
        .find_map(|op| criterion.strip_prefix(op).map(|rest| (*op, rest)))
        .unwrap_or(("", criterion));
    let wanted = operand.trim().parse::<f64>().ok();
    let order = match (number, wanted) {
        (Some(n), Some(w)) => Some(n.total_cmp(&w)),
        _ if op.starts_with(['<', '>']) && op != "<>" => {
            // Text against text compares by letters; a number never passes
            // a text bound or the other way round.
            (number.is_none() && wanted.is_none() && !shown.is_empty())
                .then(|| shown.to_lowercase().cmp(&operand.to_lowercase()))
        }
        _ => None,
    };
    let same = || match (number, wanted) {
        (Some(n), Some(w)) => n.total_cmp(&w).is_eq(),
        _ => wildcard(operand, shown),
    };
    match op {
        "" => match (number, wanted) {
            (Some(n), Some(w)) => n.total_cmp(&w).is_eq(),
            // Plain text means "begins with".
            _ => wildcard(&format!("{operand}*"), shown),
        },
        "=" if operand.is_empty() => shown.is_empty(),
        "=" => same(),
        "<>" if operand.is_empty() => !shown.is_empty(),
        "<>" => !same(),
        ">" => order.is_some_and(std::cmp::Ordering::is_gt),
        "<" => order.is_some_and(std::cmp::Ordering::is_lt),
        ">=" => order.is_some_and(std::cmp::Ordering::is_ge),
        "<=" => order.is_some_and(std::cmp::Ordering::is_le),
        _ => false,
    }
}

/// Filters a list by a criteria range: Excel's Data - Advanced.
///
/// `list` is the list with its header row; `criteria` is a header row
/// naming some of its columns (matched without regard to case) and one or
/// more rows of conditions under it. Conditions in one row must all hold;
/// a row passes if any criteria row holds for it. A criteria column whose
/// header names no column of the list is an error. With `unique`, a row
/// that shows the same as one above it is left out too ("Unique records
/// only").
///
/// In place, the rows that fail are hidden and those that pass shown;
/// copied, the header and the passing rows are written from the
/// destination down, and the cells below in those columns cleared. Returns
/// how many rows passed.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] for a criteria header the list lacks or a copy
/// that would land inside the list.
pub fn advanced_filter(
    book: &mut Spreadsheet,
    sheet: usize,
    list: Range,
    criteria: Range,
    output: FilterOutput,
    unique: bool,
) -> Result<usize> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let header = list.start.row;
    let last_row = ws.dimension().map_or_else(
        || header.index(),
        |used| list.end.row.index().min(used.end.row.index()),
    );
    let columns: Vec<Col> = (list.start.col.index()..=list.end.col.index())
        .filter_map(Col::new)
        .collect();
    let shown = |c: Col, r: Row| book.formatted(sheet, CellRef::new(c, r));
    let names: Vec<String> = columns
        .iter()
        .map(|&c| shown(c, header).to_lowercase())
        .collect();
    // Each criteria column: which list column it tests.
    let mut tests: Vec<(Col, Col)> = Vec::new();
    for index in criteria.start.col.index()..=criteria.end.col.index() {
        let Some(col) = Col::new(index) else { continue };
        let name = shown(col, criteria.start.row).to_lowercase();
        if name.is_empty() {
            continue;
        }
        let at = names
            .iter()
            .position(|n| *n == name)
            .ok_or_else(|| Error::InvalidRange(format!("the list has no column {name:?}")))?;
        tests.push((col, columns[at]));
    }
    let criteria_rows: Vec<Vec<(Col, String)>> = (criteria.start.row.index() + 1
        ..=criteria.end.row.index())
        .filter_map(Row::new)
        .map(|r| tests.iter().map(|&(cc, lc)| (lc, shown(cc, r))).collect())
        .collect();

    let mut seen = HashSet::new();
    let mut passed: Vec<Row> = Vec::new();
    let mut failed: Vec<Row> = Vec::new();
    for index in header.index() + 1..=last_row {
        let Some(row) = Row::new(index) else { continue };
        let ok = criteria_rows.is_empty()
            || criteria_rows.iter().any(|conditions| {
                conditions.iter().all(|(col, criterion)| {
                    let at = CellRef::new(*col, row);
                    let number = match ws.get(at).map(|c| c.value.result()) {
                        Some(CellValue::Number(n)) => Some(*n),
                        _ => None,
                    };
                    passes(criterion, &shown(*col, row), number)
                })
            });
        let fresh = !unique || {
            let key: Vec<String> = columns
                .iter()
                .map(|&c| shown(c, row).to_lowercase())
                .collect();
            !ok || seen.insert(key)
        };
        if ok && fresh {
            passed.push(row);
        } else {
            failed.push(row);
        }
    }

    match output {
        FilterOutput::InPlace => {
            let ws = book
                .sheet_mut(sheet)
                .ok_or(Error::SheetIndexOutOfRange(sheet))?;
            for row in &passed {
                if let Some(p) = ws.rows.get_mut(row) {
                    p.hidden = false;
                }
            }
            for row in failed {
                ws.rows.entry(row).or_default().hidden = true;
            }
        }
        FilterOutput::CopyTo(to_sheet, to) => {
            copy_rows(book, sheet, list, &columns, header, &passed, (to_sheet, to))?;
        }
    }
    Ok(passed.len())
}

/// Writes the header and the rows that passed an advanced filter from `to`
/// down, clearing what lay below in those columns.
fn copy_rows(
    book: &mut Spreadsheet,
    sheet: usize,
    list: Range,
    columns: &[Col],
    header: Row,
    passed: &[Row],
    (to_sheet, to): (usize, CellRef),
) -> Result<()> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let width = list.width();
    let height = u32::try_from(passed.len())
        .unwrap_or(u32::MAX)
        .saturating_add(1);
    let end = CellRef::new(
        Col::new(to.col.index() + width - 1).ok_or(Error::WouldPushOffSheet)?,
        Row::new(to.row.index() + height - 1).ok_or(Error::WouldPushOffSheet)?,
    );
    if to_sheet == sheet && Range::new(to, end).intersects(&list) {
        return Err(Error::InvalidRange(
            "the copy would land inside the list".into(),
        ));
    }
    let source_rows: Vec<Row> = std::iter::once(header)
        .chain(passed.iter().copied())
        .collect();
    let cells: Vec<Vec<Option<Cell>>> = source_rows
        .iter()
        .map(|&r| {
            columns
                .iter()
                .map(|&c| {
                    ws.get(CellRef::new(c, r)).map(|cell| Cell {
                        value: cell.value.result().clone(),
                        ..cell.clone()
                    })
                })
                .collect()
        })
        .collect();
    let target = book
        .sheet_mut(to_sheet)
        .ok_or(Error::SheetIndexOutOfRange(to_sheet))?;
    let stale: Vec<CellRef> = target
        .iter()
        .map(|(at, _)| at)
        .filter(|at| at.row >= to.row && at.col >= to.col && at.col <= end.col)
        .collect();
    for at in stale {
        target.remove(at);
    }
    for (r, row) in (0u32..).zip(cells) {
        for (c, cell) in (0u32..).zip(row) {
            let (Some(cell), Some(col), Some(rr)) = (
                cell,
                Col::new(to.col.index() + c),
                Row::new(to.row.index() + r),
            ) else {
                continue;
            };
            *target.entry(CellRef::new(col, rr)) = cell;
        }
    }
    Ok(())
}

// -------------------------------------------------------------- consolidate

/// One source of [`consolidate`]: a sheet and a range on it.
pub type ConsolidateSource = (usize, Range);

/// The summary of the numbers of one consolidated cell.
fn aggregate(how: super::SubtotalFunction, numbers: &[f64], filled: usize) -> Option<f64> {
    use super::SubtotalFunction as F;
    #[allow(clippy::cast_precision_loss, reason = "counts of cells")]
    let n = numbers.len() as f64;
    let sum: f64 = numbers.iter().sum();
    let mean = sum / n;
    let squares = || numbers.iter().map(|x| (x - mean).powi(2)).sum::<f64>();
    #[allow(clippy::cast_precision_loss, reason = "counts of cells")]
    match how {
        F::Count => (filled > 0).then_some(filled as f64),
        _ if numbers.is_empty() => None,
        F::Sum => Some(sum),
        F::CountNumbers => Some(n),
        F::Average => Some(mean),
        F::Max => numbers.iter().copied().reduce(f64::max),
        F::Min => numbers.iter().copied().reduce(f64::min),
        F::Product => Some(numbers.iter().product()),
        F::StdDev => (n > 1.0).then(|| (squares() / (n - 1.0)).sqrt()),
        F::StdDevP => Some((squares() / n).sqrt()),
        F::Var => (n > 1.0).then(|| squares() / (n - 1.0)),
        F::VarP => Some(squares() / n),
    }
}

/// Combines ranges into one table: Excel's Data - Consolidate, without links
/// to the sources.
///
/// With `top_row` and `left_column`, cells are matched by their labels (the
/// text of the range's first row and first column, without regard to case),
/// so sources need not line up; the labels appear in the order they are
/// first met. Without them, cells are matched by position and the result is
/// as large as the largest source. Each result cell is `how` over the
/// matching cells; one with no number is left empty (Count counts text too).
/// The result is written as values from `destination`. Returns the range it
/// covers.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] for no source.
#[allow(
    clippy::too_many_lines,
    reason = "labels, matching and writing in one pass"
)]
pub fn consolidate(
    book: &mut Spreadsheet,
    sources: &[ConsolidateSource],
    destination: (usize, CellRef),
    how: super::SubtotalFunction,
    top_row: bool,
    left_column: bool,
) -> Result<Range> {
    if sources.is_empty() {
        return Err(Error::InvalidRange("nothing to consolidate".into()));
    }
    let mut row_labels: Vec<String> = Vec::new();
    let mut col_labels: Vec<String> = Vec::new();
    let label_index = |labels: &mut Vec<String>, label: String| -> usize {
        labels
            .iter()
            .position(|l| l.to_lowercase() == label.to_lowercase())
            .unwrap_or_else(|| {
                labels.push(label);
                labels.len() - 1
            })
    };
    // (row key, column key) -> numbers and a count of filled cells.
    let mut cells: std::collections::HashMap<(usize, usize), (Vec<f64>, usize)> =
        std::collections::HashMap::new();
    let (mut rows, mut cols) = (0usize, 0usize);
    for &(sheet, range) in sources {
        let ws = book
            .sheet(sheet)
            .ok_or(Error::SheetIndexOutOfRange(sheet))?;
        let used = ws.dimension();
        let last_row = used.map_or_else(
            || range.start.row.index(),
            |u| range.end.row.index().min(u.end.row.index()),
        );
        let last_col = used.map_or_else(
            || range.start.col.index(),
            |u| range.end.col.index().min(u.end.col.index()),
        );
        let first_row = range.start.row.index() + u32::from(top_row);
        let first_col = range.start.col.index() + u32::from(left_column);
        let text = |c: u32, r: u32| match (Col::new(c), Row::new(r)) {
            (Some(c), Some(r)) => book.formatted(sheet, CellRef::new(c, r)),
            _ => String::new(),
        };
        for r in first_row..=last_row {
            for c in first_col..=last_col {
                let (Some(col), Some(row)) = (Col::new(c), Row::new(r)) else {
                    continue;
                };
                let rk = if left_column {
                    label_index(&mut row_labels, text(range.start.col.index(), r))
                } else {
                    (r - first_row) as usize
                };
                let ck = if top_row {
                    label_index(&mut col_labels, text(c, range.start.row.index()))
                } else {
                    (c - first_col) as usize
                };
                rows = rows.max(rk + 1);
                cols = cols.max(ck + 1);
                let entry = cells.entry((rk, ck)).or_default();
                match ws
                    .get(CellRef::new(col, row))
                    .map(|cell| cell.value.result())
                {
                    Some(CellValue::Number(n)) => {
                        entry.0.push(*n);
                        entry.1 += 1;
                    }
                    Some(v) if !v.is_empty() => entry.1 += 1,
                    _ => {}
                }
            }
        }
    }
    let (to_sheet, to) = destination;
    let off_r = u32::from(top_row);
    let off_c = u32::from(left_column);
    let height = u32::try_from(rows).unwrap_or(u32::MAX) + off_r;
    let width = u32::try_from(cols).unwrap_or(u32::MAX) + off_c;
    let end = CellRef::new(
        Col::new(to.col.index() + width.max(1) - 1).ok_or(Error::WouldPushOffSheet)?,
        Row::new(to.row.index() + height.max(1) - 1).ok_or(Error::WouldPushOffSheet)?,
    );
    let target = book
        .sheet_mut(to_sheet)
        .ok_or(Error::SheetIndexOutOfRange(to_sheet))?;
    let place = |dr: u32, dc: u32| {
        CellRef::new(
            Col::new(to.col.index() + dc).unwrap_or(to.col),
            Row::new(to.row.index() + dr).unwrap_or(to.row),
        )
    };
    for (i, label) in (0u32..).zip(&col_labels) {
        target.set(place(0, off_c + i), label.as_str());
    }
    for (i, label) in (0u32..).zip(&row_labels) {
        target.set(place(off_r + i, 0), label.as_str());
    }
    for ((rk, ck), (numbers, filled)) in &cells {
        let (Ok(r), Ok(c)) = (u32::try_from(*rk), u32::try_from(*ck)) else {
            continue;
        };
        let at = place(off_r + r, off_c + c);
        match aggregate(how, numbers, *filled) {
            Some(v) => target.set(at, v),
            None => {
                target.remove(at);
            }
        }
    }
    Ok(Range::new(to, end))
}
