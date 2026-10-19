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
