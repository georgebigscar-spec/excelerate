//! Inserting and removing rows and columns, and what that does to the rest of
//! the workbook.
//!
//!
//! A reference the edit deletes outright becomes `#REF!`, as it does in Excel;
//! one that merely loses part of a range is narrowed instead. `$` does not
//! protect a reference here: an absolute address names a cell, and inserting a
//! row above that cell moves the cell.

mod anchor;
mod chart;
mod data;
mod extension;
mod filter;
mod range;
mod series;

pub use data::{Deduplicated, SplitBy, remove_duplicates, text_to_columns};
pub use filter::filter_table;
pub use range::{
    SortBy, SortKey, SortOptions, copy_range, fill, insert_cells, insert_cells_with, move_range,
    move_sheet, remove_cells, sort_range, sort_range_with, sort_table,
};
pub use series::{fill_series, fill_series_back};

use crate::coordinate::{
    CellRef, Col, MAX_COL, MAX_ROW, Range, Row, parse_ref_at, scan_formula, scan_references,
};
use crate::error::{Error, Result};
use crate::model::sparkline::SparklineGroup;
use crate::model::{AutoFilter, CellValue, Spreadsheet, Worksheet};

/// Which way the grid is being edited.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Axis {
    /// Rows, numbered as the user sees them.
    Rows,
    /// Columns.
    Columns,
}

/// Where inserted rows or columns take their formatting from: Excel's
/// `CopyOrigin` on `Range.Insert`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CopyOrigin {
    /// No formatting: the new lines are blank.
    #[default]
    Blank,
    /// From the row above or the column to the left
    /// (`xlFormatFromLeftOrAbove`).
    Before,
    /// From the row below or the column to the right
    /// (`xlFormatFromRightOrBelow`).
    After,
}

/// Inserts `count` rows before row `at`, 0-based. Everything from there down
/// moves, and every reference to it follows.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_rows(book: &mut Spreadsheet, sheet: usize, at: Row, count: u32) -> Result<()> {
    insert_rows_with(book, sheet, at, count, CopyOrigin::Blank)
}

/// [`insert_rows`], with the new rows formatted like a neighbour: cell
/// styles, the row's own style and its height.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_rows_with(
    book: &mut Spreadsheet,
    sheet: usize,
    at: Row,
    count: u32,
    origin: CopyOrigin,
) -> Result<()> {
    insert_rows_many(book, sheet, &[(at, count)], origin)
}

/// Inserts rows at several places in one pass over the workbook: each
/// `(at, count)` puts `count` rows before row `at`, 0-based.
///
/// Every `at` is a row of the sheet as it is before the call, so the result
/// is that of [`insert_rows_with`] called once per place from the bottom up.
/// The order of `at` does not matter, and counts at the same row add up.
/// `origin` formats each inserted block from its own neighbour.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_rows_many(
    book: &mut Spreadsheet,
    sheet: usize,
    at: &[(Row, u32)],
    origin: CopyOrigin,
) -> Result<()> {
    let points: Vec<(u32, u32)> = at.iter().map(|&(r, n)| (r.index(), n)).collect();
    insert_many(
        book,
        sheet,
        &Shift::insert_many(Axis::Rows, &points),
        origin,
    )
}

/// Removes `count` rows starting at row `at`, 0-based.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet.
pub fn remove_rows(book: &mut Spreadsheet, sheet: usize, at: Row, count: u32) -> Result<()> {
    remove_rows_many(book, sheet, &[(at, count)])
}

/// Removes several blocks of rows in one pass over the workbook: each
/// `(at, count)` names `count` rows from row `at`, 0-based.
///
/// Every `at` is a row of the sheet as it is before the call, so the result
/// is that of [`remove_rows`] called once per block from the bottom up. The
/// order does not matter, and blocks that overlap or touch are one block: a
/// row is removed once however many blocks name it.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet.
pub fn remove_rows_many(book: &mut Spreadsheet, sheet: usize, spans: &[(Row, u32)]) -> Result<()> {
    let spans: Vec<(u32, u32)> = spans.iter().map(|&(r, n)| (r.index(), n)).collect();
    apply(book, sheet, &Shift::remove_many(Axis::Rows, &spans))
}

/// Inserts `count` columns before column `at`, 0-based.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_columns(book: &mut Spreadsheet, sheet: usize, at: Col, count: u32) -> Result<()> {
    insert_columns_with(book, sheet, at, count, CopyOrigin::Blank)
}

/// [`insert_columns`], with the new columns formatted like a neighbour: cell
/// styles, the column's own style and its width.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_columns_with(
    book: &mut Spreadsheet,
    sheet: usize,
    at: Col,
    count: u32,
    origin: CopyOrigin,
) -> Result<()> {
    insert_columns_many(book, sheet, &[(at, count)], origin)
}

/// [`insert_rows_many`] for columns: the result of [`insert_columns_with`]
/// called once per place from right to left.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::WouldPushOffSheet`] if a cell that holds something, or a merged
/// area, would be pushed past the last row or column; the book is then left
/// as it was. Row heights, column widths and styles at the edge are cut off
/// without a word, as Excel does.
pub fn insert_columns_many(
    book: &mut Spreadsheet,
    sheet: usize,
    at: &[(Col, u32)],
    origin: CopyOrigin,
) -> Result<()> {
    let points: Vec<(u32, u32)> = at.iter().map(|&(c, n)| (c.index(), n)).collect();
    insert_many(
        book,
        sheet,
        &Shift::insert_many(Axis::Columns, &points),
        origin,
    )
}

/// Removes `count` columns starting at column `at`, 0-based.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet.
pub fn remove_columns(book: &mut Spreadsheet, sheet: usize, at: Col, count: u32) -> Result<()> {
    remove_columns_many(book, sheet, &[(at, count)])
}

/// [`remove_rows_many`] for columns: the result of [`remove_columns`] called
/// once per block from right to left, with overlapping blocks merged.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet.
pub fn remove_columns_many(
    book: &mut Spreadsheet,
    sheet: usize,
    spans: &[(Col, u32)],
) -> Result<()> {
    let spans: Vec<(u32, u32)> = spans.iter().map(|&(c, n)| (c.index(), n)).collect();
    apply(book, sheet, &Shift::remove_many(Axis::Columns, &spans))
}

/// Applies an insertion, then formats each inserted block from its
/// neighbour. The blocks are found in the edited grid: block `k` starts where
/// its row now is, less its own count.
fn insert_many(
    book: &mut Spreadsheet,
    sheet: usize,
    shift: &Shift,
    origin: CopyOrigin,
) -> Result<()> {
    apply(book, sheet, shift)?;
    if origin == CopyOrigin::Blank {
        return Ok(());
    }
    let Some(target) = book.sheet_mut(sheet) else {
        return Ok(());
    };
    for span in &shift.spans {
        // Everything inserted above this block, its own count excluded.
        let Some(start) = u64::from(span.at)
            .checked_add(span.total - span.count())
            .and_then(|i| u32::try_from(i).ok())
        else {
            continue;
        };
        let count = u32::try_from(span.count()).unwrap_or(u32::MAX);
        format_inserted(target, shift.axis, start, count, origin);
    }
    Ok(())
}

/// Gives the `count` lines inserted at `at` the formatting of the line before
/// or after them. Values are not copied, only styles and size.
fn format_inserted(sheet: &mut Worksheet, axis: Axis, at: u32, count: u32, origin: CopyOrigin) {
    let source = match origin {
        CopyOrigin::Blank => return,
        CopyOrigin::Before => match at.checked_sub(1) {
            Some(source) => source,
            None => return,
        },
        CopyOrigin::After => match at.checked_add(count) {
            Some(source) => source,
            None => return,
        },
    };
    let last = at.saturating_add(count);
    match axis {
        Axis::Rows => {
            let Some(from) = Row::new(source) else { return };
            let styles: Vec<(Col, _)> = sheet
                .row_cells(from)
                .filter(|(_, cell)| cell.style != crate::style::StyleId::default())
                .map(|(col, cell)| (col, cell.style))
                .collect();
            let props = sheet.rows.get(&from).cloned();
            for row in (at..last).filter_map(Row::new) {
                for &(col, style) in &styles {
                    sheet.entry(CellRef::new(col, row)).style = style;
                }
                match &props {
                    Some(props) => {
                        let mut copy = props.clone();
                        copy.hidden = false;
                        copy.collapsed = false;
                        sheet.rows.insert(row, copy);
                    }
                    None => {
                        sheet.rows.remove(&row);
                    }
                }
            }
        }
        Axis::Columns => {
            let Some(from) = Col::new(source) else { return };
            // ponytail: a walk of the whole sheet per block, so a batch of
            // formatted columns costs blocks x cells; collect every source
            // column in one walk if that ever matters.
            let styles: Vec<(Row, _)> = sheet
                .iter()
                .filter(|(cell_at, cell)| {
                    cell_at.col == from && cell.style != crate::style::StyleId::default()
                })
                .map(|(cell_at, cell)| (cell_at.row, cell.style))
                .collect();
            let run = sheet.column_run(from).cloned();
            for col in (at..last).filter_map(Col::new) {
                for &(row, style) in &styles {
                    sheet.entry(CellRef::new(col, row)).style = style;
                }
                let entry = sheet.column_entry(col);
                let (first, last) = (entry.first, entry.last);
                *entry = run.clone().map_or_else(
                    || crate::model::ColumnRun::new(first, last),
                    |mut run| {
                        run.hidden = false;
                        run.collapsed = false;
                        run
                    },
                );
                entry.first = first;
                entry.last = last;
            }
        }
    }
}

/// A grid edit: rows or columns inserted at, or removed from, any number of
/// places at once, as a map from an index before the edit to the index after.
#[derive(Debug, Clone)]
struct Shift {
    axis: Axis,
    /// Whether the spans are inserted rather than removed.
    insert: bool,
    /// The places, sorted by `at`; a removal's spans are disjoint and never
    /// touch, so a run of removed indexes is always one span.
    spans: Vec<Span>,
}

/// One place of a [`Shift`], in indexes before the edit.
#[derive(Debug, Clone, Copy)]
struct Span {
    /// First index the place touches, 0-based.
    at: u32,
    /// One past the last index a removal takes; `at` plus the count.
    end: u64,
    /// The count of this span and of every span before it.
    total: u64,
}

impl Span {
    const fn count(self) -> u64 {
        self.end - self.at as u64
    }
}

impl Shift {
    #[cfg(test)]
    fn insert(axis: Axis, at: u32, count: u32) -> Self {
        Self::insert_many(axis, &[(at, count)])
    }

    #[cfg(test)]
    fn remove(axis: Axis, at: u32, count: u32) -> Self {
        Self::remove_many(axis, &[(at, count)])
    }

    /// Insertions at `points`, in any order; counts at one index add up.
    fn insert_many(axis: Axis, points: &[(u32, u32)]) -> Self {
        Self::new(axis, true, points)
    }

    /// Removals of `spans`, in any order; spans that overlap or touch merge.
    fn remove_many(axis: Axis, spans: &[(u32, u32)]) -> Self {
        Self::new(axis, false, spans)
    }

    fn new(axis: Axis, insert: bool, points: &[(u32, u32)]) -> Self {
        let mut sorted: Vec<(u32, u64)> = points
            .iter()
            .filter(|&&(_, count)| count > 0)
            .map(|&(at, count)| (at, u64::from(at) + u64::from(count)))
            .collect();
        sorted.sort_unstable();
        let mut spans: Vec<Span> = Vec::with_capacity(sorted.len());
        for (at, end) in sorted {
            match spans.last_mut() {
                // Insertions merge only at the same index; removals whenever
                // they meet.
                Some(last) if insert && last.at == at => last.end += end - u64::from(at),
                Some(last) if !insert && u64::from(at) <= last.end => last.end = last.end.max(end),
                _ => spans.push(Span { at, end, total: 0 }),
            }
        }
        let mut total = 0;
        for span in &mut spans {
            total += span.count();
            span.total = total;
        }
        Self {
            axis,
            insert,
            spans,
        }
    }

    /// The last span starting at or before `i`.
    fn span_at(&self, i: u32) -> Option<Span> {
        let k = self.spans.partition_point(|s| s.at <= i);
        k.checked_sub(1).and_then(|k| self.spans.get(k)).copied()
    }

    /// The removed span holding `i`, if the edit removes it.
    fn removing(&self, i: u32) -> Option<Span> {
        self.span_at(i)
            .filter(|s| !self.insert && u64::from(i) < s.end)
    }

    /// The largest index this axis has.
    const fn ceiling(&self) -> u32 {
        match self.axis {
            Axis::Rows => MAX_ROW - 1,
            Axis::Columns => MAX_COL - 1,
        }
    }

    /// Where index `i` ends up. `None` when the edit removes it, or when an
    /// insertion would push it off the sheet.
    fn moved(&self, i: u32) -> Option<u32> {
        let Some(span) = self.span_at(i) else {
            return Some(i);
        };
        let moved = if self.insert {
            u64::from(i).checked_add(span.total)?
        } else if u64::from(i) < span.end {
            return None;
        } else {
            u64::from(i).checked_sub(span.total)?
        };
        u32::try_from(moved).ok().filter(|&i| i <= self.ceiling())
    }

    /// Where index `i` ends up when it is the edge of a range rather than a
    /// cell: a removal that swallows it pulls it onto the edit point instead
    /// of deleting it, which is how a range shrinks around a removal.
    ///
    /// `start` says which edge it is: the leading one collapses onto the
    /// first index after the removed span, the trailing one onto the last
    /// index before it.
    fn moved_edge(&self, i: u32, start: bool) -> Option<u32> {
        let Some(span) = self.removing(i) else {
            return self.moved(i);
        };
        // Where the span's first index lands: less everything removed above.
        let at = u64::from(span.at).checked_sub(span.total - span.count())?;
        let at = u32::try_from(at).ok()?;
        if start { Some(at) } else { at.checked_sub(1) }
    }

    /// Where a range ends up, narrowed if the edit took part of it. `None`
    /// when nothing of it is left.
    fn range(&self, r: Range) -> Option<Range> {
        let (start, end) = match self.axis {
            Axis::Rows => (r.start.row.index(), r.end.row.index()),
            Axis::Columns => (r.start.col.index(), r.end.col.index()),
        };
        // A removal that covers the whole range leaves nothing to narrow;
        // removed indexes in a row are always one span.
        if let Some(span) = self.removing(start)
            && u64::from(end) < span.end
        {
            return None;
        }
        let (start, end) = (self.moved_edge(start, true)?, self.moved_edge(end, false)?);
        if start > end {
            return None;
        }
        Some(match self.axis {
            Axis::Rows => Range::new(
                CellRef::new(r.start.col, Row::new(start)?),
                CellRef::new(r.end.col, Row::new(end)?),
            ),
            Axis::Columns => Range::new(
                CellRef::new(Col::new(start)?, r.start.row),
                CellRef::new(Col::new(end)?, r.end.row),
            ),
        })
    }

    /// Whether an insertion would push content off the sheet: a stored cell
    /// (a blank one that only carries a style counts, as it does to Excel) or
    /// a merged area, past the last row or column. Excel refuses such an edit
    /// and so does this; what it cuts off without a word - row heights and
    /// styles, column runs, drawings, which stay pinned to the edge - is not
    /// asked about.
    ///
    /// Indexes only grow under an insertion, so the last occupied one decides
    /// for the cells: no walk of them.
    fn pushes_off(&self, sheet: &Worksheet) -> bool {
        if !self.insert {
            return false;
        }
        let gone = |i: u32| self.moved(i).is_none();
        let last = sheet.dimension().map(|d| match self.axis {
            Axis::Rows => d.end.row.index(),
            Axis::Columns => d.end.col.index(),
        });
        last.is_some_and(gone)
            || sheet.merges.iter().any(|m| {
                gone(match self.axis {
                    Axis::Rows => m.end.row.index(),
                    Axis::Columns => m.end.col.index(),
                })
            })
    }

    /// Where a cell ends up; `None` when the edit removes it.
    fn cell(&self, at: CellRef) -> Option<CellRef> {
        Some(match self.axis {
            Axis::Rows => CellRef::new(at.col, Row::new(self.moved(at.row.index())?)?),
            Axis::Columns => CellRef::new(Col::new(self.moved(at.col.index())?)?, at.row),
        })
    }
}

/// Applies a grid edit to the whole workbook.
fn apply(book: &mut Spreadsheet, sheet: usize, shift: &Shift) -> Result<()> {
    let target = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    if shift.pushes_off(target) {
        return Err(Error::WouldPushOffSheet);
    }
    let title = target.title().to_owned();

    // Formulas first, while the cells still sit where the formulas expect
    // them: every sheet may point at the edited one.
    let names: Vec<String> = book.sheets().iter().map(|s| s.title().to_owned()).collect();
    for (index, own) in names.iter().enumerate() {
        let Some(target) = book.sheet_mut(index) else {
            continue;
        };
        let rewritten: Vec<(CellRef, String)> = target
            .iter()
            .filter_map(|(at, cell)| match &cell.value {
                CellValue::Formula { formula, .. } => {
                    let text = adjust(formula, shift, &title, own == &title);
                    (&text != formula).then_some((at, text))
                }
                _ => None,
            })
            .collect();
        for (at, formula) in rewritten {
            // The cached value belonged to the formula as it was written.
            target.entry(at).value = CellValue::Formula {
                formula,
                cached: None,
            };
        }
    }
    for name in &mut book.defined_names {
        let own = name.sheet.and_then(|i| names.get(i)) == Some(&title);
        name.formula = adjust(&name.formula, shift, &title, own);
    }

    // The drawings are carried, not modelled, so their anchors are rewritten
    // inside the bytes; this reads the sheet, so it runs before the borrow.
    anchor::move_anchors(book, sheet, shift);
    // A chart keeps its series as formulas, and they name their sheet.
    chart::rewrite_series(book, |series| adjust(series, shift, &title, false));
    chart::rewrite_model(book, |series| adjust(series, shift, &title, false));
    rewrite_sparklines(book, |formula, on| {
        adjust(formula, shift, &title, on == title)
    });
    // The carried extensions name cells too: their formulas on any sheet,
    // their ranges on this one.
    let range = |r: Range| shift.range(r);
    for index in 0..book.sheets().len() {
        let Some(target) = book.sheet_mut(index) else {
            continue;
        };
        let own = target.title().eq_ignore_ascii_case(&title);
        if let Some(ext) = target.extensions.take() {
            let moved: Option<&dyn Fn(Range) -> Option<Range>> = own.then_some(&range);
            target.extensions = Some(extension::rewrite(
                &ext,
                |f| adjust(f, shift, &title, own),
                moved,
            ));
        }
    }

    let Some(target) = book.sheet_mut(sheet) else {
        return Ok(());
    };
    move_cells(target, shift);
    move_furniture(target, shift);
    Ok(())
}

/// Moves the cells themselves, dropping what the edit removed.
fn move_cells(sheet: &mut Worksheet, shift: &Shift) {
    let moved: Vec<(CellRef, Option<CellRef>)> =
        sheet.iter().map(|(at, _)| (at, shift.cell(at))).collect();
    // Taken out first, so a cell never lands on one not yet moved.
    let mut carried = Vec::with_capacity(moved.len());
    for (from, to) in moved {
        if let Some(cell) = sheet.remove(from)
            && let Some(to) = to
        {
            carried.push((to, cell));
        }
    }
    for (at, cell) in carried {
        *sheet.entry(at) = cell;
    }
}

/// A filter after the edit, with the sort it carries; `None` when its range
/// is gone.
fn moved_filter(mut filter: AutoFilter, shift: &Shift) -> Option<AutoFilter> {
    filter.range = shift.range(filter.range)?;
    filter.sort_state = filter
        .sort_state
        .take()
        .and_then(|sort| sort.moved(|r| shift.range(r)));
    Some(filter)
}

/// A sparkline is drawn in a cell and goes where the cell goes; one whose
/// cell the edit removed goes with it, and so does a group left empty.
fn move_sparklines(sheet: &mut Worksheet, shift: &Shift) {
    sheet.sparklines.retain_mut(|group| {
        group.sparklines.retain_mut(|line| {
            let at = line.location;
            let Some(moved) = shift.cell(at) else {
                return false;
            };
            line.location = moved;
            true
        });
        !group.sparklines.is_empty()
    });
}

/// Moves everything on the sheet that names a row, a column or a range.
fn move_furniture(sheet: &mut Worksheet, shift: &Shift) {
    sheet.merges = sheet
        .merges
        .iter()
        .filter_map(|&r| shift.range(r))
        .collect();
    sheet.array_formulas = sheet
        .array_formulas
        .iter()
        .filter_map(|&r| shift.range(r))
        .collect();
    sheet
        .hyperlinks
        .retain_mut(|link| match shift.range(link.range) {
            Some(range) => {
                link.range = range;
                true
            }
            None => false,
        });
    sheet.data_validations.retain_mut(|v| {
        v.sqref = v.sqref.iter().filter_map(|&r| shift.range(r)).collect();
        !v.sqref.is_empty()
    });
    sheet.conditional_formats.retain_mut(|f| {
        f.sqref = f.sqref.iter().filter_map(|&r| shift.range(r)).collect();
        !f.sqref.is_empty()
    });
    sheet.protected_ranges.retain_mut(|p| {
        p.sqref = p.sqref.iter().filter_map(|&r| shift.range(r)).collect();
        !p.sqref.is_empty()
    });
    sheet.auto_filter = sheet
        .auto_filter
        .take()
        .and_then(|f| moved_filter(f, shift));
    sheet.sort_state = sheet
        .sort_state
        .take()
        .and_then(|sort| sort.moved(|r| shift.range(r)));
    move_sparklines(sheet, shift);
    // A note is attached to a cell, so it goes where the cell goes; one whose
    // cell the edit removed goes with it. Where its box is drawn is the VML
    // part, which `anchor` moves in the bytes.
    sheet.comments = std::mem::take(&mut sheet.comments)
        .into_iter()
        .filter_map(|(at, note)| Some((shift.cell(at)?, note)))
        .collect();
    // A table whose range is gone entirely goes with it, the way a filter
    // does; one that lost part of itself narrows. Its own filter sits on the
    // header row and follows the same rule.
    sheet.tables.retain_mut(|t| {
        let Some(range) = shift.range(t.range) else {
            return false;
        };
        t.range = range;
        t.auto_filter = t.auto_filter.take().and_then(|f| moved_filter(f, shift));
        t.sort_state = t
            .sort_state
            .take()
            .and_then(|sort| sort.moved(|r| shift.range(r)));
        true
    });
    move_drawn_objects(sheet, shift);

    // Row and column properties are indexed by the axis they sit on, so only
    // the matching ones move; the others name the axis that did not change.
    match shift.axis {
        Axis::Rows => {
            sheet.rows = std::mem::take(&mut sheet.rows)
                .into_iter()
                .filter_map(|(row, props)| Some((Row::new(shift.moved(row.index())?)?, props)))
                .collect();
            sheet.row_breaks.retain_mut(|b| match shift.moved(b.at) {
                Some(at) => {
                    b.at = at;
                    true
                }
                None => false,
            });
        }
        Axis::Columns => {
            sheet.columns.retain_mut(|run| {
                match (
                    shift.moved_edge(run.first.index(), true),
                    shift.moved_edge(run.last.index(), false),
                ) {
                    (Some(first), Some(last)) if first <= last => {
                        let (Some(first), Some(last)) = (Col::new(first), Col::new(last)) else {
                            return false;
                        };
                        run.first = first;
                        run.last = last;
                        true
                    }
                    _ => false,
                }
            });
            sheet.col_breaks.retain_mut(|b| match shift.moved(b.at) {
                Some(at) => {
                    b.at = at;
                    true
                }
                None => false,
            });
        }
    }
}

/// Moves the charts, pictures and shapes of the model with the grid.
fn move_drawn_objects(sheet: &mut Worksheet, shift: &Shift) {
    // A chart's frame is anchored to cells like any drawing. The bytes of the
    // drawing were moved by `anchor`, so an untouched chart stays untouched.
    for chart in &mut sheet.charts {
        let untouched = chart.is_unchanged();
        anchor::move_anchor(&mut chart.anchor, shift);
        if untouched {
            chart.settle();
        }
    }
    for chart in &mut sheet.extended_charts {
        let untouched = chart.is_unchanged();
        anchor::move_anchor(&mut chart.anchor, shift);
        if untouched {
            chart.settle();
        }
    }
    // Pictures the same way: their drawing moved with the charts'.
    for image in &mut sheet.images {
        let untouched = image.is_unchanged();
        anchor::move_anchor(&mut image.anchor, shift);
        if untouched {
            image.settle();
        }
    }
    for shape in &mut sheet.shapes {
        let untouched = shape.is_unchanged();
        anchor::move_anchor(&mut shape.anchor, shift);
        if untouched {
            shape.settle();
        }
    }
}

/// Rewrites the references of one formula for a grid edit on `target`.
///
/// `own` says whether the formula lives on the edited sheet, which is what an
/// unqualified reference points at.
fn adjust(formula: &str, shift: &Shift, target: &str, own: bool) -> String {
    scan_references(formula, |qualifier, s| {
        let aims_at_target = qualifier.map_or(own, |q| q.eq_ignore_ascii_case(target));
        if !aims_at_target {
            return None;
        }
        let (len, first) = parse_ref_at(s)?;
        // `A1:B2` is one reference: an edit inside it narrows it, where the
        // same edit over a lone `A1` deletes it outright.
        if s.get(len) == Some(&':')
            && let Some((second_len, second)) = parse_ref_at(&s[len + 1..])
        {
            let range = Range::new(
                CellRef::new(first.col, first.row),
                CellRef::new(second.col, second.row),
            );
            let text = shift.range(range).map_or_else(
                || "#REF!".to_owned(),
                |moved| {
                    format!(
                        "{}:{}",
                        first.render(moved.start.col, moved.start.row),
                        second.render(moved.end.col, moved.end.row)
                    )
                },
            );
            return Some((len + 1 + second_len, text));
        }
        let at = CellRef::new(first.col, first.row);
        let text = shift.cell(at).map_or_else(
            || "#REF!".to_owned(),
            |moved| first.render(moved.col, moved.row),
        );
        Some((len, text))
    })
}

/// Renames the sheet at `sheet`, rewriting every formula that names it.
///
/// A sheet name lives in two places: the tab, and the qualifier of every
/// reference pointing at the sheet from anywhere in the workbook. Changing
/// only the first leaves those references naming a sheet that no longer
/// exists, which is why this is not `Worksheet::set_title`.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet,
/// [`Error::DuplicateSheetName`] if another sheet already has the name, and
/// whatever [`crate::model::Worksheet::set_title`] rejects the name with.
pub fn rename_sheet(book: &mut Spreadsheet, sheet: usize, to: &str) -> Result<()> {
    let from = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?
        .title()
        .to_owned();
    if from == to {
        return Ok(());
    }
    if let Some(other) = book.sheet_by_name(to)
        && other.title() != from
    {
        return Err(Error::DuplicateSheetName(to.to_owned()));
    }
    // The title first: it validates the name, and a rejected name must leave
    // the formulas alone.
    book.sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?
        .set_title(to)?;
    rewrite_qualifiers(book, |names| {
        if !names.iter().any(|n| n.eq_ignore_ascii_case(&from)) {
            return None;
        }
        Some(render_qualifier(names, |n| {
            if n.eq_ignore_ascii_case(&from) {
                to.to_owned()
            } else {
                n.clone()
            }
        }))
    });
    Ok(())
}

/// Removes the sheet at `sheet`, turning every reference into it into `#REF!`.
///
/// A 3-D span keeps working where it can: `Sheet1:Sheet3!A1` with `Sheet3`
/// gone becomes `Sheet1:Sheet2!A1` when a sheet remains between the ends, and
/// only collapses to `#REF!` when nothing of the span is left.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] if there is no such sheet, and
/// [`Error::Xlsx`] when it is the only one: a workbook with no sheets cannot
/// be written, and Excel refuses the same edit.
pub fn remove_sheet(book: &mut Spreadsheet, sheet: usize) -> Result<()> {
    let order: Vec<String> = book.sheets().iter().map(|s| s.title().to_owned()).collect();
    let gone = order
        .get(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?
        .clone();
    if order.len() == 1 {
        return Err(Error::Xlsx(
            "a workbook must keep at least one sheet".into(),
        ));
    }
    let survivors: Vec<String> = order
        .iter()
        .enumerate()
        .filter(|(i, _)| *i != sheet)
        .map(|(_, n)| n.clone())
        .collect();

    rewrite_qualifiers(book, |names| match names {
        [one] => one.eq_ignore_ascii_case(&gone).then(|| "#REF!".to_owned()),
        [first, last] => {
            let ends = [
                nearest_inside(&order, &survivors, first, last),
                nearest_inside(&order, &survivors, last, first),
            ];
            match ends {
                [Some(a), Some(b)] => {
                    (a != *first || b != *last).then(|| render_qualifier(&[a, b], Clone::clone))
                }
                // Nothing of the span is left to point at.
                _ => Some("#REF!".to_owned()),
            }
        }
        _ => None,
    });

    // A name that belonged to the sheet goes with it; the rest keep pointing
    // at their sheet, whose tab index has moved down by one.
    book.defined_names.retain(|n| n.sheet != Some(sheet));
    for name in &mut book.defined_names {
        if let Some(index) = name.sheet
            && index > sheet
        {
            name.sheet = Some(index - 1);
        }
    }
    book.take_sheet(sheet);
    Ok(())
}

/// The endpoint a 3-D span should use once a sheet is gone: itself when it
/// survived, otherwise the nearest surviving sheet inside the span, walking
/// from that end towards the other.
///
/// `None` when the whole span went, which is the only case that becomes
/// `#REF!`.
fn nearest_inside(
    order: &[String],
    survivors: &[String],
    end: &str,
    other: &str,
) -> Option<String> {
    let alive = |name: &str| survivors.iter().any(|s| s.eq_ignore_ascii_case(name));
    let position = |name: &str| order.iter().position(|s| s.eq_ignore_ascii_case(name));
    let (from, to) = (position(end)?, position(other)?);
    let mut inwards: Box<dyn Iterator<Item = &String>> = if from <= to {
        Box::new(order[from..=to].iter())
    } else {
        Box::new(order[to..=from].iter().rev())
    };
    inwards.find(|n| alive(n)).cloned()
}

/// Writes a qualifier back out, mapping each name through `f`, quoting what
/// has to be quoted and ending with the `!` the scanner expects.
fn render_qualifier(names: &[String], f: impl Fn(&String) -> String) -> String {
    let parts: Vec<String> = names.iter().map(|n| quote_sheet_name(&f(n))).collect();
    format!("{}!", parts.join(":"))
}

/// A sheet name as a formula spells it: bare when it can be, quoted when the
/// name holds anything else, with inner apostrophes doubled.
fn quote_sheet_name(name: &str) -> String {
    let bare = !name.is_empty()
        && name
            .chars()
            .next()
            .is_some_and(|c| c.is_alphabetic() || c == '_')
        && name
            .chars()
            .all(|c| c.is_alphanumeric() || matches!(c, '_' | '.'));
    if bare {
        name.to_owned()
    } else {
        format!("'{}'", name.replace('\'', "''"))
    }
}

/// Runs `rename` over every formula in the workbook: the cells of every sheet
/// and the defined names, which are formulas too.
fn rewrite_qualifiers(book: &mut Spreadsheet, rename: impl Fn(&[String]) -> Option<String>) {
    for index in 0..book.sheets().len() {
        let Some(target) = book.sheet_mut(index) else {
            continue;
        };
        let rewritten: Vec<(CellRef, String)> = target
            .iter()
            .filter_map(|(at, cell)| match &cell.value {
                CellValue::Formula { formula, .. } => {
                    let text = scan_formula(formula, |_, _| None, &rename);
                    (&text != formula).then_some((at, text))
                }
                _ => None,
            })
            .collect();
        for (at, formula) in rewritten {
            // The cached value answered the formula as it was written.
            target.entry(at).value = CellValue::Formula {
                formula,
                cached: None,
            };
        }
    }
    for name in &mut book.defined_names {
        name.formula = scan_formula(&name.formula, |_, _| None, &rename);
    }
    // A chart series is a formula too, and it always names its sheet.
    chart::rewrite_series(book, |series| scan_formula(series, |_, _| None, &rename));
    chart::rewrite_model(book, |series| scan_formula(series, |_, _| None, &rename));
    rewrite_sparklines(book, |formula, _| {
        scan_formula(formula, |_, _| None, &rename)
    });
    for index in 0..book.sheets().len() {
        if let Some(target) = book.sheet_mut(index)
            && let Some(ext) = target.extensions.take()
        {
            target.extensions = Some(extension::rewrite(
                &ext,
                |f| scan_formula(f, |_, _| None, &rename),
                None,
            ));
        }
    }
}

/// Runs `rewrite` over what every sparkline of the workbook reads, with the
/// title of the sheet it is drawn on. The data may sit on any sheet.
fn rewrite_sparklines(book: &mut Spreadsheet, rewrite: impl Fn(&str, &str) -> String) {
    for index in 0..book.sheets().len() {
        let Some(sheet) = book.sheet_mut(index) else {
            continue;
        };
        let title = sheet.title().to_owned();
        for formula in sheet
            .sparklines
            .iter_mut()
            .flat_map(SparklineGroup::formulas_mut)
        {
            *formula = rewrite(formula, &title);
        }
    }
}
