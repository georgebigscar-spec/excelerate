//! The outline of a sheet: Data - Group and Ungroup, Hide and Show Detail,
//! the level buttons, and Data - Subtotal, which builds an outline of its own.
//!
//! A group is a run of rows (or columns) whose outline level is at least its
//! own; its summary line is the one just below it (right of it for columns),
//! or just above (left) when the sheet's `summary_below` (`summary_right`)
//! is off. That is where Excel draws the plus and minus buttons, and where
//! it marks a collapsed group.
//!
//! Ported from the outline commands of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use crate::coordinate::{CellRef, Col, Range, Row};
use crate::error::{Error, Result};
use crate::model::{CellValue, Spreadsheet, Worksheet};

use super::CopyOrigin;

/// The deepest outline Excel keeps: eight level buttons, seven levels of
/// groups.
pub const MAX_OUTLINE_LEVEL: u8 = 7;

/// One group of an outline: lines `first..=last`, 0-based, at `level`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OutlineGroup {
    /// The depth, from 1.
    pub level: u8,
    /// First line, 0-based.
    pub first: u32,
    /// Last line, 0-based.
    pub last: u32,
}

impl OutlineGroup {
    /// Whether `line` is in the group.
    #[must_use]
    pub const fn contains(&self, line: u32) -> bool {
        self.first <= line && line <= self.last
    }

    /// The summary line, given which side summaries sit on. `None` for a
    /// group against the sheet's edge, which has none.
    #[must_use]
    pub fn summary(&self, after: bool) -> Option<u32> {
        if after {
            Some(self.last + 1).filter(|&n| n <= MAX_LINE)
        } else {
            self.first.checked_sub(1)
        }
    }
}

/// Rows or columns: the two outlines of a sheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Lines {
    Rows,
    Columns,
}

/// The last 0-based line either way; rows run further, and columns are
/// checked against their own bound where they are made.
const MAX_LINE: u32 = crate::coordinate::MAX_ROW - 1;

impl Lines {
    fn level(self, ws: &Worksheet, line: u32) -> u8 {
        match self {
            Self::Rows => Row::new(line).map_or(0, |r| ws.row_outline_level(r)),
            Self::Columns => Col::new(line).map_or(0, |c| ws.column_outline_level(c)),
        }
    }

    /// Every line with a level, ascending, as `(line, level)`.
    fn levelled(self, ws: &Worksheet) -> Vec<(u32, u8)> {
        match self {
            Self::Rows => ws
                .rows
                .iter()
                .filter(|(_, p)| p.outline_level > 0)
                .map(|(r, p)| (r.index(), p.outline_level))
                .collect(),
            Self::Columns => {
                let mut out: Vec<(u32, u8)> = ws
                    .columns
                    .iter()
                    .filter(|run| run.outline_level > 0)
                    .flat_map(|run| {
                        (run.first.index()..=run.last.index()).map(|c| (c, run.outline_level))
                    })
                    .collect();
                out.sort_unstable();
                out
            }
        }
    }

    fn summary_after(self, ws: &Worksheet) -> bool {
        match self {
            Self::Rows => ws.properties.summary_below,
            Self::Columns => ws.properties.summary_right,
        }
    }

    fn update(
        self,
        ws: &mut Worksheet,
        line: u32,
        change: impl FnOnce(&mut u8, &mut bool, &mut bool),
    ) {
        match self {
            Self::Rows => {
                let Some(row) = Row::new(line) else { return };
                let p = ws.rows.entry(row).or_default();
                change(&mut p.outline_level, &mut p.hidden, &mut p.collapsed);
                if !p.is_meaningful() {
                    ws.rows.remove(&row);
                }
            }
            Self::Columns => {
                let Some(col) = Col::new(line) else { return };
                let run = ws.column_entry(col);
                change(&mut run.outline_level, &mut run.hidden, &mut run.collapsed);
            }
        }
    }

    fn check(self, first: u32, last: u32) -> Result<()> {
        let fits = match self {
            Self::Rows => Row::new(last).is_some(),
            Self::Columns => Col::new(last).is_some(),
        };
        if first > last || !fits {
            return Err(Error::InvalidRange(format!(
                "lines {} to {} are not on the sheet",
                first + 1,
                last + 1
            )));
        }
        Ok(())
    }
}

/// The groups of one outline, inner ones after the group around them.
fn groups_of(levelled: &[(u32, u8)]) -> Vec<OutlineGroup> {
    let deepest = levelled.iter().map(|&(_, l)| l).max().unwrap_or(0);
    let mut out = Vec::new();
    for level in 1..=deepest {
        let mut open: Option<(u32, u32)> = None;
        for &(line, l) in levelled {
            if l < level {
                if let Some((first, last)) = open.take() {
                    out.push(OutlineGroup { level, first, last });
                }
                continue;
            }
            open = match open {
                Some((first, last)) if last + 1 == line => Some((first, line)),
                Some((first, last)) => {
                    out.push(OutlineGroup { level, first, last });
                    Some((line, line))
                }
                None => Some((line, line)),
            };
        }
        if let Some((first, last)) = open {
            out.push(OutlineGroup { level, first, last });
        }
    }
    out
}

/// The groups of a sheet's rows, outer ones first.
#[must_use]
pub fn row_groups(ws: &Worksheet) -> Vec<OutlineGroup> {
    groups_of(&Lines::Rows.levelled(ws))
}

/// The groups of a sheet's columns, outer ones first.
#[must_use]
pub fn column_groups(ws: &Worksheet) -> Vec<OutlineGroup> {
    groups_of(&Lines::Columns.levelled(ws))
}

fn sheet_mut(book: &mut Spreadsheet, sheet: usize) -> Result<&mut Worksheet> {
    book.sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))
}

fn group(book: &mut Spreadsheet, sheet: usize, lines: Lines, first: u32, last: u32) -> Result<()> {
    lines.check(first, last)?;
    let ws = sheet_mut(book, sheet)?;
    if let Some(line) = (first..=last).find(|&l| lines.level(ws, l) >= MAX_OUTLINE_LEVEL) {
        return Err(Error::InvalidRange(format!(
            "line {} is already {MAX_OUTLINE_LEVEL} levels deep, as deep as an outline goes",
            line + 1
        )));
    }
    for line in first..=last {
        lines.update(ws, line, |level, _, _| *level += 1);
    }
    Ok(())
}

fn ungroup(
    book: &mut Spreadsheet,
    sheet: usize,
    lines: Lines,
    first: u32,
    last: u32,
) -> Result<()> {
    lines.check(first, last)?;
    let ws = sheet_mut(book, sheet)?;
    let after = lines.summary_after(ws);
    let before = groups_of(&lines.levelled(ws));
    let touched: Vec<u32> = lines
        .levelled(ws)
        .into_iter()
        .filter(|&(l, _)| first <= l && l <= last)
        .map(|(l, _)| l)
        .collect();
    for line in touched {
        lines.update(ws, line, |level, hidden, _| {
            *level -= 1;
            // Out of every group, nothing can show it again.
            if *level == 0 {
                *hidden = false;
            }
        });
    }
    // A summary line whose group is gone has nothing left to collapse.
    let after_groups = groups_of(&lines.levelled(ws));
    for g in before {
        let Some(summary) = g.summary(after) else {
            continue;
        };
        if !after_groups
            .iter()
            .any(|h| h.summary(after) == Some(summary))
        {
            lines.update(ws, summary, |_, _, collapsed| *collapsed = false);
        }
    }
    Ok(())
}

/// The group a line names: the one it is the summary line of, as the
/// outline's plus and minus buttons do, or else the innermost one it is in.
fn group_named(groups: &[OutlineGroup], after: bool, line: u32) -> Option<OutlineGroup> {
    groups
        .iter()
        .filter(|g| g.summary(after) == Some(line))
        .max_by_key(|g| g.level)
        .or_else(|| {
            groups
                .iter()
                .filter(|g| g.contains(line))
                .max_by_key(|g| g.level)
        })
        .copied()
}

fn set_detail(
    book: &mut Spreadsheet,
    sheet: usize,
    lines: Lines,
    line: u32,
    show: bool,
) -> Result<bool> {
    let ws = sheet_mut(book, sheet)?;
    let after = lines.summary_after(ws);
    let groups = groups_of(&lines.levelled(ws));
    let Some(g) = group_named(&groups, after, line) else {
        return Ok(false);
    };
    // Inner groups collapsed on their own stay collapsed when this one opens.
    let collapsed_inside: Vec<OutlineGroup> = groups
        .iter()
        .filter(|o| o.level > g.level && g.first <= o.first && o.last <= g.last)
        .filter(|o| {
            o.summary(after)
                .filter(|&s| s != g.summary(after).unwrap_or(u32::MAX))
                .is_some_and(|s| is_collapsed(ws, lines, s))
        })
        .copied()
        .collect();
    for l in g.first..=g.last {
        let hide = !show || collapsed_inside.iter().any(|o| o.contains(l));
        lines.update(ws, l, |_, hidden, _| *hidden = hide);
    }
    if let Some(summary) = g.summary(after) {
        lines.update(ws, summary, |_, _, collapsed| *collapsed = !show);
    }
    Ok(true)
}

fn is_collapsed(ws: &Worksheet, lines: Lines, line: u32) -> bool {
    match lines {
        Lines::Rows => Row::new(line)
            .and_then(|r| ws.rows.get(&r))
            .is_some_and(|p| p.collapsed),
        Lines::Columns => Col::new(line)
            .and_then(|c| ws.column_run(c))
            .is_some_and(|run| run.collapsed),
    }
}

fn show_level(ws: &mut Worksheet, lines: Lines, level: u8) {
    let after = lines.summary_after(ws);
    for (line, l) in lines.levelled(ws) {
        lines.update(ws, line, |_, hidden, _| *hidden = l >= level);
    }
    // Inner groups first, so a summary line two groups share ends as the
    // outer one says.
    let mut groups = groups_of(&lines.levelled(ws));
    groups.sort_by_key(|g| std::cmp::Reverse(g.level));
    for g in groups {
        if let Some(summary) = g.summary(after) {
            lines.update(ws, summary, |_, _, collapsed| *collapsed = g.level >= level);
        }
    }
}

/// Groups rows `first..=last` one level deeper: Data - Group. Rows already
/// grouped become a group inside their group.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] if `first` is below `last` or a row is already
/// [`MAX_OUTLINE_LEVEL`] deep; nothing changes then.
pub fn group_rows(book: &mut Spreadsheet, sheet: usize, first: Row, last: Row) -> Result<()> {
    group(book, sheet, Lines::Rows, first.index(), last.index())
}

/// Takes rows `first..=last` one level out of their outline: Data -
/// Ungroup. A row that leaves the outline altogether is shown again.
///
/// # Errors
/// As [`group_rows`], save the depth.
pub fn ungroup_rows(book: &mut Spreadsheet, sheet: usize, first: Row, last: Row) -> Result<()> {
    ungroup(book, sheet, Lines::Rows, first.index(), last.index())
}

/// [`group_rows`] for columns.
///
/// # Errors
/// As [`group_rows`].
pub fn group_columns(book: &mut Spreadsheet, sheet: usize, first: Col, last: Col) -> Result<()> {
    group(book, sheet, Lines::Columns, first.index(), last.index())
}

/// [`ungroup_rows`] for columns.
///
/// # Errors
/// As [`group_rows`].
pub fn ungroup_columns(book: &mut Spreadsheet, sheet: usize, first: Col, last: Col) -> Result<()> {
    ungroup(book, sheet, Lines::Columns, first.index(), last.index())
}

/// Hides the detail of a group: Data - Hide Detail, or the outline's minus
/// button. `row` names the group as the button does: the group whose
/// summary row it is, or else the innermost group it is in. Returns whether
/// there was such a group.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet.
pub fn collapse_rows(book: &mut Spreadsheet, sheet: usize, row: Row) -> Result<bool> {
    set_detail(book, sheet, Lines::Rows, row.index(), false)
}

/// Shows the detail of a group again: Data - Show Detail, or the plus
/// button. Groups inside it that were collapsed on their own stay so.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet.
pub fn expand_rows(book: &mut Spreadsheet, sheet: usize, row: Row) -> Result<bool> {
    set_detail(book, sheet, Lines::Rows, row.index(), true)
}

/// [`collapse_rows`] for columns.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet.
pub fn collapse_columns(book: &mut Spreadsheet, sheet: usize, col: Col) -> Result<bool> {
    set_detail(book, sheet, Lines::Columns, col.index(), false)
}

/// [`expand_rows`] for columns.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet.
pub fn expand_columns(book: &mut Spreadsheet, sheet: usize, col: Col) -> Result<bool> {
    set_detail(book, sheet, Lines::Columns, col.index(), true)
}

/// The outline's level buttons: shows rows, columns or both down to outline
/// level `n`, where 1 shows only the outermost summaries and
/// `MAX_OUTLINE_LEVEL + 1` shows everything. `None` leaves that side alone.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] for a level outside `1..=8`.
pub fn show_outline_level(
    book: &mut Spreadsheet,
    sheet: usize,
    rows: Option<u8>,
    columns: Option<u8>,
) -> Result<()> {
    for n in [rows, columns].into_iter().flatten() {
        if !(1..=MAX_OUTLINE_LEVEL + 1).contains(&n) {
            return Err(Error::InvalidRange(format!(
                "an outline level runs from 1 to {}, not {n}",
                MAX_OUTLINE_LEVEL + 1
            )));
        }
    }
    let ws = sheet_mut(book, sheet)?;
    if let Some(n) = rows {
        show_level(ws, Lines::Rows, n);
    }
    if let Some(n) = columns {
        show_level(ws, Lines::Columns, n);
    }
    Ok(())
}

/// What a [`subtotal`] row computes: the eleven aggregates of `SUBTOTAL`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SubtotalFunction {
    /// `SUBTOTAL(9, ...)`, labelled "Total".
    #[default]
    Sum,
    /// `SUBTOTAL(3, ...)`: the cells that are not empty.
    Count,
    /// `SUBTOTAL(1, ...)`.
    Average,
    /// `SUBTOTAL(4, ...)`.
    Max,
    /// `SUBTOTAL(5, ...)`.
    Min,
    /// `SUBTOTAL(6, ...)`.
    Product,
    /// `SUBTOTAL(2, ...)`: the numbers only.
    CountNumbers,
    /// `SUBTOTAL(7, ...)`, of a sample.
    StdDev,
    /// `SUBTOTAL(8, ...)`, of a population.
    StdDevP,
    /// `SUBTOTAL(10, ...)`, of a sample.
    Var,
    /// `SUBTOTAL(11, ...)`, of a population.
    VarP,
}

impl SubtotalFunction {
    /// The function number `SUBTOTAL` takes.
    #[must_use]
    pub const fn code(self) -> u8 {
        match self {
            Self::Average => 1,
            Self::CountNumbers => 2,
            Self::Count => 3,
            Self::Max => 4,
            Self::Min => 5,
            Self::Product => 6,
            Self::StdDev => 7,
            Self::StdDevP => 8,
            Self::Sum => 9,
            Self::Var => 10,
            Self::VarP => 11,
        }
    }

    /// The word Excel's English labels use: "Ann Total", "Grand Total".
    #[must_use]
    pub const fn label(self) -> &'static str {
        match self {
            Self::Sum => "Total",
            Self::Count | Self::CountNumbers => "Count",
            Self::Average => "Average",
            Self::Max => "Max",
            Self::Min => "Min",
            Self::Product => "Product",
            Self::StdDev => "StdDev",
            Self::StdDevP => "StdDevp",
            Self::Var => "Var",
            Self::VarP => "Varp",
        }
    }
}

/// How [`subtotal`] summarises a list.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Subtotals {
    /// "At each change in": the column whose value starts a new group.
    pub group_by: Col,
    /// "Use function".
    pub function: SubtotalFunction,
    /// "Add subtotal to": the columns totalled.
    pub columns: Vec<Col>,
}

/// Adds a total row under each group of a sorted list and a grand total
/// under all of them, and outlines the result: Excel's Data - Subtotal with
/// "Summary below data".
///
/// `area` is the list with its header row on top. A group is a run of rows
/// whose `group_by` cells show the same text (without regard to case), so
/// sort the list first. Each total row is a whole inserted row, as in Excel,
/// so formulas across the workbook move with the cells they read. Its label,
/// "<name> Total" (or Count, Average... as [`SubtotalFunction::label`] says),
/// goes in the `group_by` column in bold; each total is a `SUBTOTAL` formula,
/// which a grand total over it does not count twice. Recalculate for their
/// values. The outline gets two levels: the detail rows of each group, and
/// every row down to the last group's total.
///
/// Returns how many groups were found.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] if `area` has no data row under its header, a
/// column is outside it, or the rows are already too deeply grouped to take
/// two more levels. Nothing changes then.
pub fn subtotal(book: &mut Spreadsheet, sheet: usize, area: Range, how: &Subtotals) -> Result<u32> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let inside = |c: Col| area.start.col <= c && c <= area.end.col;
    if let Some(c) = std::iter::once(&how.group_by)
        .chain(&how.columns)
        .find(|c| !inside(**c))
    {
        return Err(Error::InvalidRange(format!(
            "column {} is outside {area}",
            c.to_letters()
        )));
    }
    if how.columns.is_empty() {
        return Err(Error::InvalidRange("no column to add a subtotal to".into()));
    }
    let first = area.start.row.index() + 1;
    // A whole column ends with its data, not a million rows down.
    let last = ws
        .dimension()
        .map_or(first, |used| area.end.row.index().min(used.end.row.index()));
    if first > last {
        return Err(Error::InvalidRange(format!(
            "{area} has no data under its header"
        )));
    }
    if (first..=last + 1).any(|r| Lines::Rows.level(ws, r) + 2 > MAX_OUTLINE_LEVEL) {
        return Err(Error::InvalidRange(format!(
            "the rows of {area} are grouped too deeply to take two more levels"
        )));
    }
    if last + (last - first + 2) > MAX_LINE {
        return Err(Error::WouldPushOffSheet);
    }

    // The runs of equal keys, as (first row, last row, the text shown).
    let key = |r: u32| {
        Row::new(r).map_or_else(String::new, |row| {
            book.formatted(sheet, CellRef::new(how.group_by, row))
        })
    };
    let mut runs: Vec<(u32, u32, String)> = Vec::new();
    for r in first..=last {
        let shown = key(r);
        match runs.last_mut() {
            Some((_, end, name)) if name.to_lowercase() == shown.to_lowercase() => *end = r,
            _ => runs.push((r, r, shown)),
        }
    }

    let label = how.function.label();
    let code = how.function.code();
    let bold = |book: &mut Spreadsheet, at: CellRef| {
        let ws = book.sheet(sheet);
        let id = ws
            .and_then(|ws| ws.get(at))
            .map(|c| c.style)
            .unwrap_or_default();
        let mut style = book.styles.get(id).cloned().unwrap_or_default();
        style.font.bold = true;
        let id = book.styles.intern(style);
        if let Some(ws) = book.sheet_mut(sheet) {
            ws.entry(at).style = id;
        }
    };
    let write_total =
        |book: &mut Spreadsheet, at_row: u32, from: u32, to: u32, text: String| -> Result<()> {
            let row = Row::new(at_row).ok_or(Error::WouldPushOffSheet)?;
            let ws = sheet_mut(book, sheet)?;
            for &col in &how.columns {
                let span = format!("{0}{1}:{0}{2}", col.to_letters(), from + 1, to + 1);
                ws.set(
                    CellRef::new(col, row),
                    CellValue::formula(format!("SUBTOTAL({code},{span})")),
                );
            }
            let at = CellRef::new(how.group_by, row);
            if !how.columns.contains(&how.group_by) {
                ws.set(at, text);
            }
            bold(book, at);
            Ok(())
        };

    // Bottom up, so the rows of groups not yet done stay where they were read.
    for (start, end, name) in runs.iter().rev() {
        let below = Row::new(end + 1).ok_or(Error::WouldPushOffSheet)?;
        super::insert_rows_with(book, sheet, below, 1, CopyOrigin::Blank)?;
        write_total(book, end + 1, *start, *end, format!("{name} {label}"))?;
    }
    let groups = u32::try_from(runs.len()).unwrap_or(u32::MAX);
    let last_total = last + groups;
    let grand = Row::new(last_total + 1).ok_or(Error::WouldPushOffSheet)?;
    super::insert_rows_with(book, sheet, grand, 1, CopyOrigin::Blank)?;
    write_total(
        book,
        last_total + 1,
        first,
        last_total,
        format!("Grand {label}"),
    )?;

    // Level one: everything above the grand total; level two: each group's
    // detail, which now sits `i` rows lower for the `i` totals above it.
    group(book, sheet, Lines::Rows, first, last_total)?;
    for (i, (start, end, _)) in (0u32..).zip(&runs) {
        group(book, sheet, Lines::Rows, start + i, end + i)?;
    }
    Ok(groups)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn groups_nest_and_split() {
        // Lines 2..=7 at level 1, 4..=5 at 2, and 9 alone.
        let levelled = [(2, 1), (3, 1), (4, 2), (5, 2), (6, 1), (7, 1), (9, 1)];
        assert_eq!(
            groups_of(&levelled),
            [
                OutlineGroup {
                    level: 1,
                    first: 2,
                    last: 7
                },
                OutlineGroup {
                    level: 1,
                    first: 9,
                    last: 9
                },
                OutlineGroup {
                    level: 2,
                    first: 4,
                    last: 5
                },
            ]
        );
    }

    #[test]
    fn summary_sides() {
        let g = OutlineGroup {
            level: 1,
            first: 0,
            last: 3,
        };
        assert_eq!(g.summary(true), Some(4));
        assert_eq!(g.summary(false), None);
    }
}
