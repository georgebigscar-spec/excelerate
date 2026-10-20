//! Refreshing a pivot table: reading its source again and laying the report
//! out in the sheet's cells.
//!
//! A pivot written by this crate (or read and changed) asks Excel to rebuild
//! it on open (`refreshOnLoad`), and Excel does; but every other reader of
//! the file (an HTML or CSV export, this crate's own formulas, a viewer
//! without a pivot engine) sees only the cells, which until now held the old
//! report or nothing. [`refresh_pivot`] computes them.
//!
//! The report is laid out in Excel's tabular form: one column per row
//! field, the items of an outer field shown once at the top of their group,
//! a total row under each outer group, the column fields' items across the
//! top and grand totals last. Ported from the pivot engine of gridcraft
//! (<https://github.com/storytold/gridcraft>, MIT OR Apache-2.0).

use std::cmp::Ordering;
use std::collections::{BTreeMap, HashMap};

use crate::coordinate::{CellRef, Col, Range, Row};
use crate::error::{Error, Result};
use crate::model::pivot::{PivotTable, Subtotal};
use crate::model::{CellValue, Spreadsheet};

/// The marker a row or column field list holds for "the values go here".
const VALUES: i32 = -2;

/// An item of a field: what the cell shows, and its number when it is one,
/// which is what items sort by.
#[derive(Debug, Clone)]
struct Item {
    label: String,
    number: Option<f64>,
}

impl Item {
    fn key(&self) -> String {
        self.label.to_lowercase()
    }
}

/// Excel's order of items: numbers ascending, then text without regard to
/// case, then `(blank)`.
fn order(a: &Item, b: &Item) -> Ordering {
    let blank = |i: &Item| i.label == BLANK;
    match (a.number, b.number) {
        (Some(x), Some(y)) => x.total_cmp(&y),
        (Some(_), None) => Ordering::Less,
        (None, Some(_)) => Ordering::Greater,
        (None, None) => blank(a).cmp(&blank(b)).then_with(|| a.key().cmp(&b.key())),
    }
}

const BLANK: &str = "(blank)";

/// The values a summary reads.
#[derive(Debug, Default, Clone)]
struct Bucket {
    numbers: Vec<f64>,
    filled: usize,
}

impl Bucket {
    fn add(&mut self, value: &CellValue) {
        match value.result() {
            CellValue::Empty => {}
            CellValue::Number(n) => {
                self.numbers.push(*n);
                self.filled += 1;
            }
            _ => self.filled += 1,
        }
    }

    fn merge(&mut self, other: &Self) {
        self.numbers.extend_from_slice(&other.numbers);
        self.filled += other.filled;
    }

    /// The summary, or `None` for an empty cell of the report.
    #[allow(clippy::cast_precision_loss, reason = "counts of source rows")]
    fn summary(&self, how: Subtotal) -> Option<f64> {
        let n = self.numbers.len() as f64;
        let sum: f64 = self.numbers.iter().sum();
        let mean = sum / n;
        let squares = || self.numbers.iter().map(|x| (x - mean).powi(2)).sum::<f64>();
        match how {
            Subtotal::Count => (self.filled > 0).then_some(self.filled as f64),
            _ if self.numbers.is_empty() => None,
            Subtotal::Sum => Some(sum),
            Subtotal::CountNums => Some(n),
            Subtotal::Average => Some(mean),
            Subtotal::Max => self.numbers.iter().copied().reduce(f64::max),
            Subtotal::Min => self.numbers.iter().copied().reduce(f64::min),
            Subtotal::Product => Some(self.numbers.iter().product()),
            Subtotal::StdDev => (n > 1.0).then(|| (squares() / (n - 1.0)).sqrt()),
            Subtotal::StdDevP => Some((squares() / n).sqrt()),
            Subtotal::Var => (n > 1.0).then(|| squares() / (n - 1.0)),
            Subtotal::VarP => Some(squares() / n),
            #[allow(unreachable_patterns, reason = "the summaries may grow")]
            _ => None,
        }
    }
}

/// What a data field's caption is when the report gives none: "Sum of
/// Amount", as Excel writes it.
fn caption(how: Subtotal, field: &str) -> String {
    let verb = match how {
        Subtotal::Sum => "Sum",
        Subtotal::Count | Subtotal::CountNums => "Count",
        Subtotal::Average => "Average",
        Subtotal::Max => "Max",
        Subtotal::Min => "Min",
        Subtotal::Product => "Product",
        Subtotal::StdDev => "StdDev",
        Subtotal::StdDevP => "StdDevp",
        Subtotal::Var => "Var",
        Subtotal::VarP => "Varp",
        #[allow(
            unreachable_patterns,
            clippy::match_same_arms,
            reason = "the summaries may grow"
        )]
        _ => "Sum",
    };
    format!("{verb} of {field}")
}

/// The sheet and range a pivot cache reads.
fn source_of(book: &Spreadsheet, table: &PivotTable) -> Result<(usize, Range)> {
    let cache = book
        .pivot_caches
        .iter()
        .find(|c| c.id == table.cache_id)
        .ok_or_else(|| Error::InvalidRange(format!("pivot {:?} has no cache", table.name)))?;
    let source = &cache.source;
    let sheet_named = |name: &str| {
        book.sheets()
            .iter()
            .position(|s| s.title().eq_ignore_ascii_case(name))
    };
    if let (Some(name), Some(range)) = (&source.sheet, source.range) {
        let sheet = sheet_named(name).ok_or_else(|| {
            Error::InvalidRange(format!("the pivot's source sheet {name:?} is gone"))
        })?;
        return Ok((sheet, range));
    }
    let Some(name) = &source.name else {
        return Err(Error::InvalidRange(format!(
            "pivot {:?} reads no range this crate can find",
            table.name
        )));
    };
    // A table by its name, or a name standing for a range.
    for (index, sheet) in book.sheets().iter().enumerate() {
        if let Some(t) = sheet.tables.iter().find(|t| {
            t.name.eq_ignore_ascii_case(name) || t.display_name.eq_ignore_ascii_case(name)
        }) {
            return Ok((index, t.range));
        }
    }
    let formula = book
        .defined_names
        .iter()
        .find(|d| d.name.eq_ignore_ascii_case(name))
        .map(|d| d.formula.trim_start_matches('=').to_owned())
        .ok_or_else(|| {
            Error::InvalidRange(format!("the pivot's source {name:?} is not defined"))
        })?;
    let (sheet, range) = formula
        .rsplit_once('!')
        .ok_or_else(|| Error::InvalidRange(format!("{name:?} is not a range: {formula}")))?;
    let sheet = sheet.trim_matches('\'').replace("''", "'");
    let index = sheet_named(&sheet)
        .ok_or_else(|| Error::InvalidRange(format!("{name:?} points at a sheet that is gone")))?;
    Ok((index, Range::parse(&range.replace('$', ""))?))
}

/// One laid-out cell, relative to the report's corner.
type Laid = BTreeMap<(u32, u32), CellValue>;

/// Reads a pivot's source again and writes the report into its sheet's
/// cells, from the corner of its current location: Excel's `PivotTable` -
/// Refresh, laid out in tabular form.
///
/// The cells the old report covered are cleared first; the new extent
/// becomes the pivot's location. Items are sorted as Excel sorts them, an
/// empty source cell is the item `(blank)`, and a value field without a
/// caption gets Excel's ("Sum of Amount"). Page (filter) fields filter
/// nothing, since the model does not record which item they show. Column
/// fields get a grand total but no subtotals between them. Values are written
/// as numbers in the cells' own formats; the value fields' number formats
/// are not applied.
///
/// The pivot is marked as laid out again, so the file still asks Excel to
/// rebuild it on open: Excel's own layout replaces this one there.
///
/// Returns the range the report now covers.
///
/// # Errors
/// [`Error::SheetIndexOutOfRange`] for a missing sheet, and
/// [`Error::InvalidRange`] when `pivot` is not one of its pivots, the
/// pivot has no location or no value field, or its source cannot be found.
#[allow(clippy::too_many_lines, reason = "one layout, header to grand total")]
pub fn refresh_pivot(book: &mut Spreadsheet, sheet: usize, pivot: usize) -> Result<Range> {
    let ws = book
        .sheet(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    let table = ws
        .pivot_tables
        .get(pivot)
        .ok_or_else(|| Error::InvalidRange(format!("the sheet has no pivot table {pivot}")))?
        .clone();
    let Some(old) = table.location else {
        return Err(Error::InvalidRange(format!(
            "pivot {:?} has no location",
            table.name
        )));
    };
    if table.data_fields.is_empty() {
        return Err(Error::InvalidRange(format!(
            "pivot {:?} has no value field",
            table.name
        )));
    }
    let (source_sheet, source) = source_of(book, &table)?;
    let src = book
        .sheet(source_sheet)
        .ok_or(Error::SheetIndexOutOfRange(source_sheet))?;
    let last = src.dimension().map_or_else(
        || source.start.row.index(),
        |used| source.end.row.index().min(used.end.row.index()),
    );
    let columns: Vec<Col> = (source.start.col.index()..=source.end.col.index())
        .filter_map(Col::new)
        .collect();
    let names: Vec<String> = columns
        .iter()
        .map(|&c| book.formatted(source_sheet, CellRef::new(c, source.start.row)))
        .collect();

    let fields = |list: &[i32]| -> Vec<usize> {
        list.iter()
            .filter(|&&f| f != VALUES)
            .filter_map(|&f| usize::try_from(f).ok())
            .filter(|&f| f < columns.len())
            .collect()
    };
    let rows_by = fields(&table.row_fields);
    let cols_by = fields(&table.column_fields);
    let data: Vec<(usize, Subtotal, String)> = table
        .data_fields
        .iter()
        .filter_map(|d| {
            let field = usize::try_from(d.field)
                .ok()
                .filter(|&f| f < columns.len())?;
            let name = d
                .name
                .clone()
                .unwrap_or_else(|| caption(d.subtotal, &names[field]));
            Some((field, d.subtotal, name))
        })
        .collect();
    if data.is_empty() {
        return Err(Error::InvalidRange(format!(
            "pivot {:?} sums fields its source does not have",
            table.name
        )));
    }

    // Every record's items on the row and column fields, and its buckets.
    let item_of = |row: Row, field: usize| {
        let at = CellRef::new(columns[field], row);
        let value = src
            .get(at)
            .map(|c| c.value.result().clone())
            .unwrap_or_default();
        if value.is_empty() {
            return Item {
                label: BLANK.into(),
                number: None,
            };
        }
        Item {
            label: book.formatted(source_sheet, at),
            number: match value {
                CellValue::Number(n) => Some(n),
                _ => None,
            },
        }
    };
    let mut items: HashMap<String, Item> = HashMap::new();
    // (row keys, column keys) -> one bucket per data field.
    let mut cells: HashMap<(Vec<String>, Vec<String>), Vec<Bucket>> = HashMap::new();
    for index in source.start.row.index() + 1..=last {
        let Some(row) = Row::new(index) else { continue };
        let mut key = |fields: &[usize]| -> Vec<String> {
            fields
                .iter()
                .map(|&f| {
                    let item = item_of(row, f);
                    let k = format!("{f}\u{1}{}", item.key());
                    items.entry(k.clone()).or_insert(item);
                    k
                })
                .collect()
        };
        let (r, c) = (key(&rows_by), key(&cols_by));
        let buckets = cells
            .entry((r, c))
            .or_insert_with(|| vec![Bucket::default(); data.len()]);
        for (bucket, (field, _, _)) in buckets.iter_mut().zip(&data) {
            if let Some(cell) = src.get(CellRef::new(columns[*field], row)) {
                bucket.add(&cell.value);
            }
        }
    }
    let sorted = |keys: Vec<Vec<String>>| {
        let mut keys = keys;
        keys.sort_by(|a, b| {
            a.iter()
                .zip(b)
                .map(|(x, y)| order(&items[x], &items[y]))
                .find(|o| o.is_ne())
                .unwrap_or(Ordering::Equal)
        });
        keys.dedup();
        keys
    };
    let row_keys = sorted(cells.keys().map(|(r, _)| r.clone()).collect());
    let col_keys = sorted(cells.keys().map(|(_, c)| c.clone()).collect());
    let total = |r: &[String], c: Option<&[String]>| -> Vec<Bucket> {
        let mut out = vec![Bucket::default(); data.len()];
        for ((rk, ck), buckets) in &cells {
            if rk.starts_with(r) && c.is_none_or(|c| ck == c) {
                for (o, b) in out.iter_mut().zip(buckets) {
                    o.merge(b);
                }
            }
        }
        out
    };

    // The columns of the values area: each column key times each data
    // field, then the grand totals.
    let label_cols = u32::try_from(rows_by.len()).unwrap_or(u32::MAX);
    let mut value_cols: Vec<(Option<Vec<String>>, usize)> = Vec::new();
    for key in &col_keys {
        for d in 0..data.len() {
            value_cols.push((Some(key.clone()), d));
        }
    }
    let grand_cols = !cols_by.is_empty() && table.column_grand_totals;
    if grand_cols {
        for d in 0..data.len() {
            value_cols.push((None, d));
        }
    }
    if cols_by.is_empty() {
        value_cols.truncate(data.len());
    }

    let mut laid: Laid = BTreeMap::new();
    let mut put = |r: u32, c: u32, v: CellValue| {
        if !v.is_empty() {
            laid.insert((r, c), v);
        }
    };
    let text = |s: &str| CellValue::text(s);
    // Header rows: the column fields' items, one row each, then the data
    // captions when there are several, with the field names on the last.
    let mut header = 0u32;
    if cols_by.is_empty() {
        for (c, (_, d)) in (label_cols..).zip(&value_cols) {
            put(0, c, text(&data[*d].2));
        }
    } else {
        if data.len() == 1 {
            put(0, 0, text(&data[0].2));
        }
        for (i, &f) in (0u32..).zip(&cols_by) {
            put(0, label_cols + i, text(&names[f]));
        }
        header = 1;
        for level in 0..cols_by.len() {
            for (c, (key, d)) in (label_cols..).zip(&value_cols) {
                let label = match key {
                    Some(key) if *d == 0 || data.len() == 1 => items[&key[level]].label.clone(),
                    None if *d == 0 && level == 0 => "Grand Total".to_owned(),
                    _ => continue,
                };
                put(header, c, text(&label));
            }
            header += 1;
        }
        if data.len() > 1 {
            for (c, (key, d)) in (label_cols..).zip(&value_cols) {
                let name = &data[*d].2;
                let label = if key.is_none() {
                    format!("Total {name}")
                } else {
                    name.clone()
                };
                put(header, c, text(&label));
            }
            header += 1;
        }
        header -= 1;
    }
    for (c, &f) in (0u32..).zip(&rows_by) {
        put(header, c, text(&names[f]));
    }
    let first_data_row = header + 1;

    let write_values = |put: &mut dyn FnMut(u32, u32, CellValue), r: u32, prefix: &[String]| {
        for (c, (key, d)) in (label_cols..).zip(&value_cols) {
            let buckets = total(prefix, key.as_deref());
            if let Some(v) = buckets[*d].summary(data[*d].1) {
                put(r, c, CellValue::Number(v));
            }
        }
    };
    let mut r = first_data_row;
    if rows_by.is_empty() {
        write_values(&mut put, r, &[]);
        r += 1;
    } else {
        let subtotalled: Vec<bool> = rows_by
            .iter()
            .map(|&f| table.fields.get(f).is_none_or(|p| p.default_subtotal))
            .collect();
        for (i, key) in row_keys.iter().enumerate() {
            let previous = i.checked_sub(1).map(|p| &row_keys[p]);
            for (level, k) in key.iter().enumerate() {
                // An outer item shows once, at the top of its group.
                let fresh = previous.is_none_or(|p| p[..=level] != key[..=level]);
                if fresh {
                    put(r, u32::try_from(level).unwrap_or(0), text(&items[k].label));
                }
            }
            write_values(&mut put, r, key);
            r += 1;
            // Totals of the outer groups this row closes, innermost first.
            let next = row_keys.get(i + 1);
            for level in (0..key.len().saturating_sub(1)).rev() {
                let closes = next.is_none_or(|n| n[..=level] != key[..=level]);
                if closes && subtotalled[level] {
                    let label = format!("{} Total", items[&key[level]].label);
                    put(r, u32::try_from(level).unwrap_or(0), text(&label));
                    write_values(&mut put, r, &key[..=level]);
                    r += 1;
                }
            }
        }
        if table.row_grand_totals {
            put(r, 0, text("Grand Total"));
            write_values(&mut put, r, &[]);
            r += 1;
        }
    }

    let width = label_cols + u32::try_from(value_cols.len()).unwrap_or(0);
    let corner = old.start;
    let end_row = Row::new(corner.row.index() + r - 1).ok_or(Error::WouldPushOffSheet)?;
    let end_col =
        Col::new(corner.col.index() + width.max(1) - 1).ok_or(Error::WouldPushOffSheet)?;
    let extent = Range::new(corner, CellRef::new(end_col, end_row));

    let ws = book
        .sheet_mut(sheet)
        .ok_or(Error::SheetIndexOutOfRange(sheet))?;
    // The old report and whatever the new one now covers.
    let stale: Vec<CellRef> = ws
        .iter()
        .map(|(at, _)| at)
        .filter(|&at| old.contains(at) || extent.contains(at))
        .collect();
    for at in stale {
        ws.remove(at);
    }
    for ((dr, dc), value) in laid {
        if let (Some(row), Some(col)) = (
            Row::new(corner.row.index() + dr),
            Col::new(corner.col.index() + dc),
        ) {
            ws.set(CellRef::new(col, row), value);
        }
    }
    let table = &mut ws.pivot_tables[pivot];
    table.location = Some(extent);
    table.first_data_row = first_data_row;
    table.first_data_col = label_cols;
    if let Some(origin) = &mut table.origin {
        origin.relaid = true;
    }
    Ok(extent)
}

/// [`refresh_pivot`] for every pivot table of the workbook, as Data -
/// Refresh All does. Returns how many were refreshed.
///
/// # Errors
/// The first error a pivot gives; those before it stay refreshed.
pub fn refresh_all_pivots(book: &mut Spreadsheet) -> Result<usize> {
    let mut done = 0;
    for sheet in 0..book.sheets().len() {
        let count = book.sheet(sheet).map_or(0, |ws| ws.pivot_tables.len());
        for pivot in 0..count {
            refresh_pivot(book, sheet, pivot)?;
            done += 1;
        }
    }
    Ok(done)
}
