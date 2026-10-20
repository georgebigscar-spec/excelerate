# Editing the grid

Setting a cell is one call and touches one cell:

```rust
use excelerate::model::{Spreadsheet, Worksheet};
use excelerate::at;

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Sheet1")?;
sheet.set(at!("B2"), 42.0);
book.add_sheet(sheet)?;
# Ok::<(), excelerate::Error>(())
```

Inserting a row is not. Everything that named a cell below the insertion point
now names the wrong cell, and those names are spread over the whole workbook:
formulas on other sheets, defined names, merges, hyperlinks, data validations,
conditional formats, protected ranges, the autofilter, table ranges, chart
series, drawing anchors, page breaks. `excelerate::edit` moves them all.

## Rows and columns

```rust
use excelerate::edit::{insert_rows, insert_columns, remove_rows, remove_columns};
use excelerate::{Col, Row};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Sheet1")?)?;

// Three rows before row 5, on sheet 0. Rows are 0-based here, and
// `Row::new` answers `None` past the last row of a sheet.
let row = Row::new(4).unwrap();
insert_rows(&mut book, 0, row, 3)?;
remove_rows(&mut book, 0, row, 3)?;

let column = Col::new(1).unwrap();
insert_columns(&mut book, 0, column, 1)?;
remove_columns(&mut book, 0, column, 1)?;
# Ok::<(), excelerate::Error>(())
```

`Row::new` and `Col::new` count from zero, the way the model does. `Row(4)` is
the row a user calls 5.

An insertion that would push a cell holding something, or a merged area, past
the last row or column fails with `Error::WouldPushOffSheet` and leaves the
workbook as it was, as Excel refuses it. Row heights, column widths and
styles at the edge are cut off without a word, as Excel does.

## What happens to references

These rules are Excel's, and `tests/edit.rs` holds one test each.

**`$` does not protect anything.** It says a reference will not shift when the
*formula* is copied, which is a different question. A row inserted above the
cell moves the cell, so `$A$5` becomes `$A$6`.

**A deleted cell gives `#REF!`; a range that loses part of itself is
narrowed.** Delete row 5 and `=A5` becomes `=#REF!`, while `=SUM(A1:A10)`
becomes `=SUM(A1:A9)`. Insert inside a range and the range grows.

**The sheet qualifier survives the break.** `Data!A3` whose row is deleted
becomes `Data!#REF!`, not a bare `#REF!`. Excel writes it the same way.

**The cached result is dropped.** It was the answer to the old text. Recalculate
after an edit, or read a stale number later.

## Moving and copying a block

Copying and moving differ in exactly one way, the way Excel's Ctrl+C and
Ctrl+X differ, and it is about the formulas:

```rust
use excelerate::edit::{copy_range, move_range};
use excelerate::{Range, at};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;
let from = Range::parse("A1:B3")?;

// A copy is rewritten as if it had been written where it lands: `=A1` one
// column right reads `=B1`, `$A$1` stays where it is.
copy_range(&mut book, 0, from, 0, at!("D1"))?;

// A move keeps its answers - `=A1` still reads A1 - and the formulas
// elsewhere that read the moved cells follow them instead.
move_range(&mut book, 0, from, 0, at!("A10"))?;
# Ok::<(), excelerate::Error>(())
```

Both carry the values, the styles and the merges lying wholly inside the
block, both may overlap their own source, and both empty a target cell the
source had nothing in - which is what pasting a block does. A block that would
land off the sheet is an error rather than a silent clamp.

Two limits worth knowing. A move to another sheet qualifies what the moved
formulas read without naming a sheet, so `=Z9` becomes `=Data!Z9` - right, but
spelled out per reference, so a range reads `Data!A1:Data!A3`. And a move to
another sheet leaves the references *from elsewhere* alone; within one sheet
they follow the cells.

## Inserting cells, not rows

`insert_cells` moves part of a row, which is what Excel's "Insert Cells" does:
only the columns the block spans move, the rest of the sheet stays put.

```rust
use excelerate::edit::{Axis, insert_cells, remove_cells};
use excelerate::Range;
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;
let area = Range::parse("A1:A2")?;

insert_cells(&mut book, 0, area, Axis::Rows)?;     // push the cells below down
remove_cells(&mut book, 0, area, Axis::Rows)?;     // pull them back up
insert_cells(&mut book, 0, area, Axis::Columns)?;  // or sideways
# Ok::<(), excelerate::Error>(())
```

A reference travels only when the whole of it travels, which is why this is a
separate edit from `insert_rows` rather than a special case of it. Data that
the push would shove off the end of the sheet is an error; blank cells there
are no loss, and Excel draws the line in the same place.

## Formatting what you insert

A plain insert leaves the new rows blank. Excel's Insert takes the formatting
of the row above instead, and `CopyOrigin` gives you the same choice - the
`CopyOrigin` argument of Excel's `Range.Insert`:

```rust
use excelerate::edit::{Axis, CopyOrigin, insert_cells_with, insert_columns_with, insert_rows_with};
use excelerate::{Col, Range, Row};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;

// Two rows before row 5, styled like row 4: its cell styles, its own row
// style and its height.
insert_rows_with(&mut book, 0, Row::new(4).unwrap(), 2, CopyOrigin::Before)?;
// One column before C, styled like the column that ends up right of it.
insert_columns_with(&mut book, 0, Col::new(2).unwrap(), 1, CopyOrigin::After)?;
// Insert Cells takes the same flag; only cell styles are copied there.
insert_cells_with(&mut book, 0, Range::parse("B2:B3")?, Axis::Rows, CopyOrigin::Before)?;
# Ok::<(), excelerate::Error>(())
```

Values are never copied, and neither are hidden or collapsed outline states:
an inserted row that came out hidden would be a strange surprise.
`insert_rows` and friends are the `CopyOrigin::Blank` case. In JS the flag is
the last argument, and there it defaults to `"before"` the way Excel does.

## Sorting

`sort_range` is Excel's Data - Sort. Keys are sheet columns, the first key
decides first, and equal rows keep their order:

```rust
use excelerate::edit::{SortKey, SortOptions, sort_range, sort_range_with, sort_table};
use excelerate::{Col, Range, at};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# let mut sheet = Worksheet::new("Data")?;
# sheet.set(at!("A1"), "Region");
# sheet.set(at!("B1"), "Amount");
# book.add_sheet(sheet)?;

// Rows 2..100 by column C, largest first, then by A.
let keys = [
    SortKey::column(Col::new(2).unwrap()).descending(),
    SortKey::column(Col::new(0).unwrap()),
];
sort_range(&mut book, 0, Range::parse("A2:D100")?, &keys)?;

// With a header row: it stays on top, and keys can name its cells.
let options = SortOptions { header: true, ..SortOptions::default() };
let keys = [SortKey::header("Region"), SortKey::header("Amount").descending()];
sort_range_with(&mut book, 0, Range::parse("A1:D100")?, &keys, options)?;
# Ok::<(), excelerate::Error>(())
```

A table knows its own header and totals rows, so `sort_table` takes only its
name and the column names: `sort_table(&mut book, "Sales", &[SortKey::header("Amount")])`.
`SortOptions::orientation = Axis::Columns` sorts columns left to right, keyed
by `SortKey::row`.

The order is Excel's: numbers, then text ignoring case, then `FALSE`, `TRUE`
and errors. Empty cells go last whichever way the key runs. A formula sorts by
its cached value, so recalculate first if the cache may be stale. A moved
formula is rewritten as if copied to its new row, so `=B7*2` that lands on row
3 reads `=B3*2` - it still reads its own row. Formulas outside the range are
not retargeted, in Excel or here. A merge crossing the range is an error;
Excel refuses that sort too.

## Filling

`fill` is Ctrl+D and Ctrl+R: the first row (or column) of a range copied over
the rest, formulas rewritten as in a copy. `fill_series` is dragging the fill
handle - it continues what the first cells start:

```rust
use excelerate::edit::{Axis, fill, fill_series};
use excelerate::Range;
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;

fill(&mut book, 0, Range::parse("C2:C500")?, Axis::Rows)?;        // =A2*B2 down to row 500
fill_series(&mut book, 0, Range::parse("A1:A12")?, Axis::Rows)?;  // Jan, Feb, ... Dec
# Ok::<(), excelerate::Error>(())
```

What `fill_series` does with the leading filled cells of each line:

| Seed | Continues as |
|---|---|
| `1, 3` | `5, 7, 9` - the least-squares trend, so `1, 2, 4` goes on `5.33, 6.83` |
| `5` | `5, 5, 5` - one number is copied, as in Excel |
| a date | the next days |
| `Кв1`, `Item 007` | `Кв2`, `Item 008` - the last number counts on, zeros keep their width |
| `Jan`, `Monday`, `Январь`, `ПН` | the rest of the month or weekday list, in the seed's case |
| anything else | the seed repeated, formulas moved as in a copy |

Every filled cell takes the style of the seed cell it continues.

`fill_series_back` is the handle dragged up or left: the seed is the run of
filled cells at the end of each line, and the series runs back from it - `1, 2`
goes `0, -1`, a single `Кв3` goes `Кв2`, a single date steps back a day, a
repeat keeps its phase.

## Removing duplicates and splitting text

`remove_duplicates` is Excel's Data - Remove Duplicates: rows are compared by
the text their cells show, without regard to case, the first of each set is
kept and the rows below close up inside the range. An empty column list
compares every column of the range.

```rust
use excelerate::edit::{SplitBy, remove_duplicates, text_to_columns};
use excelerate::{Range, at};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# let mut sheet = Worksheet::new("Data")?;
# sheet.set(at!("A1"), "Name;Age");
# sheet.set(at!("A2"), "Ann;41");
# sheet.set(at!("A3"), "ann;41");
# book.add_sheet(sheet)?;

// "Ann;41" in A2 becomes "Ann" in A2 and the number 41 in B2.
text_to_columns(&mut book, 0, Range::parse("A1:A3")?, at!("A1"), &SplitBy::delimiter(';'))?;

// Row 3 repeats row 2; the first row is a header and stays.
let done = remove_duplicates(&mut book, 0, Range::parse("A1:B3")?, &[], true)?;
assert_eq!((done.removed, done.remaining), (1, 1));
# Ok::<(), excelerate::Error>(())
```

`SplitBy::Delimited` takes several delimiters, a text qualifier and Excel's
"treat consecutive delimiters as one"; `SplitBy::FixedWidth` cuts at
character positions. Pieces become numbers, booleans and errors the way a
CSV field does.

## Moving a sheet

```rust
use excelerate::edit::move_sheet;
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;
# book.add_sheet(Worksheet::new("Report")?)?;
move_sheet(&mut book, 1, 0)?;   // the report opens first now
# Ok::<(), excelerate::Error>(())
```

Formulas do not change: a sheet is named, not numbered. What does change is
every index kept beside them - the active tab, the sheet a defined name
belongs to - and those are renumbered.

## Renaming and removing sheets

`Worksheet::set_title` changes the tab and nothing else. The sheet's name also
lives in the qualifier of every reference to it, anywhere in the workbook, so
renaming through the model alone leaves those references pointing at a name
that no longer exists.

```rust
use excelerate::edit::{remove_sheet, rename_sheet};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Data")?)?;
# book.add_sheet(Worksheet::new("Report")?)?;

rename_sheet(&mut book, 0, "Sales 2026")?;   // `Data!A1` becomes `'Sales 2026'!A1`
remove_sheet(&mut book, 0)?;                 // and now it is `#REF!A1`
# Ok::<(), excelerate::Error>(())
```

Removing a sheet turns references to it into `#REF!A1` - the bang is part of
the error, not a separator after it - and fixes the sheet indices of defined
names. A 3D range `Sheet1:Sheet3!A1` narrows to what is left when an end sheet
goes and is untouched when a middle one does: both ends are still there, the
span just covers one sheet fewer.

Renaming all 38 sheets of a 19,826-formula workbook takes 360 ms; removing one
takes 10 ms.

## Grouping, subtotals and other Data tools

Rows and columns group as Data - Group does, collapse and expand as the
outline's buttons do, and `subtotal` builds a whole outline from a sorted
list:

```rust
use excelerate::edit::{SubtotalFunction, Subtotals, collapse_rows, subtotal};
use excelerate::{Col, Range, Row, at};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# let mut sheet = Worksheet::new("Sales")?;
# sheet.set(at!("A1"), "Region");
# sheet.set(at!("B1"), "Amount");
# for (r, (region, amount)) in (2..).zip([("East", 10.0), ("East", 20.0), ("West", 5.0)]) {
#     sheet.set_at(r, 1, region)?;
#     sheet.set_at(r, 2, amount)?;
# }
# book.add_sheet(sheet)?;

// "East Total" under rows 2-3, "West Total" under row 5, "Grand Total" last.
let groups = subtotal(&mut book, 0, Range::parse("A1:B4")?, &Subtotals {
    group_by: Col::from_letters("A")?,
    function: SubtotalFunction::Sum,
    columns: vec![Col::from_letters("B")?],
})?;
assert_eq!(groups, 2);

// Row 4 is the summary of East's rows: the minus button over it.
collapse_rows(&mut book, 0, Row::from_one_based(4)?)?;
# Ok::<(), excelerate::Error>(())
```

`flash_fill` fills a column by example, as Ctrl+E does, from the columns to
its left; `FlashFill::learn` does the same without a sheet.
`advanced_filter` applies a criteria range as Data - Advanced reads it
(conditions in a row all hold, rows are alternatives, `>5`, `<>x`, `=`,
wildcards), and `consolidate` combines ranges by their labels or their
positions.

A pivot table written by a program, this crate included, holds no report in
its cells until Excel opens it and rebuilds it. `refresh_pivot` lays the
report out from the pivot's source, in Excel's tabular form, so an HTML or
CSV export, or this crate's formulas, see the numbers too.

## What else moves

Drawing anchors move without going through the model: `edit::anchor` rewrites
the row and column numbers inside the carried part and leaves the rest of the
bytes alone. `editAs` decides what an insertion inside an object does -
`twoCell` (the default) drags each corner with its cell, `oneCell` moves the
object whole, `absolute` does not move it. Offsets inside a cell are not grid
indices and are left as they are.

Chart series (`<c:f>`) are rewritten the same way, by `edit::chart`. The part
is full of numbers, so a wider rewrite would corrupt them.

The sheet's `<extLst>` travels as bytes too, and the rules in it name cells:
the `x14` half of a data bar, a validation whose list sits on another sheet,
sparklines. `edit::extension` rewrites the text of two elements only:
`<xm:sqref>` moves with the cells of the sheet that holds it, and `<xm:f>`
is rewritten like any formula, on every sheet and on a rename. A rule left
with no cells is cut out with its element. `Worksheet::sparklines` moves in
the model the same way, so the writer finds the two in agreement and keeps
the bytes.

Comment addresses move in the model. Before that they did not move at all, and
a note stayed on the cell it used to describe.

## Editing in a loop

Each call walks the workbook. For a batch, prefer one call over a range to
several calls over single rows, and recalculate once at the end. When the rows
go in at many places - a subtotal under every group - the `*_many` twins take
all of them and walk the workbook once:

```rust
use excelerate::edit::{CopyOrigin, insert_rows_many, remove_rows_many};
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Sheet1")?)?;
use excelerate::Row;

// One row above rows 5, 9 and 14, counted as the sheet is now, in any order.
let places = [4, 8, 13].map(|r| (Row::new(r).unwrap(), 1));
insert_rows_many(&mut book, 0, &places, CopyOrigin::Before)?;

// Overlapping blocks are merged: this removes rows 3 to 6.
let spans = [(Row::new(2).unwrap(), 3), (Row::new(4).unwrap(), 2)];
remove_rows_many(&mut book, 0, &spans)?;
# Ok::<(), excelerate::Error>(())
```

The result is that of the single edits made from the far end. Four thousand
subtotal rows into forty thousand took 73 seconds one call at a time; the
batch takes 29 ms. `insert_columns_many` and `remove_columns_many` are the
same for columns.

For one block:

```rust
use excelerate::edit::insert_rows;
use excelerate::formula::eval::{Dependencies, recalculate};
use excelerate::progress::Options;
# use excelerate::model::{Spreadsheet, Worksheet};
# let mut book = Spreadsheet::empty();
# book.add_sheet(Worksheet::new("Sheet1")?)?;
use excelerate::Row;

insert_rows(&mut book, 0, Row::new(0).unwrap(), 10)?;   // not ten calls of one row
recalculate(&mut book, None, &Options::default());
# let _ = Dependencies::of(&book);
# Ok::<(), excelerate::Error>(())
```

If you are making a series of edits and recalculating between them, hold the
dependency index rather than rebuilding it each time - see
[Formulas](formulas.md#recalculating-after-an-edit).
