# Excelerate

Read, write and recalculate spreadsheets in pure Rust.

Opens `.xlsx`, `.xlsb`, `.xls`, `.ods`, `.csv`, `.html`, `.slk`, `.gnumeric`
and SpreadsheetML 2003, hands you the workbook as a plain Rust struct, evaluates
formulas (518 Excel functions and counting) and writes the whole thing back.
No Excel, no LibreOffice, no COM, no headless anything - just the crate.

```toml
[dependencies]
excelerate = "0.15"
```

```rust,no_run
use excelerate::{at, reader};

let book = reader::read("report.xlsx")?;
let sheet = book.active_sheet().unwrap();
println!("{}: {} non-empty cells", sheet.title(), sheet.len());
println!("B2 = {:?}", sheet.get(at!("B2")).map(|c| &c.value));
# Ok::<(), excelerate::Error>(())
```

`at!("B2")` is an address checked while compiling; `set_at(row, column, ..)`
takes numbers, and `CellRef::parse` takes text that arrives at run time.

## Why you might want it

- **The format is sniffed, not guessed from the extension.** `reader::read`
  looks at the bytes first, so a `.txt` that is secretly a zip package still
  opens as xlsx. The extension only gets a vote when the bytes stay quiet.
- **Formulas actually run.** Recalculate the whole book, one sheet, or just
  what an edit touched - the incremental path is roughly 20x cheaper on a book
  with 20k formulas.
- **Round-trips without eating your file.** Read -> write -> read keeps styles,
  merges, conditional formatting, protection, autofilters with their sort
  state, sparklines and print setup; text with a colour or font change
  mid-cell keeps its runs in xlsx, xls, ODS and HTML alike.
  Charts, pictures, shapes and comments are modelled and written back;
  anything the crate does not model yet rides through byte for byte along with
  its relationships.
- **Edits the grid like Excel does.** Insert or remove rows, columns and
  cells, copy or move a block, reorder the sheets - and the formulas, merges,
  links, tables and drawings across the whole workbook follow. A copy rewrites
  its formulas, a move keeps them and drags the references to it along.
  Inserted rows can take the formatting of their neighbour, ranges and tables
  sort in Excel's order, and the fill handle's series (`1, 3, 5`, `Jan, Feb`,
  `Q1, Q2`) are one call.
- **Builds for WebAssembly**, and ships as two npm packages:
  [`@rosperitus/excelerate`](https://www.npmjs.com/package/@rosperitus/excelerate)
  and a read-only `@rosperitus/excelerate-reader` at a third of the size.
- **No `unsafe`**, `clippy::pedantic` clean, and the deps are `zip`, `quick-xml`,
  `flate2`, `thiserror`, `regex` for the `REGEX*` functions, and `sha1`/`sha2`/
  `getrandom` for password hashes - all pure Rust, so a wasm build needs no C
  toolchain.

## Use cases, step by step

Each block below runs as a test (`cargo test --doc`), against the files in
[`tests/fixtures/`](tests/fixtures), so the asserts are the answers you get.

1. [Read a workbook and pull values out](#read-a-workbook-and-pull-values-out)
2. [Build a report from scratch](#build-a-report-from-scratch)
3. [Change an input, recalculate what depends on it](#change-an-input-recalculate-what-depends-on-it)
4. [Insert rows and sort; formulas follow](#insert-rows-and-sort-formulas-follow)
5. [Convert between formats in memory](#convert-between-formats-in-memory)
6. [Import a CSV with `;` and a decimal comma](#import-a-csv-with--and-a-decimal-comma)
7. [Ask a table by name](#ask-a-table-by-name)
8. [Notes, links, a drop-down and a password](#notes-links-a-drop-down-and-a-password)
9. [Your own function, with a progress callback](#your-own-function-with-a-progress-callback)
10. [See how a cell is styled](#see-how-a-cell-is-styled)
11. [Title, author and your own fields](#title-author-and-your-own-fields)

### Read a workbook and pull values out

```rust
use excelerate::model::CellValue;
use excelerate::{at, reader};

// 1. Open it. The bytes decide the format; the extension is a fallback.
let book = reader::read("tests/fixtures/sample.xlsx")?;

// 2. Pick a sheet by name. `book.sheets()[0]` works by position.
let data = book.sheet_index_by_name("Data").unwrap();
let sheet = &book.sheets()[data];

// 3. One cell: the stored value, formula or not.
let b1 = at!("B1");
assert_eq!(sheet.get(b1).and_then(|c| c.value.as_number()), Some(42.0));

// 4. The same cell as Excel shows it, with its number format applied.
println!("B1 shows {}", book.formatted(data, b1));

// 5. Every non-empty cell, in address order. Empty ones are never visited.
for (at, cell) in sheet.iter() {
    if let CellValue::Formula { formula, cached } = &cell.value {
        println!("{at}: ={formula}, cached as {cached:?}");
    }
}
# Ok::<(), excelerate::Error>(())
```

### Build a report from scratch

Styles live in the workbook and a cell holds an id into them. `intern` hands
back the same id for the same style, so a thousand money cells cost one entry.

```rust
use excelerate::formula::eval::recalculate;
use excelerate::model::{CellValue, Pane, PanePosition, PaneState, Spreadsheet, Worksheet};
use excelerate::progress::Options;
use excelerate::style::{NumberFormat, Style};
use excelerate::{CellRef, Col, at, reader, writer};
use std::io::Cursor;

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Sales")?;

// 1. Two styles: a bold header and a money format.
let mut header = Style::default();
header.font.bold = true;
let header = book.styles.intern(header);
let mut money = Style::default();
money.number_format = NumberFormat::Custom("#,##0.00".into());
let money = book.styles.intern(money);

// 2. The header row, then one row per record.
sheet.set_styled(at!("A1"), "Item", header);
sheet.set_styled(at!("B1"), "Amount", header);
for (row, (item, amount)) in (2..).zip([("Tea", 180.0), ("Coffee", 1350.5)]) {
    sheet.set_at(row, 1, item)?;
    sheet.set_styled(CellRef::from_row_col(row, 2)?, amount, money);
}

// 3. A total. It is only text until something computes it.
sheet.set_styled(at!("B4"), CellValue::formula("SUM(B2:B3)"), money);

// 4. A wider first column, and the header row frozen when scrolling.
sheet.set_column_width(Col::from_letters("A")?, Some(18.0));
sheet.view.pane = Some(Pane {
    x_split: 0,
    y_split: 1,
    top_left_cell: Some(at!("A2")),
    active_pane: PanePosition::BottomLeft,
    state: PaneState::Frozen,
});
book.add_sheet(sheet)?;

// 5. Compute, then write. `writer::write_xlsx(&book, "sales.xlsx")` for a file.
recalculate(&mut book, None, &Options::default());
let mut bytes = Vec::new();
writer::write_xlsx_to(&book, Cursor::new(&mut bytes))?;

// What was written reads back with the total in place, formatted.
let again = reader::read_bytes(&bytes, None)?;
assert_eq!(again.formatted(0, at!("B4")), "1,530.50");
# Ok::<(), excelerate::Error>(())
```

### Change an input, recalculate what depends on it

`recalculate` computes the whole workbook. After an edit, `Dependencies`
computes only the formulas that read what changed; build it once and keep it
for a series of edits.

```rust
use excelerate::formula::eval::{Dependencies, recalculate};
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::progress::Options;
use excelerate::at;

let (a1, a3) = (at!("A1"), at!("A3"));
let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Model")?;
sheet.set(a1, 10.0);
sheet.set(at!("A2"), CellValue::formula("A1*2"));
sheet.set(a3, CellValue::formula("A2+1"));
sheet.set(at!("B1"), CellValue::formula("NOW()")); // reads no cell
book.add_sheet(sheet)?;
recalculate(&mut book, None, &Options::default());

// 1. Index which formula reads which cell.
let deps = Dependencies::of(&book);

// 2. Change the input.
book.sheet_mut(0).unwrap().set(a1, 20.0);

// 3. Recompute what the change reaches: A2, then A3. `NOW()` is volatile,
//    so it runs on every pass too.
let computed = deps.recalculate_from(&mut book, &[(0, a1)]);
assert_eq!(computed, 3);
assert_eq!(book.sheets()[0].get(a3).and_then(|c| c.value.as_number()), Some(41.0));
# Ok::<(), excelerate::Error>(())
```

When you change a formula rather than a value, tell the index with
`deps.note(&book, sheet, cell)`.

### Insert rows and sort; formulas follow

Grid edits rewrite every formula in the workbook that points past the edit, on
every sheet, the way Excel does. A copy rewrites the copied formulas; a sort
treats each moved row as copied.

```rust
use excelerate::edit::{self, CopyOrigin, SortKey, SortOptions};
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::{Range, Row, at};

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Sales")?;
sheet.set(at!("A1"), "Item");
sheet.set(at!("B1"), "Amount");
for (row, item, amount) in [(2, "Tea", 180.0), (3, "Coffee", 350.0), (4, "Milk", 90.0)] {
    sheet.set_at(row, 1, item)?;
    sheet.set_at(row, 2, amount)?;
}
sheet.set(at!("B6"), CellValue::formula("SUM(B2:B4)"));
book.add_sheet(sheet)?;

// 1. Two rows above row 3 (rows count from 0 here), styled like row 2.
//    The total moves from B6 to B8 and its range grows to cover them.
edit::insert_rows_with(&mut book, 0, Row::new(2).unwrap(), 2, CopyOrigin::Before)?;
let total = book.sheets()[0].get(at!("B8")).map(|c| &c.value);
assert!(matches!(total, Some(CellValue::Formula { formula, .. }) if formula == "SUM(B2:B6)"));

// 2. Sort the block by its "Amount" header, largest first. The header stays.
let options = SortOptions { header: true, ..SortOptions::default() };
let keys = [SortKey::header("Amount").descending()];
edit::sort_range_with(&mut book, 0, Range::parse("A1:B6")?, &keys, options)?;
let first = book.sheets()[0].get(at!("A2")).map(|c| &c.value);
assert_eq!(first, Some(&CellValue::text("Coffee")));
# Ok::<(), excelerate::Error>(())
```

[docs/editing.md](docs/editing.md) has the rules each edit follows, and
columns, cells, copy, move, fill and sheet renames.

### Convert between formats in memory

A server that receives a file as bytes never needs a temp file. The name is
optional and only helps when the bytes are ambiguous, as CSV is.

```rust
use excelerate::{reader, writer};
use std::io::Cursor;

// 1. Bytes in: here an Excel 97 file.
let bytes = std::fs::read("tests/fixtures/sample.xls")?;
let book = reader::read_bytes(&bytes, Some("sample.xls"))?;

// 2. Bytes out, in three formats.
let mut xlsx = Vec::new();
writer::write_xlsx_to(&book, Cursor::new(&mut xlsx))?;
let mut ods = Vec::new();
writer::write_ods_to(&book, Cursor::new(&mut ods))?;
let mut csv = Vec::new();
writer::write_csv_to(&book, 0, &mut csv, ';')?; // one sheet, your delimiter

assert!(xlsx.starts_with(b"PK") && ods.starts_with(b"PK"));
# Ok::<(), Box<dyn std::error::Error>>(())
```

### Import a CSV with `;` and a decimal comma

A CSV from a Russian or German Excel writes `1 234,50`. Say which characters
are which, and the field becomes the number 1234.5.

```rust
use excelerate::reader::{CsvOptions, FormattedNumbers, read_csv_str};
use excelerate::at;

let text = "Item;Amount\nTea;1 234,50\nCoffee;99,90\n";
let book = read_csv_str(
    text,
    &CsvOptions {
        delimiter: Some(';'),
        formatted_numbers: Some(FormattedNumbers {
            decimal: ',',
            thousands: ' ',
            keep_format: true, // remember the look as the cell's number format
        }),
        ..CsvOptions::default()
    },
);
let amount = book.sheets()[0].get(at!("B2")).and_then(|c| c.value.as_number());
assert_eq!(amount, Some(1234.5));
# Ok::<(), excelerate::Error>(())
```

Without a `delimiter`, the reader picks the one that splits the lines most
evenly, and a leading `sep=;` line wins over both.

### Ask a table by name

An Excel table has a name, and formulas can use it instead of an address:
`Sales[Amount]` keeps pointing at the column when rows are added. The engine
evaluates a formula you hand it without storing anything.

```rust
use excelerate::formula::eval::{Engine, Origin};
use excelerate::formula::value::Value;
use excelerate::{at, reader};

// 1. The fixture holds one table, "Sales", with Region, Quarter and Sales.
let book = reader::read("tests/fixtures/table.xlsx")?;
let table = &book.sheets()[0].tables[0];
assert_eq!((table.display_name.as_str(), table.range.to_string().as_str()), ("Sales", "A1:C5"));

// 2. Evaluate against it, as if typed into E1 of the first sheet.
let mut engine = Engine::new(&book);
let at = Origin::new(0, at!("E1"));
let total = engine.eval(at, "SUM(Sales[Sales])");
let north = engine.eval(at, r#"SUMIFS(Sales[Sales],Sales[Region],"North")"#);
assert!(matches!(total, Value::Number(n) if n == 465.0));
assert!(matches!(north, Value::Number(n) if n == 260.0));
# Ok::<(), excelerate::Error>(())
```

### Notes, links, a drop-down and a password

Everything on a sheet that is not a cell is a plain field of `Worksheet`.

```rust
use excelerate::model::{
    Comment, DataValidation, Hyperlink, LinkTarget, PasswordHash, SheetProtection,
    Spreadsheet, TextRun, ValidationType, Worksheet,
};
use excelerate::{Range, at, reader, writer};
use std::io::Cursor;

let mut sheet = Worksheet::new("Orders")?;
let (a1, b2) = (at!("A1"), at!("B2"));

// 1. A note on B2. The author lives on the note itself.
sheet.comments.insert(b2, Comment {
    author: "Sales".to_owned(),
    text: vec![TextRun { text: "Check the rate on the day of shipping".to_owned(), font: None }],
    visible: true, // shown all the time, not only on hover
    size: None,    // Excel's default box
});

// 2. A link on A1.
sheet.hyperlinks.push(Hyperlink {
    range: Range::new(a1, a1),
    target: LinkTarget::Outside("https://example.org/orders".to_owned()),
    display: None,
    tooltip: Some("Open the order list".to_owned()),
});

// 3. A yes/no drop-down on C2:C100.
sheet.data_validations.push(DataValidation {
    sqref: vec![Range::parse("C2:C100")?],
    kind: ValidationType::List,
    formula1: "\"yes,no\"".to_owned(),
    allow_blank: true,
    ..DataValidation::default()
});

// 4. Lock the sheet. SHA-512 with a random salt, as Excel writes it; this
//    stops editing in Excel and encrypts nothing.
sheet.protection = SheetProtection { sheet: Some(true), ..SheetProtection::default() };
sheet.protection.password = PasswordHash::new("secret");

let mut book = Spreadsheet::empty();
book.add_sheet(sheet)?;
let mut bytes = Vec::new();
writer::write_xlsx_to(&book, Cursor::new(&mut bytes))?;

let again = reader::read_bytes(&bytes, None)?;
let sheet = &again.sheets()[0];
assert_eq!(sheet.comments.len(), 1);
assert!(sheet.protection.password.as_ref().is_some_and(|p| p.verify("secret")));
# Ok::<(), excelerate::Error>(())
```

Merges, autofilters, conditional formats, print setup and the rest are in
[docs/sheet-features.md](docs/sheet-features.md).

### Your own function, with a progress callback

A workbook may call a function Excel does not have, say one your add-in
provided. Register it and formulas that use it compute; leave it out and they
answer `#NAME?`, as Excel does with the add-in missing.

```rust
use excelerate::formula::custom::CustomFunctions;
use excelerate::formula::eval::recalculate;
use excelerate::formula::value::Value;
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::progress::{Options, Progress};
use excelerate::{CellError, at};
use std::cell::Cell;

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Prices")?;
sheet.set(at!("A1"), 100.0);
sheet.set(at!("A2"), CellValue::formula("WITHVAT(A1)"));
book.add_sheet(sheet)?;

// 1. The function: arguments arrive computed, a range as `Value::Array`.
let mut functions = CustomFunctions::new();
functions.register("WITHVAT", |args| match args.first() {
    Some(Value::Number(n)) => Value::Number(n * 1.2),
    _ => Value::Error(CellError::Value),
});

// 2. A progress callback. It is `Fn`, so a counter goes in a `Cell`.
let reports = Cell::new(0);
let report = |_: Progress<'_>| reports.set(reports.get() + 1);

// 3. Both ride in one `Options`, which `read_*_with`, `write_*_with` and
//    `recalculate` all take.
let options = Options::new().with_functions(&functions).reporting(&report);
recalculate(&mut book, None, &options);

assert_eq!(book.formatted(0, at!("A2")), "120");
assert!(reports.get() > 0);
# Ok::<(), excelerate::Error>(())
```

A built-in name always stays built in: registering `SUM` does not replace it.

### See how a cell is styled

```rust
use excelerate::{at, reader};

let book = reader::read("tests/fixtures/styles.xlsx")?;
let cell = book.sheets()[0].get(at!("A1")).unwrap();

// The cell holds an id; the workbook holds the style.
let style = book.styles.get(cell.style).unwrap();
assert_eq!(style.font.name, "Arial");
assert!(style.font.bold);
assert!((style.font.size_points() - 14.0).abs() < f64::EPSILON);
println!("fill {:?}, borders {:?}", style.fill, style.borders);
# Ok::<(), excelerate::Error>(())
```

[docs/styles.md](docs/styles.md) covers colours, themes, number formats and
rich text inside a cell.

### Title, author and your own fields

What Excel shows under File > Info travels with the workbook in xlsx, xls and
ODS, along with the fields a user adds under Custom.

```rust
use excelerate::model::{PropertyValue, Spreadsheet, Worksheet};
use excelerate::{reader, writer};

let mut book = Spreadsheet::empty();
book.add_sheet(Worksheet::new("Sheet1")?)?;

// 1. The standard fields are plain options. Dates are ISO 8601 text.
book.properties.title = Some("Q3 sales".into());
book.properties.creator = Some("Ann".into());
book.properties.company = Some("Acme".into());
book.properties.created = Some("2026-09-27T10:00:00Z".into());

// 2. Your own fields keep their type: text, whole number, number, yes/no, date.
book.properties.set_custom("Department", "Sales");
book.properties.set_custom("Pages", 12_i64);

// 3. They come back from xls just as from xlsx.
let mut bytes = Vec::new();
writer::write_xls_to(&book, &mut bytes)?;
let again = reader::read_bytes(&bytes, None)?;
assert_eq!(again.properties.title.as_deref(), Some("Q3 sales"));
assert_eq!(again.properties.custom("pages"), Some(&PropertyValue::Integer(12)));
# Ok::<(), excelerate::Error>(())
```

## From JavaScript

The same crate compiled to WebAssembly, for Node and the browser:

```ts
import { Book } from "@rosperitus/excelerate";
import { readFileSync, writeFileSync } from "node:fs";

const book = Book.read(readFileSync("sales.xlsx"));
book.insertRows(0, 8, 2);                              // styled like row 7
book.sortRange(0, "A1:E500", ["-Amount"], { header: true });
book.setRangeStyle(0, "A1:E1", { font: { bold: true } });
book.recalculate();
writeFileSync("sales-sorted.xlsx", book.toXlsx());
```

The whole JS API is in [docs/wasm.md](docs/wasm.md) and the package's
[README](npm/README.md).

## Formats

| Format | Read | Write |
|---|:--:|:--:|
| xlsx (OOXML) | ✅ | ✅ |
| xls (BIFF8) | ✅ | ✅ |
| xls (BIFF5, Excel 5 and 95) | ✅ | — |
| xlsb (BIFF12) | ✅ | — |
| ods (OpenDocument) | ✅ | ✅ |
| CSV | ✅ | ✅ |
| HTML | ✅ | ✅ |
| SYLK (`.slk`) | ✅ | — |
| Gnumeric | ✅ | — |
| SpreadsheetML 2003 (`.xml`) | ✅ | — |

Every format has its own sharp edges; they are all written down in
[docs/formats.md](docs/formats.md) rather than left for you to find at 3am.

## Docs

| Page | What's in it |
|---|---|
| [Getting started](docs/getting-started.md) | install, read a file, write one, the whole loop |
| [Recipes](docs/recipes.md) | short answers: read a value, convert a file, recalculate one cell |
| [Workbook model](docs/model.md) | workbook, sheet, cell, addresses and ranges |
| [Editing the grid](docs/editing.md) | inserting and removing rows, columns and cells, formatting what you insert, copying, moving, sorting and filling, and what follows each edit |
| [Sheet features](docs/sheet-features.md) | merges, comments, tables, protection, filters, validation, print setup |
| [File formats](docs/formats.md) | what each format carries and what it drops |
| [Formulas](docs/formulas.md) | evaluation, incremental recalc, driving the engine yourself |
| [Styles and number formats](docs/styles.md) | fonts, fills, borders, format strings |
| [Long operations](docs/long-operations.md) | progress, custom functions, what each step costs |
| [Links to other workbooks](docs/external-links.md) | `[1]Sheet1!A1`, the value cache, plugging in a live file |
| [WebAssembly](docs/wasm.md) | building for Node and the browser, and the whole JS API |

## Building

```text
cargo fmt && cargo clippy --all-targets -- -D warnings && cargo test
./build-npm.sh && (cd npm && npm install && npm test && npm run typecheck)
```

## Status

Formats, styles, the formula engine and the xlsx object model (protection,
autofilters, validation, conditional formatting, print setup, tables, charts,
pictures, shapes, comments, pivot tables) are in. xls keeps formulas and styles both ways. See [docs/formats.md](docs/formats.md) for the exact line
between "modelled" and "carried".

## License

MIT - see [LICENSE](LICENSE).
