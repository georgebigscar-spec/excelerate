# Formulas

The engine parses, evaluates and caches - 518 Excel functions across math,
statistics, distributions, regression, text, dates, financial, lookup, logic,
information, database, engineering and web categories, `LAMBDA` and its kin,
`GROUPBY` and `PIVOTBY`, `GETPIVOTDATA`, and the regular expressions (`REGEXTEST`,
`REGEXEXTRACT`, `REGEXREPLACE`, which have no lookaround or backreferences).
Ten more come from a 2026 specification ([Excel_2026.md](Excel_2026.md)):
`FLATTEN`, `HAS`, `HASANY` and `HASALL` from Microsoft's Beta Channel, and
`FILLDOWN`, `PARSEJSON`, `ROLLING`, `BINS`, `TABLEJOIN`, `FUZZYLOOKUP`, which
Excel does not have.

## Recalculating a workbook

```rust
use excelerate::formula::eval::recalculate;
use excelerate::progress::Options;
# use excelerate::reader;

# let mut book = reader::read("tests/fixtures/sample.xlsx")?;
let computed = recalculate(&mut book, None, &Options::default());        // whole workbook
let one_sheet = recalculate(&mut book, Some(0), &Options::default());    // just this sheet
# let _ = (computed, one_sheet);
# Ok::<(), excelerate::Error>(())
```

The third argument carries a progress callback and any functions of your own
the workbook calls; `&Options::default()` when you want neither. See
[`progress`](https://docs.rs/excelerate/latest/excelerate/progress/) for a
callback that reports formula by formula.

A workbook of more than twenty thousand formulas is parsed and computed on up
to eight threads, in waves of formulas that read nothing of each other; the
answers are what one thread gives. Your own functions, a smaller workbook or
WebAssembly keep the single-threaded pass.

Results land in each formula cell's `cached` field, which is exactly where a
file's own saved results live - so writing the book back gives every other
reader the numbers too.

## Recalculating after an edit

Full recalc on a book with 20k formulas costs ~70 ms. When you are editing,
only what depends on the change needs to move:

```rust
use excelerate::at;
use excelerate::formula::eval::recalculate_from;
# use excelerate::reader;

# let mut book = reader::read("tests/fixtures/sample.xlsx")?;
let a1 = at!("A1");
book.sheet_mut(0).unwrap().set(a1, 99.0);
let touched = recalculate_from(&mut book, &[(0, a1)]);
# let _ = touched;
# Ok::<(), excelerate::Error>(())
```

To recompute a single formula instead - the cell itself, not the cells reading
it - use `recalculate_cell`:

```rust
use excelerate::at;
use excelerate::formula::eval::recalculate_cell;
# use excelerate::reader;

# let mut book = reader::read("tests/fixtures/sample.xlsx")?;
let b4 = at!("B4");
let was_a_formula = recalculate_cell(&mut book, 0, b4);
# let _ = was_a_formula;
# Ok::<(), excelerate::Error>(())
```

Whatever that formula reads is computed on the way, but only inside the
engine's own cache: no other cell of the workbook is written.

Dependencies come out of the formulas themselves - every reference they parse -
and the wave runs to a fixed point. An edit anywhere inside a range counts as
reading that range.

Editing in a loop? Build the dependency index once and keep it:

```rust
use excelerate::at;
use excelerate::formula::eval::Dependencies;
# use excelerate::model::CellValue;
# use excelerate::reader;

# let mut book = reader::read("tests/fixtures/sample.xlsx")?;
let mut deps = Dependencies::of(&book);

let a1 = at!("A1");
book.sheet_mut(0).unwrap().set(a1, 1.0);
deps.recalculate_from(&mut book, &[(0, a1)]);

// Changed a *formula*, not a value? The index has to be told.
let b1 = at!("B1");
book.sheet_mut(0).unwrap().set(b1, CellValue::formula("A1*100"));
deps.note(&book, 0, b1);
# Ok::<(), excelerate::Error>(())
```

Numbers from a real book with 19,811 formulas: full recalc 635 ms, building the
index 16 ms, 50 edits in one batch -> 184 formulas in 13 ms. The same 50 edits
one at a time: 263 ms. Batch your edits; the pass is per call, not per cell.

## Evaluating an expression yourself

No cells involved, no book mutated:

```rust
use excelerate::at;
use excelerate::formula::{Value, eval::{Engine, Origin}};
use excelerate::model::Spreadsheet;

let book = Spreadsheet::new();
let mut engine = Engine::new(&book);

let answer = engine.eval(Origin::new(0, excelerate::at!("A1")), "ROUND(2/3, 4)");
assert_eq!(answer, Value::Number(0.6667));
# Ok::<(), excelerate::Error>(())
```

`Origin` says which cell the formula is speaking from - relative references and
`ROW()`/`COLUMN()` need it. `Engine` caches within its lifetime, so evaluating
a thousand expressions against the same book reuses everything it already
computed.

## Parsing without evaluating

```rust
use excelerate::formula::{Expr, parse};

let expr = parse("SUM(Sheet2!A1:A9)*2")?;
match expr {
    Expr::Binary(..) => {}
    _ => unreachable!(),
}
# Ok::<(), excelerate::Error>(())
```

The parser handles literals, references, ranges, calls, array constants and all
three reference operators - intersection (a space), union (`,`) and the span
operator (`:` between areas, so `A1:A2:B1` is `A1:B2`). Useful for linting
formulas, finding every cell a sheet reads, or rewriting references.

## Values

```text
pub enum Value {
    Blank,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(CellError),
    Array(Arc<Vec<Vec<Value>>>),
    Lambda(Arc<Lambda>),
}
```

Excel's own coercion rules apply throughout, including the comparison order
that trips people up: numbers < text < `FALSE` < `TRUE`. A formula that returns
an array shows its top-left value in a cell - there is no spill range here, and
one cell is one cell. A reference to that cell reads the value it shows;
`A1#` (`ANCHORARRAY`) reads the array behind it.

An array is shared, not owned: a range is read once and handed to every
formula that asks for it, so `INDEX(Data, ROW(), 1)` down forty thousand rows
does not copy `Data` forty thousand times. Make one with `Value::array(rows)`;
take the rows out with `Rc::unwrap_or_clone`.

## About that cached result

A formula cell carries the result whoever saved the file computed. That cache
is worth exactly as much as the engine that wrote it. A quick tell: if an xlsx
package has no `xl/calcChain.xml`, Excel was not the last thing to save it
(Excel always writes that part for a book with formulas), and the cached values
may be stale or plain wrong.

So: trust it for display, recalculate before you rely on it.

## Named ranges

Defined names resolve during evaluation. A sheet-scoped name beats a
workbook-scoped one of the same name - that is how two sheets can give one name
two meanings - and a name that refers to itself yields `#REF!` instead of
looping. `INDIRECT("Sales")` follows a name that stands for a reference; one
that stands for a constant is `#REF!`, as in Excel.

## Errors

Errors propagate the way they do in Excel: anything touching `#DIV/0!` becomes
`#DIV/0!`. A formula that fails to parse evaluates to `#NAME?`, which is also
what Excel shows for a function it does not know.

There is no panic path here. A formula 200 levels deep, a circular reference, a
range covering a million cells - each has a defined answer (`#REF!` for a
cycle, `#VALUE!` for an oversized range) rather than a stack overflow. A chain
of formulas reading formulas has no limit of its own: past a hundred links the
engine sets the deepest cell aside, computes it from the top and carries on.

## Goal Seek

`what_if::goal_seek` finds the value of an input cell that makes a formula
compute a target: Excel's What-If Analysis - Goal Seek. It stops after the
workbook's Maximum Iterations or within its Maximum Change (100 and 0.001
unless the file says otherwise), leaves the best value in the input cell and
recomputes what reads it.

```rust
use excelerate::at;
use excelerate::formula::what_if::goal_seek;
use excelerate::model::{CellValue, Spreadsheet, Worksheet};

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Loan")?;
sheet.set(at!("A1"), 0.01);                        // monthly rate
sheet.set(at!("B1"), CellValue::formula("PMT(A1,36,-10000)"));
book.add_sheet(sheet)?;

// Which rate makes the payment 300?
let found = goal_seek(&mut book, (0, at!("B1")), 300.0, (0, at!("A1")))?;
assert!(found.converged);
assert!((found.value - 0.0042).abs() < 0.0001);
# Ok::<(), excelerate::Error>(())
```

## Checking data validation

A sheet's `data_validations` are kept as the file wrote them;
`formula::validation` applies them. `check` says whether a value would be
accepted in a cell and hands back the rule it breaks, with the error title
and message Excel would show; `invalid_cells` is Circle Invalid Data;
`list_items` gives the choices of a drop-down list.

```rust
use excelerate::at;
use excelerate::formula::validation::{check, invalid_cells};
use excelerate::model::{
    CellValue, DataValidation, Spreadsheet, ValidationOperator, ValidationType, Worksheet,
};
use excelerate::Range;

let mut book = Spreadsheet::empty();
let mut sheet = Worksheet::new("Form")?;
sheet.set(at!("A2"), 150.0);
sheet.data_validations.push(DataValidation {
    sqref: vec![Range::parse("A1:A10")?],
    kind: ValidationType::Whole,
    operator: ValidationOperator::Between,
    formula1: "0".into(),
    formula2: "120".into(),
    allow_blank: true,
    ..DataValidation::default()
});
book.add_sheet(sheet)?;

assert!(check(&book, 0, at!("A1"), &CellValue::Number(42.0)).is_none());
assert!(check(&book, 0, at!("A1"), &CellValue::Number(4.5)).is_some());
assert_eq!(invalid_cells(&book, 0), vec![at!("A2")]);
# Ok::<(), excelerate::Error>(())
```

A rule's formulas are written for the first cell of its range and move with
the cell checked, as in Excel, so `$B$1` is a fixed bound and `B1` the cell
beside each one.
