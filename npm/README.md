# excelerate

Read, write and recalculate spreadsheets in Node and the browser - xlsx, xls,
xlsb, ods, csv, html and more - with a formula engine of 518 Excel functions.
Rust compiled to WebAssembly: no native module to build, no Python, no
headless Office.

```text
npm install @rosperitus/excelerate
```

TypeScript declarations ship with the package. Only need to read? The
`@rosperitus/excelerate-reader` package has the reading half of the API in a
wasm file of 1.4 MB instead of 3.9 MB (0.6 MB against 1.5 MB gzipped).

The examples use `import`: a project needs `"type": "module"` in its
`package.json` for that (`npm init -y` writes `"commonjs"`), or `require`
works the same - `const { Book } = require("@rosperitus/excelerate")`.

## Quick start

```ts
import { Book } from "@rosperitus/excelerate";
import { readFileSync, writeFileSync } from "node:fs";

// The format is detected from the bytes. The file name is optional and only
// settles the ambiguous cases (a .csv against an .html).
const book = Book.read(readFileSync("report.xlsx"), "report.xlsx");

console.log(book.sheetNames());       // ["Sheet1", "Data"]
console.log(book.get(0, "B4"));       // number | string | boolean | null

book.set(0, "B4", 42);
book.recalculateFrom(0, "B4");        // only what depends on B4

writeFileSync("out.xlsx", book.toXlsx());
```

Build a report from scratch - values, a formula, a styled header, a sort and
a total row:

```ts
import { Book } from "@rosperitus/excelerate";
import { writeFileSync } from "node:fs";

const book = new Book();
book.setRange(0, "A1", [
  ["Item", "Qty", "Price", "Sum"],
  ["Bolt", 40, 0.25, "=B2*C2"],
  ["Nut", 100, 0.1, null],
  ["Washer", 250, 0.05, null],
]);
book.fillDown(0, "D2:D4");                           // =B3*C3, =B4*C4
book.setRangeStyle(0, "A1:D1", { font: { bold: true }, fill: { pattern: "solid", foreground: "#FFDDEBF7" } });
book.setRangeStyle(0, "C2:D5", { numberFormat: "#,##0.00" });

book.recalculate();                                  // the sort reads cached values
book.sortRange(0, "A1:D4", ["-Sum"], { header: true });
book.insertRows(0, 5, 1);                            // styled like the row above
book.setRange(0, "A5", [["Total", null, null, "=SUM(D2:D4)"]]);

book.recalculate();
book.freezePanes(0, 1, 0);
writeFileSync("report.xlsx", book.toXlsx());
```

Continue a series the way the fill handle does:

```ts
import { Book } from "@rosperitus/excelerate";

const book = new Book();
book.setRange(0, "F1", [["Jan", "Кв1", 1], [null, null, 3]]);
book.fillSeries(0, "F1:H6", "down");   // Feb..Jun, Кв2..Кв6, 5, 7, 9, 11
```

Clean up imported data, then ask what input reaches a target:

```ts
import { Book } from "@rosperitus/excelerate";

const book = new Book();
book.setRange(0, "A1", [["Ann;41"], ["Bob;37"], ["ann;41"]]);
book.textToColumns(0, "A1:A3", { delimiters: ";" });   // Ann | 41 ...
book.removeDuplicates(0, "A1:B3");                      // { removed: 1, remaining: 2 }

book.setRange(0, "D1", [[0.01, "=PMT(D1,36,-10000)"]]);
const { value } = book.goalSeek(0, "E1", 300, "D1");    // the monthly rate: 0.0042...

book.addDataValidation(0, { sqref: "B1:B100", type: "whole", formula1: "18", formula2: "99" });
book.checkValidation(0, "B5", 12);                      // { style: "stop", ... }
```

## In the browser

The package carries a second build for browsers and bundlers (Vite, webpack,
esbuild), and the same import picks it there. The one difference from Node:
the module has to be started first, which fetches the `.wasm`:

```ts
import init, { Book } from "@rosperitus/excelerate";

await init();
const file = document.querySelector("input[type=file]") as HTMLInputElement;
const book = Book.read(new Uint8Array(await file.files![0].arrayBuffer()));
const blob = new Blob([book.toXlsx()]);
```

The `.wasm` is fetched from beside the module, so the server should send it as
`application/wasm`; then the browser compiles it as it streams. Without a
bundler, an import map pointing at `@rosperitus/excelerate/web/index.js` does
the same. `@rosperitus/excelerate/web` names the browser build explicitly,
which also runs in Node given the bytes: `await init({ module_or_path: bytes })`.

## Objects and chains

`Book` is one flat class addressed by sheet index. Over it sit `Workbook`,
`Sheet`, `Table` and `SheetRange`: a workbook hands out sheets, a sheet hands
out tables and ranges, and every call that changes something returns its
object, so calls chain.

```ts
import { Workbook } from "@rosperitus/excelerate";
import { writeFileSync } from "node:fs";

const wb = new Workbook();
const sales = wb
  .addSheet("Sales")                       // a new workbook's first sheet, renamed
  .addTableFromData("Sales", "A1", [
    ["Region", "Manager", "Amount"],
    ["North", "Ivanov", 120],
    ["South", "Petrov", 340],
    ["West", "Kozlov", 210],
    ["South", "Smirnov", 95],
  ])
  .set(1, "Amount", 125)                   // data row 1, by header
  .addFilter("Region", { values: ["South", "West"] })
  .addFilter("Amount", { custom: [{ op: ">", value: 100 }] });

sales.records({ visible: true });          // [{ Region: "South", Manager: "Petrov", Amount: 340 }, ...]
sales.range().visibleValues();             // the same with the header, as a grid
sales.sheet.style("A1:C1", { font: { bold: true } }).width("B", 14);
writeFileSync("sales.xlsx", wb.toXlsx());
```

A filter is Excel's own: the criteria go into the table's `<autoFilter>`, and
the rows they reject are hidden in the file, so Excel opens it already
filtered. Criteria on several columns all have to hold. `{ values }` keeps
the rows whose cell shows one of them (`blank: true` adds the empty ones),
`{ custom }` takes one or two comparisons (`=`, `<>`, `>`, `>=`, `<`, `<=`;
`*` and `?` in `=` and `<>`; `and: true` to need both), `{ top: n }` and
`{ bottom: n }` the extremes, `percent: true` for a share. `clearFilter()`
takes them off.

Every `Book` method that takes the sheet first is a `Sheet` method without
it, and every one that takes no sheet a `Workbook` method: what `Book`
answers with nothing returns the object, so it chains, and the rest return
their answer. So the whole flat API is there:

```ts
const sheet = wb.sheet("Sales")
  .setColumnWidth(2, 14)
  .setRowHidden(5, true)
  .insertRows(8, 2)
  .setCellStyle("A1", { font: { bold: true } });
sheet.getRange("A1:C3");                 // a query answers: CellGrid
sheet.columnWidth(2);                    // 14
wb.setDefinedName("Total", "Sales!$C$9").definedNames();
```

On top of that a few have their own shape: `write`, `range`, `style`,
`width`, `freeze`, `addTable`, `addTableFromData`, `table`, `rename`,
`activate`, `recalculate` and `remove` on a sheet; `sheet`, `addSheet` and
`Workbook.readCsv` on the workbook.

The objects are handles - a sheet index, a table name, an address - and work
through `workbook.book`. A
sheet removed or moved through the `Book` moves the indexes under them, as it
does for any index. `Workbook.read(bytes)` opens a file; `Book` and its flat
API are unchanged.

## API

```ts
type CellValue = number | string | boolean | null;
type CellGrid = CellValue[][];
type CopyOrigin = "before" | "after" | "none";   // whose formatting new rows take
type SortKeys = (number | string)[];             // 2 = column B, "-Amount" = header, largest first
type Progress = { stage: "reading" | "writing" | "recalculating"; done: number;
                  total: number | null; what: string; fraction: number | null };

class Book {
  constructor();
  static read(bytes: Uint8Array, name?: string | null, maxExpanded?: number | null, onProgress?: (p: Progress) => void): Book;
  static readCsv(bytes: Uint8Array, options: CsvOptions): Book;   // shape stated, not guessed

  // Sheets
  sheetNames(): string[];
  sheetIndex(name: string): number | undefined;   // case-insensitive, as Excel compares
  addSheet(title: string): number;
  renameSheet(sheet: number, title: string): void;
  activeSheet(): number;
  setActiveSheet(sheet: number): void;
  removeSheet(sheet: number): void;               // references to it become #REF!
  cellCount(sheet?: number | null): number;
  usedRange(sheet: number): string | undefined;   // "A1:D9", by walking the cells
  usedRangeHint(sheet: number): string | undefined; // the same, never walking; columns may be wider after a removal

  // Cells, by address
  get(sheet: number, address: string): CellValue;
  set(sheet: number, address: string, value: CellValue): void;
  clear(sheet: number, address: string): void;
  getFormula(sheet: number, address: string): string | undefined;
  getFormatted(sheet: number, address: string): string;   // through the number format
  getRange(sheet: number, range: string): CellGrid;       // one call, not one per cell
  setRange(sheet: number, at: string, values: CellGrid): void;
  getRangeStyles(sheet: number, range: string): RangeStyles;  // { styles, grid } - each style once
  setRangeStyles(sheet: number, at: string, styles: RangeStylesPatch): void;

  // Cells, by 1-based row and column - no address to build and re-parse
  getAt(sheet: number, row: number, column: number): CellValue;
  setAt(sheet: number, row: number, column: number, value: CellValue): void;
  clearAt(sheet: number, row: number, column: number): void;
  getFormulaAt(sheet: number, row: number, column: number): string | undefined;
  getFormattedAt(sheet: number, row: number, column: number): string;
  getRangeAt(sheet: number, row: number, column: number, rows: number, columns: number): CellGrid;
  setRangeAt(sheet: number, row: number, column: number, values: CellGrid): void;
  getRangeStylesAt(sheet: number, row: number, column: number, rows: number, columns: number): RangeStyles;
  setRangeStylesAt(sheet: number, row: number, column: number, styles: RangeStylesPatch): void;
  recalculateCellAt(sheet: number, row: number, column: number): boolean;
  recalculateFromAt(sheet: number, row: number, column: number): number;

  // Layout
  cellIndent(sheet: number, address: string): number;     // indent steps, 0 when none
  cellIndentAt(sheet: number, row: number, column: number): number;
  cellBold(sheet: number, address: string): boolean;      // the one flag, without the whole style
  cellBoldAt(sheet: number, row: number, column: number): boolean;
  rowLevel(sheet: number, row: number): number;           // outline depth, 0 when ungrouped
  columnLevel(sheet: number, column: string): number;     // by letters: "C"
  rowHidden(sheet: number, row: number): boolean;
  sheetVisibility(sheet: number): "visible" | "hidden" | "veryHidden";
  mergedRanges(sheet: number): string[];                  // ["A1:C1", ...]
  mergedRangesAt(sheet: number): Uint32Array;             // [r1, c1, r2, c2] per area
  merge(sheet: number, range: string): void;
  unmerge(sheet: number, range: string): boolean;
  columnWidth(sheet: number, column: number): number | undefined; // characters
  rowHeight(sheet: number, row: number): number | undefined;      // points
  setColumnWidth(sheet: number, column: number, width?: number): void; // undefined: back to the default
  setRowHeight(sheet: number, row: number, height?: number): void;
  setColumnHidden(sheet: number, column: number, hidden: boolean): void;
  setRowHidden(sheet: number, row: number, hidden: boolean): void;
  setSheetVisibility(sheet: number, state: "visible" | "hidden" | "veryHidden"): void;
  getRowAt(sheet: number, row: number, formatted?: boolean): SheetRow;

  // The grid itself: formulas everywhere in the workbook follow the cells they
  // read, and so do merges, links, validations, tables and drawings
  insertRows(sheet: number, at: number, count: number, copyOrigin?: CopyOrigin): void;
  removeRows(sheet: number, at: number, count: number): void;
  insertColumns(sheet: number, at: number, count: number, copyOrigin?: CopyOrigin): void;  // 1-based
  removeColumns(sheet: number, at: number, count: number): void;
  // Many places in one pass: [at, count] pairs, numbered as before the call
  insertRowsMany(sheet: number, places: [number, number][], copyOrigin?: CopyOrigin): void;
  removeRowsMany(sheet: number, places: [number, number][]): void;
  insertColumnsMany(sheet: number, places: [number, number][], copyOrigin?: CopyOrigin): void;
  removeColumnsMany(sheet: number, places: [number, number][]): void;
  copyRange(sheet: number, range: string, to: string, toSheet?: number): void;   // formulas rewritten
  moveRange(sheet: number, range: string, to: string, toSheet?: number): void;   // formulas kept, references follow
  insertCells(sheet: number, range: string, shift: "down" | "right", copyOrigin?: CopyOrigin): void;
  removeCells(sheet: number, range: string, shift: "up" | "left"): void;
  moveSheet(from: number, to: number): void;

  // Sorting and filling, as Excel's Data - Sort, Ctrl+D and the fill handle
  sortRange(sheet: number, range: string, keys: SortKeys, options?: SortRangeOptions): void;
  sortTable(name: string, keys: SortKeys): void;
  fillDown(sheet: number, range: string): void;
  fillRight(sheet: number, range: string): void;
  fillSeries(sheet: number, range: string, direction: "down" | "right" | "up" | "left"): void;

  // Data tools, as Excel's Remove Duplicates, Text to Columns and Data Validation
  removeDuplicates(sheet: number, range: string, options?: RemoveDuplicatesOptions): DuplicatesRemoved;
  textToColumns(sheet: number, range: string, options: TextToColumnsOptions): number;
  addDataValidation(sheet: number, rule: NewDataValidation): void;
  checkValidation(sheet: number, address: string, value: CellValue): ValidationProblem | null;
  invalidCells(sheet: number): string[];                // Circle Invalid Data
  validationList(sheet: number, address: string): string[];   // a drop-down's choices

  // What is on a sheet besides cells
  comments(sheet: number): SheetComment[];
  hyperlinks(sheet: number): SheetHyperlink[];
  tables(sheet: number): SheetTable[];
  charts(sheet: number): SheetChart[];          // with the chart and plot area fill and outline
  images(sheet: number): SheetImage[];          // without the bytes
  imageData(sheet: number, index: number): Uint8Array;
  shapes(sheet: number): SheetShape[];          // rotation, flips, fill, line, font as Excel shows them
  setShapeFormat(sheet: number, index: number, patch: ShapeFormatPatch): void;   // fill, line, font, rotation, flips
  setChartFormat(sheet: number, index: number, patch: ChartFormatPatch): void;   // chart and plot area fill and outline

  // Style: read whole, written as a patch over what the cell has
  cellStyle(sheet: number, address: string): CellStyle;   // numberFormat, font, fill, borders, alignment
  cellStyleAt(sheet: number, row: number, column: number): CellStyle;
  setCellStyle(sheet: number, address: string, patch: CellStylePatch): void;
  setCellStyleAt(sheet: number, row: number, column: number, patch: CellStylePatch): void;
  setRangeStyle(sheet: number, range: string, patch: CellStylePatch): void;
  getRichText(sheet: number, address: string): TextRun[] | null;   // [{ text, font }]
  getRichTextAt(sheet: number, row: number, column: number): TextRun[] | null;
  setRichText(sheet: number, address: string, runs: TextRun[]): void;
  setRichTextAt(sheet: number, row: number, column: number, runs: TextRun[]): void;

  // Notes, links, tables
  setComment(sheet: number, address: string, author: string, text: string): void;
  removeComment(sheet: number, address: string): boolean;
  setHyperlink(sheet: number, range: string, target: string, inside?: boolean, display?: string, tooltip?: string): void;
  removeHyperlink(sheet: number, range: string): boolean;
  addTable(sheet: number, name: string, range: string, headerRow?: boolean): void;
  removeTable(sheet: number, name: string): boolean;
  setTableFilter(name: string, column: string | number, criteria: TableFilter | null): void;  // hides the rows that fail

  // The saved view
  sheetView(sheet: number): SheetViewInfo;
  freezePanes(sheet: number, rows: number, columns: number): void;   // (0, 1, 0) pins the header, (0, 0, 0) unfreezes
  setZoom(sheet: number, percent?: number): void;
  setShowGridLines(sheet: number, show: boolean, headers?: boolean): void;

  // Names
  definedNames(): WorkbookName[];
  setDefinedName(name: string, formula: string, sheet?: number): void;
  removeDefinedName(name: string, sheet?: number): boolean;

  // File > Info: title, author, company, and fields of your own
  documentProperties(): DocumentProperties;
  setDocumentProperties(patch: DocumentPropertiesPatch): void;  // null clears a field

  // Rules a file states
  dataValidations(sheet: number): SheetValidation[];
  conditionalFormats(sheet: number): SheetConditionalFormat[];
  autoFilter(sheet: number): SheetAutoFilter | undefined;
  pivotTables(sheet: number): SheetPivotTable[];
  arrayFormulas(sheet: number): string[];       // ["B2:B4"]
  externalBooks(): (string | undefined)[];

  // Locks - they stop editing, they do not encrypt
  protectSheet(sheet: number, password?: string): void;
  unprotectSheet(sheet: number): void;
  sheetProtection(sheet: number): ProtectionInfo;
  verifySheetPassword(sheet: number, password: string): boolean;
  protectWorkbook(password?: string, windows?: boolean): void;
  unprotectWorkbook(): void;
  workbookProtection(): ProtectionInfo;

  // Calculation
  evaluate(sheet: number, address: string, formula: string): CellValue;
  recalculate(sheet?: number | null, onProgress?: (p: Progress) => void): number;
  recalculateCell(sheet: number, address: string): boolean;
  recalculateFrom(sheet: number, address: string): number;
  recalculateFromMany(sheet: number, addresses: string[]): number;
  goalSeek(sheet: number, set: string, target: number, changing: string, options?: GoalSeekOptions): GoalSeekResult;

  // Functions of your own, for names no built-in claims
  registerFunction(name: string, fn: (...args: (CellValue | CellGrid)[]) => CellValue | CellGrid): void;
  unregisterFunction(name: string): boolean;
  registeredFunctions(): string[];

  // Output
  toXlsx(onProgress?: (p: Progress) => void): Uint8Array;
  toOds(): Uint8Array;
  toXls(): Uint8Array;
  toHtml(sheet?: number | null, fragment?: boolean | null): string;
  toCsv(sheet: number, options?: CsvOptions): string;   // { delimiter: ";" } for a delimiter of your own

  free(): void;
}
```

Every method that can fail throws an `Error` with the reason - a bad address, a
missing sheet, a truncated package.

## Notes worth reading once

- **`free()` / `using`.** The workbook lives in wasm memory. Node's GC will get
  to it eventually, but in a loop over many files, release it deliberately:
  `book.free()`, or `using book = Book.read(...)` with explicit resource
  management.
- **Numbers or addresses, your call.** Every cell method has an `...At` twin
  taking 1-based `row` and `column` - the numbers `ROW()` and `COLUMN()`
  return. Looping over a grid, they save building an address string only to
  have it parsed straight back.
- **Move rectangles, not cells.** `getRange`/`setRange` cross the wasm
  boundary once for the whole block; a loop of `get` crosses once per cell.
- **Batch your edits.** `recalculateFromMany(sheet, ["A1", "B7", ...])` runs one
  pass for the whole batch. Calling `recalculateFrom` in a loop runs one pass
  per cell - on a 20k-formula workbook that is roughly 20x the work.
- **Big files.** Zip expansion is capped at 512 MB to stop a zip bomb. Real
  workbooks do exceed it (a 100 MB package can expand to 560 MB), so pass a
  larger `maxExpanded` third argument to `Book.read` when you know the source.
- **Editing the grid.** `insertRows` and its three siblings move the whole
  workbook, not one sheet: `$A$5` becomes `$A$6`, a range that lost cells
  narrows, a deleted cell reads `#REF!`, and drawings and tables follow. The
  cached result of a rewritten formula is dropped, so recalculate after.
  An insert that would push a filled cell or a merge off the last row or
  column throws and leaves the book as it was, as Excel refuses it.
- **Many rows at once.** Every edit walks the whole workbook, so a loop of
  `insertRows` costs a walk per row. `insertRowsMany(0, [[5, 1], [9, 1],
  [14, 1]])` does it in one: places are numbered as the sheet was before the
  call, in any order. Subtotal rows under 4 000 groups of 40 000 rows took
  73 s one at a time and take 29 ms this way. `removeRowsMany` and the column
  pair work the same.
- **Style is a patch.** `setCellStyle(0, "A1", { font: { bold: true } })` keeps
  the number format the cell had. Equal styles share one entry in the
  workbook's table, so painting a column costs one style, not one per cell -
  and `setRangeStyle` is one call for the whole rectangle.
- **Copy rewrites, move does not.** `copyRange` rewrites the formulas it
  carries, as if they had been written where they land; `moveRange` keeps them
  pointing at the same cells and rewrites the formulas elsewhere that read the
  moved ones. That is Ctrl+C against Ctrl+X, and it is the whole difference.
- **Inserted rows look like their neighbours.** `insertRows` copies the row
  above's cell styles and height, as Excel does. Pass `"after"` for the row
  below, `"none"` for a blank row.
- **Sort keys are numbers or names.** `sortRange(0, "A1:D99", [3, -2])` sorts
  by column C, then by B largest first; with `{ header: true }` the same reads
  `["Price", "-Qty"]`. Formulas sort by their cached value, so recalculate
  first if you have edited the inputs. `sortTable("Sales", ["-Amount"])`
  needs no range at all.
- **Read and write formatting by the block.** `getRangeStyles` returns each
  distinct style once and a grid of indexes into them; `setRangeStyles` takes
  the same shape back, so copying a block's look is two calls.
- **Charts, pictures and shapes travel byte for byte; their look is a patch.**
  `charts`, `images` and `shapes` describe what a sheet carries, and an
  untouched object goes back into the file as it came. A shape's `fill`,
  `line` and `font` are what Excel shows, its style filling in what the shape
  leaves out, and a colour comes resolved through the theme:
  `{ type: "solid", color: "#FF4472C4" }`. `setShapeFormat(0, 2, { fill: {
  type: "solid", color: "accent2" }, line: { width: 2 }, rotation: 15 })`
  changes only what it names - `null` gives a fill or line back to the style -
  and only those elements are rewritten; effects, other runs and dashes stay.
  A colour to write is `#RRGGBB`, `#AARRGGBB` or a theme name (`accent1`,
  `tx1`, `bg2`) that follows the theme. `setChartFormat` does the same for a
  chart's chart area (`format`) and plot area (`plotFormat`).
- **Functions of your own.** `book.registerFunction("MYTOTAL", (range) =>
  range.flat().reduce((a, b) => a + (b ?? 0), 0))` - arguments arrive
  evaluated, a range as a grid. A built-in name stays the built-in's, an
  unregistered one reads `#NAME?`, and a function that throws gives
  `#VALUE!` rather than an exception out of the formula.
- **Progress.** `Book.read`, `toXlsx` and `recalculate` take a callback last;
  it gets `{ stage, done, total, what, fraction }` per sheet, part or formula.
  `total` is `null` while a read has not reached the workbook part yet.
- **Cached results.** A formula cell read from a file carries whatever the app
  that saved it computed. Call `recalculate()` before relying on those numbers.

## Formats

| Format | Read | Write |
|---|:--:|:--:|
| xlsx | ✅ | ✅ (`toXlsx`) |
| xls (BIFF8) | ✅ | ✅ (`toXls`) |
| ods | ✅ | ✅ (`toOds`) |
| CSV | ✅ | ✅ (`toCsv`) |
| HTML | ✅ | ✅ (`toHtml`) |
| SYLK, Gnumeric, SpreadsheetML 2003 | ✅ | - |

SYLK, Gnumeric and SpreadsheetML are read-only in the Rust crate too - there is
nothing to bind.

`toXls` compiles formulas to BIFF8 tokens with their cached results; what
BIFF8 cannot hold (a structured reference, a row past 65,536) goes out as its
value. Call `recalculate()` first so those values are current.

## Development

This folder is both the published package's readme and the workspace that
builds and exercises it.

```
../build-npm.sh                # wasm-pack --release --target nodejs -> npm/pkg
npm install && npm test        # tests against the freshly built package
npm start                      # example.js: build a book, calculate, round-trip
npm run ts                     # typescript/basic.ts, no build step
npm run typecheck              # tsc --strict over the TypeScript examples
npm run bench [iterations]     # parse timings, median and best
../pack-npm.sh                 # pkg -> pkg-publish, then: cd pkg-publish && npm publish --otp=...
```

| Path | What it is |
|---|---|
| `pkg/` | build output for Node (git-ignored); `pkg-web/` and `pkg-bundler/` for the other targets |
| `pkg-publish/` | `pack-npm.sh`: a copy of `pkg/` renamed to `@rosperitus/excelerate`, the folder `npm publish` runs in (git-ignored) |
| `example.js`, `test.js`, `bench.js` | JavaScript examples, tests, benchmark |
| `typescript/` | TypeScript examples, run from here: `node typescript/basic.ts` (Node 22.6+). No node_modules of their own - they resolve the package from this folder |
| `browser/` | the same API in a page: `index.html` over `pkg-web/` |
| `files/` | workbooks the tests read (`gen.xlsx`, `grouped.xlsx`, committed) and whatever the benchmark and examples read and write (git-ignored) |

## License

MIT
