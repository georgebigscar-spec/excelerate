# Changelog

## Unreleased

### Fixed

- A sheet's scenarios (`<scenarios>`, What-If Analysis - Scenario Manager)
  were dropped on read, so a workbook lost them on a round trip. They are now
  `Worksheet::scenarios` (`model::scenario`) and written back in the
  schema's place.
- A grouped sheet was written without `outlineLevelRow` and
  `outlineLevelCol` in `<sheetFormatPr>`, which Excel draws the outline's
  level buttons from: it opened with nothing to collapse the groups by.

### Changed

- `zip` is built with Deflate alone (`default-features = false`,
  `deflate-flate2-zlib-rs`). Its defaults pulled in zstd, bzip2, lzma, ppmd,
  deflate64, zopfli and zip's own encryption - 28 crates, one of them C
  (`zstd-sys`), so a WebAssembly build needed clang. Every format here is
  Stored or Deflate, and Deflate stays on zlib-rs, so nothing reads or
  writes differently. A package written with one of the dropped methods
  now fails to open with `Error::Xlsx` and the like instead of opening.

### Added

- Data tools ported from [gridcraft](https://github.com/storytold/gridcraft)
  (MIT OR Apache-2.0):
  - `edit::remove_duplicates`: Excel's Remove Duplicates. Rows of a range are
    compared by the text their cells show, without regard to case, on all
    columns or the ones given; the first is kept, the rest close up inside
    the range, and moved formulas are rewritten as a sort rewrites them.
    Returns `Deduplicated { removed, remaining }`.
  - `edit::text_to_columns` with `SplitBy`: Text to Columns, by delimiters
    (with a text qualifier and consecutive delimiters as one) or by fixed
    widths. Pieces are typed as CSV fields are; a split that would run past
    the last column writes nothing and is `Error::WouldPushOffSheet`.
  - `formula::what_if::goal_seek` and `goal_seek_with`: Goal Seek, secant
    steps with bisection inside a sign change, stepping back out of inputs
    where the model errors. Limits come from the workbook's `iterateCount`
    and `iterateDelta` (`SeekLimits::of`), or are given. The best value is
    left in the input cell and what reads it is recomputed. The new
    `Error::GoalSeek` reports a set cell without a formula or an input cell
    with one.
  - `formula::validation`: `check` a value against the data validation rule
    of a cell, `invalid_cells` of a sheet (Circle Invalid Data), `rule_at` and
    `list_items` for drop-downs. Every rule type is applied, and a rule's
    formulas move from the first cell of its range to the cell checked.
- In JS: `removeDuplicates`, `textToColumns`, `goalSeek`, `checkValidation`,
  `invalidCells` and `validationList` on `Book` and `Sheet`, with their option
  and result types. `addDataValidation` adds a rule from JS, which could
  only read them before; a type or operator it does not know is an error
  rather than a silent default.
- More of Excel's Data tab, ported from gridcraft (MIT OR Apache-2.0):
  - The outline: `edit::group_rows`, `ungroup_rows`, `group_columns`,
    `ungroup_columns`, `collapse_rows`, `expand_rows` (and for columns),
    `show_outline_level` for the level buttons, `row_groups` and
    `column_groups`. `edit::subtotal` is Data - Subtotal: a `SUBTOTAL` row
    under each group of a sorted list, a grand total, and the outline.
  - `formula::conditional::ConditionalFormatting` applies a sheet's
    conditional formatting: what each cell gets from the matching rules, by
    priority with stop-if-true, and its colour scale, data bar and icon.
    `write_html` now shows it.
  - `edit::flash_fill` and `FlashFill::learn`: Flash Fill from examples, with
    several source columns, cuts inside a word and a header stepped over.
  - `edit::refresh_pivot` and `refresh_all_pivots` lay a pivot table's
    report out in its cells from its source, in tabular form, for readers
    without a pivot engine; the file still asks Excel to rebuild it.
  - `formula::what_if::add_scenario`, `show_scenario`, `scenario_summary`,
    and `data_table` for one- and two-input data tables (filled with
    values: `TABLE()` is not modelled). `edit::advanced_filter` applies a
    criteria range as Excel reads it, in place or copied, unique records
    only if asked; `edit::consolidate` combines ranges by label or position.
  - `formula::solver`: Solver with Simplex LP (branch and bound for int and
    bin) and GRG Nonlinear. `save_model` and `saved_model` keep the model in
    the hidden `solver_*` names Excel's add-in uses. The Evolutionary method
    is not included.

## 0.15.0

### Added

- Workbooks saved with a password to open: the agile encryption of Excel 2010
  and later, for xlsx, xlsm and xlsb, is decrypted by `read_bytes` and
  `read`. The password goes in `Options::password`; without one the reader
  tries the password Excel uses for a workbook encrypted only to be opened
  read-only. A password that does not open it is the new
  `Error::WrongPassword`, an encryption this does not read (Excel 2007's
  standard one, RC4 in xls) the new `Error::Encrypted`. In JS, the fifth
  argument of `Book.read`. `reader::encryption` has `decrypt` and
  `is_encrypted` on their own.
- Rows and columns inserted or removed at many places in one pass:
  `edit::{insert_rows_many, remove_rows_many, insert_columns_many,
  remove_columns_many}`. Places are given as the sheet is before the call,
  in any order, and the result is that of the single edits made from the
  far end; removed blocks that overlap are merged. Insertions take a
  `CopyOrigin` for each block. Every pass of an edit walked the whole
  workbook, so subtotals over 40 000 rows and 4 000 groups took 75 seconds
  one row at a time; the batch does it in 30 ms. The single edits are now
  a batch of one. In JS: `insertRowsMany`, `removeRowsMany`,
  `insertColumnsMany` and `removeColumnsMany`, places as `[at, count]`
  pairs, 1-based.
- In JS, `shapes` gives each shape's `rotation` (degrees), `flipH`, `flipV`
  and the `fill`, `line` and `font` Excel shows, and `charts` the
  `format` and `plotFormat` of the chart and plot area. Drawing colours come
  resolved through the workbook theme as `#AARRGGBB`.
- In JS, `setShapeFormat` and `setChartFormat` change that look with a patch
  in the same shape: a field left out stays, `null` gives a fill or outline
  back to the style, and a colour is `#RRGGBB`, `#AARRGGBB` or a theme name
  such as `accent1`.
- Ten functions from the 2026 specification (`docs/Excel_2026.md`):
  `FLATTEN`, `HAS`, `HASANY` and `HASALL` for nested arrays, from
  Microsoft's Beta Channel announcement, and six the specification proposes
  and Excel does not have: `FILLDOWN`, `PARSEJSON`, `ROLLING`, `BINS`,
  `TABLEJOIN`, `FUZZYLOOKUP`. A workbook that registered a custom function
  under one of these names now gets the built-in one.
- A sheet read row by row without building its grid:
  `progress::Options::streaming(sheet, sink)` hands each row of an xlsx sheet
  to the sink as a `RowBatch` (the row, the sheet so far, styles, epoch) and
  drops its cells; the workbook comes back with everything but that sheet's
  cells. In JS, `Book.forEachRow(bytes, name, sheet, callback, formatted?)`,
  in both packages; a callback that throws stops the calls and its error
  comes out of `forEachRow`. Other formats are read whole and their rows
  then handed over the same way.
- Shapes carry their look: `Shape::format` (fill and outline from
  `xdr:spPr`, the chart's `ShapeFormat`), `Shape::font` (the first run's
  family, size, bold, italic and colour as a `DiffFont`) and the turn of
  `a:xfrm` (`rotation` in 60 000ths of a degree clockwise, `flip_h`,
  `flip_v`). `effective_fill`, `effective_line` and `effective_font` answer
  with what Excel shows, the shape's style (`xdr:style`) filling in what the
  element leaves out. A changed fill or line rewrites only those children of
  `xdr:spPr`, a changed turn the attributes of `a:xfrm`, a changed font the
  runs of the new text; an untouched shape still goes back byte for byte, and
  a new one is written with what it was given.
- Charts carry the look of their chart area and plot area:
  `Chart::format` (`c:chartSpace/c:spPr`) and `Chart::plot_format`
  (`c:plotArea/c:spPr`), read out of the carried markup and rewritten the
  way a series' `spPr` is.
- A chart title or legend put by hand keeps its place: `Title::layout` and
  `Legend::layout` (`ManualLayout`, the corner of a `c:manualLayout` in edge
  mode as a share of the chart area, and its size when set). A layout in
  factor mode stays in the markup.

- Filtering a table: `edit::filter_table` sets a column's criterion and hides
  the data rows the table's criteria reject, as Excel does; in JS,
  `setTableFilter(name, column, { values } | { custom } | { top } |
  { bottom })`.
- In JS, objects over `Book` whose changing calls chain: `Workbook`, `Sheet`,
  `Table`, `SheetRange` - `new Workbook().addSheet("Sales")
  .addTableFromData(...).set(1, "Amount", 125).addFilter("Region", ...)
  .records({ visible: true })`. Every `Book` method that takes the sheet
  first is a `Sheet` method without it, every other a `Workbook` method;
  what `Book` answers with nothing returns the object, so it chains. The
  package's entry point is now `index.js`, which is the generated module
  plus these four.
- Both npm packages carry the browser build next to the Node one, under
  `web/`, and `exports` hands it to everything but Node, so
  `import init, { Book } from "@rosperitus/excelerate"` works in a browser
  and in a bundler after `await init()`. `…/web` names it explicitly. The
  published package used to be the Node build alone, which a browser cannot
  load.

### Changed

- The crate declares `rust-version = "1.92"` and builds on it. It did not
  before: `if let` guards in the formula engine need a newer compiler.
- An insertion of rows or columns that would push a stored cell or a merged
  area past the last row or column is refused with the new
  `Error::WouldPushOffSheet`, and the workbook is left as it was, the way
  Excel refuses it. Those cells used to vanish without a word. Row heights,
  styles and column widths at the edge are still cut off silently, as Excel
  does. `edit::insert_cells` reports the same error instead of
  `Error::Xlsx`, and counts a merged area at the edge too.
- `Shape` gains `rotation`, `flip_h`, `flip_v`, `format` and `font`: a
  `Shape { .. }` literal needs them (or `Shape::new`).
- `Title` and `Legend` gain `layout`: a literal needs it (or
  `..Default::default()`).
- `Chart` gains `format` and `plot_format`; `ChartMarkup::after_chart` and
  `after_axes` no longer hold those `c:spPr` elements.
- `Spreadsheet`, `Worksheet`, `Chart` and `Table` are `#[non_exhaustive]`,
  so a field added later no longer breaks code outside the crate. They can
  no longer be written as literals there: start from `Spreadsheet::new`,
  `Worksheet::new`, `Chart::default()` or the new `Table::new(id, name,
  range, columns)` and set fields.
- `shared::{biff_functions, date_parse, odf_formula, palette, special}` and
  `shared::unix_seconds` are no longer public; they were internals. `shared::date`
  and `shared::codepage` stay.

### Fixed

- Outside an array formula, a reference that `IF`, `CHOOSE`, `OFFSET` or
  `INDIRECT` hands back now narrows to the cell in the formula's row or
  column, as Excel does: `=IF(TRUE,A1:A3,0)` in row 2 is A2, not A1.
- Recalculation left the rest of an array formula's area (`{=...}` entered
  over several cells) holding the old answer, and formulas reading those
  cells read it too. The area now gets the new array, laid out as Excel lays
  it: a single row or column repeats, a cell past the array is `#N/A`.
  `recalculate_from` follows an edit through the area.
- A number format written the Russian way, `# ##0,00` (a space between
  placeholders groups, a comma before the decimals), rendered its comma as
  thousands and its spaces as text. It now prints `1 234,50` with the
  non-breaking space a Russian Windows uses. This is what `TEXT` gets in a
  workbook typed in a Russian Excel. `0,000` still reads as English grouping.
- `edit::insert_cells` refused to push cells down when the column just right
  of the area (or the row just below, pushing right) had something at the
  edge of the sheet: the edge it checked was one line too wide.
- xlsx: an inline string (`t="inlineStr"`) written as `<r>` runs read as
  plain text, so its fonts were lost on a rewrite. It now reads as
  `CellValue::RichText`, the same as a rich string in the shared pool.

## 0.14.0

### Changed

- `DiffFont` gains `family`: a literal needs it (or `..Default::default()`).
- `AutoFilter` gains `sort_state`, `Table` gains `sort_state`, `Worksheet`
  gains `sort_state` and `sparklines`, `Comment` gains `visible` and `size`
  and no longer derives `Eq`: literals need the new fields (or
  `..Default::default()` where there is one).

- `SeriesMarkup::before_data` no longer holds `c:spPr`, `c:marker`, `c:dPt`
  and `c:dLbls` (now fields of `Series`), and `invertIfNegative`,
  `pictureOptions` and `explosion` moved to the new
  `SeriesMarkup::after_format`; `Plot::markup` no longer holds `c:dLbls`.
- `Plot` gains `drop_lines`, `high_low_lines` and `up_down_bars`, so a
  `Plot { .. }` literal needs them (or `Plot::new`); `Plot::markup` no longer
  holds `c:dropLines`, `c:hiLowLines` and `c:upDownBars`.
- `SeriesMarker` gains `format` and `DataLabels` gains `points`: literals of
  either need them (or `..Default::default()`).
- `Plot` gains `show_markers`; `Plot::markup` no longer holds the
  `c:marker` switch of `c:lineChart`.
- `DataPoint` gains `marker`: a literal needs it (or `..Default::default()`).
- `Fill` gains `Gradient` and `Pattern`: an exhaustive `match` needs arms
  for them, and a gradient or pattern that read as `Fill::Other` now reads as
  one of them.
- `Table::auto_filter` is an `AutoFilter` instead of a bare `Range`: a
  table's filter now keeps its hidden buttons (`hiddenButton`, which Excel's
  templates set on every column) and its criteria, both of which a rewrite
  used to drop. `Table` gains `extensions` and `CfRule` gains `extensions`:
  literals need them.

- `Error`, `CellError`, `Format`, `Fill`, `ColumnFilter`, `PlotKind`, `Expr`,
  `CfScale`, `ChartText`, `DataSource`, `ImageFormat`, `PasswordHash`,
  `LinkTarget`, `Stage`, `Color` and `NumberFormat` are `#[non_exhaustive]`:
  a `match` on one needs a `_` arm, once. They are the enums whose
  vocabulary grows with the formats; this release alone added variants to
  `Fill`, and each such addition broke every exhaustive `match` downstream.
  Enums fixed by the OOXML schema, `CellValue` and the formula `Value` stay
  exhaustive.

### Fixed

- A name bound by `LET` or held by a defined name calls the lambda it
  stands for: `LET(f,LAMBDA(x,x*2),f(3))` is 6 and a Name Manager
  `LAMBDA` (recursive ones too) works as a function; both were `#NAME?`.
- `COUNT` counts only numbers inside an array constant, as inside a
  reference: `COUNT({1,"2",TRUE})` is 1, not 3.
- `SORT` honours its fourth argument and sorts columns.
- Expressions nest at most 128 deep (was 256): a formula nested past about
  245 levels overflowed a 2 MB stack in a debug build before the parser
  refused it. Excel allows 64.
- A grid edit left the newer rules in the sheet's `<extLst>` where they were:
  the `x14` half of a data bar kept its old range while the rule beside it
  moved, and a validation list on another sheet kept its old rows and name.
  `<xm:sqref>` now moves with the sheet's cells, `<xm:f>` is rewritten like
  any formula, and a rule left with no cells goes.

- ODS and HTML wrote theme and indexed colours as nothing; they are now the
  RGB they show, through the workbook's theme with the tint.
- HTML reading lost the space between `a <b>b</b>`: each text node was trimmed
  on its own. Whitespace now collapses across elements, as a browser does.
- ODS writing let a reader collapse a leading space or a run of spaces;
  they go out as `<text:s/>`, and reading honours `text:c`.
- An xlsb whose `workbook.bin` holds the book twice (one converter writes it
  so) failed with a duplicate sheet name; reading stops at `BrtEndBook`.

- A grid edit on one sheet moved the second half of a range on another:
  removing a row of `Shown` turned `Data!B3:B10` into `Data!B3:B9`, because
  `B10` was read without the qualifier of `B3`.
- A carried `<extLst>`, `<tableStyles>` or `<colors>` that used a prefix
  declared only on the part's root (`xr2:uid` on a sparkline group, `xr9:uid`
  on a table style, as Excel 365 writes them) was written with the prefix
  undeclared. It now declares what it borrowed.

- Writing xls hung forever on a workbook whose theme had ` val="` in text
  outside a tag: the theme scan resumed before its own match. It also read
  the theme once per colour of every style; now once. Found by the fuzzer.
- `LCM` multiplied past `u64` (a wrong answer in release, a panic with
  overflow checks) and `GCD` took arguments an f64 no longer holds exactly.
  Both answer `#NUM!` at 2^53, as Microsoft documents. Found by the fuzzer.

- A table part with an `<extLst>` (Excel keeps a table's alt text there as
  `<x14:table>`) read as a table with id 0 and no name: the extension shares
  the root's local name and overwrote it. Written back, such a table had
  `name=""`.
- The `<extLst>` of a table and of a conditional formatting rule is carried
  whole. A data bar keeps its `x14:id` there, which ties it to its negative
  colours and axis in the sheet's extensions; without it Excel drew the bar
  the 2007 way.

### Added

- `at!("A1")`: a cell address checked while compiling, so a literal address
  needs no `?` and a wrong one (`at!("A0")`) does not build. Behind it,
  `CellRef::from_a1` is a `const fn`, and `CellRef::parse` now goes through
  it; `$$A1` is no longer accepted.
- `Worksheet::set_at(row, column, value)` and `get_at(row, column)`, counted
  from one; `CellRef::from_row_col(row, column)`.
- `From<i32>` and `From<u32>` for `CellValue`: `sheet.set(at, 123)` no longer
  needs `123_i64`.

- Formatted text in every format that holds it: xls reads and writes the runs
  of the string table and reads `RSTRING` (BIFF5), xlsb reads `RichStr`, ODS
  reads and writes `text:span` with text styles, HTML reads `<b>`, `<i>`,
  `<u>`, `<s>`, `<sup>`, `<sub>`, `<font>` and `<span style>` inside a cell and
  writes every part of a run's font. In wasm: `getRichText`, `setRichText` and
  their `…At` twins.
- `Font::with(&DiffFont)` and `From<&Font> for DiffFont`. `DiffFont` gains
  `family`, which a rewrite of xlsx used to drop from `<rPr>`.

- Sparklines are modelled: `Worksheet::sparklines` (`model::sparkline`) reads
  the groups out of the sheet's `<extLst>`, the writer puts back only their
  `<ext>` when they changed, and grid and sheet edits move them.
- `SortState` on a sheet, an auto filter and a table (`<sortState>`), and the
  `Color` and `Icon` kinds of `ColumnFilter` (`<colorFilter>`,
  `<iconFilter>`). All three were dropped by a rewrite; COIN's sheet 7 had one.
- `SheetView` gains `show_formulas`, `show_outline_symbols`, `show_ruler`,
  `show_white_space`, `window_protection` and `grid_color`, read and written in
  xlsx and, where the format has them, xls and xlsb.
- `Comment::visible` and `Comment::size`, read from the VML frame and written
  back into it, anchor included.
- `Spreadsheet::workbook_view`: `firstSheet`, `tabRatio` and the other
  attributes of `<workbookView>` that a rewrite used to drop.

- Document properties: `Spreadsheet::properties` (`DocumentProperties`) holds
  title, subject, author, keywords, comments, last editor, category, status,
  language, identifier, revision, version, the three dates, company, manager
  and the user's own fields (`CustomProperty`, `PropertyValue`: text, integer,
  number, boolean, date), with `set_custom` and `custom` by name. Read and
  written in xlsx (`docProps/core.xml`, `app.xml`, `custom.xml`, written by
  comparison so an untouched part keeps its bytes), xls (the two MS-OLEPS
  property set streams) and ODS (`meta.xml`); read from xlsb. In JS,
  `documentProperties()` and `setDocumentProperties(patch)`.
- The xls writer puts several streams in the compound file, not one:
  `writer::ole::streams` replaces the internal `container`.

- `Spreadsheet::template`: the book is a template. The xlsx writer gives the
  main part the template content type (`spreadsheetml.template.main+xml`, or
  `ms-excel.template.macroEnabled.main+xml` with a VBA project) - Excel will
  not open an `.xltx` written as a plain workbook - and the xlsx reader sets
  the flag from that type, so a template opened and saved stays one.
- Charts: series formatting and data labels are modelled. `Series` gains
  `format` (`c:spPr` as `ShapeFormat`: `Fill` - none, solid `ChartColor`,
  other - and `LineFormat` with colour and width), `marker` (`SeriesMarker`:
  `MarkerSymbol` and size), `data_points` (`c:dPt` as `DataPoint`: index and
  its own fill, so pie slices keep their colours) and `labels`; `Plot` gains
  `labels` (`DataLabels`: `deleted`, `position` as `LabelPosition`, the five
  `show_*` flags). `ChartColor` keeps a theme colour by name with its
  `lumMod`/`lumOff`/`tint`/`shade`/`alpha` transforms, and
  `ChartColor::resolve` turns it into rgb through the workbook's theme. Each
  element keeps its `source`: while the model still says what it says it is
  written back as read, otherwise the model's fields are written into it and
  gradients, dashes, effects and extensions stay.
- Charts: the lines and bars of a stock chart are modelled. `Plot` gains
  `high_low_lines` and `drop_lines` (`ChartLines`, with their `ShapeFormat`)
  and `up_down_bars` (`UpDownBars`: `gap_width`, and the fill and outline of
  the `up` and `down` bars), written back as read while unchanged.
- Charts: a marker's own fill and outline (`SeriesMarker::format`, the
  `c:spPr` inside `c:marker`) and labels of single points
  (`DataLabels::points`: `DataLabel` with the point index, `deleted`,
  `position` and the five `show_*` flags) are modelled. A point label keeps
  its text, layout and extensions from `source`, and a cell it shows moves
  with inserted rows like the rest of the chart.
- Charts: the marker switch of a line plot (`c:lineChart/c:marker`) is
  `Plot::show_markers`. With it on, a series with no marker of its own draws
  automatic ones, as Excel's "line with markers" does.
- Charts: the marker of a single point (`c:dPt/c:marker`) is
  `DataPoint::marker`, a `SeriesMarker` like the series' own; a changed one is
  written into the point, which keeps the rest of what it said.
- Charts: gradient and pattern fills are described.
  `Fill::Gradient { stops, angle, path }` holds the stops (`GradientStop`:
  position and `ChartColor`), the angle of a linear gradient and the shape of
  a radial one (`GradientPath`); `Fill::Pattern` the preset name and the
  foreground and background colours. A changed gradient that was read keeps
  its flip, rotation, tile and focus. Picture fills stay `Fill::Other`.
- `formula::chart::source_numbers(book, sheet, chart, source)`: the cells a
  chart source reads as `DataSource::Numbers`, with the format code of its
  first cell and points numbered as the cache numbers them (text cells are
  gaps). A text cache (`c:strCache`) has no format code in the file, so a
  program drawing dates or numbers of text categories in its own locale reads
  them here.
- `formula::chart::refresh_caches(book, changed)`: reads the caches of chart
  series names, categories, values, bubble sizes and titles again from their
  cells, for every chart (`None`) or only for references covering the cells
  an edit touched, and returns how many charts changed. A reference may be a
  range, a union or a defined name (`[0]!Name` included). The caches come out
  as Excel writes them - numbers with gaps left out, labels as the cells show
  them (numbers through their format, text as is), hidden rows and columns
  skipped under `plotVisOnly` - so on the 133 charts of the test files Excel
  saved nothing changes, and an untouched chart still goes back byte for
  byte.

### Fixed

- `SUBTOTAL` counted a nested `SUBTOTAL` or `AGGREGATE` in its references
  again, so a grand total over group totals came out doubled. Cells whose
  formula calls either are now left out, as Excel does; `AGGREGATE` does the
  same with options 0 to 3 and counts them with 4 to 7.

## 0.13.0

### Fixed

- xlsx: a cell with `t="inlineStr"` read as empty - its text sits in
  `<is><t>`, and the sheet reader only looked in `<v>`. 1C exports write every
  text cell this way, so whole columns came back `null`. Runs of a rich
  inline string are joined as plain text; `<rPh>` phonetics are skipped.

### Changed

- JS bindings: `src/wasm.rs` (3200 lines) is split by topic into
  `src/wasm/` - cells, sheet, style, objects, formulas, edit, output. The
  string and numeric twins of each call share one body, sheet lookups and
  the progress callback go through one helper each, and painting a cell and
  painting a grid of styles are the same code.
- `setRangeStyle` and `setCellStyle` work out every style before writing
  any, as `setRangeStyles` already did: a bad patch leaves the sheet as it
  was instead of half painted.
- `set` and `setRange` both tell the dependency index when a formula is
  overwritten by a value; `setRange` used to notice only new formulas.

- npm docs: the README, `docs/wasm.md` and the reader's README list
  `cellStyle`, `cellStyleAt`, `getRangeStylesAt`, `setRangeStylesAt`,
  `recalculateCell`, the progress callback and `registerFunction`, which the
  package had and the docs did not. The reader's README linked to an
  unrelated unscoped `excelerate` package.
- New TypeScript example `npm/typescript/long-operations.ts`: progress over
  read, recalculation and write, and a function of your own.

### Removed (breaking)

- `Book.fromXlsx` - `Book.read` reads xlsx and every other format.
- `toCsvWith(sheet, options)` - `toCsv(sheet, options?)` takes the options.
- `unfreezePanes(sheet)` - `freezePanes(sheet, 0, 0)` does the same.

## 0.12.3

### Changed

- `Worksheet::dimension` (JS `usedRange`) answers from the column span the
  sheet keeps instead of walking every row: 23 ms to 2 us on a million rows.
  It walks only after a cell in an edge column has been removed.

### Fixed

- xlsx writer: a row with a height or hidden but no cells is written in its
  place among the others. Before, such rows came after every row with cells,
  out of the ascending order the format requires, and each one searched every
  cell of the sheet: a sheet of 4 million cells and 330 000 row records took
  17 minutes to write and now takes 3 seconds. Array formulas and
  external hyperlinks are no longer searched for each cell or link either.

## 0.12.2

### Changed

- xls reader: 2.4 times faster on large files and a quarter less memory.
  The shared string table no longer searches for `CONTINUE` breaks on
  every byte, and cells share the table's text instead of copying it.

### Added

- `edit::fill_series_back` - the fill handle dragged up or left: the seed is
  the last cells of each line and the series runs back from them (`1, 2`
  above goes `0, -1`, a single `Кв3` goes `Кв2`). JS: `fillSeries` with
  `"up"` or `"left"`.
- HTML reader: column widths from `col.x` rules and row heights from `tr.x`
  rules of a `<style>` block, the way PhpSpreadsheet sizes its export.

### Fixed

- xlsx reader: a six-digit `rgb="D8D8D8"` in `styles.xml` is read as an
  opaque colour, as Excel reads it. Before, it was dropped and the fill came
  out blank. The JS style patch takes `#RRGGBB` for the same reason.
- xls: row and column outline levels, collapsed groups and where summaries
  sit (`WSBOOL`) are read and written back, with the `GUTS` record Excel needs
  to show the outline bar. Before, a sheet with summaries above its groups
  came back with them below, and column levels were lost.

- xls reader: style 0 is the workbook's Normal style from the default cell
  format (XF 15), not the Calibri 11 default. Column widths are counted in
  digits of its font, so an Arial 8 book written back to xls or xlsx opened
  in Excel with every column wider.
- xls: frozen and split panes are read from `WINDOW2` and `PANE` and written
  back, the selection of each pane from `SELECTION`, along with the scroll position and the gridline, heading, zero and
  right-to-left switches of the window. Before, a frozen sheet read from xls
  opened unfrozen.
- xls writer: turning gridlines off cleared the default-gridline-colour bit
  instead of the gridline bit, so the grid stayed on.

## 0.12.1

### Added

- `edit::insert_rows_with` and `edit::insert_columns_with` take a
  `CopyOrigin` - `Blank`, `Before` or `After`, Excel's `CopyOrigin` on
  `Range.Insert` - and give the new lines the cell styles, own style and
  height or width of that neighbour. In JS it is the optional last argument
  of `insertRows` / `insertColumns` / `insertCells`: `"before"` (the
  default, as in Excel), `"after"` or `"none"`. `edit::insert_cells_with`
  does the same for Insert Cells.
- Sorting: `edit::sort_range`, `sort_range_with` and `sort_table`. Rows sort
  in Excel's order (numbers, text without case, `FALSE`, `TRUE`, errors;
  empty cells last both ways) and formulas are rewritten as if copied to
  their new row. A `SortKey` names a sheet column or row, or a header
  (`SortKey::header("Amount").descending()`); `SortOptions` keeps a header
  row on top or sorts columns left to right; `sort_table` takes a table's
  name and its column names. JS: `sortRange(sheet, range, [2, "-Amount"],
  { header, byColumns })` and `sortTable(name, keys)`.
- `edit::fill` - Ctrl+D and Ctrl+R: the first row or column of a range copied
  over the rest, references moving as in a copy. JS: `fillDown`, `fillRight`.
- `edit::fill_series` - the fill handle: numbers follow their trend, a date
  steps a day, `Кв1` counts on to `Кв2`, month and weekday names (English and
  Russian) go round their list, anything else repeats. JS: `fillSeries`.
- JS: `getRangeStyles` / `setRangeStyles` (and `...At`) read and write a
  rectangle's formatting in one call, as `{ styles, grid }` with each distinct
  style once; what `get` returns, `set` takes back.
- HTML round trip, checked on a 1C export (416 cells, 110 merges): a merged
  block's `<td>` now carries the right and bottom borders of the cells on
  that edge, so the page draws the whole frame; the reader spreads
  `data-format` over the merge, reads `white-space` as wrapping, and text
  with runs of spaces keeps them through `data-value`.
- Docs: sorting, filling and formatting inserted rows in `docs/editing.md`,
  three new recipes, the JS API table and the npm readme brought up to date
  (the scoped package name, 518 functions, `toXls` writes formulas).
- The HTML reader applies the page's `<style>` rules (`.x`, `td.x`) to cells
  by class, under any inline `style`; the HTML writer marks a non-General
  number format with `data-format`, and every value that would not read back
  as itself with `data-type`/`data-value` (a formula with `data-formula`),
  so a page we wrote reads back with its numbers, formulas and styles.

## 0.12.0

### Added

- `Worksheet::dimension_hint()`: the used range without walking the rows.
  Rows are exact, columns are an upper bound kept as cells are written, and a
  removed cell never narrows it. `dimension()` walks every row of the sheet -
  some twenty milliseconds on a sheet of seven hundred thousand rows, paid
  again for every sheet a viewer switches to - and a scrollbar or a clipped
  viewport does not need that precision.
- The npm packages caught up with the crate. `Book` now edits the grid
  (`insertRows`, `removeRows`, `insertColumns`, `removeColumns`,
  `removeSheet`, `merge`, `unmerge` - all under the `write` feature, so the
  reader package stays a reader), answers the cheap geometry questions
  (`usedRangeHint`, `columnWidth`, `rowHeight`), and lists what a sheet holds
  besides cells: `comments`, `hyperlinks`, `tables`, `charts`, `images` with
  `imageData` for the bytes, and `shapes`. Each object reports its anchor as
  `{ kind, row, column }` in the numbers a user sees.
- `Worksheet::set_column_width`, `set_column_hidden`, `set_row_height` and
  `set_row_hidden`, with `column_entry` under them: setting one column splits
  the run that covered it, so its neighbours keep the width they had. Until
  now a caller pushed a `ColumnRun` by hand and overlapping runs were their
  problem. The npm package gets them as `setColumnWidth`, `setRowHeight`,
  `setColumnHidden`, `setRowHidden`, along with `setSheetVisibility`.
- The rest of the wrapper's gaps, the ones a JS caller hits first: writing a
  style (`setCellStyle`, `setCellStyleAt`, `setRangeStyle` - a patch over what
  the cell had, equal styles sharing one table entry), notes, links and tables
  (`setComment`, `setHyperlink`, `addTable` and their removals), the saved view
  (`sheetView`, `freezePanes`, `setZoom`, `setShowGridLines`), defined names
  (`definedNames`, `setDefinedName`, `removeDefinedName`), the locks
  (`protectSheet`, `protectWorkbook`, `verifySheetPassword` and friends), and
  the rules a file states on the reading side (`dataValidations`,
  `conditionalFormats`, `autoFilter`, `pivotTables`, `arrayFormulas`,
  `externalBooks`). CSV gained `toCsvWith` and `Book.readCsv` for a file whose
  shape is known rather than guessed.
- `edit::copy_range` and `edit::move_range`: a rectangle of cells, values,
  styles and the merges inside it. A copy rewrites the formulas it carries the
  way Excel rewrites a copied formula; a move keeps them pointing where they
  pointed and rewrites the formulas elsewhere that read the moved cells.
  Limits, both stated in `docs/editing.md`: a cross-sheet move spells the old
  sheet's name out per reference, and leaves the references from other cells
  alone.
- `edit::insert_cells` and `edit::remove_cells`: Excel's "Insert Cells", where
  only the columns or rows the block spans move and the rest of the sheet
  stays put. A reference travels only when the whole of it does.
- `edit::move_sheet`: a sheet's place in the tab bar. Formulas name sheets
  rather than numbering them, so they do not change; the active tab and the
  sheet a defined name belongs to are renumbered.
- `Range::contains_range` and `Range::intersects`, which those needed.
- The npm packages get the five as `copyRange`, `moveRange`, `insertCells`,
  `removeCells` and `moveSheet`.
- `reader::csv::read_csv_bytes`: CSV from bytes in memory, decoding them the
  way `read_csv_with` does for a path.
- `PlotKind::element()`: the chart element a plot is drawn as, `barChart` and
  the rest. It was inside the xlsx writer, which the read-only build leaves
  out.

## 0.11.0

### Fixed

- Text in a cell whose number format has one section showed the format code
  itself: `-` in a cell formatted `#,##0` read `#,##0`. A single section
  applies to text only when it spells out `@`, which is what Excel does. 1364
  cells of one test workbook were wrong this way.
- Number format ids 5 to 8 and 41 to 43 were read as General; they are the
  currency and accounting formats the spec spells out. Ids 50 to 58, the
  second block of locale-dependent slots, now read as a short date or time
  rather than General, so a cell carrying one shows a date instead of 693597.
- Serial 0 under a date format showed 31 December 1899. Excel shows the zeroth
  of January 1900, and now so does this.

### Added

- `CellValue::formula(text)` builds a formula with no result yet, the way
  `CellValue::text` builds text. `cached: None` was written out 26 times in
  this repository alone.
- `CellValue::result()` looks through a formula to the value last computed for
  it, and `as_number`, `as_str`, `as_bool` and `as_error` look through it to
  one type. An uncomputed formula reads as `Empty` rather than as an error.
  A boolean is not a number, as Excel keeps the two apart.
- `Worksheet::set_styled(at, value, style)` writes both in one lookup of the
  cell, and `Worksheet::set_row(from, values)` writes a run along a row. The
  run stops at the last column rather than wrapping round to column A.
- `Spreadsheet::formatted_at(sheet, row, column)` is `formatted` by 1-based
  numbers, for callers that count rows rather than parse addresses.
- `CellValue::display(code, epoch)` renders a value through a number format,
  and `Spreadsheet::formatted(sheet, at)` does it with the cell's own format
  and the workbook's epoch already looked up. The second was previously
  reachable only from JavaScript, as `getFormatted`; in Rust it took twenty
  lines. The wasm wrapper and the HTML writer now call the same code.

## 0.10.0

### Added

- **BIFF5 is read**: the xls Excel 5 and 95 wrote, which was refused outright
  before. It packs the same records more narrowly - one byte per character in
  the workbook's code page, one byte per column, the relative flags of a
  reference on its row, sixteen-byte `XF` records - and names the sheet of a
  3D reference inside the token rather than through an `EXTERNSHEET` entry.
  Values, sheets, formulas, defined names and the cell format all come across.

- BIFF5 text is read in the code page the workbook's `CODEPAGE` record names.
  `shared::codepage` holds Windows 1252 and 1251, Mac Roman and DOS 866, and
  reads anything else as 1252 - three of the four BIFF5 workbooks tested say
  Mac Roman, so 1252 alone would have been wrong. `read_xls_in(path, page)`
  and `read_xls_from_in(bytes, page)` name the page for a workbook that has no
  record or whose record lies; the page named there wins.

### Fixed

- A workbook stream named in another case is found: `WorkBook` beside a stale
  BIFF5 `Book`, or `BOOK` on its own. Files like these exist and LibreOffice
  opens them; before this the reader took the first name it knew and refused
  the file as BIFF5.
- Part names written with a Windows separator (`xl\workbook.xml`) resolve. The
  format says forward slash, but packages spelled this way exist and Excel
  opens them.
- ODS: a cell covered by a merge keeps the value of its own, both read and
  written. The merge hides it rather than deletes it, and dropping it lost
  what the file said.
- ODS: a repeat count no longer expands without bound. `999999999` rows of
  `999999999` cells is thirteen hundred bytes to write, and the grid's own
  bounds still leave seventeen billion cells; copies made by repeat counts are
  now capped at four million per sheet, which leaves a sheet that really holds
  that many cells untouched.

## 0.9.0

### Breaking changes

- `ChartEx`, `PivotTable` and `PivotCache` have an `origin` field, which a
  struct literal has to fill (`..Default::default()` does).
- `CellValue::Text` holds `Arc<str>` instead of `String`. Cells read from the
  same shared string share one allocation. `CellValue::text` still builds one
  from any string; `CellValue::shared_text` takes an `Arc<str>` as it is.
- A zip package larger than the expansion cap is read when it expands no more
  than 100 times its compressed size. The cap alone refused a 70 MB export
  that expands to 634 MB and that Excel opens; the ratio is what tells such a
  workbook from a zip bomb.

### Added

- **xlsb (BIFF12) is read.** `reader::read` picks it up by the parts inside
  the package, not by the extension, so a renamed file still opens; the wasm
  `Book.read` gets it for free. Values, shared strings, formulas (shared and
  array ones expanded), defined names, merges, sheet visibility, the full cell
  style (number format, font, fill, borders, alignment, protection), row
  heights and hidden rows, column widths, the saved view with its frozen panes
  and selection, and the auto filter all come across. Checked against the
  same workbook saved as xlsx: 3769 cells, every cell style, every row and
  every sheet view agree, and so does every formula.

- Number formats read the Cyrillic date codes a Russian Excel is typed in:
  `ТЕКСТ(A1;"ДД.ММ.ГГГГ")` is a date, as it is in Excel, rather than the
  letters `ДД` beside the serial number. Month and weekday names now match
  what Excel 2019 shows for a Russian locale tag: `Январь`, `янв`, `среда`,
  `Ср`, `Я`.
- An array formula keeps the area it fills in every format that can say it:
  an `ARRAY` record in xls and a matrix span in `OpenDocument`, which are also
  read back. Before, only xlsx carried it, and a workbook that went through
  xls or ODS came back with an ordinary formula in its place - which the
  implicit intersection then read as one cell.
- `PasswordHash::new`, `iso`, `legacy` and `verify`: set a sheet or workbook
  password from its text the way Excel does (SHA-512, random 16-byte salt,
  100,000 spins), and check a password against a hash a file carries.
  Verified against hashes `excelize` wrote, and Excel 2019 unprotected a sheet
  this crate locked, with the password it was locked with.
- `Worksheet::shapes`: boxes, arrows, callouts and text boxes are read into
  `Shape` (name, alt text, anchor, preset outline, text) and written back. An
  untouched shape keeps its bytes; a rename, new outline, new text or move is
  spliced into the drawing; a removed shape is cut and a new one added.
  Inserting rows and columns moves shapes with the grid.
- 2016 charts (`ChartEx`: waterfall, funnel, treemap and the rest) are
  written back. An untouched one keeps its bytes; a change to the title or a
  series' layout, name, visibility or data is spliced into the part, which
  keeps every colour and label style; a removed one is cut from the drawing,
  and `ChartEx::new` builds one with Excel's defaults. Inserting rows moves
  them and what they read.
- Pivot tables and caches are written back. An untouched one keeps its
  bytes; a changed or new report is written from the model the way excelize
  writes one, with the cache marked `refreshOnLoad`; a removed report takes
  its part with it. excelize reads the changed and the new report with the
  same fields.
- Number formats render fractions (`# ?/?`, `?/???`, `# ??/100`) and the
  symbol in a locale tag (`[$₽-419]`); a Russian locale tag names months and
  weekdays in Russian. Before, all three were dropped from the output.
- `GETPIVOTDATA` reads a value out of a pivot report: the labels down its
  side, the headers across its top and the value where they meet, as they
  were laid out in the cells. Asking for one field of several leaves that
  field's subtotal, and asking for none leaves the grand total, as in Excel.
  `PivotTable` gained `first_data_row` and `first_data_col`, which say where
  the report's headers end.
- Ten more functions: `GROUPBY` and `PIVOTBY` (the summarising function a
  lambda or a bare name such as `SUM`, stored in files as `_xleta.SUM`),
  `PERCENTOF`, `TRIMRANGE`, `REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE`,
  `INFO`, `EUROCONVERT` and `PHONETIC`. The regular expressions run on the
  `regex` crate, a new dependency of the `formulas` feature: no lookaround or
  backreferences, and 550 KB more wasm.
- A comment added to a sheet gets the VML box Excel draws it in, and a sheet
  with no VML part gets one; the box of a removed comment is cut out. Before,
  a new comment was written without a box and Excel did not show it.

### Fixed

- Excel 2019 crashed on an xls holding an array formula: the cells the array
  covers were written as values, where BIFF wants a formula record pointing at
  the cell the array starts in. Each of them now carries that pointer, and the
  expression is said once, in the `ARRAY` record.
- A 2016 chart written from the model declares every namespace it uses on the
  element that uses it and carries the `mc:Fallback` Excel writes. Excel
  dropped the whole drawing of such a file, and this is the difference against
  the ones it writes itself.
- A rewrite dropped `<tableStyles>` and `<colors>` from the stylesheet. Slicers
  are painted with a style named there, so Excel removed all fourteen slicers
  of a dashboard this crate had rewritten; the palette an `indexed` colour
  counts into went the same way. Both are carried now.
- An array formula was written into every cell of its area rather than only
  into the cell that holds it. Excel called that a damaged file and stripped
  the sheet's cells; the xls form of it made Excel 2019 fall over.
- `OpenDocument` stated every number of a cell with no format of its own as a
  date: `is_date_format` took the `e` of `General` for the Japanese era code,
  so `10` was written as the tenth of January 1900. Found by asking Excel to
  read one of our files.
- A title typed into a 2016 chart is written as formatted text rather than as
  the form that carries a cell's formula, which has no formula to carry.
- Building the dependency graph of a pass halved: what sits inside each
  rectangle is worked out on several threads, and the edges live in vectors
  rather than hash maps (COIN: 1.53 s to 0.60 s). An engine that wants a
  formula another is already computing now waits briefly for the answer
  instead of computing the same chain beside it. A full pass over COIN takes
  6.9 s, against 13.5 s on one thread and 16.9 s before any of this.
- `recalculate_from` did not finish on a large workbook. Finding what an edit
  reaches asked every formula, on every round, whether any dirty cell was
  inside what it read; one edit in a book of 651,614 formulas ran for over ten
  minutes. It now walks the dependency graph once, computes the formulas it
  found in waves on several threads, and takes the results already stored
  beside the other cells rather than computing them again: the same edit is
  9.3 s, of which 1.1 s is building the index.
- A formula reading a defined name is recomputed when the cells that name
  stands for change, rather than on every edit anywhere. A name built at
  evaluation time, and a volatile function, still mean every pass.
- A full recalculation computed most formulas twice. The pass computed them
  in dependency order but did not keep what each one worked out, so the next
  formula reading that cell computed it again from scratch. COIN's 651,614
  formulas went from 16.9 s to 14.6 s on one thread with this alone.
- A recalculation of more than twenty thousand formulas now parses them and
  computes them on up to eight threads, in waves of formulas that read
  nothing of each other: COIN takes 8.2 s instead of 16.9 s, cell for cell
  the same values. A caller's own functions, a smaller workbook and
  WebAssembly keep the single-threaded pass.
- Writing a workbook with a changed chart, picture, shape, comment or pivot
  copied every cell of it first. The cells of a sheet are now shared by a
  clone until one of them changes: saving a million-cell export with one new
  shape peaks at 1.18 GB instead of 1.68 GB.
- A chain of formulas longer than 500 links, reached from one cell rather
  than in dependency order (`recalculate_cell`, a reference built by
  `INDIRECT`), was `#VALUE!` past the 500th link. The engine now sets the
  deepest link aside, computes it from the top and resumes, so any length
  adds up in bounded stack; a cycle longer than that still ends as a cycle.
- Array formulas keep their flag through an xlsx round trip:
  `<f t="array" ref="…">` is read into `Worksheet::array_formulas` and
  written back. Before, `{=…}` came back as a plain formula.
- Implicit intersection: a formula stored in a cell that answers with a range
  shows the cell of that range in its own row or column (`=A1:A9` in B5 is
  A5; `INDEX(A1:D6,2,0)` in column C is C2), and `#VALUE!` when its row
  misses the range. The same holds for a range met by an operator or passed
  to a value parameter (`=A1:A9*2`, `ABS(A1:A9)`, a defined name standing for
  a range); array parameters (`SUMPRODUCT`) and array formulas keep the range
  whole. A reference returned at run time by `IF`, `CHOOSE` or `OFFSET` is not
  narrowed.
- The `ARRAY` record of an xls file marks its formula as an array formula.

### Performance

- Reading a 70 MB xlsx of nine million cells takes 5.4 s and 770 MB (Excel
  2016: 7 s and 900 MB). Before, it was refused, and with the cap raised took
  7.3 s and 2.7 GB. Shared strings are no longer copied into each cell; the
  sheet and the shared string table are parsed as they inflate instead of
  being held as text; attributes are looked up in place instead of
  collected; and a sheet stores its cells as a sorted vector per row instead
  of one tree keyed by cell, each new row reserving the width of the one
  before. Walking every cell and computing the used range take a quarter of
  the time.

### Fixed

- The styles reader took the differential formats of slicer styles inside
  `<extLst>` for the book's own, and each save added them again: `chart1.xlsx`
  went from 50 to 98 to 146.

## 0.8.0

### Breaking changes

- `Value::Array` holds `Rc<Vec<Vec<Value>>>` instead of `Vec<Vec<Value>>`, so
  a range read once is shared by every formula that reads it. Build one with
  `Value::array(rows)`; take the rows out with `Rc::unwrap_or_clone`.
- `Expr::Range` has a new field, `anchors`, recording which parts of the
  reference were written with `$`. Patterns that list the fields need `..`.
- A reference to a cell whose formula computed an array reads the cell's
  top-left value, as Excel shows it. `Engine::spilled` (and `A1#` in a
  formula) returns the whole array.
- An error in a function argument is the function's answer, whichever error
  it is. `RANDBETWEEN(#REF!, 1)` is `#REF!`, where it was `#VALUE!`.
- `AVERAGEA`, `MAXA`, `MINA`, `STDEVA`, `STDEVPA`, `VARA` and `VARPA` return
  `#VALUE!` for text written directly as an argument, as Excel 2016 does.
  Text read from cells still counts as zero.
- Approximate `MATCH`, `VLOOKUP`, `HLOOKUP` and `LOOKUP` search by halving,
  among values of the looked-up value's type. On an unsorted list the answer
  is the one Excel gives, which can differ from the first match a linear scan
  found.
- Financial functions that walk a schedule period by period (`DB`, `DDB`,
  `VDB`, `IPMT`, `PPMT`, `CUMIPMT`, `CUMPRINC`, `AMORDEGRC`) return `#NUM!`
  past one million periods.
- Arrays a formula builds are capped at four million cells, the cap a
  reference already had. `EXPAND`, `RANDARRAY`, element-wise operators and
  lifted functions return `#NUM!` beyond it.
- Dates past 31 December 9999 or before the start of the calendar are
  `#NUM!` in every function that reads a date.

### Added

- xls formulas are read as text, beside their cached results: relative and
  absolute references, 3D references, add-in and newer functions, array
  constants, shared formulas expanded per cell, array formulas on their first
  cell, and defined names.
- xls formulas are written as BIFF8 tokens with their results, and defined
  names as `NAME` records. What BIFF8 cannot hold (a structured reference, a
  reference past row 65536 or column IV) is written as its value.
- The whole xls cell format is read and written: fonts, fills, borders,
  alignment and protection. Palette colours resolve to RGB on reading; the
  writer lays out its own palette, and resolves theme colours through the
  workbook's theme.
- Pictures: `Worksheet::images` models each embedded picture of a sheet's
  drawing (bytes, format, anchor, name, alt text). An untouched picture is
  written back byte for byte; a moved, renamed, replaced, removed or new one
  is spliced into the drawing. Inserting rows moves pictures with the grid.
- Functions with a value parameter are lifted over arrays, the way array
  formulas need: `ISNUMBER(A1:A3)` returns three answers. `IF`, `IFERROR` and
  `IFNA` work element by element on arrays.
- `INDIRECT` follows a defined name that stands for a reference.
- `examples/why.rs` follows a formula that disagrees with its cached result
  to the cells it reads, down to the ones whose inputs agree.
- `tests/corpus.rs`, an ignored test over the workbooks in `tests/corpus/`
  (kept out of git), with floors for how many formulas agree with the cache.
- `fuzz/`, three `cargo-fuzz` targets: any bytes through format detection,
  an xls stream, and formula text.

### Fixed

- Functions stored under the `_xlfn.` prefix, which Excel writes for every
  function added after 2007, answered `#NAME?`.
- An `.xlsm` was written with the content type of a plain workbook, which
  Excel refuses to open.
- Conditional formats and validations inside a sheet's `<extLst>` were read
  as the sheet's own, and written back as a rule with `sqref=""`.
- xls files over 7 MB were written without DIFAT sectors, losing everything
  past the first 7 MB.
- A reference to a cell holding an array formula counted the whole array
  again: `COUNTIF` over such a column returned 110 for 49 cells.
- A leading `+` converted text to a number: `=+Sheet!A1` on text was
  `#VALUE!`.
- Aggregates skip text inside computed arrays, as inside references:
  `SUM({1,"2",TRUE})` is 1.
- `NPER` refused a present value of zero.
- `DATEDIF` with `"MD"` counted from the previous month when the day had
  already come.
- `XNPV` and `XIRR` dropped a non-numeric date and paired the rest with the
  wrong cash flows; they return `#VALUE!`.
- `TBILLEQ` returned a negative yield for a discount that makes the bill
  worthless; it returns `#NUM!`.
- `IRR` and `XIRR` stopped at a tolerance of 1e-8; they now agree with
  Excel to the sixteenth digit.
- `COUNTIF` and its family counted nothing, instead of failing, when the
  range itself was an error.
- xls writing took `GETPIVOTDATA` and `RTD` to have at most 2 and 5
  arguments, and wrote a formula calling them with more as its value.
- Crashes and hangs found by fuzzing: a CSV whose first line has a
  multi-byte character in its first four bytes; an xls run of cells starting
  near the last column; a stray `]` in a structured reference, which looped
  forever; `INDEX` with a fractional position below one; `OFFSET` and
  `EDATE` with counts near the integer limits.
- A CSV cell holding an uncached array formula was written as the debug form
  of its value, `Number(1.0)`.

### Performance

- Full recalculation of a workbook with 650,000 formulas takes 16 s and
  1.4 GB. Before, it did not finish, and used 12 GB. The dependency order goes
  through one node per distinct range; ranges read are cached; names are
  indexed and their formulas parsed once; recalculation reuses the trees
  parsed for the index; text compares without allocating.
- Reading a 30 MB xls took 8.7 s and takes 0.4 s: a record break inside a
  shared string is found by binary search.
- CSV writing visits only the cells each row has, instead of looking up every
  field of the used range.
- `CUMIPMT` and `CUMPRINC` walk the schedule once instead of once per period.
- `BINOM.INV` and `CRITBINOM` search for the count, and
  `BINOM.DIST.RANGE` subtracts two distribution values, instead of summing
  millions of terms.

## 0.1.0

First version: reading and writing xlsx, xls, ods, csv and html; reading
SYLK, Gnumeric and SpreadsheetML 2003; the formula engine; styles, charts,
tables, protection, autofilters and comments; pivot tables on read; the
WebAssembly packages.
