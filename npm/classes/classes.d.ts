/**
 * Objects over `Book`, whose changing methods return themselves so calls chain:
 *
 * ```js
 * const wb = new Workbook();
 * const sales = wb.addSheet("Sales")
 *   .addTableFromData("Sales", "A1", [["Region", "Amount"], ["North", 120], ["South", 340]])
 *   .set(1, "Amount", 125)
 *   .addFilter("Region", { values: ["South"] });
 * sales.records({ visible: true }); // [{ Region: "South", Amount: 340 }]
 * ```
 *
 * Each object is a handle - a sheet index, a table name, an address - and
 * works through `workbook.book`, which is there for everything not wrapped.
 */
export declare class Workbook {
  /** A new workbook, or one over a `Book` already in hand. */
  constructor(book?: Book);
  readonly book: Book;
  /** Reads any format `Book.read` reads. */
  static read(bytes: Uint8Array, name?: string, maxExpanded?: number, onProgress?: Function, password?: string): Workbook;
  static readCsv(bytes: Uint8Array, options: CsvOptions): Workbook;
  /** Hands one sheet over row by row without keeping its grid; the rest of the workbook comes back. */
  static forEachRow(bytes: Uint8Array, name: string | null | undefined, sheet: number, callback: Function, formatted?: boolean | null): Workbook;
  /** A sheet by position (from 0) or by name. */
  sheet(which: number | string): Sheet;
  sheets(): Sheet[];
  /** Adds a sheet at the end. On a new workbook the first call takes over its empty "Sheet1". */
  addSheet(name?: string): Sheet;
  recalculate(onProgress?: Function): this;
  toXlsx(onProgress?: Function): Uint8Array;
  free(): void;
}

export declare class Sheet {
  readonly workbook: Workbook;
  /** The tab position the sheet had when it was handed out. */
  readonly index: number;
  readonly book: Book;
  readonly name: string;
  rename(name: string): this;
  /** What the cell shows, through its number format. */
  text(address: string): string;
  /** A block of values whose top left corner lands on `at`. */
  write(at: string, rows: CellGrid): this;
  range(address: string): SheetRange;
  style(address: string, patch: CellStylePatch): this;
  /** Column width in characters; the column is a letter or a number from 1. */
  width(column: string | number, characters?: number): this;
  /** Freezes the rows above and the columns to the left; zeros unfreeze. */
  freeze(rows: number, columns?: number): this;
  /** A table over a range already filled. */
  addTable(name: string, address: string, options?: { header?: boolean }): Table;
  /** Writes `rows` from `at` - the first is the header - and draws a table over them. */
  addTableFromData(name: string, at: string, rows: CellGrid): Table;
  table(name: string): Table;
  tables(): Table[];
  /** Makes this the sheet a workbook opens on. */
  activate(): this;
  /** Recalculates the formulas of this sheet. */
  recalculate(onProgress?: Function): this;
  /** Removes the sheet; handles to the sheets after it go stale. */
  remove(): Workbook;
}

export declare class Table {
  readonly sheet: Sheet;
  readonly name: string;
  readonly book: Book;
  readonly columns: string[];
  info(): SheetTable;
  /** The whole table, header and totals included; with an address, a range of the sheet. */
  range(address?: string): SheetRange;
  /** The data rows, without the header and totals. */
  data(): SheetRange;
  /** The sheet address of a data cell: `row` from 1, `column` a header or a number from 1. */
  address(row: number, column: string | number): string;
  columnIndex(header: string): number;
  set(row: number, column: string | number, value: CellValue): this;
  get(row: number, column: string | number): CellValue;
  /** Filters by a column and hides the rows that fail, as Excel does. */
  addFilter(column: string | number, criteria: TableFilter): this;
  /** Takes the criterion off one column, or off all of them. */
  clearFilter(column?: string | number): this;
  /** `["Region", "-Amount"]`: a minus sorts largest first. */
  sort(keys: SortKeys): this;
  /** The data rows keyed by header; `visible` leaves out hidden rows. */
  records(options?: { visible?: boolean }): Record<string, CellValue>[];
}

export declare class SheetRange {
  readonly sheet: Sheet;
  readonly address: string;
  readonly book: Book;
  values(): CellGrid;
  /** The values of the rows a filter or a user has not hidden. */
  visibleValues(): CellGrid;
  /** Writes a block from the top left corner of the range. */
  set(rows: CellGrid): this;
  style(patch: CellStylePatch): this;
  merge(): this;
}

/**
 * Every `Book` method that takes the sheet first is a `Sheet` method without
 * it; one that answers with nothing returns the sheet, so it chains.
 * `sheet.setColumnWidth(2, 14).setRowHidden(5, true).getRange("A1:C3")`.
 */
type SheetActionName =
  | "addDataValidation" | "clear" | "clearAt" | "copyRange" | "fillDown"
  | "fillRight" | "fillSeries" | "freezePanes" | "insertCells" | "insertColumns"
  | "insertColumnsMany" | "insertRows" | "insertRowsMany" | "merge"
  | "moveRange" | "protectSheet" | "removeCells" | "removeColumns"
  | "removeColumnsMany" | "removeRows" | "removeRowsMany" | "set" | "setAt"
  | "setCellStyle" | "setCellStyleAt" | "setChartFormat" | "setColumnHidden"
  | "setColumnWidth" | "setComment" | "setHyperlink" | "setRange" | "setRangeAt"
  | "setRangeStyle" | "setRangeStyles" | "setRangeStylesAt" | "setRichText"
  | "setRichTextAt" | "setRowHeight" | "setRowHidden" | "setShapeFormat"
  | "setSheetVisibility" | "setShowGridLines" | "setZoom" | "sortRange"
  | "unprotectSheet";
type SheetQueryName =
  | "arrayFormulas" | "autoFilter" | "cellBold" | "cellBoldAt" | "cellCount"
  | "cellIndent" | "cellIndentAt" | "cellStyle" | "cellStyleAt" | "charts"
  | "checkValidation" | "columnLevel" | "columnWidth" | "comments"
  | "conditionalFormats" | "dataValidations" | "evaluate" | "get" | "getAt"
  | "getFormatted" | "getFormattedAt" | "getFormula" | "getFormulaAt"
  | "getRange" | "getRangeAt" | "getRangeStyles" | "getRangeStylesAt"
  | "getRichText" | "getRichTextAt" | "getRowAt" | "goalSeek" | "hyperlinks"
  | "imageData" | "images" | "invalidCells" | "mergedRanges" | "mergedRangesAt"
  | "pivotTables" | "recalculateCell" | "recalculateCellAt" | "recalculateFrom"
  | "recalculateFromAt" | "recalculateFromMany" | "removeComment"
  | "removeDuplicates" | "removeHyperlink" | "removeTable" | "rowHeight"
  | "rowHidden" | "rowLevel" | "shapes" | "sheetProtection" | "sheetView"
  | "sheetVisibility" | "textToColumns" | "toCsv" | "toHtml" | "unmerge"
  | "usedRange" | "usedRangeHint" | "validationList" | "verifySheetPassword";
/** Every `Book` method without a sheet is a `Workbook` method, chaining the same way. */
type BookActionName =
  | "moveSheet" | "protectWorkbook" | "registerFunction" | "setDefinedName"
  | "setDocumentProperties" | "setTableFilter" | "sortTable" | "unprotectWorkbook";
type BookQueryName =
  | "activeSheet" | "definedNames" | "documentProperties" | "externalBooks"
  | "registeredFunctions" | "removeDefinedName" | "sheetIndex" | "sheetNames"
  | "toOds" | "toXls" | "unregisterFunction" | "workbookProtection";

type SheetActions = {
  [P in SheetActionName]: Book[P] extends (sheet: any, ...rest: infer A) => any ? (...rest: A) => Sheet : never;
};
type SheetQueries = {
  [P in SheetQueryName]: Book[P] extends (sheet: any, ...rest: infer A) => infer R ? (...rest: A) => R : never;
};
type BookActions = {
  [P in BookActionName]: Book[P] extends (...args: infer A) => any ? (...args: A) => Workbook : never;
};
type BookQueries = {
  [P in BookQueryName]: Book[P] extends (...args: infer A) => infer R ? (...args: A) => R : never;
};
export interface Sheet extends SheetActions, SheetQueries {}
export interface Workbook extends BookActions, BookQueries {}
