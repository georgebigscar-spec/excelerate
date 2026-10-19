// Objects over `Book`: a workbook hands out sheets, a sheet hands out tables
// and ranges, and every method that changes something returns its object, so
// calls chain. Each object is a handle - a sheet index, a table name, an
// address - and does its work through the `Book` underneath, which stays
// reachable as `workbook.book` for everything not wrapped here.
//
// The build wraps this file for each target (CommonJS for Node, ES modules for
// the browser and bundlers); it defines `build` and nothing else.

function build(Book) {
  const SHEET_ACTIONS = [
    "addDataValidation", "clear", "clearAt", "copyRange", "fillDown", "fillRight",
    "fillSeries", "freezePanes", "insertCells", "insertColumns",
    "insertColumnsMany", "insertRows", "insertRowsMany", "merge", "moveRange",
    "protectSheet", "removeCells", "removeColumns", "removeColumnsMany",
    "removeRows", "removeRowsMany", "set", "setAt", "setCellStyle",
    "setCellStyleAt", "setChartFormat", "setColumnHidden", "setColumnWidth",
    "setComment", "setHyperlink", "setRange", "setRangeAt", "setRangeStyle",
    "setRangeStyles", "setRangeStylesAt", "setRichText", "setRichTextAt",
    "setRowHeight", "setRowHidden", "setShapeFormat", "setSheetVisibility",
    "setShowGridLines", "setZoom", "sortRange", "unprotectSheet",
  ];
  const SHEET_QUERIES = [
    "arrayFormulas", "autoFilter", "cellBold", "cellBoldAt", "cellCount",
    "cellIndent", "cellIndentAt", "cellStyle", "cellStyleAt", "charts",
    "checkValidation", "columnLevel", "columnWidth", "comments",
    "conditionalFormats", "dataValidations", "evaluate", "get", "getAt",
    "getFormatted", "getFormattedAt", "getFormula", "getFormulaAt", "getRange",
    "getRangeAt", "getRangeStyles", "getRangeStylesAt", "getRichText",
    "getRichTextAt", "getRowAt", "goalSeek", "hyperlinks", "imageData", "images",
    "invalidCells", "mergedRanges", "mergedRangesAt", "pivotTables",
    "recalculateCell", "recalculateCellAt", "recalculateFrom", "recalculateFromAt",
    "recalculateFromMany", "removeComment", "removeDuplicates", "removeHyperlink",
    "removeTable", "rowHeight", "rowHidden", "rowLevel", "shapes",
    "sheetProtection", "sheetView", "sheetVisibility", "textToColumns", "toCsv",
    "toHtml", "unmerge", "usedRange", "usedRangeHint", "validationList",
    "verifySheetPassword",
  ];
  const BOOK_ACTIONS = [
    "moveSheet", "protectWorkbook", "registerFunction", "setDefinedName",
    "setDocumentProperties", "setTableFilter", "sortTable", "unprotectWorkbook",
  ];
  const BOOK_QUERIES = [
    "activeSheet", "definedNames", "documentProperties", "externalBooks",
    "registeredFunctions", "removeDefinedName", "sheetIndex", "sheetNames",
    "toOds", "toXls", "unregisterFunction", "workbookProtection",
  ];

  /** "A" -> 1, "AB" -> 28. */
  const columnNumber = (letters) =>
    [...letters.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  /** 1 -> "A", 28 -> "AB". */
  const columnLetters = (n) => {
    let s = "";
    for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  };
  /** "B3" -> { row: 3, column: 2 }. */
  const cell = (address) => {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(address.trim());
    if (!m) throw new Error(`not a cell address: ${address}`);
    return { row: Number(m[2]), column: columnNumber(m[1]) };
  };
  const a1 = (row, column) => `${columnLetters(column)}${row}`;
  /** "A1:C6" -> its corners; a single cell is both. */
  const corners = (address) => {
    const [from, to = from] = address.split(":");
    return { start: cell(from), end: cell(to) };
  };

  class Workbook {
    /** A new workbook, or one over a `Book` already in hand. */
    constructor(book = new Book()) {
      this.book = book;
      // A new Book comes with "Sheet1". The first `addSheet` takes it over
      // instead of leaving an empty sheet in front of the one asked for.
      this._spare = arguments.length === 0;
    }

    /** Reads xlsx, xls, xlsb, ods, csv and the rest, as `Book.read` does. */
    static read(bytes, name, maxExpanded, onProgress, password) {
      return new Workbook(Book.read(bytes, name, maxExpanded, onProgress, password));
    }

    /** Reads CSV in a stated shape, as `Book.readCsv` does. */
    static readCsv(bytes, options) {
      return new Workbook(Book.readCsv(bytes, options));
    }

    /** Hands one sheet over row by row without keeping its grid, as `Book.forEachRow` does. */
    static forEachRow(bytes, name, sheet, callback, formatted) {
      return new Workbook(Book.forEachRow(bytes, name, sheet, callback, formatted));
    }

    /** A sheet by position or by name. */
    sheet(which) {
      this._spare = false;
      const index = typeof which === "number" ? which : this.book.sheetIndex(which);
      if (index === undefined || index < 0 || index >= this.book.sheetNames().length) {
        throw new Error(`no sheet ${JSON.stringify(which)}`);
      }
      return new Sheet(this, index);
    }

    /** Every sheet, in tab order. */
    sheets() {
      return this.book.sheetNames().map((_, i) => this.sheet(i));
    }

    /** Adds a sheet at the end and returns it. */
    addSheet(name = `Sheet${this.book.sheetNames().length + 1}`) {
      if (this._spare) {
        this._spare = false;
        this.book.renameSheet(0, name);
        return new Sheet(this, 0);
      }
      return new Sheet(this, this.book.addSheet(name));
    }

    /** Recalculates every formula. */
    recalculate(onProgress) {
      this.book.recalculate(undefined, onProgress);
      return this;
    }

    toXlsx(onProgress) {
      return this.book.toXlsx(onProgress);
    }

    /** Frees the wasm memory behind the workbook. */
    free() {
      this.book.free();
    }
  }

  class Sheet {
    constructor(workbook, index) {
      this.workbook = workbook;
      this.index = index;
    }

    get book() {
      return this.workbook.book;
    }

    get name() {
      return this.book.sheetNames()[this.index];
    }

    rename(name) {
      this.book.renameSheet(this.index, name);
      return this;
    }

    /** What the cell shows, through its number format. */
    text(address) {
      return this.book.getFormatted(this.index, address);
    }

    /** A block of values whose top left corner lands on `at`. */
    write(at, rows) {
      this.book.setRange(this.index, at, rows);
      return this;
    }

    /** A range of this sheet: `"A1:C6"`, or one cell. */
    range(address) {
      return new SheetRange(this, address);
    }

    /** Lays a style patch over every cell of a range. */
    style(address, patch) {
      this.book.setRangeStyle(this.index, address, patch);
      return this;
    }

    /** Column width in characters; the column is a letter or a number. */
    width(column, characters) {
      const n = typeof column === "number" ? column : columnNumber(column);
      this.book.setColumnWidth(this.index, n, characters);
      return this;
    }

    /** Freezes the rows above and the columns to the left; zeros unfreeze. */
    freeze(rows, columns = 0) {
      this.book.freezePanes(this.index, rows, columns);
      return this;
    }

    /** Draws a table over a range already filled; its first row is the header. */
    addTable(name, address, { header = true } = {}) {
      this.book.addTable(this.index, name, address, header);
      return new Table(this, name);
    }

    /**
     * Writes `rows` from `at` and draws a table over them; the first row is
     * the header.
     */
    addTableFromData(name, at, rows) {
      if (!rows.length || !rows[0].length) throw new Error("a table needs a header row");
      const { row, column } = cell(at);
      const width = Math.max(...rows.map((r) => r.length));
      this.write(at, rows);
      return this.addTable(name, `${at}:${a1(row + rows.length - 1, column + width - 1)}`);
    }

    /** A table on this sheet, by name. */
    table(name) {
      if (!this.book.tables(this.index).some((t) => t.displayName.toLowerCase() === name.toLowerCase())) {
        throw new Error(`no table ${name} on sheet ${this.name}`);
      }
      return new Table(this, name);
    }

    tables() {
      return this.book.tables(this.index).map((t) => new Table(this, t.displayName));
    }

    /** Makes this the sheet a workbook opens on. */
    activate() {
      this.book.setActiveSheet(this.index);
      return this;
    }

    /** Recalculates the formulas of this sheet. */
    recalculate(onProgress) {
      this.book.recalculate(this.index, onProgress);
      return this;
    }

    /**
     * Removes the sheet; references to it become `#REF!`. The sheets after
     * it move down an index, so handles to them are stale. Returns the
     * workbook.
     */
    remove() {
      this.book.removeSheet(this.index);
      return this.workbook;
    }
  }

  class Table {
    constructor(sheet, name) {
      this.sheet = sheet;
      this.name = name;
    }

    get book() {
      return this.sheet.book;
    }

    /** What `Book.tables` says about this table. */
    info() {
      const found = this.book
        .tables(this.sheet.index)
        .find((t) => t.displayName.toLowerCase() === this.name.toLowerCase());
      if (!found) throw new Error(`table ${this.name} is gone`);
      return found;
    }

    get columns() {
      return this.info().columns;
    }

    /** The whole table, header and totals included, or a range of the sheet. */
    range(address) {
      return address ? this.sheet.range(address) : this.sheet.range(this.info().range);
    }

    /** The data rows, without the header and totals. */
    data() {
      const { range, headerRowCount, totalsRowCount } = this.info();
      const { start, end } = corners(range);
      const top = start.row + (headerRowCount ?? 1);
      const bottom = end.row - (totalsRowCount ?? 0);
      return this.sheet.range(`${a1(top, start.column)}:${a1(Math.max(top, bottom), end.column)}`);
    }

    /** The sheet address of a data cell: `row` from 1, `column` a header or a number from 1. */
    address(row, column) {
      const { range, headerRowCount } = this.info();
      const { start } = corners(range);
      const offset = typeof column === "number" ? column - 1 : this.columnIndex(column);
      return a1(start.row + (headerRowCount ?? 1) + row - 1, start.column + offset);
    }

    columnIndex(header) {
      const i = this.columns.findIndex((c) => c.toLowerCase() === String(header).toLowerCase());
      if (i < 0) throw new Error(`table ${this.name} has no column ${JSON.stringify(header)}`);
      return i;
    }

    /** Sets one data cell: `table.set(2, "Amount", 340)`. */
    set(row, column, value) {
      this.sheet.set(this.address(row, column), value);
      return this;
    }

    get(row, column) {
      return this.sheet.get(this.address(row, column));
    }

    /** Filters by a column and hides the rows that fail, as Excel does. */
    addFilter(column, criteria) {
      this.book.setTableFilter(this.name, column, criteria);
      return this;
    }

    /** Takes the criterion off one column, or off all of them. */
    clearFilter(column) {
      const columns = column === undefined ? this.columns : [column];
      for (const c of columns) this.book.setTableFilter(this.name, c, null);
      return this;
    }

    /** Sorts the data rows: `["Region", "-Amount"]`, a minus for largest first. */
    sort(keys) {
      this.book.sortTable(this.name, keys);
      return this;
    }

    /**
     * The data rows as objects keyed by header. `{ visible: true }` leaves
     * out the rows a filter hides.
     */
    records({ visible = false } = {}) {
      const columns = this.columns;
      const data = this.data();
      const rows = visible ? data.visibleValues() : data.values();
      return rows.map((r) => Object.fromEntries(columns.map((c, i) => [c, r[i]])));
    }
  }

  class SheetRange {
    constructor(sheet, address) {
      this.sheet = sheet;
      this.address = address;
    }

    get book() {
      return this.sheet.book;
    }

    /** The values, a row per array. */
    values() {
      return this.book.getRange(this.sheet.index, this.address);
    }

    /** The values of the rows a filter or a user has not hidden. */
    visibleValues() {
      const { start } = corners(this.address);
      return this.values().filter((_, i) => !this.sheet.rowHidden(start.row + i));
    }

    /** Writes a block from the top left corner of the range. */
    set(rows) {
      this.sheet.write(this.address.split(":")[0], rows);
      return this;
    }

    style(patch) {
      this.sheet.style(this.address, patch);
      return this;
    }

    merge() {
      this.book.merge(this.sheet.index, this.address);
      return this;
    }
  }

  // Every Book method that takes the sheet first is a Sheet method without
  // it, and every one that takes no sheet a Workbook method. What Book
  // answers with nothing returns the object, so it chains; the rest return
  // the answer. The lists follow Book's declarations; a test in npm/test.js
  // fails when Book gains a method neither covers.
  const delegate = (target, names, chain, args) => {
    for (const name of names) {
      Object.defineProperty(target.prototype, name, {
        configurable: true,
        writable: true,
        value: chain
          ? function (...rest) {
              this.book[name](...args(this), ...rest);
              return this;
            }
          : function (...rest) {
              return this.book[name](...args(this), ...rest);
            },
      });
    }
  };
  const onSheet = (sheet) => [sheet.index];
  const onBook = () => [];
  delegate(Sheet, SHEET_ACTIONS, true, onSheet);
  delegate(Sheet, SHEET_QUERIES, false, onSheet);
  delegate(Workbook, BOOK_ACTIONS, true, onBook);
  delegate(Workbook, BOOK_QUERIES, false, onBook);

  return { Workbook, Sheet, Table, SheetRange };
}
