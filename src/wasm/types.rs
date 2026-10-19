//! The TypeScript names for what crosses the boundary as a `JsValue`.
//!
//! `wasm-bindgen` would otherwise type these `any`, which is no help to a
//! caller: a cell is one of four things and nothing else.

use wasm_bindgen::prelude::*;

#[wasm_bindgen(typescript_custom_section)]
const TYPES: &'static str = r#"
/** What a cell can hold on the JS side. An error cell reads as its code, `#DIV/0!`. */
export type CellValue = number | string | boolean | null;
/** A rectangle of cells, row by row, as `getRange` returns it. */
export type CellGrid = CellValue[][];
/**
 * One row with its styling, as `getRowAt` returns it. `formatted` is `null`
 * when the call asked for values alone. `bold` is one byte per cell, 0 or 1.
 */
export interface SheetRow {
  values: CellValue[];
  formatted: string[] | null;
  bold: Uint8Array;
  indent: Uint32Array;
  hidden: boolean;
}
/** Where a drawing object sits, in the 1-based numbers a user sees. */
export interface ObjectAnchor {
  kind: "twoCell" | "oneCell" | "absolute";
  row: number | null;
  column: number | null;
}
/** How a sheet is frozen and shown, as `sheetView` returns it. */
export interface SheetViewInfo {
  frozenRows: number;
  frozenColumns: number;
  zoom: number | null;
  showGridLines: boolean;
  showRowColHeaders: boolean;
  rightToLeft: boolean;
  topLeftCell: string | null;
}
/** One of the user's own document properties. */
export interface CustomProperty {
  name: string;
  type: "text" | "integer" | "number" | "boolean" | "date";
  value: string | number | boolean;
}
/** Title, author, dates and the user's fields, as `documentProperties` returns them. Dates are ISO 8601. */
export interface DocumentProperties {
  title: string | null;
  subject: string | null;
  creator: string | null;
  keywords: string | null;
  description: string | null;
  lastModifiedBy: string | null;
  category: string | null;
  contentStatus: string | null;
  language: string | null;
  identifier: string | null;
  revision: string | null;
  version: string | null;
  created: string | null;
  modified: string | null;
  lastPrinted: string | null;
  company: string | null;
  manager: string | null;
  custom: CustomProperty[];
}
/** A name that stands for a formula, as `definedNames` returns it. */
export interface WorkbookName {
  name: string;
  formula: string;
  sheet: number | null;
  hidden: boolean;
}
/** A rule limiting what a cell accepts, as `dataValidations` returns it. */
export interface SheetValidation {
  sqref: string[];
  type: string;
  operator: string;
  formula1: string;
  formula2: string;
  allowBlank: boolean;
  showDropDown: boolean;
}
/** One conditional formatting block, as `conditionalFormats` returns it. */
export interface SheetConditionalFormat {
  sqref: string[];
  rules: { type: string; priority: number; operator: string | null; formulas: string[]; text: string | null }[];
}
/** The autofilter over a range, as `autoFilter` returns it. */
export interface SheetAutoFilter {
  range: string;
  columns: { colId: number; kind: "values" | "custom" | "dynamic" | "top10" | "color" | "icon" | "none" }[];
}
/** A pivot report on the sheet, as `pivotTables` returns it. */
export interface SheetPivotTable {
  name: string;
  location: string | null;
  cacheId: number;
  rowFields: number;
  columnFields: number;
  valueFields: number;
}
/** How a sheet or the workbook is locked, as `sheetProtection` returns it. */
export interface ProtectionInfo { locked: boolean; hasPassword: boolean }
/** How `toCsv` writes, and how `Book.readCsv` reads. */
export interface CsvOptions {
  delimiter?: string;
  contiguous?: boolean;
  preserveEmptyFields?: boolean;
}
/** A note on a cell, as `comments` returns it. */
export interface SheetComment { address: string; author: string; text: string }
/** A link over a cell or a block of them, as `hyperlinks` returns it. */
export interface SheetHyperlink {
  range: string;
  target: string;
  external: boolean;
  display: string | null;
  tooltip: string | null;
}
/** A table over a block of cells, as `tables` returns it. */
export interface SheetTable {
  name: string;
  displayName: string;
  range: string;
  headerRowCount: number | null;
  totalsRowCount: number | null;
  columns: string[];
}
/** A chart on the sheet, as `charts` returns it. */
export interface SheetChart {
  name: string;
  title: string | null;
  kinds: string[];
  seriesCount: number;
  anchor: ObjectAnchor;
  /** The chart area's fill and outline; `null` leaves them to the chart style. */
  format: DrawingFormat | null;
  /** The plot area's, the same way. */
  plotFormat: DrawingFormat | null;
}
/** A picture on the sheet, as `images` returns it. Bytes come from `imageData`. */
export interface SheetImage {
  name: string;
  description: string;
  format: string;
  byteLength: number;
  anchor: ObjectAnchor;
}
/** A drawn shape, as `shapes` returns it. */
export interface SheetShape {
  name: string;
  description: string;
  geometry: string | null;
  text: string;
  anchor: ObjectAnchor;
  /** Degrees clockwise about the centre. */
  rotation: number;
  /** Mirrored left to right, before it is turned. */
  flipH: boolean;
  /** Mirrored top to bottom, before it is turned. */
  flipV: boolean;
  /** What Excel fills it with: its own, else its style's. */
  fill: DrawingFill;
  /** Its outline, the same way. */
  line: DrawingLine;
  /** The font of its text as the first run states it, the colour from the style if not. */
  font: RunFont;
}
/**
 * A colour of a drawing, resolved through the workbook theme to `#AARRGGBB` (`FF` for opaque);
 * `null` for one no theme defines.
 */
export type DrawingColor = string | null;
/** How a drawing object or chart area is filled. `position` is in percent, `angle` in degrees. */
export type DrawingFill =
  | { type: "none" }
  | { type: "solid"; color: DrawingColor }
  | {
      type: "gradient";
      stops: { position: number; color: DrawingColor }[];
      angle: number | null;
      path: "shape" | "circle" | "rect" | null;
    }
  | { type: "pattern"; preset: string | null; foreground: DrawingColor; background: DrawingColor }
  | { type: "other" };
/** An outline: `fill` is `null` and `width` (points) is `null` where the file is silent. */
export interface DrawingLine {
  fill: DrawingFill | null;
  width: number | null;
}
/** The fill and outline an element states; `null` in either leaves it to the style. */
export interface DrawingFormat {
  fill: DrawingFill | null;
  line: DrawingLine | null;
}
/**
 * A colour as the file states it: `null` when the file leaves it to the
 * reader, `#AARRGGBB` when it names one outright, and `indexed:N` or
 * `theme:N` when it points into the legacy palette or the workbook theme -
 * a theme colour carries its tint as `theme:4@-0.25`.
 */
export type StyleColor = string | null;
/**
 * What a run of formatted text changes about the cell's font; a field left
 * out is the cell's own. `size` in points.
 */
export interface RunFont {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  strike?: boolean;
  color?: StyleColor;
  script?: "superscript" | "subscript" | "baseline";
}
/** One run of formatted text, as `getRichText` returns it. */
export interface TextRun { text: string; font?: RunFont | null }
/** One side of a cell's border. */
export interface BorderSide { style: string; color: StyleColor }
/** How a cell is painted, as `cellStyle` returns it. */
export interface CellStyle {
  numberFormat: string;
  font: {
    name: string;
    size: number;
    bold: boolean;
    italic: boolean;
    underline: string;
    strike: boolean;
    color: StyleColor;
  };
  fill: { pattern: string; foreground: StyleColor; background: StyleColor };
  borders: {
    left: BorderSide;
    right: BorderSide;
    top: BorderSide;
    bottom: BorderSide;
  };
  alignment: {
    horizontal: string | null;
    vertical: string | null;
    wrapText: boolean;
    shrinkToFit: boolean;
    indent: number;
    textRotation: number;
  };
}
/**
 * The styles of a rectangle, as `getRangeStyles` returns them: each distinct
 * style once in `styles`, and `grid` - row by row - pointing into it.
 */
export interface RangeStyles {
  styles: CellStyle[];
  grid: number[][];
}
/** Limits for `goalSeek`; the workbook's iteration settings when left out. */
export interface GoalSeekOptions {
  maxIterations?: number;
  tolerance?: number;
}
/** How `goalSeek` ended. */
export interface GoalSeekResult {
  /** The value left in the changing cell: the best one tried. */
  value: number;
  /** What the set cell computes with it; `null` when that is not a number. */
  result: number | null;
  /** How many times the set cell was computed. */
  iterations: number;
  /** Whether `result` is within the tolerance of the target. */
  converged: boolean;
}
/** The rule a value breaks, as `checkValidation` reports it. */
export interface ValidationProblem {
  /** How Excel reports it: `stop` refuses the entry, the others only warn. */
  style: "stop" | "warning" | "information";
  title: string;
  /** The rule's own message; empty when it has none. */
  message: string;
  sqref: string[];
}
"#;

/// The shapes only the writing build takes, so a read-only package does not
/// declare types for methods it does not have.
#[cfg(feature = "write")]
#[wasm_bindgen(typescript_custom_section)]
const WRITE_TYPES: &'static str = r#"
/**
 * What `setDocumentProperties` takes: a named field is set, `null` clears it,
 * a missing one stays. `custom` replaces the whole list; `type` is optional.
 */
export type DocumentPropertiesPatch = Partial<
  Record<Exclude<keyof DocumentProperties, "custom">, string | null>
> & {
  custom?: { name: string; value: string | number | boolean; type?: CustomProperty["type"] }[];
};
/**
 * A style to write, as `setCellStyle` takes it. Every field is optional and
 * what is left out keeps the value the cell had, so `{ font: { bold: true } }`
 * makes a cell bold without touching its number format.
 */
export interface CellStylePatch {
  numberFormat?: string;
  font?: Partial<{
    name: string;
    size: number;
    bold: boolean;
    italic: boolean;
    underline: string;
    strike: boolean;
    color: StyleColor;
  }>;
  fill?: Partial<{ pattern: string; foreground: StyleColor; background: StyleColor }>;
  borders?: Partial<{
    left: Partial<BorderSide>;
    right: Partial<BorderSide>;
    top: Partial<BorderSide>;
    bottom: Partial<BorderSide>;
  }>;
  alignment?: Partial<{
    horizontal: string | null;
    vertical: string | null;
    wrapText: boolean;
    shrinkToFit: boolean;
    indent: number;
    textRotation: number;
  }>;
}
/**
 * Sort keys: a 1-based column number of the sheet (a row number when sorting
 * columns) or a header's text; a minus in front sorts largest first.
 */
export type SortKeys = (number | string)[];
/** How `sortRange` reads its range. */
export interface SortRangeOptions {
  /** The first row (or column) is a header: it stays put and names keys. */
  header?: boolean;
  /** Reorder columns left to right instead of rows. */
  byColumns?: boolean;
}
/**
 * Styles to write over a rectangle, as `setRangeStyles` takes them: `grid`
 * points into `styles`, and `null` or `-1` leaves its cell as it is. What
 * `getRangeStyles` returns fits here unchanged.
 */
export interface RangeStylesPatch {
  styles: CellStylePatch[];
  grid: (number | null)[][];
}
/**
 * A colour to draw with: `#RRGGBB`, `#AARRGGBB` (the alpha is the opacity) or
 * a theme colour by name - `accent1` to `accent6`, `tx1`, `tx2`, `bg1`,
 * `bg2`, `hlink`, `folHlink` - which follows the workbook's theme.
 */
export type DrawingColorInput = string;
/** A fill to write: the form `shapes` reads, without `other`. */
export type DrawingFillInput =
  | { type: "none" }
  | { type: "solid"; color: DrawingColorInput }
  | {
      type: "gradient";
      stops: { position: number; color: DrawingColorInput }[];
      angle?: number | null;
      path?: "shape" | "circle" | "rect" | null;
    }
  | { type: "pattern"; preset?: string | null; foreground?: DrawingColorInput; background?: DrawingColorInput };
/**
 * Fill and outline to lay over an object's own: a field left out stays,
 * `null` gives it back to the style. `line.width` is in points.
 */
export interface DrawingFormatPatch {
  fill?: DrawingFillInput | null;
  line?: { fill?: DrawingFillInput | null; width?: number | null } | null;
}
/** What `setShapeFormat` takes. `rotation` is in degrees clockwise. */
export interface ShapeFormatPatch extends DrawingFormatPatch {
  font?: Pick<RunFont, "name" | "size" | "bold" | "italic" | "color">;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}
/**
 * What a table column keeps, as `setTableFilter` takes it: the shown values,
 * one or two comparisons, or the largest or smallest few.
 */
export type TableFilter =
  | { values: (string | number)[]; blank?: boolean }
  | { custom: { op: "=" | "<>" | ">" | ">=" | "<" | "<="; value: string | number }[]; and?: boolean }
  | { top: number; percent?: boolean }
  | { bottom: number; percent?: boolean };
/** What `setChartFormat` takes: the chart area and the plot area. */
export interface ChartFormatPatch {
  format?: DrawingFormatPatch | null;
  plotFormat?: DrawingFormatPatch | null;
}
/** What `removeDuplicates` did. */
export interface DuplicatesRemoved {
  /** Rows dropped as repeats of a row above them. */
  removed: number;
  /** Data rows left, the header not counted. */
  remaining: number;
}
/** How `removeDuplicates` compares rows. */
export interface RemoveDuplicatesOptions {
  /** 1-based sheet columns to compare; every column of the range when left out. */
  columns?: number[];
  /** The first row is a header: it stays and is not compared. */
  header?: boolean;
}
/** How `textToColumns` cuts text: by `delimiters` or, when given, by `widths`. */
export interface TextToColumnsOptions {
  /** Delimiter characters, each one a delimiter: `";"`, `",\t"`. A tab when left out. */
  delimiters?: string;
  /** Character positions to cut before, for fixed width. */
  widths?: number[];
  /** A run of delimiters counts as one. */
  mergeConsecutive?: boolean;
  /** The quote inside which delimiters are text; `"` when left out, `null` for none. */
  qualifier?: string | null;
  /** The first cell written; the range's own first cell when left out. */
  destination?: string;
}
/** A rule for `addDataValidation`. Formulas are written for the first cell of `sqref`. */
export interface NewDataValidation {
  /** The cells, space-separated: `"B2:B100"`, `"B2:B9 D2:D9"`. */
  sqref: string;
  type: "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
  operator?:
    | "between" | "notBetween" | "equal" | "notEqual"
    | "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual";
  /** First bound, the list (`"\"Yes,No\""` or `"$E$1:$E$5"`) or the custom formula. */
  formula1?: string;
  /** Second bound, for `between` and `notBetween`. */
  formula2?: string;
  /** Empty cells pass. On when left out. */
  allowBlank?: boolean;
  /** The in-cell arrow of a list. On when left out. */
  showDropDown?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: "stop" | "warning" | "information";
  errorTitle?: string;
  error?: string;
  promptTitle?: string;
  prompt?: string;
}
"#;
