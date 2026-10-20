//! The document model: workbook, sheet, cell.
//!
//!
//! Ownership is a tree: the workbook owns sheets, a sheet owns its cells.
//! There are no back-references from a cell to its sheet; a method that needs
//! sheet context takes it as a parameter. That removes every reason to reach
//! for `Rc<RefCell<_>>`.

pub mod autofilter;
pub mod chart;
pub mod image;
pub mod pivot;
pub mod properties;
pub mod protection;
pub mod scenario;
pub mod shape;
pub mod sparkline;
pub mod table;

use crate::coordinate::{CellRef, Col, Range, Row};
use crate::error::{CellError, Error, Result};
use crate::shared::date::Epoch;
use crate::style::{Color, DiffFont, StyleId, StyleTable};
use std::collections::BTreeMap;
use std::sync::Arc;

pub use autofilter::{
    AutoFilter, ColumnFilter, CustomFilter, DateGroup, FilterColumn, FilterOperator, SortCondition,
    SortState,
};
pub use properties::{CustomProperty, DocumentProperties, PropertyValue};
pub use protection::{PasswordHash, ProtectedRange, SheetProtection, WorkbookProtection};

/// Longest string a cell can hold (`DataType::MAX_STRING_LENGTH`).
pub const MAX_STRING_LENGTH: usize = 32_767;

/// Longest sheet name Excel accepts.
pub const SHEET_TITLE_MAX_LENGTH: usize = 31;

/// Characters banned from sheet names (`Worksheet::INVALID_CHARACTERS`).
pub const SHEET_TITLE_INVALID_CHARS: [char; 7] = ['*', ':', '/', '\\', '?', '[', ']'];

/// The value held by a cell.
///
/// Replaces the string type tags the formats use (`'n'`, `'s'`, `'f'`, ...):
/// what was a convention there is checked by the compiler here.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum CellValue {
    /// An empty cell.
    #[default]
    Empty,
    /// A number. Dates are numbers too - the cell's format is what makes one a
    /// date.
    Number(f64),
    /// A string. `TYPE_STRING`, `TYPE_STRING2` and `TYPE_INLINE` differ only in
    /// how they are written to a file, which is the writer's business, not the
    /// model's.
    ///
    /// Shared rather than owned: a sheet of seven million text cells holds
    /// perhaps eight hundred thousand distinct strings, and a copy per cell
    /// was most of the memory a large workbook took. Cells read from the
    /// same shared string point at one allocation.
    Text(Arc<str>),
    /// A boolean.
    Bool(bool),
    /// An error code held as the cell's value.
    Error(CellError),
    /// Text whose formatting changes part way through.
    RichText(Vec<TextRun>),
    /// A formula, plus its last known result.
    Formula {
        /// Formula text without the leading `=`.
        formula: String,
        /// Cached result. It is stored in the file, which is what lets a
        /// workbook be read without recalculating it.
        cached: Option<Box<Self>>,
    },
}

impl CellValue {
    /// Text, truncated to Excel's limit with line endings normalised.
    ///
    #[must_use]
    pub fn text(value: impl Into<String>) -> Self {
        let s = value.into();
        // Excel counts characters, not bytes, so truncate on a char boundary.
        let s = match s.char_indices().nth(MAX_STRING_LENGTH) {
            Some((byte_idx, _)) => s[..byte_idx].to_owned(),
            None => s,
        };
        Self::Text(s.into())
    }

    /// Text already shared with other cells, as a reader hands out the
    /// entries of a shared string table. The caller has kept it to Excel's
    /// length.
    #[must_use]
    pub const fn shared_text(value: Arc<str>) -> Self {
        Self::Text(value)
    }

    /// Whether the cell holds nothing.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    /// The text of the value with any run formatting dropped, for the places
    /// that only care what it says.
    #[must_use]
    pub fn plain_text(&self) -> Option<String> {
        match self {
            Self::Text(t) => Some(t.to_string()),
            Self::RichText(runs) => Some(runs.iter().map(|r| r.text.as_str()).collect()),
            _ => None,
        }
    }

    /// A formula with no result yet, which is what a formula written by hand
    /// is until something computes it.
    ///
    /// ```
    /// use excelerate::model::CellValue;
    /// # use excelerate::model::{Spreadsheet, Worksheet};
    /// # use excelerate::CellRef;
    /// # let mut sheet = Worksheet::new("Sheet1")?;
    /// sheet.set(CellRef::parse("C1")?, CellValue::formula("A1*B1"));
    /// # Ok::<(), excelerate::Error>(())
    /// ```
    #[must_use]
    pub fn formula(text: impl Into<String>) -> Self {
        Self::Formula {
            formula: text.into(),
            cached: None,
        }
    }

    /// What the cell holds, looking through a formula to its result.
    ///
    /// A formula cell holds two things - the text and the value last computed
    /// for it - and code that reads data wants the second. A formula with no
    /// cached result answers [`CellValue::Empty`], the way an untouched cell
    /// does; whether it was ever computed is [`CellValue::Formula`]'s own
    /// business.
    ///
    /// ```
    /// use excelerate::model::CellValue;
    ///
    /// let cell = CellValue::Formula {
    ///     formula: "1+2".to_owned(),
    ///     cached: Some(Box::new(CellValue::Number(3.0))),
    /// };
    /// assert_eq!(cell.result(), &CellValue::Number(3.0));
    /// assert_eq!(CellValue::Number(7.0).result(), &CellValue::Number(7.0));
    /// ```
    #[must_use]
    pub fn result(&self) -> &Self {
        match self {
            Self::Formula { cached, .. } => cached.as_deref().unwrap_or(&Self::Empty),
            other => other,
        }
    }

    /// The number the cell holds, or the number its formula computed.
    ///
    /// `None` for everything else, a boolean included: Excel stores `TRUE` as
    /// its own type rather than as one, and a caller asking for a number is
    /// asking about a cell it believes holds one.
    ///
    /// ```
    /// use excelerate::model::CellValue;
    ///
    /// assert_eq!(CellValue::Number(2.5).as_number(), Some(2.5));
    /// assert_eq!(CellValue::Bool(true).as_number(), None);
    /// ```
    #[must_use]
    pub fn as_number(&self) -> Option<f64> {
        match self.result() {
            Self::Number(n) => Some(*n),
            _ => None,
        }
    }

    /// The text the cell holds, or the text its formula computed.
    ///
    /// Rich text has no single string to borrow, so it answers `None` here;
    /// [`CellValue::plain_text`] builds one for both kinds.
    #[must_use]
    pub fn as_str(&self) -> Option<&str> {
        match self.result() {
            Self::Text(t) => Some(t),
            _ => None,
        }
    }

    /// The boolean the cell holds, or the one its formula computed.
    #[must_use]
    pub fn as_bool(&self) -> Option<bool> {
        match self.result() {
            Self::Bool(b) => Some(*b),
            _ => None,
        }
    }

    /// The error in the cell, or the one its formula answered with.
    #[must_use]
    pub fn as_error(&self) -> Option<crate::CellError> {
        match self.result() {
            Self::Error(e) => Some(*e),
            _ => None,
        }
    }

    /// The value as a number format shows it.
    ///
    /// A formula shows its result. `TRUE` and `FALSE` and the error codes are
    /// not run through the format: Excel does not format them either.
    /// [`Spreadsheet::formatted`] is this with the cell's own format and the
    /// workbook's epoch already looked up, which is what a caller usually
    /// wants.
    #[must_use]
    pub fn display(&self, code: &str, epoch: crate::shared::date::Epoch) -> String {
        use crate::style::format::{Value as FormatValue, format};
        match self.result() {
            Self::Number(n) => format(FormatValue::Number(*n), code, epoch),
            Self::Text(t) => format(FormatValue::Text(t), code, epoch),
            Self::Bool(b) => (if *b { "TRUE" } else { "FALSE" }).to_owned(),
            Self::Error(e) => e.as_str().to_owned(),
            Self::RichText(runs) => {
                let text: String = runs.iter().map(|r| r.text.as_str()).collect();
                format(FormatValue::Text(&text), code, epoch)
            }
            Self::Empty | Self::Formula { .. } => String::new(),
        }
    }
}

impl From<f64> for CellValue {
    fn from(v: f64) -> Self {
        Self::Number(v)
    }
}

impl From<i64> for CellValue {
    #[expect(
        clippy::cast_precision_loss,
        reason = "Excel stores every number as f64"
    )]
    fn from(v: i64) -> Self {
        Self::Number(v as f64)
    }
}

impl From<i32> for CellValue {
    /// What an integer literal is by default, so `sheet.set(at, 123)` needs
    /// no suffix.
    fn from(v: i32) -> Self {
        Self::Number(f64::from(v))
    }
}

impl From<u32> for CellValue {
    fn from(v: u32) -> Self {
        Self::Number(f64::from(v))
    }
}

impl From<bool> for CellValue {
    fn from(v: bool) -> Self {
        Self::Bool(v)
    }
}

impl From<&str> for CellValue {
    fn from(v: &str) -> Self {
        Self::text(v)
    }
}

impl From<String> for CellValue {
    fn from(v: String) -> Self {
        Self::text(v)
    }
}

/// A stretch of text inside one cell, with the formatting that overrides the
/// cell's own.
///
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TextRun {
    /// The text of this run.
    pub text: String,
    /// What it changes about the cell's font. `None` means it changes nothing.
    pub font: Option<DiffFont>,
}

/// A note attached to a cell.
///
/// The text is rich, as a comment usually starts with the author's name in
/// bold. The box it is drawn in lives in the sheet's VML drawing, which
/// travels as bytes; whether the box shows and how big it is are read out of
/// it, and the writer changes just those in the bytes.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Comment {
    /// Who wrote it. Excel keeps the authors in a table and the comment points
    /// at one; here each comment carries its own, and writing rebuilds the
    /// table.
    pub author: String,
    /// The text, in runs.
    pub text: Vec<TextRun>,
    /// Shown all the time, rather than when the pointer is over the cell.
    pub visible: bool,
    /// Width and height of the box in points; `None` for Excel's default,
    /// 108 by 59.25, which is also how a box of that size reads.
    pub size: Option<(f64, f64)>,
}

impl Comment {
    /// The text with its formatting dropped.
    #[must_use]
    pub fn plain_text(&self) -> String {
        self.text.iter().map(|r| r.text.as_str()).collect()
    }
}

/// A cell: a value and a reference to a style.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Cell {
    /// The value.
    pub value: CellValue,
    /// Style from the workbook's style table.
    pub style: StyleId,
}

impl Cell {
    /// A cell holding `value`, styled with the default style.
    #[must_use]
    pub fn new(value: impl Into<CellValue>) -> Self {
        Self {
            value: value.into(),
            style: StyleId::default(),
        }
    }
}

/// Sizing and visibility of a run of columns.
///
#[expect(
    clippy::struct_excessive_bools,
    reason = "these are the independent flags of a <col> element, not a state \
              machine that an enum could replace"
)]
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnRun {
    /// First column of the run.
    pub first: Col,
    /// Last column of the run, inclusive.
    pub last: Col,
    /// Width in characters of the default font. `None` means the sheet default.
    pub width: Option<f64>,
    /// Whether the width was set by hand rather than derived from content.
    pub custom_width: bool,
    /// Whether the columns are hidden.
    pub hidden: bool,
    /// Whether the width follows the widest cell.
    pub best_fit: bool,
    /// Style applied to otherwise unstyled cells of these columns.
    pub style: Option<StyleId>,
    /// Grouping depth.
    pub outline_level: u8,
    /// Whether the group is collapsed.
    pub collapsed: bool,
}

impl ColumnRun {
    /// A run covering one column with a given width.
    #[must_use]
    pub fn new(first: Col, last: Col) -> Self {
        Self {
            first,
            last,
            width: None,
            custom_width: false,
            hidden: false,
            best_fit: false,
            style: None,
            outline_level: 0,
            collapsed: false,
        }
    }

    /// Whether the run says anything worth writing.
    #[must_use]
    pub fn is_meaningful(&self) -> bool {
        self.width.is_some()
            || self.hidden
            || self.best_fit
            || self.style.is_some()
            || self.outline_level > 0
            || self.collapsed
    }
}

/// Sizing and visibility of one row.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct RowProperties {
    /// Height in points. `None` means the sheet default.
    pub height: Option<f64>,
    /// Whether the height was set by hand.
    pub custom_height: bool,
    /// Whether the row is hidden.
    pub hidden: bool,
    /// Style applied to otherwise unstyled cells of this row.
    pub style: Option<StyleId>,
    /// Grouping depth.
    pub outline_level: u8,
    /// Whether the group is collapsed.
    pub collapsed: bool,
}

impl RowProperties {
    /// Whether the row says anything worth writing.
    #[must_use]
    pub fn is_meaningful(&self) -> bool {
        self.height.is_some()
            || self.hidden
            || self.style.is_some()
            || self.outline_level > 0
            || self.collapsed
    }
}

/// Which pane of a split sheet something belongs to.
///
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PanePosition {
    /// `topLeft`, and the whole sheet when it is not split.
    #[default]
    TopLeft,
    /// `topRight`.
    TopRight,
    /// `bottomLeft`.
    BottomLeft,
    /// `bottomRight`.
    BottomRight,
}

impl PanePosition {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::TopLeft => "topLeft",
            Self::TopRight => "topRight",
            Self::BottomLeft => "bottomLeft",
            Self::BottomRight => "bottomRight",
        }
    }

    /// Parses the xlsx name; anything unknown reads as `topLeft`, which is what
    /// the attribute defaults to.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "topRight" => Self::TopRight,
            "bottomLeft" => Self::BottomLeft,
            "bottomRight" => Self::BottomRight,
            _ => Self::TopLeft,
        }
    }
}

/// How a sheet's panes are divided.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PaneState {
    /// Panes can be dragged; the split is measured in twips.
    #[default]
    Split,
    /// Rows and columns before the split stay put; measured in cells.
    Frozen,
    /// Frozen, and the split can still be moved.
    FrozenSplit,
}

impl PaneState {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Split => "split",
            Self::Frozen => "frozen",
            Self::FrozenSplit => "frozenSplit",
        }
    }

    /// Parses the xlsx name, defaulting to `split` as the attribute does.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "frozen" => Self::Frozen,
            "frozenSplit" => Self::FrozenSplit,
            _ => Self::Split,
        }
    }
}

/// How the sheet is displayed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SheetViewType {
    /// The ordinary grid.
    #[default]
    Normal,
    /// Page layout, showing margins and headers.
    PageLayout,
    /// Page break preview.
    PageBreakPreview,
}

impl SheetViewType {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Normal => "normal",
            Self::PageLayout => "pageLayout",
            Self::PageBreakPreview => "pageBreakPreview",
        }
    }

    /// Parses the xlsx name, defaulting to `normal`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "pageLayout" => Self::PageLayout,
            "pageBreakPreview" => Self::PageBreakPreview,
            _ => Self::Normal,
        }
    }
}

/// A frozen or split view.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Pane {
    /// Columns before the vertical split; twips when the pane is not frozen.
    pub x_split: u32,
    /// Rows before the horizontal split; twips when the pane is not frozen.
    pub y_split: u32,
    /// First visible cell of the bottom-right pane.
    pub top_left_cell: Option<CellRef>,
    /// Which pane holds the cursor.
    pub active_pane: PanePosition,
    /// Whether the split is frozen.
    pub state: PaneState,
}

/// What is selected in one pane.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Selection {
    /// The pane this selection belongs to. `None` for an unsplit sheet.
    pub pane: Option<PanePosition>,
    /// The cell the cursor sits on.
    pub active_cell: Option<CellRef>,
    /// The selected areas. Excel writes them space-separated in one attribute.
    pub sqref: Vec<Range>,
}

/// The saved view state of a sheet: zoom, scroll position, selection, panes.
///
/// Excel restores all of this when the workbook is reopened, so dropping it
/// makes a rewritten file open in a visibly different state from the original.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the independent display switches of one <sheetView> element"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetView {
    /// Display mode.
    pub view: SheetViewType,
    /// Whether this sheet's tab is selected. Excel expects it on the sheet the
    /// workbook's active tab points at.
    pub tab_selected: bool,
    /// Zoom of the current view, in percent.
    pub zoom_scale: Option<u32>,
    /// Zoom remembered for the normal view.
    pub zoom_scale_normal: Option<u32>,
    /// Zoom remembered for the page layout view.
    pub zoom_scale_page_layout: Option<u32>,
    /// Zoom remembered for the page break preview.
    pub zoom_scale_sheet_layout: Option<u32>,
    /// First visible cell - the scroll position.
    pub top_left_cell: Option<CellRef>,
    /// Whether grid lines are drawn.
    pub show_grid_lines: bool,
    /// Whether row numbers and column letters are drawn.
    pub show_row_col_headers: bool,
    /// Whether a zero is shown as `0` rather than left blank.
    pub show_zeros: bool,
    /// Whether columns run right to left.
    pub right_to_left: bool,
    /// Whether cells show their formulas instead of their values (Ctrl and the backquote key).
    pub show_formulas: bool,
    /// Whether the outline bars and buttons of grouped rows and columns are
    /// drawn.
    pub show_outline_symbols: bool,
    /// Whether the ruler is drawn in page layout view.
    pub show_ruler: bool,
    /// Whether page layout view shows the margins between pages.
    pub show_white_space: bool,
    /// Whether the window of the sheet is locked in place.
    pub window_protection: bool,
    /// Grid line colour as an index into the palette, when the sheet does not
    /// use the default one: `colorId` with `defaultGridColor="0"`.
    pub grid_color: Option<u32>,
    /// Index of the workbook window this view belongs to.
    pub workbook_view_id: u32,
    /// Frozen or split panes, if any.
    pub pane: Option<Pane>,
    /// Selection, one entry per pane.
    pub selections: Vec<Selection>,
}

impl Default for SheetView {
    fn default() -> Self {
        Self {
            view: SheetViewType::default(),
            tab_selected: false,
            zoom_scale: None,
            zoom_scale_normal: None,
            zoom_scale_page_layout: None,
            zoom_scale_sheet_layout: None,
            top_left_cell: None,
            // The three switches below default to on in xlsx; the attribute is
            // only written when it is off.
            show_grid_lines: true,
            show_row_col_headers: true,
            show_zeros: true,
            right_to_left: false,
            show_formulas: false,
            show_outline_symbols: true,
            show_ruler: true,
            show_white_space: true,
            window_protection: false,
            grid_color: None,
            workbook_view_id: 0,
            pane: None,
            selections: Vec::new(),
        }
    }
}

/// What a data validation restricts a cell to.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationType {
    /// No restriction.
    #[default]
    None,
    /// A formula that must evaluate to true.
    Custom,
    /// A date.
    Date,
    /// Any number.
    Decimal,
    /// One of a list of values - the drop-down.
    List,
    /// Text of a bounded length.
    TextLength,
    /// A time of day.
    Time,
    /// A whole number.
    Whole,
}

impl ValidationType {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Custom => "custom",
            Self::Date => "date",
            Self::Decimal => "decimal",
            Self::List => "list",
            Self::TextLength => "textLength",
            Self::Time => "time",
            Self::Whole => "whole",
        }
    }

    /// Parses the xlsx name, defaulting to `none`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "custom" => Self::Custom,
            "date" => Self::Date,
            "decimal" => Self::Decimal,
            "list" => Self::List,
            "textLength" => Self::TextLength,
            "time" => Self::Time,
            "whole" => Self::Whole,
            _ => Self::None,
        }
    }
}

/// How a rejected entry is reported.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationErrorStyle {
    /// Refuse the entry.
    #[default]
    Stop,
    /// Warn, but let it through.
    Warning,
    /// Inform only.
    Information,
}

impl ValidationErrorStyle {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Stop => "stop",
            Self::Warning => "warning",
            Self::Information => "information",
        }
    }

    /// Parses the xlsx name, defaulting to `stop`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "warning" => Self::Warning,
            "information" => Self::Information,
            _ => Self::Stop,
        }
    }
}

/// How the two formulas bound the accepted value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ValidationOperator {
    /// Between the two formulas, inclusive.
    #[default]
    Between,
    /// Outside the two formulas.
    NotBetween,
    /// Equal to the first formula.
    Equal,
    /// Different from the first formula.
    NotEqual,
    /// Greater than the first formula.
    GreaterThan,
    /// Less than the first formula.
    LessThan,
    /// Greater than or equal to the first formula.
    GreaterThanOrEqual,
    /// Less than or equal to the first formula.
    LessThanOrEqual,
}

impl ValidationOperator {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThan => "greaterThan",
            Self::LessThan => "lessThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::LessThanOrEqual => "lessThanOrEqual",
        }
    }

    /// Parses the xlsx name, defaulting to `between`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "notBetween" => Self::NotBetween,
            "equal" => Self::Equal,
            "notEqual" => Self::NotEqual,
            "greaterThan" => Self::GreaterThan,
            "lessThan" => Self::LessThan,
            "greaterThanOrEqual" => Self::GreaterThanOrEqual,
            "lessThanOrEqual" => Self::LessThanOrEqual,
            _ => Self::Between,
        }
    }
}

/// A restriction on what may be typed into a range of cells.
///
#[expect(
    clippy::struct_excessive_bools,
    reason = "the independent flags of one <dataValidation> element"
)]
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DataValidation {
    /// The cells this applies to.
    pub sqref: Vec<Range>,
    /// What is accepted.
    pub kind: ValidationType,
    /// How the formulas bound the value.
    pub operator: ValidationOperator,
    /// First bound, or the list source, without a leading `=`.
    pub formula1: String,
    /// Second bound, for the two-sided operators.
    pub formula2: String,
    /// Whether an empty cell is accepted.
    pub allow_blank: bool,
    /// Whether the in-cell arrow is **hidden**. The attribute is named
    /// `showDropDown`, but Excel writes `1` to suppress the arrow - the name
    /// says the opposite of what it does, and the value is kept as the file
    /// spells it.
    pub hide_drop_down: bool,
    /// Whether the prompt is shown when the cell is selected.
    pub show_input_message: bool,
    /// Whether the error box is shown on a rejected entry.
    pub show_error_message: bool,
    /// Title of the error box.
    pub error_title: String,
    /// Body of the error box.
    pub error: String,
    /// Title of the prompt.
    pub prompt_title: String,
    /// Body of the prompt.
    pub prompt: String,
    /// How a rejected entry is reported.
    pub error_style: ValidationErrorStyle,
}

/// Properties of the sheet as a whole.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetProperties {
    /// Colour of the sheet's tab.
    pub tab_color: Option<Color>,
    /// Whether a group's summary row sits below the group rather than above.
    pub summary_below: bool,
    /// Whether a group's summary column sits to its right.
    pub summary_right: bool,
    /// Whether the sheet is scaled to fit a number of pages rather than by a
    /// percentage.
    pub fit_to_page: bool,
    /// The name VBA knows the sheet by.
    ///
    /// Renaming a sheet in Excel leaves this alone, which is the point of it:
    /// macros address sheets through it. Dropping it breaks them silently.
    pub code_name: Option<String>,
}

impl Default for SheetProperties {
    fn default() -> Self {
        Self {
            tab_color: None,
            // Both default to on in xlsx; the attribute only appears when off.
            summary_below: true,
            summary_right: true,
            fit_to_page: false,
            code_name: None,
        }
    }
}

/// Printing margins, in inches.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct PageMargins {
    /// Left margin.
    pub left: f64,
    /// Right margin.
    pub right: f64,
    /// Top margin.
    pub top: f64,
    /// Bottom margin.
    pub bottom: f64,
    /// Distance from the top of the page to the header.
    pub header: f64,
    /// Distance from the bottom of the page to the footer.
    pub footer: f64,
}

impl Default for PageMargins {
    /// Excel's own defaults for a new sheet.
    fn default() -> Self {
        Self {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3,
        }
    }
}

/// Which way round the paper goes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Orientation {
    /// Whatever the printer decides.
    #[default]
    Default,
    /// Taller than wide.
    Portrait,
    /// Wider than tall.
    Landscape,
}

impl Orientation {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Default => "default",
            Self::Portrait => "portrait",
            Self::Landscape => "landscape",
        }
    }

    /// Parses the xlsx name.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "portrait" => Self::Portrait,
            "landscape" => Self::Landscape,
            _ => Self::Default,
        }
    }
}

/// How the sheet is laid out on paper.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the independent switches of one <pageSetup> element"
)]
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct PageSetup {
    /// Paper size code; 1 is US Letter, 9 is A4.
    pub paper_size: Option<u32>,
    /// Which way round the paper goes.
    pub orientation: Orientation,
    /// Scale in percent, when the sheet is not fitted to a page count.
    pub scale: Option<u32>,
    /// How many pages wide the sheet is squeezed into. `Some(0)` means the
    /// count is not limited in that direction.
    pub fit_to_width: Option<u32>,
    /// How many pages tall.
    pub fit_to_height: Option<u32>,
    /// Number to give the first page.
    pub first_page_number: Option<u32>,
    /// Whether that number is used rather than the automatic one.
    pub use_first_page_number: bool,
    /// Package path of the printer settings part this page setup points at.
    ///
    /// The part itself travels unparsed - it is a Windows `DEVMODE` blob - but
    /// `<pageSetup>` has to keep pointing at it, or the settings are in the
    /// package and attached to nothing.
    pub printer_settings: Option<String>,
    /// Print resolution across.
    pub horizontal_dpi: Option<u32>,
    /// Print resolution down.
    pub vertical_dpi: Option<u32>,
    /// Whether pages run across before down.
    pub over_then_down: bool,
    /// Whether colours are dropped when printing.
    pub black_and_white: bool,
    /// Whether the sheet prints without its formatting.
    pub draft: bool,
}

/// What else goes on the printed page.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the four independent switches of one <printOptions> element"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct PrintOptions {
    /// Whether the sheet is centred left to right.
    pub horizontal_centered: bool,
    /// Whether it is centred top to bottom.
    pub vertical_centered: bool,
    /// Whether row numbers and column letters are printed.
    pub headings: bool,
    /// Whether grid lines are printed.
    pub grid_lines: bool,
}

/// Headers and footers.
///
/// Each string is Excel's own little markup: `&L`, `&C` and `&R` open the left,
/// centre and right section, `&P` is the page number, `&F` the file name. It is
/// kept as written rather than parsed - nothing here needs to understand it.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the independent switches of one <headerFooter> element"
)]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HeaderFooter {
    /// Whether even pages get their own header and footer.
    pub different_odd_even: bool,
    /// Whether the first page gets its own.
    pub different_first: bool,
    /// Whether they scale with the sheet.
    pub scale_with_doc: bool,
    /// Whether they line up with the page margins.
    pub align_with_margins: bool,
    /// Header of the odd pages, and of every page unless one of the flags
    /// above says otherwise.
    pub odd_header: String,
    /// Footer of the odd pages.
    pub odd_footer: String,
    /// Header of the even pages.
    pub even_header: String,
    /// Footer of the even pages.
    pub even_footer: String,
    /// Header of the first page.
    pub first_header: String,
    /// Footer of the first page.
    pub first_footer: String,
}

impl Default for HeaderFooter {
    fn default() -> Self {
        Self {
            different_odd_even: false,
            different_first: false,
            // Both default to on in xlsx.
            scale_with_doc: true,
            align_with_margins: true,
            odd_header: String::new(),
            odd_footer: String::new(),
            even_header: String::new(),
            even_footer: String::new(),
            first_header: String::new(),
            first_footer: String::new(),
        }
    }
}

impl HeaderFooter {
    /// Whether anything here is worth writing.
    #[must_use]
    pub fn is_meaningful(&self) -> bool {
        self.different_odd_even
            || self.different_first
            || !self.scale_with_doc
            || !self.align_with_margins
            || ![
                &self.odd_header,
                &self.odd_footer,
                &self.even_header,
                &self.even_footer,
                &self.first_header,
                &self.first_footer,
            ]
            .iter()
            .all(|s| s.is_empty())
    }
}

/// A page break the user put there.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PageBreak {
    /// Row or column the break sits before, 0-based.
    pub at: u32,
    /// How far the break reaches; `None` means the whole sheet.
    pub max: Option<u32>,
    /// Whether the user put it there rather than Excel.
    pub manual: bool,
}

/// Where a hyperlink points.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum LinkTarget {
    /// Somewhere in this workbook, as `'Sheet 2'!A1`.
    Inside(String),
    /// A URL or a file, which lives in the sheet's relationships.
    Outside(String),
}

/// A hyperlink over a cell or a block of them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Hyperlink {
    /// The cells that carry the link.
    pub range: Range,
    /// Where it points.
    pub target: LinkTarget,
    /// Text shown in place of the link, when it differs from the cell.
    pub display: Option<String>,
    /// Text of the tooltip.
    pub tooltip: Option<String>,
}

/// A name standing for a formula, usually a range.
///
/// Excel also keeps its own settings here under reserved names:
/// `_xlnm.Print_Area` is the print area, `_xlnm.Print_Titles` the rows and
/// columns repeated on every page.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DefinedName {
    /// The name as written.
    pub name: String,
    /// Tab index of the sheet it belongs to; `None` when it covers the whole
    /// workbook.
    pub sheet: Option<usize>,
    /// What it stands for, without a leading `=`.
    pub formula: String,
    /// Whether Excel hides it from the name manager.
    pub hidden: bool,
}

/// What a conditional formatting rule tests.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CfRuleType {
    /// The cell's value against one or two formulas.
    #[default]
    CellIs,
    /// A formula that must come out true.
    Expression,
    /// A gradient across two or three colours.
    ColorScale,
    /// A bar drawn inside the cell.
    DataBar,
    /// A little icon beside the value.
    IconSet,
    /// The highest or lowest few.
    Top10,
    /// Values that appear once.
    UniqueValues,
    /// Values that appear more than once.
    DuplicateValues,
    /// Text holding a fragment.
    ContainsText,
    /// Text not holding it.
    NotContainsText,
    /// Text starting with it.
    BeginsWith,
    /// Text ending with it.
    EndsWith,
    /// An empty cell.
    ContainsBlanks,
    /// A cell that is not empty.
    NotContainsBlanks,
    /// A cell holding an error.
    ContainsErrors,
    /// A cell not holding one.
    NotContainsErrors,
    /// A date inside a named span, such as `lastWeek`.
    TimePeriod,
    /// Above or below the average of the range.
    AboveAverage,
}

impl CfRuleType {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::CellIs => "cellIs",
            Self::Expression => "expression",
            Self::ColorScale => "colorScale",
            Self::DataBar => "dataBar",
            Self::IconSet => "iconSet",
            Self::Top10 => "top10",
            Self::UniqueValues => "uniqueValues",
            Self::DuplicateValues => "duplicateValues",
            Self::ContainsText => "containsText",
            Self::NotContainsText => "notContainsText",
            Self::BeginsWith => "beginsWith",
            Self::EndsWith => "endsWith",
            Self::ContainsBlanks => "containsBlanks",
            Self::NotContainsBlanks => "notContainsBlanks",
            Self::ContainsErrors => "containsErrors",
            Self::NotContainsErrors => "notContainsErrors",
            Self::TimePeriod => "timePeriod",
            Self::AboveAverage => "aboveAverage",
        }
    }

    /// Parses the xlsx name, defaulting to `cellIs`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "expression" => Self::Expression,
            "colorScale" => Self::ColorScale,
            "dataBar" => Self::DataBar,
            "iconSet" => Self::IconSet,
            "top10" => Self::Top10,
            "uniqueValues" => Self::UniqueValues,
            "duplicateValues" => Self::DuplicateValues,
            "containsText" => Self::ContainsText,
            "notContainsText" => Self::NotContainsText,
            "beginsWith" => Self::BeginsWith,
            "endsWith" => Self::EndsWith,
            "containsBlanks" => Self::ContainsBlanks,
            "notContainsBlanks" => Self::NotContainsBlanks,
            "containsErrors" => Self::ContainsErrors,
            "notContainsErrors" => Self::NotContainsErrors,
            "timePeriod" => Self::TimePeriod,
            "aboveAverage" => Self::AboveAverage,
            _ => Self::CellIs,
        }
    }
}

/// How a `cellIs` rule compares.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CfOperator {
    /// Between the two formulas.
    #[default]
    Between,
    /// Outside them.
    NotBetween,
    /// Equal to the first.
    Equal,
    /// Different from it.
    NotEqual,
    /// Greater than it.
    GreaterThan,
    /// Less than it.
    LessThan,
    /// Not less than it.
    GreaterThanOrEqual,
    /// Not greater than it.
    LessThanOrEqual,
    /// Text holding the fragment.
    ContainsText,
    /// Text not holding it.
    NotContains,
    /// Text starting with it.
    BeginsWith,
    /// Text ending with it.
    EndsWith,
}

impl CfOperator {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Between => "between",
            Self::NotBetween => "notBetween",
            Self::Equal => "equal",
            Self::NotEqual => "notEqual",
            Self::GreaterThan => "greaterThan",
            Self::LessThan => "lessThan",
            Self::GreaterThanOrEqual => "greaterThanOrEqual",
            Self::LessThanOrEqual => "lessThanOrEqual",
            Self::ContainsText => "containsText",
            Self::NotContains => "notContains",
            Self::BeginsWith => "beginsWith",
            Self::EndsWith => "endsWith",
        }
    }

    /// Parses the xlsx name, defaulting to `between`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "notBetween" => Self::NotBetween,
            "equal" => Self::Equal,
            "notEqual" => Self::NotEqual,
            "greaterThan" => Self::GreaterThan,
            "lessThan" => Self::LessThan,
            "greaterThanOrEqual" => Self::GreaterThanOrEqual,
            "lessThanOrEqual" => Self::LessThanOrEqual,
            "containsText" => Self::ContainsText,
            "notContains" => Self::NotContains,
            "beginsWith" => Self::BeginsWith,
            "endsWith" => Self::EndsWith,
            _ => Self::Between,
        }
    }
}

/// Where a stop of a scale sits.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CfValueType {
    /// A number written out.
    #[default]
    Number,
    /// A percentage of the span.
    Percent,
    /// A percentile of the values.
    Percentile,
    /// The lowest value in the range.
    Min,
    /// The highest.
    Max,
    /// A formula.
    Formula,
}

impl CfValueType {
    /// The name xlsx uses.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Number => "num",
            Self::Percent => "percent",
            Self::Percentile => "percentile",
            Self::Min => "min",
            Self::Max => "max",
            Self::Formula => "formula",
        }
    }

    /// Parses the xlsx name, defaulting to `num`.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "percent" => Self::Percent,
            "percentile" => Self::Percentile,
            "min" => Self::Min,
            "max" => Self::Max,
            "formula" => Self::Formula,
            _ => Self::Number,
        }
    }
}

/// One stop of a colour scale, data bar or icon set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CfValue {
    /// What the stop is measured in.
    pub kind: CfValueType,
    /// The value itself, as written.
    pub value: String,
    /// Whether the stop is inclusive; icon sets use this.
    pub greater_or_equal: bool,
}

impl Default for CfValue {
    fn default() -> Self {
        Self {
            kind: CfValueType::default(),
            value: String::new(),
            // `gte` defaults to true in xlsx and is written only when false.
            greater_or_equal: true,
        }
    }
}

/// The graphical part of a rule, for the three kinds that have one.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum CfScale {
    /// A gradient between two or three colours.
    Color {
        /// The stops.
        values: Vec<CfValue>,
        /// The colour at each stop.
        colors: Vec<Color>,
    },
    /// A bar drawn inside the cell.
    DataBar {
        /// The two ends of the bar.
        values: Vec<CfValue>,
        /// The bar's colour.
        color: Color,
        /// Shortest bar, in percent of the cell.
        min_length: Option<u32>,
        /// Longest bar.
        max_length: Option<u32>,
        /// Whether the number is shown beside the bar.
        show_value: bool,
    },
    /// An icon beside the value.
    IconSet {
        /// The thresholds.
        values: Vec<CfValue>,
        /// Which set of icons, such as `3TrafficLights1`.
        set: Option<String>,
        /// Whether the number is shown beside the icon.
        show_value: bool,
        /// Whether the thresholds are percentages.
        percent: bool,
        /// Whether the icons run the other way.
        reverse: bool,
    },
}

/// One conditional formatting rule.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the independent flags of one <cfRule> element"
)]
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct CfRule {
    /// What the rule tests.
    pub kind: CfRuleType,
    /// Which rule wins when several match; lower is stronger.
    pub priority: i32,
    /// Index into the workbook's differential styles, when the rule paints
    /// something.
    pub dxf: Option<u32>,
    /// Whether the rules after this one are skipped when it matches.
    pub stop_if_true: bool,
    /// How a `cellIs` rule compares.
    pub operator: Option<CfOperator>,
    /// The fragment a text rule looks for.
    pub text: Option<String>,
    /// The span a `timePeriod` rule covers, such as `lastWeek`.
    pub time_period: Option<String>,
    /// How many the `top10` rule takes.
    pub rank: Option<u32>,
    /// Whether that count is a percentage.
    pub percent: bool,
    /// Whether `top10` takes them from the bottom.
    pub bottom: bool,
    /// Whether an `aboveAverage` rule looks above rather than below.
    pub above_average: bool,
    /// Whether it counts the average itself as a match.
    pub equal_average: bool,
    /// How many standard deviations from the average it reaches.
    pub std_dev: Option<i32>,
    /// The rule's formulas, in the order they are written.
    pub formulas: Vec<String>,
    /// The graphical part, for a colour scale, a data bar or an icon set.
    pub scale: Option<CfScale>,
    /// The rule's own `<extLst>` as it stands in the file. A data bar keeps
    /// its `x14:id` here, which ties it to the `x14:cfRule` in the sheet's
    /// extensions that holds its negative colours and axis; without it Excel
    /// draws the bar the 2007 way.
    pub extensions: Option<String>,
}

/// A block of cells and the rules that paint it.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ConditionalFormat {
    /// The cells the rules cover.
    pub sqref: Vec<Range>,
    /// The rules themselves.
    pub rules: Vec<CfRule>,
}

/// A part of the package this crate does not model yet, carried through
/// unchanged.
///
/// A workbook holds far more than cells: drawings, comments, the shapes that
/// position them, links to other workbooks, the document properties. Modelling
/// each is its own phase of the port, and until then the honest thing is to
/// hand the bytes back exactly as they came rather than drop them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaquePart {
    /// Path inside the package, with no leading slash.
    pub path: String,
    /// Its content type, resolved when the file was read.
    pub content_type: Option<String>,
    /// The bytes, as they were.
    pub data: Vec<u8>,
}

/// The cached values of a workbook this one links to.
///
/// A formula written `[1]Sheet1!A1` reads another file. Excel does not need
/// that file to be there: it keeps the last values it saw in
/// `xl/externalLinks/externalLinkN.xml`, and recalculates from them until the
/// link is refreshed. This is that cache, and `[N]` is the position of the
/// book in [`Spreadsheet::external`].
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ExternalBook {
    /// Where the linked file was, as the relationship spells it - usually an
    /// absolute `file:///` URL. Nothing here opens it; it says where the
    /// numbers came from.
    pub path: Option<String>,
    /// The sheets of the linked book that have cached cells.
    pub sheets: Vec<ExternalSheet>,
}

impl ExternalBook {
    /// The cached cells of one sheet, by name, case-insensitively as Excel
    /// compares sheet names.
    #[must_use]
    pub fn sheet(&self, name: &str) -> Option<&ExternalSheet> {
        self.sheets
            .iter()
            .find(|s| s.name.eq_ignore_ascii_case(name))
    }
}

/// One sheet's worth of cached values from a linked workbook.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct ExternalSheet {
    /// The sheet's name in the book it belongs to.
    pub name: String,
    /// The cells the cache holds. Only what the linking formulas asked for is
    /// there, so a gap means "not cached", which reads the same as empty.
    pub cells: BTreeMap<(Row, Col), CellValue>,
}

/// A part reached through a relationship: what kind it is and where it sits.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attachment {
    /// The relationship type, as the full URI xlsx writes.
    pub kind: String,
    /// Path of the part inside the package.
    pub target: String,
}

impl Attachment {
    /// The last segment of the relationship type, which is what tells a
    /// drawing from a comment.
    #[must_use]
    pub fn role(&self) -> &str {
        self.kind.rsplit('/').next().unwrap_or(&self.kind)
    }
}

/// Whether a sheet shows up as a tab.
///
/// `VeryHidden` is the state Excel offers only through VBA: the sheet is
/// missing from the unhide dialog as well as from the tab bar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum SheetVisibility {
    /// A tab the user sees.
    #[default]
    Visible,
    /// Hidden, but listed in the unhide dialog.
    Hidden,
    /// Hidden and not listed.
    VeryHidden,
}

impl SheetVisibility {
    /// What `<sheet state>` calls it, or `None` for the default.
    #[must_use]
    pub fn as_str(self) -> Option<&'static str> {
        match self {
            Self::Visible => None,
            Self::Hidden => Some("hidden"),
            Self::VeryHidden => Some("veryHidden"),
        }
    }

    /// The state named by that attribute; anything else is visible.
    #[must_use]
    pub fn parse(value: &str) -> Self {
        match value {
            "hidden" => Self::Hidden,
            "veryHidden" => Self::VeryHidden,
            _ => Self::Visible,
        }
    }
}

/// A sheet of a workbook.
#[derive(Debug, Clone, Default)]
#[non_exhaustive]
pub struct Worksheet {
    title: String,
    /// Whether the sheet has a tab, and whether that tab can be unhidden.
    pub visibility: SheetVisibility,
    /// The cells, by row, and inside a row sorted by column.
    ///
    /// A map of rows holding dense vectors rather than one map keyed by the
    /// cell: a tree node carries its keys, its values and the slack left for
    /// inserts, and over nine million cells that was most of what the sheet
    /// cost in memory. A row is read and written in column order, so appending
    /// is the common insert.
    ///
    /// Shared, and copied on the first change: the xlsx writer clones the
    /// workbook to rewrite a changed chart's parts, and on a sheet of a
    /// million cells a copy of them was hundreds of megabytes it never
    /// touched.
    cells: Arc<BTreeMap<Row, Vec<(Col, Cell)>>>,
    /// How many cells `cells` holds, kept rather than counted.
    count: usize,
    /// The columns cells have been written in: the leftmost and the rightmost
    /// ever seen, kept rather than searched.
    ///
    /// Exact until `span_stale` is set, an upper bound after: a removed cell
    /// does not narrow it, because narrowing means the walk this field exists
    /// to avoid. See [`Worksheet::dimension_hint`].
    col_span: Option<(Col, Col)>,
    /// Set when a cell in an edge column of `col_span` is removed: the span
    /// may be wider than the cells from then on, and [`Worksheet::dimension`]
    /// walks the rows instead of trusting it.
    span_stale: bool,
    /// Merged areas.
    pub merges: Vec<Range>,
    /// Areas of the array formulas entered with Ctrl+Shift+Enter (and the
    /// dynamic arrays Excel 365 stores the same way). The formula sits in the
    /// top left cell; the rest of the area holds its values.
    ///
    /// A formula outside these areas takes one value where it is given a
    /// range, the cell in its own row or column: Excel calls this implicit
    /// intersection.
    pub array_formulas: Vec<Range>,
    /// Column runs, in the order the file lists them.
    pub columns: Vec<ColumnRun>,
    /// Per-row properties, for rows that have any.
    pub rows: BTreeMap<Row, RowProperties>,
    /// Default column width in characters, from `<sheetFormatPr>`.
    pub default_column_width: Option<f64>,
    /// Default row height in points, from `<sheetFormatPr>`.
    pub default_row_height: Option<f64>,
    /// Saved view state: zoom, scroll position, selection, panes.
    pub view: SheetView,
    /// Restrictions on what may be typed into cells, drop-downs included.
    pub data_validations: Vec<DataValidation>,
    /// Properties of the sheet itself: tab colour, grouping direction.
    pub properties: SheetProperties,
    /// Printing margins.
    pub margins: PageMargins,
    /// How the sheet is laid out on paper.
    pub page_setup: PageSetup,
    /// What else goes on the printed page.
    pub print_options: PrintOptions,
    /// Headers and footers.
    pub header_footer: HeaderFooter,
    /// Page breaks the user put between rows.
    pub row_breaks: Vec<PageBreak>,
    /// Page breaks the user put between columns.
    pub col_breaks: Vec<PageBreak>,
    /// Hyperlinks over cells.
    pub hyperlinks: Vec<Hyperlink>,
    /// Rules that paint cells by what they hold.
    pub conditional_formats: Vec<ConditionalFormat>,
    /// What editing the sheet refuses.
    pub protection: SheetProtection,
    /// Ranges exempted from that protection, or holding a password of their
    /// own.
    pub protected_ranges: Vec<ProtectedRange>,
    /// The Scenario Manager's scenarios. See [`scenario`].
    pub scenarios: scenario::Scenarios,
    /// The filter over a range of the sheet, when it has one.
    pub auto_filter: Option<AutoFilter>,
    /// The last sort of a range made outside a filter (Data > Sort), when
    /// Excel recorded one.
    pub sort_state: Option<SortState>,
    /// Notes attached to cells, in address order.
    pub comments: BTreeMap<CellRef, Comment>,
    /// The pivot tables the sheet holds, read but not written: see
    /// [`crate::model::pivot`].
    pub pivot_tables: Vec<pivot::PivotTable>,
    /// The tables on the sheet, read and written both.
    pub tables: Vec<table::Table>,
    /// The charts drawn on the sheet, read and written both: see
    /// [`crate::model::chart`].
    pub charts: Vec<chart::Chart>,
    /// The 2016 charts on the sheet - waterfall, funnel, treemap and the
    /// rest - read but not written.
    pub extended_charts: Vec<chart::ChartEx>,
    /// The pictures on the sheet, read and written both: see
    /// [`crate::model::image`].
    pub images: Vec<image::Image>,
    /// The shapes on the sheet - boxes, arrows, callouts, text boxes - read
    /// and written both: see [`crate::model::shape`].
    pub shapes: Vec<shape::Shape>,
    /// The sparkline groups, read out of [`Self::extensions`]: see
    /// [`crate::model::sparkline`].
    pub sparklines: Vec<sparkline::SparklineGroup>,
    /// The sheet's `<extLst>`, carried as it was written.
    ///
    /// Everything newer than the 2006 schema hangs off this element:
    /// sparklines, the conditional formats and data validations that needed
    /// more than the original format could say, slicer references. None of it
    /// is modelled, and rewriting a sheet without it would drop the lot - so
    /// the element travels whole, the way an unmodelled part does.
    pub extensions: Option<String>,
    /// Parts attached to this sheet: its drawing, the shapes behind its
    /// comments, the comments themselves.
    pub attachments: Vec<Attachment>,
}

impl Worksheet {
    /// An empty sheet with a validated name.
    ///
    /// # Errors
    /// [`Error::InvalidSheetName`] if the name is empty, longer than 31
    /// characters, or contains any of `* : / \ ? [ ]`.
    pub fn new(title: impl Into<String>) -> Result<Self> {
        Ok(Self {
            title: validate_sheet_title(title.into())?,
            ..Self::default()
        })
    }

    /// The sheet name.
    #[must_use]
    pub fn title(&self) -> &str {
        &self.title
    }

    /// Renames the sheet.
    ///
    /// # Errors
    /// [`Error::InvalidSheetName`] under the same conditions as
    /// [`Worksheet::new`].
    pub fn set_title(&mut self, title: impl Into<String>) -> Result<()> {
        self.title = validate_sheet_title(title.into())?;
        Ok(())
    }

    /// The cell at `at`. A missing cell is indistinguishable from an empty one.
    #[must_use]
    pub fn get(&self, at: CellRef) -> Option<&Cell> {
        let line = self.cells.get(&at.row)?;
        line.binary_search_by_key(&at.col, |(col, _)| *col)
            .ok()
            .map(|i| &line[i].1)
    }

    /// Writes a value, keeping the style of an existing cell.
    ///
    /// For a literal address, [`at!`](crate::at) checks it while compiling:
    /// `sheet.set(at!("A1"), 123)`.
    pub fn set(&mut self, at: CellRef, value: impl Into<CellValue>) {
        self.entry(at).value = value.into();
    }

    /// The same by row and column counted from one, as Excel shows them:
    /// `set_at(1, 1, 123)` writes `A1`.
    ///
    /// # Errors
    /// [`Error::InvalidCellRef`] for a zero or a number past the sheet.
    pub fn set_at(&mut self, row: u32, column: u32, value: impl Into<CellValue>) -> Result<()> {
        self.set(CellRef::from_row_col(row, column)?, value);
        Ok(())
    }

    /// The cell at a row and column counted from one; `None` for a cell that
    /// holds nothing, and for one outside the sheet.
    #[must_use]
    pub fn get_at(&self, row: u32, column: u32) -> Option<&Cell> {
        self.get(CellRef::from_row_col(row, column).ok()?)
    }

    /// Writes a value and the style it is shown in.
    ///
    /// The pair written separately costs a second lookup of the same cell,
    /// which on a sheet being built column by column is the whole of the
    /// difference.
    ///
    /// ```
    /// use excelerate::model::{Spreadsheet, Worksheet};
    /// use excelerate::style::{Style, StyleId};
    /// use excelerate::CellRef;
    ///
    /// let mut book = Spreadsheet::empty();
    /// let mut sheet = Worksheet::new("Sheet1")?;
    /// let mut bold = Style::default();
    /// bold.font.bold = true;
    /// let bold = book.styles.intern(bold);
    ///
    /// sheet.set_styled(CellRef::parse("A1")?, "Итого", bold);
    /// # assert_ne!(bold, StyleId::default());
    /// # Ok::<(), excelerate::Error>(())
    /// ```
    pub fn set_styled(&mut self, at: CellRef, value: impl Into<CellValue>, style: StyleId) {
        let cell = self.entry(at);
        cell.value = value.into();
        cell.style = style;
    }

    /// Writes a run of values along a row, starting at `from`.
    ///
    /// Building a sheet from records means writing a row at a time, and doing
    /// that a cell at a time means building an address per cell. The run stops
    /// at the last column of the sheet rather than wrapping round to the first.
    ///
    /// ```
    /// use excelerate::model::{CellValue, Worksheet};
    /// use excelerate::{CellRef, Row};
    ///
    /// let mut sheet = Worksheet::new("Sheet1")?;
    /// let a2 = CellRef::parse("A2")?;
    /// sheet.set_row(a2, ["Чай".into(), CellValue::Number(42.0)]);
    ///
    /// assert_eq!(sheet.get(a2).map(|c| c.value.as_str()), Some(Some("Чай")));
    /// assert_eq!(
    ///     sheet.get(CellRef::parse("B2")?).and_then(|c| c.value.as_number()),
    ///     Some(42.0),
    /// );
    /// # let _ = Row::new(0);
    /// # Ok::<(), excelerate::Error>(())
    /// ```
    pub fn set_row<V: Into<CellValue>>(
        &mut self,
        from: CellRef,
        values: impl IntoIterator<Item = V>,
    ) {
        for (offset, value) in values.into_iter().enumerate() {
            let Ok(offset) = u32::try_from(offset) else {
                return;
            };
            let Some(col) = from.col.index().checked_add(offset).and_then(Col::new) else {
                return;
            };
            self.set(CellRef::new(col, from.row), value);
        }
    }

    /// The cell at `at` for modification, created empty if absent.
    pub fn entry(&mut self, at: CellRef) -> &mut Cell {
        // A new row reserves room for as many cells as the row before it
        // holds: rows of one sheet tend to be alike, and a row grown by pushes
        // to nine cells holds room for sixteen, which over a million rows was
        // most of the memory reading a sheet peaked at.
        let cells = Arc::make_mut(&mut self.cells);
        let width = cells.last_key_value().map_or(0, |(_, line)| line.len());
        let line = cells
            .entry(at.row)
            .or_insert_with(|| Vec::with_capacity(width));
        // Cells arrive in column order when a sheet is read, so the end of
        // the row is checked before a search.
        let index = match line.last() {
            Some((col, _)) if *col < at.col => Err(line.len()),
            _ => line.binary_search_by_key(&at.col, |(col, _)| *col),
        };
        let index = match index {
            Ok(i) => i,
            Err(i) => {
                line.insert(i, (at.col, Cell::default()));
                self.count += 1;
                self.col_span = Some(match self.col_span {
                    Some((lo, hi)) => (lo.min(at.col), hi.max(at.col)),
                    None => (at.col, at.col),
                });
                i
            }
        };
        &mut line[index].1
    }

    /// Removes the cell at `at`, returning it.
    pub fn remove(&mut self, at: CellRef) -> Option<Cell> {
        let line = self.cells.get(&at.row)?;
        let index = line.binary_search_by_key(&at.col, |(col, _)| *col).ok()?;
        let cells = Arc::make_mut(&mut self.cells);
        let line = cells.get_mut(&at.row)?;
        let (_, cell) = line.remove(index);
        if line.is_empty() {
            cells.remove(&at.row);
        }
        self.count -= 1;
        if self.count == 0 {
            self.col_span = None;
            self.span_stale = false;
        } else if self
            .col_span
            .is_some_and(|(lo, hi)| at.col == lo || at.col == hi)
        {
            self.span_stale = true;
        }
        Some(cell)
    }

    /// Drops every cell, keeping the rest of the sheet.
    ///
    /// What a streaming read does once it has handed a row over: the row is
    /// the caller's now, and holding it is what makes a big sheet expensive.
    pub fn clear_cells(&mut self) {
        if self.count > 0 {
            Arc::make_mut(&mut self.cells).clear();
        }
        self.count = 0;
        self.col_span = None;
        self.span_stale = false;
    }

    /// Number of cells stored.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.count
    }

    /// Whether the sheet stores no cells.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.count == 0
    }

    /// Every stored cell in row-major order.
    pub fn iter(&self) -> impl Iterator<Item = (CellRef, &Cell)> {
        self.cells.iter().flat_map(|(&row, line)| {
            line.iter()
                .map(move |(col, cell)| (CellRef::new(*col, row), cell))
        })
    }

    /// The cells of one row, in column order.
    pub fn row_cells(&self, row: Row) -> impl Iterator<Item = (Col, &Cell)> {
        self.cells
            .get(&row)
            .into_iter()
            .flat_map(|line| line.iter().map(|(col, cell)| (*col, cell)))
    }

    /// Gives back the room rows took while they grew. A reader calls this
    /// once a sheet is complete: a row grown to nine cells by pushes holds
    /// room for sixteen.
    pub fn shrink_to_fit(&mut self) {
        for line in Arc::make_mut(&mut self.cells).values_mut() {
            line.shrink_to_fit();
        }
    }

    /// The run of columns covering `col`, if the sheet describes one.
    ///
    /// Runs are searched from the last one back, so a later run overrides an
    /// earlier one covering the same column - which is how Excel reads them.
    #[must_use]
    pub fn column_run(&self, col: Col) -> Option<&ColumnRun> {
        self.columns
            .iter()
            .rev()
            .find(|run| run.first <= col && col <= run.last)
    }

    /// The width of a column, or the sheet default when it has none.
    #[must_use]
    pub fn column_width(&self, col: Col) -> Option<f64> {
        self.column_run(col)
            .and_then(|run| run.width)
            .or(self.default_column_width)
    }

    /// The run describing `col` alone, ready to be written to.
    ///
    /// A run covering more columns is split around this one, so that a change
    /// here does not quietly resize its neighbours; a sheet that says nothing
    /// about the column gets a run for it.
    pub fn column_entry(&mut self, col: Col) -> &mut ColumnRun {
        let covering = self
            .columns
            .iter()
            .rposition(|run| run.first <= col && col <= run.last);
        let at = match covering {
            None => {
                self.columns.push(ColumnRun::new(col, col));
                self.columns.len() - 1
            }
            Some(at) if self.columns[at].first == col && self.columns[at].last == col => at,
            Some(at) => {
                let run = self.columns[at].clone();
                let mut mine = run.clone();
                mine.first = col;
                mine.last = col;
                self.columns[at] = mine;
                let mut index = at;
                if run.first < col {
                    let mut left = run.clone();
                    left.last = Col::new(col.index() - 1).unwrap_or(run.first);
                    self.columns.insert(at, left);
                    index += 1;
                }
                if col < run.last {
                    let mut right = run;
                    right.first = Col::new(col.index() + 1).unwrap_or(right.last);
                    self.columns.insert(index + 1, right);
                }
                index
            }
        };
        &mut self.columns[at]
    }

    /// Sets the width of a column in characters of the default font, or clears
    /// it with `None` so the sheet default applies again.
    pub fn set_column_width(&mut self, col: Col, width: Option<f64>) {
        let run = self.column_entry(col);
        run.width = width;
        run.custom_width = width.is_some();
    }

    /// Hides a column, or shows it again.
    pub fn set_column_hidden(&mut self, col: Col, hidden: bool) {
        self.column_entry(col).hidden = hidden;
    }

    /// Sets the height of a row in points, or clears it with `None`.
    pub fn set_row_height(&mut self, row: Row, height: Option<f64>) {
        let props = self.rows.entry(row).or_default();
        props.height = height;
        props.custom_height = height.is_some();
    }

    /// Hides a row, or shows it again.
    pub fn set_row_hidden(&mut self, row: Row, hidden: bool) {
        self.rows.entry(row).or_default().hidden = hidden;
    }

    /// How deeply a column is grouped: 0 when it is not.
    #[must_use]
    pub fn column_outline_level(&self, col: Col) -> u8 {
        self.column_run(col).map_or(0, |run| run.outline_level)
    }

    /// How deeply a row is grouped: 0 when it is not.
    #[must_use]
    pub fn row_outline_level(&self, row: Row) -> u8 {
        self.rows.get(&row).map_or(0, |r| r.outline_level)
    }

    /// The height of a row, or the sheet default when it has none.
    #[must_use]
    pub fn row_height(&self, row: Row) -> Option<f64> {
        self.rows
            .get(&row)
            .and_then(|r| r.height)
            .or(self.default_row_height)
    }

    /// A rectangle covering every non-empty cell, or `None` for an empty
    /// sheet, without walking the rows.
    ///
    /// Rows are exact; columns are an upper bound, because a removed cell
    /// never narrows the span. Until a cell in the leftmost or the rightmost
    /// column is removed the two are the same, and [`Worksheet::dimension`]
    /// answers from here too; after that it walks every row, which on a sheet
    /// of seven hundred thousand rows is some twenty milliseconds.
    ///
    /// ```
    /// use excelerate::model::Worksheet;
    /// use excelerate::CellRef;
    ///
    /// let mut sheet = Worksheet::new("Sheet1")?;
    /// sheet.set(CellRef::parse("B2")?, 1.0);
    /// sheet.set(CellRef::parse("D4")?, 2.0);
    /// assert_eq!(sheet.dimension_hint(), sheet.dimension());
    ///
    /// sheet.remove(CellRef::parse("D4")?);
    /// // The hint keeps column D; the exact answer gives it up.
    /// let hint = sheet.dimension_hint().expect("a cell is left");
    /// let exact = sheet.dimension().expect("a cell is left");
    /// assert_eq!(hint.end.col, CellRef::parse("D1")?.col);
    /// assert_eq!(exact.end.col, CellRef::parse("B1")?.col);
    /// # Ok::<(), excelerate::Error>(())
    /// ```
    #[must_use]
    pub fn dimension_hint(&self) -> Option<Range> {
        let (&first_row, _) = self.cells.first_key_value()?;
        let (&last_row, _) = self.cells.last_key_value()?;
        let (min_col, max_col) = self.col_span?;
        Some(Range::new(
            CellRef::new(min_col, first_row),
            CellRef::new(max_col, last_row),
        ))
    }

    /// The smallest rectangle covering every non-empty cell, or `None` for an
    /// empty sheet.
    ///
    /// Answered without a walk unless a cell in an edge column has been
    /// removed since the sheet was read: only then can the span kept for
    /// [`Worksheet::dimension_hint`] be too wide.
    #[must_use]
    pub fn dimension(&self) -> Option<Range> {
        if !self.span_stale {
            return self.dimension_hint();
        }
        let (&first_row, _) = self.cells.first_key_value()?;
        let (&last_row, _) = self.cells.last_key_value()?;
        // Rows are sorted by column, so each gives its extremes at its ends:
        // a walk of the rows, not of the cells.
        let mut columns = self
            .cells
            .values()
            .filter_map(|line| Some((line.first()?.0, line.last()?.0)));
        let (mut min_col, mut max_col) = columns.next()?;
        for (first, last) in columns {
            min_col = min_col.min(first);
            max_col = max_col.max(last);
        }
        Some(Range::new(
            CellRef::new(min_col, first_row),
            CellRef::new(max_col, last_row),
        ))
    }
}

/// A workbook: its sheets and the tables they share.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct Spreadsheet {
    sheets: Vec<Worksheet>,
    active: usize,
    /// Style table, shared by every sheet.
    pub styles: StyleTable,
    /// The workbook's base date.
    pub epoch: Epoch,
    /// Names standing for formulas, and Excel's own reserved ones.
    pub defined_names: Vec<DefinedName>,
    /// The rest of `<workbookPr>`, as the file wrote it.
    ///
    /// Its one attribute this crate understands is `date1904`, which becomes
    /// [`Spreadsheet::epoch`]; the others are settings nothing here reads, and
    /// losing them would change how Excel opens the file.
    pub workbook_properties: Vec<(String, String)>,
    /// The attributes of `<calcPr>`: calculation mode, iteration limits, and
    /// the id of the engine that last computed the workbook.
    pub calculation_properties: Vec<(String, String)>,
    /// The rest of the first `<workbookView>`, as the file wrote it: the first
    /// tab shown in the tab bar (`firstSheet`), the tab bar's share of the
    /// window (`tabRatio`), whether tabs and scroll bars show.
    ///
    /// `activeTab` is [`Spreadsheet::set_active`], and the window's position
    /// and size are left out: they are the geometry of someone else's screen.
    pub workbook_view: Vec<(String, String)>,
    /// What changing the workbook's shape refuses.
    pub protection: WorkbookProtection,
    /// Parts attached to the workbook, such as links to other workbooks.
    pub attachments: Vec<Attachment>,
    /// Parts attached to the package itself: the document properties and a
    /// thumbnail, when the file has one.
    pub doc_props: Vec<Attachment>,
    /// Title, author, dates and the user's own fields.
    ///
    /// Written back by comparison: a part whose content the model still
    /// states travels as the bytes it came in.
    pub properties: properties::DocumentProperties,
    /// Every part carried through unmodelled, in no particular order.
    pub parts: Vec<OpaquePart>,
    /// The cached values of the workbooks this one links to, in the order
    /// `<externalReferences>` lists them: `[1]` is the first.
    pub external: Vec<ExternalBook>,
    /// The theme part, kept as the file wrote it.
    ///
    /// A theme colour is an index into this part, so replacing it with a
    /// canned one repaints every cell that uses one. Nothing here reads the
    /// theme yet, so it travels through unparsed rather than being modelled
    /// for no one.
    pub theme: Option<String>,
    /// The workbook's own `<extLst>`, carried as it was written: slicer
    /// caches and everything else the later schemas hang there.
    /// The pivot caches the workbook holds, read but not written.
    pub pivot_caches: Vec<pivot::PivotCache>,
    /// The workbook's own `<extLst>`, carried as it was written: slicer
    /// caches and everything else the later schemas hang there.
    pub workbook_extensions: Option<String>,
    /// The stylesheet's `<extLst>`, which is where slicer and table styles
    /// live. Same reasoning as [`Worksheet::extensions`].
    pub style_extensions: Option<String>,
    /// The stylesheet's `<tableStyles>`, carried as written: a slicer and a
    /// table name the style they are painted with, and a style this crate
    /// dropped took the slicer with it - Excel removed all fourteen of them
    /// from a dashboard rewritten without this.
    pub table_styles: Option<String>,
    /// The stylesheet's `<colors>`: the palette an `indexed` colour counts
    /// into, and the colours last picked in the dialog.
    pub palette: Option<String>,
    /// Whether the book is a template (`.xltx`, `.xltm`). Only the content
    /// type of the main part says so, and Excel will not open a template
    /// written as a plain workbook; the xlsx reader sets it from that type.
    pub template: bool,
}

impl Default for Spreadsheet {
    /// A workbook holding one empty sheet, as a new document starts.
    fn default() -> Self {
        Self {
            sheets: vec![Worksheet::new("Worksheet").unwrap_or_default()],
            active: 0,
            styles: StyleTable::default(),
            epoch: Epoch::default(),
            defined_names: Vec::new(),
            workbook_properties: Vec::new(),
            calculation_properties: Vec::new(),
            workbook_view: Vec::new(),
            protection: WorkbookProtection::default(),
            attachments: Vec::new(),
            doc_props: Vec::new(),
            properties: properties::DocumentProperties::default(),
            parts: Vec::new(),
            external: Vec::new(),
            theme: None,
            pivot_caches: Vec::new(),
            workbook_extensions: None,
            style_extensions: None,
            table_styles: None,
            palette: None,
            template: false,
        }
    }
}

impl Spreadsheet {
    /// A new workbook holding one empty sheet.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// A workbook with no sheets, for readers that will fill it themselves.
    #[must_use]
    pub fn empty() -> Self {
        Self {
            sheets: Vec::new(),
            active: 0,
            ..Self::default()
        }
    }

    /// Every sheet, in tab order.
    #[must_use]
    pub fn sheets(&self) -> &[Worksheet] {
        &self.sheets
    }

    /// The sheet at a tab index.
    #[must_use]
    pub fn sheet(&self, index: usize) -> Option<&Worksheet> {
        self.sheets.get(index)
    }

    /// The sheet at a tab index, for modification.
    pub fn sheet_mut(&mut self, index: usize) -> Option<&mut Worksheet> {
        self.sheets.get_mut(index)
    }

    /// A sheet by name. Excel compares sheet names case-insensitively.
    #[must_use]
    pub fn sheet_by_name(&self, name: &str) -> Option<&Worksheet> {
        self.sheets.get(self.sheet_index_by_name(name)?)
    }

    /// The tab index of a sheet by name, compared as Excel compares them -
    /// ignoring case, and not only for ASCII.
    #[must_use]
    pub fn sheet_index_by_name(&self, name: &str) -> Option<usize> {
        let name = name.to_lowercase();
        self.sheets
            .iter()
            .position(|s| s.title.to_lowercase() == name)
    }

    /// The text a cell shows: its value through its own number format.
    ///
    /// The value alone is not what the user sees. `45292` in a cell formatted
    /// `DD.MM.YYYY` reads `01.01.2024`, and the format lives in the workbook's
    /// style table while the epoch lives on the workbook, which is why this is
    /// here rather than on [`Worksheet`]. An empty cell, or one on a sheet
    /// that does not exist, gives an empty string.
    ///
    /// ```
    /// use excelerate::model::{CellValue, Spreadsheet, Worksheet};
    /// use excelerate::style::{NumberFormat, Style};
    /// use excelerate::CellRef;
    ///
    /// let mut book = Spreadsheet::empty();
    /// let mut sheet = Worksheet::new("Sheet1")?;
    /// let at = CellRef::parse("A1")?;
    /// sheet.set(at, 45292.0);
    /// let style = book.styles.intern(Style {
    ///     number_format: NumberFormat::Custom("DD.MM.YYYY".to_owned()),
    ///     ..Style::default()
    /// });
    /// sheet.entry(at).style = style;
    /// book.add_sheet(sheet)?;
    ///
    /// assert_eq!(book.formatted(0, at), "01.01.2024");
    /// # Ok::<(), excelerate::Error>(())
    /// ```
    #[must_use]
    pub fn formatted(&self, sheet: usize, at: CellRef) -> String {
        let Some(cell) = self.sheets.get(sheet).and_then(|s| s.get(at)) else {
            return String::new();
        };
        let code = self
            .styles
            .get(cell.style)
            .map_or(crate::style::format::GENERAL, |s| s.number_format.code());
        cell.value.display(code, self.epoch)
    }

    /// The same by 1-based row and column, the way a user counts them.
    ///
    /// A row or column outside the sheet's grid gives an empty string, the
    /// same answer an empty cell gives.
    #[must_use]
    pub fn formatted_at(&self, sheet: usize, row: u32, column: u32) -> String {
        let (Ok(row), Ok(column)) = (
            Row::from_one_based(u64::from(row)),
            Col::from_one_based(u64::from(column)),
        ) else {
            return String::new();
        };
        self.formatted(sheet, CellRef::new(column, row))
    }

    /// Tab index of the active sheet.
    #[must_use]
    pub const fn active_index(&self) -> usize {
        self.active
    }

    /// The active sheet. A workbook with no sheets has none.
    #[must_use]
    pub fn active_sheet(&self) -> Option<&Worksheet> {
        self.sheets.get(self.active)
    }

    /// The active sheet, for modification.
    pub fn active_sheet_mut(&mut self) -> Option<&mut Worksheet> {
        self.sheets.get_mut(self.active)
    }

    /// Takes the sheet at `index` out of the book, keeping the active tab
    /// inside the list.
    ///
    /// Nothing here rewrites the references to it; that is what
    /// [`crate::edit::remove_sheet`] is for, and it is the only caller.
    pub(crate) fn take_sheet(&mut self, index: usize) -> Option<Worksheet> {
        if index >= self.sheets.len() {
            return None;
        }
        let gone = self.sheets.remove(index);
        self.active = self.active.min(self.sheets.len().saturating_sub(1));
        Some(gone)
    }

    /// Moves the sheet at `from` to position `to`, sliding the sheets between
    /// them along.
    ///
    /// Nothing here renumbers what points at a sheet by index; that is what
    /// [`crate::edit::move_sheet`] is for, and it is the only caller.
    pub(crate) fn reorder_sheet(&mut self, from: usize, to: usize) {
        if from >= self.sheets.len() || to >= self.sheets.len() || from == to {
            return;
        }
        let sheet = self.sheets.remove(from);
        self.sheets.insert(to, sheet);
    }

    /// Makes the sheet at `index` active.
    ///
    /// # Errors
    /// [`Error::InvalidSheetName`] carrying the index if no such sheet exists.
    pub fn set_active(&mut self, index: usize) -> Result<()> {
        if index >= self.sheets.len() {
            return Err(Error::InvalidSheetName(index.to_string()));
        }
        self.active = index;
        Ok(())
    }

    /// Appends a sheet and returns its tab index.
    ///
    /// # Errors
    /// [`Error::DuplicateSheetName`] if a sheet with that name already exists -
    /// Excel compares the names case-insensitively.
    pub fn add_sheet(&mut self, sheet: Worksheet) -> Result<usize> {
        if self.sheet_by_name(sheet.title()).is_some() {
            return Err(Error::DuplicateSheetName(sheet.title));
        }
        self.sheets.push(sheet);
        Ok(self.sheets.len() - 1)
    }
}

/// Validates a sheet name against Excel's rules.
fn validate_sheet_title(title: String) -> Result<String> {
    let len = title.chars().count();
    if len == 0
        || len > SHEET_TITLE_MAX_LENGTH
        || title.contains(SHEET_TITLE_INVALID_CHARS)
        // Excel quotes sheet names in formulas, so a leading or trailing
        // apostrophe would make the reference ambiguous.
        || title.starts_with('\'')
        || title.ends_with('\'')
    {
        return Err(Error::InvalidSheetName(title));
    }
    Ok(title)
}

#[cfg(test)]
mod tests {
    /// Cells land in their row in column order whatever order they are
    /// written in, and a row with nothing left in it is gone.
    #[test]
    fn cells_keep_row_and_column_order_however_they_arrive() {
        let mut sheet = Worksheet::new("S").unwrap();
        let at = |a: &str| CellRef::parse(a).unwrap();
        for a in ["C2", "A2", "B5", "B2", "A1"] {
            sheet.set(at(a), a);
        }
        sheet.set(at("B2"), "again");
        let order: Vec<String> = sheet.iter().map(|(a, _)| a.to_string()).collect();
        assert_eq!(order, ["A1", "A2", "B2", "C2", "B5"]);
        assert_eq!(sheet.len(), 5);
        assert_eq!(sheet.dimension(), Some(Range::parse("A1:C5").unwrap()));
        assert!(sheet.remove(at("B5")).is_some());
        assert!(sheet.remove(at("B5")).is_none());
        assert_eq!(sheet.len(), 4);
        assert_eq!(sheet.dimension(), Some(Range::parse("A1:C2").unwrap()));

        assert_eq!(sheet.row_cells(Row::new(1).unwrap()).count(), 3);
    }

    use super::{CellValue, MAX_STRING_LENGTH, Spreadsheet, Worksheet};
    use crate::coordinate::{CellRef, Range, Row};
    use crate::style::StyleId;

    fn r(a: &str) -> CellRef {
        CellRef::parse(a).expect("test reference is valid")
    }

    #[test]
    fn sheet_titles_follow_excel_rules() {
        assert!(Worksheet::new("Data").is_ok());
        assert!(Worksheet::new("Данные").is_ok(), "non-ASCII names are fine");
        assert!(Worksheet::new("").is_err(), "empty name");
        assert!(Worksheet::new("a".repeat(31)).is_ok());
        assert!(
            Worksheet::new("a".repeat(32)).is_err(),
            "the limit is 31 characters"
        );
        for bad in ["a*b", "a:b", "a/b", "a\\b", "a?b", "a[b", "a]b"] {
            assert!(Worksheet::new(bad).is_err(), "banned character in {bad:?}");
        }
        assert!(
            Worksheet::new("'quoted'").is_err(),
            "apostrophe at the edges"
        );
    }

    /// `dimension` trusts the kept column span until an edge column loses a
    /// cell; this pins the cases where trusting it would be wrong or needless.
    #[test]
    fn the_used_range_stays_exact_through_removals() {
        let at = |a: &str| CellRef::parse(a).expect("a valid reference");
        let range = |a: &str| Range::parse(a).expect("a valid range");
        let mut sheet = Worksheet::new("S").expect("a valid title");
        for a in ["B2", "C5", "E3"] {
            sheet.set(at(a), 1.0);
        }
        sheet.remove(at("C5"));
        assert_eq!(sheet.dimension(), Some(range("B2:E3")), "an inner column");
        sheet.remove(at("E3"));
        assert_eq!(sheet.dimension(), Some(range("B2:B2")), "an edge column");
        sheet.remove(at("B2"));
        assert_eq!(sheet.dimension(), None, "emptied");
        sheet.set(at("D7"), 1.0);
        assert_eq!(sheet.dimension(), Some(range("D7:D7")), "filled again");
        assert_eq!(sheet.dimension_hint(), sheet.dimension());
    }

    #[test]
    fn cells_store_and_iterate_in_write_order() {
        let mut ws = Worksheet::new("S").unwrap();
        // Written out of order - iteration must still be row by row.
        ws.set(r("C1"), 3.0);
        ws.set(r("A2"), "text");
        ws.set(r("A1"), true);

        let order: Vec<String> = ws.iter().map(|(at, _)| at.to_string()).collect();
        assert_eq!(order, ["A1", "C1", "A2"]);
        assert_eq!(ws.len(), 3);
        assert_eq!(ws.get(r("A1")).unwrap().value, CellValue::Bool(true));
        assert!(ws.get(r("B1")).is_none());
    }

    #[test]
    fn set_preserves_style() {
        let mut ws = Worksheet::new("S").unwrap();
        let styled = StyleId::from_index(7);
        ws.entry(r("A1")).style = styled;
        ws.set(r("A1"), 42.0);
        assert_eq!(
            ws.get(r("A1")).unwrap().style,
            styled,
            "rewriting a value keeps the style"
        );
    }

    #[test]
    fn a_row_yields_only_its_own_cells() {
        let mut ws = Worksheet::new("S").unwrap();
        for a in ["A1", "C1", "B2", "Z2", "A3"] {
            ws.set(r(a), 1.0);
        }
        let row2 = crate::coordinate::Row::from_one_based(2).unwrap();
        let cols: Vec<u32> = ws.row_cells(row2).map(|(col, _)| col.one_based()).collect();
        assert_eq!(cols, vec![2, 26]);
    }

    #[test]
    fn dimension_covers_used_cells() {
        let mut ws = Worksheet::new("S").unwrap();
        assert!(ws.dimension().is_none(), "an empty sheet has no dimension");
        ws.set(r("C5"), 1.0);
        ws.set(r("B2"), 1.0);
        ws.set(r("D3"), 1.0);
        assert_eq!(ws.dimension().unwrap(), Range::parse("B2:D5").unwrap());
    }

    #[test]
    fn text_is_kept_as_written_and_truncated() {
        assert_eq!(
            CellValue::text("a\r\nb\rc"),
            CellValue::Text("a\r\nb\rc".into()),
            "a line break the author typed is theirs, CR and all"
        );
        let long = "я".repeat(MAX_STRING_LENGTH + 10);
        let CellValue::Text(t) = CellValue::text(long) else {
            panic!("expected text");
        };
        assert_eq!(
            t.chars().count(),
            MAX_STRING_LENGTH,
            "truncated by characters, not bytes"
        );
    }

    #[test]
    fn workbook_manages_sheets() {
        let mut wb = Spreadsheet::new();
        assert_eq!(wb.sheets().len(), 1);
        assert_eq!(wb.active_sheet().unwrap().title(), "Worksheet");

        let idx = wb.add_sheet(Worksheet::new("Второй").unwrap()).unwrap();
        assert_eq!(idx, 1);
        assert!(
            wb.sheet_by_name("ВТОРОЙ").is_some(),
            "sheet lookup is case-insensitive beyond ASCII"
        );
        assert!(
            wb.add_sheet(Worksheet::new("второй").unwrap()).is_err(),
            "duplicate name"
        );

        wb.set_active(1).unwrap();
        assert_eq!(wb.active_sheet().unwrap().title(), "Второй");
        assert!(wb.set_active(9).is_err());
    }

    /// Reading data means asking what a cell holds, and for a formula cell
    /// that is the result rather than the text.
    #[test]
    fn a_value_is_read_through_its_formula() {
        let computed = CellValue::Formula {
            formula: "1+2".to_owned(),
            cached: Some(Box::new(CellValue::Number(3.0))),
        };
        assert_eq!(computed.as_number(), Some(3.0));
        assert_eq!(computed.result(), &CellValue::Number(3.0));

        // Written by hand and not computed yet: empty, not an error.
        let fresh = CellValue::formula("1+2");
        assert_eq!(fresh.result(), &CellValue::Empty);
        assert_eq!(fresh.as_number(), None);

        assert_eq!(CellValue::text("Ёж").as_str(), Some("Ёж"));
        assert_eq!(CellValue::Bool(true).as_bool(), Some(true));
        let div0 = crate::CellError::Div0;
        assert_eq!(CellValue::Error(div0).as_error(), Some(div0));
        // A boolean is not a number, the way Excel keeps them apart.
        assert_eq!(CellValue::Bool(true).as_number(), None);
    }

    /// The number in the cell and the text the user sees are not the same
    /// thing: the format is what turns 45292 into a date.
    #[test]
    fn a_cell_shows_its_value_through_its_own_format() {
        use crate::style::{NumberFormat, Style};

        let mut book = Spreadsheet::empty();
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        let at = CellRef::parse("A1").unwrap();
        sheet.set(at, 45292.0);
        let dated = book.styles.intern(Style {
            number_format: NumberFormat::Custom("DD.MM.YYYY".to_owned()),
            ..Style::default()
        });
        sheet.entry(at).style = dated;
        let plain = CellRef::parse("A2").unwrap();
        sheet.set(plain, 45292.0);
        book.add_sheet(sheet).unwrap();

        assert_eq!(book.formatted(0, at), "01.01.2024");
        assert_eq!(book.formatted(0, plain), "45292");
        // A cell that is not there, and a sheet that is not there.
        assert_eq!(book.formatted(0, CellRef::parse("Z99").unwrap()), "");
        assert_eq!(book.formatted(7, at), "");
    }

    /// A run of values written along a row, and what happens at the edge of
    /// the sheet: the run stops rather than wrapping round to column A.
    #[test]
    fn a_row_of_values_stops_at_the_last_column() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        let last = CellRef::new(
            crate::Col::new(crate::coordinate::MAX_COL - 1).unwrap(),
            crate::Row::new(0).unwrap(),
        );
        sheet.set_row(last, [1.0, 2.0, 3.0]);
        assert_eq!(sheet.get(last).and_then(|c| c.value.as_number()), Some(1.0));
        assert_eq!(sheet.len(), 1, "the other two had nowhere to go");

        let a1 = CellRef::parse("A1").unwrap();
        sheet.set_row(a1, ["a", "b", "c"]);
        assert_eq!(
            sheet
                .get(CellRef::parse("C1").unwrap())
                .unwrap()
                .value
                .as_str(),
            Some("c")
        );
    }

    /// Writing the value and the style together leaves the same cell as
    /// writing them one after the other.
    #[test]
    fn a_value_and_its_style_are_written_together() {
        use crate::style::Style;

        let mut book = Spreadsheet::empty();
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        let mut bold = Style::default();
        bold.font.bold = true;
        let bold = book.styles.intern(bold);

        let at = CellRef::parse("A1").unwrap();
        sheet.set_styled(at, "Итого", bold);
        assert_eq!(sheet.get(at).unwrap().style, bold);
        assert_eq!(sheet.get(at).unwrap().value.as_str(), Some("Итого"));

        // And by numbers, the way a user counts rows and columns.
        book.add_sheet(sheet).unwrap();
        assert_eq!(book.formatted_at(0, 1, 1), "Итого");
        assert_eq!(book.formatted_at(0, 0, 1), "", "row zero is not a row");
    }

    #[test]
    fn the_dimension_hint_never_falls_short_of_the_exact_one() {
        let mut sheet = Worksheet::new("Sheet1").unwrap();
        assert_eq!(sheet.dimension_hint(), None, "an empty sheet has no extent");

        for a in ["C1", "A5", "F3", "B2"] {
            sheet.set(CellRef::parse(a).unwrap(), 1.0);
            assert_eq!(
                sheet.dimension_hint(),
                sheet.dimension(),
                "with nothing removed the hint is exact, at {a}"
            );
        }

        sheet.remove(CellRef::parse("F3").unwrap());
        let (hint, exact) = (sheet.dimension_hint().unwrap(), sheet.dimension().unwrap());
        assert!(
            hint.start.col <= exact.start.col && exact.end.col <= hint.end.col,
            "the hint covers the exact span: {hint:?} against {exact:?}"
        );
        assert_eq!(
            (hint.start.row, hint.end.row),
            (exact.start.row, exact.end.row),
            "rows stay exact"
        );
    }

    #[test]
    fn a_column_setting_splits_the_run_it_lands_in() {
        let mut sheet = Worksheet::new("S").unwrap();
        let mut run = crate::model::ColumnRun::new(
            crate::coordinate::Col::from_letters("A").unwrap(),
            crate::coordinate::Col::from_letters("E").unwrap(),
        );
        run.width = Some(10.0);
        sheet.columns.push(run);

        let c = crate::coordinate::Col::from_letters("C").unwrap();
        sheet.set_column_width(c, Some(30.0));

        assert_eq!(sheet.column_width(c), Some(30.0));
        for other in ["A", "B", "D", "E"] {
            let other = crate::coordinate::Col::from_letters(other).unwrap();
            assert_eq!(
                sheet.column_width(other),
                Some(10.0),
                "{other:?} keeps its width"
            );
        }
        // Three runs, none of them overlapping: A:B, C, D:E.
        assert_eq!(sheet.columns.len(), 3);
        for pair in sheet.columns.windows(2) {
            assert!(pair[0].last < pair[1].first, "runs stay apart");
        }

        // A column the sheet says nothing about gets a run of its own.
        let z = crate::coordinate::Col::from_letters("Z").unwrap();
        sheet.set_row_height(
            crate::coordinate::Row::from_one_based(2).unwrap(),
            Some(24.0),
        );
        sheet.set_column_hidden(z, true);
        assert!(sheet.column_run(z).unwrap().hidden);
        assert_eq!(
            sheet.row_height(crate::coordinate::Row::from_one_based(2).unwrap()),
            Some(24.0)
        );
    }
}
