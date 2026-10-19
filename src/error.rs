//! Crate errors and Excel cell error codes.

use crate::coordinate::{MAX_COL, MAX_ROW};

/// An error from a workbook operation.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A cell reference could not be parsed.
    #[error("invalid cell reference: {0:?}")]
    InvalidCellRef(String),

    /// A range could not be parsed.
    #[error("invalid range: {0:?}")]
    InvalidRange(String),

    /// Row number outside the sheet limits.
    #[error("row {0} is outside the range 1..={MAX_ROW}")]
    RowOutOfRange(u64),

    /// Column number outside the sheet limits.
    #[error("column {0} is outside the range 1..={MAX_COL}")]
    ColOutOfRange(u64),

    /// A date that Excel's serial format cannot represent.
    #[error("date is outside the range Excel can represent")]
    DateOutOfRange,

    /// The workbook already holds a sheet with this name.
    #[error("a sheet named {0:?} already exists")]
    DuplicateSheetName(String),

    /// A sheet was addressed by a tab index the workbook does not have.
    #[error("no sheet at index {0}")]
    SheetIndexOutOfRange(usize),

    /// An insertion would push cells that hold something, or a merged area,
    /// past the last row or column of the sheet. Excel refuses the same
    /// edit; nothing has been changed.
    #[error("the insertion would push non-empty cells off the sheet")]
    WouldPushOffSheet,

    /// A sheet name that breaks Excel's naming rules.
    #[error("invalid sheet name {0:?}")]
    InvalidSheetName(String),

    /// A formula could not be parsed.
    #[error("invalid formula: {0}")]
    InvalidFormula(String),

    /// The xlsx package is malformed, truncated, or not an xlsx at all.
    #[error("xlsx: {0}")]
    Xlsx(String),

    /// The xlsb package is malformed, truncated, or not an xlsb at all.
    #[error("xlsb: {0}")]
    Xlsb(String),

    /// A CSV file could not be read or written.
    #[error("csv: {0}")]
    Csv(String),

    /// An HTML file could not be read or written.
    #[error("html: {0}")]
    Html(String),

    /// The ODS package is malformed, truncated, or not an ODS at all.
    #[error("ods: {0}")]
    Ods(String),

    /// The workbook is encrypted in a way this crate does not read, or its
    /// encryption descriptor is malformed.
    #[error("encrypted: {0}")]
    Encrypted(String),

    /// The workbook is encrypted, and the password given (or none, and
    /// Excel's read-only default) does not open it.
    #[error("the password does not open this workbook")]
    WrongPassword,

    /// The xls file is malformed, encrypted, or not an xls at all.
    #[error("xls: {0}")]
    Xls(String),

    /// The SYLK file is malformed or not a SYLK file at all.
    #[error("slk: {0}")]
    Slk(String),

    /// The Gnumeric document is malformed or not a Gnumeric document at all.
    #[error("gnumeric: {0}")]
    Gnumeric(String),

    /// A file could not be opened or read, before any format was chosen.
    #[error("io: {0}")]
    Io(String),

    /// The `SpreadsheetML` 2003 document is malformed or not one at all.
    #[error("xml: {0}")]
    Xml2003(String),

    /// Goal Seek was pointed at cells it cannot work with: a set cell with
    /// no formula, or a changing cell that holds one.
    #[error("goal seek: {0}")]
    GoalSeek(String),
}

/// Result of a crate operation.
pub type Result<T> = core::result::Result<T, Error>;

/// An error code held *as a value* in a cell - not an operation failure.
///
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CellError {
    /// `#NULL!` - range intersection is empty.
    Null,
    /// `#DIV/0!` - division by zero.
    Div0,
    /// `#VALUE!` - wrong argument type.
    Value,
    /// `#REF!` - reference to a cell that no longer exists.
    Ref,
    /// `#NAME?` - unknown name or function.
    Name,
    /// `#NUM!` - invalid numeric value.
    Num,
    /// `#N/A` - value not available.
    Na,
    /// `#CALC!` - array calculation error.
    Calc,
}

impl CellError {
    /// The literal Excel writes for this error.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Null => "#NULL!",
            Self::Div0 => "#DIV/0!",
            Self::Value => "#VALUE!",
            Self::Ref => "#REF!",
            Self::Name => "#NAME?",
            Self::Num => "#NUM!",
            Self::Na => "#N/A",
            Self::Calc => "#CALC!",
        }
    }

    /// Parses an error literal. Case-sensitive, as Excel writes it.
    #[must_use]
    pub fn parse(s: &str) -> Option<Self> {
        Some(match s {
            "#NULL!" => Self::Null,
            "#DIV/0!" => Self::Div0,
            "#VALUE!" => Self::Value,
            "#REF!" => Self::Ref,
            "#NAME?" => Self::Name,
            "#NUM!" => Self::Num,
            "#N/A" => Self::Na,
            "#CALC!" => Self::Calc,
            _ => return None,
        })
    }
}

impl core::fmt::Display for CellError {
    fn fmt(&self, f: &mut core::fmt::Formatter<'_>) -> core::fmt::Result {
        f.write_str(self.as_str())
    }
}

#[cfg(test)]
mod tests {
    use super::CellError;

    #[test]
    fn error_codes_roundtrip() {
        for e in [
            CellError::Null,
            CellError::Div0,
            CellError::Value,
            CellError::Ref,
            CellError::Name,
            CellError::Num,
            CellError::Na,
            CellError::Calc,
        ] {
            assert_eq!(CellError::parse(e.as_str()), Some(e));
        }
        assert_eq!(CellError::parse("#OOPS!"), None);
        assert_eq!(
            CellError::parse("#n/a"),
            None,
            "Excel error literals are upper case"
        );
    }
}
