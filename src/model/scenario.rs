//! Scenarios: named sets of values for a sheet's input cells, as What-If
//! Analysis - Scenario Manager keeps them (`<scenarios>` in xlsx).

use crate::coordinate::{CellRef, Range};

/// One input cell of a scenario and the value the scenario gives it.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ScenarioInput {
    /// The cell.
    pub cell: CellRef,
    /// Its value, as text: Excel keeps a number as its digits.
    pub value: String,
    /// Whether the cell was deleted since the scenario was made.
    pub deleted: bool,
    /// Whether an undo removed it.
    pub undone: bool,
    /// The number format the value was shown in, by built-in id.
    pub number_format_id: Option<u32>,
}

/// One scenario.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Scenario {
    /// Its name, unique on the sheet.
    pub name: String,
    /// Whether it may not be changed while the sheet is protected.
    pub locked: bool,
    /// Whether it is hidden while the sheet is protected.
    pub hidden: bool,
    /// Who made or last changed it.
    pub user: Option<String>,
    /// Its comment.
    pub comment: Option<String>,
    /// The input cells and their values.
    pub inputs: Vec<ScenarioInput>,
}

/// A sheet's scenarios.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Scenarios {
    /// The scenarios, in the Scenario Manager's order.
    pub list: Vec<Scenario>,
    /// Which one was shown last, by index.
    pub current: Option<u32>,
    /// Which one is selected in the Scenario Manager, by index.
    pub show: Option<u32>,
    /// The result cells a summary report reads.
    pub result_cells: Vec<Range>,
}

impl Scenarios {
    /// Whether there is nothing to write.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.list.is_empty()
    }
}
