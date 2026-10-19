//! The Data tab's tools ported from gridcraft: Remove Duplicates, Text to
//! Columns, Goal Seek and data validation checks.

#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use excelerate::coordinate::{CellRef, Col, Range};
use excelerate::edit::{Deduplicated, SplitBy, remove_duplicates, text_to_columns};
use excelerate::formula::eval::recalculate;
use excelerate::formula::validation::{check, invalid_cells, list_items, rule_at};
use excelerate::formula::what_if::{SeekLimits, goal_seek, goal_seek_with};
use excelerate::model::{
    CellValue, DataValidation, Spreadsheet, ValidationOperator, ValidationType, Worksheet,
};
use excelerate::progress::Options;
use excelerate::style::{NumberFormat, Style};
use excelerate::{Error, at};

fn cell(address: &str) -> CellRef {
    CellRef::parse(address).unwrap()
}

fn area(range: &str) -> Range {
    Range::parse(range).unwrap()
}

/// What a cell says: a formula as its text, a value as its display.
fn text(book: &Spreadsheet, address: &str) -> String {
    match book.sheet(0).unwrap().get(cell(address)).map(|c| &c.value) {
        Some(CellValue::Formula { formula, .. }) => format!("={formula}"),
        Some(CellValue::Number(n)) => n.to_string(),
        Some(CellValue::Text(t)) => t.to_string(),
        Some(CellValue::Bool(b)) => b.to_string().to_uppercase(),
        Some(CellValue::Empty) | None => String::new(),
        other => format!("{other:?}"),
    }
}

fn book_of(rows: &[&[&str]]) -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("Data").unwrap();
    for (r, row) in (1..).zip(rows) {
        for (c, value) in (1..).zip(row.iter()) {
            if value.is_empty() {
                continue;
            }
            let value = match value.strip_prefix('=') {
                Some(f) => CellValue::formula(f),
                None => value
                    .parse::<f64>()
                    .map_or_else(|_| CellValue::text(*value), CellValue::Number),
            };
            ws.set_at(r, c, value).unwrap();
        }
    }
    book.add_sheet(ws).unwrap();
    book
}

// ------------------------------------------------------------ duplicates

#[test]
fn duplicates_go_and_the_rows_close_up() {
    let mut book = book_of(&[
        &["Name", "City", "Note"],
        &["Ann", "Oslo", "1"],
        &["Bob", "Rome", "2"],
        &["ann", "OSLO", "3"],
        &["Cid", "Rome", "4"],
        &["Bob", "Rome", "5"],
    ]);
    let done = remove_duplicates(&mut book, 0, area("A1:B6"), &[], true).unwrap();
    assert_eq!(
        done,
        Deduplicated {
            removed: 2,
            remaining: 3
        }
    );
    let column = |c: &str| {
        (1..=6)
            .map(|r| text(&book, &format!("{c}{r}")))
            .collect::<Vec<_>>()
    };
    assert_eq!(column("A"), ["Name", "Ann", "Bob", "Cid", "", ""]);
    assert_eq!(column("B"), ["City", "Oslo", "Rome", "Rome", "", ""]);
    // Column C was outside the range and stays as it was.
    assert_eq!(column("C"), ["Note", "1", "2", "3", "4", "5"]);
}

#[test]
fn duplicates_on_chosen_columns_and_formulas_follow_their_row() {
    let mut book = book_of(&[
        &["a", "1", "=B1*10"],
        &["a", "2", "=B2*10"],
        &["b", "3", "=B3*10"],
    ]);
    let done = remove_duplicates(
        &mut book,
        0,
        area("A1:C3"),
        &[Col::from_letters("A").unwrap()],
        false,
    )
    .unwrap();
    assert_eq!(done.removed, 1);
    assert_eq!(text(&book, "B2"), "3");
    assert_eq!(text(&book, "C2"), "=B2*10");
    assert_eq!(text(&book, "C3"), "");
}

#[test]
fn duplicates_compare_what_cells_show() {
    let mut book = book_of(&[&["1.01"], &["1.02"], &["1.5"]]);
    let one_place = Style {
        number_format: NumberFormat::Custom("0.0".into()),
        ..Style::default()
    };
    let id = book.styles.intern(one_place);
    for r in 1..=3 {
        book.sheet_mut(0)
            .unwrap()
            .entry(cell(&format!("A{r}")))
            .style = id;
    }
    let done = remove_duplicates(&mut book, 0, area("A1:A3"), &[], false).unwrap();
    assert_eq!(
        done,
        Deduplicated {
            removed: 1,
            remaining: 2
        }
    );
    assert_eq!(text(&book, "A2"), "1.5");
}

#[test]
fn duplicates_on_a_whole_column_only_read_the_data() {
    let mut book = book_of(&[&["x"], &["x"], &["y"]]);
    let done = remove_duplicates(&mut book, 0, area("A:A"), &[], false).unwrap();
    assert_eq!(
        done,
        Deduplicated {
            removed: 1,
            remaining: 2
        }
    );
}

#[test]
fn duplicates_refuse_a_merge_or_a_column_outside() {
    let mut book = book_of(&[&["x", "y"], &["x", "y"]]);
    book.sheet_mut(0).unwrap().merges.push(area("A1:B1"));
    assert!(matches!(
        remove_duplicates(&mut book, 0, area("A1:B2"), &[], false),
        Err(Error::InvalidRange(_))
    ));
    assert!(matches!(
        remove_duplicates(
            &mut book,
            0,
            area("A2:B2"),
            &[Col::from_letters("D").unwrap()],
            false
        ),
        Err(Error::InvalidRange(_))
    ));
}

// ------------------------------------------------------- text to columns

#[test]
fn text_splits_into_typed_columns_in_place() {
    let mut book = book_of(&[&["Ann;42;TRUE"], &["\"Smith; Bob\";7"], &[""], &["Cid;;x"]]);
    let width = text_to_columns(
        &mut book,
        0,
        area("A1:A4"),
        at!("A1"),
        &SplitBy::delimiter(';'),
    )
    .unwrap();
    assert_eq!(width, 3);
    assert_eq!(text(&book, "A1"), "Ann");
    assert_eq!(text(&book, "B1"), "42");
    assert_eq!(text(&book, "C1"), "TRUE");
    assert_eq!(text(&book, "A2"), "Smith; Bob");
    assert_eq!(text(&book, "B2"), "7");
    assert_eq!(text(&book, "A4"), "Cid");
    assert_eq!(text(&book, "B4"), "");
    assert_eq!(text(&book, "C4"), "x");
}

#[test]
fn text_splits_by_width_to_another_place() {
    let mut book = book_of(&[&["2024 Q1  15"], &["2025 Q3 230"]]);
    text_to_columns(
        &mut book,
        0,
        area("A1:A2"),
        at!("C1"),
        &SplitBy::FixedWidth(vec![4, 7]),
    )
    .unwrap();
    assert_eq!(text(&book, "A1"), "2024 Q1  15");
    assert_eq!(text(&book, "C1"), "2024");
    assert_eq!(text(&book, "D1"), "Q1");
    assert_eq!(text(&book, "E1"), "15");
    assert_eq!(text(&book, "E2"), "230");
}

#[test]
fn text_that_would_run_off_the_sheet_writes_nothing() {
    let mut book = book_of(&[&["a,b,c"]]);
    let last = CellRef::parse("XFD1").unwrap();
    book.sheet_mut(0).unwrap().set(last, "a,b");
    assert_eq!(
        text_to_columns(
            &mut book,
            0,
            area("XFD1:XFD1"),
            last,
            &SplitBy::delimiter(',')
        ),
        Err(Error::WouldPushOffSheet)
    );
    assert_eq!(text(&book, "XFD1"), "a,b");
}

// -------------------------------------------------------------- goal seek

#[test]
fn goal_seek_finds_the_input_through_a_chain() {
    // B1 = A1 squared, B2 = B1 + 1: what A1 makes B2 equal 10?
    let mut book = book_of(&[&["1", "=A1*A1"], &["", "=B1+1"]]);
    recalculate(&mut book, None, &Options::default());
    let found = goal_seek(&mut book, (0, at!("B2")), 10.0, (0, at!("A1"))).unwrap();
    assert!(found.converged, "{found:?}");
    assert!((found.value - 3.0).abs() < 1e-3, "{found:?}");
    // The answer is left in the changing cell and what reads it is current.
    let a1 = book
        .sheet(0)
        .unwrap()
        .get(at!("A1"))
        .unwrap()
        .value
        .as_number()
        .unwrap();
    assert!((a1 - found.value).abs() < f64::EPSILON);
    let b2 = book
        .sheet(0)
        .unwrap()
        .get(at!("B2"))
        .unwrap()
        .value
        .as_number()
        .unwrap();
    assert!((b2 - 10.0).abs() <= 0.001, "{b2}");
}

#[test]
fn goal_seek_solves_a_loan_payment() {
    // What rate makes a 10 000 loan over 36 months cost 300 a month?
    let mut book = book_of(&[&["0.01", "36", "10000", "=PMT(A1,B1,-C1)"]]);
    let limits = SeekLimits {
        max_iterations: 200,
        tolerance: 1e-7,
    };
    let found = goal_seek_with(
        &mut book,
        (0, at!("D1")),
        300.0,
        (0, at!("A1")),
        limits,
        &Options::default(),
    )
    .unwrap();
    assert!(found.converged, "{found:?}");
    assert!((found.result - 300.0).abs() <= 1e-7, "{found:?}");
    // 36 payments of 300 repay 10 800: about 0.42% a month.
    assert!((0.0042..0.0043).contains(&found.value), "{found:?}");
}

#[test]
fn goal_seek_reads_the_workbook_limits() {
    let mut book = book_of(&[&["1"]]);
    book.calculation_properties
        .push(("iterateCount".into(), "7".into()));
    book.calculation_properties
        .push(("iterateDelta".into(), "0.5".into()));
    assert_eq!(
        SeekLimits::of(&book),
        SeekLimits {
            max_iterations: 7,
            tolerance: 0.5
        }
    );
    assert_eq!(SeekLimits::of(&book_of(&[])), SeekLimits::default());
}

#[test]
fn goal_seek_needs_a_formula_and_a_value() {
    let mut book = book_of(&[&["1", "=A1*2", "5"]]);
    assert!(matches!(
        goal_seek(&mut book, (0, at!("C1")), 3.0, (0, at!("A1"))),
        Err(Error::GoalSeek(_))
    ));
    assert!(matches!(
        goal_seek(&mut book, (0, at!("B1")), 3.0, (0, at!("B1"))),
        Err(Error::GoalSeek(_))
    ));
    assert!(matches!(
        goal_seek(&mut book, (0, at!("B1")), f64::NAN, (0, at!("A1"))),
        Err(Error::GoalSeek(_))
    ));
    assert_eq!(
        goal_seek(&mut book, (3, at!("B1")), 3.0, (0, at!("A1"))),
        Err(Error::SheetIndexOutOfRange(3))
    );
    assert_eq!(text(&book, "A1"), "1");
}

// ------------------------------------------------------------- validation

fn rule(
    range: &str,
    kind: ValidationType,
    operator: ValidationOperator,
    f1: &str,
    f2: &str,
) -> DataValidation {
    DataValidation {
        sqref: vec![area(range)],
        kind,
        operator,
        formula1: f1.into(),
        formula2: f2.into(),
        allow_blank: true,
        show_error_message: true,
        ..DataValidation::default()
    }
}

#[test]
fn whole_numbers_between_bounds() {
    let mut book = book_of(&[&["5", "10"]]);
    book.sheet_mut(0).unwrap().data_validations.push(rule(
        "A1:A10",
        ValidationType::Whole,
        ValidationOperator::Between,
        "1",
        "$B$1",
    ));
    let ok = |v: CellValue| check(&book, 0, at!("A3"), &v).is_none();
    assert!(ok(CellValue::Number(1.0)));
    assert!(ok(CellValue::Number(10.0)));
    assert!(!ok(CellValue::Number(11.0)));
    assert!(!ok(CellValue::Number(2.5)));
    assert!(!ok(CellValue::text("3")));
    assert!(ok(CellValue::Empty));
    // Outside the rule's range anything goes.
    assert!(check(&book, 0, at!("C1"), &CellValue::Number(99.0)).is_none());
}

#[test]
fn bounds_move_with_the_cell() {
    // Each row's value must stay under the limit beside it in column B.
    let mut book = book_of(&[&["1", "5"], &["9", "20"], &["9", "3"]]);
    book.sheet_mut(0).unwrap().data_validations.push(rule(
        "A1:A3",
        ValidationType::Decimal,
        ValidationOperator::LessThan,
        "B1",
        "",
    ));
    assert_eq!(invalid_cells(&book, 0), vec![at!("A3")]);
}

#[test]
fn lists_literal_and_from_a_range() {
    let mut book = book_of(&[&["Red", ""], &["Green", ""], &["", ""]]);
    let ws = book.sheet_mut(0).unwrap();
    ws.data_validations.push(rule(
        "B1:B5",
        ValidationType::List,
        ValidationOperator::Between,
        "\"Yes,No, Maybe\"",
        "",
    ));
    ws.data_validations.push(rule(
        "C1:C5",
        ValidationType::List,
        ValidationOperator::Between,
        "$A$1:$A$3",
        "",
    ));
    let yes_no = rule_at(&book, 0, at!("B2")).unwrap();
    assert_eq!(list_items(&book, 0, yes_no), ["Yes", "No", "Maybe"]);
    let colours = rule_at(&book, 0, at!("C2")).unwrap();
    assert_eq!(list_items(&book, 0, colours), ["Red", "Green"]);
    assert!(check(&book, 0, at!("B1"), &CellValue::text("maybe")).is_none());
    assert!(check(&book, 0, at!("B1"), &CellValue::text("Perhaps")).is_some());
    assert!(check(&book, 0, at!("C4"), &CellValue::text("GREEN")).is_none());
    assert!(check(&book, 0, at!("C4"), &CellValue::text("Blue")).is_some());
}

#[test]
fn custom_rules_see_the_new_value() {
    let mut book = book_of(&[&["4"]]);
    let mut even = rule(
        "A1:A9",
        ValidationType::Custom,
        ValidationOperator::Between,
        "MOD(A1,2)=0",
        "",
    );
    even.error = "Even numbers only".into();
    book.sheet_mut(0).unwrap().data_validations.push(even);
    assert!(check(&book, 0, at!("A5"), &CellValue::Number(8.0)).is_none());
    let broken = check(&book, 0, at!("A5"), &CellValue::Number(7.0)).unwrap();
    assert_eq!(broken.error, "Even numbers only");
    // The book itself is not touched by the check.
    assert_eq!(text(&book, "A5"), "");
}

#[test]
fn text_length_and_blanks() {
    let mut book = book_of(&[&["abc"], &["abcdefgh"]]);
    let mut short = rule(
        "A1:A5",
        ValidationType::TextLength,
        ValidationOperator::LessThanOrEqual,
        "5",
        "",
    );
    short.allow_blank = false;
    book.sheet_mut(0).unwrap().data_validations.push(short);
    assert_eq!(invalid_cells(&book, 0), vec![at!("A2")]);
    assert!(check(&book, 0, at!("A3"), &CellValue::Empty).is_some());
    assert!(check(&book, 0, at!("A3"), &CellValue::Number(12345.0)).is_none());
}
