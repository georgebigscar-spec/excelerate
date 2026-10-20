//! Solver on sheet models with answers worked out by hand.

#![allow(clippy::unwrap_used)]

use std::io::Cursor;

use excelerate::coordinate::{CellRef, Range};
use excelerate::formula::eval::recalculate;
use excelerate::formula::solver::{
    Bound, Relation, SolverConstraint, SolverGoal, SolverMethod, SolverModel, SolverStatus,
    save_model, saved_model, solve,
};
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::progress::Options;
use excelerate::{at, reader::xlsx::read_xlsx_from, writer::xlsx::write_xlsx_to};

fn range(r: &str) -> Range {
    Range::parse(r).unwrap()
}

fn sheet(cells: &[(&str, CellValue)]) -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("Plan").unwrap();
    for (a, v) in cells {
        ws.set(CellRef::parse(a).unwrap(), v.clone());
    }
    book.add_sheet(ws).unwrap();
    recalculate(&mut book, None, &Options::default());
    book
}

fn num(book: &Spreadsheet, a: &str) -> f64 {
    book.sheet(0)
        .unwrap()
        .get(CellRef::parse(a).unwrap())
        .unwrap()
        .value
        .as_number()
        .unwrap()
}

fn le(cells: &str, v: f64) -> SolverConstraint {
    SolverConstraint {
        cells: range(cells),
        relation: Relation::LessOrEqual,
        bound: Bound::Value(v),
    }
}

/// x in A1, y in B1; profit 3x + 5y in C1; resources in A2:A3.
fn product_mix() -> (Spreadsheet, SolverModel) {
    let book = sheet(&[
        ("A1", CellValue::Number(0.0)),
        ("B1", CellValue::Number(0.0)),
        ("C1", CellValue::formula("3*A1+5*B1")),
        ("A2", CellValue::formula("2*B1")),
        ("A3", CellValue::formula("3*A1+2*B1")),
    ]);
    let mut model = SolverModel::new(at!("C1"), SolverGoal::Max, vec![range("A1:B1")]);
    model.constraints = vec![le("A1", 4.0), le("A2", 12.0), le("A3", 18.0)];
    (book, model)
}

#[test]
fn simplex_solves_a_product_mix() {
    let (mut book, mut model) = product_mix();
    model.method = SolverMethod::SimplexLp;
    let solved = solve(&mut book, 0, &model).unwrap();
    assert_eq!(solved.status, SolverStatus::Optimal);
    assert!((num(&book, "A1") - 2.0).abs() < 1e-9 && (num(&book, "B1") - 6.0).abs() < 1e-9);
    // The workbook is recomputed with the answer kept.
    assert!((num(&book, "C1") - 36.0).abs() < 1e-9);
    assert!((solved.objective - 36.0).abs() < 1e-9);
}

#[test]
fn grg_finds_the_same_optimum() {
    let (mut book, model) = product_mix();
    let solved = solve(&mut book, 0, &model).unwrap();
    assert!(
        matches!(
            solved.status,
            SolverStatus::Optimal | SolverStatus::Feasible
        ),
        "{solved:?}"
    );
    assert!(
        (num(&book, "C1") - 36.0).abs() < 1e-3,
        "{}",
        num(&book, "C1")
    );
}

#[test]
fn a_knapsack_of_binary_choices() {
    // Values 10, 13, 7, 8 for weights 5, 6, 3, 4; capacity 10: take items 2 and 4 (21).
    let mut cells = vec![
        ("E1", CellValue::formula("SUMPRODUCT(A1:D1,A2:D2)")),
        ("E2", CellValue::formula("SUMPRODUCT(A1:D1,A3:D3)")),
    ];
    let values = [("A2", 10.0), ("B2", 13.0), ("C2", 7.0), ("D2", 8.0)];
    let weights = [("A3", 5.0), ("B3", 6.0), ("C3", 3.0), ("D3", 4.0)];
    for (a, v) in values.iter().chain(&weights) {
        cells.push((a, CellValue::Number(*v)));
    }
    let mut book = sheet(&cells);
    let mut model = SolverModel::new(at!("E1"), SolverGoal::Max, vec![range("A1:D1")]);
    model.method = SolverMethod::SimplexLp;
    model.constraints = vec![
        le("E2", 10.0),
        SolverConstraint {
            cells: range("A1:D1"),
            relation: Relation::Binary,
            bound: Bound::None,
        },
    ];
    let solved = solve(&mut book, 0, &model).unwrap();
    assert_eq!(solved.status, SolverStatus::Optimal);
    let picks: Vec<f64> = ["A1", "B1", "C1", "D1"]
        .iter()
        .map(|a| num(&book, a))
        .collect();
    assert_eq!(picks, [0.0, 1.0, 0.0, 1.0]);
    assert!((num(&book, "E1") - 21.0).abs() < 1e-9);
}

#[test]
fn grg_on_a_curved_model() {
    // The point of x + y <= 2 nearest (3, 1) is (2, 0).
    let mut book = sheet(&[
        ("A1", CellValue::Number(0.0)),
        ("B1", CellValue::Number(0.0)),
        ("C1", CellValue::formula("(A1-3)^2+(B1-1)^2")),
        ("C2", CellValue::formula("A1+B1")),
    ]);
    let mut model = SolverModel::new(at!("C1"), SolverGoal::Min, vec![range("A1:B1")]);
    model.non_negative = false;
    model.constraints = vec![le("C2", 2.0)];
    let solved = solve(&mut book, 0, &model).unwrap();
    assert!(
        matches!(
            solved.status,
            SolverStatus::Optimal | SolverStatus::Feasible
        ),
        "{solved:?}"
    );
    assert!(
        (num(&book, "A1") - 2.0).abs() < 1e-3,
        "{}",
        num(&book, "A1")
    );
    assert!(num(&book, "B1").abs() < 1e-3, "{}", num(&book, "B1"));

    // Simplex refuses a curve, and the cells go back.
    book.sheet_mut(0).unwrap().set(at!("A1"), 0.5);
    model.method = SolverMethod::SimplexLp;
    let refused = solve(&mut book, 0, &model).unwrap();
    assert_eq!(refused.status, SolverStatus::NotLinear);
    assert!((num(&book, "A1") - 0.5).abs() < f64::EPSILON);
}

#[test]
fn value_of_infeasible_and_errors() {
    let (mut book, mut model) = product_mix();
    model.method = SolverMethod::SimplexLp;
    model.goal = SolverGoal::Value(20.0);
    let solved = solve(&mut book, 0, &model).unwrap();
    assert_eq!(solved.status, SolverStatus::Optimal);
    assert!((num(&book, "C1") - 20.0).abs() < 1e-9);

    let (mut book, mut model) = product_mix();
    model.method = SolverMethod::SimplexLp;
    model.constraints.push(SolverConstraint {
        cells: range("A1"),
        relation: Relation::GreaterOrEqual,
        bound: Bound::Value(5.0),
    });
    let solved = solve(&mut book, 0, &model).unwrap();
    assert_eq!(solved.status, SolverStatus::Infeasible);
    assert!(num(&book, "A1").abs() < f64::EPSILON);

    let (mut book, mut model) = product_mix();
    model.objective = at!("A1");
    assert!(solve(&mut book, 0, &model).is_err());
}

#[test]
fn the_model_travels_in_excel_s_names() {
    let (mut book, mut model) = product_mix();
    model.method = SolverMethod::SimplexLp;
    model.constraints.push(SolverConstraint {
        cells: range("A1:B1"),
        relation: Relation::Integer,
        bound: Bound::None,
    });
    model.constraints.push(SolverConstraint {
        cells: range("A3"),
        relation: Relation::LessOrEqual,
        bound: Bound::Cells(range("A2")),
    });
    save_model(&mut book, 0, &model);
    let opt = book
        .defined_names
        .iter()
        .find(|d| d.name == "solver_opt")
        .unwrap();
    assert_eq!(opt.formula, "Plan!$C$1");
    assert!(opt.hidden);

    let mut bytes = Vec::new();
    write_xlsx_to(&book, Cursor::new(&mut bytes)).unwrap();
    let back = read_xlsx_from(Cursor::new(bytes)).unwrap();
    let read = saved_model(&back, 0).unwrap();
    assert_eq!(read, model);
    assert_eq!(saved_model(&back, 1), None);
}
