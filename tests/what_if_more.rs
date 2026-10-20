//! Scenarios, data tables, Advanced Filter and Consolidate.

#![allow(clippy::unwrap_used)]

use std::io::Cursor;

use excelerate::coordinate::{CellRef, Range};
use excelerate::edit::{FilterOutput, SubtotalFunction, advanced_filter, consolidate};
use excelerate::formula::eval::recalculate;
use excelerate::formula::what_if::{add_scenario, data_table, scenario_summary, show_scenario};
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::progress::Options;
use excelerate::{at, reader::xlsx::read_xlsx_from, writer::xlsx::write_xlsx_to};

fn area(r: &str) -> Range {
    Range::parse(r).unwrap()
}

fn shown(book: &Spreadsheet, sheet: usize, a: &str) -> String {
    book.formatted(sheet, CellRef::parse(a).unwrap())
}

/// Price in B1, units in B2, revenue = B1*B2 in B3.
fn model() -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("Model").unwrap();
    ws.set(at!("A1"), "Price");
    ws.set(at!("B1"), 10.0);
    ws.set(at!("A2"), "Units");
    ws.set(at!("B2"), 100.0);
    ws.set(at!("A3"), "Revenue");
    ws.set(at!("B3"), CellValue::formula("B1*B2"));
    book.add_sheet(ws).unwrap();
    recalculate(&mut book, None, &Options::default());
    book
}

#[test]
fn scenarios_are_shown_summarised_and_kept_in_the_file() {
    let mut book = model();
    add_scenario(&mut book, 0, "Base", &[at!("B1"), at!("B2")]).unwrap();
    book.sheet_mut(0).unwrap().set(at!("B1"), 12.0);
    book.sheet_mut(0).unwrap().set(at!("B2"), 80.0);
    add_scenario(&mut book, 0, "Premium", &[at!("B1"), at!("B2")]).unwrap();
    assert!(add_scenario(&mut book, 0, "Bad", &[at!("B3")]).is_err());

    assert_eq!(show_scenario(&mut book, 0, "base").unwrap(), 2);
    assert_eq!(shown(&book, 0, "B3"), "1000");
    show_scenario(&mut book, 0, "Premium").unwrap();
    assert_eq!(shown(&book, 0, "B3"), "960");
    assert!(show_scenario(&mut book, 0, "Nope").is_err());

    let report = scenario_summary(&mut book, 0, &[at!("B3")]).unwrap();
    assert_eq!(book.sheet(report).unwrap().title(), "Scenario Summary");
    let row = |r: u32| {
        (1..=5)
            .map(|c| book.formatted_at(report, r, c))
            .collect::<Vec<_>>()
    };
    assert_eq!(row(2), ["", "", "Current Values:", "Base", "Premium"]);
    assert_eq!(row(4), ["", "B1", "12", "10", "12"]);
    assert_eq!(row(5), ["", "B2", "80", "100", "80"]);
    assert_eq!(row(6), ["Result Cells:", "", "", "", ""]);
    assert_eq!(row(7), ["", "B3", "960", "1000", "960"]);
    // The summary left the model as it was.
    assert_eq!(shown(&book, 0, "B1"), "12");

    let mut bytes = Vec::new();
    write_xlsx_to(&book, Cursor::new(&mut bytes)).unwrap();
    let back = read_xlsx_from(Cursor::new(bytes)).unwrap();
    let kept = &back.sheet(0).unwrap().scenarios;
    assert_eq!(kept, &book.sheet(0).unwrap().scenarios);
    assert_eq!(kept.list[0].inputs[1].value, "100");
    assert_eq!(kept.current, Some(1));
}

#[test]
fn data_tables_of_one_and_two_inputs() {
    // Prices down D2:D4, the revenue formula over them in E1.
    let mut book = model();
    let ws = book.sheet_mut(0).unwrap();
    ws.set(at!("E1"), CellValue::formula("B3"));
    for (r, p) in (2..).zip([5.0, 10.0, 20.0]) {
        ws.set_at(r, 4, p).unwrap();
    }
    assert_eq!(
        data_table(&mut book, 0, area("D1:E4"), None, Some(at!("B1"))).unwrap(),
        3
    );
    let col: Vec<String> = (2..=4).map(|r| book.formatted_at(0, r, 5)).collect();
    assert_eq!(col, ["500", "1000", "2000"]);
    // The input cell kept its value.
    assert_eq!(shown(&book, 0, "B1"), "10");

    // Prices across H1:J1 and units down G2:G3, the formula in the corner.
    let ws = book.sheet_mut(0).unwrap();
    ws.set(at!("G1"), CellValue::formula("B3"));
    for (c, p) in (8..).zip([1.0, 2.0, 3.0]) {
        ws.set_at(1, c, p).unwrap();
    }
    ws.set(at!("G2"), 10.0);
    ws.set(at!("G3"), 20.0);
    data_table(
        &mut book,
        0,
        area("G1:J3"),
        Some(at!("B1")),
        Some(at!("B2")),
    )
    .unwrap();
    let grid: Vec<Vec<String>> = (2..=3)
        .map(|r| (8..=10).map(|c| book.formatted_at(0, r, c)).collect())
        .collect();
    assert_eq!(grid, [["10", "20", "30"], ["20", "40", "60"]]);

    assert!(data_table(&mut book, 0, area("D1:E4"), None, None).is_err());
    assert!(data_table(&mut book, 0, area("A1:B3"), None, Some(at!("B1"))).is_err());
}

/// Region, Rep, Sales from A1.
fn list() -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("List").unwrap();
    let rows: [(&str, &str, f64); 6] = [
        ("East", "Ann", 500.0),
        ("West", "Bob", 1200.0),
        ("East", "Carl", 900.0),
        ("Eastern", "Dee", 50.0),
        ("West", "Bob", 1200.0),
        ("North", "Abe", 700.0),
    ];
    for (c, h) in (1..).zip(["Region", "Rep", "Sales"]) {
        ws.set_at(1, c, h).unwrap();
    }
    for (r, (region, rep, sales)) in (2..).zip(rows) {
        ws.set_at(r, 1, region).unwrap();
        ws.set_at(r, 2, rep).unwrap();
        ws.set_at(r, 3, sales).unwrap();
    }
    book.add_sheet(ws).unwrap();
    book
}

fn hidden(book: &Spreadsheet) -> Vec<u32> {
    book.sheet(0)
        .unwrap()
        .rows
        .iter()
        .filter(|(_, p)| p.hidden)
        .map(|(r, _)| r.one_based())
        .collect()
}

#[test]
fn advanced_filter_reads_criteria_as_excel_does() {
    // (Region begins with East AND Sales > 600) OR Rep is exactly "abe".
    let mut book = list();
    let ws = book.sheet_mut(0).unwrap();
    ws.set(at!("F1"), "region");
    ws.set(at!("G1"), "Sales");
    ws.set(at!("H1"), "Rep");
    ws.set(at!("F2"), "East");
    ws.set(at!("G2"), ">600");
    ws.set(at!("H3"), "=abe");
    let passed = advanced_filter(
        &mut book,
        0,
        area("A1:C7"),
        area("F1:H3"),
        FilterOutput::InPlace,
        false,
    )
    .unwrap();
    assert_eq!(passed, 2);
    assert_eq!(hidden(&book), [2, 3, 5, 6]);

    // Wildcards, and unique records copied elsewhere.
    let mut book = list();
    let ws = book.sheet_mut(0).unwrap();
    ws.set(at!("F1"), "Rep");
    ws.set(at!("F2"), "?o*");
    let passed = advanced_filter(
        &mut book,
        0,
        area("A1:C7"),
        area("F1:F2"),
        FilterOutput::CopyTo(0, at!("J1")),
        true,
    )
    .unwrap();
    assert_eq!(passed, 1);
    let copied: Vec<String> = (1..=3).map(|r| book.formatted_at(0, r, 11)).collect();
    assert_eq!(copied, ["Rep", "Bob", ""]);
    assert_eq!(hidden(&book), Vec::<u32>::new());

    let mut book = list();
    book.sheet_mut(0).unwrap().set(at!("F1"), "Nope");
    assert!(
        advanced_filter(
            &mut book,
            0,
            area("A1:C7"),
            area("F1:F2"),
            FilterOutput::InPlace,
            false
        )
        .is_err()
    );
}

#[test]
fn consolidate_by_label_and_by_position() {
    let mut book = Spreadsheet::empty();
    for (name, rows) in [
        (
            "Q1",
            vec![
                ("", "Jan", "Feb"),
                ("Apples", "1", "2"),
                ("Pears", "3", "4"),
            ],
        ),
        (
            "Q2",
            vec![
                ("", "Feb", "Mar"),
                ("pears", "10", "20"),
                ("Plums", "5", "6"),
            ],
        ),
    ] {
        let mut ws = Worksheet::new(name).unwrap();
        for (r, (a, b, c)) in (1..).zip(rows) {
            for (col, v) in (1..).zip([a, b, c]) {
                if v.is_empty() {
                    continue;
                }
                match v.parse::<f64>() {
                    Ok(n) => ws.set_at(r, col, n).unwrap(),
                    Err(_) => ws.set_at(r, col, v).unwrap(),
                }
            }
        }
        book.add_sheet(ws).unwrap();
    }
    book.add_sheet(Worksheet::new("Total").unwrap()).unwrap();
    let range = consolidate(
        &mut book,
        &[(0, area("A1:C3")), (1, area("A1:C3"))],
        (2, at!("A1")),
        SubtotalFunction::Sum,
        true,
        true,
    )
    .unwrap();
    assert_eq!(range.to_string(), "A1:D4");
    let grid: Vec<Vec<String>> = (1..=4)
        .map(|r| (1..=4).map(|c| book.formatted_at(2, r, c)).collect())
        .collect();
    assert_eq!(
        grid,
        [
            ["", "Jan", "Feb", "Mar"],
            ["Apples", "1", "2", ""],
            ["Pears", "3", "14", "20"],
            ["Plums", "", "5", "6"],
        ]
    );

    let range = consolidate(
        &mut book,
        &[(0, area("B2:C3")), (1, area("B2:C3"))],
        (2, at!("F1")),
        SubtotalFunction::Max,
        false,
        false,
    )
    .unwrap();
    assert_eq!(range.to_string(), "F1:G2");
    assert_eq!(book.formatted_at(2, 1, 6), "10");
    assert_eq!(book.formatted_at(2, 2, 7), "6");
}
