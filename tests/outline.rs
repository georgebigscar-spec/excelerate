//! The outline: grouping, hiding and showing detail, the level buttons and
//! Data - Subtotal.

#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use std::io::{Cursor, Read};

use excelerate::coordinate::{Col, Range, Row};
use excelerate::edit::{
    OutlineGroup, SubtotalFunction, Subtotals, collapse_rows, column_groups, expand_rows,
    group_columns, group_rows, row_groups, show_outline_level, subtotal, ungroup_rows,
};
use excelerate::formula::eval::recalculate;
use excelerate::model::{CellValue, Spreadsheet, Worksheet};
use excelerate::progress::Options;
use excelerate::{Error, at};

fn row(n: u32) -> Row {
    Row::from_one_based(u64::from(n)).unwrap()
}

fn col(letters: &str) -> Col {
    Col::from_letters(letters).unwrap()
}

fn book() -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    book.add_sheet(Worksheet::new("Data").unwrap()).unwrap();
    book
}

/// `(hidden, collapsed)` of 1-based rows.
fn state(book: &Spreadsheet, rows: std::ops::RangeInclusive<u32>) -> Vec<(bool, bool)> {
    let ws = book.sheet(0).unwrap();
    rows.map(|n| {
        ws.rows
            .get(&row(n))
            .map_or((false, false), |p| (p.hidden, p.collapsed))
    })
    .collect()
}

fn level(book: &Spreadsheet, n: u32) -> u8 {
    book.sheet(0).unwrap().row_outline_level(row(n))
}

#[test]
fn groups_nest_and_ungroup() {
    let mut book = book();
    group_rows(&mut book, 0, row(5), row(8)).unwrap();
    group_rows(&mut book, 0, row(6), row(7)).unwrap();
    let ws = book.sheet(0).unwrap();
    assert_eq!(
        row_groups(ws),
        [
            OutlineGroup {
                level: 1,
                first: 4,
                last: 7
            },
            OutlineGroup {
                level: 2,
                first: 5,
                last: 6
            },
        ]
    );
    ungroup_rows(&mut book, 0, row(5), row(8)).unwrap();
    assert_eq!((level(&book, 5), level(&book, 6)), (0, 1));
    ungroup_rows(&mut book, 0, row(1), row(100)).unwrap();
    assert_eq!(row_groups(book.sheet(0).unwrap()), []);
    // Nothing worth keeping is left behind in the row map.
    assert!(book.sheet(0).unwrap().rows.is_empty());
}

#[test]
fn seven_levels_and_no_more() {
    let mut book = book();
    for _ in 0..7 {
        group_rows(&mut book, 0, row(2), row(3)).unwrap();
    }
    assert!(matches!(
        group_rows(&mut book, 0, row(1), row(3)),
        Err(Error::InvalidRange(_))
    ));
    // The failed call changed nothing, row 1 included.
    assert_eq!((level(&book, 1), level(&book, 2)), (0, 7));
}

#[test]
fn collapsing_keeps_inner_groups_as_they_were() {
    let mut book = book();
    group_rows(&mut book, 0, row(5), row(8)).unwrap();
    group_rows(&mut book, 0, row(6), row(7)).unwrap();
    // The inner group's summary is row 8: collapse it on its own first.
    assert!(collapse_rows(&mut book, 0, row(8)).unwrap());
    assert_eq!(
        state(&book, 6..=8),
        [(true, false), (true, false), (false, true)]
    );
    // Row 9 is the outer group's summary.
    assert!(collapse_rows(&mut book, 0, row(9)).unwrap());
    assert!(state(&book, 5..=8).iter().all(|&(hidden, _)| hidden));
    assert_eq!(state(&book, 9..=9), [(false, true)]);
    // Opening the outer group leaves the inner one shut.
    assert!(expand_rows(&mut book, 0, row(9)).unwrap());
    assert_eq!(
        state(&book, 5..=9),
        [
            (false, false),
            (true, false),
            (true, false),
            (false, true),
            (false, false)
        ]
    );
    assert!(!expand_rows(&mut book, 0, row(30)).unwrap());
}

#[test]
fn level_buttons() {
    let mut book = book();
    group_rows(&mut book, 0, row(2), row(9)).unwrap();
    group_rows(&mut book, 0, row(3), row(5)).unwrap();
    show_outline_level(&mut book, 0, Some(2), None).unwrap();
    assert_eq!(
        state(&book, 2..=6),
        [
            (false, false),
            (true, false),
            (true, false),
            (true, false),
            (false, true)
        ]
    );
    show_outline_level(&mut book, 0, Some(1), None).unwrap();
    assert!(state(&book, 2..=9).iter().all(|&(hidden, _)| hidden));
    assert_eq!(state(&book, 10..=10), [(false, true)]);
    show_outline_level(&mut book, 0, Some(8), None).unwrap();
    assert!(state(&book, 2..=10).iter().all(|&s| s == (false, false)));
    assert!(show_outline_level(&mut book, 0, Some(0), None).is_err());
}

#[test]
fn columns_group_too() {
    let mut book = book();
    book.sheet_mut(0)
        .unwrap()
        .set_column_width(col("C"), Some(20.0));
    group_columns(&mut book, 0, col("B"), col("D")).unwrap();
    let ws = book.sheet(0).unwrap();
    assert_eq!(
        column_groups(ws),
        [OutlineGroup {
            level: 1,
            first: 1,
            last: 3
        }]
    );
    // The width of C survives being split out of its run.
    assert_eq!(ws.column_width(col("C")), Some(20.0));
}

/// Region, Rep, Amount: three regions, sorted.
fn sales() -> Spreadsheet {
    let mut book = book();
    let rows: [(&str, &str, f64); 6] = [
        ("East", "Ann", 10.0),
        ("East", "Bob", 20.0),
        ("north", "Cid", 5.0),
        ("North", "Dee", 7.0),
        ("West", "Eve", 1.0),
        ("West", "Fay", 2.0),
    ];
    let ws = book.sheet_mut(0).unwrap();
    ws.set(at!("A1"), "Region");
    ws.set(at!("B1"), "Rep");
    ws.set(at!("C1"), "Amount");
    for (r, (region, rep, amount)) in (2..).zip(rows) {
        ws.set_at(r, 1, region).unwrap();
        ws.set_at(r, 2, rep).unwrap();
        ws.set_at(r, 3, amount).unwrap();
    }
    // A formula below the list, which must follow the rows it reads.
    ws.set(at!("E20"), CellValue::formula("C7"));
    book
}

fn shown(book: &Spreadsheet, address: &str) -> String {
    book.formatted(0, excelerate::coordinate::CellRef::parse(address).unwrap())
}

#[test]
fn subtotals_by_region() {
    let mut book = sales();
    let groups = subtotal(
        &mut book,
        0,
        Range::parse("A1:C7").unwrap(),
        &Subtotals {
            group_by: col("A"),
            function: SubtotalFunction::Sum,
            columns: vec![col("C")],
        },
    )
    .unwrap();
    assert_eq!(groups, 3);
    recalculate(&mut book, None, &Options::default());

    let column = |c: &str| {
        (1..=11)
            .map(|r| shown(&book, &format!("{c}{r}")))
            .collect::<Vec<_>>()
    };
    assert_eq!(
        column("A"),
        [
            "Region",
            "East",
            "East",
            "East Total",
            "north",
            "North",
            "north Total",
            "West",
            "West",
            "West Total",
            "Grand Total"
        ]
    );
    assert_eq!(
        column("C"),
        [
            "Amount", "10", "20", "30", "5", "7", "12", "1", "2", "3", "45"
        ]
    );
    let ws = book.sheet(0).unwrap();
    let formula = |address: &str| match &ws
        .get(excelerate::coordinate::CellRef::parse(address).unwrap())
        .unwrap()
        .value
    {
        CellValue::Formula { formula, .. } => formula.clone(),
        other => panic!("{other:?}"),
    };
    assert_eq!(formula("C4"), "SUBTOTAL(9,C2:C3)");
    // The grand total spans the group totals and does not count them twice.
    assert_eq!(formula("C11"), "SUBTOTAL(9,C2:C10)");
    // The formula below moved down with the row it read.
    assert_eq!(formula("E24"), "C9");
    // Labels are bold.
    let bold = |address: &str| {
        let id = ws
            .get(excelerate::coordinate::CellRef::parse(address).unwrap())
            .unwrap()
            .style;
        book.styles.get(id).unwrap().font.bold
    };
    assert!(bold("A4") && bold("A11") && !bold("A2"));
    // Outline: details at 2, totals at 1, the grand total and header at 0.
    let levels: Vec<u8> = (1..=11).map(|r| level(&book, r)).collect();
    assert_eq!(levels, [0, 2, 2, 1, 2, 2, 1, 2, 2, 1, 0]);
}

#[test]
fn subtotal_refuses_what_it_cannot_do() {
    let mut book = sales();
    let how = |group_by: &str, columns: &[&str]| Subtotals {
        group_by: col(group_by),
        function: SubtotalFunction::Count,
        columns: columns.iter().map(|c| col(c)).collect(),
    };
    let area = Range::parse("A1:C7").unwrap();
    assert!(subtotal(&mut book, 0, area, &how("D", &["C"])).is_err());
    assert!(subtotal(&mut book, 0, area, &how("A", &[])).is_err());
    assert!(
        subtotal(
            &mut book,
            0,
            Range::parse("A1:C1").unwrap(),
            &how("A", &["C"])
        )
        .is_err()
    );
    assert_eq!(shown(&book, "A4"), "north");
}

#[test]
fn the_outline_levels_reach_the_file() {
    let mut book = sales();
    subtotal(
        &mut book,
        0,
        Range::parse("A1:C7").unwrap(),
        &Subtotals {
            group_by: col("A"),
            function: SubtotalFunction::Average,
            columns: vec![col("C")],
        },
    )
    .unwrap();
    group_columns(&mut book, 0, col("B"), col("B")).unwrap();
    let mut bytes = Vec::new();
    excelerate::writer::xlsx::write_xlsx_to(&book, Cursor::new(&mut bytes)).unwrap();

    let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
    let mut xml = String::new();
    zip.by_name("xl/worksheets/sheet1.xml")
        .unwrap()
        .read_to_string(&mut xml)
        .unwrap();
    assert!(xml.contains(r#"outlineLevelRow="2""#), "{xml}");
    assert!(xml.contains(r#"outlineLevelCol="1""#), "{xml}");
    assert!(xml.contains(r#"defaultRowHeight="15""#));

    let back = excelerate::reader::xlsx::read_xlsx_from(Cursor::new(bytes)).unwrap();
    let levels: Vec<u8> = (1..=11).map(|r| level(&back, r)).collect();
    assert_eq!(levels, [0, 2, 2, 1, 2, 2, 1, 2, 2, 1, 0]);
    assert_eq!(shown(&back, "A11"), "Grand Average");
}
