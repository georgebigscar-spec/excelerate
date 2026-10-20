//! Flash Fill on a sheet.

#![allow(clippy::unwrap_used)]

use excelerate::at;
use excelerate::edit::flash_fill;
use excelerate::model::{CellValue, Spreadsheet, Worksheet};

#[test]
fn fills_the_column_by_example_past_a_header() {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("People").unwrap();
    let rows = [
        ("First", "Last", "Short"),
        ("Ada", "Lovelace", "Lovelace A."),
        ("Alan", "Turing", ""),
        ("Grace", "Hopper", ""),
        ("Edsger", "Dijkstra", ""),
    ];
    for (r, (first, last, short)) in (1..).zip(rows) {
        ws.set_at(r, 1, first).unwrap();
        ws.set_at(r, 2, last).unwrap();
        if !short.is_empty() {
            ws.set_at(r, 3, short).unwrap();
        }
    }
    book.add_sheet(ws).unwrap();

    assert_eq!(flash_fill(&mut book, 0, at!("C3")).unwrap(), 3);
    let shown: Vec<String> = (1..=5).map(|r| book.formatted_at(0, r, 3)).collect();
    assert_eq!(
        shown,
        [
            "Short",
            "Lovelace A.",
            "Turing A.",
            "Hopper G.",
            "Dijkstra E."
        ]
    );
    // Written as text, as Excel writes it.
    assert_eq!(
        book.sheet(0).unwrap().get(at!("C4")).unwrap().value,
        CellValue::text("Hopper G.")
    );
    // Nothing to learn from: nothing filled.
    assert_eq!(flash_fill(&mut book, 0, at!("E2")).unwrap(), 0);
}
