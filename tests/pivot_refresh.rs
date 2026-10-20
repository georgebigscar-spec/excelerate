//! Refreshing a pivot table: the report laid out in the cells.

#![allow(clippy::unwrap_used)]

use std::io::{Cursor, Read};

use excelerate::Range;
use excelerate::edit::{refresh_all_pivots, refresh_pivot};
use excelerate::model::pivot::{
    CacheField, CacheSource, DataField, PivotAxis, PivotCache, PivotField, PivotTable, Subtotal,
};
use excelerate::model::table::Table;
use excelerate::model::{Spreadsheet, Worksheet};
use excelerate::reader::xlsx::read_xlsx_from;
use excelerate::writer::xlsx::write_xlsx_to;

fn fixture() -> Spreadsheet {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/tests/fixtures/pivot.xlsx"
    ))
    .unwrap();
    read_xlsx_from(Cursor::new(bytes)).unwrap()
}

fn grid(book: &Spreadsheet, sheet: usize, range: Range) -> Vec<Vec<String>> {
    (range.start.row.one_based()..=range.end.row.one_based())
        .map(|r| {
            (range.start.col.one_based()..=range.end.col.one_based())
                .map(|c| book.formatted_at(sheet, r, c))
                .collect()
        })
        .collect()
}

#[test]
fn a_pivot_written_without_its_cells_gets_them() {
    let mut book = fixture();
    // As written by a program without a pivot engine: no report in the cells.
    let before = book.sheet(0).unwrap().pivot_tables[0].location.unwrap();
    assert!(
        grid(&book, 0, before)
            .iter()
            .flatten()
            .all(String::is_empty)
    );

    let range = refresh_pivot(&mut book, 0, 0).unwrap();
    assert_eq!(range.to_string(), "F1:J9");
    assert_eq!(
        grid(&book, 0, range),
        [
            ["Сумма продаж", "", "Год", "", ""],
            ["Регион", "Товар", "2024", "2025", "Grand Total"],
            ["Север", "Болт", "150", "", "150"],
            ["", "Гайка", "100", "300", "400"],
            ["Север Total", "", "250", "300", "550"],
            ["Юг", "Болт", "120", "250", "370"],
            ["", "Гайка", "200", "", "200"],
            ["Юг Total", "", "320", "250", "570"],
            ["Grand Total", "", "570", "550", "1120"],
        ]
    );
    let table = &book.sheet(0).unwrap().pivot_tables[0];
    assert_eq!((table.first_data_row, table.first_data_col), (2, 2));

    // The file keeps the cells and still asks Excel to lay the report out.
    let mut bytes = Vec::new();
    write_xlsx_to(&book, Cursor::new(&mut bytes)).unwrap();
    let mut zip = zip::ZipArchive::new(Cursor::new(&bytes)).unwrap();
    let cache_part = book.pivot_caches[0].definition_part.clone();
    let mut cache = String::new();
    zip.by_name(&cache_part)
        .unwrap()
        .read_to_string(&mut cache)
        .unwrap();
    assert!(cache.contains(r#"refreshOnLoad="1""#), "{cache}");
    let back = read_xlsx_from(Cursor::new(bytes)).unwrap();
    assert_eq!(back.formatted_at(0, 9, 10), "1120");
    assert_eq!(
        back.sheet(0).unwrap().pivot_tables[0]
            .location
            .map(|r| r.to_string()),
        Some("F1:J9".into())
    );
}

/// Region, Rep, Units, Price on a sheet, as a table named Sales.
fn sales() -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("Data").unwrap();
    let rows: [(&str, &str, f64, f64); 5] = [
        ("West", "Ann", 3.0, 10.0),
        ("East", "Bob", 1.0, 20.0),
        ("West", "Cid", 5.0, 30.0),
        ("east", "Dee", 2.0, 40.0),
        ("", "Eve", 4.0, 50.0),
    ];
    for (c, name) in (1..).zip(["Region", "Rep", "Units", "Price"]) {
        ws.set_at(1, c, name).unwrap();
    }
    for (r, (region, rep, units, price)) in (2..).zip(rows) {
        if !region.is_empty() {
            ws.set_at(r, 1, region).unwrap();
        }
        ws.set_at(r, 2, rep).unwrap();
        ws.set_at(r, 3, units).unwrap();
        ws.set_at(r, 4, price).unwrap();
    }
    ws.tables.push(Table::new(
        1,
        "Sales",
        Range::parse("A1:D6").unwrap(),
        ["Region", "Rep", "Units", "Price"],
    ));
    book.add_sheet(ws).unwrap();
    book.add_sheet(Worksheet::new("Report").unwrap()).unwrap();
    book.pivot_caches.push(PivotCache {
        id: 1,
        source: CacheSource {
            name: Some("Sales".into()),
            ..CacheSource::default()
        },
        fields: ["Region", "Rep", "Units", "Price"]
            .iter()
            .map(|n| CacheField {
                name: (*n).into(),
                ..CacheField::default()
            })
            .collect(),
        ..PivotCache::default()
    });
    let field = |axis| PivotField {
        axis,
        default_subtotal: true,
        ..PivotField::default()
    };
    book.sheet_mut(1).unwrap().pivot_tables.push(PivotTable {
        name: "ByRegion".into(),
        cache_id: 1,
        location: Some(Range::parse("B2:D4").unwrap()),
        fields: vec![
            field(PivotAxis::Row),
            field(PivotAxis::Unused),
            field(PivotAxis::Values),
            field(PivotAxis::Values),
        ],
        row_fields: vec![0],
        column_fields: vec![-2],
        data_fields: vec![
            DataField {
                field: 2,
                subtotal: Subtotal::Sum,
                ..DataField::default()
            },
            DataField {
                name: Some("Mean price".into()),
                field: 3,
                subtotal: Subtotal::Average,
                ..DataField::default()
            },
        ],
        row_grand_totals: true,
        column_grand_totals: true,
        ..PivotTable::default()
    });
    book
}

#[test]
fn a_pivot_over_a_table_with_two_value_fields() {
    let mut book = sales();
    // Something in the way of the larger report is cleared.
    book.sheet_mut(1).unwrap().set_at(6, 3, "stale").unwrap();
    assert_eq!(refresh_all_pivots(&mut book).unwrap(), 1);
    let range = book.sheet(1).unwrap().pivot_tables[0].location.unwrap();
    assert_eq!(range.to_string(), "B2:D6");
    assert_eq!(
        grid(&book, 1, range),
        [
            ["Region", "Sum of Units", "Mean price"],
            ["East", "3", "30"],
            ["West", "8", "20"],
            ["(blank)", "4", "50"],
            ["Grand Total", "15", "30"],
        ]
    );
}

#[test]
fn what_cannot_be_refreshed() {
    let mut book = sales();
    assert!(refresh_pivot(&mut book, 1, 3).is_err());
    assert!(refresh_pivot(&mut book, 7, 0).is_err());
    book.sheet_mut(1).unwrap().pivot_tables[0]
        .data_fields
        .clear();
    assert!(refresh_pivot(&mut book, 1, 0).is_err());
    let mut book = sales();
    book.pivot_caches[0].source.name = Some("Nowhere".into());
    assert!(refresh_pivot(&mut book, 1, 0).is_err());
}
