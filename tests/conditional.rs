//! Conditional formatting applied: which rules match a cell and what they
//! give it, and how the HTML writer shows it.

#![allow(clippy::unwrap_used, clippy::expect_used, clippy::panic)]

use excelerate::coordinate::{CellRef, Range};
use excelerate::formula::conditional::{CfLook, ConditionalFormatting};
use excelerate::model::{
    CfOperator, CfRule, CfRuleType, CfScale, CfValue, CfValueType, ConditionalFormat, Spreadsheet,
    Worksheet,
};
use excelerate::style::{Color, DiffFill, DiffFont, DifferentialStyle};

fn at(address: &str) -> CellRef {
    CellRef::parse(address).unwrap()
}

/// A1:A10 holding 1..=10, B1 holding "Apple pie", and two differential
/// styles: 0 a red fill, 1 bold green text.
fn book() -> Spreadsheet {
    let mut book = Spreadsheet::empty();
    let mut ws = Worksheet::new("S").unwrap();
    for n in 1..=10 {
        ws.set_at(n, 1, f64::from(n)).unwrap();
    }
    ws.set(at("B1"), "Apple pie");
    ws.set(at("B2"), "apple");
    ws.set(at("B3"), "pear");
    ws.set(at("B4"), "APPLE");
    book.add_sheet(ws).unwrap();
    book.styles.differential.push(DifferentialStyle {
        fill: Some(DiffFill {
            background: Some(Color::Argb(0xFFFF_0000)),
            ..DiffFill::default()
        }),
        ..DifferentialStyle::default()
    });
    book.styles.differential.push(DifferentialStyle {
        font: Some(DiffFont {
            bold: Some(true),
            color: Some(Color::Argb(0xFF00_8000)),
            ..DiffFont::default()
        }),
        fill: Some(DiffFill {
            background: Some(Color::Argb(0xFF00_00FF)),
            ..DiffFill::default()
        }),
        ..DifferentialStyle::default()
    });
    book
}

fn add(book: &mut Spreadsheet, range: &str, rule: CfRule) {
    book.sheet_mut(0)
        .unwrap()
        .conditional_formats
        .push(ConditionalFormat {
            sqref: vec![Range::parse(range).unwrap()],
            rules: vec![rule],
        });
}

fn rule(kind: CfRuleType, priority: i32, dxf: u32) -> CfRule {
    CfRule {
        kind,
        priority,
        dxf: Some(dxf),
        ..CfRule::default()
    }
}

/// The rows of column `col` from 1 to `last` a rule matched.
fn hits(book: &Spreadsheet, col: &str, last: u32) -> Vec<u32> {
    let mut cf = ConditionalFormatting::new(book, 0);
    (1..=last)
        .filter(|r| !cf.look(at(&format!("{col}{r}"))).rules.is_empty())
        .collect()
}

#[test]
fn cell_value_rules() {
    let mut book = book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            operator: Some(CfOperator::Between),
            formulas: vec!["7".into(), "3".into()],
            ..rule(CfRuleType::CellIs, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [3, 4, 5, 6, 7]);

    let mut book = self::book();
    add(
        &mut book,
        "B1:B4",
        CfRule {
            operator: Some(CfOperator::Equal),
            formulas: vec!["\"Apple\"".into()],
            ..rule(CfRuleType::CellIs, 1, 0)
        },
    );
    assert_eq!(hits(&book, "B", 4), [2, 4]);
}

#[test]
fn formulas_move_with_the_cell() {
    let mut book = book();
    // Each cell against the one above it: written for A2, read for A2..A10.
    add(
        &mut book,
        "A2:A10",
        CfRule {
            formulas: vec!["MOD(A2,3)=0".into()],
            ..rule(CfRuleType::Expression, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [3, 6, 9]);
}

#[test]
fn text_blank_and_error_rules() {
    let mut book = book();
    add(
        &mut book,
        "B1:B5",
        CfRule {
            text: Some("APP".into()),
            ..rule(CfRuleType::ContainsText, 1, 0)
        },
    );
    assert_eq!(hits(&book, "B", 5), [1, 2, 4]);

    let mut book = self::book();
    add(
        &mut book,
        "B1:B5",
        CfRule {
            text: Some("pie".into()),
            ..rule(CfRuleType::EndsWith, 1, 0)
        },
    );
    assert_eq!(hits(&book, "B", 5), [1]);

    let mut book = self::book();
    add(&mut book, "B1:B5", rule(CfRuleType::ContainsBlanks, 1, 0));
    assert_eq!(hits(&book, "B", 5), [5]);
}

#[test]
fn ranked_and_statistical_rules() {
    let mut book = book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            rank: Some(3),
            ..rule(CfRuleType::Top10, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [8, 9, 10]);

    let mut book = self::book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            rank: Some(20),
            percent: true,
            bottom: true,
            ..rule(CfRuleType::Top10, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [1, 2]);

    let mut book = self::book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            above_average: true,
            ..rule(CfRuleType::AboveAverage, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [6, 7, 8, 9, 10]);

    // One population standard deviation of 1..=10 is about 2.87 over 5.5.
    let mut book = self::book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            above_average: true,
            std_dev: Some(1),
            ..rule(CfRuleType::AboveAverage, 1, 0)
        },
    );
    assert_eq!(hits(&book, "A", 10), [9, 10]);

    let mut book = self::book();
    add(&mut book, "B1:B4", rule(CfRuleType::DuplicateValues, 1, 0));
    assert_eq!(hits(&book, "B", 4), [2, 4]);
}

#[test]
fn the_stronger_rule_wins_each_property_and_stop_if_true_stops() {
    let mut book = book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            formulas: vec!["A1>5".into()],
            ..rule(CfRuleType::Expression, 2, 0)
        },
    );
    add(
        &mut book,
        "A1:A10",
        CfRule {
            formulas: vec!["A1>8".into()],
            ..rule(CfRuleType::Expression, 1, 1)
        },
    );
    let mut cf = ConditionalFormatting::new(&book, 0);
    let look = cf.look(at("A9"));
    assert_eq!(look.rules, [1, 0]);
    // Blue from the stronger rule, not red; bold from it too.
    let fill = look.style.fill.unwrap();
    assert_eq!(fill.background, Some(Color::Argb(0xFF00_00FF)));
    assert_eq!(look.style.font.unwrap().bold, Some(true));
    assert_eq!(cf.look(at("A6")).rules, [0]);
    assert!(cf.look(at("A2")).is_empty());

    book.sheet_mut(0).unwrap().conditional_formats[1].rules[0].stop_if_true = true;
    let mut cf = ConditionalFormatting::new(&book, 0);
    assert_eq!(cf.look(at("A9")).rules, [1]);
}

fn scale_rule(scale: CfScale, kind: CfRuleType) -> CfRule {
    CfRule {
        kind,
        priority: 1,
        scale: Some(scale),
        ..CfRule::default()
    }
}

fn point(kind: CfValueType, value: &str) -> CfValue {
    CfValue {
        kind,
        value: value.into(),
        ..CfValue::default()
    }
}

#[test]
fn colour_scales_bars_and_icons() {
    let mut book = book();
    add(
        &mut book,
        "A1:A10",
        scale_rule(
            CfScale::Color {
                values: vec![
                    point(CfValueType::Min, ""),
                    point(CfValueType::Percentile, "50"),
                    point(CfValueType::Max, ""),
                ],
                colors: vec![
                    Color::Argb(0xFFF8_696B),
                    Color::Argb(0xFFFF_EB84),
                    Color::Argb(0xFF63_BE7B),
                ],
            },
            CfRuleType::ColorScale,
        ),
    );
    let mut cf = ConditionalFormatting::new(&book, 0);
    assert_eq!(cf.look(at("A1")).color_scale, Some(0xF8_696B));
    assert_eq!(cf.look(at("A10")).color_scale, Some(0x63_BE7B));
    // 5.5 is the median: halfway between 5 and 6 lies the middle colour.
    let five = cf.look(at("A5")).color_scale.unwrap();
    assert!(five != 0xF8_696B && five != 0xFF_EB84, "{five:06X}");

    let mut book = self::book();
    add(
        &mut book,
        "A1:A10",
        scale_rule(
            CfScale::DataBar {
                values: vec![point(CfValueType::Min, ""), point(CfValueType::Max, "")],
                color: Color::Argb(0xFF63_8EC6),
                min_length: Some(0),
                max_length: Some(100),
                show_value: false,
            },
            CfRuleType::DataBar,
        ),
    );
    let mut cf = ConditionalFormatting::new(&book, 0);
    let bar = cf.look(at("A4")).data_bar.unwrap();
    assert!((bar.length - 3.0 / 9.0).abs() < 1e-9, "{bar:?}");
    assert_eq!(bar.color, 0x63_8EC6);
    assert!(cf.look(at("A4")).hides_value());

    let mut book = self::book();
    let mut thirds = vec![
        point(CfValueType::Percent, "0"),
        point(CfValueType::Percent, "33"),
        point(CfValueType::Percent, "67"),
    ];
    thirds[2].greater_or_equal = true;
    add(
        &mut book,
        "A1:A10",
        scale_rule(
            CfScale::IconSet {
                values: thirds,
                set: Some("3Arrows".into()),
                show_value: true,
                percent: true,
                reverse: false,
            },
            CfRuleType::IconSet,
        ),
    );
    let mut cf = ConditionalFormatting::new(&book, 0);
    let icon = |cf: &mut ConditionalFormatting<'_>, a: &str| cf.look(at(a)).icon.unwrap().index;
    assert_eq!(
        (
            icon(&mut cf, "A1"),
            icon(&mut cf, "A5"),
            icon(&mut cf, "A10")
        ),
        (0, 1, 2)
    );
}

#[test]
fn the_html_page_shows_it() {
    let mut book = book();
    add(
        &mut book,
        "A1:A10",
        CfRule {
            formulas: vec!["A1>8".into()],
            ..rule(CfRuleType::Expression, 1, 1)
        },
    );
    add(
        &mut book,
        "A1:A10",
        scale_rule(
            CfScale::DataBar {
                values: vec![point(CfValueType::Min, ""), point(CfValueType::Max, "")],
                color: Color::Argb(0xFF63_8EC6),
                min_length: None,
                max_length: None,
                show_value: true,
            },
            CfRuleType::DataBar,
        ),
    );
    let mut bytes = Vec::new();
    excelerate::writer::html::write_html_to(
        &book,
        &mut bytes,
        &excelerate::writer::html::HtmlOptions::default(),
    )
    .unwrap();
    let page = String::from_utf8(bytes).unwrap();
    let row = |n: u32| {
        let start = page.find(&format!("<tr class=\"row{n}\">")).unwrap();
        let end = page[start..].find("</tr>").unwrap();
        page[start..start + end].to_owned()
    };
    assert!(row(9).contains("background-color: #0000FF"), "{}", row(9));
    assert!(row(9).contains("font-weight:bold"), "{}", row(9));
    assert!(!row(2).contains("#0000FF"));
    assert!(
        row(2).contains("linear-gradient(to right, #638EC6"),
        "{}",
        row(2)
    );
    assert!(row(2).contains(">2<"), "the value stays: {}", row(2));
    // Nothing for a sheet without rules.
    assert!(CfLook::default().is_empty());
}
