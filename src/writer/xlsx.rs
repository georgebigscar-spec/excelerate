//! Writing xlsx workbooks.
//!
//!
//! The parts written are the minimum a reader - ours, Excel, or `LibreOffice` -
//! needs to open the file: content types, package and workbook relationships,
//! the workbook, one part per sheet, the shared string pool and the style
//! table. Every path here is fixed, since we are the ones laying out the
//! package; the reader stays general because other writers are not.
//!
//! XML is written by hand rather than through a serializer: the documents are
//! small, their shape is fixed, and escaping is the only subtlety - which
//! [`escape`] handles in one place.

use crate::error::{Error, Result};
use crate::model::{CellValue, Spreadsheet};
use crate::progress::{Options, Stage};
use crate::style::{
    Alignment, BorderStyle, Borders, Color, DiagonalDirection, Fill, Font, NumberFormat, Pattern,
    Protection, Script, Underline,
};
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::{Seek, Write};

/// Number format ids below this are reserved for Excel's built-ins; custom
/// formats are numbered from here up.
const FIRST_CUSTOM_FORMAT_ID: u16 = 164;

/// Writes a workbook to an xlsx file.
///
/// # Errors
/// [`Error::Xlsx`] if the file cannot be created or written.
pub fn write_xlsx(book: &Spreadsheet, path: impl AsRef<std::path::Path>) -> Result<()> {
    let file = std::fs::File::create(path).map_err(|e| Error::Xlsx(e.to_string()))?;
    write_xlsx_to(book, std::io::BufWriter::new(file))
}

/// Writes a workbook to any seekable sink.
///
/// # Errors
/// [`Error::Xlsx`] if the workbook holds no sheets, or the sink fails.
pub fn write_xlsx_to<W: Write + Seek>(book: &Spreadsheet, sink: W) -> Result<()> {
    write_xlsx_to_with(book, sink, &Options::default())
}

/// The same, reporting its progress as it goes.
///
/// Parts are the unit, and the total is known from the start: the fixed parts,
/// one worksheet each, the notes of the sheets that have them, and whatever
/// the workbook carries unparsed.
///
/// # Errors
/// Same as [`write_xlsx_to`].
pub fn write_xlsx_to_with<W: Write + Seek>(
    book: &Spreadsheet,
    sink: W,
    options: &Options<'_>,
) -> Result<()> {
    if book.sheets().is_empty() {
        return Err(Error::Xlsx("cannot write a workbook with no sheets".into()));
    }
    // Charts and pictures are applied to the parts first; an untouched book
    // passes through.
    let charts = super::pivot::prepare(super::chart_ex::prepare(super::chart::prepare(book)?)?)?;
    let prepared = super::properties::prepare(super::comment::prepare(super::shape::prepare(
        super::image::prepare(charts),
    )));
    let book: &Spreadsheet = &prepared;
    let mut zip = zip::ZipWriter::new(sink);
    let opts = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    let strings = collect_shared_strings(book);

    // Five fixed parts, a part per sheet plus its relationships and notes, and
    // everything carried unparsed.
    let total = 5
        + book.sheets().len() * 2
        + book
            .sheets()
            .iter()
            .filter(|s| !s.comments.is_empty())
            .count()
        + book.sheets().iter().map(|s| s.tables.len()).sum::<usize>()
        + book.parts.len();
    let mut written = 0usize;
    // Table parts are numbered from one across the workbook.
    let mut table_number = 1usize;
    let mut part = |name: &str, body: &str| -> Result<()> {
        options.report(Stage::Writing, written, Some(total), name);
        written += 1;
        zip.start_file::<_, ()>(name, opts)
            .map_err(|e| Error::Xlsx(e.to_string()))?;
        zip.write_all(body.as_bytes())
            .map_err(|e| Error::Xlsx(e.to_string()))
    };

    part("[Content_Types].xml", &content_types(book))?;
    part("_rels/.rels", &root_rels(book))?;
    part("xl/workbook.xml", &workbook(book))?;
    part("xl/_rels/workbook.xml.rels", &workbook_rels(book))?;
    part("xl/styles.xml", &styles(book))?;
    // The workbook's own theme when it had one: a theme colour is an index
    // into it, so substituting ours would repaint every cell that uses one.
    part(
        "xl/theme/theme1.xml",
        book.theme.as_deref().unwrap_or(THEME),
    )?;
    part("xl/sharedStrings.xml", &shared_strings(&strings))?;
    for (i, sheet) in book.sheets().iter().enumerate() {
        // An external hyperlink keeps its address in the sheet's own
        // relationships, so the two parts are written together.
        let outside = external_links(sheet);
        part(
            &format!("xl/worksheets/sheet{}.xml", i + 1),
            &worksheet(sheet, &strings, &outside),
        )?;
        // The notes are modelled, so their part is rebuilt rather than
        // carried; the VML that places them travels whole in `parts`.
        if !sheet.comments.is_empty() {
            part(&format!("xl/comments{}.xml", i + 1), &comments_xml(sheet))?;
        }
        // A table part is numbered across the workbook, not within its sheet.
        let numbers: Vec<usize> = (0..sheet.tables.len()).map(|n| table_number + n).collect();
        for (table, number) in sheet.tables.iter().zip(&numbers) {
            part(&format!("xl/tables/table{number}.xml"), &table_xml(table))?;
        }
        table_number += sheet.tables.len();
        if !outside.is_empty()
            || !sheet.attachments.is_empty()
            || !sheet.comments.is_empty()
            || !numbers.is_empty()
        {
            part(
                &format!("xl/worksheets/_rels/sheet{}.xml.rels", i + 1),
                &sheet_rels(&outside, sheet, i + 1, &numbers),
            )?;
        }
    }

    // Everything the crate does not model yet, byte for byte.
    for opaque in &book.parts {
        options.report(Stage::Writing, written, Some(total), &opaque.path);
        written += 1;
        zip.start_file::<_, ()>(opaque.path.as_str(), opts)
            .map_err(|e| Error::Xlsx(e.to_string()))?;
        zip.write_all(&opaque.data)
            .map_err(|e| Error::Xlsx(e.to_string()))?;
    }

    zip.finish().map_err(|e| Error::Xlsx(e.to_string()))?;
    Ok(())
}

/// Renders a sheet's notes.
///
/// The authors go into a table and each note points at its own by index, which
/// is how the format states it; a note with no author still takes a slot, so
/// the empty author is a table entry like any other.
fn comments_xml(sheet: &crate::model::Worksheet) -> String {
    let mut authors: Vec<&str> = Vec::new();
    for comment in sheet.comments.values() {
        if !authors.contains(&comment.author.as_str()) {
            authors.push(&comment.author);
        }
    }
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            "<authors>"
        ),
        decl = XML_DECL
    );
    for author in &authors {
        let _ = write!(s, "<author>{}</author>", escape(author));
    }
    s.push_str("</authors><commentList>");
    for (at, comment) in &sheet.comments {
        let author = authors
            .iter()
            .position(|a| *a == comment.author)
            .unwrap_or(0);
        let _ = write!(s, r#"<comment ref="{at}" authorId="{author}"><text>"#);
        for run in &comment.text {
            s.push_str("<r>");
            if let Some(font) = &run.font {
                s.push_str(&run_font_xml(font));
            }
            let _ = write!(s, r#"<t xml:space="preserve">{}</t>"#, escape(&run.text));
            s.push_str("</r>");
        }
        s.push_str("</text></comment>");
    }
    s.push_str("</commentList></comments>");
    s
}

/// Renders an `<rPr>`: what one run of a rich string changes about the font.
///
/// The order of the children is the one the schema fixes, not the one the
/// fields happen to be declared in.
fn run_font_xml(font: &crate::style::DiffFont) -> String {
    let mut s = String::from("<rPr>");
    if let Some(name) = &font.name {
        let _ = write!(s, r#"<rFont val="{}"/>"#, escape(name));
    }
    for (tag, on) in [
        ("b", font.bold),
        ("i", font.italic),
        ("strike", font.strike),
    ] {
        if let Some(on) = on {
            let _ = write!(s, r#"<{tag} val="{}"/>"#, u8::from(on));
        }
    }
    if let Some(color) = &font.color {
        let _ = write!(s, "<color{}/>", color_attr(color));
    }
    if let Some(size) = font.size {
        let _ = write!(s, r#"<sz val="{}"/>"#, trim_number(f64::from(size) / 100.0));
    }
    if let Some(u) = font.underline {
        let _ = write!(s, r#"<u val="{}"/>"#, u.as_str());
    }
    match font.script {
        Some(crate::style::Script::Superscript) => {
            s.push_str(r#"<vertAlign val="superscript"/>"#);
        }
        Some(crate::style::Script::Subscript) => s.push_str(r#"<vertAlign val="subscript"/>"#),
        Some(crate::style::Script::Baseline) | None => {}
    }
    if let Some(family) = font.family {
        let _ = write!(s, r#"<family val="{family}"/>"#);
    }
    if let Some(charset) = font.charset {
        let _ = write!(s, r#"<charset val="{charset}"/>"#);
    }
    if let Some(scheme) = font.scheme {
        let _ = write!(s, r#"<scheme val="{}"/>"#, scheme.as_str());
    }
    s.push_str("</rPr>");
    s
}

pub(crate) use super::xmlesc::escape;

/// Builds the shared string pool: text value -> index, in first-seen order.
/// One entry of the string pool.
enum Pooled<'a> {
    /// Text that reads the same all the way through.
    Plain(&'a str),
    /// Text whose formatting changes part way through.
    Rich(&'a [crate::model::TextRun]),
}

/// The pool, and the two indexes that find a cell's place in it.
struct StringPool<'a> {
    entries: Vec<Pooled<'a>>,
    plain: HashMap<&'a str, usize>,
    /// Keyed by a rendering of the runs, because a run list is not hashable
    /// and there are few enough of them for that to cost nothing.
    rich: HashMap<String, usize>,
}

/// A key for a run list, used only to spot two equal ones.
fn rich_key(runs: &[crate::model::TextRun]) -> String {
    format!("{runs:?}")
}

fn collect_shared_strings(book: &Spreadsheet) -> StringPool<'_> {
    let mut pool = StringPool {
        entries: Vec::new(),
        plain: HashMap::new(),
        rich: HashMap::new(),
    };
    for sheet in book.sheets() {
        for (_, cell) in sheet.iter() {
            let next = pool.entries.len();
            match &cell.value {
                CellValue::Text(s) => {
                    if let std::collections::hash_map::Entry::Vacant(slot) =
                        pool.plain.entry(s.as_ref())
                    {
                        slot.insert(next);
                        pool.entries.push(Pooled::Plain(s.as_ref()));
                    }
                }
                CellValue::RichText(runs) => {
                    if let std::collections::hash_map::Entry::Vacant(slot) =
                        pool.rich.entry(rich_key(runs))
                    {
                        slot.insert(next);
                        pool.entries.push(Pooled::Rich(runs));
                    }
                }
                _ => {}
            }
        }
    }
    pool
}

const XML_DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// The package's own relationships: the workbook, and the document properties
/// carried over from the file that was read.
fn root_rels(book: &Spreadsheet) -> String {
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            r#"<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/>"#
        ),
        decl = XML_DECL,
        rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    );
    for (i, doc) in book.doc_props.iter().enumerate() {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{}" Target="{}"/>"#,
            i + 2,
            escape(&doc.kind),
            escape(&relative_target("", &doc.target))
        );
    }
    s.push_str("</Relationships>");
    s
}

#[expect(
    dead_code,
    reason = "kept as the shape of a package with nothing carried"
)]
const ROOT_RELS: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    r#"<Relationship Id="rId1" "#,
    r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" "#,
    r#"Target="xl/workbook.xml"/>"#,
    "</Relationships>",
);

/// The content type of the workbook part. A workbook that carries a VBA
/// project is macro-enabled, and says so here: Excel refuses to open an `.xlsm`
/// whose main part claims to be a plain workbook, macros and all. A template
/// is the same story for `.xltx` and `.xltm`.
fn main_content_type(book: &Spreadsheet) -> &'static str {
    let macros = book
        .parts
        .iter()
        .any(|p| p.content_type.as_deref() == Some("application/vnd.ms-office.vbaProject"));
    match (book.template, macros) {
        (false, false) => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        }
        (false, true) => "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        (true, false) => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
        }
        (true, true) => "application/vnd.ms-excel.template.macroEnabled.main+xml",
    }
}

fn content_types(book: &Spreadsheet) -> String {
    let sheet_count = book.sheets().len();
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
            r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#,
            r#"<Default Extension="xml" ContentType="application/xml"/>"#,
            r#"<Override PartName="/xl/workbook.xml" ContentType="{main}"/>"#,
            r#"<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#,
            r#"<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#,
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>"#,
        ),
        decl = XML_DECL,
        main = main_content_type(book),
    );
    for i in 1..=sheet_count {
        let _ = write!(
            s,
            r#"<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#
        );
    }
    // A sheet with notes has a part for them, written from the model.
    for (i, sheet) in book.sheets().iter().enumerate() {
        if !sheet.comments.is_empty() {
            let _ = write!(
                s,
                r#"<Override PartName="/xl/comments{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"#,
                i + 1
            );
        }
    }
    let mut table_number = 1;
    for sheet in book.sheets() {
        for _ in &sheet.tables {
            let _ = write!(
                s,
                r#"<Override PartName="/xl/tables/table{table_number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>"#,
            );
            table_number += 1;
        }
    }
    // A carried part keeps the type it had. An `<Override>` covers a part on
    // its own, so no `<Default>` per extension is needed; the ones for the
    // relationship parts are already declared above.
    for part in &book.parts {
        // A relationship part is covered by the `rels` default above.
        if std::path::Path::new(&part.path)
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("rels"))
        {
            continue;
        }
        if let Some(kind) = &part.content_type {
            let _ = write!(
                s,
                r#"<Override PartName="/{}" ContentType="{}"/>"#,
                escape(&part.path),
                escape(kind)
            );
        }
    }
    s.push_str("</Types>");
    s
}

/// The carried attributes of `<workbookView>`. `firstSheet` past the last
/// sheet, or past the active one, is a file Excel repairs, and a sheet removal
/// makes one; it is pulled back.
fn workbook_view_attrs(book: &Spreadsheet) -> String {
    let mut s = String::new();
    for (name, value) in &book.workbook_view {
        let value = if name == "firstSheet" {
            let first: usize = value.parse().unwrap_or(0);
            first.min(book.active_index()).to_string()
        } else {
            value.clone()
        };
        let _ = write!(s, r#" {name}="{}""#, escape(&value));
    }
    s
}

fn workbook(book: &Spreadsheet) -> String {
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            "{properties}",
            // `<workbookProtection>` precedes the views in the schema's order.
            "{protection}",
            "<bookViews><workbookView{view} activeTab=\"{active}\"/></bookViews>",
            "<sheets>"
        ),
        decl = XML_DECL,
        // `<fileVersion>` is deliberately absent: it names the application
        // that last wrote the file, and this is not that application.
        properties = workbook_pr(book),
        protection = workbook_protection_xml(&book.protection),
        view = workbook_view_attrs(book),
        active = book.active_index()
    );
    for (i, sheet) in book.sheets().iter().enumerate() {
        let (id, name) = (i + 1, escape(sheet.title()));
        let state = sheet
            .visibility
            .as_str()
            .map_or(String::new(), |v| format!(r#" state="{v}""#));
        let _ = write!(
            s,
            r#"<sheet name="{name}" sheetId="{id}"{state} r:id="rId{id}"/>"#
        );
    }
    s.push_str("</sheets>");
    // `<externalReferences>` sits between the sheets and the names, and each
    // entry is a pointer at the relationship holding the other workbook.
    let external: Vec<usize> = book
        .attachments
        .iter()
        .enumerate()
        .filter(|(_, a)| a.role() == "externalLink")
        .map(|(i, _)| i)
        .collect();
    if !external.is_empty() {
        s.push_str("<externalReferences>");
        for i in external {
            let _ = write!(
                s,
                r#"<externalReference r:id="rId{}"/>"#,
                book.sheets().len() + 4 + i
            );
        }
        s.push_str("</externalReferences>");
    }
    // `<definedNames>` follows `<sheets>` in the schema's fixed order.
    if !book.defined_names.is_empty() {
        s.push_str("<definedNames>");
        for name in &book.defined_names {
            let _ = write!(s, r#"<definedName name="{}""#, escape(&name.name));
            if let Some(sheet) = name.sheet {
                let _ = write!(s, r#" localSheetId="{sheet}""#);
            }
            if name.hidden {
                s.push_str(r#" hidden="1""#);
            }
            let _ = write!(s, ">{}</definedName>", escape(&name.formula));
        }
        s.push_str("</definedNames>");
    }
    // `<calcPr>` closes the workbook part, after the names.
    if !book.calculation_properties.is_empty() {
        s.push_str("<calcPr");
        for (name, value) in &book.calculation_properties {
            let _ = write!(s, r#" {}="{}""#, escape(name), escape(value));
        }
        s.push_str("/>");
    }
    // `<pivotCaches>` names the caches the reports read from. The parts
    // travel unparsed, but without this element the workbook does not point at
    // them and every pivot in it is broken.
    let caches: Vec<(u32, usize)> = book
        .pivot_caches
        .iter()
        .filter_map(|cache| {
            let index = book
                .attachments
                .iter()
                .position(|a| a.target == cache.definition_part)?;
            Some((cache.id, index))
        })
        .collect();
    if !caches.is_empty() {
        s.push_str("<pivotCaches>");
        for (id, index) in caches {
            let _ = write!(
                s,
                r#"<pivotCache cacheId="{id}" r:id="rId{}"/>"#,
                book.sheets().len() + 4 + index
            );
        }
        s.push_str("</pivotCaches>");
    }
    // Whatever the later schemas hung on the workbook, back where it was.
    if let Some(extensions) = &book.workbook_extensions {
        s.push_str(extensions);
    }
    s.push_str("</workbook>");
    s
}

/// Renders `<workbookPr>`, which is where the base date is declared.
fn workbook_pr(book: &Spreadsheet) -> String {
    let mac = book.epoch == crate::shared::date::Epoch::Mac1904;
    if !mac && book.workbook_properties.is_empty() {
        return String::new();
    }
    let mut s = String::from("<workbookPr");
    if mac {
        s.push_str(r#" date1904="1""#);
    }
    for (name, value) in &book.workbook_properties {
        let _ = write!(s, r#" {}="{}""#, escape(name), escape(value));
    }
    s.push_str("/>");
    s
}

/// Renders `<workbookProtection>`, or nothing when the workbook locks nothing.
fn workbook_protection_xml(protection: &crate::model::WorkbookProtection) -> String {
    if protection.is_empty() {
        return String::new();
    }
    let mut s = String::from("<workbookProtection");
    for (name, value) in protection.locks() {
        if let Some(on) = value {
            // Here the file spells the words out, unlike the sheet's `1`/`0`:
            // that is what Excel writes and what Excel reads back.
            let _ = write!(s, r#" {name}="{on}""#);
        }
    }
    for (attrs, password) in [
        (
            crate::model::protection::PasswordAttrs::WORKBOOK,
            &protection.workbook_password,
        ),
        (
            crate::model::protection::PasswordAttrs::REVISIONS,
            &protection.revisions_password,
        ),
    ] {
        if let Some(password) = password {
            for (name, value) in password.to_attrs(attrs) {
                let _ = write!(s, r#" {name}="{}""#, escape(&value));
            }
        }
    }
    s.push_str("/>");
    s
}

fn workbook_rels(book: &Spreadsheet) -> String {
    let sheet_count = book.sheets().len();
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
        ),
        decl = XML_DECL
    );
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for i in 1..=sheet_count {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{i}" Type="{rel}/worksheet" Target="worksheets/sheet{i}.xml"/>"#
        );
    }
    // Styles, the string pool and the theme follow the sheets, so sheet N is
    // always rIdN and the ids in workbook.xml need no lookup table.
    //
    // The theme must be declared here and not merely be present in the
    // package: a colour written as `theme="9"` is an index into it, and a
    // reader that cannot find the part has nothing to resolve the index
    // against.
    let _ = write!(
        s,
        concat!(
            r#"<Relationship Id="rId{styles}" Type="{rel}/styles" Target="styles.xml"/>"#,
            r#"<Relationship Id="rId{strings}" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/>"#,
            r#"<Relationship Id="rId{theme}" Type="{rel}/theme" Target="theme/theme1.xml"/>"#
        ),
        rel = rel,
        styles = sheet_count + 1,
        strings = sheet_count + 2,
        theme = sheet_count + 3
    );
    for (i, attached) in book.attachments.iter().enumerate() {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{}" Target="{}"/>"#,
            sheet_count + 4 + i,
            escape(&attached.kind),
            escape(&relative_target("xl", &attached.target))
        );
    }
    s.push_str("</Relationships>");
    s
}

/// A package path written the way a relationship inside `base` refers to it.
///
/// `/xl/pivotTables/pivotTable1.xml` from `xl/worksheets` is
/// `../pivotTables/pivotTable1.xml`. An absolute target is legal in the
/// package format and Excel reads it, but plenty of readers resolve targets by
/// hand and only handle the relative form - excelize walks off a null pointer
/// on one - so what goes out is what Excel itself writes.
pub(super) fn relative_target(base: &str, target: &str) -> String {
    let target = target.trim_start_matches('/');
    let base: Vec<&str> = base.split('/').filter(|p| !p.is_empty()).collect();
    let parts: Vec<&str> = target.split('/').collect();
    let shared = base.iter().zip(&parts).take_while(|(a, b)| a == b).count();
    let mut out = String::new();
    for _ in shared..base.len() {
        out.push_str("../");
    }
    out.push_str(&parts[shared..].join("/"));
    out
}

/// The default Office theme.
///
/// A style may refer to a colour as `theme="4"` rather than by rgb, and the
/// reference only resolves against this part. Writing styles that point at a
/// theme without shipping one produces a file Excel refuses to open, so the
/// standard palette goes out with every workbook.
const THEME: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">"#,
    "<a:themeElements><a:clrScheme name=\"Office\">",
    r#"<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>"#,
    r#"<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>"#,
    r#"<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>"#,
    r#"<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>"#,
    r#"<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>"#,
    r#"<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>"#,
    r#"<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>"#,
    "</a:clrScheme>",
    r#"<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/>"#,
    "<a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>",
    r#"<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>"#,
    "</a:fontScheme>",
    r#"<a:fmtScheme name="Office">"#,
    r#"<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
    r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
    r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>"#,
    r#"<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>"#,
    r#"<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>"#,
    r#"<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>"#,
    r#"<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>"#,
    r#"<a:effectStyle><a:effectLst/></a:effectStyle>"#,
    r#"<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>"#,
    r#"<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
    r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>"#,
    r#"<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>"#,
    "</a:fmtScheme></a:themeElements></a:theme>",
);

fn shared_strings(pool: &StringPool<'_>) -> String {
    let count = pool.entries.len();
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"count="{count}" uniqueCount="{count}">"#
        ),
        decl = XML_DECL,
        count = count
    );
    for entry in &pool.entries {
        // xml:space="preserve" keeps leading and trailing whitespace, which a
        // reader is otherwise free to trim.
        match entry {
            Pooled::Plain(text) => {
                let _ = write!(
                    s,
                    r#"<si><t xml:space="preserve">{}</t></si>"#,
                    escape(text)
                );
            }
            Pooled::Rich(runs) => {
                s.push_str("<si>");
                for run in *runs {
                    s.push_str("<r>");
                    if let Some(font) = &run.font {
                        s.push_str(&run_font_xml(font));
                    }
                    let _ = write!(
                        s,
                        r#"<t xml:space="preserve">{}</t></r>"#,
                        escape(&run.text)
                    );
                }
                s.push_str("</si>");
            }
        }
    }
    s.push_str("</sst>");
    s
}

/// Collects distinct values in first-seen order, returning the table and the
/// index each style maps to.
///
/// xlsx keeps fonts, fills and borders in tables of their own, with `cellXfs`
/// pointing into them, so whole styles have to be taken apart on the way out
/// the same way they were assembled on the way in.
fn dedupe<T: Clone + PartialEq>(
    values: impl Iterator<Item = T>,
    seed: Vec<T>,
) -> (Vec<T>, Vec<u32>) {
    let mut table = seed;
    let mut indices = Vec::new();
    for v in values {
        let pos = table.iter().position(|t| *t == v).unwrap_or_else(|| {
            table.push(v);
            table.len() - 1
        });
        indices.push(u32::try_from(pos).unwrap_or(0));
    }
    (table, indices)
}

fn styles(book: &Spreadsheet) -> String {
    let all = book.styles.all();
    let (custom, format_ids) = custom_format_ids(all);

    let (fonts, font_ids) = dedupe(all.iter().map(|s| s.font.clone()), Vec::new());
    // Excel reserves the first two fills for `none` and `gray125` and mis-renders
    // a file that omits them, so the fill table starts seeded with both.
    let (fills, fill_ids) = dedupe(
        all.iter().map(|s| s.fill.clone()),
        vec![
            Fill::default(),
            Fill {
                pattern: Pattern::Named("gray125"),
                ..Fill::default()
            },
        ],
    );
    let (borders, border_ids) = dedupe(all.iter().map(|s| s.borders.clone()), Vec::new());

    let mut s = format!(
        concat!(
            "{decl}",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#
        ),
        decl = XML_DECL
    );

    if !custom.is_empty() {
        let _ = write!(s, r#"<numFmts count="{}">"#, custom.len());
        for (i, code) in custom.iter().enumerate() {
            let id = FIRST_CUSTOM_FORMAT_ID + u16::try_from(i).unwrap_or(0);
            let _ = write!(
                s,
                r#"<numFmt numFmtId="{id}" formatCode="{}"/>"#,
                escape(code)
            );
        }
        s.push_str("</numFmts>");
    }

    // The schema fixes this order. Excel wants a font and a border at index 0
    // even where nothing uses one; the fill table is seeded already.
    component_table(&mut s, "fonts", &fonts, &Font::default(), font_xml);
    let _ = write!(s, r#"<fills count="{}">"#, fills.len());
    for fill in &fills {
        s.push_str(&fill_xml(fill));
    }
    s.push_str("</fills>");
    component_table(
        &mut s,
        "borders",
        &borders,
        &Borders::default(),
        borders_xml,
    );

    s.push_str(
        r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
    );

    let _ = write!(s, r#"<cellXfs count="{}">"#, all.len());
    for (i, style) in all.iter().enumerate() {
        s.push_str(&cell_xf_xml(
            style,
            [
                u32::from(format_ids[i]),
                font_ids[i],
                fill_ids[i],
                border_ids[i],
            ],
        ));
    }
    s.push_str("</cellXfs>");

    s.push_str(
        r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
    );
    // `<dxfs>` follows `<cellStyles>` in the schema's fixed order; conditional
    // formatting rules index into it.
    let dxfs = &book.styles.differential;
    if !dxfs.is_empty() {
        let _ = write!(s, r#"<dxfs count="{}">"#, dxfs.len());
        for dxf in dxfs {
            s.push_str(&dxf_xml(dxf));
        }
        s.push_str("</dxfs>");
    }
    // The schema's order: table styles, then the palette, then whatever a
    // later schema hung on the stylesheet.
    if let Some(styles) = &book.table_styles {
        s.push_str(styles);
    }
    if let Some(palette) = &book.palette {
        s.push_str(palette);
    }
    if let Some(extensions) = &book.style_extensions {
        s.push_str(extensions);
    }
    s.push_str("</styleSheet>");
    s
}

/// Number format ids for every style, plus the custom format strings they refer
/// to. A custom string gets an id of its own above the built-in range.
fn custom_format_ids(all: &[crate::style::Style]) -> (Vec<&str>, Vec<u16>) {
    // A `Vec`, not a `HashMap`: a workbook usually holds about ten custom
    // formats, and a linear scan over those beats hashing. A hundred would show,
    // but no one has brought such a workbook yet.
    let mut custom: Vec<&str> = Vec::new();
    let ids = all
        .iter()
        .map(|s| match &s.number_format {
            NumberFormat::Custom(code) => {
                let pos = custom.iter().position(|c| c == code).unwrap_or_else(|| {
                    custom.push(code);
                    custom.len() - 1
                });
                FIRST_CUSTOM_FORMAT_ID + u16::try_from(pos).unwrap_or(0)
            }
            NumberFormat::General => 0, // 0 == General
            NumberFormat::Builtin(id) => *id,
        })
        .collect();
    (custom, ids)
}

/// One `<fonts>`- or `<borders>`-style table, which must hold at least the one
/// entry every `<xf>` falls back to.
// The fallback entry is not optional: Excel opens a file with <fonts count="0"/>
// but then has nothing to draw the cells with.
fn component_table<T>(
    out: &mut String,
    tag: &str,
    items: &[T],
    fallback: &T,
    render: fn(&T) -> String,
) {
    let _ = write!(out, r#"<{tag} count="{}">"#, items.len().max(1));
    if items.is_empty() {
        out.push_str(&render(fallback));
    }
    for item in items {
        out.push_str(&render(item));
    }
    let _ = write!(out, "</{tag}>");
}

/// One `<xf>`: a style as the indices it is assembled from, with an `apply*`
/// flag for each part that is not the default.
fn cell_xf_xml(style: &crate::style::Style, [fmt, font, fill, border]: [u32; 4]) -> String {
    let al = alignment_xml(&style.alignment);
    let pr = protection_xml(style.protection);
    let mut out = String::new();
    let _ = write!(
        out,
        concat!(
            r#"<xf xfId="0" numFmtId="{fmt}" fontId="{font}" fillId="{fill}" borderId="{border}""#,
            r#" applyNumberFormat="{af}" applyFont="{afont}" applyFill="{afill}""#,
            r#" applyBorder="{ab}" applyAlignment="{aa}" applyProtection="{ap}""#
        ),
        fmt = fmt,
        font = font,
        fill = fill,
        border = border,
        af = u8::from(fmt != 0),
        afont = u8::from(font != 0),
        afill = u8::from(fill != 0),
        ab = u8::from(border != 0),
        aa = u8::from(!al.is_empty()),
        ap = u8::from(!pr.is_empty()),
    );
    // An empty <xf/> collapses; otherwise the diff against the source is noise.
    if al.is_empty() && pr.is_empty() {
        out.push_str("/>");
    } else {
        let _ = write!(out, ">{al}{pr}</xf>");
    }
    out
}

/// Renders a colour attribute, or nothing for an automatic colour.
fn color_attr(color: &Color) -> String {
    match color {
        Color::Auto => String::new(),
        Color::Argb(_) => color
            .to_argb_str()
            .map_or_else(String::new, |v| format!(r#" rgb="{v}""#)),
        Color::Indexed(i) => format!(r#" indexed="{i}""#),
        Color::Theme { id, tint } => {
            if *tint == 0 {
                format!(r#" theme="{id}""#)
            } else {
                format!(r#" theme="{id}" tint="{}""#, f64::from(*tint) / 1_000_000.0)
            }
        }
    }
}

fn font_xml(font: &Font) -> String {
    let mut s = String::from("<font>");
    if font.bold {
        s.push_str(r#"<b val="1"/>"#);
    }
    if font.italic {
        s.push_str(r#"<i val="1"/>"#);
    }
    if font.strike {
        s.push_str(r#"<strike val="1"/>"#);
    }
    if font.underline != Underline::None {
        let _ = write!(s, r#"<u val="{}"/>"#, font.underline.as_str());
    }
    match font.script {
        Script::Superscript => s.push_str(r#"<vertAlign val="superscript"/>"#),
        Script::Subscript => s.push_str(r#"<vertAlign val="subscript"/>"#),
        Script::Baseline => {}
    }
    let _ = write!(s, r#"<sz val="{}"/>"#, trim_number(font.size_points()));
    let color = color_attr(&font.color);
    if !color.is_empty() {
        let _ = write!(s, "<color{color}/>");
    }
    let _ = write!(s, r#"<name val="{}"/>"#, escape(&font.name));
    // The character set and the scheme come after the name, which is the
    // order the schema fixes for them.
    if let Some(family) = font.family {
        let _ = write!(s, r#"<family val="{family}"/>"#);
    }
    if let Some(charset) = font.charset {
        let _ = write!(s, r#"<charset val="{charset}"/>"#);
    }
    if let Some(scheme) = font.scheme {
        let _ = write!(s, r#"<scheme val="{}"/>"#, scheme.as_str());
    }
    s.push_str("</font>");
    s
}

fn fill_xml(fill: &Fill) -> String {
    if fill.pattern == Pattern::None {
        return r#"<fill><patternFill patternType="none"/></fill>"#.to_owned();
    }
    let mut s = format!(
        r#"<fill><patternFill patternType="{}">"#,
        fill.pattern.as_str()
    );
    let fg = color_attr(&fill.foreground);
    if !fg.is_empty() {
        let _ = write!(s, "<fgColor{fg}/>");
    }
    let bg = color_attr(&fill.background);
    if !bg.is_empty() {
        let _ = write!(s, "<bgColor{bg}/>");
    }
    s.push_str("</patternFill></fill>");
    s
}

fn borders_xml(b: &Borders) -> String {
    let mut s = String::from("<border");
    match b.diagonal_direction {
        DiagonalDirection::Up => s.push_str(r#" diagonalUp="1""#),
        DiagonalDirection::Down => s.push_str(r#" diagonalDown="1""#),
        DiagonalDirection::Both => s.push_str(r#" diagonalUp="1" diagonalDown="1""#),
        DiagonalDirection::None => {}
    }
    s.push('>');
    for (name, side) in [
        ("left", &b.left),
        ("right", &b.right),
        ("top", &b.top),
        ("bottom", &b.bottom),
        ("diagonal", &b.diagonal),
    ] {
        if side.style == BorderStyle::None {
            let _ = write!(s, "<{name}/>");
            continue;
        }
        let _ = write!(s, r#"<{name} style="{}">"#, side.style.as_str());
        let color = color_attr(&side.color);
        if !color.is_empty() {
            let _ = write!(s, "<color{color}/>");
        }
        let _ = write!(s, "</{name}>");
    }
    s.push_str("</border>");
    s
}

fn alignment_xml(a: &Alignment) -> String {
    if *a == Alignment::default() {
        return String::new();
    }
    let mut s = String::from("<alignment");
    if let Some(h) = a.horizontal.as_str() {
        let _ = write!(s, r#" horizontal="{h}""#);
    }
    if let Some(v) = a.vertical.as_str() {
        let _ = write!(s, r#" vertical="{v}""#);
    }
    if a.text_rotation != 0 {
        let _ = write!(s, r#" textRotation="{}""#, a.text_rotation);
    }
    if a.wrap_text {
        s.push_str(r#" wrapText="1""#);
    }
    if a.shrink_to_fit {
        s.push_str(r#" shrinkToFit="1""#);
    }
    if a.indent != 0 {
        let _ = write!(s, r#" indent="{}""#, a.indent);
    }
    if a.reading_order != 0 {
        let _ = write!(s, r#" readingOrder="{}""#, a.reading_order);
    }
    s.push_str("/>");
    s
}

fn protection_xml(p: Protection) -> String {
    if p == Protection::default() {
        return String::new();
    }
    let mut s = String::from("<protection");
    if let Some(v) = p.locked.as_bool() {
        let _ = write!(s, r#" locked="{v}""#);
    }
    if let Some(v) = p.hidden.as_bool() {
        let _ = write!(s, r#" hidden="{v}""#);
    }
    s.push_str("/>");
    s
}

/// `<sheetFormatPr>`, or nothing when the sheet has no default sizes and no
/// outline.
fn sheet_format_xml(sheet: &crate::model::Worksheet) -> String {
    // Excel draws the outline's level buttons from these two counts, not
    // from the rows and columns themselves: without them a grouped sheet
    // opens with no buttons to collapse it by.
    let outline_rows = sheet
        .rows
        .values()
        .map(|r| r.outline_level)
        .max()
        .unwrap_or(0);
    let outline_cols = sheet
        .columns
        .iter()
        .map(|c| c.outline_level)
        .max()
        .unwrap_or(0);
    let mut s = String::new();
    if sheet.default_column_width.is_some()
        || sheet.default_row_height.is_some()
        || outline_rows > 0
        || outline_cols > 0
    {
        s.push_str("<sheetFormatPr");
        if let Some(w) = sheet.default_column_width {
            let _ = write!(s, r#" defaultColWidth="{}""#, trim_number(w));
        }
        // Required by the schema; 15 points is Excel's own for Calibri 11.
        let _ = write!(
            s,
            r#" defaultRowHeight="{}""#,
            trim_number(sheet.default_row_height.unwrap_or(15.0))
        );
        if outline_rows > 0 {
            let _ = write!(s, r#" outlineLevelRow="{outline_rows}""#);
        }
        if outline_cols > 0 {
            let _ = write!(s, r#" outlineLevelCol="{outline_cols}""#);
        }
        s.push_str("/>");
    }
    s
}

/// Renders a number without a trailing `.0`, which Excel dislikes on font sizes.
fn trim_number(n: f64) -> String {
    if n.fract() == 0.0 {
        format!("{n:.0}")
    } else {
        n.to_string()
    }
}

fn worksheet(
    sheet: &crate::model::Worksheet,
    pool: &StringPool<'_>,
    outside: &[&crate::model::Hyperlink],
) -> String {
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#
        ),
        decl = XML_DECL
    );
    // `<sheetPr>` comes before everything else in the schema's fixed order.
    s.push_str(&sheet_pr_xml(&sheet.properties));
    if let Some(dim) = sheet.dimension() {
        let _ = write!(s, r#"<dimension ref="{dim}"/>"#);
    }

    s.push_str(&sheet_views_xml(&sheet.view));

    // `<sheetFormatPr>` and `<cols>` must precede `<sheetData>`; Excel rejects
    // a sheet whose elements are out of order.
    s.push_str(&sheet_format_xml(sheet));
    if !sheet.columns.is_empty() {
        s.push_str("<cols>");
        for run in &sheet.columns {
            s.push_str(&column_xml(run));
        }
        s.push_str("</cols>");
    }

    s.push_str("<sheetData>");

    // Cells arrive row by row already, so a row closes as soon as the row index
    // changes; xlsx requires both rows and cells in ascending order. A row can
    // carry a height or be hidden without holding any cell: both sequences are
    // sorted, so those are merged in on the way rather than searched for.
    let arrays = ArrayFormulas::of(&sheet.array_formulas);
    let mut bare_rows = sheet.rows.iter().peekable();
    let mut open_row = None;
    for (at, cell) in sheet.iter() {
        if open_row != Some(at.row) {
            if open_row.is_some() {
                s.push_str("</row>");
            }
            while let Some((row, props)) = bare_rows.next_if(|(row, _)| **row < at.row) {
                let _ = write!(s, "{}</row>", row_open_tag(*row, Some(props)));
            }
            let props = bare_rows
                .next_if(|(row, _)| **row == at.row)
                .map(|(_, p)| p);
            s.push_str(&row_open_tag(at.row, props));
            open_row = Some(at.row);
        }
        s.push_str(&cell_xml(at, cell, pool, &arrays));
    }
    if open_row.is_some() {
        s.push_str("</row>");
    }
    for (row, props) in bare_rows {
        let _ = write!(s, "{}</row>", row_open_tag(*row, Some(props)));
    }
    s.push_str("</sheetData>");

    // `<sheetProtection>`, `<protectedRanges>` and `<autoFilter>` all sit
    // between the cells and the merges in the schema's fixed order.
    s.push_str(&guards_xml(sheet));

    if !sheet.merges.is_empty() {
        let _ = write!(s, r#"<mergeCells count="{}">"#, sheet.merges.len());
        for m in &sheet.merges {
            let _ = write!(s, r#"<mergeCell ref="{m}"/>"#);
        }
        s.push_str("</mergeCells>");
    }

    // `<conditionalFormatting>` sits between them in the schema's fixed order.
    for block in &sheet.conditional_formats {
        let _ = write!(
            s,
            r#"<conditionalFormatting sqref="{}">"#,
            sqref(&block.sqref)
        );
        for rule in &block.rules {
            s.push_str(&cf_rule_xml(rule));
        }
        s.push_str("</conditionalFormatting>");
    }

    // `<dataValidations>` follows `<mergeCells>` in the schema's fixed order.
    if !sheet.data_validations.is_empty() {
        let _ = write!(
            s,
            r#"<dataValidations count="{}">"#,
            sheet.data_validations.len()
        );
        for dv in &sheet.data_validations {
            s.push_str(&data_validation_xml(dv));
        }
        s.push_str("</dataValidations>");
    }

    s.push_str(&print_tail_xml(sheet, outside));
    s.push_str(&attached_parts_xml(sheet));
    // `<extLst>` closes the element, and what is in it came from the file
    // unread, except for the sparklines.
    if let Some(extensions) = sheet_extensions(sheet) {
        s.push_str(&extensions);
    }
    s.push_str("</worksheet>");
    s
}

/// The sheet's `<extLst>`: as it was read, unless the sparkline groups in the
/// model no longer say what it says. Then only their `<ext>` is replaced, and
/// the extensions beside it stay byte for byte.
fn sheet_extensions(sheet: &crate::model::Worksheet) -> Option<std::borrow::Cow<'_, str>> {
    use crate::model::sparkline::{EXTENSION_URI, read};
    use std::borrow::Cow;
    let carried = sheet.extensions.as_deref();
    if carried.map(read).unwrap_or_default() == sheet.sparklines {
        return carried.map(Cow::Borrowed);
    }
    let ours = if sheet.sparklines.is_empty() {
        String::new()
    } else {
        sparklines_ext_xml(&sheet.sparklines)
    };
    let Some(carried) = carried else {
        return Some(Cow::Owned(format!("<extLst>{ours}</extLst>")));
    };
    // The carried `<ext>` holding them, if there is one.
    let old = carried
        .match_indices("<ext ")
        .map(|(at, _)| at)
        .find(|&at| {
            let tag_end = carried[at..].find('>').map_or(carried.len(), |e| at + e);
            carried[at..tag_end].contains(EXTENSION_URI)
        })
        .and_then(|at| Some((at, at + carried[at..].find("</ext>")? + "</ext>".len())));
    // Without one, they go in last.
    let (from, to) = old.unwrap_or_else(|| {
        let at = carried.rfind("</extLst>").unwrap_or(carried.len());
        (at, at)
    });
    let spliced = format!("{}{ours}{}", &carried[..from], &carried[to..]);
    // An `<extLst>` left with no `<ext>` in it says nothing.
    spliced.contains("<ext ").then_some(Cow::Owned(spliced))
}

/// The `<ext>` holding a sheet's sparkline groups.
fn sparklines_ext_xml(groups: &[crate::model::sparkline::SparklineGroup]) -> String {
    let mut s = format!(
        concat!(
            r#"<ext uri="{uri}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#,
            r#"<x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#,
        ),
        uri = crate::model::sparkline::EXTENSION_URI,
    );
    for group in groups {
        s.push_str("<x14:sparklineGroup");
        if let Some(kind) = group.kind.as_str() {
            let _ = write!(s, r#" type="{kind}""#);
        }
        for (name, value) in &group.attributes {
            let _ = write!(s, r#" {name}="{}""#, escape(value));
        }
        for (name, on) in group.flags() {
            if on {
                let _ = write!(s, r#" {name}="1""#);
            }
        }
        s.push('>');
        // The schema's order: colours, the date axis, then the sparklines.
        s.push_str(&group.colors);
        if let Some(axis) = &group.date_axis {
            let _ = write!(s, "<xm:f>{}</xm:f>", escape(axis));
        }
        s.push_str("<x14:sparklines>");
        for line in &group.sparklines {
            s.push_str("<x14:sparkline>");
            if let Some(data) = &line.data {
                let _ = write!(s, "<xm:f>{}</xm:f>", escape(data));
            }
            let _ = write!(s, "<xm:sqref>{}</xm:sqref></x14:sparkline>", line.location);
        }
        s.push_str("</x14:sparklines></x14:sparklineGroup>");
    }
    s.push_str("</x14:sparklineGroups></ext>");
    s
}

/// Renders one `xl/tables/tableN.xml`.
///
/// Written from the model rather than carried, so that a range moved by an
/// edit takes its table with it.
fn table_xml(table: &crate::model::table::Table) -> String {
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main""#,
            r#" id="{id}" name="{name}" displayName="{display}" ref="{range}""#,
        ),
        decl = XML_DECL,
        id = table.id,
        name = escape(&table.name),
        display = escape(&table.display_name),
        range = table.range,
    );
    // Both counts default in the schema, so an absent one stays absent: the
    // file said nothing and neither do we.
    if let Some(n) = table.header_row_count {
        let _ = write!(s, r#" headerRowCount="{n}""#);
    }
    if let Some(n) = table.totals_row_count {
        let _ = write!(s, r#" totalsRowCount="{n}""#);
    }
    s.push('>');
    if let Some(filter) = &table.auto_filter {
        s.push_str(&auto_filter_xml(filter));
    }
    if let Some(sort) = &table.sort_state {
        s.push_str(&sort_state_xml(sort));
    }
    let _ = write!(s, r#"<tableColumns count="{}">"#, table.columns.len());
    for column in &table.columns {
        let _ = write!(
            s,
            r#"<tableColumn id="{}" name="{}""#,
            column.id,
            escape(&column.name)
        );
        if let Some(f) = &column.totals_row_function {
            let _ = write!(s, r#" totalsRowFunction="{}""#, escape(f));
        }
        if let Some(l) = &column.totals_row_label {
            let _ = write!(s, r#" totalsRowLabel="{}""#, escape(l));
        }
        match &column.calculated_formula {
            Some(f) => {
                let _ = write!(
                    s,
                    "><calculatedColumnFormula>{}</calculatedColumnFormula></tableColumn>",
                    escape(f)
                );
            }
            None => s.push_str("/>"),
        }
    }
    s.push_str("</tableColumns>");
    if let Some(style) = &table.style {
        s.push_str("<tableStyleInfo");
        if let Some(name) = &style.name {
            let _ = write!(s, r#" name="{}""#, escape(name));
        }
        let _ = write!(
            s,
            concat!(
                r#" showFirstColumn="{first}" showLastColumn="{last}""#,
                r#" showRowStripes="{rows}" showColumnStripes="{columns}"/>"#,
            ),
            first = u8::from(style.show_first_column),
            last = u8::from(style.show_last_column),
            rows = u8::from(style.show_row_stripes),
            columns = u8::from(style.show_column_stripes),
        );
    }
    if let Some(extensions) = &table.extensions {
        s.push_str(extensions);
    }
    s.push_str("</table>");
    s
}

/// Renders the elements that point at a sheet's own parts: its drawings, the
/// shapes behind its comments, and its tables.
///
/// Each is a pointer at a relationship rather than content, and the schema
/// fixes their order at the end of the sheet.
fn attached_parts_xml(sheet: &crate::model::Worksheet) -> String {
    // The external hyperlinks take the first relationship ids, so a sheet's
    // own parts are numbered after them.
    let first = external_links(sheet).len() + 1;
    let mut s = String::new();
    for (i, attached) in sheet.attachments.iter().enumerate() {
        let tag = match attached.role() {
            "drawing" => "drawing",
            "vmlDrawing" => "legacyDrawing",
            // A comment part is found through the relationship alone; a
            // printer setting hangs off `<pageSetup>`, which is not written
            // with one here.
            _ => continue,
        };
        let _ = write!(s, r#"<{tag} r:id="rId{}"/>"#, first + i);
    }
    // `<tableParts>` is last but for `<extLst>`. Without the element the part
    // is still in the package and still related, and Excel shows the cells as
    // an ordinary range - the table is gone.
    if !sheet.tables.is_empty() {
        let start = table_rel_base(sheet);
        let _ = write!(s, r#"<tableParts count="{}">"#, sheet.tables.len());
        for i in 0..sheet.tables.len() {
            let _ = write!(s, r#"<tablePart r:id="rId{}"/>"#, start + i);
        }
        s.push_str("</tableParts>");
    }
    s
}

/// The relationship id the sheet's first table takes. The tables are related
/// last, after the hyperlinks, the carried parts and the notes.
fn table_rel_base(sheet: &crate::model::Worksheet) -> usize {
    external_links(sheet).len()
        + 1
        + sheet.attachments.len()
        + usize::from(!sheet.comments.is_empty())
}

/// Renders the tail of a sheet part: the hyperlinks and everything about
/// printing, in the order the schema fixes them.
fn print_tail_xml(sheet: &crate::model::Worksheet, outside: &[&crate::model::Hyperlink]) -> String {
    let mut s = String::new();
    if !sheet.hyperlinks.is_empty() {
        s.push_str("<hyperlinks>");
        // `outside` lists the external links in this same order, so the n-th
        // one met here holds the n-th relationship id.
        let mut next_id = 0;
        for link in &sheet.hyperlinks {
            let id = match link.target {
                crate::model::LinkTarget::Outside(_) if next_id < outside.len() => {
                    next_id += 1;
                    Some(next_id)
                }
                _ => None,
            };
            s.push_str(&hyperlink_xml(link, id));
        }
        s.push_str("</hyperlinks>");
    }
    let print = sheet.print_options;
    if print != crate::model::PrintOptions::default() {
        s.push_str("<printOptions");
        for (name, on) in [
            ("horizontalCentered", print.horizontal_centered),
            ("verticalCentered", print.vertical_centered),
            ("headings", print.headings),
            ("gridLines", print.grid_lines),
        ] {
            if on {
                let _ = write!(s, r#" {name}="1""#);
            }
        }
        s.push_str("/>");
    }
    let m = sheet.margins;
    let _ = write!(
        s,
        r#"<pageMargins left="{}" right="{}" top="{}" bottom="{}" header="{}" footer="{}"/>"#,
        trim_number(m.left),
        trim_number(m.right),
        trim_number(m.top),
        trim_number(m.bottom),
        trim_number(m.header),
        trim_number(m.footer)
    );
    s.push_str(&page_setup_xml(&sheet.page_setup, sheet, outside.len()));
    s.push_str(&header_footer_xml(&sheet.header_footer));
    for (tag, breaks) in [
        ("rowBreaks", &sheet.row_breaks),
        ("colBreaks", &sheet.col_breaks),
    ] {
        if breaks.is_empty() {
            continue;
        }
        let manual = breaks.iter().filter(|b| b.manual).count();
        let _ = write!(
            s,
            r#"<{tag} count="{}" manualBreakCount="{manual}">"#,
            breaks.len()
        );
        for b in breaks {
            let _ = write!(s, r#"<brk id="{}""#, b.at);
            if let Some(max) = b.max {
                let _ = write!(s, r#" max="{max}""#);
            }
            if b.manual {
                s.push_str(r#" man="1""#);
            }
            s.push_str("/>");
        }
        let _ = write!(s, "</{tag}>");
    }
    s
}

/// Renders one `<dxf>`: only the parts a differential format overrides.
fn dxf_xml(dxf: &crate::style::DifferentialStyle) -> String {
    let mut s = String::from("<dxf>");
    if let Some(font) = &dxf.font {
        s.push_str("<font>");
        for (tag, on) in [
            ("b", font.bold),
            ("i", font.italic),
            ("strike", font.strike),
        ] {
            if let Some(on) = on {
                let _ = write!(s, r#"<{tag} val="{}"/>"#, u8::from(on));
            }
        }
        if let Some(u) = font.underline {
            let _ = write!(s, r#"<u val="{}"/>"#, u.as_str());
        }
        if let Some(size) = font.size {
            let _ = write!(s, r#"<sz val="{}"/>"#, trim_number(f64::from(size) / 100.0));
        }
        if let Some(color) = &font.color {
            let _ = write!(s, "<color{}/>", color_attr(color));
        }
        if let Some(name) = &font.name {
            let _ = write!(s, r#"<name val="{}"/>"#, escape(name));
        }
        s.push_str("</font>");
    }
    if let Some(format) = &dxf.number_format {
        match format {
            crate::style::NumberFormat::Custom(code) => {
                let _ = write!(
                    s,
                    r#"<numFmt numFmtId="{FIRST_CUSTOM_FORMAT_ID}" formatCode="{}"/>"#,
                    escape(code)
                );
            }
            crate::style::NumberFormat::Builtin(id) => {
                let _ = write!(s, r#"<numFmt numFmtId="{id}"/>"#);
            }
            crate::style::NumberFormat::General => {}
        }
    }
    if let Some(fill) = &dxf.fill {
        s.push_str("<fill><patternFill");
        if let Some(pattern) = &fill.pattern {
            let _ = write!(s, r#" patternType="{}""#, pattern.as_str());
        }
        s.push('>');
        if let Some(color) = &fill.foreground {
            let _ = write!(s, "<fgColor{}/>", color_attr(color));
        }
        if let Some(color) = &fill.background {
            let _ = write!(s, "<bgColor{}/>", color_attr(color));
        }
        s.push_str("</patternFill></fill>");
    }
    if let Some(borders) = &dxf.borders {
        s.push_str("<border>");
        for (tag, side) in [
            ("left", &borders.left),
            ("right", &borders.right),
            ("top", &borders.top),
            ("bottom", &borders.bottom),
            ("diagonal", &borders.diagonal),
        ] {
            let Some(border) = side else { continue };
            let _ = write!(s, r#"<{tag} style="{}">"#, border.style.as_str());
            let color = color_attr(&border.color);
            if !color.is_empty() {
                let _ = write!(s, "<color{color}/>");
            }
            let _ = write!(s, "</{tag}>");
        }
        s.push_str("</border>");
    }
    if let Some(alignment) = &dxf.alignment {
        s.push_str(&alignment_xml(alignment));
    }
    if let Some(protection) = dxf.protection {
        s.push_str(&protection_xml(protection));
    }
    s.push_str("</dxf>");
    s
}

/// Renders one `<cfRule>`.
fn cf_rule_xml(rule: &crate::model::CfRule) -> String {
    let mut s = format!(r#"<cfRule type="{}""#, rule.kind.as_str());
    if let Some(dxf) = rule.dxf {
        let _ = write!(s, r#" dxfId="{dxf}""#);
    }
    let _ = write!(s, r#" priority="{}""#, rule.priority);
    if let Some(op) = rule.operator {
        let _ = write!(s, r#" operator="{}""#, op.as_str());
    }
    if let Some(text) = &rule.text {
        let _ = write!(s, r#" text="{}""#, escape(text));
    }
    if let Some(period) = &rule.time_period {
        let _ = write!(s, r#" timePeriod="{}""#, escape(period));
    }
    if let Some(rank) = rule.rank {
        let _ = write!(s, r#" rank="{rank}""#);
    }
    if let Some(dev) = rule.std_dev {
        let _ = write!(s, r#" stdDev="{dev}""#);
    }
    for (name, on) in [
        ("percent", rule.percent),
        ("bottom", rule.bottom),
        ("equalAverage", rule.equal_average),
        ("stopIfTrue", rule.stop_if_true),
    ] {
        if on {
            let _ = write!(s, r#" {name}="1""#);
        }
    }
    // Written only when false: an absent attribute means "above".
    if !rule.above_average {
        s.push_str(r#" aboveAverage="0""#);
    }
    if rule.formulas.is_empty() && rule.scale.is_none() && rule.extensions.is_none() {
        s.push_str("/>");
        return s;
    }
    s.push('>');
    for formula in &rule.formulas {
        let _ = write!(s, "<formula>{}</formula>", escape(formula));
    }
    if let Some(scale) = &rule.scale {
        s.push_str(&cf_scale_xml(scale));
    }
    if let Some(extensions) = &rule.extensions {
        s.push_str(extensions);
    }
    s.push_str("</cfRule>");
    s
}

/// Renders the graphical part of a rule.
fn cf_scale_xml(scale: &crate::model::CfScale) -> String {
    use crate::model::CfScale;
    let stops = |values: &[crate::model::CfValue], gte: bool| {
        let mut out = String::new();
        for v in values {
            let _ = write!(out, r#"<cfvo type="{}""#, v.kind.as_str());
            if !v.value.is_empty() {
                let _ = write!(out, r#" val="{}""#, escape(&v.value));
            }
            // `gte` is written only when false, and only where it means
            // something.
            if gte && !v.greater_or_equal {
                out.push_str(r#" gte="0""#);
            }
            out.push_str("/>");
        }
        out
    };
    match scale {
        CfScale::Color { values, colors } => {
            let mut s = String::from("<colorScale>");
            s.push_str(&stops(values, false));
            for color in colors {
                let _ = write!(s, "<color{}/>", color_attr(color));
            }
            s.push_str("</colorScale>");
            s
        }
        CfScale::DataBar {
            values,
            color,
            min_length,
            max_length,
            show_value,
        } => {
            let mut s = String::from("<dataBar");
            if let Some(min) = min_length {
                let _ = write!(s, r#" minLength="{min}""#);
            }
            if let Some(max) = max_length {
                let _ = write!(s, r#" maxLength="{max}""#);
            }
            if !show_value {
                s.push_str(r#" showValue="0""#);
            }
            s.push('>');
            s.push_str(&stops(values, false));
            let _ = write!(s, "<color{}/></dataBar>", color_attr(color));
            s
        }
        CfScale::IconSet {
            values,
            set,
            show_value,
            percent,
            reverse,
        } => {
            let mut s = String::from("<iconSet");
            if let Some(set) = set {
                let _ = write!(s, r#" iconSet="{}""#, escape(set));
            }
            if !*show_value {
                s.push_str(r#" showValue="0""#);
            }
            if !*percent {
                s.push_str(r#" percent="0""#);
            }
            if *reverse {
                s.push_str(r#" reverse="1""#);
            }
            s.push('>');
            s.push_str(&stops(values, true));
            s.push_str("</iconSet>");
            s
        }
    }
}

/// The hyperlinks of a sheet that point outside the workbook, in the order
/// their relationship ids are handed out.
fn external_links(sheet: &crate::model::Worksheet) -> Vec<&crate::model::Hyperlink> {
    sheet
        .hyperlinks
        .iter()
        .filter(|l| matches!(l.target, crate::model::LinkTarget::Outside(_)))
        .collect()
}

/// The relationships part of one sheet: its external hyperlinks first, then
/// whatever is attached to it.
fn sheet_rels(
    outside: &[&crate::model::Hyperlink],
    sheet: &crate::model::Worksheet,
    index: usize,
    tables: &[usize],
) -> String {
    let mut s = format!(
        concat!(
            "{decl}",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
        ),
        decl = XML_DECL
    );
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for (i, link) in outside.iter().enumerate() {
        let crate::model::LinkTarget::Outside(target) = &link.target else {
            continue;
        };
        // An address outside the package is never resolved against it, which
        // is what `TargetMode="External"` says.
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{rel}/hyperlink" Target="{}" TargetMode="External"/>"#,
            i + 1,
            escape(target)
        );
    }
    for (i, attached) in sheet.attachments.iter().enumerate() {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{}" Target="{}"/>"#,
            outside.len() + 1 + i,
            escape(&attached.kind),
            escape(&relative_target("xl/worksheets", &attached.target))
        );
    }
    if !sheet.comments.is_empty() {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{rel}/comments" Target="../comments{index}.xml"/>"#,
            outside.len() + 1 + sheet.attachments.len(),
        );
    }
    let base = table_rel_base(sheet);
    for (i, number) in tables.iter().enumerate() {
        let _ = write!(
            s,
            r#"<Relationship Id="rId{}" Type="{rel}/table" Target="../tables/table{number}.xml"/>"#,
            base + i,
        );
    }
    s.push_str("</Relationships>");
    s
}

/// Renders one `<hyperlink>` element; `id` numbers the relationship of an external link.
fn hyperlink_xml(link: &crate::model::Hyperlink, id: Option<usize>) -> String {
    let mut s = format!(r#"<hyperlink ref="{}""#, sqref(&[link.range]));
    match &link.target {
        crate::model::LinkTarget::Inside(location) => {
            let _ = write!(s, r#" location="{}""#, escape(location));
        }
        crate::model::LinkTarget::Outside(_) => {
            if let Some(id) = id {
                let _ = write!(s, r#" r:id="rId{id}""#);
            }
        }
    }
    if let Some(display) = &link.display {
        let _ = write!(s, r#" display="{}""#, escape(display));
    }
    if let Some(tooltip) = &link.tooltip {
        let _ = write!(s, r#" tooltip="{}""#, escape(tooltip));
    }
    s.push_str("/>");
    s
}

/// Renders what guards the sheet - its protection, the ranges exempted from
/// it, and its filter - in the order the schema fixes them.
fn guards_xml(sheet: &crate::model::Worksheet) -> String {
    let mut s = sheet_protection_xml(&sheet.protection);
    if !sheet.protected_ranges.is_empty() {
        s.push_str("<protectedRanges>");
        for range in &sheet.protected_ranges {
            s.push_str(&protected_range_xml(range));
        }
        s.push_str("</protectedRanges>");
    }
    s.push_str(&scenarios_xml(&sheet.scenarios));
    if let Some(filter) = &sheet.auto_filter {
        s.push_str(&auto_filter_xml(filter));
    }
    // `<sortState>` follows `<autoFilter>` in the schema's fixed order.
    if let Some(sort) = &sheet.sort_state {
        s.push_str(&sort_state_xml(sort));
    }
    s
}

/// Renders `<sheetProtection>`, or nothing when the sheet forbids nothing.
///
/// A flag left unset stays out of the element: the attribute defaults differ
/// from flag to flag, and writing `0` for an absent one would tell Excel
/// something the file never said.
fn sheet_protection_xml(protection: &crate::model::SheetProtection) -> String {
    if protection.is_empty() {
        return String::new();
    }
    let mut s = String::from("<sheetProtection");
    // The password comes first, as the schema orders the attributes.
    if let Some(password) = &protection.password {
        for (name, value) in password.to_attrs(crate::model::protection::PasswordAttrs::PLAIN) {
            let _ = write!(s, r#" {name}="{}""#, escape(&value));
        }
    }
    for (name, value) in protection.flags() {
        if let Some(on) = value {
            let _ = write!(s, r#" {name}="{}""#, u8::from(on));
        }
    }
    s.push_str("/>");
    s
}

/// Renders `<scenarios>`, which sits after `<protectedRanges>`.
fn scenarios_xml(scenarios: &crate::model::scenario::Scenarios) -> String {
    let mut s = String::new();
    if scenarios.is_empty() {
        return s;
    }
    s.push_str("<scenarios");
    if let Some(n) = scenarios.current {
        let _ = write!(s, r#" current="{n}""#);
    }
    if let Some(n) = scenarios.show {
        let _ = write!(s, r#" show="{n}""#);
    }
    if !scenarios.result_cells.is_empty() {
        let refs: Vec<String> = scenarios
            .result_cells
            .iter()
            .map(ToString::to_string)
            .collect();
        let _ = write!(s, r#" sqref="{}""#, refs.join(" "));
    }
    s.push('>');
    for scenario in &scenarios.list {
        let _ = write!(
            s,
            r#"<scenario name="{}" count="{}""#,
            escape(&scenario.name),
            scenario.inputs.len()
        );
        if scenario.locked {
            s.push_str(r#" locked="1""#);
        }
        if scenario.hidden {
            s.push_str(r#" hidden="1""#);
        }
        if let Some(user) = &scenario.user {
            let _ = write!(s, r#" user="{}""#, escape(user));
        }
        if let Some(comment) = &scenario.comment {
            let _ = write!(s, r#" comment="{}""#, escape(comment));
        }
        s.push('>');
        for input in &scenario.inputs {
            let _ = write!(
                s,
                r#"<inputCells r="{}" val="{}""#,
                input.cell,
                escape(&input.value)
            );
            if input.deleted {
                s.push_str(r#" deleted="1""#);
            }
            if input.undone {
                s.push_str(r#" undone="1""#);
            }
            if let Some(id) = input.number_format_id {
                let _ = write!(s, r#" numFmtId="{id}""#);
            }
            s.push_str("/>");
        }
        s.push_str("</scenario>");
    }
    s.push_str("</scenarios>");
    s
}

/// Renders one `<protectedRange>`.
fn protected_range_xml(range: &crate::model::ProtectedRange) -> String {
    let mut s = format!(
        r#"<protectedRange name="{}" sqref="{}""#,
        escape(&range.name),
        sqref(&range.sqref)
    );
    if let Some(password) = &range.password {
        for (name, value) in password.to_attrs(crate::model::protection::PasswordAttrs::PLAIN) {
            let _ = write!(s, r#" {name}="{}""#, escape(&value));
        }
    }
    if !range.security_descriptor.is_empty() {
        let _ = write!(
            s,
            r#" securityDescriptor="{}""#,
            escape(&range.security_descriptor)
        );
    }
    s.push_str("/>");
    s
}

/// Renders `<autoFilter>` with the criteria of every column that has any.
fn auto_filter_xml(filter: &crate::model::AutoFilter) -> String {
    let mut s = format!(r#"<autoFilter ref="{}""#, filter.range);
    if filter.columns.is_empty() && filter.sort_state.is_none() {
        s.push_str("/>");
        return s;
    }
    s.push('>');
    for column in &filter.columns {
        let _ = write!(s, r#"<filterColumn colId="{}""#, column.col_id);
        if column.hidden_button {
            s.push_str(r#" hiddenButton="1""#);
        }
        match &column.filter {
            None => s.push_str("/>"),
            Some(criteria) => {
                s.push('>');
                s.push_str(&column_filter_xml(criteria));
                s.push_str("</filterColumn>");
            }
        }
    }
    if let Some(sort) = &filter.sort_state {
        s.push_str(&sort_state_xml(sort));
    }
    s.push_str("</autoFilter>");
    s
}

/// Renders `<sortState>` and its keys.
fn sort_state_xml(sort: &crate::model::SortState) -> String {
    let mut s = format!(r#"<sortState ref="{}""#, sort.range);
    if sort.column_sort {
        s.push_str(r#" columnSort="1""#);
    }
    if sort.case_sensitive {
        s.push_str(r#" caseSensitive="1""#);
    }
    if let Some(method) = &sort.sort_method {
        let _ = write!(s, r#" sortMethod="{}""#, escape(method));
    }
    if sort.conditions.is_empty() {
        s.push_str("/>");
        return s;
    }
    s.push('>');
    for key in &sort.conditions {
        s.push_str("<sortCondition");
        if key.descending {
            s.push_str(r#" descending="1""#);
        }
        if let Some(by) = &key.sort_by {
            let _ = write!(s, r#" sortBy="{}""#, escape(by));
        }
        let _ = write!(s, r#" ref="{}""#, key.range);
        if let Some(list) = &key.custom_list {
            let _ = write!(s, r#" customList="{}""#, escape(list));
        }
        if let Some(dxf) = key.dxf {
            let _ = write!(s, r#" dxfId="{dxf}""#);
        }
        if let Some(set) = &key.icon_set {
            let _ = write!(s, r#" iconSet="{}""#, escape(set));
        }
        if let Some(id) = key.icon_id {
            let _ = write!(s, r#" iconId="{id}""#);
        }
        s.push_str("/>");
    }
    s.push_str("</sortState>");
    s
}

/// The attributes of a colour or icon filter; nothing for the other kinds.
fn marked_filter_attrs(filter: &crate::model::ColumnFilter) -> String {
    use crate::model::ColumnFilter;
    let mut s = String::new();
    match filter {
        ColumnFilter::Color { dxf, cell_color } => {
            if let Some(dxf) = dxf {
                let _ = write!(s, r#" dxfId="{dxf}""#);
            }
            if !*cell_color {
                s.push_str(r#" cellColor="0""#);
            }
        }
        ColumnFilter::Icon { icon_set, icon_id } => {
            let _ = write!(s, r#" iconSet="{}""#, escape(icon_set));
            if let Some(id) = icon_id {
                let _ = write!(s, r#" iconId="{id}""#);
            }
        }
        _ => {}
    }
    s
}

/// Renders the one child of a `<filterColumn>`.
fn column_filter_xml(filter: &crate::model::ColumnFilter) -> String {
    use crate::model::ColumnFilter;
    let mut s = format!("<{}", filter.tag());
    match filter {
        ColumnFilter::Values {
            blank,
            values,
            date_groups,
        } => {
            if *blank {
                s.push_str(r#" blank="1""#);
            }
            if values.is_empty() && date_groups.is_empty() {
                s.push_str("/>");
                return s;
            }
            s.push('>');
            for value in values {
                let _ = write!(s, r#"<filter val="{}"/>"#, escape(value));
            }
            for group in date_groups {
                s.push_str("<dateGroupItem");
                for (name, value) in group.parts() {
                    let _ = write!(s, r#" {name}="{value}""#);
                }
                let _ = write!(s, r#" dateTimeGrouping="{}"/>"#, escape(&group.grouping));
            }
        }
        ColumnFilter::Custom { and, rules } => {
            if *and {
                s.push_str(r#" and="1""#);
            }
            if rules.is_empty() {
                s.push_str("/>");
                return s;
            }
            s.push('>');
            for rule in rules {
                s.push_str("<customFilter");
                // `equal` is the schema's default and Excel leaves it off.
                if rule.operator != crate::model::FilterOperator::Equal {
                    let _ = write!(s, r#" operator="{}""#, rule.operator.as_str());
                }
                let _ = write!(s, r#" val="{}"/>"#, escape(&rule.value));
            }
        }
        ColumnFilter::Dynamic {
            kind,
            value,
            max_value,
        } => {
            let _ = write!(s, r#" type="{}""#, escape(kind));
            if let Some(value) = value {
                let _ = write!(s, r#" val="{}""#, escape(value));
            }
            if let Some(value) = max_value {
                let _ = write!(s, r#" maxVal="{}""#, escape(value));
            }
            s.push_str("/>");
            return s;
        }
        ColumnFilter::Top10 {
            value,
            percent,
            top,
            filter_value,
        } => {
            if *percent {
                s.push_str(r#" percent="1""#);
            }
            if !*top {
                s.push_str(r#" top="0""#);
            }
            if let Some(value) = value {
                let _ = write!(s, r#" val="{}""#, escape(value));
            }
            if let Some(value) = filter_value {
                let _ = write!(s, r#" filterVal="{}""#, escape(value));
            }
            s.push_str("/>");
            return s;
        }
        ColumnFilter::Color { .. } | ColumnFilter::Icon { .. } => {
            s.push_str(&marked_filter_attrs(filter));
            s.push_str("/>");
            return s;
        }
    }
    let _ = write!(s, "</{}>", filter.tag());
    s
}

/// Renders the `<sheetPr>` element, or nothing when it would say nothing.
fn sheet_pr_xml(props: &crate::model::SheetProperties) -> String {
    if *props == crate::model::SheetProperties::default() {
        return String::new();
    }
    let mut s = String::from("<sheetPr");
    // The name VBA addresses the sheet by, which survives a rename.
    if let Some(code_name) = &props.code_name {
        let _ = write!(s, r#" codeName="{}""#, escape(code_name));
    }
    s.push('>');
    if let Some(color) = &props.tab_color {
        let _ = write!(s, "<tabColor{}/>", color_attr(color));
    }
    if !props.summary_below || !props.summary_right {
        s.push_str("<outlinePr");
        if !props.summary_below {
            s.push_str(r#" summaryBelow="0""#);
        }
        if !props.summary_right {
            s.push_str(r#" summaryRight="0""#);
        }
        s.push_str("/>");
    }
    if props.fit_to_page {
        s.push_str(r#"<pageSetUpPr fitToPage="1"/>"#);
    }
    s.push_str("</sheetPr>");
    s
}

/// Renders the `<pageSetup>` element.
fn page_setup_xml(
    setup: &crate::model::PageSetup,
    sheet: &crate::model::Worksheet,
    outside: usize,
) -> String {
    if *setup == crate::model::PageSetup::default() {
        return String::new();
    }
    let mut s = String::from("<pageSetup");
    // The relationship id of the printer settings part, which is numbered
    // after the external hyperlinks the same way the other attachments are.
    if let Some(part) = &setup.printer_settings
        && let Some(index) = sheet.attachments.iter().position(|a| &a.target == part)
    {
        let _ = write!(s, r#" r:id="rId{}""#, outside + 1 + index);
    }
    for (name, value) in [
        ("paperSize", setup.paper_size),
        ("scale", setup.scale),
        ("firstPageNumber", setup.first_page_number),
        ("fitToWidth", setup.fit_to_width),
        ("fitToHeight", setup.fit_to_height),
        ("horizontalDpi", setup.horizontal_dpi),
        ("verticalDpi", setup.vertical_dpi),
    ] {
        if let Some(v) = value {
            let _ = write!(s, r#" {name}="{v}""#);
        }
    }
    if setup.over_then_down {
        s.push_str(r#" pageOrder="overThenDown""#);
    }
    if setup.orientation != crate::model::Orientation::Default {
        let _ = write!(s, r#" orientation="{}""#, setup.orientation.as_str());
    }
    for (name, on) in [
        ("usePrinterDefaults", false),
        ("blackAndWhite", setup.black_and_white),
        ("draft", setup.draft),
        ("useFirstPageNumber", setup.use_first_page_number),
    ] {
        if on {
            let _ = write!(s, r#" {name}="1""#);
        }
    }
    s.push_str("/>");
    s
}

/// Renders the `<headerFooter>` element.
fn header_footer_xml(hf: &crate::model::HeaderFooter) -> String {
    if !hf.is_meaningful() {
        return String::new();
    }
    let mut s = String::from("<headerFooter");
    if hf.different_odd_even {
        s.push_str(r#" differentOddEven="1""#);
    }
    if hf.different_first {
        s.push_str(r#" differentFirst="1""#);
    }
    if !hf.scale_with_doc {
        s.push_str(r#" scaleWithDoc="0""#);
    }
    if !hf.align_with_margins {
        s.push_str(r#" alignWithMargins="0""#);
    }
    s.push('>');
    for (tag, text) in [
        ("oddHeader", &hf.odd_header),
        ("oddFooter", &hf.odd_footer),
        ("evenHeader", &hf.even_header),
        ("evenFooter", &hf.even_footer),
        ("firstHeader", &hf.first_header),
        ("firstFooter", &hf.first_footer),
    ] {
        if !text.is_empty() {
            let _ = write!(s, "<{tag}>{}</{tag}>", escape(text));
        }
    }
    s.push_str("</headerFooter>");
    s
}

/// Renders an `sqref` attribute value: areas separated by spaces, with a
/// one-cell area written as the bare cell Excel writes.
fn sqref(ranges: &[crate::coordinate::Range]) -> String {
    ranges
        .iter()
        .map(|r| {
            if r.start == r.end {
                r.start.to_string()
            } else {
                r.to_string()
            }
        })
        .collect::<Vec<_>>()
        .join(" ")
}

/// Renders the `<sheetViews>` element: zoom, scroll position, panes and
/// selection. Excel restores the sheet to this state when the file is opened,
/// so it is written for every sheet, defaults included.
fn sheet_views_xml(view: &crate::model::SheetView) -> String {
    let mut s = String::from("<sheetViews><sheetView");
    if view.view != crate::model::SheetViewType::Normal {
        let _ = write!(s, r#" view="{}""#, view.view.as_str());
    }
    if view.tab_selected {
        s.push_str(r#" tabSelected="1""#);
    }
    // These three default to on, so the attribute only appears when off.
    if !view.show_grid_lines {
        s.push_str(r#" showGridLines="0""#);
    }
    if !view.show_row_col_headers {
        s.push_str(r#" showRowColHeaders="0""#);
    }
    if !view.show_zeros {
        s.push_str(r#" showZeros="0""#);
    }
    if view.right_to_left {
        s.push_str(r#" rightToLeft="1""#);
    }
    for (on, name) in [
        (view.show_formulas, "showFormulas"),
        (view.window_protection, "windowProtection"),
    ] {
        if on {
            let _ = write!(s, r#" {name}="1""#);
        }
    }
    for (on, name) in [
        (view.show_outline_symbols, "showOutlineSymbols"),
        (view.show_ruler, "showRuler"),
        (view.show_white_space, "showWhiteSpace"),
    ] {
        if !on {
            let _ = write!(s, r#" {name}="0""#);
        }
    }
    if let Some(color) = view.grid_color {
        let _ = write!(s, r#" defaultGridColor="0" colorId="{color}""#);
    }
    if let Some(cell) = view.top_left_cell {
        let _ = write!(s, r#" topLeftCell="{cell}""#);
    }
    for (name, zoom) in [
        ("zoomScale", view.zoom_scale),
        ("zoomScaleNormal", view.zoom_scale_normal),
        ("zoomScalePageLayoutView", view.zoom_scale_page_layout),
        ("zoomScaleSheetLayoutView", view.zoom_scale_sheet_layout),
    ] {
        if let Some(z) = zoom {
            let _ = write!(s, r#" {name}="{z}""#);
        }
    }
    let _ = write!(s, r#" workbookViewId="{}""#, view.workbook_view_id);

    if view.pane.is_none() && view.selections.is_empty() {
        s.push_str("/></sheetViews>");
        return s;
    }
    s.push('>');
    if let Some(p) = &view.pane {
        s.push_str("<pane");
        if p.x_split > 0 {
            let _ = write!(s, r#" xSplit="{}""#, p.x_split);
        }
        if p.y_split > 0 {
            let _ = write!(s, r#" ySplit="{}""#, p.y_split);
        }
        if let Some(cell) = p.top_left_cell {
            let _ = write!(s, r#" topLeftCell="{cell}""#);
        }
        let _ = write!(
            s,
            r#" activePane="{}" state="{}"/>"#,
            p.active_pane.as_str(),
            p.state.as_str()
        );
    }
    for sel in &view.selections {
        s.push_str("<selection");
        if let Some(pane) = sel.pane {
            let _ = write!(s, r#" pane="{}""#, pane.as_str());
        }
        if let Some(cell) = sel.active_cell {
            let _ = write!(s, r#" activeCell="{cell}""#);
        }
        if !sel.sqref.is_empty() {
            let _ = write!(s, r#" sqref="{}""#, sqref(&sel.sqref));
        }
        s.push_str("/>");
    }
    s.push_str("</sheetView></sheetViews>");
    s
}

/// Renders one `<dataValidation>` element.
fn data_validation_xml(dv: &crate::model::DataValidation) -> String {
    let mut s = String::from("<dataValidation");
    if dv.kind != crate::model::ValidationType::None {
        let _ = write!(s, r#" type="{}""#, dv.kind.as_str());
    }
    if dv.error_style != crate::model::ValidationErrorStyle::Stop {
        let _ = write!(s, r#" errorStyle="{}""#, dv.error_style.as_str());
    }
    if dv.operator != crate::model::ValidationOperator::Between {
        let _ = write!(s, r#" operator="{}""#, dv.operator.as_str());
    }
    for (name, on) in [
        ("allowBlank", dv.allow_blank),
        ("showDropDown", dv.hide_drop_down),
        ("showInputMessage", dv.show_input_message),
        ("showErrorMessage", dv.show_error_message),
    ] {
        if on {
            let _ = write!(s, r#" {name}="1""#);
        }
    }
    for (name, text) in [
        ("errorTitle", &dv.error_title),
        ("error", &dv.error),
        ("promptTitle", &dv.prompt_title),
        ("prompt", &dv.prompt),
    ] {
        if !text.is_empty() {
            let _ = write!(s, r#" {name}="{}""#, escape(text));
        }
    }
    let _ = write!(s, r#" sqref="{}">"#, sqref(&dv.sqref));
    if !dv.formula1.is_empty() {
        let _ = write!(s, "<formula1>{}</formula1>", escape(&dv.formula1));
    }
    if !dv.formula2.is_empty() {
        let _ = write!(s, "<formula2>{}</formula2>", escape(&dv.formula2));
    }
    s.push_str("</dataValidation>");
    s
}

/// Renders one `<col>` element.
fn column_xml(run: &crate::model::ColumnRun) -> String {
    let mut s = format!(
        r#"<col min="{}" max="{}""#,
        run.first.one_based(),
        run.last.one_based()
    );
    if let Some(width) = run.width {
        let _ = write!(s, r#" width="{}""#, trim_number(width));
    }
    if run.custom_width {
        s.push_str(r#" customWidth="1""#);
    }
    if run.hidden {
        s.push_str(r#" hidden="1""#);
    }
    if run.best_fit {
        s.push_str(r#" bestFit="1""#);
    }
    if let Some(style) = run.style {
        let _ = write!(s, r#" style="{}""#, style.index());
    }
    if run.outline_level > 0 {
        let _ = write!(s, r#" outlineLevel="{}""#, run.outline_level);
    }
    if run.collapsed {
        s.push_str(r#" collapsed="1""#);
    }
    s.push_str("/>");
    s
}

/// Renders the opening tag of a `<row>`.
fn row_open_tag(
    row: crate::coordinate::Row,
    props: Option<&crate::model::RowProperties>,
) -> String {
    let mut s = format!(r#"<row r="{}""#, row.one_based());
    if let Some(p) = props {
        if let Some(height) = p.height {
            let _ = write!(s, r#" ht="{}""#, trim_number(height));
        }
        if p.custom_height {
            s.push_str(r#" customHeight="1""#);
        }
        if p.hidden {
            s.push_str(r#" hidden="1""#);
        }
        if let Some(style) = p.style {
            // A row style only applies when customFormat says so.
            let _ = write!(s, r#" s="{}" customFormat="1""#, style.index());
        }
        if p.outline_level > 0 {
            let _ = write!(s, r#" outlineLevel="{}""#, p.outline_level);
        }
        if p.collapsed {
            s.push_str(r#" collapsed="1""#);
        }
    }
    s.push('>');
    s
}

/// A sheet's array formulas, indexed once so a cell does not search them all.
struct ArrayFormulas<'a> {
    /// Each array by its top-left cell, where its formula is written.
    starts: HashMap<crate::coordinate::CellRef, &'a crate::coordinate::Range>,
    /// The arrays wider than one cell: only these have cells inside that are
    /// not their start.
    // ponytail: linear scan per formula cell; old books are mostly single-cell
    // arrays, so this stays short. An interval index if blocks run to thousands.
    blocks: Vec<&'a crate::coordinate::Range>,
}

impl<'a> ArrayFormulas<'a> {
    fn of(ranges: &'a [crate::coordinate::Range]) -> Self {
        Self {
            // Reversed, so that of two arrays claiming one start the first
            // wins, as a search in file order would pick it.
            starts: ranges.iter().rev().map(|r| (r.start, r)).collect(),
            blocks: ranges.iter().filter(|r| r.start != r.end).collect(),
        }
    }

    /// Whether `at` lies inside somebody else's array.
    fn inside_another(&self, at: crate::coordinate::CellRef) -> bool {
        self.blocks.iter().any(|r| r.contains(at) && r.start != at)
    }
}

/// Renders one `<c>` element.
fn cell_xml(
    at: crate::coordinate::CellRef,
    cell: &crate::model::Cell,
    pool: &StringPool<'_>,
    arrays: &ArrayFormulas<'_>,
) -> String {
    let style = cell.style.index();
    let attrs = if style == 0 {
        format!(r#" r="{at}""#)
    } else {
        format!(r#" r="{at}" s="{style}""#)
    };

    match &cell.value {
        CellValue::Empty => format!("<c{attrs}/>"),
        CellValue::Number(n) => format!("<c{attrs}><v>{}</v></c>", number(*n)),
        CellValue::Bool(b) => format!(r#"<c{attrs} t="b"><v>{}</v></c>"#, u8::from(*b)),
        CellValue::Error(e) => format!(r#"<c{attrs} t="e"><v>{}</v></c>"#, escape(e.as_str())),
        CellValue::RichText(runs) => match pool.rich.get(&rich_key(runs)) {
            Some(i) => format!(r#"<c{attrs} t="s"><v>{i}</v></c>"#),
            None => format!("<c{attrs}/>"),
        },
        CellValue::Text(text) => match pool.plain.get(text.as_ref()) {
            Some(i) => format!(r#"<c{attrs} t="s"><v>{i}</v></c>"#),
            // A string missing from the pool would be a bug in the pool builder;
            // write it inline rather than silently emit a dangling index.
            None => format!(
                r#"<c{attrs} t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>"#,
                escape(text)
            ),
        },
        CellValue::Formula { formula, cached } => {
            // Inside somebody else's array the cell shows its part of the
            // result and holds no formula of its own. Excel refuses a file
            // that says otherwise: it strips the cells of the sheet.
            if arrays.inside_another(at) {
                let value = cached.as_deref().unwrap_or(&CellValue::Empty).clone();
                return cell_xml(
                    at,
                    &crate::model::Cell {
                        value,
                        ..cell.clone()
                    },
                    pool,
                    arrays,
                );
            }
            let value = match cached.as_deref() {
                Some(CellValue::Number(n)) => format!("<v>{}</v>", number(*n)),
                Some(CellValue::Bool(b)) => format!("<v>{}</v>", u8::from(*b)),
                Some(CellValue::Text(t)) => format!("<v>{}</v>", escape(t)),
                Some(CellValue::Error(e)) => format!("<v>{}</v>", escape(e.as_str())),
                _ => String::new(),
            };
            // The cached type has to be declared on the cell, or a reader takes
            // a cached string for a number.
            let t = match cached.as_deref() {
                Some(CellValue::Text(_)) => r#" t="str""#,
                Some(CellValue::Bool(_)) => r#" t="b""#,
                Some(CellValue::Error(_)) => r#" t="e""#,
                _ => "",
            };
            let f = match arrays.starts.get(&at) {
                Some(r) => format!(r#"<f t="array" ref="{r}">"#),
                None => "<f>".to_owned(),
            };
            format!("<c{attrs}{t}>{f}{}</f>{value}</c>", escape(formula))
        }
    }
}

/// Renders a number the way xlsx expects: no exponent for ordinary magnitudes,
/// no trailing `.0` on integers.
fn number(n: f64) -> String {
    if n.fract() == 0.0 && n.abs() < 1e15 {
        format!("{n:.0}")
    } else {
        let s = n.to_string();
        // Rust prints 1e-7 as "1e-7"; xlsx wants "1E-7".
        if s.contains('e') {
            s.replace('e', "E")
        } else {
            s
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{auto_filter_xml, sheet_protection_xml, worksheet};
    use crate::coordinate::Range;
    use crate::model::CellValue;
    use crate::model::{
        Attachment, AutoFilter, ColumnFilter, CustomFilter, FilterOperator, PasswordHash,
        SheetProtection, Worksheet,
    };

    fn range(s: &str) -> Range {
        Range::parse(s).unwrap_or_else(|e| unreachable!("{e}"))
    }

    #[test]
    fn a_sheet_that_forbids_nothing_gets_no_element() {
        assert_eq!(sheet_protection_xml(&SheetProtection::default()), "");
    }

    #[test]
    fn only_the_flags_the_sheet_decided_are_written() {
        let xml = sheet_protection_xml(&SheetProtection {
            sheet: Some(true),
            objects: Some(false),
            password: Some(PasswordHash::Legacy("CC1A".into())),
            ..SheetProtection::default()
        });
        assert_eq!(
            xml,
            r#"<sheetProtection password="CC1A" sheet="1" objects="0"/>"#
        );
    }

    #[test]
    fn a_filter_with_no_criteria_is_one_empty_element() {
        assert_eq!(
            auto_filter_xml(&AutoFilter::new(range("A1:C9"))),
            r#"<autoFilter ref="A1:C9"/>"#
        );
    }

    #[test]
    fn the_default_operator_is_left_off_as_excel_leaves_it() {
        let mut filter = AutoFilter::new(range("A1:A9"));
        filter.column_at(0).filter = Some(ColumnFilter::Custom {
            and: false,
            rules: vec![
                CustomFilter {
                    operator: FilterOperator::Equal,
                    value: "x".into(),
                },
                CustomFilter {
                    operator: FilterOperator::LessThan,
                    value: "9".into(),
                },
            ],
        });
        let xml = auto_filter_xml(&filter);
        assert!(
            xml.contains(r#"<customFilter val="x"/>"#),
            "equal is the schema default: {xml}"
        );
        assert!(xml.contains(r#"<customFilter operator="lessThan" val="9"/>"#));
    }

    #[test]
    fn a_table_is_written_as_its_own_part_and_pointed_at() {
        // Two things have to line up: the part carries the table, and
        // `<tableParts>` is what makes Excel treat the range as a table rather
        // than as ordinary cells.
        let mut sheet = Worksheet::new("T").unwrap_or_default();
        sheet.tables.push(crate::model::table::Table {
            id: 1,
            name: "Sales".into(),
            display_name: "Sales".into(),
            range: crate::coordinate::Range::parse("A1:C5").expect("a written range"),
            header_row_count: None,
            totals_row_count: None,
            auto_filter: crate::coordinate::Range::parse("A1:C5")
                .ok()
                .map(crate::model::AutoFilter::new),
            sort_state: None,
            columns: vec![crate::model::table::TableColumn {
                id: 1,
                name: "Region".into(),
                ..crate::model::table::TableColumn::default()
            }],
            extensions: None,
            style: Some(crate::model::table::TableStyle {
                name: Some("TableStyleMedium2".into()),
                show_row_stripes: true,
                ..crate::model::table::TableStyle::default()
            }),
        });
        let empty = crate::model::Spreadsheet::empty();
        let pool = super::collect_shared_strings(&empty);
        let xml = worksheet(&sheet, &pool, &[]);
        assert!(
            xml.contains(r#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#),
            "{xml}"
        );

        let part = super::table_xml(&sheet.tables[0]);
        assert!(
            part.contains(r#"name="Sales" displayName="Sales" ref="A1:C5">"#),
            "{part}"
        );
        assert!(
            part.contains(r#"<tableColumn id="1" name="Region"/>"#),
            "{part}"
        );
        assert!(part.contains(r#"showRowStripes="1""#), "{part}");
        // Neither count was in the file, so neither is written back.
        assert!(!part.contains("headerRowCount"), "{part}");
        assert!(!part.contains("totalsRowCount"), "{part}");

        // The relationship the element points at has to exist.
        let rels = super::sheet_rels(&[], &sheet, 1, &[1]);
        assert!(
            rels.contains(r#"Id="rId1""#) && rels.contains("../tables/table1.xml"),
            "{rels}"
        );
    }

    #[test]
    fn a_row_without_cells_takes_its_place_among_the_others() {
        // ECMA-376 wants `<row>` in ascending order; a row that only carries
        // a height must not trail after the rows that hold cells.
        let mut sheet = Worksheet::new("R").unwrap_or_default();
        sheet.set(
            crate::coordinate::CellRef::parse("A5").expect("a written reference"),
            1.0,
        );
        let row = |n| crate::coordinate::Row::from_one_based(n).expect("a written row");
        sheet.set_row_height(row(2), Some(30.0));
        sheet.set_row_height(row(5), Some(20.0));
        sheet.set_row_hidden(row(9), true);
        let empty = crate::model::Spreadsheet::empty();
        let pool = super::collect_shared_strings(&empty);
        let xml = worksheet(&sheet, &pool, &[]);
        let rows: Vec<&str> = xml
            .split("<row r=\"")
            .skip(1)
            .filter_map(|tail| tail.split('"').next())
            .collect();
        assert_eq!(rows, ["2", "5", "9"], "{xml}");
        assert!(xml.contains(r#"<row r="5" ht="20" customHeight="1"><c r="A5">"#));
    }

    #[test]
    fn many_rows_with_heights_are_written_in_one_pass() {
        // Each row with a height used to search every cell of the sheet:
        // 20 000 rows made it 4e8 steps.
        let mut sheet = Worksheet::new("R").unwrap_or_default();
        for n in 1..=20_000u32 {
            let row = crate::coordinate::Row::new(n * 2).expect("a small row");
            sheet.set_row_height(row, Some(15.0));
            sheet.set(
                crate::coordinate::CellRef::new(crate::coordinate::Col::new(0).expect("A"), row),
                f64::from(n),
            );
            let empty_row = crate::coordinate::Row::new(n * 2 + 1).expect("a small row");
            sheet.set_row_height(empty_row, Some(15.0));
        }
        let empty = crate::model::Spreadsheet::empty();
        let pool = super::collect_shared_strings(&empty);
        let started = std::time::Instant::now();
        let xml = worksheet(&sheet, &pool, &[]);
        assert!(
            started.elapsed() < std::time::Duration::from_secs(2),
            "took {:?}",
            started.elapsed()
        );
        assert_eq!(xml.matches("<row ").count(), 40_000);
    }

    #[test]
    fn many_array_formulas_are_looked_up_not_searched() {
        // Old workbooks hold a single-cell array formula in every row; each
        // formula cell used to walk the whole list of them, twice.
        let mut sheet = Worksheet::new("R").unwrap_or_default();
        let a = crate::coordinate::Col::new(0).expect("A");
        for n in 0..60_000u32 {
            let at = crate::coordinate::CellRef::new(
                a,
                crate::coordinate::Row::new(n).expect("a small row"),
            );
            sheet.set(
                at,
                CellValue::Formula {
                    formula: "SUM(B1:B2*C1:C2)".into(),
                    cached: None,
                },
            );
            sheet.array_formulas.push(Range::new(at, at));
        }
        // One real block, so the "inside another array" path still runs.
        sheet.array_formulas.push(range("B1:B2"));
        sheet.set(
            crate::coordinate::CellRef::parse("B2").expect("a written reference"),
            CellValue::Formula {
                formula: "ROW()".into(),
                cached: Some(Box::new(CellValue::Number(2.0))),
            },
        );
        let empty = crate::model::Spreadsheet::empty();
        let pool = super::collect_shared_strings(&empty);
        let started = std::time::Instant::now();
        let xml = worksheet(&sheet, &pool, &[]);
        assert!(
            started.elapsed() < std::time::Duration::from_secs(2),
            "took {:?}",
            started.elapsed()
        );
        assert_eq!(xml.matches(r#"<f t="array" ref="A"#).count(), 60_000);
        assert!(xml.contains(r#"<c r="B2"><v>2</v></c>"#), "{xml:.2000}");
    }

    #[test]
    fn many_outside_links_are_numbered_without_a_search() {
        // Each external link used to look itself up among all of them.
        let mut sheet = Worksheet::new("R").unwrap_or_default();
        let a = crate::coordinate::Col::new(0).expect("A");
        for n in 0..60_000u32 {
            let at = crate::coordinate::CellRef::new(
                a,
                crate::coordinate::Row::new(n).expect("a small row"),
            );
            let target = if n == 1 {
                crate::model::LinkTarget::Inside("R!A1".into())
            } else {
                crate::model::LinkTarget::Outside(format!("https://example.com/{n}"))
            };
            sheet.hyperlinks.push(crate::model::Hyperlink {
                range: Range::new(at, at),
                target,
                display: None,
                tooltip: None,
            });
        }
        let outside = super::external_links(&sheet);
        let started = std::time::Instant::now();
        let xml = super::print_tail_xml(&sheet, &outside);
        assert!(
            started.elapsed() < std::time::Duration::from_secs(2),
            "took {:?}",
            started.elapsed()
        );
        // The inside link takes no id, so the one after it gets the second.
        assert!(
            xml.contains(r#"<hyperlink ref="A1" r:id="rId1"/>"#),
            "{xml:.300}"
        );
        assert!(
            xml.contains(r#"<hyperlink ref="A3" r:id="rId2"/>"#),
            "{xml:.300}"
        );
        assert!(xml.contains(r#"<hyperlink ref="A60000" r:id="rId59999"/>"#));
    }

    /// The names of the elements one level inside `root`, in the order written.
    fn children_of(xml: &str, root: &str) -> Vec<String> {
        let mut reader = quick_xml::Reader::from_str(xml);
        let mut depth = 0usize;
        let mut out = Vec::new();
        loop {
            match reader.read_event() {
                Err(e) => panic!("what we wrote does not parse: {e}"),
                Ok(quick_xml::events::Event::Eof) => break,
                Ok(quick_xml::events::Event::Start(e)) => {
                    let name = e.local_name().as_ref().to_owned();
                    if depth == 1 {
                        out.push(name);
                    }
                    depth += 1;
                }
                Ok(quick_xml::events::Event::Empty(e)) => {
                    if depth == 1 {
                        out.push(e.local_name().as_ref().to_owned());
                    }
                }
                Ok(quick_xml::events::Event::End(_)) => depth -= 1,
                Ok(_) => {}
            }
        }
        assert!(xml.contains(&format!("<{root}")), "no <{root}> in {xml}");
        out
    }

    /// Asserts that `written` is a subsequence of `schema`: every element is one
    /// the schema names, and no two of them are out of order.
    fn assert_in_schema_order(written: &[String], schema: &[&str], what: &str) {
        let mut expected = schema.iter();
        for name in written {
            assert!(
                schema.contains(&name.as_str()),
                "{what}: <{name}> is not in the schema's sequence; \
                 add it to the list at the position the schema gives it"
            );
            assert!(
                expected.any(|s| *s == name),
                "{what}: <{name}> is written out of order; \
                 the schema's sequence is fixed and Excel rejects a file that breaks it"
            );
        }
    }

    /// The order `CT_Worksheet` fixes, as far as this writer emits it.
    const WORKSHEET_ORDER: [&str; 24] = [
        "sheetPr",
        "dimension",
        "sheetViews",
        "sheetFormatPr",
        "cols",
        "sheetData",
        "sheetProtection",
        "protectedRanges",
        "scenarios",
        "autoFilter",
        "sortState",
        "mergeCells",
        "conditionalFormatting",
        "dataValidations",
        "hyperlinks",
        "printOptions",
        "pageMargins",
        "pageSetup",
        "headerFooter",
        "rowBreaks",
        "colBreaks",
        "drawing",
        "legacyDrawing",
        "tableParts",
    ];

    /// The order `CT_Stylesheet` fixes, as far as this writer emits it.
    const STYLES_ORDER: [&str; 8] = [
        "numFmts",
        "fonts",
        "fills",
        "borders",
        "cellStyleXfs",
        "cellXfs",
        "cellStyles",
        "dxfs",
    ];

    /// A sheet using every element the writer knows how to emit, so the order
    /// check has something to check.
    fn crowded_sheet() -> Worksheet {
        use crate::model::{
            CfRule, ConditionalFormat, DataValidation, Hyperlink, LinkTarget, PageBreak,
            PrintOptions, ProtectedRange,
        };
        let mut sheet = Worksheet::new("Crowded").unwrap_or_default();
        sheet.properties.code_name = Some("Sheet1".into());
        let a = crate::coordinate::Col::new(0).expect("column A");
        sheet.columns.push(crate::model::ColumnRun {
            first: a,
            last: a,
            width: Some(12.0),
            custom_width: true,
            hidden: false,
            best_fit: false,
            style: None,
            outline_level: 0,
            collapsed: false,
        });
        *sheet.entry(range("A1:A1").start) = crate::model::Cell {
            value: crate::model::CellValue::Number(1.0),
            style: crate::style::StyleId::default(),
        };
        sheet.protection.sheet = Some(true);
        sheet.protected_ranges.push(ProtectedRange {
            name: "r1".into(),
            sqref: vec![range("A1:B2")],
            password: None,
            security_descriptor: String::new(),
        });
        sheet.auto_filter = Some(AutoFilter::new(range("A1:C9")));
        sheet.merges.push(range("D1:E1"));
        sheet.conditional_formats.push(ConditionalFormat {
            sqref: vec![range("A1:A9")],
            rules: vec![CfRule::default()],
        });
        sheet.data_validations.push(DataValidation {
            sqref: vec![range("B1:B9")],
            formula1: "1".into(),
            ..DataValidation::default()
        });
        sheet.hyperlinks.push(Hyperlink {
            range: range("C1:C1"),
            target: LinkTarget::Inside("Crowded!A1".into()),
            display: None,
            tooltip: None,
        });
        sheet.print_options = PrintOptions {
            headings: true,
            ..PrintOptions::default()
        };
        sheet.header_footer.odd_header = "&Ltitle".into();
        sheet.row_breaks.push(PageBreak {
            at: 5,
            max: None,
            manual: true,
        });
        sheet.col_breaks.push(PageBreak {
            at: 3,
            max: None,
            manual: true,
        });
        for kind in [
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        ] {
            sheet.attachments.push(Attachment {
                kind: kind.into(),
                target: "xl/whatever.xml".into(),
            });
        }
        sheet
    }

    #[test]
    fn a_sheet_writes_its_elements_in_the_order_the_schema_fixes() {
        // Excel rejects a worksheet part whose elements are in the wrong order,
        // and reading our own file back does not notice: the reader dispatches
        // on the element name and takes them in any order at all. So this is
        // the only test that would catch a writer emitting them shuffled.
        let sheet = crowded_sheet();
        let empty = crate::model::Spreadsheet::empty();
        let pool = super::collect_shared_strings(&empty);
        let xml = worksheet(&sheet, &pool, &[]);
        let written = children_of(&xml, "worksheet");
        assert!(
            written.len() >= 18,
            "the fixture stopped covering the vocabulary: only {written:?}"
        );
        assert_in_schema_order(&written, &WORKSHEET_ORDER, "worksheet");
    }

    #[test]
    fn the_style_tables_are_written_in_the_order_the_schema_fixes() {
        let mut book = crate::model::Spreadsheet::empty();
        book.styles.intern(crate::style::Style {
            number_format: crate::style::NumberFormat::Custom("0.000".into()),
            font: crate::style::Font {
                bold: true,
                ..crate::style::Font::default()
            },
            ..crate::style::Style::default()
        });
        book.styles
            .differential
            .push(crate::style::DifferentialStyle::default());
        let written = children_of(&super::styles(&book), "styleSheet");
        assert_eq!(
            written,
            [
                "numFmts",
                "fonts",
                "fills",
                "borders",
                "cellStyleXfs",
                "cellXfs",
                "cellStyles",
                "dxfs"
            ],
            "every style table is written, and in this order"
        );
        assert_in_schema_order(&written, &STYLES_ORDER, "styleSheet");
    }
}
