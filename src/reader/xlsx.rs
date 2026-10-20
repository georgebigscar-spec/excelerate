//! Reading xlsx workbooks.
//!
//!
//! An xlsx file is a zip of XML parts wired together by relationship files.
//! The chain this reader walks is:
//!
//! ```text
//! _rels/.rels          -> xl/workbook.xml            (the officeDocument)
//! xl/workbook.xml      -> <sheet name r:id>          (names and tab order)
//! xl/_rels/workbook.xml.rels -> r:id to part path    (which XML holds a sheet)
//! xl/sharedStrings.xml -> the string pool cells refer to by index
//! xl/worksheets/N.xml  -> the cells themselves
//! ```
//!
//! Everything read here comes from an untrusted file, so sizes are capped and
//! nothing is allocated on the strength of a number the file supplies. Parsing
//! is streaming: no XML document is ever held in memory whole.

use crate::coordinate::{CellRef, Col, Range, Row, shift_references};
use crate::error::{CellError, Error, Result};
use crate::model::chart::{Chart, ChartEx, ChartExOrigin, ChartOrigin};
use crate::model::pivot::{
    CacheField, CacheSource, DataField, PivotAxis, PivotCache, PivotField, PivotOrigin,
    PivotStyleInfo, PivotTable, Subtotal,
};
use crate::model::protection::PasswordAttrs;
use crate::model::table::{Table, TableColumn, TableStyle};
use crate::model::{
    Attachment, AutoFilter, Cell, CellValue, CfOperator, CfRule, CfRuleType, CfScale, CfValue,
    CfValueType, ColumnFilter, ColumnRun, Comment, ConditionalFormat, CustomFilter, DataValidation,
    DateGroup, DefinedName, ExternalBook, ExternalSheet, FilterOperator, Hyperlink, LinkTarget,
    OpaquePart, Orientation, PageBreak, Pane, PanePosition, PaneState, PasswordHash,
    ProtectedRange, RowProperties, Selection, SheetView, SheetViewType, SheetVisibility,
    SortCondition, SortState, Spreadsheet, TextRun, ValidationErrorStyle, ValidationOperator,
    ValidationType, WorkbookProtection, Worksheet,
};
use crate::progress::{Options, RowBatch, Stage};
use crate::shared::date::Epoch;
use crate::style::{
    Alignment, Border, BorderStyle, Borders, Color, DiagonalDirection, DiffBorders, DiffFill,
    DiffFont, DifferentialStyle, Fill, Font, FontScheme, HorizontalAlign, NumberFormat, Pattern,
    Protection, ProtectionState, Script, Style, StyleId, StyleTable, Underline, VerticalAlign,
};
use quick_xml::Reader;
use quick_xml::events::Event;
use std::collections::{BTreeMap, HashMap};
use std::io::{BufReader, Read, Seek};

pub use super::zipxml::MAX_UNCOMPRESSED_SIZE;
use super::zipxml::{attr, attrs, is_true, push_entity, uncompressed_size};

/// Reads an xlsx workbook from a file.
///
/// # Errors
/// [`Error::Xlsx`] if the file is not a readable xlsx package, and the parsing
/// errors of the parts it contains.
pub fn read_xlsx(path: impl AsRef<std::path::Path>) -> Result<Spreadsheet> {
    let file = std::fs::File::open(path).map_err(|e| Error::Xlsx(e.to_string()))?;
    read_xlsx_from(BufReader::new(file))
}

/// Reads an xlsx workbook from any seekable reader.
///
/// # Errors
/// Same as [`read_xlsx`].
pub fn read_xlsx_from<R: Read + Seek>(source: R) -> Result<Spreadsheet> {
    read_xlsx_from_limited(source, MAX_UNCOMPRESSED_SIZE)
}

/// The same, with the cap on how far the package may expand given explicitly.
///
/// The default [`MAX_UNCOMPRESSED_SIZE`] is what stops a zip bomb, and a real
/// workbook rarely comes near it - but "rarely" is not "never", and a caller
/// who knows where the file came from can raise it.
///
/// # Errors
/// Same as [`read_xlsx`].
pub fn read_xlsx_from_limited<R: Read + Seek>(source: R, max_expanded: u64) -> Result<Spreadsheet> {
    read_xlsx_from_with(source, max_expanded, &Options::default())
}

/// The same, reporting its progress as it goes.
///
/// Sheets are the unit: one report before each, so a caller can name the sheet
/// being read. The total is unknown until the workbook part has been read, so
/// the first report carries `None` for it.
///
/// # Errors
/// Same as [`read_xlsx`].
pub fn read_xlsx_from_with<R: Read + Seek>(
    source: R,
    max_expanded: u64,
    options: &Options<'_>,
) -> Result<Spreadsheet> {
    let mut zip = zip::ZipArchive::new(source).map_err(|e| Error::Xlsx(e.to_string()))?;

    let total = uncompressed_size(&mut zip);
    let compressed = super::zipxml::compressed_size(&mut zip);
    if !super::zipxml::expansion_allowed(total, compressed, max_expanded) {
        return Err(Error::Xlsx(format!(
            "package expands to {total} bytes from {compressed}, over the {max_expanded} limit \
             and more than {} times its compressed size",
            super::zipxml::MAX_COMPRESSION_RATIO
        )));
    }

    let types = read_content_types(&mut zip);

    // The workbook part is normally xl/workbook.xml, but the package is free to
    // put it elsewhere and point at it from the root relationships.
    let workbook_path = find_workbook_part(&mut zip)?;
    let base = workbook_path.rsplit_once('/').map_or("", |(dir, _)| dir);

    let rels = read_relationships(&mut zip, &rels_path_for(&workbook_path))?;
    let shared = match rels.values().find(|r| r.kind.ends_with("/sharedStrings")) {
        Some(r) => read_shared_strings(&mut zip, &resolve(base, &r.target))?,
        None => Vec::new(),
    };

    let mut book = Spreadsheet::empty();
    book.template = matches!(types.get(&workbook_path), Some(t) if t.contains(".template."));
    if let Some(r) = rels.values().find(|r| r.kind.ends_with("/styles")) {
        let styles_path = resolve(base, &r.target);
        if let Ok(xml) = read_part(&mut zip, &styles_path) {
            book.style_extensions = trailing_extensions(&xml);
            let root = root_namespaces(&xml);
            let own = |e: Option<String>| e.map(|e| self_contained(&e, &root));
            book.table_styles = own(element_of(&xml, "tableStyles"));
            book.palette = own(element_of(&xml, "colors"));
        }
        book.styles = read_styles(&mut zip, &styles_path)?;
    }
    if let Some(r) = rels.values().find(|r| r.kind.ends_with("/theme")) {
        book.theme = read_part(&mut zip, &resolve(base, &r.target)).ok();
    }
    if let Ok(xml) = read_part(&mut zip, &workbook_path) {
        book.workbook_extensions = trailing_extensions(&xml);
    }
    let header = read_sheet_list(&mut zip, &workbook_path)?;
    book.epoch = header.epoch;
    book.workbook_properties = header.properties;
    book.calculation_properties = header.calculation;
    book.workbook_view = header.view;
    book.protection = header.protection;
    let sheet_count = header.sheets.len();
    for (done, (name, rel_id, visibility)) in header.sheets.into_iter().enumerate() {
        options.report(Stage::Reading, done, Some(sheet_count), &name);
        let rel = rels
            .get(&rel_id)
            .ok_or_else(|| Error::Xlsx(format!("sheet {name:?} points at unknown {rel_id:?}")))?;
        let path = resolve(base, &rel.target);
        // A sheet has relationships of its own; an external hyperlink keeps its
        // address there rather than in the sheet part, and a drawing or a
        // comment is reached only through them.
        let links = read_relationships(&mut zip, &rels_path_for(&path)).unwrap_or_default();
        let stream = options.row_sink(done).map(|sink| RowStream {
            sink,
            styles: &book.styles,
            epoch: book.epoch,
        });
        let mut sheet = read_sheet(&mut zip, &path, &name, &shared, &links, stream)?;
        sheet.shrink_to_fit();
        sheet.visibility = visibility;
        let sheet_base = path.rsplit_once('/').map_or("", |(dir, _)| dir);
        // The notes are modelled, so their part is read rather than carried;
        // the VML behind them still travels whole, because where a note sits
        // is a drawing and drawings are not modelled.
        if let Some(rel) = links
            .values()
            .find(|r| !r.external && r.kind.ends_with("/comments"))
        {
            sheet.comments =
                read_comments(&mut zip, &resolve(sheet_base, &rel.target)).unwrap_or_default();
        }
        read_note_boxes(&mut zip, &links, sheet_base, &mut sheet.comments);
        sheet.pivot_tables = read_sheet_pivots(&mut zip, &links, sheet_base);
        // The tables are modelled and written back, so their parts are read
        // rather than carried, the way the notes are.
        sheet.tables = read_sheet_tables(&mut zip, &links, sheet_base);
        // The charts are modelled and their parts carried as well: an
        // untouched chart goes back as the bytes it came in.
        // Charts, pictures and shapes all live in the drawings, which are
        // inflated once for the three.
        let drawings = sheet_drawings(&mut zip, &links, sheet_base);
        (sheet.charts, sheet.extended_charts) = read_sheet_charts(&mut zip, &drawings);
        sheet.images = read_sheet_images(&mut zip, &drawings);
        sheet.shapes = read_sheet_shapes(&drawings, book.theme.as_deref());
        sheet.attachments = attachments(&links, sheet_base, &["hyperlink", "comments", "table"]);
        book.add_sheet(sheet)?;
    }
    if book.sheets().is_empty() {
        return Err(Error::Xlsx("workbook declares no sheets".into()));
    }
    // Which tab Excel opens on. Out-of-range values are ignored rather than
    // rejected: the sheet list is what matters, the tab is only a view.
    let _ = book.set_active(header.active);
    book.defined_names = read_defined_names(&mut zip, &workbook_path)?;
    book.pivot_caches = read_pivot_caches(&mut zip, &header.pivot_caches, &rels, base);
    // The books this one links to, in the order `[N]` counts them. Their parts
    // stay in `parts` as well: nothing here writes them back, so what leaves
    // is what arrived.
    for id in &header.external {
        let Some(rel) = rels.get(id) else { continue };
        let path = resolve(base, &rel.target);
        book.external
            .push(read_external_link(&mut zip, &path).unwrap_or_default());
    }

    // Everything the workbook and the package point at that this crate does
    // not model. `calcChain` is left behind on purpose: it lists the order the
    // formulas were computed in, and after a rewrite it would be a lie.
    book.attachments = attachments(
        &rels,
        base,
        &["worksheet", "styles", "sharedStrings", "theme", "calcChain"],
    );
    let root = read_relationships(&mut zip, "_rels/.rels").unwrap_or_default();
    book.doc_props = attachments(&root, "", &["officeDocument"]);
    carry(&mut zip, &carried_roots(&book), &types, &mut book.parts);
    read_properties(&mut book);
    Ok(book)
}

/// The parts the workbook, the package and the sheets point at that travel
/// unmodelled; `carry` follows their relationships from there.
fn carried_roots(book: &Spreadsheet) -> Vec<String> {
    book.attachments
        .iter()
        .chain(&book.doc_props)
        .map(|a| a.target.clone())
        .chain(
            book.sheets()
                .iter()
                .flat_map(|s| s.attachments.iter().map(|a| a.target.clone())),
        )
        .collect()
}

/// The document properties of a package read straight from the zip, for a
/// reader that carries no parts: xlsb keeps them in the same XML parts as
/// xlsx.
pub(crate) fn package_properties<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> crate::model::DocumentProperties {
    use super::properties::{APP_REL, CORE_REL, CUSTOM_REL, read_app, read_core, read_custom};
    let root = read_relationships(zip, "_rels/.rels").unwrap_or_default();
    let mut text = |kind: &str| {
        let rel = root.values().find(|r| r.kind == kind && !r.external)?;
        read_part(zip, &resolve("", &rel.target)).ok()
    };
    let (core, app, custom) = (text(CORE_REL), text(APP_REL), text(CUSTOM_REL));
    let mut props = crate::model::DocumentProperties::default();
    if let Some(xml) = core {
        read_core(&xml, &mut props);
    }
    if let Some(xml) = app {
        read_app(&xml, &mut props);
    }
    if let Some(xml) = custom {
        props.custom = read_custom(&xml);
    }
    props
}

/// Fills [`Spreadsheet::properties`] from the parts the package points at.
/// The parts stay in `parts`: the writer keeps their bytes while the model
/// still says what they say.
fn read_properties(book: &mut Spreadsheet) {
    use super::properties::{APP_REL, CORE_REL, CUSTOM_REL, read_app, read_core, read_custom};
    let text = |kind: &str| {
        let target = &book.doc_props.iter().find(|a| a.kind == kind)?.target;
        let part = book.parts.iter().find(|p| &p.path == target)?;
        core::str::from_utf8(&part.data).ok().map(str::to_owned)
    };
    let (core, app, custom) = (text(CORE_REL), text(APP_REL), text(CUSTOM_REL));
    let props = &mut book.properties;
    if let Some(xml) = core {
        read_core(&xml, props);
    }
    if let Some(xml) = app {
        read_app(&xml, props);
    }
    if let Some(xml) = custom {
        props.custom = read_custom(&xml);
    }
}

/// One entry of a `.rels` part.
pub(crate) struct Relationship {
    pub(crate) kind: String,
    pub(crate) target: String,
    /// Whether the target lives outside the package and so is never a part.
    pub(crate) external: bool,
}

/// The relationships of a part that point at something this crate does not
/// handle itself, resolved to package paths.
fn attachments(
    rels: &HashMap<String, Relationship>,
    base: &str,
    handled: &[&str],
) -> Vec<Attachment> {
    let mut out: Vec<Attachment> = rels
        .values()
        .filter(|r| !r.external)
        .filter(|r| {
            let role = r.kind.rsplit('/').next().unwrap_or(&r.kind);
            !handled.contains(&role)
        })
        .map(|r| Attachment {
            kind: r.kind.clone(),
            target: resolve(base, &r.target),
        })
        .collect();
    // The relationships arrive from a map, so an order has to be imposed for
    // the result to be the same on every read.
    out.sort_by(|a, b| a.target.cmp(&b.target));
    out
}

/// The content type of every part, as `[Content_Types].xml` states it.
///
/// A carried part has to keep its type, or the package it is written into
/// declares nothing about it and no reader will touch it.
fn read_content_types<R: Read + Seek>(zip: &mut zip::ZipArchive<R>) -> HashMap<String, String> {
    let mut defaults: HashMap<String, String> = HashMap::new();
    let mut overrides: HashMap<String, String> = HashMap::new();
    let Ok(xml) = read_part(zip, "[Content_Types].xml") else {
        return overrides;
    };
    let mut reader = Reader::from_str(&xml);
    while let Ok(event) = reader.read_event() {
        match event {
            Event::Empty(ref e) | Event::Start(ref e) => {
                let kind = e.local_name();
                let Some(content) = attr(e, "ContentType") else {
                    continue;
                };
                match kind.as_ref() {
                    "Default" => {
                        if let Some(ext) = attr(e, "Extension") {
                            defaults.insert(ext.to_lowercase(), content);
                        }
                    }
                    "Override" => {
                        if let Some(part) = attr(e, "PartName") {
                            overrides.insert(part.trim_start_matches('/').to_owned(), content);
                        }
                    }
                    _ => {}
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    // Fold the defaults into the map as the parts are carried, so the writer
    // needs no extension table of its own.
    let mut out = overrides;
    for i in 0..zip.len() {
        let Ok(entry) = zip.by_index_raw(i) else {
            continue;
        };
        let path = entry.name().to_owned();
        if out.contains_key(&path) {
            continue;
        }
        if let Some(ext) = path.rsplit('.').next()
            && let Some(content) = defaults.get(&ext.to_lowercase())
        {
            out.insert(path, content.clone());
        }
    }
    out
}

/// Reads a part as bytes, refusing anything oversized.
fn read_bytes<R: Read + Seek>(zip: &mut zip::ZipArchive<R>, path: &str) -> Result<Vec<u8>> {
    super::zipxml::read_bytes(zip, path).map_err(Error::Xlsx)
}

/// Carries a part and everything it points at, so a drawing keeps its images
/// and a comment keeps the shapes that place it.
fn carry<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    roots: &[String],
    types: &HashMap<String, String>,
    out: &mut Vec<OpaquePart>,
) {
    let mut queue: Vec<String> = roots.to_vec();
    while let Some(path) = queue.pop() {
        if out.iter().any(|p| p.path == path) {
            continue;
        }
        let Ok(data) = read_bytes(zip, &path) else {
            continue;
        };
        out.push(OpaquePart {
            content_type: types.get(&path).cloned(),
            data,
            path: path.clone(),
        });
        // The part's own relationships come along, and so do their targets.
        let rels_path = rels_path_for(&path);
        let Ok(rels) = read_relationships(zip, &rels_path) else {
            continue;
        };
        if let Ok(data) = read_bytes(zip, &rels_path) {
            out.push(OpaquePart {
                path: rels_path.clone(),
                content_type: types.get(&rels_path).cloned(),
                data,
            });
        }
        let base = path.rsplit_once('/').map_or("", |(dir, _)| dir);
        for rel in rels.values() {
            if !rel.external {
                queue.push(resolve(base, &rel.target));
            }
        }
    }
}

/// Reads a part fully, refusing anything oversized.
fn read_part<R: Read + Seek>(zip: &mut zip::ZipArchive<R>, path: &str) -> Result<String> {
    super::zipxml::read_part(zip, path).map_err(Error::Xlsx)
}

/// The `.rels` path belonging to a part: `xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`.
pub(crate) fn rels_path_for(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((dir, file)) => format!("{dir}/_rels/{file}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// Resolves a relationship target against the directory of the part that
/// declared it. Absolute targets (`/xl/styles.xml`) are taken as package paths.
pub(crate) fn resolve(base: &str, target: &str) -> String {
    if let Some(abs) = target.strip_prefix('/') {
        return abs.to_owned();
    }
    if base.is_empty() {
        return target.to_owned();
    }
    // Targets such as "../sharedStrings.xml" walk up from the part's directory.
    let mut parts: Vec<&str> = base.split('/').collect();
    let mut rest = target;
    while let Some(up) = rest.strip_prefix("../") {
        parts.pop();
        rest = up;
    }
    if parts.is_empty() {
        rest.to_owned()
    } else {
        format!("{}/{}", parts.join("/"), rest)
    }
}

/// Parses a `.rels` part into id -> relationship.
pub(crate) fn read_relationships<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<HashMap<String, Relationship>> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out = HashMap::new();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(e) | Event::Start(e)) if e.local_name().as_ref() == "Relationship" => {
                let (mut id, mut kind, mut target) = (None, None, None);
                let mut external = false;
                for attr in e.attributes().flatten() {
                    let value = attr
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .into_owned();
                    match attr.key.local_name().as_ref() {
                        "Id" => id = Some(value),
                        "Type" => kind = Some(value),
                        "Target" => target = Some(value),
                        "TargetMode" => external = value == "External",
                        _ => {}
                    }
                }
                if let (Some(id), Some(kind), Some(target)) = (id, kind, target) {
                    out.insert(
                        id,
                        Relationship {
                            kind,
                            target,
                            external,
                        },
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Finds the workbook part through the package root relationships.
pub(crate) fn find_workbook_part<R: Read + Seek>(zip: &mut zip::ZipArchive<R>) -> Result<String> {
    let rels = read_relationships(zip, "_rels/.rels")?;
    rels.values()
        .find(|r| r.kind.ends_with("/officeDocument"))
        .map(|r| resolve("", &r.target))
        .ok_or_else(|| Error::Xlsx("package declares no office document".into()))
}

/// Reads the sheet names and their relationship ids in tab order, plus the
/// index of the tab the workbook was saved on.
fn read_sheet_list<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<WorkbookHeader> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out = Vec::new();
    let mut active = 0;
    let mut epoch = Epoch::Windows1900;
    let mut properties = Vec::new();
    let mut calculation = Vec::new();
    let mut view = Vec::new();
    let mut seen_view = false;
    let mut protection = WorkbookProtection::default();
    let mut external = Vec::new();
    let mut pivot_caches: Vec<(u32, String)> = Vec::new();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(e) | Event::Start(e)) if e.local_name().as_ref() == "workbookPr" => {
                // The one attribute with a meaning here is the base date; a
                // workbook written on a Mac counts from 1904, and reading it as
                // 1900 moves every date in it by 1462 days.
                for (key, value) in attrs(&e) {
                    if key == "date1904" {
                        if is_true(&value) {
                            epoch = Epoch::Mac1904;
                        }
                    } else {
                        properties.push((key, value));
                    }
                }
            }
            Ok(Event::Empty(e) | Event::Start(e)) if e.local_name().as_ref() == "calcPr" => {
                calculation = attrs(&e);
            }
            Ok(Event::Empty(e) | Event::Start(e))
                if e.local_name().as_ref() == "workbookProtection" =>
            {
                for (name, slot) in protection.lock_slots() {
                    // `false` is spelled out here rather than left off, so the
                    // string is what decides, not the attribute's presence.
                    *slot = attr(&e, name).map(|v| is_true(&v));
                }
                protection.workbook_password =
                    PasswordHash::from_attrs(PasswordAttrs::WORKBOOK, |name| attr(&e, name));
                protection.revisions_password =
                    PasswordHash::from_attrs(PasswordAttrs::REVISIONS, |name| attr(&e, name));
            }
            // A workbook with several windows has a view for each; the first
            // is the one Excel opens with.
            Ok(Event::Empty(e) | Event::Start(e))
                if e.local_name().as_ref() == "workbookView" && !seen_view =>
            {
                seen_view = true;
                view = workbook_view(&e, &mut active);
            }
            Ok(Event::Empty(e) | Event::Start(e)) if e.local_name().as_ref() == "pivotCache" => {
                if let (Some(id), Some(rel)) = (
                    attr(&e, "cacheId").and_then(|v| v.parse().ok()),
                    attr(&e, "id"),
                ) {
                    pivot_caches.push((id, rel));
                }
            }
            Ok(Event::Empty(e) | Event::Start(e))
                if e.local_name().as_ref() == "externalReference" =>
            {
                if let Some(id) = attr(&e, "id") {
                    external.push(id);
                }
            }
            Ok(Event::Empty(e) | Event::Start(e)) if e.local_name().as_ref() == "sheet" => {
                let (mut name, mut rid) = (None, None);
                let mut state = SheetVisibility::Visible;
                for attr in e.attributes().flatten() {
                    let value = attr
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default()
                        .into_owned();
                    // The id attribute is namespaced (r:id), so match the local
                    // part rather than the raw key.
                    match attr.key.local_name().as_ref() {
                        "name" => name = Some(value),
                        "id" => rid = Some(value),
                        "state" => state = SheetVisibility::parse(&value),
                        _ => {}
                    }
                }
                if let (Some(name), Some(rid)) = (name, rid) {
                    out.push((name, rid, state));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(WorkbookHeader {
        sheets: out,
        active,
        epoch,
        properties,
        calculation,
        view,
        protection,
        pivot_caches,
        external,
    })
}

/// The first `<workbookView>`: the active tab, and the attributes carried
/// as they are - all but the window's geometry.
fn workbook_view(
    e: &quick_xml::events::BytesStart<'_>,
    active: &mut usize,
) -> Vec<(String, String)> {
    let mut view = Vec::new();
    for (key, value) in attrs(e) {
        match key.as_str() {
            "activeTab" => *active = value.parse().unwrap_or(0),
            "xWindow" | "yWindow" | "windowWidth" | "windowHeight" => {}
            _ => view.push((key, value)),
        }
    }
    view
}

/// What the workbook part says before its sheets are read.
struct WorkbookHeader {
    /// Sheet names and the relationship ids their parts hang off.
    sheets: Vec<(String, String, SheetVisibility)>,
    /// Index of the tab the workbook was saved on.
    active: usize,
    /// The base date its serial numbers count from.
    epoch: Epoch,
    /// The rest of `<workbookPr>`.
    properties: Vec<(String, String)>,
    /// The attributes of `<calcPr>`.
    calculation: Vec<(String, String)>,
    /// The rest of the first `<workbookView>`.
    view: Vec<(String, String)>,
    /// What `<workbookProtection>` locks.
    protection: WorkbookProtection,
    /// The pivot caches the workbook declares: the id a report points at, and
    /// the relationship its definition hangs off.
    pivot_caches: Vec<(u32, String)>,
    /// Relationship ids of `<externalReference>`, in the order written: the
    /// first is what a formula spells `[1]`.
    external: Vec<String>,
}

/// The pivot reports of one sheet, in name order.
///
/// Whether each note shows and how big its box is: that lives in the VML.
fn read_note_boxes<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    links: &HashMap<String, Relationship>,
    base: &str,
    comments: &mut BTreeMap<CellRef, Comment>,
) {
    let Some(rel) = links
        .values()
        .find(|r| !r.external && r.kind.ends_with("/vmlDrawing"))
    else {
        return;
    };
    let Ok(vml) = read_part(zip, &resolve(base, &rel.target)) else {
        return;
    };
    for shape in super::vml::note_shapes(&vml) {
        if let Some(note) = shape.cell.and_then(|at| comments.get_mut(&at)) {
            note.visible = shape.visible;
            // Excel's own size is what `None` means.
            note.size = shape.size.filter(|&s| s != (108.0, 59.25));
        }
    }
}

/// They are read as well as carried: nothing writes them back, so the parts
/// still travel whole and this is a view of them.
fn read_sheet_pivots<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    links: &HashMap<String, Relationship>,
    base: &str,
) -> Vec<PivotTable> {
    let mut out: Vec<PivotTable> = links
        .values()
        .filter(|r| !r.external && r.kind.ends_with("/pivotTable"))
        .filter_map(|rel| {
            let part = resolve(base, &rel.target);
            let mut table = read_pivot_table(zip, &part).ok()?;
            table.origin = Some(PivotOrigin {
                part,
                read: Box::new(table.clone()),
                relaid: false,
            });
            Some(table)
        })
        .collect();
    out.sort_by(|a, b| a.name.cmp(&b.name));
    out
}

/// The caches the reports read from, in the order the workbook lists them.
fn read_pivot_caches<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    declared: &[(u32, String)],
    rels: &HashMap<String, Relationship>,
    base: &str,
) -> Vec<PivotCache> {
    let mut out = Vec::new();
    for (id, rel_id) in declared {
        let Some(rel) = rels.get(rel_id) else {
            continue;
        };
        let path = resolve(base, &rel.target);
        if let Ok(mut cache) = read_pivot_cache(zip, &path) {
            cache.id = *id;
            cache.definition_part = path;
            cache.origin = Some(Box::new(cache.clone()));
            out.push(cache);
        }
    }
    out
}

/// The drawing parts of one sheet with their text, in path order.
fn sheet_drawings<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    links: &HashMap<String, Relationship>,
    base: &str,
) -> Vec<(String, String)> {
    let mut paths: Vec<String> = links
        .values()
        .filter(|r| !r.external && r.kind.ends_with("/drawing"))
        .map(|r| resolve(base, &r.target))
        .collect();
    paths.sort();
    paths
        .into_iter()
        .filter_map(|path| read_part(zip, &path).ok().map(|xml| (path, xml)))
        .collect()
}

/// The charts drawn on one sheet, classic and 2016, in drawing order.
fn read_sheet_charts<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    drawings: &[(String, String)],
) -> (Vec<Chart>, Vec<ChartEx>) {
    let mut charts = Vec::new();
    let mut extended = Vec::new();
    for (drawing, xml) in drawings {
        let rels = read_relationships(zip, &rels_path_for(drawing)).unwrap_or_default();
        let dir = drawing.rsplit_once('/').map_or("", |(dir, _)| dir);
        let objects = super::chart::scan_drawing(xml);
        for object in objects {
            let Some(anchor) = object.anchor else {
                continue;
            };
            for frame in object.frames {
                let Some(rel) = rels.get(&frame.rel) else {
                    continue;
                };
                let part = resolve(dir, &rel.target);
                let Ok(body) = read_part(zip, &part) else {
                    continue;
                };
                if frame.extended {
                    if let Some(mut chart) = super::chart::read_chart_ex(&body) {
                        chart.name = frame.name;
                        chart.anchor = anchor;
                        chart.part = part;
                        chart.origin = Some(ChartExOrigin {
                            drawing: drawing.clone(),
                            id: frame.id,
                            grouped: frame.grouped,
                            read: Box::new(chart.clone()),
                        });
                        extended.push(chart);
                    }
                } else if let Some((mut chart, root_attributes, prefix)) =
                    super::chart::read_chart(&body)
                {
                    chart.name = frame.name;
                    chart.anchor = anchor;
                    chart.origin = Some(ChartOrigin {
                        drawing: drawing.clone(),
                        part,
                        read: Box::new(chart.clone()),
                        root_attributes,
                        prefix,
                        grouped: frame.grouped,
                    });
                    charts.push(chart);
                }
            }
        }
    }
    (charts, extended)
}

/// The shapes drawn on one sheet, in drawing order.
fn read_sheet_shapes(
    drawings: &[(String, String)],
    theme: Option<&str>,
) -> Vec<crate::model::shape::Shape> {
    use crate::model::shape::{Shape, ShapeOrigin};
    let mut shapes = Vec::new();
    for (drawing, xml) in drawings {
        if !xml.contains("sp") {
            continue;
        }
        let first = shapes.len();
        for object in super::shape::scan_shapes(xml, theme) {
            let Some(anchor) = object.anchor else {
                continue;
            };
            for element in object.shapes {
                shapes.push(Shape {
                    origin: Some(ShapeOrigin {
                        drawing: drawing.clone(),
                        id: element.id,
                        grouped: element.grouped,
                        name: element.name.clone(),
                        description: element.description.clone(),
                        anchor,
                        rotation: element.rotation,
                        flip_h: element.flip_h,
                        flip_v: element.flip_v,
                        geometry: element.geometry.clone(),
                        format: element.format.clone(),
                        text: element.text.clone(),
                        font: element.font.clone(),
                        style: element.style,
                        read_from_drawing: 0,
                    }),
                    name: element.name,
                    description: element.description,
                    anchor,
                    rotation: element.rotation,
                    flip_h: element.flip_h,
                    flip_v: element.flip_v,
                    geometry: element.geometry,
                    format: element.format,
                    text: element.text,
                    font: element.font,
                });
            }
        }
        let read = shapes.len() - first;
        for shape in &mut shapes[first..] {
            if let Some(origin) = &mut shape.origin {
                origin.read_from_drawing = read;
            }
        }
    }
    shapes
}

/// The pictures drawn on one sheet, in drawing order. A picture linked to a
/// file outside the package has no bytes to model and stays in the drawing
/// as written.
fn read_sheet_images<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    drawings: &[(String, String)],
) -> Vec<crate::model::image::Image> {
    use crate::model::image::{Image, ImageFormat, ImageOrigin, hash};
    let mut images = Vec::new();
    for (drawing, xml) in drawings {
        if !xml.contains("pic") {
            continue;
        }
        let rels = read_relationships(zip, &rels_path_for(drawing)).unwrap_or_default();
        let dir = drawing.rsplit_once('/').map_or("", |(dir, _)| dir);
        for object in super::image::scan_pictures(xml) {
            let Some(anchor) = object.anchor else {
                continue;
            };
            for picture in object.pictures {
                let Some(rel) = picture.rel.as_ref().and_then(|id| rels.get(id)) else {
                    continue;
                };
                if rel.external {
                    continue;
                }
                let part = resolve(dir, &rel.target);
                let Ok(data) = read_bytes(zip, &part) else {
                    continue;
                };
                let Some(format) =
                    ImageFormat::sniff(&data).or_else(|| ImageFormat::from_extension(&part))
                else {
                    continue;
                };
                images.push(Image {
                    origin: Some(ImageOrigin {
                        drawing: drawing.clone(),
                        id: picture.id,
                        grouped: picture.grouped,
                        name: picture.name.clone(),
                        description: picture.description.clone(),
                        anchor,
                        data_len: data.len(),
                        data_hash: hash(&data),
                    }),
                    name: picture.name,
                    description: picture.description,
                    anchor,
                    format,
                    data,
                });
            }
        }
    }
    images
}

/// The tables of one sheet, in the order their parts are related.
///
/// Unlike the pivot reports beside them these are written back from the model,
/// so their parts are dropped from the attachments the sheet carries.
fn read_sheet_tables<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    links: &HashMap<String, Relationship>,
    base: &str,
) -> Vec<Table> {
    let mut out: Vec<Table> = links
        .values()
        .filter(|r| !r.external && r.kind.ends_with("/table"))
        .filter_map(|rel| read_table(zip, &resolve(base, &rel.target)).ok())
        .collect();
    out.sort_by_key(|t| t.id);
    out
}

/// Reads one `xl/tables/tableN.xml`.
fn read_table<R: Read + Seek>(zip: &mut zip::ZipArchive<R>, path: &str) -> Result<Table> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out = Table {
        id: 0,
        name: String::new(),
        display_name: String::new(),
        range: Range::parse("A1")?,
        header_row_count: None,
        totals_row_count: None,
        auto_filter: None,
        sort_state: None,
        columns: Vec::new(),
        style: None,
        extensions: trailing_extensions(&xml),
    };
    let mut seen = false;
    let mut filter_col = None;
    let mut in_filter = false;
    loop {
        let event = reader.read_event();
        if let Ok(Event::Start(ref e)) = event
            && e.local_name().as_ref() == "autoFilter"
        {
            in_filter = true;
        }
        match event {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                // Last in the schema; its `x14:table` shares the local name
                // and would overwrite the table with empty attributes.
                "extLst" => break,
                "table" => {
                    seen = true;
                    out.id = attr(e, "id").and_then(|v| v.parse().ok()).unwrap_or(0);
                    out.name = attr(e, "name").unwrap_or_default();
                    // A file may name only one of the two; Excel treats the
                    // display name as the one formulas use, so it wins.
                    out.display_name = attr(e, "displayName")
                        .or_else(|| attr(e, "name"))
                        .unwrap_or_default();
                    if out.name.is_empty() {
                        out.name.clone_from(&out.display_name);
                    }
                    if let Some(r) = attr(e, "ref").and_then(|r| Range::parse(&r).ok()) {
                        out.range = r;
                    }
                    out.header_row_count = attr(e, "headerRowCount").and_then(|v| v.parse().ok());
                    out.totals_row_count = attr(e, "totalsRowCount").and_then(|v| v.parse().ok());
                }
                "tableColumn" => out.columns.push(TableColumn {
                    id: attr(e, "id").and_then(|v| v.parse().ok()).unwrap_or(0),
                    name: attr(e, "name").unwrap_or_default(),
                    totals_row_function: attr(e, "totalsRowFunction"),
                    totals_row_label: attr(e, "totalsRowLabel"),
                    calculated_formula: None,
                }),
                "calculatedColumnFormula" => {
                    // The text arrives as the next event; the column it belongs
                    // to is the one being built.
                    if let Ok(Event::Text(t)) = reader.read_event()
                        && let Some(column) = out.columns.last_mut()
                    {
                        column.calculated_formula = Some(t.xml10_content().into_owned());
                    }
                }
                "tableStyleInfo" => {
                    out.style = Some(TableStyle {
                        name: attr(e, "name"),
                        show_first_column: attr(e, "showFirstColumn").is_some_and(|v| is_true(&v)),
                        show_last_column: attr(e, "showLastColumn").is_some_and(|v| is_true(&v)),
                        show_row_stripes: attr(e, "showRowStripes").is_some_and(|v| is_true(&v)),
                        show_column_stripes: attr(e, "showColumnStripes")
                            .is_some_and(|v| is_true(&v)),
                    });
                }
                name => {
                    let sort = sort_target(&mut out.sort_state, &mut out.auto_filter, in_filter);
                    let _ = read_sort_element(sort, name, e)
                        || read_filter_element(&mut out.auto_filter, &mut filter_col, name, e);
                }
            },
            Ok(Event::End(ref e)) => match e.local_name().as_ref() {
                "filterColumn" => filter_col = None,
                "autoFilter" => in_filter = false,
                _ => {}
            },
            Ok(Event::Eof) | Err(_) => break,
            Ok(_) => {}
        }
    }
    if seen {
        Ok(out)
    } else {
        Err(Error::Xlsx(format!("{path} is not a table part")))
    }
}

/// Reads a pivot table definition: the report's shape, not its data.
fn read_pivot_table<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<PivotTable> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out = PivotTable::default();
    // Which list of field indexes is being read: the same `<field x="n"/>`
    // element serves rows, columns and page filters.
    let mut axis: Option<PivotAxis> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "pivotTableDefinition" => {
                    out.name = attr(e, "name").unwrap_or_default();
                    out.cache_id = attr(e, "cacheId").and_then(|v| v.parse().ok()).unwrap_or(0);
                    out.row_grand_totals = attr(e, "rowGrandTotals").is_none_or(|v| is_true(&v));
                    out.column_grand_totals = attr(e, "colGrandTotals").is_none_or(|v| is_true(&v));
                }
                "location" => {
                    out.location = attr(e, "ref").and_then(|r| Range::parse(&r).ok());
                    let count = |name: &str| attr(e, name).and_then(|v| v.parse().ok());
                    out.first_data_row = count("firstDataRow").unwrap_or(1);
                    out.first_data_col = count("firstDataCol").unwrap_or(1);
                }
                "pivotField" => out.fields.push(PivotField {
                    axis: attr(e, "axis")
                        .map(|a| PivotAxis::parse(&a))
                        .unwrap_or_default(),
                    data_field: attr(e, "dataField").is_some_and(|v| is_true(&v)),
                    default_subtotal: attr(e, "defaultSubtotal").is_none_or(|v| is_true(&v)),
                    show_all: attr(e, "showAll").is_some_and(|v| is_true(&v)),
                }),
                "rowFields" => axis = Some(PivotAxis::Row),
                "colFields" => axis = Some(PivotAxis::Column),
                "pageFields" => axis = Some(PivotAxis::Page),
                // A page filter names its field with `fld`, the other two axes
                // with `x`; `-2` in the latter is the values field.
                "field" | "pageField" => {
                    let index = attr(e, "x")
                        .or_else(|| attr(e, "fld"))
                        .and_then(|v| v.parse::<i32>().ok());
                    if let (Some(index), Some(axis)) = (index, axis) {
                        match axis {
                            PivotAxis::Row => out.row_fields.push(index),
                            PivotAxis::Column => out.column_fields.push(index),
                            PivotAxis::Page => out.page_fields.push(index),
                            _ => {}
                        }
                    }
                }
                "dataField" => out.data_fields.push(DataField {
                    name: attr(e, "name"),
                    field: attr(e, "fld").and_then(|v| v.parse().ok()).unwrap_or(0),
                    subtotal: attr(e, "subtotal")
                        .map(|v| Subtotal::parse(&v))
                        .unwrap_or_default(),
                    number_format: attr(e, "numFmtId").and_then(|v| v.parse().ok()),
                }),
                "pivotTableStyleInfo" => {
                    out.style = PivotStyleInfo {
                        name: attr(e, "name"),
                        show_row_headers: attr(e, "showRowHeaders").is_some_and(|v| is_true(&v)),
                        show_column_headers: attr(e, "showColHeaders").is_some_and(|v| is_true(&v)),
                        show_row_stripes: attr(e, "showRowStripes").is_some_and(|v| is_true(&v)),
                        show_column_stripes: attr(e, "showColStripes").is_some_and(|v| is_true(&v)),
                    };
                }
                _ => {}
            },
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                "rowFields" | "colFields" | "pageFields" => axis = None,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Reads a pivot cache definition: where the data came from and what its
/// columns are.
fn read_pivot_cache<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<PivotCache> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out = PivotCache::default();
    // Shared items belong to the field being read; they are only listed when
    // the cache was saved with its data.
    let mut in_shared = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "worksheetSource" => {
                    out.source = CacheSource {
                        sheet: attr(e, "sheet"),
                        range: attr(e, "ref").and_then(|r| Range::parse(&r).ok()),
                        name: attr(e, "name"),
                    };
                }
                "cacheField" => out.fields.push(CacheField {
                    name: attr(e, "name").unwrap_or_default(),
                    number_format: attr(e, "numFmtId").and_then(|v| v.parse().ok()),
                    shared_items: Vec::new(),
                }),
                "sharedItems" => in_shared = true,
                // Inside `<sharedItems>` the tag says the type and `v` the
                // value: `s` text, `n` number, `b` boolean, `d` date, `m` a
                // blank, which has no value at all.
                "s" | "n" | "b" | "d" | "e" if in_shared => {
                    if let (Some(field), Some(value)) = (out.fields.last_mut(), attr(e, "v")) {
                        field.shared_items.push(value);
                    }
                }
                _ => {}
            },
            Ok(Event::End(e)) if e.local_name().as_ref() == "sharedItems" => in_shared = false,
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// One element of a part, from its opening tag to its closing one, as it
/// stands in the file. For the parts of the stylesheet this crate carries
/// rather than models.
fn element_of(xml: &str, name: &str) -> Option<String> {
    let open = xml.find(&format!("<{name}"))?;
    let close = format!("</{name}>");
    if let Some(end) = xml[open..].find(&close) {
        return Some(xml[open..open + end + close.len()].to_owned());
    }
    // An empty element: `<tableStyles count="0"/>`.
    let end = xml[open..].find("/>")? + 2;
    Some(xml[open..open + end].to_owned())
}

/// The sheet's own `<extLst>`, as it stands in the file.
///
/// The element is the last child of `<worksheet>`, and the ones that appear
/// earlier belong to a rule or a validation inside the sheet, so the last
/// opening tag is the one wanted. Nothing here looks inside it.
fn trailing_extensions(xml: &str) -> Option<String> {
    let open = xml.rfind("<extLst")?;
    let close = xml[open..].find("</extLst>")? + open + "</extLst>".len();
    Some(self_contained(&xml[open..close], &root_namespaces(xml)))
}

/// The `xmlns:p` declarations on the root element of a part.
fn root_namespaces(xml: &str) -> Vec<(String, String)> {
    let mut reader = Reader::from_str(xml);
    loop {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => return namespaces_of(&e),
            Ok(Event::Eof) | Err(_) => return Vec::new(),
            Ok(_) => {}
        }
    }
}

/// The `xmlns:p` declarations an element makes, as (prefix, uri).
fn namespaces_of(e: &quick_xml::events::BytesStart<'_>) -> Vec<(String, String)> {
    e.attributes()
        .flatten()
        .filter_map(|a| {
            let prefix = a.key.as_ref().strip_prefix("xmlns:")?.to_owned();
            let uri = a
                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                .ok()?
                .into_owned();
            Some((prefix, uri))
        })
        .collect()
}

/// An element cut out of its part - an `<extLst>`, `<tableStyles>` -
/// declaring on itself the prefixes it borrowed from the part's root.
///
/// Excel 365 puts `xr2:uid` on a sparkline group and declares `xr2` only on
/// `<worksheet>`. Our root declares its own namespaces, not the file's, so the
/// element written back as it was read would use an undeclared prefix.
fn self_contained(ext: &str, root: &[(String, String)]) -> String {
    let mut borrowed = String::new();
    for (prefix, uri) in root {
        let used = ext.contains(&format!("<{prefix}:")) || ext.contains(&format!(" {prefix}:"));
        if used && !ext.contains(&format!("xmlns:{prefix}=")) {
            let _ = std::fmt::Write::write_fmt(
                &mut borrowed,
                format_args!(r#" xmlns:{prefix}="{}""#, escape_attr(uri)),
            );
        }
    }
    if borrowed.is_empty() {
        return ext.to_owned();
    }
    // Just past the element's name.
    let at = ext
        .find(|c: char| c.is_ascii_whitespace() || c == '>' || c == '/')
        .unwrap_or(ext.len());
    format!("{}{borrowed}{}", &ext[..at], &ext[at..])
}

/// A namespace URI back in an attribute. They never hold markup in practice;
/// this keeps a hostile one from breaking the element.
fn escape_attr(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
}

/// Reads a sheet's notes: the author table, then the notes that point into it.
fn read_comments<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<BTreeMap<CellRef, Comment>> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut authors: Vec<String> = Vec::new();
    let mut out = BTreeMap::new();
    let mut in_authors = false;
    let mut in_author = false;
    let mut author = String::new();
    // The note being read: where it sits, whose it is, and its runs.
    let mut at: Option<CellRef> = None;
    let mut author_id = 0usize;
    let mut runs: Vec<TextRun> = Vec::new();
    let mut run = TextRun::default();
    let mut in_text = false;
    let mut in_rpr = false;
    let mut saw_run = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "authors" => in_authors = true,
                "author" if in_authors => {
                    in_author = true;
                    author.clear();
                }
                "comment" => {
                    at = attr(e, "ref").and_then(|r| CellRef::parse(&r).ok());
                    author_id = attr(e, "authorId")
                        .and_then(|id| id.parse().ok())
                        .unwrap_or(0);
                    runs.clear();
                    run = TextRun::default();
                    saw_run = false;
                }
                "r" => {
                    run = TextRun::default();
                    saw_run = true;
                }
                "rPr" => {
                    in_rpr = true;
                    run.font = Some(DiffFont::default());
                }
                "t" => in_text = true,
                name if in_rpr => apply_run_font(run.font.as_mut(), name, e),
                _ => {}
            },
            Ok(Event::Text(t)) if in_author => author.push_str(&t.xml10_content()),
            Ok(Event::Text(t)) if in_text => run.text.push_str(&t.borrow().into_inner()),
            Ok(Event::GeneralRef(r)) if in_author => push_entity(&mut author, &r),
            Ok(Event::GeneralRef(r)) if in_text => push_entity(&mut run.text, &r),
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                "authors" => in_authors = false,
                "author" => {
                    in_author = false;
                    authors.push(std::mem::take(&mut author));
                }
                "r" => runs.push(std::mem::take(&mut run)),
                "rPr" => in_rpr = false,
                "t" => in_text = false,
                "comment" => {
                    // A note written without `<r>` is one plain run.
                    if !saw_run {
                        runs.push(std::mem::take(&mut run));
                    }
                    if let Some(at) = at.take() {
                        out.insert(
                            at,
                            Comment {
                                author: authors.get(author_id).cloned().unwrap_or_default(),
                                text: std::mem::take(&mut runs),
                                ..Comment::default()
                            },
                        );
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Reads the shared string pool.
///
/// A `<si>` may be a single `<t>` or a run of `<r><t>` fragments with different
/// formatting; the text is the concatenation either way. Formatting runs are
/// rich text, which belongs to a later phase.
///
/// Each entry is kept as the value a cell holding it gets, so a cell takes a
/// clone of a shared `Arc<str>` rather than a copy of the text: a sheet of
/// seven million text cells over eight hundred thousand distinct strings
/// held 728 MB of copies.
fn read_shared_strings<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<CellValue>> {
    // Streamed as it inflates: the table of a large export is a quarter of a
    // gigabyte of XML, and none of it is needed once its strings are taken.
    let part = super::zipxml::open_part(zip, path).map_err(Error::Xlsx)?;
    let mut reader = Reader::from_reader(std::io::BufReader::with_capacity(1 << 16, part));
    let mut buf = Vec::new();
    let mut out: Vec<CellValue> = Vec::new();
    let mut runs: Vec<TextRun> = Vec::new();
    let mut run = TextRun::default();
    let mut in_si = false;
    let mut in_text = false;
    let mut in_rpr = false;
    loop {
        buf.clear();
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e) | Event::Empty(ref e)) => match e.local_name().as_ref() {
                "si" => {
                    in_si = true;
                    runs.clear();
                    run = TextRun::default();
                }
                "r" => run = TextRun::default(),
                "rPr" => {
                    in_rpr = true;
                    run.font = Some(DiffFont::default());
                }
                "t" => in_text = true,
                name if in_rpr => apply_run_font(run.font.as_mut(), name, e),
                _ => {}
            },
            Ok(Event::Text(t)) if in_si && in_text => {
                // Raw content, not `xml10_content`: XML line-ending
                // normalization would fold the CR of a CRLF typed into a cell
                // into the LF alone, and Excel keeps both.
                run.text.push_str(&t.borrow().into_inner());
            }
            Ok(Event::GeneralRef(r)) if in_si && in_text => push_entity(&mut run.text, &r),
            Ok(Event::End(e)) => match e.local_name().as_ref() {
                "si" => {
                    in_si = false;
                    // A string with no `<r>` at all is one plain run.
                    if runs.is_empty() {
                        runs.push(std::mem::take(&mut run));
                    }
                    out.push(pooled(&runs));
                    runs.clear();
                }
                "r" => runs.push(std::mem::take(&mut run)),
                "rPr" => in_rpr = false,
                "t" => in_text = false,
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Applies one child of an `<rPr>`: the formatting of a single run.
///
/// `rPr` states only what the run changes, which is the same shape as a
/// differential format, so the same type holds it.
fn apply_run_font(font: Option<&mut DiffFont>, name: &str, e: &quick_xml::events::BytesStart<'_>) {
    let Some(font) = font else { return };
    match name {
        "b" => font.bold = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
        "i" => font.italic = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
        "strike" => font.strike = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
        "u" => {
            font.underline =
                Some(attr(e, "val").map_or(Underline::Single, |v| Underline::parse(&v)));
        }
        "sz" => {
            font.size = attr(e, "val")
                .and_then(|v| v.parse::<f64>().ok())
                .map(points_to_hundredths);
        }
        "color" => font.color = Some(read_color(e)),
        "rFont" => font.name = attr(e, "val"),
        "charset" => font.charset = attr(e, "val").and_then(|v| v.parse().ok()),
        "family" => font.family = attr(e, "val").and_then(|v| v.parse().ok()),
        "scheme" => font.scheme = attr(e, "val").and_then(|v| FontScheme::parse(&v)),
        "vertAlign" => {
            font.script = Some(match attr(e, "val").as_deref() {
                Some("superscript") => Script::Superscript,
                Some("subscript") => Script::Subscript,
                _ => Script::Baseline,
            });
        }
        // `family` says which font to substitute when the named one is
        // missing; nothing here chooses fonts.
        _ => {}
    }
}

/// Reads the style table.
///
/// xlsx keeps fonts, fills and borders in tables of their own and has each
/// `cellXfs` entry point into them by index. This assembles them back into
/// whole [`Style`] values, one per `cellXfs` position.
fn read_styles<R: Read + Seek>(zip: &mut zip::ZipArchive<R>, path: &str) -> Result<StyleTable> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut state = StylesReader::default();

    loop {
        let event = match reader.read_event() {
            Ok(e) => e,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
        };
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                state.start(e, matches!(event, Event::Empty(_)));
            }
            Event::End(ref e) => state.end(e.local_name().as_ref()),
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(state.finish())
}

/// The component tables of `styles.xml` as they are being filled in.
///
/// Cells do not carry a style; they carry an index into `cellXfs`, whose entry
/// indexes in turn into the font, fill and border tables. So the part has to be
/// read whole before a single [`Style`] can be assembled, which is what
/// [`StylesReader::finish`] does.
#[derive(Default)]
struct StylesReader {
    custom_formats: HashMap<u16, String>,
    fonts: Vec<Font>,
    fills: Vec<Fill>,
    borders: Vec<Borders>,
    xfs: Vec<Xf>,
    dxfs: Vec<DifferentialStyle>,
    /// Which table we are inside; the same element names appear in several, and
    /// `cellStyleXfs` holds `xf` elements that cells never index into.
    section: Section,
    side: BorderSide,
    /// How deep inside `<extLst>` the reader is. The extension travels whole
    /// in `Spreadsheet::style_extensions`, and what it holds repeats the
    /// part's own names: the `<x14:dxfs>` of slicer styles, read as the
    /// part's `<dxfs>`, were appended to the book's differential formats and
    /// written out a second time on every save.
    ext_depth: u32,
}

impl StylesReader {
    fn start(&mut self, e: &quick_xml::events::BytesStart<'_>, empty: bool) {
        let name = e.local_name();
        let name = name.as_ref();
        if name == "extLst" && !empty {
            self.ext_depth += 1;
        }
        if self.ext_depth > 0 {
            return;
        }
        if self.start_section(name) || self.start_number_format(name, e) {
            return;
        }
        if self.section == Section::Dxfs {
            self.start_dxf(name, e, empty);
        } else {
            self.start_component(name, e, empty);
        }
    }

    /// The elements that open a table, and so say what the ones inside mean.
    fn start_section(&mut self, name: &str) -> bool {
        self.section = match name {
            "fonts" => Section::Fonts,
            "fills" => Section::Fills,
            "borders" => Section::Borders,
            "cellXfs" => Section::CellXfs,
            "cellStyleXfs" => Section::Other,
            "dxfs" => Section::Dxfs,
            _ => return false,
        };
        true
    }

    /// A number format, which reads differently inside a `<dxf>`: there it is
    /// the format itself rather than an entry in the book's table of them.
    fn start_number_format(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>) -> bool {
        if name != "numFmt" {
            return false;
        }
        if self.section == Section::Dxfs {
            if let Some(dxf) = self.dxfs.last_mut() {
                dxf.number_format = Some(match attr(e, "formatCode") {
                    Some(code) => NumberFormat::Custom(code),
                    None => attr(e, "numFmtId")
                        .and_then(|v| v.parse().ok())
                        .map_or(NumberFormat::General, NumberFormat::Builtin),
                });
            }
            return true;
        }
        let (mut id, mut code) = (None, None);
        for (key, value) in attrs(e) {
            match key.as_str() {
                "numFmtId" => id = value.parse::<u16>().ok(),
                "formatCode" => code = Some(value),
                _ => {}
            }
        }
        if let (Some(id), Some(code)) = (id, code) {
            self.custom_formats.insert(id, code);
        }
        true
    }

    /// One element inside a `<dxf>`, the partial style of a conditional format.
    /// Each part is `Option`: a dxf says *what to replace*, so an absent font
    /// is not a default font.
    fn start_dxf(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>, empty: bool) {
        if name == "dxf" {
            self.dxfs.push(DifferentialStyle::default());
            return;
        }
        let Some(dxf) = self.dxfs.last_mut() else {
            self.side = apply_dxf_child(None, self.side, name, e, empty);
            return;
        };
        match name {
            "font" => dxf.font = Some(DiffFont::default()),
            "fill" => dxf.fill = Some(DiffFill::default()),
            "border" => {
                dxf.borders = Some(DiffBorders::default());
                self.side = BorderSide::None;
            }
            "alignment" => dxf.alignment = Some(read_alignment(e)),
            "protection" => dxf.protection = Some(read_protection(e)),
            _ => self.side = apply_dxf_child(Some(dxf), self.side, name, e, empty),
        }
    }

    /// One element of a component table or of `cellXfs`; the names overlap
    /// between the tables, so the open section decides which one this is.
    fn start_component(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>, empty: bool) {
        match name {
            "dxf" => self.dxfs.push(DifferentialStyle::default()),
            "font" => self.fonts.push(Font::default()),
            "fill" => self.fills.push(Fill::default()),
            "border" => {
                self.borders.push(Borders {
                    diagonal_direction: diagonal_direction(e),
                    ..Borders::default()
                });
                self.side = BorderSide::None;
            }
            "xf" if self.section == Section::CellXfs => self.xfs.push(read_xf(e)),
            "alignment" if self.section == Section::CellXfs => {
                if let Some(xf) = self.xfs.last_mut() {
                    xf.alignment = read_alignment(e);
                }
            }
            "protection" if self.section == Section::CellXfs => {
                if let Some(xf) = self.xfs.last_mut() {
                    xf.protection = read_protection(e);
                }
            }
            _ => match self.section {
                Section::Fonts => {
                    if let Some(font) = self.fonts.last_mut() {
                        apply_font_child(font, name, e);
                    }
                }
                Section::Fills => {
                    if let Some(fill) = self.fills.last_mut() {
                        apply_fill_child(fill, name, e);
                    }
                }
                Section::Borders => {
                    self.side =
                        apply_border_child(self.borders.last_mut(), self.side, name, e, empty);
                }
                _ => {}
            },
        }
    }

    fn end(&mut self, name: &str) {
        if name == "extLst" {
            self.ext_depth = self.ext_depth.saturating_sub(1);
            return;
        }
        if self.ext_depth > 0 {
            return;
        }
        match name {
            "fonts" | "fills" | "borders" | "cellXfs" | "cellStyleXfs" | "dxfs" => {
                self.section = Section::None;
            }
            "left" | "right" | "top" | "bottom" | "diagonal" => self.side = BorderSide::None,
            _ => {}
        }
    }

    /// Resolves every `cellXfs` entry against the component tables.
    fn finish(self) -> StyleTable {
        let styles = self
            .xfs
            .into_iter()
            .map(|xf| Style {
                number_format: match (xf.number_format, self.custom_formats.get(&xf.number_format))
                {
                    (_, Some(code)) => NumberFormat::Custom(code.clone()),
                    (0, None) => NumberFormat::General,
                    (id, None) => NumberFormat::Builtin(id),
                },
                font: self
                    .fonts
                    .get(xf.font as usize)
                    .cloned()
                    .unwrap_or_default(),
                fill: self
                    .fills
                    .get(xf.fill as usize)
                    .cloned()
                    .unwrap_or_default(),
                borders: self
                    .borders
                    .get(xf.border as usize)
                    .cloned()
                    .unwrap_or_default(),
                alignment: xf.alignment,
                protection: xf.protection,
            })
            .collect();
        let mut table = StyleTable::from_styles(styles);
        table.differential = self.dxfs;
        table
    }
}

/// Applies one child element of a `<dxf>`.
///
/// A differential format states only what it changes, so each piece lands in
/// its own `Option` rather than filling in a whole component.
fn apply_dxf_child(
    dxf: Option<&mut DifferentialStyle>,
    side: BorderSide,
    name: &str,
    e: &quick_xml::events::BytesStart<'_>,
    empty: bool,
) -> BorderSide {
    let Some(dxf) = dxf else { return side };
    // Inside a border the element names are the edges, and the colour that
    // follows belongs to whichever edge was opened last.
    if dxf.borders.is_some() && matches!(name, "left" | "right" | "top" | "bottom" | "diagonal") {
        let opened = BorderSide::parse(name);
        if let Some(borders) = dxf.borders.as_mut()
            && let Some(slot) = diff_side(borders, opened)
        {
            *slot = Some(Border {
                style: attr(e, "style")
                    .map(|v| BorderStyle::parse(&v))
                    .unwrap_or_default(),
                color: Color::default(),
            });
        }
        return if empty { BorderSide::None } else { opened };
    }
    if name == "color" {
        let color = read_color(e);
        if let Some(borders) = dxf.borders.as_mut()
            && side != BorderSide::None
            && let Some(Some(border)) = diff_side(borders, side)
        {
            border.color = color;
            return side;
        }
        if let Some(font) = dxf.font.as_mut() {
            font.color = Some(color);
        }
        return side;
    }
    if let Some(fill) = dxf.fill.as_mut() {
        match name {
            "patternFill" => {
                if let Some(pattern) = attr(e, "patternType") {
                    fill.pattern = Some(Pattern::parse(&pattern));
                }
            }
            "fgColor" => fill.foreground = Some(read_color(e)),
            "bgColor" => fill.background = Some(read_color(e)),
            _ => {}
        }
    }
    if let Some(font) = dxf.font.as_mut() {
        match name {
            "b" => font.bold = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
            "i" => font.italic = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
            "strike" => font.strike = Some(attr(e, "val").is_none_or(|v| is_true(&v))),
            "u" => {
                font.underline =
                    Some(attr(e, "val").map_or(Underline::Single, |v| Underline::parse(&v)));
            }
            "sz" => {
                font.size = attr(e, "val")
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(points_to_hundredths);
            }
            "name" | "rFont" => font.name = attr(e, "val"),
            _ => {}
        }
    }
    side
}

/// The slot of a differential border for one edge.
fn diff_side(borders: &mut DiffBorders, side: BorderSide) -> Option<&mut Option<Border>> {
    Some(match side {
        BorderSide::Left => &mut borders.left,
        BorderSide::Right => &mut borders.right,
        BorderSide::Top => &mut borders.top,
        BorderSide::Bottom => &mut borders.bottom,
        BorderSide::Diagonal => &mut borders.diagonal,
        BorderSide::None => return None,
    })
}

/// Points to the hundredths of a point a [`Font`] size is kept in.
fn points_to_hundredths(points: f64) -> u32 {
    #[expect(
        clippy::cast_possible_truncation,
        clippy::cast_sign_loss,
        reason = "a font size is a small positive number"
    )]
    let value = (points * 100.0).round() as u32;
    value
}

/// Applies one child element of `<fill>`.
fn apply_fill_child(fill: &mut Fill, name: &str, e: &quick_xml::events::BytesStart<'_>) {
    match name {
        "patternFill" => {
            fill.pattern = attr(e, "patternType").map_or(Pattern::None, |v| Pattern::parse(&v));
        }
        "fgColor" => fill.foreground = read_color(e),
        "bgColor" => fill.background = read_color(e),
        _ => {}
    }
}

/// Applies one child element of `<border>`, returning the side now open.
fn apply_border_child(
    borders: Option<&mut Borders>,
    side: BorderSide,
    name: &str,
    e: &quick_xml::events::BytesStart<'_>,
    empty: bool,
) -> BorderSide {
    let Some(borders) = borders else {
        return side;
    };
    if name == "color" {
        if let Some(target) = side.get_mut(borders) {
            target.color = read_color(e);
        }
        return side;
    }
    let opened = BorderSide::parse(name);
    if opened == BorderSide::None {
        return side;
    }
    if let Some(target) = opened.get_mut(borders) {
        target.style = attr(e, "style").map_or(BorderStyle::None, |v| BorderStyle::parse(&v));
    }
    // A self-closing side has no colour child to wait for.
    if empty { BorderSide::None } else { opened }
}

/// Which table of `styles.xml` the parser is inside.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum Section {
    #[default]
    None,
    Fonts,
    Fills,
    Borders,
    CellXfs,
    /// The differential formats conditional formatting refers to.
    Dxfs,
    /// A section whose elements share names with the ones we want but must be
    /// ignored, such as `cellStyleXfs`.
    Other,
}

/// The border side currently being read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum BorderSide {
    #[default]
    None,
    Left,
    Right,
    Top,
    Bottom,
    Diagonal,
}

impl BorderSide {
    fn parse(name: &str) -> Self {
        match name {
            "left" => Self::Left,
            "right" => Self::Right,
            "top" => Self::Top,
            "bottom" => Self::Bottom,
            "diagonal" => Self::Diagonal,
            _ => Self::None,
        }
    }

    fn get_mut(self, b: &mut Borders) -> Option<&mut Border> {
        Some(match self {
            Self::Left => &mut b.left,
            Self::Right => &mut b.right,
            Self::Top => &mut b.top,
            Self::Bottom => &mut b.bottom,
            Self::Diagonal => &mut b.diagonal,
            Self::None => return None,
        })
    }
}

/// One `cellXfs` entry: indices into the component tables, plus the parts that
/// live on the entry itself.
#[derive(Debug, Default)]
struct Xf {
    number_format: u16,
    font: u32,
    fill: u32,
    border: u32,
    alignment: Alignment,
    protection: Protection,
}

fn read_xf(e: &quick_xml::events::BytesStart<'_>) -> Xf {
    let mut xf = Xf::default();
    for (key, value) in attrs(e) {
        match key.as_str() {
            "numFmtId" => xf.number_format = value.parse().unwrap_or(0),
            "fontId" => xf.font = value.parse().unwrap_or(0),
            "fillId" => xf.fill = value.parse().unwrap_or(0),
            "borderId" => xf.border = value.parse().unwrap_or(0),
            _ => {}
        }
    }
    xf
}

fn read_alignment(e: &quick_xml::events::BytesStart<'_>) -> Alignment {
    let mut a = Alignment::default();
    for (key, value) in attrs(e) {
        match key.as_str() {
            "horizontal" => a.horizontal = HorizontalAlign::parse(&value),
            "vertical" => a.vertical = VerticalAlign::parse(&value),
            "wrapText" => a.wrap_text = is_true(&value),
            "shrinkToFit" => a.shrink_to_fit = is_true(&value),
            "indent" => a.indent = value.parse().unwrap_or(0),
            "textRotation" => a.text_rotation = value.parse().unwrap_or(0),
            "readingOrder" => a.reading_order = value.parse().unwrap_or(0),
            _ => {}
        }
    }
    a
}

fn read_protection(e: &quick_xml::events::BytesStart<'_>) -> Protection {
    let mut p = Protection::default();
    for (key, value) in attrs(e) {
        match key.as_str() {
            "locked" => p.locked = ProtectionState::from_attr(&value),
            "hidden" => p.hidden = ProtectionState::from_attr(&value),
            _ => {}
        }
    }
    p
}

fn diagonal_direction(e: &quick_xml::events::BytesStart<'_>) -> DiagonalDirection {
    let up = attr(e, "diagonalUp").is_some_and(|v| is_true(&v));
    let down = attr(e, "diagonalDown").is_some_and(|v| is_true(&v));
    match (up, down) {
        (true, true) => DiagonalDirection::Both,
        (true, false) => DiagonalDirection::Up,
        (false, true) => DiagonalDirection::Down,
        (false, false) => DiagonalDirection::None,
    }
}

/// Applies one child element of `<font>`.
fn apply_font_child(font: &mut Font, name: &str, e: &quick_xml::events::BytesStart<'_>) {
    // A flag element with no `val` means on: `<b/>` is bold.
    let flag = || attr(e, "val").is_none_or(|v| is_true(&v));
    match name {
        "b" => font.bold = flag(),
        "i" => font.italic = flag(),
        "strike" => font.strike = flag(),
        "u" => font.underline = attr(e, "val").map_or(Underline::Single, |v| Underline::parse(&v)),
        "sz" => {
            if let Some(points) = attr(e, "val").and_then(|v| v.parse::<f64>().ok()) {
                font.set_size_points(points);
            }
        }
        "name" => {
            if let Some(v) = attr(e, "val") {
                font.name = v;
            }
        }
        "color" => font.color = read_color(e),
        "charset" => font.charset = attr(e, "val").and_then(|v| v.parse().ok()),
        "family" => font.family = attr(e, "val").and_then(|v| v.parse().ok()),
        "scheme" => font.scheme = attr(e, "val").and_then(|v| FontScheme::parse(&v)),
        "vertAlign" => {
            font.script = match attr(e, "val").as_deref() {
                Some("superscript") => Script::Superscript,
                Some("subscript") => Script::Subscript,
                _ => Script::Baseline,
            };
        }
        _ => {}
    }
}

/// Reads a colour element, in the order of precedence xlsx uses.
fn read_color(e: &quick_xml::events::BytesStart<'_>) -> Color {
    if let Some(rgb) = attr(e, "rgb")
        && let Some(c) = Color::from_argb_str(&rgb)
    {
        return c;
    }
    if let Some(theme) = attr(e, "theme").and_then(|v| v.parse::<u32>().ok()) {
        let tint = attr(e, "tint")
            .and_then(|v| v.parse::<f64>().ok())
            .map_or(0, |t| {
                #[expect(clippy::cast_possible_truncation, reason = "tint is in -1.0..=1.0")]
                {
                    (t * 1_000_000.0).round() as i32
                }
            });
        return Color::Theme { id: theme, tint };
    }
    if let Some(indexed) = attr(e, "indexed").and_then(|v| v.parse::<u32>().ok()) {
        return Color::Indexed(indexed);
    }
    Color::Auto
}

/// An xlsx boolean, which may be written either way round.
/// Reads one worksheet part.
fn read_sheet<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
    name: &str,
    shared: &[CellValue],
    links: &HashMap<String, Relationship>,
    stream: Option<RowStream<'_>>,
) -> Result<Worksheet> {
    let part = super::zipxml::open_part(zip, path).map_err(Error::Xlsx)?;
    let sheet = Worksheet::new(name)?;
    let mut state = SheetReader::new(sheet, path, shared, links, stream);

    // The part is parsed as it is inflated rather than read into memory
    // first. The one stretch kept as text is the sheet's own `<extLst>`, which
    // travels as written: the recorder holds the bytes from a mark on, the
    // mark follows the parser, and it stays put while that element is read.
    let mut reader = Reader::from_reader(std::io::BufReader::with_capacity(
        1 << 16,
        Recorder::new(part),
    ));
    let mut buf = Vec::new();
    let mut depth = 0u32;
    // Where the `<extLst>` being copied out began, how deep it sits and
    // whether it is a rule's rather than the sheet's.
    let mut extension_start: Option<(u64, u32, bool)> = None;
    // Whether a `<cfRule>` of the sheet is open: its `<extLst>` is kept on
    // the rule. The `x14:cfRule` inside the sheet's own extensions shares
    // the local name, and is never looked at: nothing starts while copying.
    let mut in_rule = false;
    loop {
        buf.clear();
        let before = reader.buffer_position();
        let event = match reader.read_event_into(&mut buf) {
            Ok(e) => e,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
        };
        match event {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let empty = matches!(event, Event::Empty(_));
                if depth == 0 {
                    state.root_namespaces = namespaces_of(e);
                }
                let name = e.local_name();
                if name.as_ref() == "cfRule" && extension_start.is_none() {
                    in_rule = !empty;
                }
                if name.as_ref() == "extLst" && extension_start.is_none() && (depth == 1 || in_rule)
                {
                    if empty {
                        let bytes = reader
                            .get_ref()
                            .get_ref()
                            .since(before, reader.buffer_position());
                        state.keep_extensions(in_rule, &bytes);
                    } else {
                        extension_start = Some((before, depth, in_rule));
                    }
                }
                state.start(e, empty);
                if !empty {
                    depth += 1;
                }
            }
            Event::Text(ref t) => state.text(&t.borrow().into_inner()),
            Event::GeneralRef(ref r) => state.entity(r),
            Event::End(ref e) => {
                depth = depth.saturating_sub(1);
                state.end(e.local_name().as_ref());
                if extension_start.is_none() && e.local_name().as_ref() == "cfRule" {
                    in_rule = false;
                }
                if let Some((start, at, rule)) = extension_start
                    && at == depth
                {
                    extension_start = None;
                    let bytes = reader
                        .get_ref()
                        .get_ref()
                        .since(start, reader.buffer_position());
                    state.keep_extensions(rule, &bytes);
                }
            }
            Event::Eof => break,
            _ => {}
        }
        if extension_start.is_none() {
            let at = reader.buffer_position();
            reader.get_mut().get_mut().forget_before(at);
        }
    }
    Ok(state.sheet)
}

/// A reader that keeps what passed through it from a mark on, so a stretch of
/// a stream a parser has already consumed can still be copied out whole.
struct Recorder<R> {
    inner: R,
    /// Where in the stream `kept` begins.
    base: u64,
    kept: Vec<u8>,
}

impl<R: Read> Recorder<R> {
    const fn new(inner: R) -> Self {
        Self {
            inner,
            base: 0,
            kept: Vec::new(),
        }
    }

    /// Drops what lies before `at`. Only now and then, so dropping stays
    /// cheaper than keeping.
    fn forget_before(&mut self, at: u64) {
        let drop = usize::try_from(at.saturating_sub(self.base)).unwrap_or(usize::MAX);
        if drop >= 1 << 20 && drop <= self.kept.len() {
            self.kept.drain(..drop);
            self.base = at;
        }
    }

    /// The bytes from `start` to `end`, positions in the stream.
    fn since(&self, start: u64, end: u64) -> Vec<u8> {
        let from = usize::try_from(start.saturating_sub(self.base)).unwrap_or(usize::MAX);
        let to = usize::try_from(end.saturating_sub(self.base)).unwrap_or(usize::MAX);
        self.kept
            .get(from..to.min(self.kept.len()))
            .unwrap_or_default()
            .to_vec()
    }
}

impl<R: Read> Read for Recorder<R> {
    fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(out)?;
        self.kept.extend_from_slice(&out[..n]);
        Ok(n)
    }
}

/// Where a streamed sheet's rows go, and what they need to be understood.
#[derive(Clone, Copy)]
struct RowStream<'a> {
    sink: crate::progress::RowSink<'a>,
    styles: &'a crate::style::StyleTable,
    epoch: Epoch,
}

/// The sheet being built and everything the event loop threads through it.
///
/// A sheet part is one flat stream of elements whose meaning depends on what is
/// open around them - a `<color>` inside a colour scale is not a cell colour, a
/// `<formula>` inside a `<cfRule>` is not a cell formula. The flags below are
/// that context; keeping them in one place is what lets the handlers split by
/// the group of elements they read rather than by which variable they touch.
#[expect(
    clippy::struct_excessive_bools,
    reason = "the flags of an XML state machine, not options of a config"
)]
struct SheetReader<'a> {
    sheet: Worksheet,
    sheet_dir: &'a str,
    shared: &'a [CellValue],
    links: &'a HashMap<String, Relationship>,

    /// The cell being read, and what has been gathered for it so far.
    at: Option<CellRef>,
    kind: CellKind,
    style: StyleId,
    value: String,
    formula: String,
    in_value: bool,
    in_formula: bool,
    /// Inside `<rPh>` of an inline string: a reading guide for the text, not
    /// part of it.
    in_phonetic: bool,
    /// The `<r>` runs of an inline string, read as the shared string pool
    /// reads them; empty for a cell written with a bare `<t>`.
    runs: Vec<TextRun>,
    /// Inside `<rPr>` of such a run.
    in_run_font: bool,
    /// `si` of the shared formula this cell takes part in, and the master cell
    /// of each group: xlsx writes the text once and leaves every other cell of
    /// the run to offset it.
    shared_index: Option<String>,
    masters: HashMap<String, (CellRef, String)>,

    /// Which header or footer the text now being read belongs to, and whether
    /// the breaks now being read are between rows or between columns.
    header_part: Option<String>,
    breaks_are_rows: bool,
    /// Whether a `<color>` now being read belongs to a scale, and whether text
    /// now being read is a conditional formatting formula.
    in_scale: bool,
    in_cf_formula: bool,

    /// The data validation being read.
    validation: Option<DataValidation>,
    in_formula1: bool,
    in_formula2: bool,

    /// Offset of the `<filterColumn>` being read. Its criteria arrive as
    /// children, so the column they belong to has to be remembered.
    filter_col: Option<u32>,
    /// Inside an open `<autoFilter>`: a `<sortState>` there is the filter's,
    /// one after it the sheet's.
    in_filter: bool,
    /// The namespaces `<worksheet>` declares, which a carried `<extLst>` may
    /// lean on.
    root_namespaces: Vec<(String, String)>,

    /// How deep inside `<extLst>` the reader is. What an extension holds
    /// travels whole in [`Worksheet::extensions`], and the names inside it
    /// repeat the sheet's own under another namespace: an
    /// `<x14:conditionalFormatting>` read as a `<conditionalFormatting>`
    /// becomes a rule with no range that the writer then puts in the sheet.
    ext_depth: u32,

    /// Where finished rows go, when the caller asked for them one at a time.
    stream: Option<RowStream<'a>>,
    /// The row the cells now being read belong to, as `<row r>` names it.
    row: Option<Row>,
}

impl<'a> SheetReader<'a> {
    fn new(
        sheet: Worksheet,
        path: &'a str,
        shared: &'a [CellValue],
        links: &'a HashMap<String, Relationship>,
        stream: Option<RowStream<'a>>,
    ) -> Self {
        SheetReader {
            sheet,
            sheet_dir: path.rsplit_once('/').map_or("", |(dir, _)| dir),
            shared,
            links,
            at: None,
            kind: CellKind::Number,
            style: StyleId::default(),
            value: String::new(),
            formula: String::new(),
            in_value: false,
            in_formula: false,
            in_phonetic: false,
            runs: Vec::new(),
            in_run_font: false,
            shared_index: None,
            masters: HashMap::new(),
            header_part: None,
            breaks_are_rows: true,
            in_scale: false,
            in_cf_formula: false,
            validation: None,
            in_formula1: false,
            in_formula2: false,
            filter_col: None,
            in_filter: false,
            root_namespaces: Vec::new(),
            ext_depth: 0,
            stream,
            row: None,
        }
    }
}

impl SheetReader<'_> {
    /// Hands a finished row to the stream, if there is one, and forgets its
    /// cells: keeping them is what a streaming read exists to avoid.
    fn flush_row(&mut self) {
        let Some(stream) = self.stream else {
            return;
        };
        // A `<row>` without `r` takes its number from its cells; with neither
        // there is nothing to name it by, and nothing in it to hand over.
        let row = self
            .row
            .take()
            .or_else(|| self.sheet.iter().next().map(|(at, _)| at.row));
        if let Some(row) = row {
            (stream.sink)(&RowBatch {
                row,
                sheet: &self.sheet,
                styles: stream.styles,
                epoch: stream.epoch,
            });
        }
        self.sheet.clear_cells();
    }

    /// Dispatches an opening element to whichever group of the sheet vocabulary
    /// claims it. Each handler answers whether the name was its own.
    fn start(&mut self, e: &quick_xml::events::BytesStart<'_>, empty: bool) {
        // The order is not arbitrary: cells outnumber everything else by an order
        // of magnitude, so `start_cell` goes first.
        let name = e.local_name();
        let name = name.as_ref();
        if name == "extLst" && !empty {
            self.ext_depth += 1;
        }
        if self.ext_depth > 0 {
            return;
        }
        let _ = self.start_cell(name, e, empty)
            || self.start_furniture(name, e)
            || self.start_page(name, e)
            || self.start_conditional(name, e)
            || self.start_validation(name, e, empty)
            || self.start_filter(name, e, empty);
        // `<row/>` never yields an End event, and a streaming read is still
        // owed it: no cells, but maybe a height or a hidden flag.
        if empty && name == "row" {
            self.flush_row();
        }
    }

    /// `<c>` and the elements that live inside one.
    fn start_cell(
        &mut self,
        name: &str,
        e: &quick_xml::events::BytesStart<'_>,
        empty: bool,
    ) -> bool {
        match name {
            "c" => {
                self.at = None;
                self.kind = CellKind::Number;
                self.style = StyleId::default();
                self.value.clear();
                self.formula.clear();
                self.shared_index = None;
                self.runs.clear();
                for attr in e.attributes().flatten() {
                    let v = attr
                        .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                        .unwrap_or_default();
                    // r = address, t = type, s = style.
                    match attr.key.local_name().as_ref() {
                        "r" => self.at = CellRef::parse(&v).ok(),
                        "t" => self.kind = CellKind::from_attr(&v),
                        "s" => {
                            self.style = v.parse().map(StyleId::from_index).unwrap_or_default();
                        }
                        _ => {}
                    }
                }
                // A cell with no children is written as <c r=".." s=".."/>,
                // which never yields an End event. Commit it here or a
                // style-only cell is lost.
                if empty
                    && let Some(at) = self.at.take()
                    && self.style != StyleId::default()
                {
                    *self.sheet.entry(at) = Cell {
                        value: CellValue::Empty,
                        style: self.style,
                    };
                }
            }
            // A self-closing element has no End event to lower the flag again.
            // `<f t="shared" si="0"/>` is common, and leaving the flag up made
            // every later run of text in the sheet - down to the formulas of
            // its data validations - be swallowed as formula source.
            "v" => self.in_value = !empty,
            // `t="inlineStr"` keeps its text in `<is><t>`, or in a `<t>` per
            // run of `<is><r>`, each with its own `<rPr>`.
            "rPh" => self.in_phonetic = !empty,
            "r" if self.kind == CellKind::InlineString && !self.in_phonetic => {
                self.runs.push(TextRun::default());
            }
            "rPr" if self.kind == CellKind::InlineString => {
                if let Some(run) = self.runs.last_mut() {
                    run.font = Some(DiffFont::default());
                    self.in_run_font = !empty;
                }
            }
            name if self.in_run_font => {
                apply_run_font(self.runs.last_mut().and_then(|r| r.font.as_mut()), name, e);
            }
            "t" if self.at.is_some()
                && self.kind == CellKind::InlineString
                && !self.in_phonetic =>
            {
                self.in_value = !empty;
            }
            "f" => {
                self.in_formula = !empty;
                self.shared_index = attr(e, "si");
                if attr(e, "t").as_deref() == Some("array")
                    && let Some(r) = attr(e, "ref").and_then(|r| Range::parse(&r).ok())
                {
                    self.sheet.array_formulas.push(r);
                }
            }
            _ => return false,
        }
        true
    }

    /// Rows, columns, views and everything else attached to the sheet itself.
    fn start_furniture(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>) -> bool {
        match name {
            "sheetFormatPr" => {
                self.sheet.default_column_width =
                    attr(e, "defaultColWidth").and_then(|v| v.parse().ok());
                self.sheet.default_row_height =
                    attr(e, "defaultRowHeight").and_then(|v| v.parse().ok());
            }
            "col" => {
                if let Some(run) = read_column_run(e) {
                    self.sheet.columns.push(run);
                }
            }
            "row" => {
                // Kept whether or not the row has properties worth storing:
                // it is what a streaming read hands the caller. Nobody else
                // wants it, so an ordinary read does not parse it.
                if self.stream.is_some() {
                    self.row = attr(e, "r")
                        .and_then(|v| v.parse().ok())
                        .and_then(|n| Row::from_one_based(n).ok());
                }
                if let Some((row, props)) = read_row_properties(e) {
                    self.sheet.rows.insert(row, props);
                }
            }
            "sheetView" => self.sheet.view = read_sheet_view(e),
            "pane" => self.sheet.view.pane = Some(read_pane(e)),
            "selection" => self.sheet.view.selections.push(read_selection(e)),
            "sheetPr" => {
                self.sheet.properties.fit_to_page = false;
                self.sheet.properties.code_name = attr(e, "codeName");
            }
            "tabColor" => self.sheet.properties.tab_color = Some(read_color(e)),
            "outlinePr" => {
                self.sheet.properties.summary_below =
                    attr(e, "summaryBelow").is_none_or(|v| is_true(&v));
                self.sheet.properties.summary_right =
                    attr(e, "summaryRight").is_none_or(|v| is_true(&v));
            }
            "pageSetUpPr" => {
                self.sheet.properties.fit_to_page =
                    attr(e, "fitToPage").is_some_and(|v| is_true(&v));
            }
            "hyperlink" => {
                if let Some(link) = read_hyperlink(e, self.links) {
                    self.sheet.hyperlinks.push(link);
                }
            }
            "sheetProtection" => read_sheet_protection(&mut self.sheet.protection, e),
            "protectedRange" => self.sheet.protected_ranges.push(read_protected_range(e)),
            "scenarios" => {
                let scenarios = &mut self.sheet.scenarios;
                scenarios.current = attr(e, "current").and_then(|v| v.parse().ok());
                scenarios.show = attr(e, "show").and_then(|v| v.parse().ok());
                scenarios.result_cells =
                    attr(e, "sqref").map(|v| read_sqref(&v)).unwrap_or_default();
            }
            "scenario" => self
                .sheet
                .scenarios
                .list
                .push(crate::model::scenario::Scenario {
                    name: attr(e, "name").unwrap_or_default(),
                    locked: attr(e, "locked").is_some_and(|v| is_true(&v)),
                    hidden: attr(e, "hidden").is_some_and(|v| is_true(&v)),
                    user: attr(e, "user"),
                    comment: attr(e, "comment"),
                    inputs: Vec::new(),
                }),
            "inputCells" => {
                let cell = attr(e, "r").and_then(|v| CellRef::parse(&v).ok());
                if let (Some(cell), Some(scenario)) = (cell, self.sheet.scenarios.list.last_mut()) {
                    scenario.inputs.push(crate::model::scenario::ScenarioInput {
                        cell,
                        value: attr(e, "val").unwrap_or_default(),
                        deleted: attr(e, "deleted").is_some_and(|v| is_true(&v)),
                        undone: attr(e, "undone").is_some_and(|v| is_true(&v)),
                        number_format_id: attr(e, "numFmtId").and_then(|v| v.parse().ok()),
                    });
                }
            }
            "mergeCell" => {
                for attr in e.attributes().flatten() {
                    if attr.key.local_name().as_ref() == "ref"
                        && let Ok(r) = Range::parse(
                            &attr
                                .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                                .unwrap_or_default(),
                        )
                    {
                        self.sheet.merges.push(r);
                    }
                }
            }
            _ => return false,
        }
        true
    }

    /// Margins, printing, headers and page breaks.
    fn start_page(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>) -> bool {
        match name {
            "pageMargins" => {
                // Excel's own defaults, in inches; it converts to centimetres itself.
                let margin = |name: &str, fallback: f64| {
                    attr(e, name)
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(fallback)
                };
                self.sheet.margins = crate::model::PageMargins {
                    left: margin("left", 0.7),
                    right: margin("right", 0.7),
                    top: margin("top", 0.75),
                    bottom: margin("bottom", 0.75),
                    header: margin("header", 0.3),
                    footer: margin("footer", 0.3),
                };
            }
            "pageSetup" => {
                self.sheet.page_setup = read_page_setup(e);
                // The printer settings hang off a relationship of the sheet;
                // the part travels unparsed, the pointer does not.
                self.sheet.page_setup.printer_settings = attr(e, "id")
                    .and_then(|id| self.links.get(&id))
                    .filter(|rel| !rel.external)
                    .map(|rel| resolve(self.sheet_dir, &rel.target));
            }
            "printOptions" => {
                self.sheet.print_options = crate::model::PrintOptions {
                    horizontal_centered: attr(e, "horizontalCentered").is_some_and(|v| is_true(&v)),
                    vertical_centered: attr(e, "verticalCentered").is_some_and(|v| is_true(&v)),
                    headings: attr(e, "headings").is_some_and(|v| is_true(&v)),
                    grid_lines: attr(e, "gridLines").is_some_and(|v| is_true(&v)),
                };
            }
            "headerFooter" => {
                self.sheet.header_footer.different_odd_even =
                    attr(e, "differentOddEven").is_some_and(|v| is_true(&v));
                self.sheet.header_footer.different_first =
                    attr(e, "differentFirst").is_some_and(|v| is_true(&v));
                self.sheet.header_footer.scale_with_doc =
                    attr(e, "scaleWithDoc").is_none_or(|v| is_true(&v));
                self.sheet.header_footer.align_with_margins =
                    attr(e, "alignWithMargins").is_none_or(|v| is_true(&v));
            }
            "oddHeader" | "oddFooter" | "evenHeader" | "evenFooter" | "firstHeader"
            | "firstFooter" => {
                self.header_part = Some(name.to_owned());
            }
            "rowBreaks" => self.breaks_are_rows = true,
            "colBreaks" => self.breaks_are_rows = false,
            "brk" => {
                if let Some(brk) = read_break(e) {
                    if self.breaks_are_rows {
                        self.sheet.row_breaks.push(brk);
                    } else {
                        self.sheet.col_breaks.push(brk);
                    }
                }
            }
            _ => return false,
        }
        true
    }

    /// Conditional formatting: the blocks, their rules and the scales inside.
    fn start_conditional(&mut self, name: &str, e: &quick_xml::events::BytesStart<'_>) -> bool {
        match name {
            "conditionalFormatting" => {
                self.sheet.conditional_formats.push(ConditionalFormat {
                    sqref: attr(e, "sqref").map(|v| read_sqref(&v)).unwrap_or_default(),
                    rules: Vec::new(),
                });
            }
            "cfRule" => {
                if let Some(block) = self.sheet.conditional_formats.last_mut() {
                    block.rules.push(read_cf_rule(e));
                }
            }
            "colorScale" | "dataBar" | "iconSet" => {
                self.in_scale = true;
                if let Some(rule) = last_rule(&mut self.sheet) {
                    rule.scale = Some(new_scale(name, e));
                }
            }
            "cfvo" => {
                if let Some(rule) = last_rule(&mut self.sheet) {
                    push_cfvo(rule, e);
                }
            }
            "color" if self.in_scale => {
                if let Some(rule) = last_rule(&mut self.sheet) {
                    push_scale_color(rule, read_color(e));
                }
            }
            "formula" => {
                self.in_cf_formula = true;
                if let Some(rule) = last_rule(&mut self.sheet) {
                    rule.formulas.push(String::new());
                }
            }
            _ => return false,
        }
        true
    }

    /// Data validations and their two formulas.
    fn start_validation(
        &mut self,
        name: &str,
        e: &quick_xml::events::BytesStart<'_>,
        empty: bool,
    ) -> bool {
        match name {
            "dataValidation" => {
                let dv = read_data_validation(e);
                // <dataValidation .../> with no formula children never yields
                // an End event, so it is committed right here.
                if empty {
                    self.sheet.data_validations.push(dv);
                } else {
                    self.validation = Some(dv);
                }
            }
            "formula1" => self.in_formula1 = !empty,
            "formula2" => self.in_formula2 = !empty,
            _ => return false,
        }
        true
    }

    /// An `<extLst>` copied out whole: the rule's being read, or the sheet's.
    fn keep_extensions(&mut self, rule: bool, bytes: &[u8]) {
        let text = Some(self_contained(
            String::from_utf8_lossy(bytes).trim_start(),
            &self.root_namespaces,
        ));
        if !rule {
            self.sheet.sparklines = text
                .as_deref()
                .map(crate::model::sparkline::read)
                .unwrap_or_default();
            self.sheet.extensions = text;
        } else if let Some(rule) = last_rule(&mut self.sheet) {
            rule.extensions = text;
        }
    }

    /// The autofilter, its columns and each kind of criterion they carry.
    fn start_filter(
        &mut self,
        name: &str,
        e: &quick_xml::events::BytesStart<'_>,
        empty: bool,
    ) -> bool {
        if name == "autoFilter" {
            self.in_filter = !empty;
        }
        let sheet = &mut self.sheet;
        let sort = sort_target(
            &mut sheet.sort_state,
            &mut sheet.auto_filter,
            self.in_filter,
        );
        // A table has an `<autoFilter>` of its own, but it lives in its own
        // part; the one inside a sheet is always the sheet's.
        read_sort_element(sort, name, e)
            || read_filter_element(&mut sheet.auto_filter, &mut self.filter_col, name, e)
    }

    /// A closing element: everything here is a flag the opening element raised,
    /// plus the two elements that commit what they gathered.
    // TODO: too many flags by now; a small stack of open elements would replace them.
    fn end(&mut self, name: &str) {
        if name == "extLst" {
            self.ext_depth = self.ext_depth.saturating_sub(1);
            return;
        }
        if self.ext_depth > 0 {
            return;
        }
        match name {
            "v" | "t" => self.in_value = false,
            "rPh" => self.in_phonetic = false,
            "rPr" => self.in_run_font = false,
            "f" => self.in_formula = false,
            "colorScale" | "dataBar" | "iconSet" => self.in_scale = false,
            "formula" => self.in_cf_formula = false,
            "oddHeader" | "oddFooter" | "evenHeader" | "evenFooter" | "firstHeader"
            | "firstFooter" => self.header_part = None,
            "formula1" => self.in_formula1 = false,
            "formula2" => self.in_formula2 = false,
            "filterColumn" => self.filter_col = None,
            "autoFilter" => self.in_filter = false,
            "row" => self.flush_row(),
            "dataValidation" => {
                if let Some(dv) = self.validation.take() {
                    self.sheet.data_validations.push(dv);
                }
            }
            "c" => {
                if let Some(at) = self.at.take() {
                    let text = shared_formula(
                        &mut self.masters,
                        self.shared_index.as_deref(),
                        at,
                        &self.formula,
                    );
                    let value = if self.runs.is_empty() {
                        build_value(self.kind, &self.value, &text, self.shared)
                    } else {
                        pooled(&std::mem::take(&mut self.runs))
                    };
                    let cell = Cell {
                        value,
                        style: self.style,
                    };
                    // An empty cell carrying only a style still matters:
                    // dropping it would lose the formatting.
                    if !cell.value.is_empty() || cell.style != StyleId::default() {
                        *self.sheet.entry(at) = cell;
                    }
                }
            }
            _ => {}
        }
    }

    /// Character data, which belongs to whichever element is open.
    fn text(&mut self, text: &str) {
        if let Some(out) = self.text_sink() {
            out.push_str(text);
        }
    }

    /// An entity reference, which lands wherever text would.
    fn entity(&mut self, r: &quick_xml::events::BytesRef<'_>) {
        if let Some(out) = self.text_sink() {
            push_entity(out, r);
        }
    }

    /// Where text now being read belongs, if anywhere.
    fn text_sink(&mut self) -> Option<&mut String> {
        // The order of these checks is their priority. Do not reshuffle.
        if self.in_cf_formula {
            let rule = last_rule(&mut self.sheet)?;
            if rule.formulas.is_empty() {
                rule.formulas.push(String::new());
            }
            return rule.formulas.last_mut();
        }
        if self.header_part.is_some() {
            return header_slot(&mut self.sheet, self.header_part.as_deref());
        }
        if self.in_value {
            return Some(match self.runs.last_mut() {
                Some(run) => &mut run.text,
                None => &mut self.value,
            });
        }
        if self.in_formula {
            return Some(&mut self.formula);
        }
        let dv = self.validation.as_mut()?;
        if self.in_formula1 {
            Some(&mut dv.formula1)
        } else if self.in_formula2 {
            Some(&mut dv.formula2)
        } else {
            None
        }
    }
}

/// Resolves a cell's formula text within its shared-formula group.
///
/// The first cell of a group carries the text and becomes the master; the rest
/// are written as `<f t="shared" si="N"/>` with nothing inside and mean "the
/// master's formula, moved here".
fn shared_formula(
    masters: &mut HashMap<String, (CellRef, String)>,
    si: Option<&str>,
    at: CellRef,
    formula: &str,
) -> String {
    let Some(si) = si else {
        return formula.to_owned();
    };
    if !formula.is_empty() {
        masters.insert(si.to_owned(), (at, formula.to_owned()));
        return formula.to_owned();
    }
    match masters.get(si) {
        Some((master, text)) => shift_references(
            text,
            i64::from(at.col.index()) - i64::from(master.col.index()),
            i64::from(at.row.index()) - i64::from(master.row.index()),
        ),
        // A group whose master is missing: nothing can be reconstructed, and
        // the cached value stays as the cell's value.
        None => String::new(),
    }
}

/// The rule now being filled in: the last one of the last block.
fn last_rule(sheet: &mut Worksheet) -> Option<&mut CfRule> {
    sheet.conditional_formats.last_mut()?.rules.last_mut()
}

/// Reads the attributes of a `<cfRule>` element.
fn read_cf_rule(e: &quick_xml::events::BytesStart<'_>) -> CfRule {
    let flag = |name: &str| attr(e, name).is_some_and(|v| is_true(&v));
    CfRule {
        kind: attr(e, "type")
            .map(|v| CfRuleType::parse(&v))
            .unwrap_or_default(),
        priority: attr(e, "priority")
            .and_then(|v| v.parse().ok())
            .unwrap_or(1),
        dxf: attr(e, "dxfId").and_then(|v| v.parse().ok()),
        stop_if_true: flag("stopIfTrue"),
        operator: attr(e, "operator").map(|v| CfOperator::parse(&v)),
        text: attr(e, "text"),
        time_period: attr(e, "timePeriod"),
        rank: attr(e, "rank").and_then(|v| v.parse().ok()),
        percent: flag("percent"),
        bottom: flag("bottom"),
        // The attribute is written only when it is false, so its absence means
        // the rule looks above the average.
        above_average: attr(e, "aboveAverage").is_none_or(|v| is_true(&v)),
        equal_average: flag("equalAverage"),
        std_dev: attr(e, "stdDev").and_then(|v| v.parse().ok()),
        formulas: Vec::new(),
        scale: None,
        extensions: None,
    }
}

/// Starts the graphical part of a rule.
fn new_scale(tag: &str, e: &quick_xml::events::BytesStart<'_>) -> CfScale {
    let number = |name: &str| attr(e, name).and_then(|v| v.parse().ok());
    match tag {
        "dataBar" => CfScale::DataBar {
            values: Vec::new(),
            color: Color::default(),
            min_length: number("minLength"),
            max_length: number("maxLength"),
            // Both are written only when false.
            show_value: attr(e, "showValue").is_none_or(|v| is_true(&v)),
        },
        "iconSet" => CfScale::IconSet {
            values: Vec::new(),
            set: attr(e, "iconSet"),
            show_value: attr(e, "showValue").is_none_or(|v| is_true(&v)),
            percent: attr(e, "percent").is_none_or(|v| is_true(&v)),
            reverse: attr(e, "reverse").is_some_and(|v| is_true(&v)),
        },
        _ => CfScale::Color {
            values: Vec::new(),
            colors: Vec::new(),
        },
    }
}

/// Appends a `<cfvo>` stop to whichever scale the rule carries.
fn push_cfvo(rule: &mut CfRule, e: &quick_xml::events::BytesStart<'_>) {
    let stop = CfValue {
        kind: attr(e, "type")
            .map(|v| CfValueType::parse(&v))
            .unwrap_or_default(),
        value: attr(e, "val").unwrap_or_default(),
        greater_or_equal: attr(e, "gte").is_none_or(|v| is_true(&v)),
    };
    match &mut rule.scale {
        Some(
            CfScale::Color { values, .. }
            | CfScale::DataBar { values, .. }
            | CfScale::IconSet { values, .. },
        ) => values.push(stop),
        None => {}
    }
}

/// Appends a colour to whichever scale the rule carries.
fn push_scale_color(rule: &mut CfRule, value: Color) {
    match &mut rule.scale {
        Some(CfScale::Color { colors, .. }) => colors.push(value),
        Some(CfScale::DataBar { color, .. }) => *color = value,
        _ => {}
    }
}

/// The header or footer string a run of text belongs to.
fn header_slot<'a>(sheet: &'a mut Worksheet, part: Option<&str>) -> Option<&'a mut String> {
    let hf = &mut sheet.header_footer;
    Some(match part? {
        "oddHeader" => &mut hf.odd_header,
        "oddFooter" => &mut hf.odd_footer,
        "evenHeader" => &mut hf.even_header,
        "evenFooter" => &mut hf.even_footer,
        "firstHeader" => &mut hf.first_header,
        "firstFooter" => &mut hf.first_footer,
        _ => return None,
    })
}

/// Reads a `<pageSetup>` element.
fn read_page_setup(e: &quick_xml::events::BytesStart<'_>) -> crate::model::PageSetup {
    let number = |name: &str| attr(e, name).and_then(|v| v.parse().ok());
    let flag = |name: &str| attr(e, name).is_some_and(|v| is_true(&v));
    crate::model::PageSetup {
        paper_size: number("paperSize"),
        orientation: attr(e, "orientation")
            .map(|v| Orientation::parse(&v))
            .unwrap_or_default(),
        scale: number("scale"),
        fit_to_width: number("fitToWidth"),
        fit_to_height: number("fitToHeight"),
        first_page_number: number("firstPageNumber"),
        use_first_page_number: flag("useFirstPageNumber"),
        printer_settings: None,
        horizontal_dpi: number("horizontalDpi"),
        vertical_dpi: number("verticalDpi"),
        over_then_down: attr(e, "pageOrder").is_some_and(|v| v == "overThenDown"),
        black_and_white: flag("blackAndWhite"),
        draft: flag("draft"),
    }
}

/// Reads one `<brk>` element.
fn read_break(e: &quick_xml::events::BytesStart<'_>) -> Option<PageBreak> {
    Some(PageBreak {
        at: attr(e, "id")?.parse().ok()?,
        max: attr(e, "max").and_then(|v| v.parse().ok()),
        manual: attr(e, "man").is_some_and(|v| is_true(&v)),
    })
}

/// Reads one `<hyperlink>` element, resolving an external one through the
/// sheet's relationships.
fn read_hyperlink(
    e: &quick_xml::events::BytesStart<'_>,
    links: &HashMap<String, Relationship>,
) -> Option<Hyperlink> {
    let range = Range::parse(&attr(e, "ref")?).ok()?;
    // `location` points inside the workbook; anything else is a relationship
    // holding a URL or a path.
    let target = if let Some(inside) = attr(e, "location") {
        LinkTarget::Inside(inside)
    } else {
        let rid = attrs(e)
            .into_iter()
            .find(|(k, _)| k == "id")
            .map(|(_, v)| v)?;
        LinkTarget::Outside(links.get(&rid)?.target.clone())
    };
    Some(Hyperlink {
        range,
        target,
        display: attr(e, "display"),
        tooltip: attr(e, "tooltip"),
    })
}

/// Reads the workbook's defined names.
fn read_defined_names<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<Vec<DefinedName>> {
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut out: Vec<DefinedName> = Vec::new();
    let mut open = false;
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) if e.local_name().as_ref() == "definedName" => {
                let Some(name) = attr(&e, "name") else {
                    continue;
                };
                out.push(DefinedName {
                    name,
                    sheet: attr(&e, "localSheetId").and_then(|v| v.parse().ok()),
                    formula: String::new(),
                    hidden: attr(&e, "hidden").is_some_and(|v| is_true(&v)),
                });
                open = true;
            }
            Ok(Event::Text(t)) if open => {
                if let Some(last) = out.last_mut() {
                    last.formula.push_str(&t.xml10_content());
                }
            }
            Ok(Event::GeneralRef(r)) if open => {
                if let Some(last) = out.last_mut() {
                    push_entity(&mut last.formula, &r);
                }
            }
            Ok(Event::End(e)) if e.local_name().as_ref() == "definedName" => open = false,
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(out)
}

/// Reads one `xl/externalLinks/externalLinkN.xml`: the values this workbook
/// last saw in the book it links to.
///
/// Excel recalculates a link from this cache while the other file is closed,
/// and so does [`crate::formula`]. A part that links to something other than a
/// workbook - DDE, OLE - has no `<externalBook>` and reads as an empty entry,
/// which keeps `[N]` counting straight.
fn read_external_link<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
    path: &str,
) -> Result<ExternalBook> {
    // The file it points at is named by the part's own relationships, as an
    // external target.
    let path_of_book = read_relationships(zip, &rels_path_for(path))
        .unwrap_or_default()
        .into_values()
        .find(|r| r.external)
        .map(|r| r.target);
    let xml = read_part(zip, path)?;
    let mut reader = Reader::from_str(&xml);
    let mut names: Vec<String> = Vec::new();
    let mut sheets: Vec<ExternalSheet> = Vec::new();
    // The sheet a `<sheetData>` belongs to, and the cell a `<v>` belongs to.
    let mut sheet: Option<usize> = None;
    let mut cell: Option<(CellRef, String)> = None;
    let mut text = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Empty(e) | Event::Start(e)) => match e.local_name().as_ref() {
                "sheetName" => names.extend(attr(&e, "val")),
                "sheetData" => {
                    let id: usize = attr(&e, "sheetId")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0);
                    sheets.push(ExternalSheet {
                        name: names.get(id).cloned().unwrap_or_default(),
                        cells: BTreeMap::new(),
                    });
                    sheet = Some(sheets.len() - 1);
                }
                "cell" => {
                    cell = attr(&e, "r")
                        .and_then(|r| CellRef::parse(&r).ok())
                        .map(|at| (at, attr(&e, "t").unwrap_or_default()));
                    text.clear();
                }
                _ => {}
            },
            Ok(Event::Text(t)) if cell.is_some() => text.push_str(&t.borrow().into_inner()),
            Ok(Event::GeneralRef(r)) if cell.is_some() => push_entity(&mut text, &r),
            Ok(Event::End(e)) if e.local_name().as_ref() == "cell" => {
                if let (Some((at, kind)), Some(index)) = (cell.take(), sheet)
                    && let Some(s) = sheets.get_mut(index)
                {
                    s.cells.insert((at.row, at.col), cached_value(&kind, &text));
                }
            }
            Ok(Event::Eof) => break,
            Err(e) => return Err(Error::Xlsx(format!("{path}: {e}"))),
            _ => {}
        }
    }
    Ok(ExternalBook {
        path: path_of_book,
        sheets,
    })
}

/// One cached value, by the type tag the cache writes on it.
///
/// The tags are the cell ones minus `s`: a linked book's strings are spelled
/// out here, not shared with a pool this package does not have.
fn cached_value(kind: &str, text: &str) -> CellValue {
    match kind {
        "b" => CellValue::Bool(text == "1"),
        "e" => CellValue::Error(CellError::parse(text).unwrap_or(CellError::Value)),
        "str" | "s" => CellValue::text(text),
        _ => text.parse().map_or(CellValue::Empty, CellValue::Number),
    }
}

/// Parses an `sqref` attribute: areas separated by spaces.
fn read_sqref(value: &str) -> Vec<Range> {
    value
        .split_whitespace()
        .filter_map(|r| Range::parse(r).ok())
        .collect()
}

/// Reads the attributes of a `<sheetView>` element.
fn read_sheet_view(e: &quick_xml::events::BytesStart<'_>) -> SheetView {
    let zoom = |name: &str| {
        attr(e, name)
            .and_then(|v| v.parse::<u32>().ok())
            // A zero or unparsable scale is meaningless; substitutes
            // 100, but dropping the attribute leaves Excel's own default.
            .filter(|&z| z > 0)
    };
    let flag = |name: &str, default: bool| attr(e, name).map_or(default, |v| is_true(&v));
    SheetView {
        view: attr(e, "view")
            .map(|v| SheetViewType::parse(&v))
            .unwrap_or_default(),
        tab_selected: flag("tabSelected", false),
        zoom_scale: zoom("zoomScale"),
        zoom_scale_normal: zoom("zoomScaleNormal"),
        zoom_scale_page_layout: zoom("zoomScalePageLayoutView"),
        zoom_scale_sheet_layout: zoom("zoomScaleSheetLayoutView"),
        top_left_cell: attr(e, "topLeftCell").and_then(|v| CellRef::parse(&v).ok()),
        show_grid_lines: flag("showGridLines", true),
        show_row_col_headers: flag("showRowColHeaders", true),
        show_zeros: flag("showZeros", true),
        right_to_left: flag("rightToLeft", false),
        show_formulas: flag("showFormulas", false),
        show_outline_symbols: flag("showOutlineSymbols", true),
        show_ruler: flag("showRuler", true),
        show_white_space: flag("showWhiteSpace", true),
        window_protection: flag("windowProtection", false),
        // `colorId` means something only when the default is switched off.
        grid_color: if flag("defaultGridColor", true) {
            None
        } else {
            attr(e, "colorId").and_then(|v| v.parse().ok())
        },
        workbook_view_id: attr(e, "workbookViewId")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0),
        pane: None,
        selections: Vec::new(),
    }
}

/// Reads a `<pane>` element.
fn read_pane(e: &quick_xml::events::BytesStart<'_>) -> Pane {
    Pane {
        x_split: attr(e, "xSplit").and_then(|v| v.parse().ok()).unwrap_or(0),
        y_split: attr(e, "ySplit").and_then(|v| v.parse().ok()).unwrap_or(0),
        top_left_cell: attr(e, "topLeftCell").and_then(|v| CellRef::parse(&v).ok()),
        active_pane: attr(e, "activePane")
            .map(|v| PanePosition::parse(&v))
            .unwrap_or_default(),
        state: attr(e, "state")
            .map(|v| PaneState::parse(&v))
            .unwrap_or_default(),
    }
}

/// Reads a `<selection>` element.
fn read_selection(e: &quick_xml::events::BytesStart<'_>) -> Selection {
    Selection {
        pane: attr(e, "pane").map(|v| PanePosition::parse(&v)),
        active_cell: attr(e, "activeCell").and_then(|v| CellRef::parse(&v).ok()),
        sqref: attr(e, "sqref").map(|v| read_sqref(&v)).unwrap_or_default(),
    }
}

/// Reads the attributes of a `<dataValidation>` element; its formulas are
/// child elements and are filled in as they are read.
fn read_data_validation(e: &quick_xml::events::BytesStart<'_>) -> DataValidation {
    let flag = |name: &str| attr(e, name).is_some_and(|v| is_true(&v));
    DataValidation {
        sqref: attr(e, "sqref").map(|v| read_sqref(&v)).unwrap_or_default(),
        kind: attr(e, "type")
            .map(|v| ValidationType::parse(&v))
            .unwrap_or_default(),
        operator: attr(e, "operator")
            .map(|v| ValidationOperator::parse(&v))
            .unwrap_or_default(),
        formula1: String::new(),
        formula2: String::new(),
        allow_blank: flag("allowBlank"),
        hide_drop_down: flag("showDropDown"),
        show_input_message: flag("showInputMessage"),
        show_error_message: flag("showErrorMessage"),
        error_title: attr(e, "errorTitle").unwrap_or_default(),
        error: attr(e, "error").unwrap_or_default(),
        prompt_title: attr(e, "promptTitle").unwrap_or_default(),
        prompt: attr(e, "prompt").unwrap_or_default(),
        error_style: attr(e, "errorStyle")
            .map(|v| ValidationErrorStyle::parse(&v))
            .unwrap_or_default(),
    }
}

/// Reads the flags and the password of `<sheetProtection>`.
fn read_sheet_protection(
    protection: &mut crate::model::SheetProtection,
    e: &quick_xml::events::BytesStart<'_>,
) {
    for (name, slot) in protection.slots() {
        *slot = attr(e, name).map(|v| is_true(&v));
    }
    protection.password = PasswordHash::from_attrs(PasswordAttrs::PLAIN, |name| attr(e, name));
}

/// Reads one `<protectedRange>`.
fn read_protected_range(e: &quick_xml::events::BytesStart<'_>) -> ProtectedRange {
    ProtectedRange {
        name: attr(e, "name").unwrap_or_default(),
        sqref: attr(e, "sqref").map(|v| read_sqref(&v)).unwrap_or_default(),
        password: PasswordHash::from_attrs(PasswordAttrs::PLAIN, |name| attr(e, name)),
        security_descriptor: attr(e, "securityDescriptor").unwrap_or_default(),
    }
}

/// One element of an `<autoFilter>`, the sheet's or a table's: `col` is the
/// `<filterColumn>` open at the moment.
fn read_filter_element(
    auto_filter: &mut Option<AutoFilter>,
    col: &mut Option<u32>,
    name: &str,
    e: &quick_xml::events::BytesStart<'_>,
) -> bool {
    match name {
        "autoFilter" => {
            *auto_filter = attr(e, "ref")
                .and_then(|v| Range::parse(&v).ok())
                .map(AutoFilter::new);
        }
        "filterColumn" => {
            *col = attr(e, "colId").and_then(|v| v.parse().ok());
            if let Some((filter, col_id)) = auto_filter.as_mut().zip(*col) {
                let column = filter.column_at(col_id);
                column.hidden_button = attr(e, "hiddenButton").is_some_and(|v| is_true(&v));
            }
        }
        "filters" | "customFilters" | "dynamicFilter" | "top10" | "colorFilter" | "iconFilter" => {
            if let Some(column) = filter_column(auto_filter, *col) {
                column.filter = ColumnFilter::empty(name);
                if let Some(filter) = column.filter.as_mut() {
                    read_filter_attrs(filter, e);
                }
            }
        }
        "filter" => {
            if let Some(ColumnFilter::Values { values, .. }) =
                filter_column(auto_filter, *col).and_then(|c| c.filter.as_mut())
            {
                values.push(attr(e, "val").unwrap_or_default());
            }
        }
        "dateGroupItem" => {
            if let Some(ColumnFilter::Values { date_groups, .. }) =
                filter_column(auto_filter, *col).and_then(|c| c.filter.as_mut())
            {
                date_groups.push(read_date_group(e));
            }
        }
        "customFilter" => {
            if let Some(ColumnFilter::Custom { rules, .. }) =
                filter_column(auto_filter, *col).and_then(|c| c.filter.as_mut())
            {
                rules.push(CustomFilter {
                    operator: FilterOperator::parse(&attr(e, "operator").unwrap_or_default()),
                    value: attr(e, "val").unwrap_or_default(),
                });
            }
        }
        _ => return false,
    }
    true
}

/// The column of the filter the criteria now being read belong to.
fn filter_column(
    auto_filter: &mut Option<AutoFilter>,
    col_id: Option<u32>,
) -> Option<&mut crate::model::FilterColumn> {
    let col_id = col_id?;
    Some(auto_filter.as_mut()?.column_at(col_id))
}

/// Reads the attributes a filter element carries itself, as opposed to the
/// ones its children do.
fn read_filter_attrs(filter: &mut ColumnFilter, e: &quick_xml::events::BytesStart<'_>) {
    let flag = |name: &str| attr(e, name).is_some_and(|v| is_true(&v));
    match filter {
        ColumnFilter::Values { blank, .. } => *blank = flag("blank"),
        ColumnFilter::Custom { and, .. } => *and = flag("and"),
        ColumnFilter::Dynamic {
            kind,
            value,
            max_value,
        } => {
            *kind = attr(e, "type").unwrap_or_default();
            *value = attr(e, "val");
            *max_value = attr(e, "maxVal");
        }
        ColumnFilter::Top10 {
            value,
            percent,
            top,
            filter_value,
        } => {
            *value = attr(e, "val");
            *percent = flag("percent");
            // `top` defaults to on, so only an explicit `0` turns it off.
            *top = attr(e, "top").is_none_or(|v| is_true(&v));
            *filter_value = attr(e, "filterVal");
        }
        ColumnFilter::Color { dxf, cell_color } => {
            *dxf = attr(e, "dxfId").and_then(|v| v.parse().ok());
            // `cellColor` defaults to on, so only an explicit `0` turns it off.
            *cell_color = attr(e, "cellColor").is_none_or(|v| is_true(&v));
        }
        ColumnFilter::Icon { icon_set, icon_id } => {
            *icon_set = attr(e, "iconSet").unwrap_or_default();
            *icon_id = attr(e, "iconId").and_then(|v| v.parse().ok());
        }
    }
}

/// Where a `<sortState>` being read belongs: the open filter's, or the one
/// beside it on the sheet or table.
fn sort_target<'a>(
    outer: &'a mut Option<SortState>,
    filter: &'a mut Option<AutoFilter>,
    in_filter: bool,
) -> Option<&'a mut Option<SortState>> {
    if in_filter {
        filter.as_mut().map(|f| &mut f.sort_state)
    } else {
        Some(outer)
    }
}

/// A `<sortState>` or one of its `<sortCondition>` keys.
fn read_sort_element(
    state: Option<&mut Option<SortState>>,
    name: &str,
    e: &quick_xml::events::BytesStart<'_>,
) -> bool {
    let flag = |name: &str| attr(e, name).is_some_and(|v| is_true(&v));
    let range = || attr(e, "ref").and_then(|v| Range::parse(&v).ok());
    match name {
        "sortState" => {
            if let Some(state) = state {
                *state = range().map(|range| SortState {
                    range,
                    column_sort: flag("columnSort"),
                    case_sensitive: flag("caseSensitive"),
                    sort_method: attr(e, "sortMethod"),
                    conditions: Vec::new(),
                });
            }
        }
        "sortCondition" => {
            if let (Some(Some(state)), Some(range)) = (state, range()) {
                state.conditions.push(SortCondition {
                    range,
                    descending: flag("descending"),
                    sort_by: attr(e, "sortBy"),
                    custom_list: attr(e, "customList"),
                    dxf: attr(e, "dxfId").and_then(|v| v.parse().ok()),
                    icon_set: attr(e, "iconSet"),
                    icon_id: attr(e, "iconId").and_then(|v| v.parse().ok()),
                });
            }
        }
        _ => return false,
    }
    true
}

/// Reads one `<dateGroupItem>`.
fn read_date_group(e: &quick_xml::events::BytesStart<'_>) -> DateGroup {
    let mut group = DateGroup::default();
    for (name, slot) in group.slots() {
        *slot = attr(e, name).and_then(|v| v.parse().ok());
    }
    group.grouping = attr(e, "dateTimeGrouping").unwrap_or_default();
    group
}

/// Reads one `<col>` element.
fn read_column_run(e: &quick_xml::events::BytesStart<'_>) -> Option<ColumnRun> {
    let first = Col::from_one_based(attr(e, "min")?.parse().ok()?).ok()?;
    let last = Col::from_one_based(attr(e, "max")?.parse().ok()?).ok()?;
    let mut run = ColumnRun::new(first.min(last), first.max(last));
    run.width = attr(e, "width").and_then(|v| v.parse().ok());
    run.custom_width = attr(e, "customWidth").is_some_and(|v| is_true(&v));
    run.hidden = attr(e, "hidden").is_some_and(|v| is_true(&v));
    run.best_fit = attr(e, "bestFit").is_some_and(|v| is_true(&v));
    run.style = attr(e, "style")
        .and_then(|v| v.parse::<u32>().ok())
        .map(StyleId::from_index);
    run.outline_level = attr(e, "outlineLevel")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    run.collapsed = attr(e, "collapsed").is_some_and(|v| is_true(&v));
    // A run that says nothing is not worth keeping; xlsx writers emit them.
    run.is_meaningful().then_some(run)
}

/// Reads the attributes of a `<row>` element.
fn read_row_properties(e: &quick_xml::events::BytesStart<'_>) -> Option<(Row, RowProperties)> {
    let row = Row::from_one_based(attr(e, "r")?.parse().ok()?).ok()?;
    let props = RowProperties {
        height: attr(e, "ht").and_then(|v| v.parse().ok()),
        custom_height: attr(e, "customHeight").is_some_and(|v| is_true(&v)),
        hidden: attr(e, "hidden").is_some_and(|v| is_true(&v)),
        style: attr(e, "s")
            .filter(|_| attr(e, "customFormat").is_some_and(|v| is_true(&v)))
            .and_then(|v| v.parse::<u32>().ok())
            .map(StyleId::from_index),
        outline_level: attr(e, "outlineLevel")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0),
        collapsed: attr(e, "collapsed").is_some_and(|v| is_true(&v)),
    };
    props.is_meaningful().then_some((row, props))
}

/// The `t` attribute of a `<c>` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CellKind {
    /// No `t`, or `t="n"`: the value is a number.
    Number,
    /// `t="s"`: the value is an index into the shared string pool.
    SharedString,
    /// `t="str"`: a string produced by a formula.
    FormulaString,
    /// `t="inlineStr"`: the text sits inside the cell.
    InlineString,
    /// `t="b"`: `0` or `1`.
    Bool,
    /// `t="e"`: an error literal.
    Error,
    /// `t="d"`: an ISO 8601 date.
    IsoDate,
}

impl CellKind {
    fn from_attr(value: &str) -> Self {
        match value {
            "s" => Self::SharedString,
            "str" => Self::FormulaString,
            "inlineStr" => Self::InlineString,
            "b" => Self::Bool,
            "e" => Self::Error,
            "d" => Self::IsoDate,
            _ => Self::Number,
        }
    }
}

/// Turns the raw text of a cell into a value.
fn build_value(kind: CellKind, raw: &str, formula: &str, shared: &[CellValue]) -> CellValue {
    if !formula.is_empty() {
        // The stored `<v>` is the last result Excel computed. Keeping it is what
        // lets a workbook be read without a formula engine.
        let cached = (!raw.is_empty()).then(|| Box::new(scalar(kind, raw, shared)));
        return CellValue::Formula {
            formula: formula.to_owned(),
            cached,
        };
    }
    scalar(kind, raw, shared)
}

/// A pooled string as a cell value: plain when nothing about it changes part
/// way through, and rich when something does.
fn pooled(runs: &[TextRun]) -> CellValue {
    match runs {
        [] => CellValue::Empty,
        [only] if only.font.is_none() => CellValue::text(only.text.clone()),
        _ => CellValue::RichText(runs.to_vec()),
    }
}

/// Reads a non-formula value.
fn scalar(kind: CellKind, raw: &str, shared: &[CellValue]) -> CellValue {
    if raw.is_empty() {
        return CellValue::Empty;
    }
    match kind {
        CellKind::SharedString => raw
            .parse::<usize>()
            .ok()
            .and_then(|i| shared.get(i))
            .cloned()
            .unwrap_or(CellValue::Empty),
        CellKind::FormulaString | CellKind::InlineString | CellKind::IsoDate => {
            CellValue::text(raw)
        }
        CellKind::Bool => CellValue::Bool(raw != "0"),
        CellKind::Error => {
            CellError::parse(raw).map_or_else(|| CellValue::text(raw), CellValue::Error)
        }
        CellKind::Number => raw
            .parse::<f64>()
            .map_or_else(|_| CellValue::text(raw), CellValue::Number),
    }
}

#[cfg(test)]
mod tests {
    use super::{read_xlsx_from, rels_path_for, resolve};
    use crate::coordinate::CellRef;
    use crate::model::CellValue;
    use std::io::Cursor;

    #[test]
    fn rels_paths_follow_the_part() {
        assert_eq!(
            rels_path_for("xl/workbook.xml"),
            "xl/_rels/workbook.xml.rels"
        );
        assert_eq!(
            rels_path_for("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
        assert_eq!(rels_path_for("workbook.xml"), "_rels/workbook.xml.rels");
    }

    #[test]
    fn targets_resolve_against_their_part() {
        assert_eq!(resolve("xl", "styles.xml"), "xl/styles.xml");
        assert_eq!(
            resolve("xl", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
        assert_eq!(
            resolve("xl", "/xl/styles.xml"),
            "xl/styles.xml",
            "absolute target"
        );
        assert_eq!(resolve("", "xl/workbook.xml"), "xl/workbook.xml");
        assert_eq!(
            resolve("xl/worksheets", "../sharedStrings.xml"),
            "xl/sharedStrings.xml"
        );
        assert_eq!(resolve("xl", "../sharedStrings.xml"), "sharedStrings.xml");
    }

    #[test]
    fn rejects_input_that_is_not_a_package() {
        assert!(read_xlsx_from(Cursor::new(b"not a zip at all".to_vec())).is_err());
        assert!(
            read_xlsx_from(Cursor::new(Vec::new())).is_err(),
            "empty input"
        );
    }

    /// Builds a one-sheet package around a given `<sheetData>` and whatever
    /// follows it, so a sheet part can be handed to the reader verbatim.
    fn package(sheet_body: &str) -> Vec<u8> {
        package_with("", sheet_body, &[])
    }

    /// The same with a say over the `<sheet>` attributes and room for extra
    /// parts, spelled however the caller wants them spelled; a part named
    /// twice keeps the caller's.
    fn package_with(sheet_attrs: &str, sheet_body: &str, extra: &[(&str, &str)]) -> Vec<u8> {
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let parts = [
            (
                "_rels/.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
                ),
            ),
            (
                "xl/workbook.xml".to_owned(),
                format!(
                    r#"<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="S" sheetId="1"{sheet_attrs} r:id="rId1"/></sheets></workbook>"#
                ),
            ),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                format!(
                    r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
                ),
            ),
            (
                "xl/worksheets/sheet1.xml".to_owned(),
                format!(r#"<worksheet xmlns="{main}">{sheet_body}</worksheet>"#),
            ),
        ];
        let extra: Vec<(String, String)> = extra
            .iter()
            .map(|(name, body)| ((*name).to_owned(), (*body).to_owned()))
            .collect();
        let mut parts: Vec<(String, String)> = parts
            .into_iter()
            .filter(|(name, _)| !extra.iter().any(|(over, _)| over == name))
            .collect();
        parts.extend(extra);
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            for (name, body) in parts {
                w.start_file::<_, ()>(name, zip::write::SimpleFileOptions::default())
                    .unwrap();
                std::io::Write::write_all(&mut w, body.as_bytes()).unwrap();
            }
            w.finish().unwrap();
        }
        buf
    }

    /// A package whose parts are spelled with capitals Excel does not use, a
    /// shared string holding a CRLF, and a sheet hidden from the tab bar.
    #[test]
    fn odd_spelling_crlf_and_a_hidden_sheet_all_survive() {
        let main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        let strings = format!(
            "<sst xmlns=\"{main}\" count=\"1\" uniqueCount=\"1\"><si><t>one\r\ntwo</t></si></sst>"
        );
        let links = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId9" Type="{rel}/sharedStrings" Target="SharedStrings.xml"/></Relationships>"#
        );
        let book = read_xlsx_from(Cursor::new(package_with(
            r#" state="hidden""#,
            r#"<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>"#,
            &[
                ("xl/SharedStrings.xml", strings.as_str()),
                ("xl/_rels/workbook.xml.rels", links.as_str()),
            ],
        )))
        .expect("package reads");
        let sheet = book.sheet(0).expect("one sheet");
        assert_eq!(sheet.visibility, crate::model::SheetVisibility::Hidden);
        assert_eq!(
            sheet
                .get(CellRef::parse("A1").unwrap())
                .and_then(|c| c.value.plain_text()),
            Some("one\r\ntwo".to_owned()),
            "the CR of a CRLF is the author's, not the XML's"
        );
    }

    #[test]
    fn a_streamed_sheet_hands_over_every_row_and_keeps_none() {
        use crate::progress::{Options, RowBatch};
        use std::cell::RefCell;

        let seen: RefCell<Vec<(u32, usize)>> = RefCell::new(Vec::new());
        let sink = |batch: &RowBatch<'_>| {
            seen.borrow_mut().push((
                batch.row.index() + 1,
                batch.sheet.row_cells(batch.row).count(),
            ));
        };
        let book = super::read_xlsx_from_with(
            Cursor::new(package(concat!(
                r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="C1"><v>3</v></c></row>"#,
                r#"<row r="2" hidden="1"/>"#,
                r#"<row><c r="B3"><v>5</v></c></row>"#,
                "</sheetData>"
            ))),
            u64::MAX,
            &Options::new().streaming(0, &sink),
        )
        .expect("package reads");
        assert_eq!(
            seen.into_inner(),
            vec![(1, 2), (2, 0), (3, 1)],
            "an empty row is still a row, and one without `r` is named by its cells"
        );
        let sheet = book.sheet(0).expect("one sheet");
        assert_eq!(sheet.len(), 0, "nothing of the grid is kept");
        assert_eq!(sheet.dimension(), None);
        assert!(
            sheet.rows.values().any(|props| props.hidden),
            "row properties stay"
        );
    }

    #[test]
    fn an_inline_string_is_read_from_inside_the_cell() {
        // 1C writes every text cell as `t="inlineStr"`; reading only `<v>`
        // left a whole column of such an export empty.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            r#"<sheetData><row r="1">"#,
            r#"<c r="A1" t="inlineStr"><is><t xml:space="preserve">41.01 </t></is></c>"#,
            r#"<c r="B1" t="inlineStr"><is><r><t>Bold</t></r><r><rPr><b/></rPr><t> part</t></r>"#,
            r#"<rPh sb="0" eb="1"><t>ignored</t></rPh></is></c>"#,
            r#"<c r="C1" t="inlineStr"><is><t/></is></c><c r="D1"><v>5</v></c>"#,
            r#"</row></sheetData>"#
        ))))
        .expect("package reads");
        let sheet = book.sheet(0).expect("one sheet");
        let text = |a: &str| {
            sheet
                .get(CellRef::parse(a).unwrap())
                .and_then(|c| c.value.plain_text())
        };
        assert_eq!(text("A1"), Some("41.01 ".to_owned()));
        assert_eq!(
            text("B1"),
            Some("Bold part".to_owned()),
            "runs join, phonetics do not"
        );
        let Some(CellValue::RichText(runs)) = sheet
            .get(CellRef::parse("B1").unwrap())
            .map(|c| c.value.clone())
        else {
            panic!("runs with a font read as rich text");
        };
        assert_eq!(runs.len(), 2);
        assert_eq!(runs[0].font, None);
        assert_eq!(runs[1].font.as_ref().and_then(|f| f.bold), Some(true));
        assert_eq!(text("C1"), None);
        assert_eq!(
            sheet
                .get(CellRef::parse("D1").unwrap())
                .map(|c| c.value.clone()),
            Some(CellValue::Number(5.0))
        );
    }

    #[test]
    fn a_self_closing_element_does_not_swallow_the_rest_of_the_sheet() {
        // `<f t="shared" si="0"/>` never yields an End event. Treating it as an
        // open element left the reader convinced everything after it was
        // formula source, which quietly emptied the data validations that
        // follow it in the part.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            r#"<sheetData><row r="1"><c r="A1"><f t="shared" si="0"/><v>7</v></c></row></sheetData>"#,
            r#"<dataValidations count="1"><dataValidation type="list" sqref="B1">"#,
            r#"<formula1>'Lists'!$A$1:$C$1</formula1></dataValidation></dataValidations>"#
        ))))
        .expect("package reads");
        let sheet = book.sheet(0).expect("one sheet");
        assert_eq!(sheet.data_validations.len(), 1);
        assert_eq!(sheet.data_validations[0].formula1, "'Lists'!$A$1:$C$1");
    }

    #[test]
    fn a_rule_inside_an_extension_is_not_read_as_the_sheets_own() {
        // Excel 2010 keeps rules that refer to other sheets in `<extLst>`,
        // under names the sheet's own rules and validations also use. They
        // travel with the extension; read as the sheet's, each became a rule
        // with no range, written back as `sqref=""`.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            r#"<sheetData/><conditionalFormatting sqref="A1"><cfRule type="expression" priority="2">"#,
            r#"<formula>A1>0</formula></cfRule></conditionalFormatting>"#,
            r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" "#,
            r#"xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">"#,
            r#"<x14:conditionalFormattings><x14:conditionalFormatting "#,
            r#"xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#,
            r#"<x14:cfRule type="expression" priority="1"><xm:f>Other!A1</xm:f></x14:cfRule>"#,
            r#"<xm:sqref>B1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings>"#,
            r#"<x14:dataValidations count="1"><x14:dataValidation type="list">"#,
            r#"<x14:formula1><xm:f>Other!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C1</xm:sqref>"#,
            r#"</x14:dataValidation></x14:dataValidations></ext></extLst>"#
        ))))
        .expect("package reads");
        let sheet = book.sheet(0).expect("one sheet");
        assert_eq!(sheet.conditional_formats.len(), 1);
        assert_eq!(sheet.conditional_formats[0].rules[0].formulas, ["A1>0"]);
        assert_eq!(sheet.data_validations, []);
        assert!(
            sheet
                .extensions
                .as_deref()
                .is_some_and(|x| x.contains("Other!A1"))
        );
    }

    #[test]
    fn shared_formulas_are_restored_for_every_cell_of_the_run() {
        // Only the master carries the text; the rest of the run says "the same
        // formula, moved here" and would otherwise be read as bare numbers.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            "<sheetData>",
            r#"<row r="1"><c r="C1"><f t="shared" ref="C1:C3" si="0">A1+$B$1</f><v>1</v></c></row>"#,
            r#"<row r="2"><c r="C2"><f t="shared" si="0"/><v>2</v></c></row>"#,
            r#"<row r="3"><c r="C3"><f t="shared" si="0"></f><v>3</v></c></row>"#,
            "</sheetData>"
        ))))
        .expect("package reads");
        let sheet = book.sheet(0).expect("one sheet");
        let formula = |a: &str| match &sheet.get(CellRef::parse(a).unwrap()).unwrap().value {
            CellValue::Formula { formula, .. } => formula.clone(),
            other => panic!("{a} is {other:?}, not a formula"),
        };
        assert_eq!(formula("C1"), "A1+$B$1");
        assert_eq!(formula("C2"), "A2+$B$1");
        assert_eq!(formula("C3"), "A3+$B$1");
    }

    #[test]
    fn a_filter_column_is_found_by_its_offset_not_by_its_position() {
        // The file is free to list the columns in any order, and colId counts
        // from the first column of the filter's range rather than from A.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            "<sheetData/>",
            r#"<autoFilter ref="C1:F9">"#,
            r#"<filterColumn colId="3"><filters><filter val="late"/></filters></filterColumn>"#,
            r#"<filterColumn colId="0"><top10 val="5" top="0"/></filterColumn>"#,
            "</autoFilter>"
        ))))
        .expect("package reads");
        let filter = book
            .sheet(0)
            .and_then(|s| s.auto_filter.clone())
            .expect("the sheet has a filter");
        assert_eq!(filter.range.to_string(), "C1:F9");
        assert_eq!(
            filter.columns.iter().map(|c| c.col_id).collect::<Vec<_>>(),
            vec![3, 0],
            "the columns keep the order the file listed them in"
        );
        assert!(matches!(
            filter.columns[1].filter,
            Some(crate::model::ColumnFilter::Top10 { top: false, .. })
        ));
    }

    #[test]
    fn a_protection_flag_the_file_omits_stays_unset() {
        // Three states, not two: `0` is a decision the user made, an absent
        // attribute is Excel's own default, and the two are not the same.
        let book = read_xlsx_from(Cursor::new(package(concat!(
            "<sheetData/>",
            r#"<sheetProtection sheet="1" formatCells="0" password="CC1A"/>"#
        ))))
        .expect("package reads");
        let protection = book
            .sheet(0)
            .map(|s| s.protection.clone())
            .expect("one sheet");
        assert_eq!(protection.sheet, Some(true));
        assert_eq!(protection.format_cells, Some(false));
        assert_eq!(protection.format_rows, None);
        assert_eq!(
            protection.password,
            Some(crate::model::PasswordHash::Legacy("CC1A".into()))
        );
    }

    #[test]
    fn rejects_a_zip_without_a_workbook() {
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            w.start_file::<_, ()>("readme.txt", zip::write::SimpleFileOptions::default())
                .unwrap();
            std::io::Write::write_all(&mut w, b"hello").unwrap();
            w.finish().unwrap();
        }
        let err = read_xlsx_from(Cursor::new(buf)).expect_err("must not be read as a workbook");
        assert!(
            err.to_string().contains("_rels/.rels"),
            "reports the missing part: {err}"
        );
    }
}
