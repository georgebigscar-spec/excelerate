const test = require("node:test");
const assert = require("node:assert");
const { Book } = require("excelerate");

test("values and formulas", () => {
  const book = new Book();
  book.set(0, "A1", 2);
  book.set(0, "A2", 3);
  assert.strictEqual(book.get(0, "A1"), 2);
  assert.strictEqual(book.evaluate(0, "A3", "=A1*A2+SUM(A1:A2)"), 11);
  assert.strictEqual(book.evaluate(0, "A3", "=UPPER(\"ой\")"), "ОЙ");
  assert.strictEqual(book.evaluate(0, "A3", "=1/0"), "#DIV/0!");
});

test("a round trip through xlsx", () => {
  const book = new Book();
  book.set(0, "A1", "Гайка");
  book.set(0, "B1", 12.5);
  book.set(0, "C1", true);
  book.addSheet("Второй");

  const reread = Book.read(book.toXlsx());
  assert.deepStrictEqual(reread.sheetNames(), book.sheetNames());
  assert.strictEqual(reread.get(0, "A1"), "Гайка");
  assert.strictEqual(reread.get(0, "B1"), 12.5);
  assert.strictEqual(reread.get(0, "C1"), true);
  assert.strictEqual(reread.get(0, "D1"), null);
});

test("csv", () => {
  const book = new Book();
  book.set(0, "A1", "a,b");
  book.set(0, "B1", 1);
  assert.strictEqual(book.toCsv(0).trim(), '"a,b",1');
});

test("the errors are readable", () => {
  const book = new Book();
  assert.throws(() => book.get(9, "A1"), /no such sheet/);
  assert.throws(() => book.set(0, "нея", 1));
});

test("a round trip with a real file: read -> edit -> write", () => {
  const fs = require("node:fs");
  const os = require("node:os");
  const path = require("node:path");

  // files/gen.xlsx: on sheet "01_Математика" column C holds formulas, D holds
  // the expectations, and the data they read sits in A101 and below.
  const book = Book.read(fs.readFileSync(path.join(__dirname, "files", "gen.xlsx")));
  assert.strictEqual(book.sheetNames()[0], "01_Математика");
  assert.strictEqual(book.get(0, "A101"), 10);
  assert.strictEqual(book.get(0, "D3"), 100);
  assert.strictEqual(book.evaluate(0, "C4", "=SUM(A100:A109)"), 450);

  book.set(0, "A101", 110);
  assert.strictEqual(book.get(0, "A101"), 110);
  assert.strictEqual(book.evaluate(0, "C4", "=SUM(A100:A109)"), 550);

  const out = path.join(os.tmpdir(), `excelerate-${process.pid}.xlsx`);
  fs.writeFileSync(out, book.toXlsx());
  try {
    const reread = Book.read(fs.readFileSync(out));
    assert.deepStrictEqual(reread.sheetNames(), book.sheetNames());
    assert.strictEqual(reread.get(0, "A101"), 110);
    assert.strictEqual(reread.get(0, "D3"), 100);
    assert.strictEqual(reread.evaluate(0, "C4", "=SUM(A100:A109)"), 550);
  } finally {
    fs.unlinkSync(out);
  }
});

test("Book.read works out the format on its own", () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const bytes = fs.readFileSync(path.join(__dirname, "files", "gen.xlsx"));

  // With no file name, the signature alone answers.
  assert.strictEqual(Book.read(bytes).get(0, "A101"), 10);
  // The name changes nothing, even when it lies.
  assert.strictEqual(Book.read(bytes, "gen.csv").get(0, "A101"), 10);

  // CSV announces itself by extension only; without one it is still CSV, by default.
  const csv = Buffer.from("a,b\n1,2\n");
  assert.strictEqual(Book.read(csv, "data.csv").get(0, "B2"), 2);
  assert.strictEqual(Book.read(csv).get(0, "A1"), "a");

  // HTML is recognized by its own tag.
  const html = Buffer.from("<table><tr><td>7</td></tr></table>");
  assert.strictEqual(Book.read(html).get(0, "A1"), 7);

  assert.throws(() => Book.read(Buffer.from("PK\x03\x04не архив")));
});

test("recalculating formulas when a workbook is opened", () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const book = Book.read(fs.readFileSync(path.join(__dirname, "files", "gen.xlsx")));

  // In files/gen.xlsx the formulas carry no cache: column C is the formula, D the expectation.
  assert.strictEqual(book.get(0, "C3"), null);
  assert.strictEqual(book.get(1, "C3"), null);

  // One sheet.
  const onSheet = book.recalculate(0);
  assert.ok(onSheet > 0, `sheet 0 should hold formulas, not ${onSheet}`);
  assert.strictEqual(book.get(0, "C3"), book.get(0, "D3"));
  assert.strictEqual(book.get(1, "C3"), null, "the neighbouring sheet is untouched");

  // The whole workbook: it holds more formulas than any one sheet.
  const whole = book.recalculate();
  assert.ok(whole > onSheet, `${whole} > ${onSheet}`);
  assert.strictEqual(book.get(1, "C3"), book.get(1, "D3"));

  // Recalculation is idempotent: the same formulas, the same values.
  assert.strictEqual(book.recalculate(), whole);
  assert.strictEqual(book.get(0, "C3"), book.get(0, "D3"));

  // And the result survives writing: the cache goes into the file.
  const reread = Book.read(book.toXlsx());
  assert.strictEqual(reread.get(0, "C3"), book.get(0, "D3"));
  assert.strictEqual(reread.get(1, "C3"), book.get(1, "D3"));
});

test("recalculation sees the edit", () => {
  const book = new Book();
  book.set(0, "A1", 2);
  book.set(0, "A2", 3);
  book.set(0, "A3", "=A1+A2");
  assert.strictEqual(book.get(0, "A3"), null);

  assert.strictEqual(book.recalculate(), 1);
  assert.strictEqual(book.get(0, "A3"), 5);

  book.set(0, "A1", 10);
  book.recalculate(0);
  assert.strictEqual(book.get(0, "A3"), 13);
  assert.strictEqual(book.toCsv(0).trim(), "10\r\n3\r\n13");
});

test("an edit recalculates only the cells that depend on it", () => {
  const book = new Book();
  book.set(0, "A1", 2);
  book.set(0, "A2", 3);
  book.set(0, "A3", "=A1+A2");
  book.set(0, "A4", "=A3*10");
  book.set(0, "B1", "=7*6");
  assert.strictEqual(book.recalculate(), 3);

  book.set(0, "A1", 10);
  // A3 reads A1, A4 reads A3; B1 reads none of them.
  assert.strictEqual(book.recalculateFrom(0, "A1"), 2);
  assert.strictEqual(book.get(0, "A3"), 13);
  assert.strictEqual(book.get(0, "A4"), 130);
  assert.strictEqual(book.get(0, "B1"), 42);

  // An edit nobody reads costs nothing.
  book.set(0, "Z99", 1);
  assert.strictEqual(book.recalculateFrom(0, "Z99"), 0);

  assert.throws(() => book.recalculateFrom(9, "A1"), /no such sheet/);
});

test("on a large workbook an edit is cheaper than a full recalculation", () => {
  const fs = require("node:fs");
  const path = require("node:path");
  const book = Book.read(fs.readFileSync(path.join(__dirname, "files", "gen.xlsx")));

  const all = book.recalculate();
  assert.ok(all > 100, `${all} formulas`);
  const touched = book.recalculateFrom(0, "A101");
  assert.ok(touched > 0 && touched < all / 10, `${touched} of ${all}`);
});

test("a batch of edits is one pass", () => {
  const book = new Book();
  for (let row = 1; row <= 200; row++) {
    book.set(0, `A${row}`, row);
    book.set(0, `B${row}`, `=A${row}*2`);
  }
  book.set(0, "D1", "=SUM(B1:B200)");
  assert.strictEqual(book.recalculate(), 201);
  assert.strictEqual(book.get(0, "D1"), 40200);

  // Ten edits, one pass, and it touches only their formulas plus the sum.
  const edited = [];
  for (let row = 1; row <= 10; row++) {
    book.set(0, `A${row}`, 1000 + row);
    edited.push(`A${row}`);
  }
  assert.strictEqual(book.recalculateFromMany(0, edited), 11);
  assert.strictEqual(book.get(0, "B1"), 2002);
  assert.strictEqual(book.get(0, "D1"), 60200);

  // The index survives an edited value and an edited formula.
  book.set(0, "B1", "=A1*3");
  assert.strictEqual(book.recalculateFrom(0, "A1"), 2);
  assert.strictEqual(book.get(0, "B1"), 3003);
  assert.strictEqual(book.get(0, "D1"), 61201);
});

test("ranges are read and written in one call", () => {
  const book = new Book();
  book.setRange(0, "A1", [
    ["Товар", "Цена"],
    ["Болт", 7.5],
    ["Гайка", 12.5],
    ["Итого", "=SUM(B2:B3)"],
  ]);
  book.recalculate();

  assert.strictEqual(book.usedRange(0), "A1:B4");
  assert.deepStrictEqual(book.getRange(0, "A1:B4"), [
    ["Товар", "Цена"],
    ["Болт", 7.5],
    ["Гайка", 12.5],
    ["Итого", 20],
  ]);
  // Beyond what is filled in it answers null, not an error.
  assert.deepStrictEqual(book.getRange(0, "D9:D10"), [[null], [null]]);
  assert.throws(() => book.setRange(0, "A1", "not an array"));
});

test("a formula, formatting and clearing a cell", () => {
  const book = new Book();
  book.set(0, "A1", 0.256);
  book.set(0, "A2", "=A1*2");
  book.recalculate();

  assert.strictEqual(book.getFormula(0, "A2"), "A1*2");
  assert.strictEqual(book.getFormula(0, "A1"), undefined);
  // With no format on the cell it is General.
  assert.strictEqual(book.getFormatted(0, "A1"), "0.256");

  book.clear(0, "A1");
  assert.strictEqual(book.get(0, "A1"), null);
  assert.strictEqual(book.usedRange(0), "A2:A2");
});

test("sheets: name, index, active tab", () => {
  const book = new Book();
  book.addSheet("Данные");

  assert.strictEqual(book.sheetIndex("данные"), 1, "a sheet name is matched case-insensitively");
  assert.strictEqual(book.sheetIndex("no such sheet"), undefined);

  book.renameSheet(0, "Отчёт");
  assert.deepStrictEqual(book.sheetNames(), ["Отчёт", "Данные"]);

  assert.strictEqual(book.activeSheet(), 0);
  book.setActiveSheet(1);
  assert.strictEqual(book.activeSheet(), 1);
  assert.throws(() => book.setActiveSheet(9));
});

test("output in ods, xls and html", () => {
  const book = new Book();
  book.set(0, "A1", "Гайка");
  book.set(0, "B1", 12.5);

  assert.ok(book.toOds().length > 0);
  assert.ok(book.toXls().length > 0);

  const page = book.toHtml(0, true);
  assert.ok(page.includes("Гайка"), "the value lands in the table");
  assert.ok(!page.includes("<html"), "a fragment comes without the wrapper");
  assert.ok(book.toHtml().includes("<html"));
});

test("cell indent and outline levels", () => {
  // files/grouped.xlsx is committed: A2 indented by 3, row 2 at level 2,
  // columns C:D at level 1.
  const fs = require("node:fs");
  const path = require("node:path");
  const file = path.join(__dirname, "files", "grouped.xlsx");

  const book = Book.read(fs.readFileSync(file), "grouped.xlsx");
  assert.strictEqual(book.cellIndent(0, "A1"), 0, "with no style there is no indent");
  assert.strictEqual(book.cellIndent(0, "A2"), 3);

  assert.strictEqual(book.rowLevel(0, 1), 0, "a row outside any group");
  assert.strictEqual(book.rowLevel(0, 2), 2);

  assert.strictEqual(book.columnLevel(0, "A"), 0);
  assert.strictEqual(book.columnLevel(0, "C"), 1, "a column inside the C:D run");
  assert.strictEqual(book.columnLevel(0, "E"), 0);

  assert.throws(() => book.rowLevel(0, 0), "there is no row zero");
  assert.throws(() => book.columnLevel(0, "not letters"));
});

test("addressing by numbers instead of strings", () => {
  const book = new Book();
  book.setRangeAt(0, 1, 1, [
    ["Товар", "Цена"],
    ["Болт", 7.5],
    ["Гайка", 12.5],
    ["Итого", "=SUM(B2:B3)"],
  ]);
  book.recalculate();

  // The numbers are the ones the user sees: row 4, column 2 is B4.
  assert.strictEqual(book.getAt(0, 4, 2), book.get(0, "B4"));
  assert.strictEqual(book.getFormulaAt(0, 4, 2), "SUM(B2:B3)");
  assert.strictEqual(book.getFormattedAt(0, 4, 2), "20");
  assert.strictEqual(book.cellIndentAt(0, 1, 1), 0);

  assert.deepStrictEqual(book.getRangeAt(0, 1, 1, 2, 2), [
    ["Товар", "Цена"],
    ["Болт", 7.5],
  ]);

  book.setAt(0, 2, 2, 20);
  assert.strictEqual(book.recalculateFromAt(0, 2, 2), 1);
  assert.strictEqual(book.getAt(0, 4, 2), 32.5);

  book.clearAt(0, 1, 1);
  assert.strictEqual(book.getAt(0, 1, 1), null);

  assert.throws(() => book.getAt(0, 0, 1), "numbering starts at one, there is no row 0");
  assert.throws(() => book.getRangeAt(0, 1, 1, 0, 2), "an empty range");
});

test("a custom function in a formula", () => {
  const book = new Book();
  book.set(0, "A1", 2);
  book.set(0, "A2", 3);
  book.registerFunction("МОЙИТОГ", (range) => {
    let total = 0;
    for (const row of range) for (const v of row) total += Number(v) || 0;
    return total * 2;
  });
  book.registerFunction("myrate", (x) => x * 1.5 + 0.5);

  assert.strictEqual(book.evaluate(0, "B1", "=МОЙИТОГ(A1:A2)"), 10);
  // The name is matched case-insensitively, as a builtin is.
  assert.strictEqual(book.evaluate(0, "B1", "=MYRATE(10)"), 15.5);
  assert.strictEqual(book.evaluate(0, "B1", "=SUM(MYRATE(1),MYRATE(2))"), 5.5);
  // An array from JS becomes an array of the engine.
  book.registerFunction("МАССИВ", () => [[1, 2], [3, 4]]);
  assert.strictEqual(book.evaluate(0, "B1", "=SUM(МАССИВ())"), 10);
  // A string holding an error code is an error, not text.
  book.registerFunction("СЛОМАН", () => "#N/A");
  assert.strictEqual(book.evaluate(0, "B1", "=СЛОМАН()"), "#N/A");
  // An exception inside the function does not bring the process down.
  book.registerFunction("ПЛОХАЯ", () => { throw new Error("no"); });
  assert.strictEqual(book.evaluate(0, "B1", "=ПЛОХАЯ()"), "#VALUE!");
  // A builtin name does not become the caller's.
  book.registerFunction("SUM", () => -1);
  assert.strictEqual(book.evaluate(0, "B1", "=SUM(A1:A2)"), 5);

  assert.ok(book.registeredFunctions().includes("МОЙИТОГ"));
  assert.strictEqual(book.unregisterFunction("МОЙИТОГ"), true);
  assert.strictEqual(book.evaluate(0, "B1", "=МОЙИТОГ(A1:A2)"), "#NAME?");
});

test("recalculation sees the custom functions", () => {
  const book = new Book();
  book.set(0, "A1", 4);
  book.registerFunction("УДВОЙ", (x) => x * 2);
  book.set(0, "B1", "=УДВОЙ(A1)");
  assert.strictEqual(book.recalculate(), 1);
  assert.strictEqual(book.get(0, "B1"), 8);
});

test("progress while reading, writing and recalculating", () => {
  const book = new Book();
  book.set(0, "A1", 1);
  book.set(0, "B1", "=A1+1");
  book.addSheet("Второй");

  const written = [];
  const bytes = book.toXlsx((p) => {
    assert.strictEqual(p.stage, "writing");
    written.push(p.what);
  });
  assert.ok(written.includes("xl/workbook.xml"), written.join(","));

  const read = [];
  const back = Book.read(bytes, "b.xlsx", undefined, (p) => read.push(p));
  assert.strictEqual(back.sheetNames().length, 2);
  assert.strictEqual(read.length, 2, "one report per sheet");
  assert.strictEqual(read[0].stage, "reading");
  assert.strictEqual(read[0].done, 0);
  assert.strictEqual(read[0].total, 2);
  assert.strictEqual(read[1].fraction, 0.5);

  const steps = [];
  back.recalculate(undefined, (p) => steps.push(p.stage));
  assert.deepStrictEqual(steps, ["recalculating"]);
});

const fs = require("node:fs");
const path = require("node:path");

/** A fixture the repository carries, read as bytes. */
function fixture(name) {
  return fs.readFileSync(path.join(__dirname, "..", "tests", "fixtures", name));
}

test("inserting and removing rows moves the formulas with them", () => {
  const book = new Book();
  book.set(0, "A1", 1);
  book.set(0, "A2", 2);
  book.set(0, "B1", "=SUM(A1:A2)");

  book.insertRows(0, 1, 2);
  assert.strictEqual(book.get(0, "A3"), 1);
  assert.strictEqual(book.getFormula(0, "B3"), "SUM(A3:A4)");

  book.removeRows(0, 1, 2);
  assert.strictEqual(book.getFormula(0, "B1"), "SUM(A1:A2)");

  book.insertColumns(0, 1, 1);
  assert.strictEqual(book.getFormula(0, "C1"), "SUM(B1:B2)");
  book.removeColumns(0, 1, 1);
  assert.strictEqual(book.get(0, "A1"), 1);
});

test("batch edits take places as the sheet was before the call", () => {
  const book = new Book();
  for (let r = 1; r <= 6; r++) book.set(0, `A${r}`, r);
  book.set(0, "B1", "=SUM(A1:A6)");

  book.insertRowsMany(0, [[5, 1], [3, 2]]);
  assert.deepStrictEqual(
    [3, 4, 5, 6, 7, 8, 9].map((r) => book.get(0, `A${r}`)),
    [null, null, 3, 4, null, 5, 6],
  );
  assert.strictEqual(book.getFormula(0, "B1"), "SUM(A1:A9)");

  book.removeRowsMany(0, [[7, 1], [3, 2]]);
  assert.strictEqual(book.getFormula(0, "B1"), "SUM(A1:A6)");
  assert.strictEqual(book.get(0, "A6"), 6);

  book.insertColumnsMany(0, [[1, 1]], "none");
  assert.strictEqual(book.getFormula(0, "C1"), "SUM(B1:B6)");
  book.removeColumnsMany(0, [[1, 1]]);
  assert.strictEqual(book.get(0, "A1"), 1);
  assert.throws(() => book.insertRowsMany(0, [[1]]));
});

test("a removed sheet leaves #REF! behind", () => {
  const book = new Book();
  const second = book.addSheet("Данные");
  book.set(second, "A1", 5);
  book.set(0, "A1", "=Данные!A1");

  book.removeSheet(second);
  assert.strictEqual(book.sheetNames().length, 1);
  assert.match(book.getFormula(0, "A1"), /#REF!/);
});

test("merges", () => {
  const book = new Book();
  book.merge(0, "A1:C1");
  assert.deepStrictEqual(book.mergedRanges(0), ["A1:C1"]);
  assert.deepStrictEqual(Array.from(book.mergedRangesAt(0)), [1, 1, 1, 3]);
  assert.strictEqual(book.unmerge(0, "A1:C1"), true);
  assert.strictEqual(book.unmerge(0, "A1:C1"), false);
  assert.deepStrictEqual(book.mergedRanges(0), []);
});

test("the shape of a sheet without walking it", () => {
  const book = Book.read(fixture("sample.xlsx"), "sample.xlsx");
  assert.strictEqual(book.usedRangeHint(0), book.usedRange(0));
  const height = book.rowHeight(0, 1);
  assert.ok(height == null || height > 0);
  const width = book.columnWidth(0, 1);
  assert.ok(width == null || width > 0);
});

test("tables, comments and links", () => {
  const book = Book.read(fixture("table.xlsx"), "table.xlsx");
  const [table] = book.tables(0);
  assert.ok(table.name.length > 0);
  assert.match(table.range, /^[A-Z]+\d+:[A-Z]+\d+$/);
  assert.ok(Array.isArray(table.columns));
  assert.ok(Array.isArray(book.comments(0)));
  assert.ok(Array.isArray(book.hyperlinks(0)));
});

test("pictures, shapes and charts", () => {
  const pictures = Book.read(fixture("pictures.xlsx"), "pictures.xlsx");
  const images = pictures.images(0);
  assert.ok(images.length > 0);
  assert.strictEqual(images[0].byteLength, pictures.imageData(0, 0).length);
  assert.ok(images[0].anchor.row >= 1);
  assert.throws(() => pictures.imageData(0, 99), /no such image/);

  const shapes = Book.read(fixture("shapes.xlsx"), "shapes.xlsx").shapes(0);
  assert.ok(shapes.length > 0);
  assert.strictEqual(typeof shapes[0].text, "string");
  // A text box from excelize: no fill of its own and a style that adds none,
  // a bold white first run.
  assert.deepStrictEqual(shapes[0].fill, { type: "none" });
  assert.strictEqual(shapes[0].rotation, 0);
  assert.strictEqual(shapes[0].flipH, false);
  assert.strictEqual(shapes[0].font.bold, true);
  assert.strictEqual(shapes[0].font.color, "#FFFFFFFF");

  const charts = Book.read(fixture("chart.xlsx"), "chart.xlsx").charts(0);
  assert.ok(charts.length > 0);
  assert.ok(charts[0].kinds.every((k) => k.endsWith("Chart")));
  assert.ok(charts[0].seriesCount >= 1);
  assert.deepStrictEqual(charts[0].format.fill, { type: "solid", color: "#FFFFFFFF" });
  assert.strictEqual(charts[0].plotFormat, null);
});

test("a shape's and a chart's look survive a write", () => {
  const book = Book.read(fixture("shapes.xlsx"), "shapes.xlsx");
  book.setShapeFormat(0, 0, {
    fill: { type: "solid", color: "accent2" },
    line: { fill: { type: "solid", color: "#80112233" }, width: 2.5 },
    font: { size: 14, italic: true, color: "#FF0000FF" },
    rotation: -90,
    flipH: true,
  });
  book.setShapeFormat(0, 1, {
    fill: { type: "gradient", stops: [{ position: 0, color: "#FF0000" }, { position: 100, color: "bg1" }], angle: 90 },
  });
  const shapes = Book.read(book.toXlsx(), "out.xlsx").shapes(0);
  // accent2 of Office's theme, which excelize's file carries.
  assert.deepStrictEqual(shapes[0].fill, { type: "solid", color: "#FFED7D31" });
  assert.deepStrictEqual(shapes[0].line, { fill: { type: "solid", color: "#80112233" }, width: 2.5 });
  assert.strictEqual(shapes[0].rotation, 270);
  assert.strictEqual(shapes[0].flipH, true);
  assert.strictEqual(shapes[0].font.size, 14);
  assert.strictEqual(shapes[0].font.italic, true);
  assert.strictEqual(shapes[0].font.bold, true, "a field the patch left out stays");
  assert.strictEqual(shapes[0].font.color, "#FF0000FF");
  assert.strictEqual(shapes[0].text, "Итого за квартал");
  assert.deepStrictEqual(shapes[1].fill, {
    type: "gradient",
    stops: [{ position: 0, color: "#FFFF0000" }, { position: 100, color: "#FFFFFFFF" }],
    angle: 90,
    path: null,
  });
  assert.throws(() => book.setShapeFormat(0, 0, { fill: { type: "solid", color: "accent9" } }), /theme colour/);
  assert.throws(() => book.setShapeFormat(0, 99, {}), /no such shape/);

  const charts = Book.read(fixture("chart.xlsx"), "chart.xlsx");
  charts.setChartFormat(0, 0, { format: { fill: { type: "none" } }, plotFormat: { fill: { type: "solid", color: "#EEEEEE" } } });
  const chart = Book.read(charts.toXlsx(), "out.xlsx").charts(0)[0];
  assert.deepStrictEqual(chart.format.fill, { type: "none" });
  assert.deepStrictEqual(chart.plotFormat, { fill: { type: "solid", color: "#FFEEEEEE" }, line: null });
});

test("objects chain: a table from data, a filter, the rows it leaves", () => {
  const { Workbook } = require("excelerate");
  const wb = new Workbook();
  const sales = wb
    .addSheet("Продажи")
    .addTableFromData("Sales", "B2", [
      ["Регион", "Сумма"],
      ["Север", 120],
      ["Юг", 340],
      ["Юг", 95],
      ["Запад", 210],
    ])
    .set(4, "Сумма", 220)
    .addFilter("Регион", { values: ["юг", "Запад"] })
    .addFilter(2, { custom: [{ op: ">", value: 100 }] });

  assert.deepStrictEqual(wb.book.sheetNames(), ["Продажи"], "the first addSheet takes the spare sheet");
  assert.strictEqual(sales.range().address, "B2:C6");
  assert.deepStrictEqual(sales.records({ visible: true }), [
    { Регион: "Юг", Сумма: 340 },
    { Регион: "Запад", Сумма: 220 },
  ]);
  assert.strictEqual(sales.records().length, 4);

  // The criteria and the hidden rows are in the file, not only in memory.
  const back = Workbook.read(wb.toXlsx(), "out.xlsx").sheet("Продажи").table("Sales");
  assert.deepStrictEqual(back.range().visibleValues(), [["Регион", "Сумма"], ["Юг", 340], ["Запад", 220]]);

  sales.clearFilter();
  assert.strictEqual(sales.records({ visible: true }).length, 4);
  sales.addFilter("Сумма", { bottom: 1 });
  assert.deepStrictEqual(sales.records({ visible: true }), [{ Регион: "Юг", Сумма: 95 }]);
  assert.throws(() => sales.addFilter("Нет", { top: 1 }), /no column/);
  assert.throws(() => wb.sheet("Нет"), /no sheet/);

  // A second sheet is added, not taken over.
  assert.strictEqual(wb.addSheet("Итоги").index, 1);
});

test("Sheet and Workbook carry every Book method, chaining what returns nothing", () => {
  const { Workbook, Sheet } = require("excelerate");
  // Every Book method that takes the sheet first, from the declarations.
  const dts = fs.readFileSync(require.resolve("excelerate/excelerate.d.ts"), "utf8");
  const body = dts.slice(dts.indexOf("export class Book"));
  const methods = [...body.slice(0, body.indexOf("\n}\n")).matchAll(/^ {4}(static )?(\w+)\((\w*)/gm)];
  const ownWay = new Set(["renameSheet", "removeSheet", "setActiveSheet", "recalculate", "addTable", "tables", "addSheet", "free", "toXlsx"]);
  for (const [, isStatic, name, first] of methods) {
    if (ownWay.has(name)) continue;
    const owner = isStatic ? Workbook : first === "sheet" ? Sheet.prototype : Workbook.prototype;
    assert.strictEqual(typeof owner[name], "function", `${first === "sheet" ? "Sheet" : "Workbook"} lacks ${name}`);
  }

  const wb = new Workbook();
  const sheet = wb.addSheet("Data")
    .setRange("A1", [["x", "y"], [1, 2], [3, 4]])
    .set("C2", "=A2+B2")
    .setColumnWidth(1, 20)
    .setRowHidden(3, true)
    .setCellStyle("A1", { font: { bold: true } })
    .merge("A5:B5")
    .setComment("A1", "me", "header");
  assert.strictEqual(sheet.recalculate().get("C2"), 3);
  assert.strictEqual(sheet.columnWidth(1), 20);
  assert.strictEqual(sheet.rowHidden(3), true);
  assert.strictEqual(sheet.cellBold("A1"), true);
  assert.deepStrictEqual(sheet.mergedRanges(), ["A5:B5"]);
  assert.strictEqual(sheet.unmerge("A5:B5"), true, "a query returns its answer, not the sheet");
  assert.strictEqual(wb.setDefinedName("Total", "Data!$C$2").definedNames()[0].name, "Total");
  assert.strictEqual(wb.addSheet("Two").activate().workbook.activeSheet(), 1);
  assert.strictEqual(wb.sheet("Two").remove().sheetNames().length, 1);
});

test("column widths, row heights and a hidden sheet survive a round trip", () => {
  const book = new Book();
  book.set(0, "A1", "ширина");
  book.setColumnWidth(0, 1, 32);
  book.setRowHeight(0, 1, 28);
  book.setColumnHidden(0, 3, true);
  book.setRowHidden(0, 2, true);
  const hidden = book.addSheet("Служебный");
  book.setSheetVisibility(hidden, "veryHidden");

  const back = Book.read(book.toXlsx());
  assert.strictEqual(back.columnWidth(0, 1), 32);
  assert.strictEqual(back.rowHeight(0, 1), 28);
  assert.strictEqual(back.rowHidden(0, 2), true);
  assert.strictEqual(back.sheetVisibility(hidden), "veryHidden");
  assert.strictEqual(back.sheetVisibility(0), "visible");

  book.setColumnWidth(0, 1, undefined);
  assert.strictEqual(book.columnWidth(0, 1), undefined);
  assert.throws(() => book.setSheetVisibility(0, "почти"), /veryHidden/);
});

test("painting cells", () => {
  const book = new Book();
  book.set(0, "A1", "Отчёт");
  book.set(0, "B2", 1234.5);
  book.setCellStyle(0, "A1", {
    font: { bold: true, size: 14, color: "#FF1F4E79" },
    fill: { pattern: "solid", foreground: "#FFFFE699" },
    alignment: { horizontal: "center", wrapText: true },
  });
  book.setCellStyle(0, "B2", { numberFormat: "#,##0.00" });
  book.setRangeStyle(0, "A3:C3", { borders: { bottom: { style: "thin", color: "#FF000000" } } });

  const back = Book.read(book.toXlsx());
  const title = back.cellStyle(0, "A1");
  assert.strictEqual(title.font.bold, true);
  assert.strictEqual(title.font.size, 14);
  assert.strictEqual(title.font.color, "#FF1F4E79");
  assert.strictEqual(title.fill.foreground, "#FFFFE699");
  assert.strictEqual(title.alignment.horizontal, "center");
  assert.strictEqual(title.alignment.wrapText, true);
  assert.strictEqual(back.cellStyle(0, "B2").numberFormat, "#,##0.00");
  assert.strictEqual(back.getFormatted(0, "B2"), "1,234.50");
  assert.strictEqual(back.cellStyle(0, "C3").borders.bottom.style, "thin");

  // A patch keeps what it does not mention.
  book.setCellStyle(0, "A1", { font: { italic: true } });
  const both = book.cellStyle(0, "A1");
  assert.strictEqual(both.font.bold, true, "bold survives a later patch");
  assert.strictEqual(both.font.italic, true);
  assert.throws(() => book.setCellStyle(0, "A1", { font: { color: "розовый" } }), /colour/);
});

test("notes, links and tables can be written", () => {
  const book = new Book();
  book.setRange(0, "A1", [["Товар", "Цена"], ["Гайка", 10], ["Болт", 20]]);
  book.setComment(0, "B1", "Иванов", "с НДС");
  book.setHyperlink(0, "A2", "https://example.com", false, "прайс", "открыть");
  book.addTable(0, "Продажи", "A1:B3");

  const back = Book.read(book.toXlsx());
  assert.deepStrictEqual(back.comments(0), [{ address: "B1", author: "Иванов", text: "с НДС" }]);
  const [link] = back.hyperlinks(0);
  assert.strictEqual(link.target, "https://example.com");
  assert.strictEqual(link.external, true);
  assert.strictEqual(link.display, "прайс");
  const [table] = back.tables(0);
  assert.strictEqual(table.displayName, "Продажи");
  assert.deepStrictEqual(table.columns, ["Товар", "Цена"]);
  // A structured reference resolves against it.
  assert.strictEqual(back.evaluate(0, "D1", "=SUM(Продажи[Цена])"), 30);

  assert.throws(() => book.addTable(0, "Продажи", "A1:B3"), /already/);
  assert.strictEqual(book.removeComment(0, "B1"), true);
  assert.strictEqual(book.removeHyperlink(0, "A2"), true);
  assert.strictEqual(book.removeTable(0, "Продажи"), true);
  assert.strictEqual(book.removeTable(0, "Продажи"), false);
});

test("freezing panes and the sheet view", () => {
  const book = new Book();
  book.freezePanes(0, 1, 2);
  book.setZoom(0, 85);
  book.setShowGridLines(0, false, false);

  const back = Book.read(book.toXlsx());
  const view = back.sheetView(0);
  assert.strictEqual(view.frozenRows, 1);
  assert.strictEqual(view.frozenColumns, 2);
  assert.strictEqual(view.zoom, 85);
  assert.strictEqual(view.showGridLines, false);
  assert.strictEqual(view.showRowColHeaders, false);

  book.freezePanes(0, 0, 0);
  assert.strictEqual(book.sheetView(0).frozenRows, 0);
  assert.throws(() => book.setZoom(0, 5), /between 10 and 400/);
});

test("defined names", () => {
  const book = new Book();
  book.set(0, "A1", 6);
  book.set(0, "A2", 7);
  book.setDefinedName("Итог", "Worksheet!$A$1:$A$2");
  assert.strictEqual(book.evaluate(0, "B1", "=SUM(Итог)"), 13);

  const back = Book.read(book.toXlsx());
  const [name] = back.definedNames();
  assert.strictEqual(name.name, "Итог");
  assert.strictEqual(name.sheet, null);
  assert.strictEqual(book.removeDefinedName("Итог"), true);
  assert.strictEqual(book.removeDefinedName("Итог"), false);
  assert.throws(() => book.setDefinedName("Локальный", "A1", 9), /no such sheet/);
});

test("csv with a delimiter of its own", () => {
  const book = new Book();
  book.setRange(0, "A1", [["a", 1], ["b", 2]]);
  assert.strictEqual(book.toCsv(0, { delimiter: ";" }), "a;1\r\nb;2\r\n");
  assert.strictEqual(book.toCsv(0), "a,1\r\nb,2\r\n");

  const read = Book.readCsv(Buffer.from("a;1\n\nb;2\n"), { delimiter: ";", contiguous: true });
  assert.strictEqual(read.get(0, "A2"), "b", "the empty row is skipped");
  assert.throws(() => book.toCsv(0, { delimiter: ";;" }), /one character/);
});

test("protection", () => {
  const book = new Book();
  assert.deepStrictEqual(book.sheetProtection(0), { locked: false, hasPassword: false });
  book.protectSheet(0, "проба");
  book.protectWorkbook("книга", true);

  const back = Book.read(book.toXlsx());
  assert.deepStrictEqual(back.sheetProtection(0), { locked: true, hasPassword: true });
  assert.strictEqual(back.verifySheetPassword(0, "проба"), true);
  assert.strictEqual(back.verifySheetPassword(0, "не проба"), false);
  assert.strictEqual(back.workbookProtection().locked, true);

  book.unprotectSheet(0);
  book.unprotectWorkbook();
  assert.strictEqual(book.sheetProtection(0).locked, false);
  assert.strictEqual(book.workbookProtection().locked, false);
});

test("what a sheet states about its cells", () => {
  const book = Book.read(fixture("sample.xlsx"), "sample.xlsx");
  assert.ok(Array.isArray(book.dataValidations(0)));
  assert.ok(Array.isArray(book.conditionalFormats(0)));
  assert.ok(Array.isArray(book.arrayFormulas(0)));
  assert.ok(Array.isArray(book.externalBooks()));
  // A sheet with no autofilter answers undefined, not an empty object.
  assert.strictEqual(book.autoFilter(0), undefined);

  const pivots = Book.read(fixture("pivot.xlsx"), "pivot.xlsx");
  const found = pivots.sheetNames().flatMap((_, sheet) => pivots.pivotTables(sheet));
  assert.ok(found.length > 0);
  assert.ok(found[0].location.includes(":"));
});

test("copying, moving and reordering", () => {
  const book = new Book();
  book.setRange(0, "A1", [[1], [2], [3]]);
  book.set(0, "B1", "=SUM(A1:A3)");
  book.set(0, "C1", "=A1*2");

  // A copy is rewritten as if it had been written where it lands.
  book.copyRange(0, "B1:B1", "D5");
  assert.strictEqual(book.getFormula(0, "D5"), "SUM(C5:C7)");

  // A move keeps its answers, and drags the references to it along.
  book.moveRange(0, "A1:A3", "F10");
  assert.strictEqual(book.getAt(0, 10, 6), 1);
  assert.strictEqual(book.get(0, "A1"), null);
  assert.strictEqual(book.getFormula(0, "C1"), "F10*2");
  assert.strictEqual(book.getFormula(0, "B1"), "SUM(F10:F12)");

  // Insert Cells: only the column the range spans moves.
  book.set(0, "H1", "beside");
  book.insertCells(0, "F10:F10", "down");
  assert.strictEqual(book.getAt(0, 11, 6), 1);
  assert.strictEqual(book.get(0, "H1"), "beside");
  book.removeCells(0, "F10:F10", "up");
  assert.strictEqual(book.getAt(0, 10, 6), 1);
  assert.throws(() => book.insertCells(0, "A1:A1", "вбок"), /"down", "up"/);

  const second = book.addSheet("Второй");
  book.moveSheet(second, 0);
  assert.deepStrictEqual(book.sheetNames()[0], "Второй");
  assert.strictEqual(book.sheetIndex("Второй"), 0);
  // The formulas name sheets, so moving one changes nothing in them.
  assert.strictEqual(book.getFormula(1, "B1"), "SUM(F10:F12)");
});

test("sorting, filling a series and styles by the block", () => {
  const book = new Book();
  book.setRange(0, "A1", [
    ["Name", "Score"],
    ["ann", 5],
    ["bob", 9],
    ["cy", 5],
  ]);
  book.sortRange(0, "A1:B4", ["-Score", "Name"], { header: true });
  assert.deepStrictEqual(book.getRange(0, "A1:A4").flat(), ["Name", "bob", "ann", "cy"]);
  book.sortRange(0, "A2:B4", [1]);
  assert.deepStrictEqual(book.getRange(0, "A2:A4").flat(), ["ann", "bob", "cy"]);
  assert.throws(() => book.sortRange(0, "A1:B4", ["Nope"], { header: true }), /no header/);

  book.setRange(0, "D1", [[1, "Кв1", "Jan"], [3]]);
  book.fillSeries(0, "D1:F4", "down");
  assert.deepStrictEqual(book.getRange(0, "D4:F4"), [[7, "Кв4", "Apr"]]);

  book.setRangeStyle(0, "A1:B1", { font: { bold: true } });
  const styles = book.getRangeStyles(0, "A1:B2");
  assert.strictEqual(styles.styles.length, 2);
  assert.deepStrictEqual(styles.grid, [[0, 0], [1, 1]]);
  assert.strictEqual(styles.styles[0].font.bold, true);
  // What comes out goes back in: the header's look copied under it.
  book.setRangeStyles(0, "A5", styles);
  assert.strictEqual(book.cellBold(0, "A5"), true);
  assert.strictEqual(book.cellBold(0, "A6"), false);
  book.setRangeStyles(0, "C1", { styles: [{ font: { italic: true } }], grid: [[0, null]] });
  assert.strictEqual(book.cellStyle(0, "C1").font.italic, true);
  assert.strictEqual(book.cellStyle(0, "D1").font.italic, false);
});

test("document properties", () => {
  const book = new Book();
  assert.strictEqual(book.documentProperties().title, null);
  book.setDocumentProperties({
    title: "Отчёт",
    creator: "Анна",
    company: "ООО «Ромашка»",
    created: "2026-09-27T10:00:00Z",
    custom: [
      { name: "Отдел", value: "Продажи" },
      { name: "Страниц", value: 12 },
      { name: "Курс", value: 92.5 },
      { name: "Проверено", value: true },
      { name: "Срок", value: "2026-12-31T00:00:00Z", type: "date" },
    ],
  });
  // A field the patch leaves out stays; null clears one.
  book.setDocumentProperties({ creator: null, subject: "Продажи" });

  for (const format of ["toXlsx", "toXls", "toOds"]) {
    const props = Book.read(book[format]()).documentProperties();
    assert.strictEqual(props.title, "Отчёт", format);
    assert.strictEqual(props.subject, "Продажи", format);
    assert.strictEqual(props.creator, null, format);
    assert.deepStrictEqual(
      props.custom.map((p) => [p.name, p.type, p.value]),
      [
        ["Отдел", "text", "Продажи"],
        ["Страниц", "integer", 12],
        ["Курс", "number", 92.5],
        ["Проверено", "boolean", true],
        ["Срок", "date", "2026-12-31T00:00:00Z"],
      ],
      format,
    );
  }
  // ODS has no place for a company.
  assert.strictEqual(Book.read(book.toXlsx()).documentProperties().company, "ООО «Ромашка»");
  assert.strictEqual(Book.read(book.toOds()).documentProperties().company, null);
  assert.throws(() => book.setDocumentProperties({ title: 5 }), /string or null/);
});

test("formatted text", () => {
  const book = new Book();
  const runs = [
    { text: "Итого: ", font: { bold: true, color: "#FFFF0000" } },
    { text: "x", font: null },
    { text: "2", font: { script: "superscript", size: 8 } },
  ];
  book.setRichText(0, "A1", runs);
  assert.strictEqual(book.get(0, "A1"), "Итого: x2");
  assert.strictEqual(book.getRichText(0, "B1"), null);
  book.set(0, "B1", "plain");
  assert.strictEqual(book.getRichTextAt(0, 1, 2), null);
  // Every format that keeps runs gives them back.
  for (const [format, bytes] of [
    ["xlsx", book.toXlsx()],
    ["ods", book.toOds()],
    ["xls", book.toXls()],
  ]) {
    const back = Book.read(bytes).getRichText(0, "A1");
    assert.deepStrictEqual(
      back.map((r) => r.text),
      ["Итого: ", "x", "2"],
      format,
    );
    assert.strictEqual(back[0].font.bold, true, format);
    assert.strictEqual(back[0].font.color, "#FFFF0000", format);
    assert.strictEqual(back[2].font.script, "superscript", format);
    assert.strictEqual(back[2].font.size, 8, format);
  }
  assert.throws(() => book.setRichText(0, "A1", [{ font: {} }]), /text/);
});

test("remove duplicates and text to columns", () => {
  const book = new Book();
  book.setRange(0, "A1", [["Ann;Oslo"], ["Bob;Rome"], ["ann;OSLO"], ["Cid;Rome;3"]]);
  assert.strictEqual(book.textToColumns(0, "A1:A4", { delimiters: ";" }), 3);
  assert.deepStrictEqual(book.getRange(0, "A4:C4"), [["Cid", "Rome", 3]]);

  const done = book.removeDuplicates(0, "A1:B4", { columns: [1, 2] });
  assert.deepStrictEqual(done, { removed: 1, remaining: 3 });
  assert.deepStrictEqual(book.getRange(0, "A1:A4").flat(), ["Ann", "Bob", "Cid", null]);

  book.setRange(0, "E1", [["2024 Q1  15"]]);
  book.textToColumns(0, "E1", { widths: [4, 7], destination: "F1" });
  assert.deepStrictEqual(book.getRange(0, "F1:H1"), [[2024, "Q1", 15]]);

  assert.throws(() => book.removeDuplicates(0, "A1:B4", { columns: [0] }), /counts from 1/);
  assert.throws(() => book.textToColumns(0, "A1", { delimiters: "" }), /delimiters or widths/);
});

test("goal seek", () => {
  const { Workbook } = require("excelerate");
  const book = new Book();
  book.setRange(0, "A1", [[0.01, 36, 10000, "=PMT(A1,B1,-C1)"]]);
  const found = book.goalSeek(0, "D1", 300, "A1", { tolerance: 1e-7 });
  assert.strictEqual(found.converged, true);
  assert.ok(Math.abs(found.result - 300) < 1e-6, JSON.stringify(found));
  assert.strictEqual(book.get(0, "A1"), found.value);
  assert.ok(Math.abs(book.get(0, "D1") - 300) < 1e-6);
  assert.throws(() => book.goalSeek(0, "A1", 1, "B1"), /must contain a formula/);

  // Through the Sheet object, the sheet comes first implicitly.
  const sheet = new Workbook().sheet(0);
  sheet.set("A1", 1).set("B1", "=A1*A1");
  assert.ok(Math.abs(sheet.goalSeek("B1", 9, "A1").value - 3) < 1e-3);
});

test("data validation rules, checked", () => {
  const book = new Book();
  book.setRange(0, "E1", [["Red"], ["Green"]]);
  book.addDataValidation(0, {
    sqref: "A1:A10",
    type: "whole",
    operator: "between",
    formula1: "1",
    formula2: "10",
    errorTitle: "Out of range",
    error: "From 1 to 10",
  });
  book.addDataValidation(0, { sqref: "B1:B10", type: "list", formula1: "$E$1:$E$2" });
  book.addDataValidation(0, { sqref: "C1:C10", type: "list", formula1: "\"Yes,No\"" });

  assert.strictEqual(book.checkValidation(0, "A3", 7), null);
  const problem = book.checkValidation(0, "A3", 11);
  assert.deepStrictEqual(problem, {
    style: "stop",
    title: "Out of range",
    message: "From 1 to 10",
    sqref: ["A1:A10"],
  });
  assert.notStrictEqual(book.checkValidation(0, "A3", 2.5), null);
  assert.strictEqual(book.checkValidation(0, "B2", "green"), null);
  assert.deepStrictEqual(book.validationList(0, "B5"), ["Red", "Green"]);
  assert.deepStrictEqual(book.validationList(0, "C1"), ["Yes", "No"]);
  assert.deepStrictEqual(book.validationList(0, "D1"), []);

  book.setRange(0, "A1", [[5], [50], [3]]);
  assert.deepStrictEqual(book.invalidCells(0), ["A2"]);

  // The rule survives a round trip as Excel would read it.
  const back = Book.read(book.toXlsx());
  assert.deepStrictEqual(back.invalidCells(0), ["A2"]);
  assert.strictEqual(back.dataValidations(0).length, 3);

  assert.throws(() => book.addDataValidation(0, { sqref: "A1", type: "nope" }), /type is/);
  assert.throws(
    () => book.addDataValidation(0, { sqref: "A1", type: "whole", operator: "bigger" }),
    /not a validation operator/,
  );
});
