"use strict";
// Offline tests for the ReadyCode Reader engine (no network).
// Run: node test/engine.test.cjs
const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const zlib = require("zlib");
const e = require("../src/engine.cjs");

// A minimal .docx: a zip holding word/document.xml (deflated), written here so
// the test needs no Word files and no personal documents.
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) { c = (crc ^ buf[n]) & 0xff; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; crc = (crc >>> 8) ^ c; }
  return (crc ^ 0xffffffff) >>> 0;
}
function zip(entries) {
  const locals = [], centrals = [];
  let offset = 0;
  for (const [name, text] of entries) {
    const raw = Buffer.from(text, "utf8"), data = zlib.deflateRawSync(raw), nameBuf = Buffer.from(name);
    const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8); lh.writeUInt32LE(crc32(raw), 14); lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(nameBuf.length, 26);
    const ch = Buffer.alloc(46); ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(8, 10); ch.writeUInt32LE(crc32(raw), 16); ch.writeUInt32LE(data.length, 20); ch.writeUInt32LE(raw.length, 24); ch.writeUInt16LE(nameBuf.length, 28); ch.writeUInt32LE(offset, 42);
    locals.push(lh, nameBuf, data); centrals.push(ch, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  }
  const cd = Buffer.concat(centrals);
  const end = Buffer.alloc(22); end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(entries.length, 8); end.writeUInt16LE(entries.length, 10); end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, cd, end]);
}
const p = (text, style) => `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`;
const docXml = `<?xml version="1.0"?><w:document xmlns:w="w"><w:body>${[
  p("Installation", "Heading1"), p("Unpack the heater and check the parts list against the box."),
  p("Wiring", "Heading2"), p("Only use a 230 V, 10 A grounded circuit &amp; never share it."),
  p("Care", "Heading1"), p("Wipe the glass with a soft cloth; never use abrasive cleaners."),
].join("")}</w:body></w:document>`;

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

test("Word (.docx): sections follow headings and are labelled by their heading path", async () => {
  const file = path.join(os.tmpdir(), `reader-test-${process.pid}.docx`);
  fs.writeFileSync(file, zip([["word/document.xml", docXml]]));
  const sections = await e.readDocx(file);
  assert.deepStrictEqual(sections.map((s) => s.label), ['section "Installation", paragraphs 1-2', 'section "Installation > Wiring", paragraphs 3-4', 'section "Care", paragraphs 5-6']);
  assert.match(sections[1].text, /230 V, 10 A grounded circuit & never share it/);
  fs.unlinkSync(file);
});

test("passages are whole and at most ~1,200 characters, split at line breaks", () => {
  const long = Array.from({ length: 60 }, (_, i) => `Line ${i} of a long page with some words in it.`).join("\n");
  const out = e.passagesOf([{ page: 1, text: long }]);
  assert.ok(out.length >= 2);
  assert.ok(out.every((x) => x.text.length <= e.LIMITS.passageChars + 200));
  assert.strictEqual(out.map((x) => x.text).join("\n"), long, "nothing is lost between passages");
});

test("printed page numbers: a constant offset for main pages, roman numerals for front matter", () => {
  const pages = [];
  for (let n = 1; n <= 30; n++) {
    const printed = n <= 4 ? ["", "", "iii Book Title", "iv Book Title"][n - 1] : `${n - 4} Book Title`;
    pages.push({ page: n, text: `${printed}\nBody text with enough words to count as a page.` });
  }
  const pr = e.printedPages(pages);
  assert.strictEqual(pr[3], "iii");
  assert.strictEqual(pr[4], "iv");
  assert.strictEqual(pr[10], "6");
  assert.strictEqual(pr[1], undefined);
  assert.strictEqual(e.romanValue("xiv"), 14);
  assert.strictEqual(e.romanValue("hello"), null);
});

test("search finds the passage that matches the question words", () => {
  const passages = [{ page: 1, text: "The heater needs a grounded circuit." }, { page: 2, text: "Clean the glass with a soft cloth." }];
  const top = e.makeSearch(passages)("How do I clean the glass?", 1);
  assert.strictEqual(top[0].page, 2);
});

test("Excel (.xlsx): header row, shared strings, dates, repeated headers and duplicate records", async () => {
  const file = path.join(os.tmpdir(), `reader-test-${process.pid}.xlsx`);
  const ss = ["Name", "Phone", "Joined", "Ada Lovelace", "555-0101", "Alan Turing", "555-0199"];
  const c = (ref, v, t, s) => `<c r="${ref}"${t ? ` t="${t}"` : ""}${s ? ` s="${s}"` : ""}><v>${v}</v></c>`;
  const row = (n, cells) => `<row r="${n}">${cells.join("")}</row>`;
  const rows = [row(1, [c("A1", 0, "s"), c("B1", 1, "s"), c("C1", 2, "s")])];
  for (let n = 2; n <= 400; n++) rows.push(row(n, [c(`A${n}`, 5, "s"), c(`B${n}`, 6, "s"), c(`C${n}`, 45000, "", 1)]));
  rows.push(row(401, [c("A401", 0, "s"), c("B401", 1, "s"), c("C401", 2, "s")]));
  rows.push(row(402, [c("A402", 3, "s"), c("B402", 4, "s"), c("C402", 45123, "", 1)]));
  fs.writeFileSync(file, zip([
    ["xl/workbook.xml", `<workbook><sheets><sheet name="People" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>`],
    ["xl/styles.xml", `<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>`],
    ["xl/sharedStrings.xml", `<sst>${ss.map((s) => `<si><t>${s}</t></si>`).join("")}</sst>`],
    ["xl/worksheets/sheet1.xml", `<worksheet><sheetData>${rows.join("")}</sheetData></worksheet>`],
  ]));
  const t = await e.readXlsx(file);
  assert.deepStrictEqual(t.sheets[0].headers, ["Name", "Phone", "Joined"]);
  assert.strictEqual(t.rowStart.length - 1, 400, "the repeated header row is not a record");
  const reader = e.createReader({ apiKey: "" });
  const loaded = await reader.load(file);
  assert.strictEqual(loaded.rows, 400);
  assert.deepStrictEqual(loaded.coverage, { read: "all 400 rows in 1 sheet", repeated_header_rows_skipped: 1 });
  assert.match(loaded.document_tokens_note, /estimate/);
  const r = await reader.ask({ question: "What is Ada Lovelace's phone number?" });
  assert.strictEqual(r.passages[0].row, 402);
  // Only the columns the question needs: the name and the phone, not Joined.
  assert.strictEqual(r.passages[0].text, "Name: Ada Lovelace | Phone: 555-0101");
  assert.deepStrictEqual(r.columns_shown, ["Name", "Phone"]);
  assert.strictEqual(r.passages[0].match, "exact");
  const all = await reader.ask({ question: "Tell me about Ada Lovelace" });
  assert.match(all.passages[0].text, /Name: Ada Lovelace \| Phone: 555-0101 \| Joined: 2023-07-16/);
  const dup = await reader.ask({ question: "Alan Turing phone" });
  assert.strictEqual(dup.passages.length, 1, "399 identical records come back once");
  assert.strictEqual(dup.passages[0].identical_records, 399);
  fs.unlinkSync(file);
});

// A workbook written from plain rows: text as inline strings, numbers as numbers.
function workbook(file, sheets) {
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const cellXml = (ref, v) => (typeof v === "number" ? `<c r="${ref}"><v>${v}</v></c>` : `<c r="${ref}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`);
  const sheetXml = (rows) => `<worksheet><sheetData>${rows.map((cells, i) => `<row r="${i + 1}">${cells.map((v, c) => (v === "" ? "" : cellXml(`${e.colName(c)}${i + 1}`, v))).join("")}</row>`).join("")}</sheetData></worksheet>`;
  fs.writeFileSync(file, zip([
    ["xl/workbook.xml", `<workbook><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`],
    ["xl/_rels/workbook.xml.rels", `<Relationships>${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}</Relationships>`],
    ...sheets.map((s, i) => [`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s.rows)]),
  ]));
}
const STAFF = ["Name", "Company", "Job Title", "Salary"];
const staffFile = () => {
  const file = path.join(os.tmpdir(), `reader-calc-${process.pid}.xlsx`);
  workbook(file, [
    { name: "Staff", rows: [STAFF,
      ["Ada Lovelace", "Roob Inc", "Engineer", 100],
      ["Ada Lovelace", "Roob Inc", "Engineer", 100],
      ["Alan Turing", "Roob Inc", "Analyst", 200],
      ["Grace Hopper", "Mayer Inc", "Engineer", 300],
      ["Nat Becker", "Mayer Inc", "Clerk", 50],
      ["Prof. Nat Becker II", "Smith-Hickle", "Clerk", "n/a"]] },
    { name: "Contractors (2)", rows: [STAFF,
      ["Nat Becker", "Roob Inc", "Clerk", 70],
      ["Linus Torvalds", "roob inc", "Engineer", 80]] },
  ]);
  return file;
};

test("calculations: counts of rows and of people, rankings, distinct values, sums and averages, computed over every row", async () => {
  const file = staffFile();
  const reader = e.createReader({ apiKey: "" });
  const loaded = await reader.load(file);
  assert.strictEqual(loaded.rows, 8);
  assert.strictEqual(loaded.coverage.read, "all 8 rows in 2 sheets");
  const ask = (question) => reader.ask({ question });

  const count = await ask("How many people work at Roob Inc?");
  assert.strictEqual(count.verdict, "calculated");
  assert.strictEqual(count.plan_checked, false, "no key: the plan is computed but marked unchecked");
  assert.strictEqual(count.calculation.result, 5, "case and spacing differences are the same company");
  assert.strictEqual(count.calculation.unique, 4, "Ada Lovelace appears twice");
  assert.deepStrictEqual(count.calculation.matching_rows_by_sheet, { Staff: 3, "Contractors (2)": 2 });
  assert.match(count.answer, /^5 rows where Company is "Roob Inc" in all 2 sheets, holding 4 different Name values\.$/);

  const top = await ask("Which company appears most often?");
  assert.strictEqual(top.calculation.operation, "top");
  assert.deepStrictEqual(top.calculation.result[0], { value: "Roob Inc", rows: 5, unique: 4 });
  assert.match(top.answer, /^Roob Inc \(5 rows, 4 different Name values\) has the most rows by Company/);

  assert.strictEqual((await ask("How many different companies are there?")).calculation.result, 3);
  assert.strictEqual((await ask("How many rows are in sheet Contractors (2)?")).calculation.result, 2);

  const list = await ask("List everyone whose job title is Clerk.");
  assert.strictEqual(list.calculation.result, 2, "two different people");
  assert.deepStrictEqual(list.calculation.entries.map((x) => [x.Name, x.rows]), [["Nat Becker", 2], ["Prof. Nat Becker II", 1]]);

  const total = await ask("What is the total Salary at Roob Inc?");
  assert.strictEqual(total.calculation.result, 550);
  const avg = await ask("What is the average salary?");
  assert.strictEqual(avg.calculation.result, Number((900 / 7).toPrecision(15)));
  assert.strictEqual(avg.calculation.cells_not_numbers, 1, "a text cell is left out and reported");

  // A lookup is not a calculation, and an exact name comes before a longer one.
  const look = await ask("What is the job title of Nat Becker?");
  assert.notStrictEqual(look.verdict, "calculated");
  assert.strictEqual(look.passages[0].match, "exact");
  assert.strictEqual(look.passages[0].text, "Name: Nat Becker | Job Title: Clerk");
  assert.strictEqual(look.passages[0].identical_records, 2, "rows the same in the columns shown come back once");
  assert.ok(look.passages.some((p) => p.match === "partial" && /Prof\. Nat Becker II/.test(p.text)));
  fs.unlinkSync(file);
});

test("calculate: explicit plans, people rather than rows, one sheet, clear errors", async () => {
  const file = staffFile();
  const reader = e.createReader({ apiKey: "" });
  await reader.load(file);
  const byPeople = await reader.calculate({ operation: "top", column: "Job Title", unique_by: "Name", rank_by: "unique", n: 2 });
  assert.deepStrictEqual(byPeople.result, [{ value: "Engineer", rows: 4, unique: 3 }, { value: "Clerk", rows: 3, unique: 2 }]);
  const fewest = await reader.calculate({ operation: "top", column: "Company", order: "fewest", n: 1 });
  assert.strictEqual(fewest.result[0].value, "Smith-Hickle");
  const oneSheet = await reader.calculate({ operation: "count", sheet: "Staff", where: [{ column: "Job Title", equals: "engineer" }] });
  assert.strictEqual(oneSheet.result, 3);
  assert.strictEqual(oneSheet.sheets, "Staff");
  const max = await reader.calculate({ operation: "max", column: "Salary" });
  assert.deepStrictEqual([max.result, max.at], [300, ["Staff row 5"]]);
  const above = await reader.calculate({ operation: "list", where: [{ column: "Salary", above: 90 }, { column: "Company", not_equals: "Mayer Inc" }], columns: ["Name", "Salary"] });
  assert.deepStrictEqual(above.entries.map((x) => x.Name), ["Ada Lovelace", "Ada Lovelace", "Alan Turing"]);
  await assert.rejects(() => reader.calculate({ operation: "sum", column: "Wage" }), /No column named "Wage"\. Columns: Name, Company, Job Title, Salary\./);
  await assert.rejects(() => reader.calculate({ operation: "count", sheet: "Nope" }), /No sheet named "Nope"/);
  fs.unlinkSync(file);
});

test("rows that differ only in columns not shown are counted exactly, however many there are", async () => {
  const file = path.join(os.tmpdir(), `reader-same-${process.pid}.xlsx`);
  const rows = [STAFF];
  for (let i = 0; i < 30; i++) rows.push(["Nat Becker", "Roob Inc", "Clerk", 1000 + i]);
  rows.push(["Prof. Nat Becker II", "Mayer Inc", "Clerk", 5]);
  workbook(file, [{ name: "Staff", rows }]);
  const reader = e.createReader({ apiKey: "" });
  await reader.load(file);
  const r = await reader.ask({ question: "What is the job title of Nat Becker?" });
  assert.strictEqual(r.passages[0].text, "Name: Nat Becker | Job Title: Clerk");
  assert.strictEqual(r.passages[0].identical_records, 30, "30 rows share these shown values, though each has its own salary");
  fs.unlinkSync(file);
});

test("a file still being read reports how far it has got", async () => {
  const core = require("../src/core.cjs");
  const reader = core.createReaderCore({ apiKey: "" });
  let finish;
  const alive = setTimeout(() => {}, 5000); // the wait timer does not keep Node running by itself
  const r = await reader.load("big.xlsx", (prog) => { Object.assign(prog, { phase: "reading rows", sheet: "Data", sheets: 3, sheetsDone: 1, rows: 250000, total: 1000, read: 400 }); return new Promise((resolve) => { finish = resolve; }); }, { waitMs: 20 });
  assert.strictEqual(r.status, "still_reading");
  assert.deepStrictEqual(r.progress, { phase: "reading rows", sheet: "Data", sheets_done: 1, sheets: 3, rows_read: 250000, percent: 40 });
  finish({ id: "big.xlsx", pages: 0, tokens: 0 });
  clearTimeout(alive);
});

test("old .doc files get a clear message", async () => {
  await assert.rejects(() => e.readPages("x.doc"), /save it as \.docx/i);
});

(async () => {
  let passed = 0;
  for (const t of tests) {
    try { await t.fn(); passed++; console.log(`ok   ${t.name}`); } catch (err) { console.log(`FAIL ${t.name}\n     ${String(err && err.stack || err).split("\n").slice(0, 3).join("\n     ")}`); }
  }
  console.log(`${passed}/${tests.length} passed`);
  process.exitCode = passed === tests.length ? 0 : 1;
})();
