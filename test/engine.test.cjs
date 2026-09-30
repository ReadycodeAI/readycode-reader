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
  const r = await reader.ask({ question: "What is Ada Lovelace's phone number?" });
  assert.strictEqual(r.passages[0].row, 402);
  assert.match(r.passages[0].text, /Name: Ada Lovelace \| Phone: 555-0101 \| Joined: 2023-07-16/);
  const dup = await reader.ask({ question: "Alan Turing phone" });
  assert.strictEqual(dup.passages.length, 1, "399 identical records come back once");
  assert.strictEqual(dup.passages[0].identical_records, 399);
  fs.unlinkSync(file);
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
