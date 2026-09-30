"use strict";
// ReadyCode Reader for this computer: reads files from disk (streaming large
// Office files, caching documents by content) and hands them to the shared
// core. The MCP server (server.cjs) uses this module.
const fs = require("fs");
const os = require("os");
const path = require("path");
const zlib = require("zlib");
const crypto = require("crypto");
const { StringDecoder } = require("string_decoder");
const core = require("./core.cjs");

const { LIMITS } = core;
const CACHE = path.join(process.env.LOCALAPPDATA || process.env.XDG_CACHE_HOME || path.join(os.homedir(), ".cache"), "readycode-reader");

// ---------- zip on disk ----------
function zipEntries(file) {
  const fd = fs.openSync(file, "r");
  try {
    const size = fs.fstatSync(fd).size;
    const tail = Buffer.alloc(Math.min(size, 65557 + 22));
    fs.readSync(fd, tail, 0, tail.length, size - tail.length);
    let e = -1;
    for (let i = tail.length - 22; i >= 0; i--) if (tail.readUInt32LE(i) === 0x06054b50) { e = i; break; }
    if (e < 0) throw new Error("Not a valid Office file (zip).");
    const count = tail.readUInt16LE(e + 10), cdSize = tail.readUInt32LE(e + 12), cdOff = tail.readUInt32LE(e + 16);
    if (cdOff === 0xffffffff || count === 0xffff) throw new Error("This file uses the ZIP64 format, which is not supported yet.");
    const cd = Buffer.alloc(cdSize);
    fs.readSync(fd, cd, 0, cdSize, cdOff);
    const out = new Map();
    let p = 0;
    for (let k = 0; k < count; k++) {
      const method = cd.readUInt16LE(p + 10), csize = cd.readUInt32LE(p + 20), nlen = cd.readUInt16LE(p + 28), elen = cd.readUInt16LE(p + 30), clen = cd.readUInt16LE(p + 32), local = cd.readUInt32LE(p + 42);
      const name = cd.toString("utf8", p + 46, p + 46 + nlen);
      const lh = Buffer.alloc(30);
      fs.readSync(fd, lh, 0, 30, local);
      out.set(name, { method, csize, start: local + 30 + lh.readUInt16LE(26) + lh.readUInt16LE(28) });
      p += 46 + nlen + elen + clen;
    }
    return out;
  } finally { fs.closeSync(fd); }
}
function diskArchive(file) {
  const entries = zipEntries(file);
  const buffer = (e) => {
    const fd = fs.openSync(file, "r");
    try {
      const raw = Buffer.alloc(e.csize);
      fs.readSync(fd, raw, 0, e.csize, e.start);
      return e.method ? zlib.inflateRawSync(raw) : raw;
    } finally { fs.closeSync(fd); }
  };
  return {
    has: (name) => entries.has(name),
    text: async (name) => buffer(entries.get(name)).toString("utf8"),
    async elements(name, tag, onElement) {
      const e = entries.get(name);
      const raw = fs.createReadStream(file, { start: e.start, end: e.start + e.csize - 1, highWaterMark: 1 << 20 });
      const stream = e.method ? raw.pipe(zlib.createInflateRaw({ chunkSize: 1 << 20 })) : raw;
      const decoder = new StringDecoder("utf8");
      const split = core.elementSplitter(tag, onElement);
      for await (const chunk of stream) split.write(decoder.write(chunk));
      split.write(decoder.end());
      split.end();
    },
  };
}

// ---------- reading ----------
async function readPages(file) {
  const kind = core.kindOf(path.extname(file).toLowerCase());
  if (kind === "pdf") {
    const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
    return core.readPdfPages(getDocument, new Uint8Array(fs.readFileSync(file)));
  }
  if (kind === "docx") return core.readDocxPages(diskArchive(file));
  if (kind === "text") return core.textParts(fs.readFileSync(file, "utf8"));
  throw new Error("Spreadsheets are read with readXlsx.");
}
const readDocx = (file) => core.readDocxPages(diskArchive(file));
const readXlsx = (file) => core.readXlsxTable(diskArchive(file));

// Documents are cached by content; spreadsheets are re-read (far larger).
async function open(file, id, ext, hash) {
  const t0 = Date.now();
  if (core.kindOf(ext) === "xlsx") return core.recordFromTable(id, ext, await readXlsx(file), t0);
  const cached = path.join(CACHE, `${hash}.json`);
  let rec = null;
  try { rec = JSON.parse(fs.readFileSync(cached, "utf8")); } catch (_) { rec = null; }
  if (!rec || rec.version !== 6) {
    rec = core.recordFromPages(id, ext, await readPages(file));
    fs.mkdirSync(CACHE, { recursive: true });
    fs.writeFileSync(cached, JSON.stringify(rec));
  }
  return core.withStore(rec);
}

function createReader({ apiKey = () => process.env.OPENROUTER_API_KEY || "", log = null } = {}) {
  const reader = core.createReaderCore({ apiKey, log });
  async function load(file) {
    if (!file || !fs.existsSync(file)) throw new Error(`File not found: ${file}`);
    const stat = fs.statSync(file);
    if (stat.size > LIMITS.maxFileBytes) throw new Error(`File is larger than ${LIMITS.maxFileBytes / 1024 / 1024} MB.`);
    const ext = path.extname(file).toLowerCase();
    core.kindOf(ext);
    const hash = crypto.createHash("sha256").update(`${file}|${stat.size}|${stat.mtimeMs}`).digest("hex").slice(0, 16);
    const id = `${path.basename(file).replace(/[^\w.-]+/g, "_")}#${hash.slice(0, 6)}`;
    return reader.load(id, () => open(file, id, ext, hash));
  }
  return { load, ask: reader.ask, list: reader.list };
}

module.exports = {
  createReader, readPages, readXlsx, readDocx, zipEntries,
  passagesOf: core.passagesOf, printedPages: core.printedPages, romanValue: core.romanValue,
  makeSearch: core.makeSearch, makeTableSearch: core.makeTableSearch, excelDate: core.excelDate, colName: core.colName, LIMITS,
};
