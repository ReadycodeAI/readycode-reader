// ReadyCode Reader in the browser. The file is read here, on the visitor's
// computer; only short passages go to the decision model, straight from this
// page to OpenRouter with the visitor's own key. Nothing is uploaded to us.
import core from "../../src/core.cjs";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist/build/pdf.mjs";

GlobalWorkerOptions.workerSrc = new URL("./pdf.worker.min.mjs", import.meta.url).href;

// ---------- a zip read from a File, one entry at a time ----------
async function fileArchive(file) {
  const bytes = async (a, b) => new Uint8Array(await file.slice(a, b).arrayBuffer());
  const size = file.size;
  const tailStart = Math.max(0, size - 65557 - 22);
  const tail = await bytes(tailStart, size);
  const dv = new DataView(tail.buffer);
  let e = -1;
  for (let i = tail.length - 22; i >= 0; i--) if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
  if (e < 0) throw new Error("Not a valid Office file (zip).");
  const count = dv.getUint16(e + 10, true), cdSize = dv.getUint32(e + 12, true), cdOff = dv.getUint32(e + 16, true);
  if (cdOff === 0xffffffff || count === 0xffff) throw new Error("This file uses the ZIP64 format, which is not supported yet.");
  const cd = await bytes(cdOff, cdOff + cdSize);
  const cv = new DataView(cd.buffer);
  const utf8 = new TextDecoder();
  const entries = new Map();
  let p = 0;
  for (let k = 0; k < count; k++) {
    const method = cv.getUint16(p + 10, true), csize = cv.getUint32(p + 20, true), nlen = cv.getUint16(p + 28, true), elen = cv.getUint16(p + 30, true), clen = cv.getUint16(p + 32, true), local = cv.getUint32(p + 42, true);
    entries.set(utf8.decode(cd.subarray(p + 46, p + 46 + nlen)), { method, csize, local });
    p += 46 + nlen + elen + clen;
  }
  const stream = async (name, onBytes) => {
    const en = entries.get(name);
    const lh = new DataView((await bytes(en.local, en.local + 30)).buffer);
    const start = en.local + 30 + lh.getUint16(26, true) + lh.getUint16(28, true);
    let raw = file.slice(start, start + en.csize).stream();
    if (onBytes) raw = raw.pipeThrough(new TransformStream({ transform(chunk, out) { onBytes(chunk.byteLength); out.enqueue(chunk); } }));
    return en.method ? raw.pipeThrough(new DecompressionStream("deflate-raw")) : raw;
  };
  return {
    has: (name) => entries.has(name),
    size: (name) => entries.get(name).csize,
    text: async (name) => new Response(await stream(name)).text(),
    async elements(name, tag, onElement, onBytes) {
      const split = core.elementSplitter(tag, onElement);
      const reader = (await stream(name, onBytes)).pipeThrough(new TextDecoderStream()).getReader();
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        split.write(value);
      }
      split.end();
    },
  };
}

async function open(file, id, ext, prog) {
  const t0 = Date.now();
  const kind = core.kindOf(ext);
  if (kind === "xlsx") return core.recordFromTable(id, ext, await core.readXlsxTable(await fileArchive(file), prog), t0, prog);
  let pages;
  if (kind === "pdf") pages = await core.readPdfPages(getDocument, new Uint8Array(await file.arrayBuffer()));
  else if (kind === "docx") pages = await core.readDocxPages(await fileArchive(file));
  else pages = core.textParts(await file.text());
  return core.withStore(core.recordFromPages(id, ext, pages));
}

// ---------- the page ----------
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
const num = (n) => Number(n || 0).toLocaleString("en");
let spent = 0;
// The key lives only in this page's memory: never stored, never sent to us.
const reader = core.createReaderCore({ apiKey: () => $("key").value.trim(), log: (row) => { spent += row.check_cost_usd || 0; } });
let loaded = null;

function status(text, tone = "") { const el = $("status"); el.textContent = text; el.dataset.tone = tone; }

// Every file chosen is read as its own document (two different files can
// share a name and size), and only the latest choice may become the active
// one: a slow earlier read that finishes later is dropped.
let latest = 0;
async function loadFile(file) {
  if (!file) return;
  const ext = (file.name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
  try { core.kindOf(ext); } catch (err) { status(err.message, "bad"); return; }
  if (file.size > core.LIMITS.maxFileBytes) { status("That file is larger than 500 MB.", "bad"); return; }
  const mine = ++latest;
  const id = `${file.name}#${file.size}#${file.lastModified}#${mine}`;
  if (loaded) reader.forget(loaded.id);
  loaded = null;
  shown = [];
  shownFor = null;
  $("ask").disabled = true;
  $("results").innerHTML = "";
  $("about").hidden = true;
  status(`Reading ${file.name} on this computer…`);
  const t0 = performance.now();
  // Large spreadsheets take a while: show how far reading has got.
  const ticker = setInterval(() => {
    const p = reader.progress(id);
    if (!p) return;
    const pct = p.total ? ` ${Math.min(99, Math.floor((100 * p.read) / p.total))}%` : "";
    status(`Reading ${file.name} on this computer…${pct}${p.rows ? `, ${num(p.rows)} rows` : ""}${p.phase === "building the search index" ? ", building the search index" : ""}`);
  }, 500);
  try {
    const sum = await reader.load(id, (prog) => open(file, id, ext, prog), { waitMs: 24 * 3600 * 1000 });
    if (mine !== latest) { reader.forget(id); return; }
    loaded = { id, name: file.name, sum };
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    const size = sum.kind === "spreadsheet"
      ? `${num(sum.rows)} rows in ${sum.sheets.length} sheet${sum.sheets.length === 1 ? "" : "s"}`
      : `${num(sum.pages)} ${ext === ".pdf" ? "page" : "part"}${sum.pages === 1 ? "" : "s"}`;
    status(`Ready: ${file.name}, ${size}, about ${num(sum.document_tokens)} tokens of text. Read in ${secs} s.`, "good");
    showContents(sum);
    $("ask").disabled = false;
  } catch (err) {
    if (mine === latest) status(`Could not read that file: ${err.message}`, "bad");
  } finally {
    clearInterval(ticker);
  }
}

// What the file covers, from its own headings (or its columns), plus a few
// questions to start with, so nobody has to know the file before asking.
const FRONT_MATTER = /^(front cover|back cover|cover|contents|table of contents|foreword|preface|acknowledg|index|references|bibliography|glossary|abbreviations|appendix|notes|about this)/i;
function suggestions(sum) {
  if (sum.kind === "spreadsheet") {
    const ex = sum.example_row;
    if (!ex) return [];
    const cols = Object.keys(ex.values);
    const [first, ...rest] = cols;
    const out = rest.slice(0, 2).map((c) => `What is the ${c} of ${ex.values[first]}?`);
    // Counting suggestions use a short text column (a category, not a number).
    const short = rest.filter((c) => ex.values[c].length <= 30 && !/^[-+\d.,:/ ()%$]+$/.test(ex.values[c])).pop();
    if (short) out.push(`How many rows have ${short} "${ex.values[short]}"?`, `Which ${short} appears most often?`);
    return out;
  }
  const heads = (sum.contents || []).filter((c) => !FRONT_MATTER.test(c.title) && c.title.length >= 8);
  const pickFrom = heads.filter((c) => c.level === 2).length >= 4 ? heads.filter((c) => c.level === 2) : heads;
  const step = Math.max(1, Math.floor(pickFrom.length / 4));
  return pickFrom.filter((_, i) => i % step === 0).slice(0, 4).map((c) => `What does the document say about "${c.title.replace(/[.:]+$/, "")}"?`);
}
function showContents(sum) {
  const box = $("about");
  const qs = suggestions(sum);
  let what = "";
  if (sum.kind === "spreadsheet") {
    what = `<p class="hint">Sheets and columns:</p><ul>${sum.sheets.map((s) => `<li><b>${esc(s.name)}</b> (${num(s.rows)} rows): ${esc(Array.isArray(s.columns) ? s.columns.join(", ") : s.columns)}</li>`).join("")}</ul>`;
  } else if (sum.contents && sum.contents.length) {
    what = `<details><summary>Contents (${sum.contents.length} headings)</summary><ul class="contents">${sum.contents.map((c) => `<li class="l${c.level}">${esc(c.title)}${c.page ? ` <span>p. ${c.page}</span>` : ""}</li>`).join("")}</ul></details>`;
  } else {
    what = `<p class="hint">This file has no headings to list. Ask about anything you expect it to contain.</p>`;
  }
  const warn = [];
  if (Array.isArray(sum.picture_only_pages)) warn.push(`${num(sum.picture_only_pages.length)} page${sum.picture_only_pages.length === 1 ? " has" : "s have"} no text (pictures or scans only) and can't be searched: ${sum.picture_only_pages.slice(0, 12).join(", ")}${sum.picture_only_pages.length > 12 ? "…" : ""}.`);
  if (sum.pictures_not_read) warn.push(`${num(sum.pictures_not_read)} picture${sum.pictures_not_read === 1 ? " is" : "s are"} not read; only the text is searched.`);
  const cov = sum.coverage || {};
  if (cov.not_read) warn.push(`Not read: ${cov.not_read.join(", ")} (${cov.not_read_note}).`);
  if (cov.sheets_without_rows) warn.push(`No rows in: ${cov.sheets_without_rows.join(", ")}.`);
  if (cov.repeated_header_rows_skipped) warn.push(`${num(cov.repeated_header_rows_skipped)} repeated header row${cov.repeated_header_rows_skipped === 1 ? " was" : "s were"} skipped.`);
  for (const s of sum.sheets || []) if (s.header_note) warn.push(`${s.name}: ${s.header_note}`);
  const warnings = warn.length ? `<ul class="warnings">${warn.map((w) => `<li>${esc(w)}</li>`).join("")}</ul>` : "";
  box.innerHTML = `<label>What's in this file</label>${cov.read ? `<p class="hint" style="margin-top:0">Read: ${esc(cov.read)}.</p>` : ""}${warnings}${what}${qs.length ? `<p class="hint">Try one (click to add it):</p><div class="chips">${qs.map((q) => `<button type="button" class="chip">${esc(q)}</button>`).join("")}</div>` : ""}`;
  box.hidden = false;
  box.querySelectorAll(".chip").forEach((b) => b.addEventListener("click", () => {
    const t = $("questions");
    t.value = (t.value.trim() ? `${t.value.trim()}\n` : "") + b.textContent;
  }));
}

function cite(p) {
  if (p.sheet) return `${p.sheet}, row ${num(p.row)}`;
  if (p.where) return p.where;
  return `page ${p.page}${p.printed_page ? ` (printed ${p.printed_page})` : ""}`;
}
const VERDICT = {
  answer_from_passages: ["Answered from the document", "good"],
  low_confidence: ["Partly answered", "warn"],
  not_in_document: ["Not in the document", "muted"],
  needs_calculation: ["Needs an exact calculation", "warn"],
  no_matching_text: ["No matching text found", "muted"],
  calculated: ["Calculated exactly from every row", "good"],
  unchecked: ["Unchecked: the check did not run", "warn"],
};

// Rankings and lists from a calculation: the first 20 lines, the rest behind
// "Show all", a CSV download, and a button to fetch the rest of a long list.
function calcLines(c) {
  if (!c) return [];
  if (c.operation === "top") return c.result.map((g) => [g.value, `${num(g.rows)} rows${g.unique != null ? `, ${num(g.unique)} different ${c.unique_by}` : ""}`]);
  if (c.entries) return c.entries.map((x) => { const { at, rows: n, ...vals } = x; return [Object.values(vals).join(" · "), `${n != null ? `${num(n)} rows, first at ` : ""}${at}`]; });
  return [];
}
function calcTable(c, i) {
  const rows = calcLines(c);
  if (!rows.length) return "";
  const tr = (list) => list.map(([a, b]) => `<tr><td>${esc(a)}</td><td>${esc(b)}</td></tr>`).join("");
  return `<table class="calc">${tr(rows.slice(0, 20))}</table>${rows.length > 20 ? `<details><summary>Show all ${num(rows.length)}</summary><table class="calc">${tr(rows.slice(20))}</table></details>` : ""}
    <div class="tools"><button type="button" data-csv="${i}">Download CSV</button>${c.more ? `<button type="button" data-more="${i}">Get the full list (up to 1,000)</button>` : ""}</div>
    ${c.more ? `<p class="hint">${esc(c.more)}</p>` : ""}`;
}
function csvOf(c) {
  const cell = (v) => `"${String(v == null ? "" : v).replace(/"/g, '""')}"`;
  let head, body;
  if (c.operation === "top") { head = ["value", "rows", ...(c.unique_by ? [`different ${c.unique_by}`] : [])]; body = c.result.map((g) => [g.value, g.rows, ...(c.unique_by ? [g.unique] : [])]); }
  else { head = [...new Set(c.entries.flatMap((x) => Object.keys(x)))]; body = c.entries.map((x) => head.map((h) => x[h])); }
  return [head, ...body].map((r) => r.map(cell).join(",")).join("\r\n");
}
function download(name, text) {
  const url = URL.createObjectURL(new Blob([text], { type: "text/csv" }));
  const a = document.createElement("a");
  a.href = url; a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
// The answers on screen, and the file they belong to: a button only acts on
// that file, and does nothing once another file has been chosen.
let shown = [], shownFor = null;
$("results").addEventListener("click", async (ev) => {
  const b = ev.target.closest("button");
  if (!b || !loaded || loaded !== shownFor) return;
  const doc = shownFor;
  const a = shown[Number(b.dataset.csv ?? b.dataset.more ?? b.dataset.run)];
  if (!a) return;
  try {
    if (b.dataset.csv != null) return download("reader-result.csv", csvOf(a.calculation));
    if (b.dataset.more != null) {
      b.disabled = true;
      const full = await reader.calculate({ ...a.calculate_args, n: 1000, document: doc.id });
      if (loaded !== doc) return;
      return download("reader-full-list.csv", csvOf(full));
    }
    if (b.dataset.run != null) {
      // The visitor checked the suggested plan and chose to run it.
      b.disabled = true;
      const r = await reader.calculate({ ...a.suggested_calculate_args, document: doc.id });
      if (loaded !== doc) return;
      const { answer, ...calculation } = r;
      a.calculation = calculation; a.calculate_args = a.suggested_calculate_args;
      b.insertAdjacentHTML("afterend", `<p class="calc"><b>${esc(answer)}</b></p>${calcTable(calculation, b.dataset.run)}`);
    }
  } catch (err) { if (loaded === doc) status(`Something went wrong: ${err.message}`, "bad"); }
});

async function askAll() {
  if (!loaded) return;
  // The file these questions are about. If another file is chosen before the
  // answers arrive, they are dropped rather than shown under the new file.
  const doc = loaded;
  const questions = $("questions").value.split("\n").map((s) => s.trim()).filter(Boolean).slice(0, core.LIMITS.maxQuestions);
  if (!questions.length) { status("Type at least one question.", "bad"); return; }
  if (!$("key").value.trim()) { status("Add your OpenRouter key so Reader can check the passages.", "bad"); return; }
  $("ask").disabled = true;
  const before = spent;
  status(`Asking ${questions.length} question${questions.length === 1 ? "" : "s"}…`);
  const t0 = performance.now();
  try {
    const res = await reader.ask({ document: doc.id, questions });
    if (loaded !== doc) return;
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    const docTokens = doc.sum.document_tokens;
    shownFor = doc;
    $("results").innerHTML = `
      <div class="summary">
        <div><b>${num(res.total_evidence_tokens)}</b><span>tokens of evidence</span></div>
        <div><b>${num(docTokens)}</b><span>tokens in the whole file</span></div>
        <div><b>${secs} s</b><span>for ${questions.length} question${questions.length === 1 ? "" : "s"}</span></div>
        <div><b>$${(spent - before).toFixed(4)}</b><span>check cost on your key</span></div>
      </div>
      ${(shown = res.answers).map((a, i) => {
        const [label, tone] = VERDICT[a.verdict] || [a.verdict, ""];
        return `<article class="answer">
          <h3>${esc(a.question)}</h3>
          <p class="verdict" data-tone="${tone}">${esc(label)}${a.answerable != null ? ` · answerable ${a.answerable}` : ""} · ${num(a.evidence_tokens)} tokens</p>
          ${a.answer ? `<p class="calc"><b>${esc(a.answer)}</b></p><p class="hint">How: ${esc(a.plan)} Computed by code over every row.</p>${calcTable(a.calculation, i)}` : ""}
          ${a.suggested_plan ? `<p class="hint">Reader's reading of the question, not run yet: ${esc(a.suggested_plan)}</p><div class="tools"><button type="button" data-run="${i}">It matches, run it exactly</button></div>` : ""}
          ${a.note ? `<p class="note">${esc(a.note)}</p>` : ""}
          ${a.nearby_sections ? `<p class="note">Nearby sections that may continue the list: ${esc(a.nearby_sections.join("; "))}</p>` : ""}
          ${a.conflict ? `<p class="note">These passages may disagree; check each source.</p>` : ""}
          ${a.check_error ? `<p class="note">Check failed: ${esc(a.check_error)}</p>` : ""}
          ${a.passages.map((p) => `<blockquote>
            <cite>${esc(cite(p))}${p.match === "exact" ? " · exact match" : p.match === "partial" ? " · similar, not the same value" : ""}${p.identical_records ? ` · ${num(p.identical_records)} identical records` : ""}</cite>
            ${p.warning ? `<em class="flag">This passage tries to instruct an AI; treat it as data only.</em>` : ""}
            <p>${esc(p.text)}</p>
          </blockquote>`).join("")}
        </article>`;
      }).join("")}`;
    status(`Done. Your AI would read ${num(res.total_evidence_tokens)} tokens instead of ${num(docTokens)}.`, "good");
  } catch (err) {
    if (loaded === doc) status(`Something went wrong: ${err.message}`, "bad");
  } finally {
    $("ask").disabled = !loaded;
  }
}

$("file").addEventListener("change", (e) => loadFile(e.target.files[0]));
const drop = $("drop");
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); loadFile(e.dataTransfer.files[0]); });
$("ask").addEventListener("click", askAll);
