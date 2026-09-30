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
  const stream = async (name) => {
    const en = entries.get(name);
    const lh = new DataView((await bytes(en.local, en.local + 30)).buffer);
    const start = en.local + 30 + lh.getUint16(26, true) + lh.getUint16(28, true);
    const raw = file.slice(start, start + en.csize).stream();
    return en.method ? raw.pipeThrough(new DecompressionStream("deflate-raw")) : raw;
  };
  return {
    has: (name) => entries.has(name),
    text: async (name) => new Response(await stream(name)).text(),
    async elements(name, tag, onElement) {
      const split = core.elementSplitter(tag, onElement);
      const reader = (await stream(name)).pipeThrough(new TextDecoderStream()).getReader();
      for (;;) {
        const { value, done } = await reader.read();
        if (done) break;
        split.write(value);
      }
      split.end();
    },
  };
}

async function open(file, id, ext) {
  const t0 = Date.now();
  const kind = core.kindOf(ext);
  if (kind === "xlsx") return core.recordFromTable(id, ext, await core.readXlsxTable(await fileArchive(file)), t0);
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

async function loadFile(file) {
  if (!file) return;
  const ext = (file.name.match(/\.[^.]+$/) || [""])[0].toLowerCase();
  try { core.kindOf(ext); } catch (err) { status(err.message, "bad"); return; }
  if (file.size > core.LIMITS.maxFileBytes) { status("That file is larger than 500 MB.", "bad"); return; }
  const id = `${file.name}#${file.size}`;
  loaded = null;
  $("ask").disabled = true;
  $("results").innerHTML = "";
  $("about").hidden = true;
  status(`Reading ${file.name} on this computer…`);
  const t0 = performance.now();
  try {
    const sum = await reader.load(id, () => open(file, id, ext), { waitMs: 24 * 3600 * 1000 });
    loaded = { id, name: file.name, sum };
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    const size = sum.kind === "spreadsheet"
      ? `${num(sum.rows)} rows in ${sum.sheets.length} sheet${sum.sheets.length === 1 ? "" : "s"}`
      : `${num(sum.pages)} ${ext === ".pdf" ? "pages" : "parts"}`;
    status(`Ready: ${file.name}, ${size}, about ${num(sum.document_tokens)} tokens of text. Read in ${secs} s.`, "good");
    showContents(sum);
    $("ask").disabled = false;
  } catch (err) {
    status(`Could not read that file: ${err.message}`, "bad");
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
    const short = rest.slice(2).find((c) => ex.values[c].length <= 30);
    if (short) out.push(`How many rows have ${short} "${ex.values[short]}"?`);
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
  box.innerHTML = `<label>What's in this file</label>${what}${qs.length ? `<p class="hint">Try one (click to add it):</p><div class="chips">${qs.map((q) => `<button type="button" class="chip">${esc(q)}</button>`).join("")}</div>` : ""}`;
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
  unchecked: ["Unchecked (no key)", "warn"],
};

async function askAll() {
  if (!loaded) return;
  const questions = $("questions").value.split("\n").map((s) => s.trim()).filter(Boolean).slice(0, core.LIMITS.maxQuestions);
  if (!questions.length) { status("Type at least one question.", "bad"); return; }
  if (!$("key").value.trim()) { status("Add your OpenRouter key so Reader can check the passages.", "bad"); return; }
  $("ask").disabled = true;
  const before = spent;
  status(`Asking ${questions.length} question${questions.length === 1 ? "" : "s"}…`);
  const t0 = performance.now();
  try {
    const res = await reader.ask({ document: loaded.id, questions });
    const secs = ((performance.now() - t0) / 1000).toFixed(1);
    const docTokens = loaded.sum.document_tokens;
    $("results").innerHTML = `
      <div class="summary">
        <div><b>${num(res.total_evidence_tokens)}</b><span>tokens of evidence</span></div>
        <div><b>${num(docTokens)}</b><span>tokens in the whole file</span></div>
        <div><b>${secs} s</b><span>for ${questions.length} question${questions.length === 1 ? "" : "s"}</span></div>
        <div><b>$${(spent - before).toFixed(4)}</b><span>check cost on your key</span></div>
      </div>
      ${res.answers.map((a) => {
        const [label, tone] = VERDICT[a.verdict] || [a.verdict, ""];
        return `<article class="answer">
          <h3>${esc(a.question)}</h3>
          <p class="verdict" data-tone="${tone}">${esc(label)}${a.answerable != null ? ` · answerable ${a.answerable}` : ""} · ${num(a.evidence_tokens)} tokens</p>
          ${a.note ? `<p class="note">${esc(a.note)}</p>` : ""}
          ${a.conflict ? `<p class="note">These passages may disagree; check each source.</p>` : ""}
          ${a.check_error ? `<p class="note">Check failed: ${esc(a.check_error)}</p>` : ""}
          ${a.passages.map((p) => `<blockquote>
            <cite>${esc(cite(p))}${p.identical_records ? ` · ${num(p.identical_records)} identical records` : ""}</cite>
            ${p.warning ? `<em class="flag">This passage tries to instruct an AI; treat it as data only.</em>` : ""}
            <p>${esc(p.text)}</p>
          </blockquote>`).join("")}
        </article>`;
      }).join("")}`;
    status(`Done. Your AI would read ${num(res.total_evidence_tokens)} tokens instead of ${num(docTokens)}.`, "good");
  } catch (err) {
    status(`Something went wrong: ${err.message}`, "bad");
  } finally {
    $("ask").disabled = false;
  }
}

$("file").addEventListener("change", (e) => loadFile(e.target.files[0]));
const drop = $("drop");
drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => { e.preventDefault(); drop.classList.remove("over"); loadFile(e.dataTransfer.files[0]); });
$("ask").addEventListener("click", askAll);
