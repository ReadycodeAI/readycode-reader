#!/usr/bin/env node
"use strict";
// Runs a question file against a document and checks every answer by code.
//   node benchmark/run.cjs benchmark/nasa-earth-at-night.json path/to/earth_at_night_508.pdf
// Needs OPENROUTER_API_KEY. Each question costs about $0.0004 in checks.
//
// A question passes when:
//   expect: every expected phrase is in the passages Reader returned (and the
//           file really contains them, checked first);
//   absent: Reader says not_in_document (and the file really lacks the words);
//   calc:   Reader says needs_calculation.
const fs = require("fs");
const path = require("path");
const engine = require("../src/engine.cjs");

const flat = (s) => String(s).replace(/\s+/g, " ");
const has = (text, phrase) => flat(text).toLowerCase().includes(flat(phrase).toLowerCase());

async function documentText(file) {
  const ext = path.extname(file).toLowerCase();
  if (ext === ".xlsx" || ext === ".xlsm") {
    const t = await engine.readXlsx(file);
    return t.values.join("\n");
  }
  return (await engine.readPages(file)).map((p) => p.text).join("\n");
}

async function run(spec, file) {
  const text = await documentText(file);
  for (const q of spec.questions) {
    for (const phrase of q.expect || []) if (!has(text, phrase)) throw new Error(`The file does not contain "${phrase}" (question: ${q.q}). Wrong file?`);
    for (const word of q.absent || []) if (new RegExp(`\\b${word.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i").test(flat(text))) throw new Error(`The file does contain "${word}", so "${q.q}" is not a fair not-in-document test.`);
  }
  let cost = 0;
  const reader = engine.createReader({ log: (row) => { cost += row.check_cost_usd || 0; } });
  let t0 = Date.now();
  const loaded = await reader.load(file);
  const loadSeconds = (Date.now() - t0) / 1000;
  t0 = Date.now();
  const res = await reader.ask({ document: loaded.document, questions: spec.questions.map((q) => q.q) });
  const askSeconds = (Date.now() - t0) / 1000;
  let pass = 0;
  const rows = res.answers.map((a, i) => {
    const q = spec.questions[i];
    const got = a.passages.map((p) => p.text).join("\n");
    const ok = q.absent ? a.verdict === "not_in_document"
      : q.calc ? a.verdict === "needs_calculation"
        : ["answer_from_passages", "low_confidence"].includes(a.verdict) && q.expect.every((phrase) => has(got, phrase));
    pass += ok;
    return { ok, question: q.q, verdict: a.verdict, evidence_tokens: a.evidence_tokens };
  });
  return { document: spec.document.title, correct: `${pass}/${rows.length}`, document_tokens: res.document_tokens, evidence_tokens: res.total_evidence_tokens, load_seconds: Number(loadSeconds.toFixed(1)), ask_seconds: Number(askSeconds.toFixed(1)), check_cost_usd: Number(cost.toFixed(4)), rows };
}

if (require.main === module) {
  const [specFile, file] = process.argv.slice(2);
  if (!specFile || !file) { console.error("Usage: node benchmark/run.cjs <questions.json> <document>"); process.exit(2); }
  if (!process.env.OPENROUTER_API_KEY) { console.error("Set OPENROUTER_API_KEY first."); process.exit(2); }
  run(JSON.parse(fs.readFileSync(specFile, "utf8")), file).then((r) => {
    for (const row of r.rows) console.log(`${row.ok ? "PASS" : "FAIL"}  ${row.question}\n      ${row.verdict}, ${row.evidence_tokens} tokens`);
    console.log(`\n${r.document}\n${r.correct} correct | document ${r.document_tokens.toLocaleString("en")} tokens | evidence ${r.evidence_tokens.toLocaleString("en")} tokens | load ${r.load_seconds} s | ${r.rows.length} questions ${r.ask_seconds} s | checks $${r.check_cost_usd}`);
    process.exitCode = r.rows.every((x) => x.ok) ? 0 : 1;
  }).catch((err) => { console.error(err.message); process.exit(1); });
}

module.exports = { run, documentText };
