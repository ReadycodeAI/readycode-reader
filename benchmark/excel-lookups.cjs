#!/usr/bin/env node
"use strict";
// Builds a question file from any large people/contacts spreadsheet with
// Name, Email, Phone, Address, Company and Job Title columns, with every
// expected answer computed by code over all rows, then runs it.
//   node --max-old-space-size=8192 benchmark/excel-lookups.cjs path/to/100mb.xlsx
// The benchmark used a synthetic 100 MB sample workbook (fake people made
// with the Faker library, 1.1 million rows in 4 sheets).
const engine = require("../src/engine.cjs");
const { run } = require("./run.cjs");

(async () => {
  const file = process.argv[2];
  if (!file) { console.error("Usage: node benchmark/excel-lookups.cjs <workbook.xlsx>"); process.exit(2); }
  if (!process.env.OPENROUTER_API_KEY) { console.error("Set OPENROUTER_API_KEY first."); process.exit(2); }
  const t = await engine.readXlsx(file);
  const N = t.rowStart.length - 1;
  const row = (r) => { const o = {}; const h = t.sheets[t.rowSheet[r]].headers; for (let k = t.rowStart[r]; k < t.rowStart[r + 1]; k++) o[h[t.cellCol[k]]] = t.values[t.cellVal[k]]; return o; };
  const where = (field, value) => { const out = []; for (let r = 0; r < N; r++) { const o = row(r); if (o[field] === value) out.push(o); } return out; };
  let seed = 7;
  const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
  const pick = () => row(Math.floor(rnd() * N));
  const people = [];
  while (people.length < 5) { const o = pick(); if (o.Name && o.Phone && o.Company && o["Job Title"] && o.Email && o.Address) people.push(o); }
  const asks = [["Phone", (n) => `What is the phone number of ${n}?`], ["Company", (n) => `Which company does ${n} work for?`], ["Job Title", (n) => `What is the job title of ${n}?`], ["Email", (n) => `What is the email address of ${n}?`], ["Address", (n) => `What is the address of ${n}?`]];
  const questions = people.map((p, i) => ({ q: asks[i][1](p.Name), expect: [...new Set(where("Name", p.Name).map((o) => o[asks[i][0]]))] }));
  const byEmail = pick(), byPhone = pick();
  questions.push({ q: `Who has the email address ${byEmail.Email}?`, expect: [...new Set(where("Email", byEmail.Email).map((o) => o.Name))] });
  questions.push({ q: `Whose phone number is ${byPhone.Phone}?`, expect: [...new Set(where("Phone", byPhone.Phone).map((o) => o.Name))] });
  const trap = "Zebulon Quackenbush-Vantablack";
  if (where("Name", trap).length) throw new Error("The trap name exists in this file.");
  questions.push({ q: `What is the phone number of ${trap}?`, absent: ["Quackenbush"] });
  const company = pick().Company, title = pick()["Job Title"];
  questions.push({ q: `How many people in the file work at ${company}?`, calc: true, truth: where("Company", company).length });
  questions.push({ q: `List everyone whose job title is ${title}.`, calc: true, truth: where("Job Title", title).length });
  const counts = new Map();
  for (let r = 0; r < N; r++) { const c = row(r).Company; if (c) counts.set(c, (counts.get(c) || 0) + 1); }
  const [topValue, topRows] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
  questions.push({ q: "Which company appears most often in the spreadsheet?", calc: true, top: { value: topValue, rows: topRows } });
  const r = await run({ document: { title: `${file.split(/[\\/]/).pop()}: ${N.toLocaleString("en")} rows` }, questions }, file);
  for (const row of r.rows) console.log(`${row.ok ? "PASS" : "FAIL"}  ${row.question}\n      ${row.verdict}, ${row.evidence_tokens} tokens${row.answer ? `\n      ${row.answer}${row.expected ? ` (expected ${row.expected})` : ""}` : ""}`);
  console.log(`\n${r.document}\n${r.correct} correct | document ${r.document_tokens.toLocaleString("en")} tokens | evidence ${r.evidence_tokens.toLocaleString("en")} tokens | load ${r.load_seconds} s | ${r.rows.length} questions ${r.ask_seconds} s | checks $${r.check_cost_usd}`);
  process.exitCode = r.rows.every((x) => x.ok) ? 0 : 1;
})().catch((err) => { console.error(err.message); process.exit(1); });
