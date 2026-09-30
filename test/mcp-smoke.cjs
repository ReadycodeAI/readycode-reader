"use strict";
// Talks to src/server.cjs over stdio exactly like an MCP client.
const { spawn } = require("child_process");
const path = require("path");
const child = spawn(process.execPath, [path.join(__dirname, "..", "src", "server.cjs")], { stdio: ["pipe", "pipe", "inherit"] });
let buf = ""; const wait = new Map(); let n = 0;
child.stdout.on("data", (d) => { buf += d; let i; while ((i = buf.indexOf("\n")) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); if (wait.has(m.id)) { wait.get(m.id)(m); wait.delete(m.id); } } });
const rpc = (method, params) => new Promise((r) => { const id = ++n; wait.set(id, r); child.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n"); });
const tool = async (name, args) => { const r = await rpc("tools/call", { name, arguments: args }); if (r.result.isError) throw new Error(r.result.content[0].text); return JSON.parse(r.result.content[0].text); };
(async () => {
  const init = await rpc("initialize", { protocolVersion: "2025-06-18", capabilities: {}, clientInfo: { name: "smoke", version: "1" } });
  const tools = await rpc("tools/list", {});
  console.log(init.result.serverInfo.name, init.result.serverInfo.version, "| tools:", tools.result.tools.map((t) => t.name).join(", "));
  let t0 = Date.now();
  const loaded = await tool("load_document", { path: process.argv[2] });
  console.log(`load ${Date.now() - t0} ms: ${loaded.pages} pages, ${loaded.pages_with_text} with text, ${loaded.document_tokens} tokens`);
  const qs = ["How many channels does the VIIRS instrument on Suomi NPP have, and what wavelength range do they cover?", "When did iceberg A-68 break away from the Larsen C ice shelf?", "What causes auroras?", "How many acres had the Okanogan Complex fire burned as of August 20?", "Over which periods was the book's cover image data acquired?", "How much did the Suomi NPP satellite cost to build?", "What had Shanghai's population grown to by 2016?", "Which company built the VIIRS instrument?"];
  t0 = Date.now();
  const r = await tool("ask_document", { questions: qs });
  console.log(`8 questions in one call: ${Date.now() - t0} ms, ${r.total_evidence_tokens} evidence tokens`);
  r.answers.forEach((a, i) => console.log(`  Q${i + 1} ${a.verdict.padEnd(21)} ${a.passages.map((p) => `${p.page}${p.printed_page ? `(p.${p.printed_page})` : ""}`).join(", ")}`));
  child.kill();
})().catch((e) => { console.error("FAIL", e.message); child.kill(); process.exit(1); });
