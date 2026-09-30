#!/usr/bin/env node
"use strict";
// ReadyCode Reader as a local MCP server (stdio). Claude Code, Cursor, Codex
// and other MCP clients start it; it loads documents from this computer and
// returns only the passages that answer each question.
const fs = require("fs");
const path = require("path");
const os = require("os");
const { execFileSync } = require("child_process");
const { createReader } = require("./engine.cjs");

const VERSION = require("../package.json").version;

// The key comes from OPENROUTER_API_KEY. Some MCP clients pass only a few
// environment variables; on Windows it is then read from the user's own
// environment settings. It is kept in memory only, never printed or logged.
let cachedKey = null;
function apiKey() {
  if (cachedKey != null) return cachedKey;
  cachedKey = process.env.OPENROUTER_API_KEY || "";
  if (!cachedKey && process.platform === "win32") {
    try {
      cachedKey = execFileSync("powershell.exe", ["-NoProfile", "-Command", "[Environment]::GetEnvironmentVariable('OPENROUTER_API_KEY','User')"], { encoding: "utf8", windowsHide: true, stdio: ["ignore", "pipe", "ignore"] }).trim();
    } catch (_) { cachedKey = ""; }
  }
  return cachedKey;
}

// Local usage log, only when READER_LOG=1 (questions stay on this computer).
const LOG = process.env.READER_LOG === "1" ? path.join(process.env.LOCALAPPDATA || path.join(os.homedir(), ".cache"), "readycode-reader", "usage.jsonl") : null;
const reader = createReader({ apiKey, log: LOG ? (row) => { fs.mkdirSync(path.dirname(LOG), { recursive: true }); fs.appendFileSync(LOG, `${JSON.stringify(row)}\n`); } : null });

const TOOLS = [
  { name: "load_document", description: "Load a document (PDF, Word .docx, Excel .xlsx, TXT, MD, CSV or JSON) from a local file path so it can be questioned. Returns its size, its contents (PDF bookmarks with pages, Word headings, or spreadsheet sheets, columns and an example row) and what was read (coverage), so you can decide what to ask. document_tokens estimates the whole file as text; it is not model usage. A very large file may return status still_reading with progress: call ask_document anyway, it waits.", inputSchema: { type: "object", properties: { path: { type: "string", description: "Absolute path to the file." }, headers: { type: "string", enum: ["auto", "first_row", "none"], description: "Spreadsheets only: how to treat each sheet's first row. auto (default) uses it as column names when it looks like labels; none reads every row as data; first_row always uses it as column names." } }, required: ["path"] } },
  { name: "ask_document", description: "Ask questions about a loaded document. Returns the passages judged to answer each question, with where they came from (PDF page and printed page, Word section and paragraphs, or spreadsheet sheet and row), how answerable it is from the document (0-1), a note if passages disagree, and token counts. Put several questions in \"questions\" to answer them in one call. Write your answer from these passages and cite their pages. If the verdict is not_in_document, say the document does not appear to contain it. no_matching_text means no passage shares a word with the question: ask again in other words (synonyms, abbreviations) before concluding it is absent. unchecked means the evidence check could not run, or skipped passages (marked checked: false): use the passages with care, and don't conclude absence. Spreadsheets: rows show only the columns the question needs (columns_shown); match \"exact\" rows hold the value named in the question, \"partial\" rows only share words with it (a different, longer name, say) and must not be merged with them. Counting, ranking, totals and \"list every\" questions come back with verdict calculated: a number computed by code over every row, with the plan used; report it with the sheets and column. needs_calculation gives no number: if it includes suggested_calculate_args, check they cover the whole question (add any missing conditions) and run them with the calculate tool; otherwise build the calculation yourself.", inputSchema: { type: "object", properties: { question: { type: "string", description: "One question." }, questions: { type: "array", items: { type: "string" }, description: "Several questions at once (up to 20)." }, document: { type: "string", description: "Document id from load_document (optional: the last loaded one)." } } } },
  { name: "calculate", description: "Exact calculations over a loaded spreadsheet, computed by code over every row; nothing is estimated. Operations: count (rows), distinct (different values in a column), top (values with the most, or fewest, rows or different people), sum, average, min, max, and list (every matching row, or each different person with how many rows they have). Filter with conditions, choose one sheet or all, and set unique_by (for example Name) to count people rather than rows. ask_document already answers plainly worded counting and ranking questions this way; use calculate when it returns needs_calculation or to adjust a calculation. Returns the number with the sheets, column and conditions used, and rows cited as sheet and row.", inputSchema: { type: "object", properties: {
    operation: { type: "string", enum: ["count", "distinct", "top", "sum", "average", "min", "max", "list"] },
    column: { type: "string", description: "The column to count distinct values of, rank, add up, average, or find the lowest or highest in (a header name, or a letter for sheets without headers)." },
    where: { type: "array", description: "Conditions every counted row must meet (all of them).", items: { type: "object", properties: { column: { type: "string" }, equals: { type: "string" }, contains: { type: "string" }, not_equals: { type: "string" }, above: { type: "number" }, below: { type: "number" }, at_least: { type: "number" }, at_most: { type: "number" } }, required: ["column"] } },
    sheet: { type: "string", description: "One sheet name, or \"all\" (default)." },
    unique_by: { type: "string", description: "Also count different values of this column (for example Name, to count people rather than rows); list gives one entry per value." },
    rank_by: { type: "string", enum: ["rows", "unique"], description: "For top: rank by rows (default) or by different unique_by values." },
    order: { type: "string", enum: ["most", "fewest"], description: "For top (default most)." },
    columns: { type: "array", items: { type: "string" }, description: "For list: columns to show for each entry." },
    n: { type: "number", description: "For top: how many values (default 5, up to 100). For list: how many entries (default 100, up to 1,000)." },
    document: { type: "string", description: "Document id from load_document (optional: the last loaded one)." },
  }, required: ["operation"] } },
  { name: "list_documents", description: "List the loaded documents.", inputSchema: { type: "object", properties: {} } },
];

async function callTool(name, args) {
  if (name === "load_document") return reader.load((args || {}).path, { headers: (args || {}).headers || "auto" });
  if (name === "ask_document") return reader.ask(args || {});
  if (name === "calculate") return reader.calculate(args || {});
  if (name === "list_documents") return reader.list();
  throw new Error(`Unknown tool: ${name}`);
}

const send = (msg) => process.stdout.write(`${JSON.stringify(msg)}\n`);
let buffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  buffer += chunk;
  let nl;
  while ((nl = buffer.indexOf("\n")) >= 0) {
    const line = buffer.slice(0, nl).trim();
    buffer = buffer.slice(nl + 1);
    if (line) handle(line).catch((e) => fail(null, -32603, `Internal error: ${String((e && e.message) || e).slice(0, 200)}`));
  }
});
// MCP protocol versions this server speaks, newest first. A client asking for
// another version is offered the newest; it decides whether to continue.
const PROTOCOLS = ["2025-06-18", "2025-03-26", "2024-11-05"];
const fail = (id, code, message) => send({ jsonrpc: "2.0", id, error: { code, message } });

// Every message is checked before use: bad input gets a JSON-RPC error (or is
// ignored, for notifications) and never stops the server.
async function handle(line) {
  let msg;
  try { msg = JSON.parse(line); } catch (_) { return fail(null, -32700, "Parse error: not valid JSON."); }
  if (Array.isArray(msg)) return fail(null, -32600, "Batches are not supported; send one request per line.");
  if (!msg || typeof msg !== "object") return fail(null, -32600, "Invalid request: expected a JSON-RPC object.");
  const { id, method, params } = msg;
  const validId = typeof id === "string" || (typeof id === "number" && Number.isFinite(id));
  if (id === undefined) return; // a notification: nothing to answer
  if (!validId) return fail(null, -32600, "Invalid request: id must be a string or a number.");
  if (msg.jsonrpc !== "2.0" || typeof method !== "string") return fail(id, -32600, "Invalid request: needs jsonrpc \"2.0\" and a method.");
  if (params != null && typeof params !== "object") return fail(id, -32602, "Invalid params: expected an object.");
  if (method === "initialize") {
    const asked = params && params.protocolVersion;
    return send({ jsonrpc: "2.0", id, result: { protocolVersion: PROTOCOLS.includes(asked) ? asked : PROTOCOLS[0], capabilities: { tools: {} }, serverInfo: { name: "readycode-reader", version: VERSION } } });
  }
  if (method === "ping") return send({ jsonrpc: "2.0", id, result: {} });
  if (method === "tools/list") return send({ jsonrpc: "2.0", id, result: { tools: TOOLS } });
  if (method === "tools/call") {
    if (!params || typeof params.name !== "string") return fail(id, -32602, "Invalid params: tools/call needs a tool name.");
    if (params.arguments != null && (typeof params.arguments !== "object" || Array.isArray(params.arguments))) return fail(id, -32602, "Invalid params: arguments must be an object.");
    try {
      const out = await callTool(params && params.name, params && params.arguments);
      return send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text: JSON.stringify(out) }] } });
    } catch (e) {
      return send({ jsonrpc: "2.0", id, result: { isError: true, content: [{ type: "text", text: String((e && e.message) || e).slice(0, 400) }] } });
    }
  }
  return send({ jsonrpc: "2.0", id, error: { code: -32601, message: `Method not found: ${method}` } });
}
