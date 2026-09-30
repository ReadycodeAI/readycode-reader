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
  { name: "load_document", description: "Load a document (PDF, Word .docx, Excel .xlsx, TXT, MD, CSV or JSON) from a local file path so it can be questioned. Returns its size and tokens and its contents (PDF bookmarks with pages, Word headings, or spreadsheet sheets, columns and an example row), so you can decide what to ask. A very large file may return status still_reading: call ask_document anyway, it waits.", inputSchema: { type: "object", properties: { path: { type: "string", description: "Absolute path to the file." } }, required: ["path"] } },
  { name: "ask_document", description: "Ask questions about a loaded document. Returns only the passages that answer each question, with where they came from (PDF page and printed page, Word section and paragraphs, or spreadsheet sheet and row), how answerable it is from the document (0-1), a note if passages disagree, and token counts. Put several questions in \"questions\" to answer them in one call. Write your answer from these passages and cite their pages. If the verdict is not_in_document, say the document does not contain it.", inputSchema: { type: "object", properties: { question: { type: "string", description: "One question." }, questions: { type: "array", items: { type: "string" }, description: "Several questions at once (up to 20)." }, document: { type: "string", description: "Document id from load_document (optional: the last loaded one)." } } } },
  { name: "list_documents", description: "List the loaded documents.", inputSchema: { type: "object", properties: {} } },
];

async function callTool(name, args) {
  if (name === "load_document") return reader.load((args || {}).path);
  if (name === "ask_document") return reader.ask(args || {});
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
    if (line) handle(line);
  }
});
async function handle(line) {
  let msg;
  try { msg = JSON.parse(line); } catch (_) { return; }
  const { id, method, params } = msg;
  if (id === undefined) return;
  if (method === "initialize") return send({ jsonrpc: "2.0", id, result: { protocolVersion: (params && params.protocolVersion) || "2025-06-18", capabilities: { tools: {} }, serverInfo: { name: "readycode-reader", version: VERSION } } });
  if (method === "ping") return send({ jsonrpc: "2.0", id, result: {} });
  if (method === "tools/list") return send({ jsonrpc: "2.0", id, result: { tools: TOOLS } });
  if (method === "tools/call") {
    try {
      const out = await callTool(params && params.name, params && params.arguments);
      return send({ jsonrpc: "2.0", id, result: { content: [{ type: "text", text: JSON.stringify(out, null, 1) }] } });
    } catch (e) {
      return send({ jsonrpc: "2.0", id, result: { isError: true, content: [{ type: "text", text: String((e && e.message) || e).slice(0, 400) }] } });
    }
  }
  return send({ jsonrpc: "2.0", id, error: { code: -32601, message: `Method not found: ${method}` } });
}
