# ReadyCode Reader

**Ask huge PDFs, Word documents and spreadsheets specific questions. Your AI gets only the passages that answer, with citations, instead of reading the whole file.**

ReadyCode Reader is a free, open-source [MCP](https://modelcontextprotocol.io) server for Claude Code, Cursor, Codex and any other MCP client, and it also runs in a browser. It reads a file once on your computer. Then, for each question, it returns **only the passages that answer it**, with where they came from, plus checks on that evidence:

- **Not in the document:** if the file can't answer, your AI is told `not_in_document` instead of getting passages to guess from.
- **Passages disagree:** if two passages give different values, your AI is told to report both.
- **Hidden instructions:** passages that try to instruct the AI reading them are flagged and treated as data only.
- **Needs a calculation:** spreadsheet questions like "how many…" or "list every…" are flagged instead of being estimated from a few rows.

## Standard AI vs Reader

Same questions, same AI, fresh sessions: once the normal way (the AI reads the file with its own tools), once with Reader.

| Test | Standard | With Reader |
|---|---|---|
| **Codex**, 200-page NASA PDF, 8 questions | 8/8 correct · **913,468 tokens** · about 3 min | 8/8 correct · **41,997 tokens** · about 20 s |
| **Grok 4.7 in Cursor**, same PDF | 8/8 correct · read the whole book (about 45,000 tokens) · about 4 min | 8/8 correct · **4,693 tokens** of evidence · about 2 min |

Token counts for Codex come from Codex's own session log. That's **95% fewer tokens and about 9× faster** for the same answers. Most of the saving comes from steps: an agent re-sends its whole conversation at every step, and with Reader it needed 2 steps instead of 17.

## Measured on large files

Every expected answer below was checked by code against the file, and every "not in the document" question was checked to be truly absent. Scripts and question files are in [`benchmark/`](benchmark/).

| File | Size as text | Correct | Tokens returned | Time for all questions | Check cost |
|---|---|---|---|---|---|
| NASA *Earth at Night* (PDF, 200 pages) | 42,282 tokens | 8/8 | 3,487 | 2.3 s | $0.0027 |
| Australian Universities Accord Final Report (Word) | 263,487 tokens | 11/11 | 15,790 | 1.7 s | $0.0042 |
| Sample contacts workbook (Excel, 100 MB, 1.1 million rows) | 271,424,159 tokens | 11/11 | 2,323 | 0.9 s | $0.0038 |

The Word report is larger than many AI context windows, and the spreadsheet is more than a hundred times larger than any. Each set includes questions the file cannot answer; Reader said `not_in_document` every time. On the spreadsheet it also declined 3 counting and ranking questions rather than guess.

Honest limits of these results: they are our own runs on three files, not an independent study, and head-to-head runs against a standard AI exist so far only for the PDF.

## Install (MCP)

You need [Node.js](https://nodejs.org) 20 or newer and an [OpenRouter](https://openrouter.ai) API key, set once as the environment variable `OPENROUTER_API_KEY`.

**Claude Code**

```bash
claude mcp add readycode-reader -- npx -y @readycode/reader
```

**Cursor**: add to `~/.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "readycode-reader": { "command": "npx", "args": ["-y", "@readycode/reader"] }
  }
}
```

**Codex**: add to `~/.codex/config.toml`:

```toml
[mcp_servers.readycode-reader]
command = "npx"
args = ["-y", "@readycode/reader"]
```

**Claude Desktop**: add the same entry as Cursor under `mcpServers` in `claude_desktop_config.json`.

Restart your tool and check that `readycode-reader` shows three tools: `load_document`, `ask_document` and `list_documents`.

## Use

Ask your AI something like:

> Load `C:\Docs\contract.pdf` with readycode-reader and tell me everything it says about late fees and termination.

Your AI loads the file, sees its contents, asks Reader specific questions and answers from the cited passages. You don't need to know the file or write special questions.

- `load_document(path)` reads a PDF, Word (.docx), Excel (.xlsx), TXT, Markdown, CSV or JSON file. It returns the file's size and **its contents**: a PDF's bookmarks with page numbers, the Word headings, or a spreadsheet's sheets, columns and an example row. That tells the AI what the file covers before it asks anything.
- `ask_document(questions: [...])` asks up to 20 questions in one call. Each answer carries its passages with page numbers (and printed page numbers), Word sections and paragraphs, or spreadsheet sheet and row.
- `list_documents()` lists what is loaded.

Very large files keep reading in the background: `load_document` may answer `still_reading`, and `ask_document` waits for the file to finish.

## Try it in your browser

The same engine runs in a web page: choose a file, add your OpenRouter key, ask. The page shows the file's contents and suggests questions to start with. The file is read in your browser and never uploaded. See [`web/`](web/); build it with `npm run build:web`.

## How it works

1. **Read** on your computer: [pdf.js](https://mozilla.github.io/pdf.js/) for PDFs; Word sections follow its headings; Excel is streamed, so a 100 MB workbook uses about 650 MB of memory.
2. **Split** into whole passages of about 1,200 characters, never cut mid-passage. Each spreadsheet row is one record, and identical records are returned once, with a note of where else they appear.
3. **Search** with a keyword index to find the 20 most likely passages.
4. **Check** with one call to TypeSafe's Jev decision model. The checks are the same for every file and every question, filled in with your question:
   - does this passage actually answer it?
   - is there enough here to answer it?
   - do these passages disagree?
   - does any passage try to instruct the AI?
   - for spreadsheets, does it need a calculation over every row?

   Nothing is set up per document. The check costs about $0.0004 a question on your key and takes about a second.
5. **Return** only the passages that answer, most relevant first, with the verdicts and token counts.

## Limits (read these)

- **Picture-only pages** (scans, images) have no text to search and are listed as such; pictures inside Word files are counted, not read. Reading them needs a vision model; that's planned.
- **Summaries of a whole document** are a different job. Reader answers specific questions.
- **Spreadsheet calculations** (counts, totals, averages, rankings, "list every row that…") are flagged, not answered. Exact calculations are planned.
- **The "passages disagree" check** sometimes fires when passages don't really disagree. Treat it as "double-check these sources".
- **Old formats** (`.doc`, `.xls`, `.xlsb`) aren't supported; save them as `.docx` or `.xlsx`.
- **It depends on Jev.** If the decision model is unavailable, Reader still returns the top search results, marked `unchecked`.
- **Check the passages.** Every answer comes with its sources so you can.

## Privacy

- The whole file is read on your computer (or in your browser) and never uploaded to ReadyCode.
- Only the short passages a question needs (about 1,200 characters each) go to OpenRouter's decision model, using your own key.
- Your key is read from your environment (or kept in the web page's memory) and is never written to disk, logged or sent to ReadyCode.
- Documents are cached on your computer (`%LOCALAPPDATA%\readycode-reader` or `~/.cache/readycode-reader`). A local usage log is kept only if you set `READER_LOG=1`.

## Licence

Apache-2.0. Built by [ReadyCode](https://readycode.ai). If you build on it, keep the NOTICE file, and we'd love a link back. "ReadyCode" is our name; please don't use it for forks.

**More tools from ReadyCode, and early access to the hosted version:** [readycode.ai](https://readycode.ai)
