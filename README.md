# ReadyCode Reader (Readycode.AI)

**Ask huge PDFs, Word documents and spreadsheets specific questions. Your AI gets the passages that answer, checked and cited, instead of reading the whole file.**

[Website](https://readycode.ai/reader) · [Get updates and early access](https://readycode.ai/reader#get-started) · [Live browser demo](https://readycodeai.github.io/readycode-reader/) · [Install from npm](https://www.npmjs.com/package/@readycode/reader)

[ReadyCode.AI](https://readycode.ai/reader) has contributed ReadyCode Reader - a free, open-source [MCP](https://modelcontextprotocol.io) server for Claude Code, Cursor, Codex and any other MCP client, and it also runs in a browser. It reads a file once on your computer. Then, for each question, it returns **the passages judged to answer it**, with where they came from, plus checks on that evidence:

- **Not in the document:** if none of the passages checked supports an answer, your AI is told `not_in_document` instead of getting passages to guess from. If no passage even shares a word with the question, it is told `no_matching_text`, since the file may use other words.
- **Passages disagree:** when the check finds two relevant passages giving different values, your AI is told to report both.
- **Hidden instructions:** passages the check judges to be instructing the AI reading them are flagged, so your AI treats them as data only.
- **Exact spreadsheet calculations:** "how many…", "which … appears most often", totals, averages and "list everyone who…" are computed by code over every row and come back as numbers, with the sheets, column and conditions used. A model never estimates them.
- **Exact matches first:** a spreadsheet row holding the exact name you asked about comes before look-alikes ("Nat Becker" before "Prof. Nat Becker II"), and rows are trimmed to the columns the question names.

## Standard AI vs Reader

Same questions, same AI, fresh sessions: once the normal way (the AI reads the file with its own tools), once with Reader.

| Test | Standard | With Reader |
|---|---|---|
| **Codex**, 200-page NASA PDF, 8 questions | 8/8 correct · **913,468 tokens** · about 3 min | 8/8 correct · **41,997 tokens** · about 20 s |
| **Grok 4.7 in Cursor**, same PDF | 8/8 correct · read the whole book (about 45,000 tokens) · about 4 min | 8/8 correct · **4,693 tokens** of evidence · about 2 min |
| **Codex** and **Grok 4.7 in Cursor**, 263,000-token Word report, 11 questions | 11/11 · 4–5 tool calls · 2–4 min · no citations | 11/11 · 2 calls · 30 s–1 min · section and paragraph on every answer |
| **Codex** and **Grok 4.7 in Cursor**, 100 MB Excel workbook (1.1 million rows), 11 questions | 11/11 · wrote its own parser script · 7 steps | 8/8 lookups in 2 calls, no code; declined the 3 counting questions (since fixed, see below) |

Token counts for Codex come from Codex's own session log. On the PDF that's **95% fewer tokens and about 9× faster** for the same answers. Most of the saving comes from steps: an agent re-sends its whole conversation at every step, and with Reader it needed 2 steps instead of 17. On the Word report the gain was speed, fewer steps and citations rather than tokens. Full results: [`benchmark/head-to-head.md`](benchmark/head-to-head.md).

## Measured on large files

Every expected answer below was checked by code against the file, and every "not in the document" question was checked by code: its key words appear nowhere in the file. Scripts and question files are in [`benchmark/`](benchmark/).

| File | Size as text | Correct | Tokens returned | Time for all questions | Check cost |
|---|---|---|---|---|---|
| NASA *Earth at Night* (PDF, 200 pages) | 42,306 tokens | 8/8 | 3,818 | 1.8 s | $0.0028 |
| Australian Universities Accord Final Report (Word) | 263,691 tokens | 11/11 | 15,687 | 0.7 s | $0.0044 |
| Sample contacts workbook (Excel, 100 MB, 1.1 million rows) | 271,424,159 tokens (estimate) | 11/11 | 2,494 | 2.4 s | $0.0022 |

Tokens returned vary by a few percent from run to run, because the checker's relevance scores vary slightly; the scores above have been the same in every run.

The Word report is larger than many AI context windows, and the spreadsheet is more than a hundred times larger than any. Each set includes questions the file cannot answer; Reader said `not_in_document` every time.

On the spreadsheet, 8 questions are lookups. Each is answered with only the name and the column asked about, 13–40 tokens each. The other 3 are exact calculations over all 1.1 million rows:

- "How many people work at Smith-Hickle?": 9 rows, and Reader adds that they hold only 1 different name.
- "List everyone whose job title is Floor Layer": 869 rows, 150 different people, listed.
- "Which company appears most often?": Roob Inc, 286 rows.

In the Word set, "Who else was on the panel?" needs all 7 other members, including two ex-officio members found only through section headings.

Honest limits of these results: they are our own runs on three files, not an independent study. The head-to-head runs above were made before exact calculations existed and have not been re-run since.

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

Restart your tool and check that `readycode-reader` shows four tools: `load_document`, `ask_document`, `calculate` and `list_documents`.

## Use

Ask your AI something like:

> Load `C:\Docs\contract.pdf` with readycode-reader and tell me everything it says about late fees and termination.

Your AI loads the file, sees its contents, asks Reader specific questions and answers from the cited passages. You don't need to know the file or write special questions.

- `load_document(path)` reads a PDF, Word (.docx), Excel (.xlsx), TXT, Markdown, CSV or JSON file. It returns the file's size and **its contents**: a PDF's bookmarks with page numbers, the Word headings, or a spreadsheet's sheets, columns and an example row. That tells the AI what the file covers before it asks anything. For spreadsheets it says which row it took as column names; `headers: "none"` reads every row as data, and `headers: "first_row"` always uses row 1.
- `ask_document(questions: [...])` asks up to 20 questions in one call. Each answer carries its passages with page numbers (and printed page numbers), Word sections and paragraphs, or spreadsheet sheet and row.
- `calculate(operation, column, where, sheet, unique_by, n)` runs an exact calculation over a loaded spreadsheet: `count`, `distinct`, `top` (most or fewest), `sum`, `average`, `min`, `max` or `list`. It takes conditions (equals, contains, not equals, above, below), one sheet or all, and can count people instead of rows (`unique_by: "Name"`). You rarely need it: `ask_document` turns plainly worded questions into these plans, has each plan checked against the question, and answers with verdict `calculated`.
- `list_documents()` lists what is loaded.

A calculated answer looks like this:

```json
{ "question": "How many people in the file work at Smith-Hickle?", "verdict": "calculated",
  "answer": "9 rows where Company is \"Smith-Hickle\" in all 4 sheets, holding 1 different Name value.",
  "plan": "Count the rows where Company is \"Smith-Hickle\" in all sheets, and count the different Name values among them.",
  "calculation": { "operation": "count", "unique_by": "Name", "where": ["Company is \"Smith-Hickle\""], "sheets": "all",
    "rows_checked": 1116271, "matching_rows": 9,
    "matching_rows_by_sheet": { "Worksheet (2)": 0, "Worksheet (3)": 0, "Tablo3": 1, "Worksheet": 8 },
    "result": 9, "unique": 1, "method": "Computed by code over every row; nothing estimated." } }
```

Each answer has a verdict:

| Verdict | Meaning |
|---|---|
| `answer_from_passages` | The passages returned answer the question. |
| `low_confidence` | The passages may not fully answer it; answer only what they state. |
| `not_in_document` | None of the passages checked supports an answer. For a spreadsheet, no row holds the words asked about. |
| `no_matching_text` | No passage shares a word with the question. Ask again with other words before concluding it isn't there. |
| `calculated` | An exact number computed by code over every row, with the plan used and `calculate_args` to re-run or adjust it. |
| `needs_calculation` | A calculation is needed but Reader could not confirm one that covers the whole question (for example "salary above 150" or "not at Acme"). It gives no number; it offers `suggested_calculate_args` for your AI to check, complete and run with `calculate`. |
| `unchecked` | The check could not run (no key, or the service failed). Search results are returned as they are, and no calculation is run. |

Very large files keep reading in the background: `load_document` may answer `still_reading` with how far it has got, and `ask_document` waits for the file to finish. `load_document` also reports what was read (every sheet and row, and anything skipped), so you can see what a "not in the document" answer covered. `document_tokens` is an estimate of the whole file's size as text (about 3.5 characters per token), for comparison; it is not model usage.

## Try it in your browser

The same engine runs in a web page: choose a file, add your OpenRouter key, ask. The page shows the file's contents and suggests questions to start with. The file is read in your browser and never uploaded. See [`web/`](web/); build it with `npm run build:web`.

## How it works

1. **Read** on your computer: [pdf.js](https://mozilla.github.io/pdf.js/) for PDFs; Word sections follow its headings; Excel is streamed, so a 100 MB workbook uses about 650 MB of memory.
2. **Split** at line breaks into passages of about 1,200 characters, so a passage usually holds whole lines and paragraphs; a single line longer than about 1,400 characters (common in some PDFs) is cut into pieces. Short text is kept. Each spreadsheet row is one record, and identical records are returned once, with a note of where else they appear.
3. **Search** with a keyword index to find the 20 most likely passages.
4. **Check** with one call to TypeSafe's Jev decision model. The checks are the same for every file and every question, filled in with your question:
   - does this passage actually answer it?
   - is there enough here to answer it?
   - do these passages disagree?
   - does any passage try to instruct the AI?
   - for spreadsheets, does it need a calculation over every row, and does the calculation plan Reader made fit the question?

   Nothing is set up per document. The check costs about $0.0004 a question on your key and takes about a second.
5. **Calculate** (spreadsheets): when a question needs every row, code runs the plan over the whole workbook. Values that differ only in capitals or spacing count as one, and rows can be counted as rows or as different people. A 1.1-million-row workbook takes about 1–2 seconds.
6. **Return** the passages judged to answer, most relevant first, with the verdicts and token counts. For "list everyone" questions in Word files, the headings of neighbouring sections come too, since a list often runs over several sections.

## Limits (read these)

- **Picture-only pages** (scans, images) have no text to search and are listed as such; pictures inside Word files are counted, not read. Reading them needs a vision model; that's planned.
- **Summaries of a whole document** are a different job. Reader answers specific questions.
- **Spreadsheet calculations** cover counts, distinct values, rankings, totals, averages, lowest and highest values, and lists, with simple conditions. Formulas, pivot tables, date ranges and combining sheets with different columns are not supported yet. When Reader can't turn a question into a plan it is sure of, it answers `needs_calculation` rather than guess, and your AI can call `calculate` directly.
- **The "passages disagree" check** now needs two relevant passages about the same thing, but it can still fire when passages don't really disagree. Treat it as "double-check these sources".
- **Old formats** (`.doc`, `.xls`, `.xlsb`) aren't supported; save them as `.docx` or `.xlsx`.
- **It depends on Jev.** If the decision model is unavailable, Reader still returns the top search results, marked `unchecked`, and runs no calculation on its own: it suggests one for your AI (or you) to confirm.
- **Search is by words.** Plural and singular forms match, but synonyms don't ("car" won't find "automobile"). That's why a question sharing no words with the file gets `no_matching_text`, not `not_in_document`.
- **Header rows are guessed.** A first row of short, distinct labels is taken as column names unless its values appear again below. `load_document` says which row it used, and you can override it.
- **Check the passages.** Every answer comes with its sources so you can.

## Privacy

- The whole file is read on your computer (or in your browser) and never uploaded to ReadyCode. ReadyCode never receives your file, your questions or your key.
- For each question, the question and the candidate passages found by the local search go to OpenRouter's decision model with your own key. That's up to about 20 passages of about 1,200 characters, or spreadsheet rows. For calculation questions, a one-line description of the planned calculation (or the column names) goes too.
- Exact calculations run entirely on your computer and send nothing.
- Your key is read from your environment (or kept in the web page's memory) and is never written to disk, logged, or sent anywhere but OpenRouter.
- Text documents are cached on your computer (`%LOCALAPPDATA%\readycode-reader` or `~/.cache/readycode-reader`). A local usage log is kept only if you set `READER_LOG=1`.

Full details: [PRIVACY.md](PRIVACY.md).

## Credits and related work

Reader combines established pieces rather than inventing a new retrieval method. The search-then-check pattern (a keyword shortlist, then Jev relevance judgments) follows TypeSafe's own guidance for Jev. Others have built Jev-based PDF search, MCP document retrieval and MCP spreadsheet operations. What Reader adds is one small, local tool covering PDF, Word and Excel. It combines evidence checks (relevance, enough to answer, disagreement, hidden instructions), citations, honest "not found" answers and exact spreadsheet calculations, through MCP or in a browser.

## Licence

Apache-2.0. Built by [ReadyCode](https://readycode.ai). If you build on it, keep the NOTICE file, and we'd love a link back. "ReadyCode" is our name; please don't use it for forks.

**Source:** [github.com/ReadycodeAI/readycode-reader](https://github.com/ReadycodeAI/readycode-reader). Issues and ideas are welcome there.

**News about Reader and new ReadyCode tools, and early access to the hosted version:** [readycode.ai/reader](https://readycode.ai/reader)
