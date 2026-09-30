# Head-to-head: standard AI vs Reader

Each AI answered the same questions about the same file twice, in fresh chats: once with its normal built-in tools (terminal, scripts, file search), once using only Reader. The prompts didn't say what Reader was. Answers were checked against a key computed by code from the file. Reader's side is also confirmed by its own local log.

## Word: Australian Universities Accord Final Report (263,487 tokens, 11 questions)

| | Codex, standard | Codex, Reader | Cursor (Grok 4.7 low), standard | Cursor (Grok 4.7 low), Reader |
|---|---|---|---|---|
| Correct | 11/11 | 11/11 | 11/11 | 11/11 |
| Tool calls | 4 (one failed on quoting) | 2 | 5 (one search missed the file) | 2 |
| Time | about 3–4 min (1 m 17 s of work shown) | about 30 s end to end; Reader answered all 11 in about 1 s | about 2–3 min | about 1 min |
| Text brought into the chat | many search dumps (not measured) | 14,439 tokens of evidence (5.5% of the report) | about 50–70k characters of search output (estimated) | 15,291 tokens of evidence (5.8%) |
| Citations | none | section and paragraphs on every answer | none | section and paragraphs on every answer |
| The 2 questions the report can't answer | "Not in the document" ✓ | ✓ | ✓ | ✓ |

Neither app shows its exact billed tokens for these runs, so the token comparison is between what each method brought into the chat, not the bill.

**What the AIs themselves said:**

- Codex: "For a large, mostly textual report and a set of factual questions, readycode-reader is the better default." It put the time saving at roughly 80–95%, and Reader's output at 60–75% smaller than its own search dumps.
- Cursor: "The reader was the better shape for this job. One call took all 11 questions, each answer came with a section and paragraph, an answerable score, and a clear not_in_document."

**Where Reader was weaker:**

- **"List everyone" questions.** For "who else was on the panel", Reader's answerable score was only 0.53. With Reader, Codex still found the full list (including ex-officio members Tony Cook and Ben Rimmer, from Appendix B). With Reader, Cursor gave only the six names in the covering letter, while standard search found all eight. Questions whose answer is a complete list need more than the top passages.
- **Output size in Codex.** The 11-answer reply was large enough for Codex to cut off part of it on screen. Fixed after this run: batch output is now compact JSON, with the document details given once.
- **Misleading picture warning.** Reader reported "23 picture-only pages" for the Word file; those were only short paragraphs. Fixed: only PDFs can have picture-only pages.
- **Pictures aren't read.** The report has 44 pictures; anything that exists only inside a picture can't be found, so "not in the document" means "not in its text".
- **Forensic work.** Both AIs noted that direct file inspection is better for exact string counts, formatting, metadata and revision history.

**What this shows:** strong agents with a terminal get these answers right on their own, by searching. Reader's advantage on a report like this is **speed (2–4× faster), fewer steps (2 instead of 4–5), cited answers and clean "not in the document" verdicts**, rather than a large token cut. The large token savings appear when an agent would otherwise read a lot of the file: in the NASA PDF run, Codex used 913,468 tokens normally and 41,997 with Reader.

**A Codex setup note:** Codex exposes MCP tools through tool search. Twice it said Reader wasn't available until the prompt asked it to use tool search to find `readycode-reader`.

## Excel: 100 MB workbook (1.1 million rows, 11 questions)

| | Codex, standard | Codex, Reader | Cursor (Grok 4.7 low), standard | Cursor (Grok 4.7 low), Reader |
|---|---|---|---|---|
| Lookups and the missing name (questions 1–8) | 8/8 | 8/8 | 8/8 | 8/8 |
| Counts and ranking (questions 9–11) | 3/3 (9, 869, Roob Inc 286) | declined (needs_calculation) | 3/3 | declined (needs_calculation) |
| Work | 7 commands plus writing a parser script | 2 calls, no code | 7 steps plus a parser script | 2 calls, no code |
| Time | about 73 s of tool time (the scan itself about 26 s) | about 46 s of tool time | 2–3 min (the scan about 16 s) | 2–3 min, mostly waiting for the load |
| Text returned | about 6–8k tokens (estimated) | 2,323 evidence tokens plus metadata (about 3–4k) | about 48 KB of full rows | 2,323 evidence tokens |

**The honest reading:** agents with a terminal can write a streaming parser and answer everything, including exact counts, in a few minutes. Reader answered every lookup with sheet-and-row citations in two calls and no code. It grouped duplicate records, kept "Nat Becker" apart from "Prof. Nat Becker II", and refused to guess counts. For spreadsheets, Reader wins on ease and setup, not on raw speed, and it loses on counts until exact calculations exist.

**Update, after these runs (our benchmark, not a re-run of the head-to-head):**

Reader now calculates exactly. The same 11 questions score 11/11, with the 3 counts computed by code over all 1.1 million rows:

- 9 rows at Smith-Hickle. Reader adds that they hold only 1 different name: the same person repeated.
- 869 Floor Layer rows, 150 different people.
- Roob Inc, 286 rows.

Lookups now return only the columns a question needs, so the 8 lookups came to 193 tokens in all. The Word panel question now also returns the headings of neighbouring sections, which name all 7 other members, including Jenny Macklin, Fiona Nash and the two ex-officio members. The status of each requested improvement is below.

**Improvements both AIs asked for, in priority order:**

1. **Exact calculations:** count, distinct, top N and sums over whole columns. The person chooses all sheets or one, and rows or unique people; the answer comes back as a number with the sheet and column. *Done: `ask_document` answers these with verdict `calculated`, and the new `calculate` tool takes explicit plans.*
2. **Exact field matches first:** exact name or email matches ahead of partial word matches, with near-matches clearly separated. *Done: rows are marked `match: exact` or `partial`.*
3. **Return only the relevant columns:** a phone question needs the name, the phone and the source, not long Text and Description cells. *Done: `columns_shown` says which.*
4. **Load progress:** `still_reading` should say how far along it is. *Done: sheet, rows read and percent.*
5. **Explain the token figures:** `document_tokens` estimates the file's size as text (about 3.5 characters per token); it is not model usage. *Done: `document_tokens_note`.*
6. **Coverage report:** say which sheets and how many rows were read, and anything skipped, so an absence result can be trusted. *Done: `coverage` on load, and `searched` on spreadsheet "not in the document" answers.*
