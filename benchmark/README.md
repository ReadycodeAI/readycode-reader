# Benchmark

Everything needed to check our numbers yourself. Each question file lists the questions and what a right answer must contain; the runner checks every answer by code.

## Run it

```bash
export OPENROUTER_API_KEY=sk-or-...        # your own key; about $0.0004 per question
node benchmark/run.cjs benchmark/nasa-earth-at-night.json path/to/earth_at_night_508.pdf
node benchmark/run.cjs benchmark/universities-accord-word.json "path/to/Australian Universities Accord - Final Report.docx"
node --max-old-space-size=8192 benchmark/excel-lookups.cjs path/to/workbook.xlsx
```

Before asking anything, the runner confirms that the file really contains every expected answer and really lacks the words of every "not in the document" question, so a wrong file or an unfair question stops the run.

A question passes when:

- **answer questions:** Reader's verdict is `answer_from_passages` (or `low_confidence`) and the passages it returned contain every expected phrase;
- **not-in-the-document questions:** the verdict is `not_in_document`;
- **calculation questions** (spreadsheets): the verdict is `calculated` and the number matches the one the benchmark computes itself over every row (the matching row count, or the top value and its row count).

For a question that asks for a complete list, the headings Reader returns in `nearby_sections` count as returned text: they are how a list that runs over several sections is found.

## The files

- **NASA *Earth at Night*** (2019, 200-page PDF, NP-2019-07-2739-HQ): free from nasa.gov. Questions: [`nasa-earth-at-night.json`](nasa-earth-at-night.json), 6 answerable and 2 that the book cannot answer.
- **Australian Universities Accord – Final Report** (Australian Government, Word .docx, about 263,000 tokens): published by the Department of Education. Questions: [`universities-accord-word.json`](universities-accord-word.json), 9 answerable and 2 that the report cannot answer. "Who else was on the Review Panel?" needs all 7 other members, two of them ex-officio.
- **A 100 MB contacts workbook** (1.1 million rows of fake people generated with the Faker library, 4 sheets, columns Name, Email, Phone, Address, Company, Text, Description, Job Title). [`excel-lookups.cjs`](excel-lookups.cjs) picks rows at random (fixed seed) and computes each expected answer over every row: 5 lookups by name, 2 reverse lookups (by email and by phone), 1 name that isn't in the file, and 3 counting, listing or ranking questions whose answers it also computes itself. It works on any workbook with those columns.

## Results (30 September 2026, Jev `typesafe/jev-1.13`)

| File | Correct | Document tokens | Tokens returned | Load | All questions | Check cost |
|---|---|---|---|---|---|---|
| NASA PDF | 8/8 | 42,306 | 4,150 | 0.6 s | 1.9 s | $0.0028 |
| Word report | 11/11 | 263,691 | 15,336 | under 1 s (cached) | 1.2 s | $0.0044 |
| Excel workbook | 11/11 | 271,424,159 (estimate) | 2,494 | 14.7 s | 3.3 s | $0.0022 |

On the workbook, the 8 lookups return only the columns each question needs (13–40 tokens each). The 3 calculations are exact:

- 9 rows at Smith-Hickle, which hold 1 different name.
- 869 Floor Layer rows, 150 different people, listed.
- Roob Inc as the most frequent company, with 286 rows.

The list is 1,934 of the 2,494 tokens.

Since these runs Reader keeps short passages it used to drop (a short last line, a short page) and matches plural and singular forms, so the PDF and Word files return a little more text (3,487 → 4,150 and 14,991 → 15,336 tokens) with the same scores. Earlier results on the same day, before exact calculations, column trimming and the list change: Word 15,790 tokens returned; Excel 2,323 tokens returned, with the 3 calculation questions declined (`needs_calculation`) rather than answered.

## Standard AI vs Reader (NASA PDF)

The same 8 questions were given to each AI twice, in fresh sessions: once reading the PDF with its own tools, once through Reader. The prompts didn't say what Reader was; they asked the AI to measure its own usage.

| AI | Standard | With Reader | Source of the numbers |
|---|---|---|---|
| Codex (ChatGPT) | 8/8 · 913,468 tokens · 17 steps · about 3 min | 8/8 · 41,997 tokens · 2 steps · about 20 s | Codex session log |
| Grok 4.7 (low) in Cursor | 8/8 · whole book read (about 45,000 tokens of text) · about 4 min | 8/8 · 4,693 tokens of evidence · about 2 min | Cursor's own report |

Most of Codex's saving is cached tokens, which cost less, so the cost saving is smaller than the token saving: about 85% rather than 95%.

Head-to-head runs for the Word and Excel files, against Codex and Grok, are in [`head-to-head.md`](head-to-head.md).
