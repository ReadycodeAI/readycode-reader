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
- **calculation questions** (spreadsheets): the verdict is `needs_calculation`.

## The files

- **NASA *Earth at Night*** (2019, 200-page PDF, NP-2019-07-2739-HQ): free from nasa.gov. Questions: [`nasa-earth-at-night.json`](nasa-earth-at-night.json), 6 answerable and 2 that the book cannot answer.
- **Australian Universities Accord – Final Report** (Australian Government, Word .docx, about 263,000 tokens): published by the Department of Education. Questions: [`universities-accord-word.json`](universities-accord-word.json), 9 answerable and 2 that the report cannot answer.
- **A 100 MB contacts workbook** (1.1 million rows of fake people generated with the Faker library, 4 sheets, columns Name, Email, Phone, Address, Company, Text, Description, Job Title). [`excel-lookups.cjs`](excel-lookups.cjs) picks rows at random (fixed seed) and computes each expected answer over every row: 5 lookups by name, 2 reverse lookups (by email and by phone), 1 name that isn't in the file, and 3 counting or ranking questions. It works on any workbook with those columns.

## Results (30 September 2026, Jev `typesafe/jev-1.13`)

| File | Correct | Document tokens | Tokens returned | Load | All questions | Check cost |
|---|---|---|---|---|---|---|
| NASA PDF | 8/8 | 42,282 | 3,487 | 0.6 s | 2.3 s | $0.0027 |
| Word report | 11/11 | 263,487 | 15,790 | under 1 s (cached) | 1.7 s | $0.0042 |
| Excel workbook | 11/11 | 271,424,159 | 2,323 | 14.3 s | 0.9 s | $0.0038 |

## Standard AI vs Reader (NASA PDF)

The same 8 questions were given to each AI twice, in fresh sessions: once reading the PDF with its own tools, once through Reader. The prompts didn't say what Reader was; they asked the AI to measure its own usage.

| AI | Standard | With Reader | Source of the numbers |
|---|---|---|---|
| Codex (ChatGPT) | 8/8 · 913,468 tokens · 17 steps · about 3 min | 8/8 · 41,997 tokens · 2 steps · about 20 s | Codex session log |
| Grok 4.7 (low) in Cursor | 8/8 · whole book read (about 45,000 tokens of text) · about 4 min | 8/8 · 4,693 tokens of evidence · about 2 min | Cursor's own report |

Most of Codex's saving is cached tokens, which cost less, so the cost saving is smaller than the token saving: about 85% rather than 95%.

Head-to-head runs for the Word and Excel files haven't been done yet.
