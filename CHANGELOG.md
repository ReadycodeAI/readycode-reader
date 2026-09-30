# Changelog

## 0.2.0

Fixes from an independent review, each with a regression test:

- **Calculations never answer a different question.**
  - A plan that doesn't account for every part of the question ("salary above 150", "not at Acme", "work on solar panels") is not run, even if the checker approves it.
  - Without a working check, no calculation is run at all.
  - In both cases Reader returns `needs_calculation` with `suggested_calculate_args` for the AI to confirm, complete and run with `calculate`.
- **A name and a company together find that person first.** Before, "Bob's phone at Acme" could return other Acme rows and miss Bob.
- **Short text is kept.** A short page or last line ("The access code is 123456.") used to be dropped.
- **Honest absence.**
  - A question that shares no words with a document now gets `no_matching_text` (the file may use other words), not `not_in_document`.
  - Plural and singular forms now match.
- **Headerless spreadsheets keep their first record.**
  - A first row whose values reappear below is read as data.
  - `load_document` reports which row it used as column names.
  - `headers: "none" | "first_row"` overrides the choice.
- **Checker answers are validated.** An incomplete check is treated as unchecked, not as a row of "no"s. An "enough" score below 0.5 is `low_confidence`.
- **Sturdier MCP server.** Malformed messages (including JSON `null`) get JSON-RPC errors instead of stopping the server, and the protocol version is negotiated.
- **Browser demo.**
  - Each chosen file is its own document, and a slower earlier load can't replace a later one.
  - Loading warnings are shown.
  - Result tables have "Show all", a CSV download and "Get the full list".
  - Suggested calculations can be run with one click after checking them.
  - It has clear links to ReadyCode updates and early access.
- **Privacy.** A Reader-specific [PRIVACY.md](PRIVACY.md) states exactly what is sent to OpenRouter: the question and the candidate passages, not just the ones returned.

## 0.1.0

First release: PDF, Word and Excel through MCP and in the browser, with Jev evidence checks, citations and exact spreadsheet calculations.
