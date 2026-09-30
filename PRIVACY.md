# ReadyCode Reader: privacy

This covers ReadyCode Reader: the MCP server (`@readycode/reader`) and the browser demo. It does not cover other ReadyCode products, which have their own policies.

## In one line

Your file is read on your own computer, or in your own browser. ReadyCode never receives your file, your questions or your key. The only thing that leaves your machine is what goes to OpenRouter to check each question's evidence, sent with your own OpenRouter key.

## What stays on your computer

- **The whole file.** It is read and searched locally, and never uploaded to ReadyCode or anyone else.
- **A cache of text documents** (PDF, Word and text files), so they open faster next time. It is stored in `%LOCALAPPDATA%\readycode-reader` (Windows) or `~/.cache/readycode-reader`, and you can delete it at any time. Spreadsheets are not cached. The browser demo keeps nothing after you close the page.
- **A usage log**, only if you turn it on with `READER_LOG=1`. It records each question, its verdict and token counts, and stays in the same local folder.

## What is sent to OpenRouter, for each question

Reader checks each answer with TypeSafe's Jev decision model (`typesafe/jev-1.13`) through OpenRouter's decisions API, using **your** OpenRouter key. Each question sends:

- **the question**;
- **the candidate passages the local search found**, before any are filtered: up to about 20, or 30 for "list everyone" questions. Each is about 1,200 characters of document text, or one spreadsheet row of up to about 3,000 characters (often fewer columns, when the question only needs some);
- **for spreadsheet calculation questions**, a one-line description of the calculation Reader plans to run. When no row matched the question's words, the sheet's column names are sent instead of rows.

Reader sends nothing else from the file. Exact calculations (counts, totals, lists) and the explicit `calculate` tool run entirely on your computer and send nothing. OpenRouter's and the model provider's own terms and privacy policies apply to what they receive; see [openrouter.ai/privacy](https://openrouter.ai/privacy). If a document is confidential, check whether those terms allow it before asking questions about it.

Requests identify the app as "ReadyCode Reader" (standard OpenRouter attribution headers). They carry no information about you beyond your own key.

## Your key

- **MCP server:** the key is read from the `OPENROUTER_API_KEY` environment variable, or, on Windows, from your user environment settings. It is kept in memory only: never written to disk, logged, or sent anywhere except OpenRouter.
- **Browser demo:** the key you paste stays in that page's memory. It is not saved, not put in the address bar, and sent only to OpenRouter.

## The browser demo's hosting

The demo is a static page hosted on GitHub Pages. GitHub may log standard web-server data (such as IP addresses) for visits; see GitHub's privacy statement. The page has no analytics or tracking scripts and sets no cookies. Links to readycode.ai take you to our website, which has its own privacy policy.

## Contact

Questions: readycodeai@gmail.com, or open an issue at [github.com/ReadycodeAI/readycode-reader](https://github.com/ReadycodeAI/readycode-reader/issues).
