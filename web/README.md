# Reader in the browser

The same engine as the MCP server (`src/core.cjs`), running in a web page. The visitor chooses a file and adds their own OpenRouter key; the file is read in the browser and never uploaded, and only the short passages each question needs go from the page to OpenRouter.

```bash
npm install
npm run build:web      # writes web/dist: index.html, reader-web.js, pdf.worker.min.mjs
```

Host the three files in `web/dist` on any static site (they must be served over http(s), not opened as a file). No server code is needed, so it costs nothing however many people use it.
