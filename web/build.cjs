// Builds the browser demo into web/dist: one page, one script and the PDF
// worker. Run: npm run build:web
const fs = require("fs");
const path = require("path");
const esbuild = require("esbuild");
const out = path.join(__dirname, "dist");
fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
esbuild.buildSync({ entryPoints: [path.join(__dirname, "src", "main.js")], bundle: true, format: "esm", minify: true, target: "es2022", platform: "browser", outfile: path.join(out, "reader-web.js"), logLevel: "warning" });
fs.copyFileSync(path.join(__dirname, "index.html"), path.join(out, "index.html"));
fs.copyFileSync(require.resolve("pdfjs-dist/build/pdf.worker.min.mjs"), path.join(out, "pdf.worker.min.mjs"));
for (const f of fs.readdirSync(out)) console.log(`${f}  ${(fs.statSync(path.join(out, f)).size / 1024).toFixed(0)} KB`);
