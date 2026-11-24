import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, "..");
const nm = path.join(rootDir, "node_modules");
const distDir = path.join(rootDir, "dist");

const jszipSrc = fs.readFileSync(path.join(nm, "jszip/dist/jszip.min.js"), "utf8");
const docxSrc = fs.readFileSync(path.join(nm, "docx-preview/dist/docx-preview.min.js"), "utf8");
const pdfSrc = fs.readFileSync(path.join(nm, "pdfjs-dist/build/pdf.min.js"), "utf8");
const pdfWorkerSrc = fs.readFileSync(path.join(nm, "pdfjs-dist/build/pdf.worker.min.js"), "utf8");
const coreSrc = fs.readFileSync(path.join(rootDir, "src/plugin-core.js"), "utf8");

const bundle = `// BetterGravity Office & PDF Preview Plugin (bundled offline)
// Includes JSZip, docx-preview (Apache-2.0), and PDF.js 3.11.174 (Apache-2.0)

(function () {
  const JSZip = (function () {
    const module = { exports: {} };
    const exports = module.exports;
    const define = undefined;
    const require = undefined;
${jszipSrc}
    return module.exports;
  })();

  const docx = (function () {
    const module = { exports: {} };
    const exports = module.exports;
    const define = undefined;
    const require = function (id) {
      if (id === "jszip") return JSZip;
      throw new Error("Unexpected dependency: " + id);
    };
${docxSrc}
    return module.exports;
  })();

  const pdfjsLib = (function () {
    const module = { exports: {} };
    const exports = module.exports;
    const define = undefined;
    const require = undefined;
${pdfSrc}
    return module.exports;
  })();

  const PDF_WORKER_SOURCE = ${JSON.stringify(pdfWorkerSrc)};

${coreSrc}
})();
`;

fs.mkdirSync(path.join(distDir, "styles"), { recursive: true });
fs.writeFileSync(path.join(distDir, "index.js"), bundle, "utf8");
fs.copyFileSync(path.join(rootDir, "plugin.json"), path.join(distDir, "plugin.json"));
fs.copyFileSync(path.join(rootDir, "styles/doc-preview.css"), path.join(distDir, "styles/doc-preview.css"));

console.log("Built dist/index.js (" + (bundle.length / 1024).toFixed(1) + " KB)");
