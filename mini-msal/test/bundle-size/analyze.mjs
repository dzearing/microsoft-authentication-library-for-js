/**
 * Attributes minified bytes (and their gzip share) to source files/packages
 * using the bundle's source map — same technique as source-map-explorer.
 *
 * Usage: node analyze.mjs [variant] [--files]
 */
import { readFileSync } from "node:fs";
import { gzipSync } from "node:zlib";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SourceMapConsumer } from "source-map";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const variant = process.argv[2] ?? "msal-stack-only";
const showFiles = process.argv.includes("--files");

const bundlePath = path.join(__dirname, "../../dist", variant, "bundle.js");
const code = readFileSync(bundlePath, "utf8");
const map = readFileSync(bundlePath + ".map", "utf8");
const lines = code.split("\n");

// bucket a source path into a readable group
function groupOf(source) {
    if (!source) return "(unmapped: bundler runtime/glue)";
    const m = source.match(/node_modules\/(@[^/]+\/[^/]+|[^/]+)\/(.*)/);
    if (m) {
        if (showFiles) return `${m[1]}/${m[2]}`;
        // group one level below the package for msal packages
        const sub = m[2].split("/").slice(0, 2).join("/");
        return m[1].startsWith("@azure/msal") ? `${m[1]} ${sub}` : m[1];
    }
    return source.replace(/^webpack:\/\/[^/]*\//, "");
}

const bySource = new Map();
function add(key, bytes) {
    bySource.set(key, (bySource.get(key) ?? 0) + bytes);
}

const consumer = await new SourceMapConsumer(map);
// collect mappings per generated line, then measure spans between them
const perLine = new Map();
consumer.eachMapping((m) => {
    const arr = perLine.get(m.generatedLine) ?? [];
    arr.push(m);
    perLine.set(m.generatedLine, arr);
});

let mappedTotal = 0;
for (const [lineNo, mappings] of perLine) {
    const lineText = lines[lineNo - 1] ?? "";
    mappings.sort((a, b) => a.generatedColumn - b.generatedColumn);
    for (let i = 0; i < mappings.length; i++) {
        const start = mappings[i].generatedColumn;
        const end =
            i + 1 < mappings.length
                ? mappings[i + 1].generatedColumn
                : lineText.length;
        const bytes = Buffer.byteLength(lineText.slice(start, end));
        add(groupOf(mappings[i].source), bytes);
        mappedTotal += bytes;
    }
}
consumer.destroy();

const totalBytes = Buffer.byteLength(code);
add("(unmapped: bundler runtime/glue)", totalBytes - mappedTotal);

const gzipTotal = gzipSync(code, { level: 9 }).length;
const ratio = gzipTotal / totalBytes;

const rows = [...bySource.entries()].sort((a, b) => b[1] - a[1]);
console.log(`\n${variant}: ${(totalBytes / 1024).toFixed(1)} KB minified, ${(gzipTotal / 1024).toFixed(1)} KB gzip\n`);
console.log(
    "source".padEnd(64) + "min KB".padStart(9) + "~gz KB".padStart(9) + "%".padStart(7)
);
for (const [src, bytes] of rows) {
    if (bytes < 200 && !showFiles) continue;
    console.log(
        src.slice(0, 63).padEnd(64) +
            (bytes / 1024).toFixed(1).padStart(9) +
            ((bytes * ratio) / 1024).toFixed(1).padStart(9) +
            ((bytes / totalBytes) * 100).toFixed(1).padStart(6) + "%"
    );
}
