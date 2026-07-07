import { readFileSync, readdirSync, existsSync } from "node:fs";
import { gzipSync, brotliCompressSync, constants } from "node:zlib";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const distDir = path.join(__dirname, "../../dist");

function fmt(bytes) {
    return (bytes / 1024).toFixed(1).padStart(8) + " KB";
}

const rows = [];
for (const variant of readdirSync(distDir)) {
    const bundle = path.join(distDir, variant, "bundle.js");
    if (!existsSync(bundle) || variant.includes("mock") || variant.includes("aad")) continue;
    const code = readFileSync(bundle);
    const gz = gzipSync(code, { level: 9 });
    const br = brotliCompressSync(code, {
        params: { [constants.BROTLI_PARAM_QUALITY]: 11 },
    });
    rows.push({ variant, raw: code.length, gzip: gz.length, brotli: br.length });
}

rows.sort((a, b) => b.raw - a.raw);
console.log(
    "variant".padEnd(20) + "minified".padStart(11) + "gzip -9".padStart(11) + "brotli".padStart(11)
);
for (const r of rows) {
    console.log(r.variant.padEnd(20) + fmt(r.raw) + fmt(r.gzip) + fmt(r.brotli));
}
