/** Minimal static server for dist/ on :4173 (no npx dependency). */
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

// dist/ lives at the package root; this file lives in test/infra/
const root = path.join(
    path.dirname(fileURLToPath(import.meta.url)),
    "../../dist"
);
const types = { ".js": "text/javascript", ".html": "text/html", ".map": "application/json" };

const server = createServer(async (req, res) => {
    let p = new URL(req.url, "http://x").pathname;
    if (p.endsWith(".html") || p.endsWith("/") || p.endsWith("bundle.js")) {
        console.log(new Date().toISOString().slice(11, 19), p);
    }
    if (p.endsWith("/")) p += "index.html";
    try {
        const file = path.join(root, decodeURIComponent(p));
        const data = await readFile(file);
        res.setHeader("Content-Type", types[path.extname(file)] ?? "application/octet-stream");
        res.end(data);
    } catch {
        res.writeHead(404);
        res.end("not found");
    }
});
server.keepAliveTimeout = 120_000;
server.headersTimeout = 125_000;
server.listen(4173, () => console.log("serving dist on :4173"));
