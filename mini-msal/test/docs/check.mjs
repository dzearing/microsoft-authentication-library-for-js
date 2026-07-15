/**
 * Doc-sample check (D6): every fenced ```ts / ```tsx code block in the
 * consumer docs must type-check (strict) against the packages' built dist
 * types — copy-paste-runnable is enforced mechanically, not by review.
 *
 * A sample that is intentionally partial can opt out with an HTML comment
 * on the line right above its fence:  <!-- docs-check:skip <reason> -->
 * Keep those rare and justified.
 *
 *   npm run docs:check          (no network; builds package dist via tsc)
 */
import { execFileSync } from "node:child_process";
import { mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const TSC = path.join(ROOT, "node_modules", ".bin", "tsc");
const DOCS = ["README.md", "docs/README.md", "docs/UPGRADING.md", "docs/ALACARTE.md", "docs/SIZE.md"];
const SCRATCH = path.join(ROOT, "test", "docs", ".samples");

// 1. fresh dist (browser first: compat/react need its d.ts)
for (const p of ["browser", "compat", "react"]) {
    console.log(`tsc: packages/${p}`);
    execFileSync(TSC, ["-p", "tsconfig.json"], {
        cwd: path.join(ROOT, "packages", p),
        stdio: process.env.DEBUG ? "inherit" : "pipe",
    });
}

// 2. extract fenced ts/tsx samples
const samples = [];
let skipped = 0;
for (const doc of DOCS) {
    const lines = readFileSync(path.join(ROOT, doc), "utf8").split("\n");
    for (let i = 0; i < lines.length; i++) {
        const m = lines[i].match(/^```(ts|tsx)$/);
        if (!m) continue;
        const start = i + 1;
        let end = start;
        while (end < lines.length && lines[end] !== "```") end++;
        // nearest non-empty line above the fence may opt the sample out
        let above = i - 1;
        while (above >= 0 && lines[above].trim() === "") above--;
        if (above >= 0 && /<!--\s*docs-check:skip\b/.test(lines[above])) {
            console.log(`  skip ${doc}:${start + 1} (${lines[above].trim()})`);
            skipped++;
        } else {
            samples.push({ doc, line: start + 1, ext: m[1], code: lines.slice(start, end).join("\n") });
        }
        i = end;
    }
}
if (!samples.length) {
    console.error("DOCS CHECK FAILED: no ts/tsx samples extracted — extraction broken?");
    process.exit(1);
}

// 3. scratch consumer inside the repo: @mini-msal/* resolve through the
// workspace symlinks, whose exports "types" condition points at dist
rmSync(SCRATCH, { recursive: true, force: true });
mkdirSync(SCRATCH, { recursive: true });
for (const s of samples) {
    const name = `${s.doc.replace(/[/.]/g, "-").replace(/-md$/, "")}-L${s.line}.${s.ext}`;
    // export {} forces module scope so samples can't collide across files
    writeFileSync(path.join(SCRATCH, name), `// ${s.doc}:${s.line}\n${s.code}\nexport {};\n`);
}
writeFileSync(
    path.join(SCRATCH, "tsconfig.json"),
    JSON.stringify(
        {
            compilerOptions: {
                target: "ES2022",
                module: "ESNext",
                moduleResolution: "bundler",
                lib: ["ES2022", "DOM", "DOM.Iterable"],
                jsx: "react-jsx",
                strict: true,
                skipLibCheck: true,
                noEmit: true,
            },
            include: ["*.ts", "*.tsx"],
        },
        null,
        2
    )
);

// 4. strict type-check; keep the scratch dir on failure for debugging
console.log(`tsc --noEmit (${samples.length} samples, ${skipped} skipped)…`);
try {
    execFileSync(TSC, ["-p", "tsconfig.json"], { cwd: SCRATCH, stdio: "pipe" });
} catch (e) {
    console.error("DOCS CHECK FAILED (samples kept in test/docs/.samples):");
    console.error(String(e.stdout ?? e.message));
    process.exit(1);
}
rmSync(SCRATCH, { recursive: true, force: true });
console.log(`docs check OK — ${samples.length} samples type-check against dist types`);
