/**
 * Conformance runner: executes scenarios against a live harness page running
 * either real @azure/msal-browser (target=real) or mini-msal (target=mini).
 *
 *   node test/conformance/run.mjs --target=real            # capture snapshots
 *   node test/conformance/run.mjs --target=real --check    # re-run, diff vs snapshots (determinism)
 *   node test/conformance/run.mjs --target=mini            # run + diff vs real snapshots
 *   --area=core,silent   run a subset of areas
 *   --grep=<substr>      run matching scenario ids
 *
 * Snapshots: test/conformance/snapshots/real/<id>.json (normalized observations).
 * Results:   test/conformance/results/<target>.json
 */
import { chromium } from "playwright-core";
import { spawn } from "node:child_process";
import { setTimeout as sleep } from "node:timers/promises";
import { readdirSync, readFileSync, writeFileSync, mkdirSync, existsSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { normalize, diff, idp, BASE } from "./lib.mjs";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const args = Object.fromEntries(
    process.argv.slice(2).map((a) => {
        const m = a.match(/^--([^=]+)(?:=(.*))?$/);
        return m ? [m[1], m[2] ?? true] : [a, true];
    })
);
const TARGET = args.target ?? "real";
const CHECK = TARGET === "mini" || !!args.check;
const AREAS = args.area ? String(args.area).split(",") : null;
const GREP = args.grep ? String(args.grep) : null;
const SCENARIO_TIMEOUT = 60_000;

if (!["real", "mini"].includes(TARGET)) {
    console.error(`unknown target: ${TARGET}`);
    process.exit(2);
}

// ---- load scenarios ---------------------------------------------------------
const scenarioDir = path.join(__dirname, "scenarios");
const scenarios = [];
for (const f of readdirSync(scenarioDir).sort()) {
    if (!f.endsWith(".mjs")) continue;
    const mod = await import(path.join(scenarioDir, f));
    for (const s of mod.scenarios) {
        scenarios.push({ area: mod.area, ...s });
    }
}
const selected = scenarios.filter(
    (s) =>
        (!AREAS || AREAS.includes(s.area)) &&
        (!GREP || s.id.includes(GREP))
);
console.log(
    `target=${TARGET} ${CHECK ? "(diff mode)" : "(capture mode)"} — ${selected.length}/${scenarios.length} scenarios`
);

// ---- servers + browser ------------------------------------------------------
const stdio = process.env.DEBUG ? "inherit" : "ignore";
const ROOT = path.resolve(__dirname, "../..");
const servers = [
    spawn("node", ["test/infra/mock-idp.mjs"], { stdio, cwd: ROOT }),
    spawn("node", ["test/infra/serve.mjs"], { stdio, cwd: ROOT }),
];
await sleep(600);

const browser = await chromium.launch({ channel: "chrome", headless: true });
const snapDir = path.join(__dirname, "snapshots", "real");
const resultsDir = path.join(__dirname, "results");
mkdirSync(snapDir, { recursive: true });
mkdirSync(resultsDir, { recursive: true });

const results = [];

try {
    for (const s of selected) {
        const started = Date.now();
        await idp.reset();
        const context = await browser.newContext({ ignoreHTTPSErrors: true });
        const page = await context.newPage();
        const navs = [];
        page.on("framenavigated", (f) => {
            if (f !== page.mainFrame()) return;
            try {
                const u = new URL(f.url());
                navs.push({
                    page: u.origin + u.pathname,
                    queryKeys: [...u.searchParams.keys()].sort(),
                    hashKeys: [
                        ...new URLSearchParams(u.hash.slice(1)).keys(),
                    ].sort(),
                });
            } catch {
                navs.push({ page: f.url() });
            }
        });
        const pageErrors = [];
        page.on("pageerror", (e) => pageErrors.push(String(e).slice(0, 200)));

        const ctx = {
            page,
            context,
            browser,
            target: TARGET,
            harnessUrl: `${BASE}/conformance-${TARGET}/`,
            popupUrl: `${BASE}/conformance-${TARGET}/popup.html`,
            reactHarnessUrl: `${BASE}/conformance-react-${TARGET}/`,
            reactPopupUrl: `${BASE}/conformance-react-${TARGET}/popup.html`,
            navs,
            pageErrors,
            idp,
        };

        let observation;
        try {
            observation = await Promise.race([
                s.run(ctx),
                sleep(s.timeout ?? SCENARIO_TIMEOUT).then(() => {
                    throw new Error(
                        `scenario timeout after ${s.timeout ?? SCENARIO_TIMEOUT}ms`
                    );
                }),
            ]);
        } catch (e) {
            observation = {
                fatal: String(e?.stack ?? e).slice(0, 500),
            };
        }
        await context.close().catch(() => {});

        const normalized = normalize(observation);
        const snapPath = path.join(snapDir, `${s.id}.json`);

        if (TARGET === "real" && !CHECK) {
            writeFileSync(snapPath, JSON.stringify(normalized, null, 2) + "\n");
            const status = observation?.fatal ? "error" : "captured";
            results.push({ id: s.id, area: s.area, status });
            console.log(
                `  ${status === "error" ? "ERR " : "CAP "} ${s.id} (${Date.now() - started}ms)` +
                    (observation?.fatal ? ` — ${observation.fatal.slice(0, 120)}` : "")
            );
        } else {
            if (!existsSync(snapPath)) {
                results.push({ id: s.id, area: s.area, status: "no-snapshot" });
                console.log(`  ???  ${s.id} — no real snapshot`);
                continue;
            }
            const real = JSON.parse(readFileSync(snapPath, "utf8"));
            const diffs = diff(real, normalized);
            const status = observation?.fatal
                ? "error"
                : diffs.length === 0
                  ? "pass"
                  : "diff";
            results.push({
                id: s.id,
                area: s.area,
                status,
                ...(diffs.length && { diffs }),
                ...(observation?.fatal && { fatal: observation.fatal }),
                observed: normalized,
            });
            console.log(
                `  ${status === "pass" ? "PASS" : status === "diff" ? "DIFF" : "ERR "} ${s.id} (${Date.now() - started}ms)` +
                    (diffs.length ? ` — ${diffs.length} diff(s)` : "")
            );
        }
    }
} finally {
    await browser.close().catch(() => {});
    servers.forEach((sv) => sv.kill());
}

const outPath = path.join(
    resultsDir,
    `${TARGET}${TARGET === "real" && CHECK ? "-check" : ""}.json`
);
writeFileSync(outPath, JSON.stringify(results, null, 2) + "\n");

const counts = {};
for (const r of results) counts[r.status] = (counts[r.status] ?? 0) + 1;
console.log(
    `\n${TARGET}: ${results.length} scenarios — ` +
        Object.entries(counts)
            .map(([k, v]) => `${k}: ${v}`)
            .join(", ")
);
console.log(`results: ${path.relative(process.cwd(), outPath)}`);
// exit codes: real capture fails on scenario errors; real --check fails on any
// nondeterminism; mini always exits 0 (diffs ARE the deliverable — see report)
const failed =
    TARGET === "real" &&
    (counts.error || counts["no-snapshot"] || (CHECK && counts.diff));
process.exit(failed ? 1 : 0);
