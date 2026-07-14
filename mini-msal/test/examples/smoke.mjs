/**
 * Example smoke check (D3): every examples/* profile, built against the mock
 * IdP (the *-smoke rspack variants with ./authConfig.js swapped), must
 * complete a REAL sign-in in headless Chrome and render a silently acquired
 * token. Reuses the e2e harness pattern: starts the mock IdP + static server
 * itself. Prereq: `npm run build`.
 */
import { chromium } from "playwright-core";

// the mock IdP uses a self-signed cert; this process only talks to localhost
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
import { spawn } from "node:child_process";
import { setTimeout as sleep } from "node:timers/promises";
import { fileURLToPath } from "node:url";

// package root (this file lives in test/examples/)
const ROOT = fileURLToPath(new URL("../..", import.meta.url));

// interaction: how the example's #signin button acquires the login
// (redirect roundtrip vs popup + bridge page) — assertions are identical
const EXAMPLES = [
    { app: "example-core-redirect-smoke", interaction: "redirect" },
    { app: "example-core-popup-smoke", interaction: "popup" },
    { app: "example-compat-smoke", interaction: "popup" },
    { app: "example-react-smoke", interaction: "redirect" },
];

const results = [];
const check = (app, name, ok, detail = "") => {
    results.push({ app, name, ok, detail });
    console.log(`  ${ok ? "PASS" : "FAIL"}  ${name}${detail ? ` — ${detail}` : ""}`);
};

const stdio = process.env.DEBUG ? "inherit" : "ignore";
const servers = [
    spawn("node", ["test/infra/mock-idp.mjs"], { stdio, cwd: ROOT }),
    spawn("node", ["test/infra/serve.mjs"], { stdio, cwd: ROOT }),
];
await sleep(500);

const browser = await chromium.launch({ channel: "chrome", headless: true });

try {
    for (const { app, interaction } of EXAMPLES) {
        console.log(`\n=== ${app} (${interaction}) ===`);
        await fetch("https://localhost:4599/reset");
        const context = await browser.newContext({ ignoreHTTPSErrors: true });
        const page = await context.newPage();
        page.on("pageerror", (e) => console.log("  [pageerror]", e.message));
        if (process.env.DEBUG) {
            page.on("console", (m) => console.log("  [console]", m.text().slice(0, 200)));
        }

        try {
            await page.goto(`http://localhost:4173/${app}/`);
            await page.locator("#signin").click({ timeout: 10000 });
            await page
                .getByText("Hello ada@contoso.com")
                .waitFor({ timeout: 15000 });
            check(app, "signs in against the mock IdP", true);

            const token = page.getByText(/^token: \d+ chars$/);
            await token.waitFor({ timeout: 10000 });
            check(app, "renders a silently acquired token", true,
                await token.textContent());
        } catch (e) {
            check(app, "signs in against the mock IdP", false,
                String(e).slice(0, 300));
        }
        await context.close();
    }
} finally {
    await browser.close();
    servers.forEach((s) => s.kill());
}

const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} checks passed`);
process.exit(failed.length ? 1 : 0);
