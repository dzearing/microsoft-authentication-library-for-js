/**
 * Seam-hardening checks (D1): on a core-only client, every feature-owned
 * public API must throw BrowserAuthError("feature_not_configured") with a
 * message naming the missing "@mini-msal/browser/<feature>" import — never
 * undefined-is-not-a-function. Composing the feature must replace the stub.
 *
 *   node test/unit/seams.mjs        (requires a prior `npm run build`)
 *
 * Runs against dist/conformance-mini-core (harness-mini-core.ts), which
 * exposes __core + __features globals.
 */
import { chromium } from "playwright-core";
import { spawn } from "node:child_process";
import { setTimeout as sleep } from "node:timers/promises";

const BASE = "http://localhost:4173";
const stdio = process.env.DEBUG ? "inherit" : "ignore";
const server = spawn("node", ["test/infra/serve.mjs"], { stdio });
await sleep(400);

const browser = await chromium.launch({ channel: "chrome", headless: true });
let failed = 0;

try {
    const page = await browser.newPage();
    await page.goto(`${BASE}/conformance-mini-core/`);
    await page.waitForFunction(() => !!globalThis.__core);

    const results = await page.evaluate(() => {
        const { createClient } = globalThis.__core;
        const feats = globalThis.__features;
        const config = { auth: { clientId: "11111111-2222-3333-4444-555555555555" } };
        const out = [];
        const check = (id, fn) => {
            try {
                const detail = fn();
                out.push({ id, ok: true, ...(detail && { detail }) });
            } catch (e) {
                out.push({ id, ok: false, detail: String(e?.stack ?? e).slice(0, 300) });
            }
        };
        const throws = (client, api, mod) => {
            try {
                client[api]({});
            } catch (e) {
                if (e.name !== "BrowserAuthError")
                    throw new Error(`${api}: name ${e.name}`);
                if (e.errorCode !== "feature_not_configured")
                    throw new Error(`${api}: code ${e.errorCode}`);
                const want = `@mini-msal/browser/${mod}`;
                if (!e.errorMessage.includes(want) || !e.errorMessage.includes(api))
                    throw new Error(`${api}: message "${e.errorMessage}"`);
                return;
            }
            throw new Error(`${api} did not throw on a core-only client`);
        };

        // core-only: all feature-owned APIs throw the documented error
        const core = createClient(config);
        const OWNED = [
            ["loginPopup", "popup"],
            ["acquireTokenPopup", "popup"],
            ["logoutPopup", "popup"],
            ["acquireTokenByCode", "broker"],
            ["addPerformanceCallback", "telemetry"],
            ["removePerformanceCallback", "telemetry"],
        ];
        for (const [api, mod] of OWNED) {
            check(`core-only.${api}`, () => throws(core, api, mod));
        }
        // core APIs unaffected by the stubs
        check("core-only.core-apis-intact", () => {
            if (core.getAllAccounts().length !== 0) throw new Error("accounts");
            if (typeof core.loginRedirect !== "function") throw new Error("loginRedirect");
        });

        // composing a feature replaces its stubs (uninitialized client, so a
        // replaced API fails with the normal preflight error instead)
        const notStub = (client, api, arg) => {
            try {
                const r = client[api](arg);
                // async APIs reject later; swallow to keep the page clean
                if (r && typeof r.catch === "function") r.catch(() => {});
            } catch (e) {
                if (e.errorCode === "feature_not_configured")
                    throw new Error(`${api} still stubbed: ${e.errorMessage}`);
            }
        };
        const withPopup = createClient(config, [feats.popup]);
        check("popup-composed.replaces-stubs", () => {
            for (const api of ["loginPopup", "acquireTokenPopup", "logoutPopup"])
                notStub(withPopup, api, {});
            // non-composed features still guarded on a partial composition
            throws(withPopup, "acquireTokenByCode", "broker");
            throws(withPopup, "addPerformanceCallback", "telemetry");
        });
        const withBroker = createClient(config, [feats.broker]);
        check("broker-composed.replaces-stub", () => {
            notStub(withBroker, "acquireTokenByCode", { code: "x" });
            throws(withBroker, "loginPopup", "popup");
        });
        const withTelemetry = createClient(config, [feats.telemetry]);
        check("telemetry-composed.replaces-stubs", () => {
            // real's stub perf client returns "" as the callback id (C20)
            const id = withTelemetry.addPerformanceCallback(() => {});
            if (typeof id !== "string") throw new Error(`callback id ${id}`);
            if (typeof withTelemetry.removePerformanceCallback(id) !== "boolean")
                throw new Error("removePerformanceCallback");
            throws(withTelemetry, "loginPopup", "popup");
        });
        return out;
    });

    for (const r of results) {
        if (!r.ok) failed++;
        console.log(`  ${r.ok ? "PASS" : "FAIL"} ${r.id}${r.ok ? "" : ` — ${r.detail}`}`);
    }
    console.log(`\nseams: ${results.length - failed}/${results.length} checks pass`);
} finally {
    await browser.close().catch(() => {});
    server.kill();
}
process.exit(failed ? 1 : 0);
