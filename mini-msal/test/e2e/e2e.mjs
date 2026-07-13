/**
 * E2E test: runs the SAME full-auth-flow scenario against both the real
 * MSAL app and the mini-msal app, using the local mock IdP + headless
 * system Chrome (playwright-core, no browser download).
 *
 * Scenario per app:
 *   1. load page -> ssoSilent auto-attempt fails (login_required) and/or
 *      MsalAuthenticationTemplate auto-loginRedirect kicks in
 *   2. redirect roundtrip completes -> authenticated UI ("Hello ...")
 *   3. acquireTokenSilent -> served (cache), correct token
 *   4. acquireTokenPopup -> popup roundtrip, correct token
 *   5. add a second account (popup, prompt=select_account), switch active
 *      account both ways, per-account token isolation
 *   6. logoutPopup(account) removes only that account; final logout clears all
 *
 * Prereqs: `npm run build` and `node mock-idp.mjs` + `node serve.mjs` running,
 * or just run this script — it starts both servers itself.
 */
import { chromium } from "playwright-core";

// the mock IdP uses a self-signed cert; this process only talks to localhost
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
import { spawn } from "node:child_process";
import { setTimeout as sleep } from "node:timers/promises";
import { fileURLToPath } from "node:url";

// package root (this file lives in test/e2e/)
const ROOT = fileURLToPath(new URL("../..", import.meta.url));

const APPS = (process.env.APPS ?? "mini-mock-app,real-mock-app").split(",");
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
servers.forEach((s, i) =>
    s.on("exit", (code) => {
        if (code !== null && code !== 0) {
            console.log(`  [server ${i} died with code ${code}]`);
        }
    })
);
await sleep(500);

const browser = await chromium.launch({ channel: "chrome", headless: true });

try {
    for (const app of APPS) {
        console.log(`\n=== ${app} ===`);
        await fetch("https://localhost:4599/reset");
        const context = await browser.newContext({ ignoreHTTPSErrors: true });
        const page = await context.newPage();
        page.on("pageerror", (e) => console.log("  [pageerror]", e.message));
        page.on("framenavigated", (f) => {
            if (f === page.mainFrame()) console.log("  [nav]", f.url().slice(0, 110));
        });
        if (process.env.DEBUG) {
            page.on("requestfailed", (r) =>
                console.log("  [requestfailed]", r.failure()?.errorText, r.method(), r.url().slice(0, 120))
            );
            page.on("console", (m) =>
                console.log("  [console]", m.text().slice(0, 200))
            );
            context.on("page", (p) => {
                console.log("  [popup opened]", p.url().slice(0, 110));
                p.on("framenavigated", (f) =>
                    console.log("  [popup nav]", f.url().slice(0, 110))
                );
            });
        }

        // 1+2: load -> auto login via redirect roundtrip -> authenticated UI
        await page.goto(`http://localhost:4173/${app}/`);
        try {
            await page.getByText(/^Hello/).waitFor({ timeout: 15000 });
            check(app, "auto-login redirect roundtrip renders profile", true);
        } catch {
            const paragraphs = await page
                .locator("p")
                .allTextContents()
                .catch(() => []);
            check(app, "auto-login redirect roundtrip renders profile", false,
                JSON.stringify(paragraphs));
            await context.close();
            continue;
        }

        const banner = await page.locator("small").textContent();
        check(app, "withMsal status banner shows 1 account",
            /1 account/.test(banner ?? ""), banner ?? "");

        const protectedVisible = await page
            .getByText("Protected content")
            .isVisible();
        check(app, "MsalAuthenticationTemplate shows protected content",
            protectedVisible);

        const accounts = await page.evaluate(() =>
            globalThis.__msal.getAllAccounts()
        );
        check(app, "getAllAccounts returns signed-in account",
            accounts.length === 1 && accounts[0].username === "ada@contoso.com",
            JSON.stringify(accounts.map((a) => a.username)));

        // 3: silent from cache
        const silent = await page.evaluate(async () => {
            const [account] = globalThis.__msal.getAllAccounts();
            const r = await globalThis.__msal.acquireTokenSilent({
                scopes: ["User.Read"],
                account,
            });
            return { token: r.accessToken, fromCache: r.fromCache };
        });
        check(app, "acquireTokenSilent returns cached token",
            silent.token === "mock-access-token-0" && silent.fromCache !== false,
            JSON.stringify(silent));

        // 4: popup acquisition
        try {
            const popup = await page.evaluate(async () => {
                const r = await globalThis.__msal.acquireTokenPopup({
                    scopes: ["User.Read"],
                    redirectUri: location.pathname + "popup.html",
                });
                return r.accessToken;
            });
            check(app, "acquireTokenPopup completes popup roundtrip",
                popup === "mock-access-token-0", popup);
        } catch (e) {
            check(app, "acquireTokenPopup completes popup roundtrip", false,
                String(e).slice(0, 200));
        }

        // 5: add a second account via popup with prompt=select_account
        try {
            const added = await page.evaluate(async () => {
                const r = await globalThis.__msal.loginPopup({
                    scopes: ["User.Read"],
                    redirectUri: location.pathname + "popup.html",
                    prompt: "select_account",
                });
                globalThis.__msal.setActiveAccount(r.account);
                return {
                    user: r.account.username,
                    token: r.accessToken,
                    count: globalThis.__msal.getAllAccounts().length,
                };
            });
            check(app, "add second account (select_account popup)",
                added.user === "grace@contoso.com" &&
                    added.token === "mock-access-token-1" &&
                    added.count === 2,
                JSON.stringify(added));
        } catch (e) {
            check(app, "add second account (select_account popup)", false,
                String(e.stack ?? e).slice(0, 600));
        }

        // 6: switcher UI lists both accounts, banner counts 2
        await page.waitForTimeout(300);
        const adaBtn = page.getByRole("button", { name: /ada@contoso\.com/ });
        const graceBtn = page.getByRole("button", { name: /grace@contoso\.com/ });
        check(app, "switcher lists both accounts",
            (await adaBtn.count()) === 1 && (await graceBtn.count()) === 1,
            `banner: ${await page.locator("small").textContent()}`);

        // 7: switch active back to ada via the UI; per-account token isolation
        try {
            // after adding grace she is active; the greeting must follow the
            // active account WITHOUT a reload...
            await page.getByText(/\(grace@contoso\.com\)/).waitFor({ timeout: 3000 });
            await adaBtn.click();
            // ...and follow it back when switching via the UI
            await page.getByText(/\(ada@contoso\.com\)/).waitFor({ timeout: 3000 });
            const sw = await page.evaluate(async () => {
                const active = globalThis.__msal.getActiveAccount();
                const adaToken = (
                    await globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: active,
                    })
                ).accessToken;
                const grace = globalThis.__msal
                    .getAllAccounts()
                    .find((a) => a.username === "grace@contoso.com");
                const graceToken = (
                    await globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: grace,
                    })
                ).accessToken;
                return { active: active?.username, adaToken, graceToken };
            });
            check(app, "switch active account live-updates UI + per-account tokens",
                sw.active === "ada@contoso.com" &&
                    sw.adaToken === "mock-access-token-0" &&
                    sw.graceToken === "mock-access-token-1",
                JSON.stringify(sw));
        } catch (e) {
            check(app, "switch active account live-updates UI + per-account tokens", false,
                String(e).slice(0, 200));
        }

        // 8: logoutPopup(grace) removes only grace
        try {
            const afterOne = await page.evaluate(async () => {
                const grace = globalThis.__msal
                    .getAllAccounts()
                    .find((a) => a.username === "grace@contoso.com");
                await globalThis.__msal.logoutPopup({ account: grace });
                return globalThis.__msal.getAllAccounts().map((a) => a.username);
            });
            check(app, "logoutPopup(account) removes only that account",
                afterOne.length === 1 && afterOne[0] === "ada@contoso.com",
                JSON.stringify(afterOne));
        } catch (e) {
            check(app, "logoutPopup(account) removes only that account", false,
                String(e).slice(0, 200));
        }

        // 9: final logout clears the rest
        try {
            const remaining = await page.evaluate(async () => {
                await globalThis.__msal.logoutPopup();
                return globalThis.__msal.getAllAccounts().length;
            });
            check(app, "logoutPopup clears remaining account", remaining === 0,
                `${remaining} account(s) left`);
        } catch (e) {
            check(app, "logoutPopup clears remaining account", false,
                String(e).slice(0, 200));
        }

        await context.close();
    }

    // cache interop: the two stacks must share sign-in state (drop-in).
    // Login on one stack, navigate the SAME TAB to the other, assert signed-in
    // with no new IdP roundtrip — both directions.
    for (const [first, second] of [
        ["real-mock-app", "mini-mock-app"],
        ["mini-mock-app", "real-mock-app"],
    ]) {
        console.log(`\n=== interop: ${first} -> ${second} ===`);
        await fetch("https://localhost:4599/reset");
        const context = await browser.newContext({ ignoreHTTPSErrors: true });
        const page = await context.newPage();
        await page.goto(`http://localhost:4173/${first}/`);
        await page.getByText(/^Hello/).waitFor({ timeout: 15000 });

        let authNavs = 0;
        page.on("framenavigated", (f) => {
            if (f === page.mainFrame() && f.url().includes("4599")) authNavs++;
        });
        await page.goto(`http://localhost:4173/${second}/`);
        try {
            await page.getByText(/\(ada@contoso\.com\)/).waitFor({ timeout: 8000 });
            const state = await page.evaluate(async () => {
                const m = globalThis.__msal;
                const accounts = m.getAllAccounts().map((a) => a.username);
                const active = m.getActiveAccount()?.username ?? null;
                const r = await m.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: m.getAllAccounts()[0],
                });
                return { accounts, active, token: r.accessToken, fromCache: r.fromCache };
            });
            check(second, `reads ${first}'s cache (no re-auth)`,
                authNavs === 0 &&
                    state.accounts.join() === "ada@contoso.com" &&
                    state.active === "ada@contoso.com" &&
                    state.token === "mock-access-token-0" &&
                    state.fromCache !== false,
                JSON.stringify({ authNavs, ...state }));
        } catch (e) {
            check(second, `reads ${first}'s cache (no re-auth)`, false,
                String(e).slice(0, 250));
        }
        await context.close();
    }

    // cancelled-login regression (mini): an #error=access_denied response with a
    // pending request must surface as an error and NOT trigger an auto-login loop
    // (real MSAL's behavior, user-verified against live AAD)
    if (APPS.includes("mini-mock-app")) {
        console.log("\n=== mini-mock-app (cancelled login) ===");
        // cold IdP session: a user who cancelled sign-in has no session, so
        // the auto silent-SSO attempt must fail rather than silently sign in
        await fetch("https://localhost:4599/reset");
        const context = await browser.newContext({ ignoreHTTPSErrors: true });
        const page = await context.newPage();
        await page.addInitScript(() => {
            if (!sessionStorage.getItem("msal.request")) {
                sessionStorage.setItem(
                    "msal.request",
                    JSON.stringify({ verifier: "v", state: "S", scopes: ["User.Read"] })
                );
                // a real mid-flow state also carries the initiating page
                // (C13: without it the library replays to the homepage)
                sessionStorage.setItem(
                    "msal.11111111-2222-3333-4444-555555555555.request.origin",
                    "http://localhost:4173/mini-mock-app/"
                );
            }
        });
        await page.goto(
            "http://localhost:4173/mini-mock-app/#error=access_denied&error_description=User+cancelled&state=S"
        );
        try {
            await page.getByText(/access_denied/).first().waitFor({ timeout: 5000 });
            await page.waitForTimeout(1500); // would have redirected by now if looping
            const stillHere = page.url().startsWith("http://localhost:4173/mini-mock-app/");
            const signInVisible = await page.getByText("Please sign in.").isVisible();
            check("mini-mock-app", "cancelled login shows error, no auto-retry loop",
                stillHere && signInVisible, page.url().slice(0, 100));
        } catch (e) {
            check("mini-mock-app", "cancelled login shows error, no auto-retry loop",
                false, String(e).slice(0, 200));
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
