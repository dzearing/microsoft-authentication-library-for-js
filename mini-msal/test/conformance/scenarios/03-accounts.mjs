/**
 * Area 3 — accounts & cache: multi-account, getAccount filters, active
 * account persistence, logout variants, cache entity shapes, localStorage
 * incl. cross-tab events.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    capture,
    storageDump,
    loginViaPopup,
    digestIdpLog,
    idp,
} from "../lib.mjs";

export const area = "accounts";

const ACCOUNT_DIGEST = `(a) => a && {
    username: a.username,
    homeAccountId: a.homeAccountId,
    localAccountId: a.localAccountId,
    tenantId: a.tenantId,
    name: a.name ?? null,
    environment: a.environment ?? null,
    hasIdTokenClaims: !!a.idTokenClaims && Object.keys(a.idTokenClaims).length > 0,
}`;

async function loginBoth(ctx, config) {
    await gotoHarness(ctx);
    const first = await loginViaPopup(ctx, config);
    const second = await loginViaPopup(ctx, config, {
        prompt: "select_account",
    });
    return { first, second };
}

export const scenarios = [
    {
        id: "accounts.multi-account",
        note: "two popup logins: account list shape/order, active account, events",
        async run(ctx) {
            const config = stdConfig(ctx);
            const logins = await loginBoth(ctx, config);
            const accounts = await tryEval(
                ctx,
                (digestSrc) =>
                    globalThis.__msal
                        .getAllAccounts()
                        .map((0, eval)(digestSrc)),
                ACCOUNT_DIGEST
            );
            const active = await tryEval(
                ctx,
                (digestSrc) =>
                    ((0, eval)(digestSrc))(
                        globalThis.__msal.getActiveAccount()
                    ),
                ACCOUNT_DIGEST
            );
            const cap = await capture(ctx);
            return {
                logins,
                accounts,
                activeAfterSecondLogin: active,
                events: cap.events,
            };
        },
    },
    {
        id: "accounts.get-account-filters",
        note: "getAccount / getAccountBy* filter semantics incl. case-insensitivity and misses",
        async run(ctx) {
            const config = stdConfig(ctx);
            await loginBoth(ctx, config);
            return {
                filters: await tryEval(ctx, () => {
                    const m = globalThis.__msal;
                    const u = (a) => a?.username ?? null;
                    const ada = m
                        .getAllAccounts()
                        .find((a) => a.username === "ada@contoso.com");
                    return {
                        byUsername: u(
                            m.getAccount?.({ username: "ada@contoso.com" })
                        ),
                        byUsernameUpper: u(
                            m.getAccount?.({ username: "ADA@CONTOSO.COM" })
                        ),
                        byHomeAccountId: u(
                            m.getAccount?.({
                                homeAccountId: ada?.homeAccountId,
                            })
                        ),
                        byLocalAccountId: u(
                            m.getAccount?.({
                                localAccountId: ada?.localAccountId,
                            })
                        ),
                        byUnknown: u(
                            m.getAccount?.({ username: "nobody@contoso.com" })
                        ),
                        emptyFilter: u(m.getAccount?.({})),
                        legacyByUsername: u(
                            m.getAccountByUsername?.("ada@contoso.com") ??
                                null
                        ),
                        legacyByHomeId: u(
                            m.getAccountByHomeId?.(ada?.homeAccountId) ?? null
                        ),
                        legacyByLocalId: u(
                            m.getAccountByLocalId?.(ada?.localAccountId) ??
                                null
                        ),
                        allWithFilter:
                            m
                                .getAllAccounts({
                                    username: "grace@contoso.com",
                                })
                                ?.map((a) => a.username) ?? null,
                    };
                }),
            };
        },
    },
    {
        id: "accounts.active-account-persistence",
        note: "setActiveAccount persists across a page reload + fresh client instance",
        async run(ctx) {
            const config = stdConfig(ctx);
            await loginBoth(ctx, config);
            await tryEval(ctx, () => {
                const m = globalThis.__msal;
                const ada = m
                    .getAllAccounts()
                    .find((a) => a.username === "ada@contoso.com");
                m.setActiveAccount(ada);
            });
            const activeFilterEntry = await ctx.page.evaluate(() => {
                const k = Object.keys(sessionStorage).find((x) =>
                    x.includes("active-account")
                );
                return k
                    ? { key: k, value: JSON.parse(sessionStorage.getItem(k)) }
                    : null;
            });
            await ctx.page.reload();
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await create(ctx, config);
            const cap0 = await capture(ctx);
            const afterReload = await tryEval(ctx, () => ({
                active: globalThis.__msal.getActiveAccount()?.username ?? null,
                accounts: globalThis.__msal
                    .getAllAccounts()
                    .map((a) => a.username),
            }));
            const clearActive = await tryEval(ctx, () => {
                globalThis.__msal.setActiveAccount(null);
                return globalThis.__msal.getActiveAccount()?.username ?? null;
            });
            return {
                activeFilterEntry,
                afterReload,
                clearActive,
                eventsAfterReload: cap0.events,
            };
        },
    },
    {
        id: "accounts.logout-popup-per-account",
        note: "logoutPopup({account}) removes only that account's entities; events + IdP logout params",
        async run(ctx) {
            const config = stdConfig(ctx);
            await loginBoth(ctx, config);
            await idp.reset();
            const logout = await tryEval(
                ctx,
                async (popupUrl) => {
                    const m = globalThis.__msal;
                    const grace = m
                        .getAllAccounts()
                        .find((a) => a.username === "grace@contoso.com");
                    // the logout popup lands on the bridge page (v5 popup
                    // completion is bridge-based, like login popups)
                    await m.logoutPopup({
                        account: grace,
                        postLogoutRedirectUri: popupUrl,
                    });
                    return {
                        remaining: m.getAllAccounts().map((a) => a.username),
                        active: m.getActiveAccount()?.username ?? null,
                    };
                },
                ctx.popupUrl
            );
            const cap = await capture(ctx);
            const reqs = await idp.requests();
            return {
                logout,
                logoutRequest: digestIdpLog(
                    reqs.filter((r) => r.endpoint === "logout")
                ),
                events: cap.events.filter((e) =>
                    /logout|account/i.test(e.eventType)
                ),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "accounts.logout-redirect",
        note: "logoutRedirect: end_session navigation params, storage cleared after roundtrip",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await loginViaPopup(ctx, config);
            await idp.reset();
            // the logout roundtrip lands back on the SAME url, so wait for a
            // navigation event rather than a URL change
            const nav = ctx.page.waitForEvent("framenavigated", {
                timeout: 15000,
            });
            await ctx.page.evaluate(() => {
                globalThis.__msal
                    .logoutRedirect()
                    .catch(
                        (e) =>
                            (globalThis.__logoutErr =
                                globalThis.__serializeError(e))
                    );
            });
            await nav;
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await create(ctx, config);
            const after = await tryEval(ctx, () => ({
                accounts: globalThis.__msal.getAllAccounts().length,
                active: globalThis.__msal.getActiveAccount()?.username ?? null,
            }));
            const reqs = await idp.requests();
            return {
                after,
                logoutRequest: digestIdpLog(
                    reqs.filter((r) => r.endpoint === "logout")
                ),
                storage: await storageDump(ctx),
                navs: ctx.navs,
            };
        },
    },
    {
        id: "accounts.local-storage",
        note: "cacheLocation localStorage: where entities land, what stays in sessionStorage",
        async run(ctx) {
            const config = stdConfig(ctx, {
                cache: { cacheLocation: "localStorage" },
            });
            await gotoHarness(ctx);
            const login = await loginViaPopup(ctx, config);
            const silentAfter = await tryEval(ctx, async () => {
                const m = globalThis.__msal;
                const r = await m.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: m.getAllAccounts()[0],
                });
                return { fromCache: r.fromCache, token: r.accessToken };
            });
            return {
                login,
                silentAfter,
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "accounts.cross-tab-events",
        note: "localStorage: a second tab sees the account; logout in tab1 -> events/state in tab2",
        async run(ctx) {
            const config = stdConfig(ctx, {
                cache: { cacheLocation: "localStorage" },
            });
            await gotoHarness(ctx);
            await loginViaPopup(ctx, config);

            const page2 = await ctx.context.newPage();
            await page2.goto(ctx.harnessUrl);
            await page2.waitForFunction(() => !!globalThis.__create);
            await page2.evaluate((c) => globalThis.__create(c), config);
            const tab2Initial = await page2.evaluate(() =>
                globalThis.__msal.getAllAccounts().map((a) => a.username)
            );

            // sign the account out in tab 1, then observe tab 2
            await tryEval(
                ctx,
                async (popupUrl) => {
                    const m = globalThis.__msal;
                    await m.logoutPopup({
                        account: m.getAllAccounts()[0],
                        postLogoutRedirectUri: popupUrl,
                    });
                },
                ctx.popupUrl
            );
            await page2.waitForTimeout(1000);
            const tab2After = await page2.evaluate(() => ({
                accounts: globalThis.__msal
                    .getAllAccounts()
                    .map((a) => a.username),
                events: globalThis.__cap.events,
            }));
            await page2.close();
            return { tab2Initial, tab2After };
        },
    },
];
