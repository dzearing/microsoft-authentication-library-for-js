/**
 * Area 4 — errors & guards: full error taxonomy (class + errorCode) for each
 * failure mode, plus environment guards (iframe/popup/interaction lock).
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    fire,
    pending,
    loginViaPopup,
    idp,
} from "../lib.mjs";

export const area = "errors";

export const scenarios = [
    {
        id: "errors.cancelled-login-redirect",
        note: "IdP returns error=access_denied on redirect login (user cancelled)",
        async run(ctx) {
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            await create(ctx, config);
            await idp.inject({
                endpoint: "authorize",
                error: "access_denied",
                errorDescription: "AADSTS65004: User declined to consent.",
                count: "1",
            });
            await ctx.page.evaluate(() => {
                globalThis.__msal
                    .loginRedirect({ scopes: ["User.Read"] })
                    .catch(
                        (e) =>
                            (globalThis.__redirectErr =
                                globalThis.__serializeError(e))
                    );
            });
            await ctx.page.waitForURL((u) => /[#&]error=/.test(u.hash), {
                timeout: 15000,
            });
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await create(ctx, config);
            const result = await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const cap = await capture(ctx);
            return {
                result,
                events: cap.events,
                hashAfter: new URL(ctx.page.url()).hash,
            };
        },
    },
    {
        id: "errors.popup-blocked",
        note: "window.open returns null: popup-blocked error taxonomy",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await ctx.page.evaluate(() => {
                window.open = () => null;
            });
            const result = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const cap = await capture(ctx);
            return { result, events: cap.events };
        },
    },
    {
        id: "errors.popup-closed-by-user",
        note: "user closes the login popup mid-flow (real v5 has NO close detection: bridge timeout only)",
        async run(ctx) {
            await gotoHarness(ctx);
            // short bridge timeout so the capture completes quickly — real
            // v5.16 only notices a closed popup when the bridge times out
            await create(
                ctx,
                stdConfig(ctx, {
                    system: {
                        popupBridgeTimeout: 3000,
                        windowHashTimeout: 3000,
                    },
                })
            );
            await idp.inject({ endpoint: "authorize", delay: "6000" });
            // capture the popup handle in-page so the "user" can close it at
            // a fixed 500ms after the call on BOTH stacks (playwright's page
            // event fires at different times for the two implementations)
            await ctx.page.evaluate(() => {
                const orig = window.open.bind(window);
                window.open = (...args) =>
                    (globalThis.__lastPopup = orig(...args));
            });
            await fire(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await ctx.page.waitForTimeout(500);
            const closed = await ctx.page.evaluate(() => {
                if (globalThis.__lastPopup) {
                    globalThis.__lastPopup.close();
                    return true;
                }
                return false;
            });
            const result = await pending(ctx);
            return { result, popupWasOpenWhenClosed: closed };
        },
    },
    {
        id: "errors.iframe-timeout",
        note: "silent iframe never returns (IdP hangs): timeout error taxonomy",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(
                ctx,
                stdConfig(ctx, {
                    system: {
                        iframeBridgeTimeout: 1500,
                        iframeHashTimeout: 1500,
                    },
                })
            );
            await idp.inject({ endpoint: "authorize", delay: "8000" });
            const result = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.ssoSilent({
                        scopes: ["User.Read"],
                        loginHint: "ada@contoso.com",
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            return { result };
        },
        timeout: 30000,
    },
    {
        id: "errors.interaction-required-variants",
        note: "interaction_required / consent_required / login_required error classification",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const out = {};
            for (const code of [
                "interaction_required",
                "consent_required",
                "login_required",
            ]) {
                await idp.inject({
                    endpoint: "authorize",
                    error: code,
                    errorDescription: `AADSTS50076: ${code} description`,
                    count: "1",
                });
                out[code] = await tryResult(
                    ctx,
                    ([popupUrl]) =>
                        globalThis.__msal.ssoSilent({
                            scopes: ["User.Read"],
                            loginHint: "ada@contoso.com",
                            redirectUri: popupUrl,
                        }),
                    [ctx.popupUrl]
                );
            }
            return out;
        },
    },
    {
        id: "errors.uninitialized-client",
        note: "API calls before initialize(): which throw, with what code",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { noInit: true });
            const calls = {};
            calls.loginPopup = await tryEval(ctx, async () => {
                await globalThis.__msal.loginPopup({ scopes: ["User.Read"] });
                return "resolved";
            });
            calls.acquireTokenSilent = await tryEval(ctx, async () => {
                await globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                });
                return "resolved";
            });
            calls.handleRedirectPromise = await tryEval(ctx, async () => {
                await globalThis.__msal.handleRedirectPromise();
                return "resolved";
            });
            calls.getAllAccounts = await tryEval(ctx, () => {
                const a = globalThis.__msal.getAllAccounts();
                return `resolved:${a.length}`;
            });
            calls.logoutRedirect = await tryEval(ctx, async () => {
                await globalThis.__msal.logoutRedirect();
                return "resolved";
            });
            return calls;
        },
    },
    {
        id: "errors.redirect-in-iframe",
        note: "loginRedirect from inside an iframe is blocked",
        async run(ctx) {
            await gotoHarness(ctx);
            await ctx.page.evaluate((src) => {
                const f = document.createElement("iframe");
                f.src = src + "?frame=1";
                document.body.append(f);
            }, ctx.harnessUrl);
            const frame = await ctx.page.waitForFunction(() => {
                return window.frames.length === 1;
            });
            const child = ctx.page
                .frames()
                .find((f) => f.url().includes("frame=1"));
            await child.waitForFunction(() => !!globalThis.__create);
            await child.evaluate(
                (c) => globalThis.__create(c),
                stdConfig(ctx)
            );
            const result = await child.evaluate(async () => {
                try {
                    await globalThis.__msal.loginRedirect({
                        scopes: ["User.Read"],
                    });
                    return { ok: "navigated" };
                } catch (e) {
                    return { err: globalThis.__serializeError(e) };
                }
            });
            return { result };
        },
    },
    {
        id: "errors.nested-popup-guard",
        note: "auth APIs inside an msal-named popup window are blocked",
        async run(ctx) {
            await gotoHarness(ctx);
            await ctx.page.evaluate(() => {
                window.name = "msal.11111111-2222-3333-4444-555555555555";
            });
            await create(ctx, stdConfig(ctx));
            const popupCall = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const silentCall = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                })
            );
            return { popupCall, silentCall };
        },
    },
    {
        id: "errors.interaction-in-progress",
        note: "second interactive call while a popup flow is pending",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await idp.inject({ endpoint: "authorize", delay: "4000" });
            await fire(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await ctx.page.waitForTimeout(500);
            const second = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const first = await pending(ctx);
            return { second, firstEventuallyResolved: first };
        },
        timeout: 30000,
    },
    {
        id: "errors.silent-unknown-account",
        note: "acquireTokenSilent for an account that is not in the cache",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const result = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: {
                        homeAccountId: "unknown-oid.unknown-tenant",
                        environment: "login.windows.net",
                        tenantId: "unknown-tenant",
                        username: "ghost@contoso.com",
                        localAccountId: "unknown-oid",
                        idTokenClaims: {},
                    },
                })
            );
            const noAccount = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                })
            );
            return { unknownAccount: result, noAccountAtAll: noAccount };
        },
    },
];
