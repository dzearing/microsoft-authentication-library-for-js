/**
 * Area 1 — core flows: loginRedirect/handleRedirectPromise, loginPopup,
 * acquireTokenPopup, ssoSilent, acquireTokenRedirect, state validation,
 * PKCE params, response_mode, hash cleanup.
 */
import { createHash } from "node:crypto";
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    storageDump,
    loginViaRedirect,
    loginViaPopup,
    digestIdpLog,
    armOpenRecorder,
    openCalls,
    idp,
} from "../lib.mjs";

export const area = "core";

/** node-side PKCE check: S256(code_verifier from /token) === code_challenge from /authorize */
function pkceValid(reqs) {
    const authorize = reqs.find((r) => r.endpoint === "authorize");
    const token = reqs.find((r) => r.endpoint === "token");
    const challenge = authorize?.query?.code_challenge;
    const verifier = token?.body?.code_verifier;
    if (!challenge || !verifier) return { challenge: !!challenge, verifier: !!verifier };
    const computed = createHash("sha256")
        .update(verifier)
        .digest("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");
    return computed === challenge;
}

export const scenarios = [
    {
        id: "core.login-redirect-roundtrip",
        note: "loginRedirect -> IdP -> handleRedirectPromise; PKCE, response_mode, hash cleanup, events",
        async run(ctx) {
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            const result = await loginViaRedirect(ctx, config);
            const again = await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const reqs = await idp.requests();
            const cap = await capture(ctx);
            const url = new URL(ctx.page.url());
            return {
                result,
                secondHandleRedirectCall: again,
                hashAfterProcessing: url.hash,
                pkceValid: pkceValid(reqs),
                idp: digestIdpLog(reqs),
                events: cap.events,
                accounts: await tryEval(ctx, () =>
                    globalThis.__msal.getAllAccounts().map((a) => a.username)
                ),
            };
        },
    },
    {
        id: "core.handle-redirect-clean-load",
        note: "handleRedirectPromise on a clean page load returns null, emits no login events",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const result = await tryEval(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const cap = await capture(ctx);
            return { result, events: cap.events };
        },
    },
    {
        id: "core.login-popup-roundtrip",
        note: "loginPopup full roundtrip: result shape, events, cache entities created",
        async run(ctx) {
            await gotoHarness(ctx);
            const result = await loginViaPopup(ctx, stdConfig(ctx));
            const reqs = await idp.requests();
            const cap = await capture(ctx);
            return {
                result,
                pkceValid: pkceValid(reqs),
                idp: digestIdpLog(reqs),
                events: cap.events,
            };
        },
    },
    {
        id: "core.acquire-token-popup",
        note: "acquireTokenPopup for an additional scope set after popup login",
        async run(ctx) {
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx));
            const result = await tryResult(
                ctx,
                async (popupUrl) =>
                    globalThis.__msal.acquireTokenPopup({
                        scopes: ["Mail.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const cap = await capture(ctx);
            return { result, events: cap.events };
        },
    },
    {
        id: "core.sso-silent-cold",
        note: "ssoSilent with no IdP session -> login_required error taxonomy",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
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
            const reqs = await idp.requests();
            const cap = await capture(ctx);
            return {
                result,
                idp: digestIdpLog(reqs),
                events: cap.events,
            };
        },
    },
    {
        id: "core.sso-silent-warm",
        note: "ssoSilent with a warm IdP session succeeds via hidden iframe (prompt=none)",
        async run(ctx) {
            await idp.session({ active: "1", user: "0" });
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
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
            const reqs = await idp.requests();
            const cap = await capture(ctx);
            return {
                result,
                pkceValid: pkceValid(reqs),
                idp: digestIdpLog(reqs),
                events: cap.events,
            };
        },
    },
    {
        id: "core.acquire-token-redirect",
        note: "acquireTokenRedirect roundtrip for new scopes with an existing session",
        async run(ctx) {
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            await loginViaRedirect(ctx, config);
            await idp.reset();
            await idp.session({ active: "1" });
            await ctx.page.evaluate(() => {
                const account = globalThis.__msal.getAllAccounts()[0];
                globalThis.__msal
                    .acquireTokenRedirect({ scopes: ["Mail.Read"], account })
                    .catch(
                        (e) =>
                            (globalThis.__redirectErr =
                                globalThis.__serializeError(e))
                    );
            });
            await ctx.page.waitForURL((u) => /[#&](code|error)=/.test(u.hash), {
                timeout: 15000,
            });
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await create(ctx, config);
            const result = await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const reqs = await idp.requests();
            return {
                result,
                idp: digestIdpLog(reqs),
                hashAfterProcessing: new URL(ctx.page.url()).hash,
            };
        },
    },
    {
        id: "core.redirect-state-tampered",
        note: "auth response with a forged state: how does handleRedirectPromise react?",
        async run(ctx) {
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            await create(ctx, config);
            await ctx.page.evaluate(() => {
                globalThis.__msal
                    .loginRedirect({ scopes: ["User.Read"] })
                    .catch(
                        (e) =>
                            (globalThis.__redirectErr =
                                globalThis.__serializeError(e))
                    );
            });
            await ctx.page.waitForURL((u) => /[#&]code=/.test(u.hash), {
                timeout: 15000,
            });
            // forge the state before the library ever sees the response
            await ctx.page.evaluate(() => {
                location.hash = location.hash.replace(
                    /state=[^&]*/,
                    "state=FORGED-STATE-VALUE"
                );
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
                hashAfter: new URL(ctx.page.url()).hash.replace(
                    /state=[^&]*/,
                    "state=<forged>"
                ),
            };
        },
    },
    {
        id: "core.popup-open-timing",
        note: "default navigatePopups: window.open call timing/url/name/features during loginPopup (popup-blocker window)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await armOpenRecorder(ctx);
            const result = await tryResult(
                ctx,
                async (popupUrl) => {
                    globalThis.__inApiCall = true;
                    const p = globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    });
                    globalThis.__inApiCall = false;
                    return p;
                },
                ctx.popupUrl
            );
            return { result, openCalls: await openCalls(ctx) };
        },
    },
    {
        id: "core.popup-open-timing-async",
        note: "system.navigatePopups=false: popup opens late, directly at the authorize URL",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(
                ctx,
                stdConfig(ctx, { system: { navigatePopups: false } })
            );
            await armOpenRecorder(ctx);
            const result = await tryResult(
                ctx,
                async (popupUrl) => {
                    globalThis.__inApiCall = true;
                    const p = globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    });
                    globalThis.__inApiCall = false;
                    return p;
                },
                ctx.popupUrl
            );
            return { result, openCalls: await openCalls(ctx) };
        },
    },
    {
        id: "core.popup-window-attributes",
        note: "request.popupWindowAttributes: popupSize/popupPosition drive the window.open features string",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await armOpenRecorder(ctx);
            const result = await tryResult(
                ctx,
                async (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        popupWindowAttributes: {
                            popupSize: { width: 300, height: 301 },
                            popupPosition: { top: 11, left: 22 },
                        },
                    }),
                ctx.popupUrl
            );
            return { result, openCalls: await openCalls(ctx) };
        },
    },
    {
        id: "core.storage-shape-after-login",
        note: "exact storage keys + entity shapes written by a popup login (v5 msal.3 schema)",
        async run(ctx) {
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx));
            return { storage: await storageDump(ctx) };
        },
    },
];
