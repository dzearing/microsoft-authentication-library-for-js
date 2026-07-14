/**
 * Area 16 — config knobs + logout params (C19): system.allowRedirectInIframe,
 * system.tokenRenewalOffsetSeconds, logoutHint / id_token_hint /
 * extraQueryParameters on the end_session URL, system.serverTelemetryEnabled.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    loginViaPopup,
    patchAccessTokens,
    storageDump,
    digestIdpLog,
    idp,
} from "../lib.mjs";

export const area = "config";

export const scenarios = [
    {
        id: "config.allow-redirect-in-iframe",
        note: "system.allowRedirectInIframe=true: loginRedirect navigates the iframe to the IdP and the roundtrip completes",
        async run(ctx) {
            await gotoHarness(ctx);
            // iframe src EXACTLY equals redirectUri: the post-redirect URL
            // then matches the stored login-request URL, so the response is
            // processed in place (no replay navigation racing our evaluates)
            await ctx.page.evaluate((src) => {
                const f = document.createElement("iframe");
                f.src = src;
                document.body.append(f);
            }, ctx.harnessUrl);
            await ctx.page.waitForFunction(() => window.frames.length === 1);
            const child = ctx.page
                .frames()
                .find((f) => f !== ctx.page.mainFrame());
            await child.waitForFunction(() => !!globalThis.__create);
            const config = stdConfig(ctx, {
                system: { allowRedirectInIframe: true },
            });
            await child.evaluate((c) => globalThis.__create(c), config);
            // the navigation can destroy the context before evaluate returns
            await child
                .evaluate(() => {
                    globalThis.__msal
                        .loginRedirect({ scopes: ["User.Read"] })
                        .catch(
                            (e) =>
                                (globalThis.__redirectErr =
                                    globalThis.__serializeError(e))
                        );
                })
                .catch(() => {});
            // the iframe itself navigates to the IdP and back with a code
            await child.waitForURL((u) => /[#&](code|error)=/.test(u.hash), {
                timeout: 15000,
            });
            await child.waitForFunction(() => !!globalThis.__create);
            await child.evaluate((c) => globalThis.__create(c), config);
            const completed = await child.evaluate(async () => {
                try {
                    const r =
                        await globalThis.__msal.handleRedirectPromise();
                    return {
                        ok: r && {
                            username: r.account?.username ?? null,
                            hasAccessToken: !!r.accessToken,
                            fromCache: r.fromCache ?? null,
                        },
                    };
                } catch (e) {
                    return { err: globalThis.__serializeError(e) };
                }
            });
            const reqs = await idp.requests();
            return {
                redirectErr: await child.evaluate(
                    () => globalThis.__redirectErr ?? null
                ),
                completed,
                authorizeCalls: reqs.filter((r) => r.endpoint === "authorize")
                    .length,
                tokenCalls: reqs.filter((r) => r.endpoint === "token").length,
            };
        },
    },
    {
        id: "config.token-renewal-offset",
        note: "system.tokenRenewalOffsetSeconds drives the AT expiry buffer (large offset forces refresh; 0 serves near-expiry ATs from cache)",
        async run(ctx) {
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx));
            // AT expiring in ~600s + offset 1800 -> treated as expired
            const soon600 = String(Math.floor(Date.now() / 1000) + 600);
            await patchAccessTokens(ctx, {
                expiresOn: soon600,
                extendedExpiresOn: soon600,
            });
            await idp.reset();
            await create(
                ctx,
                stdConfig(ctx, {
                    system: { tokenRenewalOffsetSeconds: 1800 },
                })
            );
            const reported = await tryEval(
                ctx,
                () =>
                    globalThis.__msal.getConfiguration().system
                        .tokenRenewalOffsetSeconds
            );
            const bigOffset = await tryResult(ctx, () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            });
            const bigOffsetGrants = (await idp.requests())
                .filter((r) => r.endpoint === "token")
                .map((r) => r.body?.grant_type ?? null);
            // AT expiring in ~120s + offset 0 -> still a cache hit
            const soon = String(Math.floor(Date.now() / 1000) + 120);
            await patchAccessTokens(ctx, {
                expiresOn: soon,
                extendedExpiresOn: soon,
            });
            await idp.reset();
            await create(
                ctx,
                stdConfig(ctx, { system: { tokenRenewalOffsetSeconds: 0 } })
            );
            const zeroOffset = await tryResult(ctx, () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            });
            const zeroOffsetGrants = (await idp.requests())
                .filter((r) => r.endpoint === "token")
                .map((r) => r.body?.grant_type ?? null);
            return {
                reported,
                bigOffset: {
                    fromCache: bigOffset.ok?.fromCache ?? bigOffset,
                    grants: bigOffsetGrants,
                },
                zeroOffset: {
                    fromCache: zeroOffset.ok?.fromCache ?? zeroOffset,
                    grants: zeroOffsetGrants,
                },
            };
        },
    },
    {
        id: "config.logout-hint-params",
        note: "end_session URL params: explicit logoutHint, account-derived logout_hint (login_hint claim), idTokenHint + extraQueryParameters",
        async run(ctx) {
            // make the mock IdP put a login_hint claim in its id_tokens
            await idp.config({ "claims.login_hint": "ada.hint@contoso.com" });
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            await loginViaPopup(ctx, config);
            const doLogout = (req) =>
                tryEval(
                    ctx,
                    async ([r, popupUrl]) => {
                        const m = globalThis.__msal;
                        if (r.useAccount) {
                            r.account = m.getAllAccounts()[0];
                            delete r.useAccount;
                        }
                        await m.logoutPopup({
                            postLogoutRedirectUri: popupUrl,
                            ...r,
                        });
                        return "done";
                    },
                    [req, ctx.popupUrl]
                );
            const relogin = async () => {
                await idp.session({ active: "1" });
                await tryResult(
                    ctx,
                    (popupUrl) =>
                        globalThis.__msal.loginPopup({
                            scopes: ["User.Read"],
                            redirectUri: popupUrl,
                        }),
                    ctx.popupUrl
                );
            };
            const explicit = await doLogout({
                logoutHint: "explicit-hint@contoso.com",
            });
            await relogin();
            const derived = await doLogout({ useAccount: true });
            await relogin();
            const idTokenHint = await doLogout({
                useAccount: true,
                idTokenHint: "the-id-token-hint",
                extraQueryParameters: { customParam: "custom-value" },
            });
            const reqs = await idp.requests();
            return {
                calls: { explicit, derived, idTokenHint },
                logoutRequests: digestIdpLog(
                    reqs.filter((r) => r.endpoint === "logout")
                ),
            };
        },
    },
    {
        id: "config.server-telemetry-enabled",
        note: "system.serverTelemetryEnabled=true: populated x-client-current/last-telemetry body params + server-telemetry cache entry on failure, cleared on success",
        async run(ctx) {
            const config = stdConfig(ctx, {
                system: { serverTelemetryEnabled: true },
            });
            await gotoHarness(ctx);
            await loginViaPopup(ctx, config);
            const loginToken = (await idp.requests()).find(
                (r) => r.endpoint === "token"
            );
            const telemetryOf = (body) => ({
                current: body?.["x-client-current-telemetry"] ?? null,
                last: body?.["x-client-last-telemetry"] ?? null,
            });
            const serverTelemetryEntry = async () => {
                const dump = await storageDump(ctx);
                const all = { ...dump.localStorage, ...dump.sessionStorage };
                const key = Object.keys(all).find((k) =>
                    k.startsWith("server-telemetry-")
                );
                return key ? all[key] : null;
            };
            // fail a silent call completely: RT redemption rejected AND no
            // IdP session for the iframe fallback
            await idp.reset();
            await idp.session({ active: "0" });
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await idp.inject({
                endpoint: "token",
                error: "invalid_grant",
                count: "1",
            });
            await create(ctx, config);
            const silentCall = (popupUrl) => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                    redirectUri: popupUrl,
                });
            };
            const failed = await tryResult(ctx, silentCall, ctx.popupUrl);
            const entryAfterFailure = await serverTelemetryEntry();
            // retry with the IdP healthy: last-telemetry reports the failures
            await idp.session({ active: "1" });
            const retried = await tryResult(ctx, silentCall, ctx.popupUrl);
            const entryAfterSuccess = await serverTelemetryEntry();
            const tokens = (await idp.requests()).filter(
                (r) => r.endpoint === "token"
            );
            return {
                loginTelemetry: telemetryOf(loginToken?.body),
                failed: failed.err?.errorCode ?? failed,
                entryAfterFailure,
                retried: retried.ok?.fromCache ?? retried,
                retryTelemetry: tokens.map((r) => telemetryOf(r.body)),
                entryAfterSuccess,
            };
        },
    },
];
