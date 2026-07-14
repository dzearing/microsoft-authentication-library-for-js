/**
 * Area 14 — programmatic token APIs (C17): clearCache local sign-out,
 * hydrateCache SSR cache seeding, the top-level loadExternalTokens export,
 * and the hybrid-SPA acquireTokenByCode({code}) redemption path.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    storageDump,
    loginViaPopup,
    digestIdpLog,
    idp,
    RESULT_DIGEST,
} from "../lib.mjs";

export const area = "token-apis";

/** IdP traffic digest EXCLUDING discovery (eager-vs-lazy discovery timing is
 * C18's concern; this area pins token/authorize/logout traffic only). */
const traffic = async () =>
    digestIdpLog(await idp.requests()).filter(
        (r) => r.endpoint !== "discovery"
    );

const ACCOUNTS_DIGEST = `() => globalThis.__msal.getAllAccounts().map((a) => ({
    username: a.username,
    homeAccountId: a.homeAccountId,
    tenantId: a.tenantId,
}))`;

const accounts = (ctx) =>
    ctx.page.evaluate((src) => (0, eval)(src)(), ACCOUNTS_DIGEST);

const eventCount = (ctx) =>
    ctx.page.evaluate(() => globalThis.__cap.events.length);

const eventsSince = async (ctx, base) =>
    (await capture(ctx)).events.slice(base);

export const scenarios = [
    {
        id: "token-apis.clear-cache",
        note: "clearCache() signs out locally: cache purged, accounts gone, NO navigation and NO end_session request",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            const login = await loginViaPopup(ctx, config);
            if (login.err) {
                throw new Error(`seed login failed: ${JSON.stringify(login.err)}`);
            }
            await idp.reset();
            const evBase = await eventCount(ctx);
            const cleared = await tryEval(ctx, async () => {
                globalThis.__msal.setActiveAccount(
                    globalThis.__msal.getAllAccounts()[0]
                );
                const ret = globalThis.__msal.clearCache();
                const isPromise = typeof ret?.then === "function";
                return { isPromise, resolved: (await ret) ?? null };
            });
            return {
                accountsBefore: 1,
                cleared,
                accountsAfter: await accounts(ctx),
                activeAfter: await tryEval(ctx, () =>
                    globalThis.__msal.getActiveAccount()
                ),
                events: await eventsSince(ctx, evBase),
                // still on the harness page: local sign-out must not navigate
                samePage: await ctx.page.evaluate(
                    () => !!globalThis.__msal && location.hash === ""
                ),
                network: await traffic(),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "token-apis.hydrate-cache",
        note: "hydrateCache(result, request) writes account + id/access token entities; silent then serves from cache with no network",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await create(ctx, config);
            // acquire a real result to seed with, then wipe all storage
            const seed = await tryEval(
                ctx,
                async (req) => {
                    globalThis.__seed = await globalThis.__msal.loginPopup(req);
                    return { username: globalThis.__seed.account.username };
                },
                { scopes: ["User.Read"], redirectUri: ctx.harnessUrl + "popup.html" }
            );
            if (seed.err) {
                throw new Error(`seed login failed: ${JSON.stringify(seed.err)}`);
            }
            await ctx.page.evaluate(() => sessionStorage.clear());
            await idp.reset();
            await create(ctx, config); // fresh client, same page (keeps __seed)
            const hydrated = await tryEval(
                ctx,
                async () =>
                    (await globalThis.__msal.hydrateCache(globalThis.__seed, {
                        scopes: ["User.Read"],
                    })) ?? null
            );
            const silent = await tryResult(ctx, async () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: globalThis.__msal.getAllAccounts()[0],
                })
            );
            return {
                hydrated,
                accounts: await accounts(ctx),
                silent,
                network: await traffic(),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "token-apis.load-external-tokens",
        note: "top-level loadExternalTokens export seeds the cache from an external token response and returns an AuthenticationResult",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await idp.reset();
            const typeofExport = await ctx.page.evaluate(
                () => typeof globalThis.__lib.loadExternalTokens
            );
            const loaded = await ctx.page.evaluate(
                async ([cfg, digestSrc]) => {
                    try {
                        const b64url = (o) =>
                            btoa(JSON.stringify(o))
                                .replace(/\+/g, "-")
                                .replace(/\//g, "_")
                                .replace(/=+$/, "");
                        const now = Math.floor(Date.now() / 1000);
                        const user = {
                            sub: "sub-123",
                            oid: "oid-123",
                            tid: "tenant-123",
                            preferred_username: "ada@contoso.com",
                            name: "Ada Lovelace",
                        };
                        const response = {
                            token_type: "Bearer",
                            scope: "openid profile User.Read",
                            expires_in: 3600,
                            access_token: "mock-access-token-ext",
                            refresh_token: "mock-rt-ext",
                            id_token: `${b64url({ alg: "none" })}.${b64url({
                                aud: cfg.auth.clientId,
                                iss: "https://localhost:4599/tenant/v2.0",
                                ...user,
                                exp: now + 3600,
                                iat: now,
                            })}.sig`,
                            client_info: b64url({
                                uid: "oid-123",
                                utid: "tenant-123",
                            }),
                        };
                        const r = await globalThis.__lib.loadExternalTokens(
                            cfg,
                            { scopes: ["User.Read"] },
                            response,
                            {}
                        );
                        return { ok: (0, eval)(digestSrc)(r) };
                    } catch (e) {
                        return { err: globalThis.__serializeError(e) };
                    }
                },
                [config, RESULT_DIGEST]
            );
            await create(ctx, config);
            const silent = await tryResult(ctx, async () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: globalThis.__msal.getAllAccounts()[0],
                })
            );
            return {
                typeofExport,
                loaded,
                accounts: await accounts(ctx),
                silent,
                network: await traffic(),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "token-apis.acquire-token-by-code",
        note: "hybrid-SPA acquireTokenByCode({code}) redeems at the token endpoint (no PKCE) and dedupes concurrent same-code calls",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await idp.reset();
            await create(ctx, config);
            const evBase = await eventCount(ctx);
            const results = await ctx.page.evaluate(async (digestSrc) => {
                const digest = (0, eval)(digestSrc);
                const call = () =>
                    globalThis.__msal.acquireTokenByCode({
                        code: "mock-code-0",
                        scopes: ["User.Read"],
                    });
                try {
                    const [a, b] = await Promise.all([call(), call()]);
                    // after settling the dedupe entry is gone: fresh POST
                    const later = await call();
                    return {
                        ok: {
                            a: digest(a),
                            concurrentDeduped: a === b,
                            later: digest(later),
                            laterIsFresh: later !== a,
                        },
                    };
                } catch (e) {
                    return { err: globalThis.__serializeError(e) };
                }
            }, RESULT_DIGEST);
            return {
                results,
                accounts: await accounts(ctx),
                events: await eventsSince(ctx, evBase),
                network: await traffic(),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "token-apis.acquire-token-by-code-errors",
        note: "code+nativeAccountId -> spa_code_and_nativeAccountId_present; neither -> auth_code_or_nativeAccountId_required; no wire traffic",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await idp.reset();
            await create(ctx, config);
            const evBase = await eventCount(ctx);
            const both = await tryEval(ctx, () =>
                globalThis.__msal.acquireTokenByCode({
                    code: "mock-code-0",
                    nativeAccountId: "native-account-1",
                    scopes: ["User.Read"],
                })
            );
            const neither = await tryEval(ctx, () =>
                globalThis.__msal.acquireTokenByCode({ scopes: ["User.Read"] })
            );
            return {
                both,
                neither,
                events: await eventsSince(ctx, evBase),
                network: await traffic(),
            };
        },
    },
];
