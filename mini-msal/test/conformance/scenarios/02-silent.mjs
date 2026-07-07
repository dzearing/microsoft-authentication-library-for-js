/**
 * Area 2 — silent acquisition: cache-hit semantics, expiry-window refresh,
 * refresh-token grant, iframe fallback, CacheLookupPolicy values, concurrent
 * dedupe, forceRefresh.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    loginViaPopup,
    patchAccessTokens,
    removeCreds,
    digestIdpLog,
    idp,
} from "../lib.mjs";

export const area = "silent";

/** Sign in via popup, then reset the IdP request log (keeps the session warm). */
async function seedLogin(ctx, config) {
    await gotoHarness(ctx);
    const login = await loginViaPopup(ctx, config);
    if (login.err) throw new Error(`seed login failed: ${JSON.stringify(login.err)}`);
    // clear the request log but preserve the warm IdP session for iframe flows
    await idp.reset();
    await idp.session({ active: "1", user: "0" });
    return login;
}

/** Recreate the client (fresh instance re-reads storage) and run a silent call. */
async function silent(ctx, config, request = {}) {
    await create(ctx, config);
    return tryResult(
        ctx,
        async (req) => {
            const account =
                globalThis.__msal.getAllAccounts()[0] ?? undefined;
            return globalThis.__msal.acquireTokenSilent({
                account,
                ...req,
            });
        },
        { scopes: ["User.Read"], redirectUri: ctx.popupUrl, ...request }
    );
}

const tokenGrants = (reqs) =>
    reqs.map((r) =>
        r.endpoint === "token"
            ? `token:${r.body?.grant_type}`
            : `${r.endpoint}${r.query?.prompt ? `:prompt=${r.query.prompt}` : ""}`
    );

export const scenarios = [
    {
        id: "silent.cache-hit-fresh",
        note: "valid cached AT: served from cache with zero network traffic",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const result = await silent(ctx, config);
            const reqs = await idp.requests();
            return {
                result,
                networkCalls: tokenGrants(reqs),
                // expiresOn must reflect the cached entity, not "now"
                expiresRoughlyOneHour: await tryEval(ctx, async () => {
                    const account = globalThis.__msal.getAllAccounts()[0];
                    const r = await globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account,
                    });
                    const mins =
                        (+new Date(r.expiresOn) - Date.now()) / 60000;
                    return mins > 50 && mins < 70;
                }),
            };
        },
    },
    {
        id: "silent.expiry-window-refresh",
        note: "AT expiring inside the renewal offset window is refreshed, not served",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const soon = String(Math.floor(Date.now() / 1000) + 120); // < 300s offset
            await patchAccessTokens(ctx, {
                expiresOn: soon,
                extendedExpiresOn: soon,
            });
            const result = await silent(ctx, config);
            const reqs = await idp.requests();
            return { result, networkCalls: tokenGrants(reqs) };
        },
    },
    {
        id: "silent.refresh-token-grant",
        note: "expired AT + valid RT: refresh_token grant at the token endpoint",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            const result = await silent(ctx, config);
            const reqs = await idp.requests();
            return {
                result,
                networkCalls: tokenGrants(reqs),
                tokenRequest: digestIdpLog(
                    reqs.filter((r) => r.endpoint === "token")
                ),
            };
        },
    },
    {
        id: "silent.iframe-fallback-no-rt",
        note: "expired AT, no RT, warm IdP session: hidden-iframe prompt=none fallback",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await removeCreds(ctx, "refreshToken");
            const result = await silent(ctx, config);
            const reqs = await idp.requests();
            return {
                result,
                networkCalls: tokenGrants(reqs),
                authorizeParams: digestIdpLog(
                    reqs.filter((r) => r.endpoint === "authorize")
                ),
            };
        },
    },
    {
        id: "silent.policy-access-token-valid",
        note: "CacheLookupPolicy.AccessToken with a valid AT: cache hit",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const result = await silent(ctx, config, { cacheLookupPolicy: 1 });
            return { result, networkCalls: tokenGrants(await idp.requests()) };
        },
    },
    {
        id: "silent.policy-access-token-expired",
        note: "CacheLookupPolicy.AccessToken with an expired AT: error, no fallback",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            const result = await silent(ctx, config, { cacheLookupPolicy: 1 });
            return { result, networkCalls: tokenGrants(await idp.requests()) };
        },
    },
    {
        id: "silent.policy-at-and-rt",
        note: "CacheLookupPolicy.AccessTokenAndRefreshToken: expired AT -> RT grant",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            const result = await silent(ctx, config, { cacheLookupPolicy: 2 });
            return { result, networkCalls: tokenGrants(await idp.requests()) };
        },
    },
    {
        id: "silent.policy-refresh-token",
        note: "CacheLookupPolicy.RefreshToken: valid AT is IGNORED, RT grant used",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const result = await silent(ctx, config, { cacheLookupPolicy: 3 });
            return { result, networkCalls: tokenGrants(await idp.requests()) };
        },
    },
    {
        id: "silent.policy-rt-and-network",
        note: "CacheLookupPolicy.RefreshTokenAndNetwork: RT grant fails -> iframe fallback",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            await idp.inject({
                endpoint: "token",
                error: "invalid_grant",
                suberror: "bad_token",
                count: "1",
            });
            const result = await silent(ctx, config, { cacheLookupPolicy: 4 });
            const reqs = await idp.requests();
            return { result, networkCalls: tokenGrants(reqs) };
        },
    },
    {
        id: "silent.policy-skip",
        note: "CacheLookupPolicy.Skip: valid AT ignored entirely, straight to iframe",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const result = await silent(ctx, config, { cacheLookupPolicy: 5 });
            const reqs = await idp.requests();
            return { result, networkCalls: tokenGrants(reqs) };
        },
    },
    {
        id: "silent.concurrent-dedupe",
        note: "two parallel identical acquireTokenSilent calls with an expired AT: one network refresh or two?",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await create(ctx, config);
            const results = await tryEval(ctx, async () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                const req = { scopes: ["User.Read"], account };
                const [a, b] = await Promise.all([
                    globalThis.__msal.acquireTokenSilent(req),
                    globalThis.__msal.acquireTokenSilent(req),
                ]);
                return {
                    sameToken: a.accessToken === b.accessToken,
                    aFromCache: a.fromCache,
                    bFromCache: b.fromCache,
                };
            });
            const reqs = await idp.requests();
            return {
                results,
                tokenEndpointHits: reqs.filter((r) => r.endpoint === "token")
                    .length,
                networkCalls: tokenGrants(reqs),
            };
        },
    },
    {
        id: "silent.force-refresh",
        note: "forceRefresh: true bypasses a valid cached AT",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const result = await silent(ctx, config, { forceRefresh: true });
            const reqs = await idp.requests();
            const cap = await capture(ctx);
            return {
                result,
                networkCalls: tokenGrants(reqs),
                events: cap.events,
            };
        },
    },
];
