/**
 * Area 9 — resilience: 429/Retry-After throttling, 5xx handling, network
 * failure mapping, proactive refresh (refresh_in / refreshOn).
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    loginViaPopup,
    patchAccessTokens,
    idp,
} from "../lib.mjs";

export const area = "resilience";

async function seedExpired(ctx, config) {
    await gotoHarness(ctx);
    const login = await loginViaPopup(ctx, config);
    if (login.err) throw new Error(`seed login failed: ${JSON.stringify(login.err)}`);
    await idp.reset();
    await idp.session({ active: "1", user: "0" });
    const past = String(Math.floor(Date.now() / 1000) - 600);
    await patchAccessTokens(ctx, { expiresOn: past, extendedExpiresOn: past });
    await create(ctx, config);
}

const silentCall = (ctx) =>
    tryResult(ctx, () => {
        const account = globalThis.__msal.getAllAccounts()[0];
        return globalThis.__msal.acquireTokenSilent({
            scopes: ["User.Read"],
            account,
        });
    });

export const scenarios = [
    {
        id: "resilience.throttle-429-retry-after",
        note: "token endpoint 429 + Retry-After: error surfaced, throttle cache blocks the immediate retry",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedExpired(ctx, config);
            await idp.inject({
                endpoint: "token",
                status: "429",
                retryAfter: "5",
                error: "temporarily_throttled",
                errorDescription: "AADSTS90101: too many requests",
                count: "1",
            });
            const first = await silentCall(ctx);
            const throttleKeys = await ctx.page.evaluate(() =>
                Object.keys(sessionStorage).filter((k) =>
                    /throttl/i.test(k)
                ).length
            );
            const retry = await silentCall(ctx);
            const reqs = await idp.requests();
            return {
                first,
                throttleEntriesInStorage: throttleKeys,
                immediateRetry: retry,
                tokenEndpointHits: reqs.filter((r) => r.endpoint === "token")
                    .length,
            };
        },
    },
    {
        id: "resilience.5xx-server-error",
        note: "token endpoint 503 once: is the request retried? final outcome",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedExpired(ctx, config);
            await idp.inject({
                endpoint: "token",
                status: "503",
                error: "service_unavailable",
                count: "1",
            });
            const result = await silentCall(ctx);
            const reqs = await idp.requests();
            return {
                result,
                tokenEndpointHits: reqs.filter((r) => r.endpoint === "token")
                    .length,
            };
        },
    },
    {
        id: "resilience.network-drop",
        note: "token endpoint socket drop: network failure error taxonomy",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedExpired(ctx, config);
            await idp.inject({ endpoint: "token", drop: "1", count: "1" });
            const result = await silentCall(ctx);
            const reqs = await idp.requests();
            return {
                result,
                tokenEndpointHits: reqs.filter((r) => r.endpoint === "token")
                    .length,
            };
        },
    },
    {
        id: "resilience.proactive-refresh",
        note: "refresh_in in the token response: cached result served + background refresh fired",
        async run(ctx) {
            const config = stdConfig(ctx);
            // token responses carry refresh_in=1s so the refreshOn point
            // passes almost immediately while expiry stays 1h away
            await idp.config({ refresh_in: "1" });
            await gotoHarness(ctx);
            const login = await loginViaPopup(ctx, config);
            await idp.reset();
            await idp.config({ refresh_in: "1" });
            await idp.session({ active: "1" });
            await ctx.page.waitForTimeout(1500);
            await create(ctx, config);
            const result = await silentCall(ctx);
            // give any fire-and-forget background refresh time to land
            await ctx.page.waitForTimeout(1500);
            const reqs = await idp.requests();
            return {
                login: { ok: !!login.ok },
                result,
                refreshOnStored: await ctx.page.evaluate(() => {
                    const keysKey = Object.keys(sessionStorage).find((k) =>
                        k.startsWith("msal.3.token.keys")
                    );
                    if (!keysKey) return null;
                    const keys = JSON.parse(sessionStorage.getItem(keysKey));
                    const at = JSON.parse(
                        sessionStorage.getItem(keys.accessToken[0])
                    );
                    return at.refreshOn ? "present" : "absent";
                }),
                backgroundTokenRequests: reqs.filter(
                    (r) => r.endpoint === "token"
                ).length,
            };
        },
        timeout: 30000,
    },
];
