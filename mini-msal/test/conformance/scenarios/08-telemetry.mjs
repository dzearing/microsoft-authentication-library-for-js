/**
 * Area 7 — telemetry: addPerformanceCallback event shapes, x-client-* /
 * client-request-id server telemetry headers, correlationId propagation.
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
    idp,
} from "../lib.mjs";

export const area = "telemetry";

/**
 * C11: full-shape perf-event digest. Values that cannot be deterministic
 * across runs (ids, clocks, network timing) are reduced to `<typeof>`;
 * `ext` (aggregated sub-measurement data) is compared by its key set;
 * `context` (real's internal call-tree debug blob) by string presence;
 * every other key AND value must match real exactly.
 */
const VOLATILE = new Set([
    "eventId",
    "correlationId",
    "startTimeMs",
    "durationMs",
    "networkRtt",
    "networkEffectiveType",
    "logs",
    "errorStack",
    "context",
]);
function normPerfEvent(e) {
    const out = {};
    for (const k of Object.keys(e).sort()) {
        const v = e[k];
        if (v === null) {
            out[k] = null;
        } else if (VOLATILE.has(k)) {
            out[k] = `<${typeof v}>`;
        } else if (k === "ext" && typeof v === "object") {
            const ext = {};
            for (const ek of Object.keys(v).sort()) {
                ext[ek] = typeof v[ek] === "number" ? "<number>" : v[ek];
            }
            out.ext = ext;
        } else {
            out[k] = v;
        }
    }
    return out;
}
const clearPerfRaw = (ctx) =>
    ctx.page.evaluate(() => {
        globalThis.__cap.perfRaw.length = 0;
    });
async function perfShapes(ctx) {
    await ctx.page.waitForTimeout(500);
    return (await capture(ctx)).perfRaw.map(normPerfEvent);
}

export const scenarios = [
    {
        id: "telemetry.perf-events-popup-login",
        note: "performance events emitted for a popup login (names, success, durations)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { perfClient: true });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            // performance events flush asynchronously
            await ctx.page.waitForTimeout(500);
            const cap = await capture(ctx);
            return { perf: cap.perf };
        },
    },
    {
        id: "telemetry.perf-events-silent",
        note: "performance events for acquireTokenSilent: cache hit vs network refresh",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await ctx.page.evaluate(() => {
                globalThis.__cap.perf.length = 0;
            });
            await tryResult(ctx, () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            });
            await ctx.page.waitForTimeout(500);
            const cacheHitPerf = (await capture(ctx)).perf;
            await ctx.page.evaluate(() => {
                globalThis.__cap.perf.length = 0;
            });
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await tryResult(ctx, () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            });
            await ctx.page.waitForTimeout(500);
            const refreshPerf = (await capture(ctx)).perf;
            return { cacheHitPerf, refreshPerf };
        },
    },
    {
        id: "telemetry.perf-event-shape",
        note: "full field shape of emitted perf events: initialize + popup login (C11)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { perfClient: true });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            return { shapes: await perfShapes(ctx) };
        },
    },
    {
        id: "telemetry.perf-event-shape-silent",
        note: "full perf-event shape: silent cache hit, network refresh, ssoSilent, failed refresh (C11)",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const silentCall = () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            };
            await clearPerfRaw(ctx);
            await tryResult(ctx, silentCall);
            const cacheHit = await perfShapes(ctx);
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await clearPerfRaw(ctx);
            await tryResult(ctx, silentCall);
            const refresh = await perfShapes(ctx);
            await idp.session({ active: "1", user: "0" });
            await clearPerfRaw(ctx);
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.ssoSilent({
                        scopes: ["User.Read"],
                        loginHint: "ada@contoso.com",
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const sso = await perfShapes(ctx);
            // failure shape: RT redemption fails (invalid_grant), iframe
            // fallback fails fast with login_required (no IdP session)
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await idp.session({ active: "0" });
            await idp.inject({
                endpoint: "token",
                error: "invalid_grant",
                count: "1",
            });
            await clearPerfRaw(ctx);
            await tryResult(
                ctx,
                (popupUrl) => {
                    const account = globalThis.__msal.getAllAccounts()[0];
                    return globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account,
                        redirectUri: popupUrl,
                    });
                },
                ctx.popupUrl
            );
            const failure = await perfShapes(ctx);
            return { cacheHit, refresh, sso, failure };
        },
    },
    {
        id: "telemetry.token-request-headers",
        note: "server telemetry params on token POSTs (v5 sends them as BODY params, not headers)",
        async run(ctx) {
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx));
            const token = (await idp.requests()).find(
                (r) => r.endpoint === "token"
            );
            const b = token?.body ?? {};
            return {
                sku: b["x-client-SKU"] ?? null,
                ver: b["x-client-VER"] ?? null,
                currentTelemetry: b["x-client-current-telemetry"] ?? null,
                lastTelemetry: b["x-client-last-telemetry"] ?? null,
                libCapability: b["x-ms-lib-capability"] ?? null,
                // custom x-client HTTP headers are NOT sent by v5
                anyTelemetryHttpHeaders: Object.keys(
                    token?.headers ?? {}
                ).filter((k) => k.startsWith("x-client")),
            };
        },
    },
    {
        id: "telemetry.correlation-id-propagation",
        note: "request.correlationId flows into headers, result, and perf events",
        async run(ctx) {
            const CID = "11223344-5566-7788-99aa-bbccddeeff00";
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { perfClient: true });
            const result = await tryEval(
                ctx,
                async ([cid, popupUrl]) => {
                    const r = await globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        correlationId: cid,
                    });
                    return { resultCorrelationId: r.correlationId ?? null };
                },
                [CID, ctx.popupUrl]
            );
            await ctx.page.waitForTimeout(500);
            const cap = await capture(ctx);
            const reqs = await idp.requests();
            const token = reqs.find((r) => r.endpoint === "token");
            const authorize = reqs.find((r) => r.endpoint === "authorize");
            return {
                resultMatches: result.ok?.resultCorrelationId === CID,
                tokenBodyClientRequestId:
                    token?.body?.["client-request-id"] === CID
                        ? "matches"
                        : (token?.body?.["client-request-id"] ?? null),
                authorizeParamMatches:
                    authorize?.query?.["client-request-id"] === CID,
                perfEventWithCid: cap.perf.some(
                    (p) => p.correlationId === CID
                ),
            };
        },
    },
    {
        id: "telemetry.last-telemetry-after-failure",
        note: "x-client-last-telemetry carries the previous failure on the next token request",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await loginViaPopup(ctx, config);
            await idp.reset();
            await idp.session({ active: "1" });
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            // one failing refresh...
            await idp.inject({
                endpoint: "token",
                error: "invalid_grant",
                count: "1",
            });
            await create(ctx, config);
            const failed = await tryResult(ctx, (popupUrl) => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                    redirectUri: popupUrl,
                });
            }, ctx.popupUrl);
            // ...then whatever token call happens next should report it
            const tokens = (await idp.requests()).filter(
                (r) => r.endpoint === "token"
            );
            const last = tokens[tokens.length - 1]?.body ?? {};
            return {
                failedCall: failed,
                tokenCalls: tokens.length,
                lastTelemetryParam: last["x-client-last-telemetry"] ?? null,
                lastTelemetryMentionsError:
                    (last["x-client-last-telemetry"] ?? "").includes(
                        "invalid_grant"
                    ),
            };
        },
    },
];
