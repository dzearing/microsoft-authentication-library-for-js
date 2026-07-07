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
