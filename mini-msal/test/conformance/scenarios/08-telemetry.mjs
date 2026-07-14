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
    {
        id: "telemetry.perf-redirect-event",
        note: "handleRedirectPromise emits ONE root acquireTokenRedirect event per completed roundtrip; clean loads and repeat (memoized) calls emit nothing (C20)",
        async run(ctx) {
            const config = stdConfig(ctx);
            // clean load: no interaction in progress -> no measurement at all
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true });
            await tryEval(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            await ctx.page.waitForTimeout(500);
            const cleanLoadEvents = (await capture(ctx)).perf.filter(
                (p) => p.name === "acquireTokenRedirect"
            );
            // full redirect roundtrip with a perf client on the return page
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true });
            await ctx.page.evaluate(() => {
                globalThis.__msal
                    .loginRedirect({ scopes: ["User.Read"] })
                    .catch(() => {});
            });
            await ctx.page.waitForURL((u) => /[#&](code|error)=/.test(u.hash), {
                timeout: 15000,
            });
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await create(ctx, config, { perfClient: true });
            const result = await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            // memoized: repeat calls reuse the first promise, no extra events
            await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const shapes = await perfShapes(ctx);
            return {
                cleanLoadEvents,
                result,
                eventNames: shapes.map((s) => `${s.name}:${s.success}`),
                redirectEvents: shapes.filter(
                    (s) => s.name === "acquireTokenRedirect"
                ),
            };
        },
    },
    {
        id: "telemetry.perf-failure-correlation",
        note: "failed silent WITHOUT an app correlationId: perf event cid === error.correlationId === wire client-request-id (C20)",
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
            const past = String(Math.floor(Date.now() / 1000) - 600);
            await patchAccessTokens(ctx, {
                expiresOn: past,
                extendedExpiresOn: past,
            });
            await idp.reset();
            await idp.session({ active: "0" });
            await idp.inject({
                endpoint: "token",
                error: "invalid_grant",
                count: "1",
            });
            await clearPerfRaw(ctx);
            const err = await ctx.page.evaluate(async (popupUrl) => {
                try {
                    const account = globalThis.__msal.getAllAccounts()[0];
                    await globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account,
                        redirectUri: popupUrl,
                    });
                    return null;
                } catch (e) {
                    return {
                        ...globalThis.__serializeError(e),
                        correlationId: e?.correlationId ?? null,
                    };
                }
            }, ctx.popupUrl);
            await ctx.page.waitForTimeout(500);
            const events = (await capture(ctx)).perfRaw.filter(
                (e) => e.name === "acquireTokenSilent"
            );
            const eventCid = events[0]?.correlationId;
            const token = (await idp.requests()).find(
                (r) => r.endpoint === "token"
            );
            return {
                failedEventCount: events.length,
                errorCode: err?.errorCode ?? null,
                errorHasCid:
                    typeof err?.correlationId === "string" &&
                    err.correlationId.length > 0,
                eventCidEqualsErrorCid:
                    !!eventCid && eventCid === err?.correlationId,
                // v5 sends client-request-id on the token QUERY string
                eventCidEqualsWireCid:
                    !!eventCid &&
                    eventCid === token?.query?.["client-request-id"],
            };
        },
    },
    {
        id: "telemetry.perf-callback-dedupe",
        note: "addPerformanceCallback dedupes identical callbacks (same id, single delivery); stub client returns the constant callback-id (C20)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { perfClient: true });
            const ids = await ctx.page.evaluate(() => {
                globalThis.__dupCount = 0;
                function dupCb(events) {
                    globalThis.__dupCount += events.length;
                }
                const id1 = globalThis.__msal.addPerformanceCallback(dupCb);
                const id2 = globalThis.__msal.addPerformanceCallback(dupCb);
                return { sameId: id1 === id2 };
            });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await ctx.page.waitForTimeout(500);
            const counts = await ctx.page.evaluate(() => ({
                // events emitted AFTER dupCb registered (popup login only)
                emitted: globalThis.__cap.perf.filter(
                    (p) => p.name === "acquireTokenPopup"
                ).length,
                received: globalThis.__dupCount,
            }));
            // stub path: no telemetry.client configured
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const stubIds = await ctx.page.evaluate(() => {
                function stubCb() {}
                return [
                    globalThis.__msal.addPerformanceCallback(stubCb),
                    globalThis.__msal.addPerformanceCallback(stubCb),
                ];
            });
            return {
                ...ids,
                emittedSome: counts.emitted > 0,
                deliveredOncePerEvent: counts.received === counts.emitted,
                stubIds,
            };
        },
    },
    {
        id: "telemetry.perf-preflight-failures",
        note: "no-account silent failure emits NO perf event (abandoned measurement); uninitialized preflight failures DO emit success:false (C20)",
        async run(ctx) {
            const config = stdConfig(ctx);
            // (a) initialized client, empty cache -> no_account_error
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true });
            await clearPerfRaw(ctx);
            const noAccount = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({ scopes: ["User.Read"] })
            );
            await ctx.page.waitForTimeout(500);
            const noAccountEvents = (await capture(ctx)).perf.filter(
                (p) => p.name === "acquireTokenSilent"
            );
            // (b) uninitialized client -> preflight ends the measurement
            await gotoHarness(ctx);
            await create(ctx, config, { perfClient: true, noInit: true });
            const uninitSilent = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({ scopes: ["User.Read"] })
            );
            const uninitPopup = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.acquireTokenPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await ctx.page.waitForTimeout(500);
            const uninitEvents = (await capture(ctx)).perf.map((p) => ({
                name: p.name,
                success: p.success,
                errorCode: p.errorCode,
            }));
            return {
                noAccount,
                noAccountEvents,
                uninitSilent,
                uninitPopup,
                uninitEvents,
            };
        },
    },
    {
        id: "telemetry.perf-init-once",
        note: "initialize() is measured at most once — repeat calls return early without a perf event (C20)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { perfClient: true });
            await tryEval(ctx, async () => {
                await globalThis.__msal.initialize();
                await globalThis.__msal.initialize();
            });
            await ctx.page.waitForTimeout(500);
            const initEvents = (await capture(ctx)).perf.filter(
                (p) => p.name === "initializeClientApplication"
            );
            return { initEventCount: initEvents.length, initEvents };
        },
    },
    {
        id: "telemetry.performance-marks",
        note: "sessionStorage msal.browser.performance.enabled=1 + perf client -> performance.mark/measure timeline entries (C20, LOW)",
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
            // flag off: no msal timeline entries at all
            await ctx.page.evaluate(() => {
                performance.clearMarks();
                performance.clearMeasures();
            });
            await tryResult(ctx, silentCall);
            await ctx.page.waitForTimeout(500);
            const withoutFlag = await ctx.page.evaluate(
                () =>
                    performance
                        .getEntries()
                        .filter((e) => e.name.startsWith("msal.")).length
            );
            // flag on: measures for the root + every completed sub-measurement
            await ctx.page.evaluate(() => {
                sessionStorage.setItem(
                    "msal.browser.performance.enabled",
                    "1"
                );
                performance.clearMarks();
                performance.clearMeasures();
            });
            await tryResult(ctx, silentCall);
            await ctx.page.waitForTimeout(500);
            const entries = await ctx.page.evaluate(() => ({
                measures: performance
                    .getEntriesByType("measure")
                    .map((e) => e.name),
                marks: performance
                    .getEntriesByType("mark")
                    .map((e) => e.name),
            }));
            const stripCid = (n) =>
                n.replace(/\.[0-9a-f]{8}-[0-9a-f-]{27}$/, ".<cid>");
            return {
                withoutFlag,
                measures: [...new Set(entries.measures.map(stripCid))].sort(),
                markKinds: [
                    ...new Set(
                        entries.marks.map((n) => n.split(".")[1])
                    ),
                ].sort(),
            };
        },
    },
];
