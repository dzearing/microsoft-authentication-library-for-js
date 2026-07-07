/**
 * Area 5 — platform broker (WAM). A fake in-page broker answers real MSAL's
 * probe/requests so the suite records the real broker contract:
 *   - DOM path (navigator.platformAuthentication, v5 experimental)
 *   - extension path (CustomEvent-free: window.postMessage + MessageChannel)
 * Protocol notes: conformance/notes/broker-protocol.md
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    loginViaPopup,
    idp,
} from "../lib.mjs";

export const area = "broker";

const BROKER_CONFIG = {
    system: { allowPlatformBroker: true },
    experimental: { allowPlatformBrokerWithDOM: true },
};

/** Install the fake DOM broker (must run BEFORE __create). */
function installDomBroker(ctx, mode = "success", error = null) {
    return ctx.page.evaluate(
        ([m, err]) => {
            const b64u = (o) =>
                btoa(JSON.stringify(o))
                    .replace(/\+/g, "-")
                    .replace(/\//g, "_")
                    .replace(/=+$/, "");
            const now = Math.floor(Date.now() / 1000);
            const idToken = `${b64u({ alg: "none" })}.${b64u({
                aud: "11111111-2222-3333-4444-555555555555",
                iss: "https://localhost:4599/tenant/v2.0",
                oid: "oid-123",
                sub: "sub-123",
                tid: "tenant-123",
                preferred_username: "ada@contoso.com",
                name: "Ada Lovelace",
                exp: now + 3600,
                iat: now,
            })}.sig`;
            globalThis.__brokerLog = [];
            window.navigator.platformAuthentication = {
                async getSupportedContracts(brokerId) {
                    globalThis.__brokerLog.push({
                        call: "getSupportedContracts",
                        brokerId,
                    });
                    return ["get-token-and-sign-out"];
                },
                async executeGetToken(req) {
                    globalThis.__brokerLog.push({
                        call: "executeGetToken",
                        requestKeys: Object.keys(req).sort(),
                        clientId: req.clientId,
                        scope: req.scope,
                        accountId: req.accountId,
                        brokerId: req.brokerId,
                        redirectUri: req.redirectUri,
                        isSecurityTokenService: req.isSecurityTokenService,
                        hasCorrelationId: !!req.correlationId,
                        extraParameterKeys: req.extraParameters
                            ? Object.keys(req.extraParameters).sort()
                            : null,
                    });
                    if (m !== "success") {
                        return { isSuccess: false, error: err };
                    }
                    return {
                        isSuccess: true,
                        accessToken: "broker-access-token",
                        idToken,
                        clientInfo: b64u({
                            uid: "oid-123",
                            utid: "tenant-123",
                        }),
                        account: { id: req.accountId, properties: {} },
                        scopes: req.scope,
                        expiresIn: 3600,
                        state: req.state,
                        properties: {},
                    };
                },
            };
            return true;
        },
        [mode, error]
    );
}

export const scenarios = [
    {
        id: "broker.dom-probe-and-interactive",
        note: "DOM broker: probe contract, acquireTokenByCode({nativeAccountId}) brokered, then brokered silent",
        async run(ctx) {
            const config = stdConfig(ctx, BROKER_CONFIG);
            await gotoHarness(ctx);
            await installDomBroker(ctx);
            await create(ctx, config);
            // acquireTokenByCode({nativeAccountId}) is the designed entry
            // point into the broker (hybrid/native bridge enablement)
            const result = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenByCode({
                    nativeAccountId: "native-acc-1",
                    scopes: ["User.Read"],
                })
            );
            // the account now carries nativeAccountId -> silent should route
            // through the broker as well
            const silentResult = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["Mail.Read"],
                    account: globalThis.__msal.getAllAccounts()[0],
                })
            );
            const brokerLog = await ctx.page.evaluate(
                () => globalThis.__brokerLog
            );
            const accountShape = await tryEval(ctx, () => {
                const a = globalThis.__msal.getAllAccounts()[0];
                return a
                    ? {
                          username: a.username,
                          nativeAccountId: a.nativeAccountId ?? null,
                      }
                    : null;
            });
            const cap = await capture(ctx);
            return {
                result,
                silentResult,
                accountShape,
                brokerLog,
                events: cap.events.filter((e) =>
                    /acquireToken/i.test(e.eventType)
                ),
            };
        },
    },
    {
        id: "broker.dom-first-login",
        note: "cold cache + broker present: does a FIRST login use the broker? (hybrid contract)",
        async run(ctx) {
            await gotoHarness(ctx);
            await installDomBroker(ctx);
            await create(ctx, stdConfig(ctx, BROKER_CONFIG));
            const result = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.acquireTokenPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const brokerLog = await ctx.page.evaluate(
                () => globalThis.__brokerLog
            );
            return { result, brokerLog };
        },
    },
    {
        id: "broker.dom-error-mapping",
        note: "broker error statuses -> msal error classes (USER_CANCEL / ACCOUNT_UNAVAILABLE)",
        async run(ctx) {
            const out = {};
            for (const [status, code] of [
                ["USER_CANCEL", "user_cancel"],
                ["ACCOUNT_UNAVAILABLE", "account_unavailable"],
                ["NO_NETWORK", "no_network"],
            ]) {
                await idp.reset();
                await gotoHarness(ctx);
                await installDomBroker(ctx, "error", {
                    code,
                    errorCode: "1000",
                    description: `fake broker ${status}`,
                    protocolError: "",
                    status,
                    properties: {},
                });
                await create(ctx, stdConfig(ctx, BROKER_CONFIG));
                out[status] = await tryResult(ctx, () =>
                    globalThis.__msal.acquireTokenByCode({
                        nativeAccountId: "native-acc-1",
                        scopes: ["User.Read"],
                    })
                );
            }
            return out;
        },
        timeout: 90000,
    },
    {
        id: "broker.dom-disabled-fallback",
        note: "fatal broker error (DISABLED): error surfaced, broker dropped for later calls",
        async run(ctx) {
            await gotoHarness(ctx);
            await installDomBroker(ctx, "error", {
                code: "disabled",
                errorCode: "1001",
                description: "broker disabled by policy",
                protocolError: "",
                status: "DISABLED",
                properties: {},
            });
            await create(ctx, stdConfig(ctx, BROKER_CONFIG));
            const result = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenByCode({
                    nativeAccountId: "native-acc-1",
                    scopes: ["User.Read"],
                })
            );
            const brokerCallsAfterFirst = await ctx.page.evaluate(
                () => globalThis.__brokerLog.length
            );
            // broker should now be dropped: a second attempt must not call it
            const second = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenByCode({
                    nativeAccountId: "native-acc-1",
                    scopes: ["User.Read"],
                })
            );
            const brokerCallsAfterSecond = await ctx.page.evaluate(
                () => globalThis.__brokerLog.length
            );
            return {
                result,
                second,
                brokerCallsAfterFirst,
                brokerCallsAfterSecond,
            };
        },
    },
    {
        id: "broker.extension-handshake-capture",
        note: "extension path: exact handshake postMessage msal emits when no extension answers",
        async run(ctx) {
            await gotoHarness(ctx);
            // record (but do not answer) the handshake broadcast
            await ctx.page.evaluate(() => {
                globalThis.__handshakes = [];
                window.addEventListener(
                    "message",
                    (event) => {
                        const d = event.data;
                        if (d && d.channel) {
                            globalThis.__handshakes.push({
                                channel: d.channel,
                                extensionId: d.extensionId ?? null,
                                method: d.body?.method ?? null,
                                hasResponseId: !!d.responseId,
                                portsTransferred: event.ports.length,
                            });
                        }
                    },
                    true
                );
            });
            const initResult = await tryEval(ctx, async (config) => {
                const started = Date.now();
                await globalThis.__create(config);
                return { initMs: Date.now() - started < 5000 ? "<5s" : ">=5s" };
            }, stdConfig(ctx, {
                system: {
                    allowPlatformBroker: true,
                    nativeBrokerHandshakeTimeout: 700,
                },
            }));
            const handshakes = await ctx.page.evaluate(
                () => globalThis.__handshakes
            );
            // web flows still work after the failed probe
            const login = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            return {
                initResult,
                handshakes,
                webFallbackLogin: { ok: !!login.ok, err: login.err ?? null },
            };
        },
        timeout: 30000,
    },
    {
        id: "broker.extension-fake-e2e",
        note: "extension path end-to-end: fake extension answers Handshake + GetToken over MessageChannel",
        async run(ctx) {
            await gotoHarness(ctx);
            await ctx.page.evaluate(() => {
                const b64u = (o) =>
                    btoa(JSON.stringify(o))
                        .replace(/\+/g, "-")
                        .replace(/\//g, "_")
                        .replace(/=+$/, "");
                const now = Math.floor(Date.now() / 1000);
                const idToken = `${b64u({ alg: "none" })}.${b64u({
                    aud: "11111111-2222-3333-4444-555555555555",
                    iss: "https://localhost:4599/tenant/v2.0",
                    oid: "oid-123",
                    sub: "sub-123",
                    tid: "tenant-123",
                    preferred_username: "ada@contoso.com",
                    name: "Ada Lovelace",
                    exp: now + 3600,
                    iat: now,
                })}.sig`;
                globalThis.__extLog = [];
                window.addEventListener(
                    "message",
                    (event) => {
                        const d = event.data;
                        if (
                            !d ||
                            d.channel !==
                                "53ee284d-920a-4b59-9d30-a60315b26836" ||
                            d.body?.method !== "Handshake"
                        ) {
                            return;
                        }
                        // hide the bounce from msal's own no-extension check
                        event.stopImmediatePropagation();
                        const port = event.ports[0];
                        globalThis.__extLog.push({
                            method: "Handshake",
                            extensionId: d.extensionId ?? null,
                        });
                        port.onmessage = (m) => {
                            const req = m.data;
                            globalThis.__extLog.push({
                                method: req.body?.method,
                                requestKeys: req.body?.request
                                    ? Object.keys(req.body.request).sort()
                                    : null,
                                accountId: req.body?.request?.accountId,
                                scope: req.body?.request?.scope,
                                tokenType: req.body?.request?.tokenType,
                                extraParameters:
                                    req.body?.request?.extraParameters ?? null,
                            });
                            if (req.body?.method === "GetToken") {
                                port.postMessage({
                                    channel: req.channel,
                                    extensionId: req.extensionId,
                                    responseId: req.responseId,
                                    body: {
                                        method: "Response",
                                        response: {
                                            status: "Success",
                                            result: {
                                                access_token:
                                                    "ext-broker-access-token",
                                                id_token: idToken,
                                                client_info: b64u({
                                                    uid: "oid-123",
                                                    utid: "tenant-123",
                                                }),
                                                account: {
                                                    id:
                                                        req.body.request
                                                            .accountId ?? "",
                                                    properties: {},
                                                },
                                                scope: req.body.request.scope,
                                                expires_in: 3600,
                                                state: "",
                                                properties: {},
                                            },
                                        },
                                    },
                                });
                            }
                        };
                        port.postMessage({
                            channel: d.channel,
                            extensionId:
                                d.extensionId ??
                                "ppnbnpeolgkicgegkbkbjmhlideopiji",
                            responseId: d.responseId,
                            body: {
                                method: "HandshakeResponse",
                                version: "1.0.0",
                            },
                        });
                    },
                    true
                );
            });
            await create(
                ctx,
                stdConfig(ctx, { system: { allowPlatformBroker: true } })
            );
            const result = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenByCode({
                    nativeAccountId: "native-acc-1",
                    scopes: ["User.Read"],
                })
            );
            const extLog = await ctx.page.evaluate(() => globalThis.__extLog);
            return { result, extLog };
        },
        timeout: 30000,
    },
    {
        id: "broker.is-platform-broker-available",
        note: "isPlatformBrokerAvailable() export behavior with and without a DOM broker",
        async run(ctx) {
            await gotoHarness(ctx);
            const without = await tryEval(ctx, async () => {
                const f = globalThis.__lib.isPlatformBrokerAvailable;
                if (typeof f !== "function") return "not-exported";
                return await f();
            });
            await installDomBroker(ctx);
            const withBroker = await tryEval(ctx, async () => {
                const f = globalThis.__lib.isPlatformBrokerAvailable;
                if (typeof f !== "function") return "not-exported";
                return await f();
            });
            return { without, withBroker };
        },
    },
];
