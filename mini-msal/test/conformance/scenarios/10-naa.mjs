/**
 * Area 6 — nested app auth (NAA): createNestablePublicClientApplication with
 * a mocked bridge host. Protocol notes: conformance/notes/naa-protocol.md
 */
import { stdConfig, gotoHarness, tryEval, tryResult, idp } from "../lib.mjs";

export const area = "naa";

/** Install the fake window.nestedAppAuthBridge (before createNestablePCA). */
function installBridge(ctx, mode = "success", error = null) {
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
            globalThis.__naaLog = [];
            window.nestedAppAuthBridge = {
                _cb: null,
                addEventListener(type, cb) {
                    this._cb = cb;
                },
                postMessage(str) {
                    const req = JSON.parse(str);
                    globalThis.__naaLog.push({
                        messageType: req.messageType,
                        method: req.method,
                        clientLibrary: req.clientLibrary,
                        clientLibraryVersion: req.clientLibraryVersion,
                        hasRequestId: !!req.requestId,
                        hasSendTime: typeof req.sendTime === "number",
                        tokenParams: req.tokenParams
                            ? {
                                  keys: Object.keys(req.tokenParams).sort(),
                                  clientId: req.tokenParams.clientId,
                                  scope: req.tokenParams.scope,
                                  authenticationScheme:
                                      req.tokenParams.authenticationScheme,
                                  hasCorrelationId:
                                      !!req.tokenParams.correlationId,
                                  forceRefresh:
                                      req.tokenParams.forceRefresh ?? null,
                              }
                            : null,
                    });
                    const reply = (payload) =>
                        setTimeout(
                            () =>
                                this._cb(
                                    JSON.stringify({
                                        requestId: req.requestId,
                                        ...payload,
                                    })
                                ),
                            0
                        );
                    if (req.method === "GetInitContext") {
                        reply({
                            success: true,
                            initContext: {
                                sdkName: "fake-naa-host",
                                sdkVersion: "1.0.0",
                                accountContext: null,
                                capabilities: { queryAccount: false },
                            },
                        });
                    } else if (
                        req.method === "GetToken" ||
                        req.method === "GetTokenPopup"
                    ) {
                        if (m === "success") {
                            reply({
                                success: true,
                                token: {
                                    access_token: "naa-access-token",
                                    id_token: idToken,
                                    scope: req.tokenParams?.scope,
                                    expires_in: 3600,
                                    authority:
                                        "https://localhost:4599/tenant/",
                                },
                                account: {
                                    environment: "localhost:4599",
                                    homeAccountId: "oid-123.tenant-123",
                                    tenantId: "tenant-123",
                                    username: "ada@contoso.com",
                                    localAccountId: "oid-123",
                                    name: "Ada Lovelace",
                                },
                            });
                        } else {
                            reply({ success: false, error: err });
                        }
                    }
                },
            };
            return true;
        },
        [mode, error]
    );
}

function createNestable(ctx, config) {
    return tryEval(
        ctx,
        async (c) => {
            const f = globalThis.__lib.createNestablePublicClientApplication;
            if (typeof f !== "function") {
                throw new Error(
                    "createNestablePublicClientApplication is not exported"
                );
            }
            globalThis.__msal = await f(c);
            return "created";
        },
        config
    );
}

export const scenarios = [
    {
        id: "naa.init-handshake",
        note: "createNestablePublicClientApplication: GetInitContext envelope + NAA controller active",
        async run(ctx) {
            await gotoHarness(ctx);
            await installBridge(ctx);
            const created = await createNestable(ctx, stdConfig(ctx));
            const log = await ctx.page.evaluate(() => globalThis.__naaLog);
            // NAA controller rejects redirect APIs — the discriminator that
            // the bridge (not the standard controller) is active
            const redirectApi = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenRedirect({
                    scopes: ["User.Read"],
                })
            );
            return { created, handshake: log, redirectApi };
        },
    },
    {
        id: "naa.get-token-popup",
        note: "acquireTokenPopup through the bridge: request mapping + result mapping",
        async run(ctx) {
            await gotoHarness(ctx);
            await installBridge(ctx);
            await createNestable(ctx, stdConfig(ctx));
            const result = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenPopup({
                    scopes: ["User.Read"],
                })
            );
            const log = await ctx.page.evaluate(() =>
                globalThis.__naaLog.filter((m) => m.method !== "GetInitContext")
            );
            return { result, bridgeRequests: log };
        },
    },
    {
        id: "naa.silent-cache-then-bridge",
        note: "first acquireTokenSilent -> bridge; second -> served from browser cache",
        async run(ctx) {
            await gotoHarness(ctx);
            await installBridge(ctx);
            await createNestable(ctx, stdConfig(ctx));
            const first = await tryResult(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: {
                        homeAccountId: "oid-123.tenant-123",
                        environment: "localhost:4599",
                        tenantId: "tenant-123",
                        username: "ada@contoso.com",
                        localAccountId: "oid-123",
                        idTokenClaims: {},
                    },
                })
            );
            const callsAfterFirst = await ctx.page.evaluate(
                () =>
                    globalThis.__naaLog.filter((m) => m.method === "GetToken")
                        .length
            );
            const second = await tryResult(ctx, () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                return globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
            });
            const callsAfterSecond = await ctx.page.evaluate(
                () =>
                    globalThis.__naaLog.filter((m) => m.method === "GetToken")
                        .length
            );
            return { first, callsAfterFirst, second, callsAfterSecond };
        },
    },
    {
        id: "naa.error-mapping",
        note: "bridge error statuses -> msal error classes",
        async run(ctx) {
            const out = {};
            for (const status of [
                "USER_INTERACTION_REQUIRED",
                "USER_CANCEL",
                "PERSISTENT_ERROR",
                "NO_NETWORK",
            ]) {
                await gotoHarness(ctx);
                await installBridge(ctx, "error", {
                    status,
                    code: "fake_code",
                    description: `fake ${status}`,
                });
                await createNestable(ctx, stdConfig(ctx));
                out[status] = await tryResult(ctx, () =>
                    globalThis.__msal.acquireTokenPopup({
                        scopes: ["User.Read"],
                    })
                );
            }
            return out;
        },
    },
    {
        id: "naa.unsupported-apis",
        note: "NAA controller: which APIs throw unsupported_method",
        async run(ctx) {
            await gotoHarness(ctx);
            await installBridge(ctx);
            await createNestable(ctx, stdConfig(ctx));
            const calls = {};
            for (const api of [
                "loginRedirect",
                "logoutRedirect",
                "logoutPopup",
                "acquireTokenByCode",
            ]) {
                calls[api] = await tryEval(
                    ctx,
                    async (name) => {
                        await globalThis.__msal[name]({});
                        return "resolved";
                    },
                    api
                );
            }
            calls.addPerformanceCallback = await tryEval(ctx, () => {
                globalThis.__msal.addPerformanceCallback(() => {});
                return "resolved";
            });
            calls.handleRedirectPromise = await tryEval(ctx, async () => {
                const r = await globalThis.__msal.handleRedirectPromise();
                return r === null ? "null" : "non-null";
            });
            return calls;
        },
    },
    {
        id: "naa.no-bridge-fallback",
        note: "no window.nestedAppAuthBridge: falls back to a standard PCA (web flows work)",
        async run(ctx) {
            await gotoHarness(ctx);
            const created = await createNestable(ctx, stdConfig(ctx));
            // standard controller accepts addPerformanceCallback; NAA throws
            const perfApi = await tryEval(ctx, () => {
                const id = globalThis.__msal.addPerformanceCallback(() => {});
                return typeof id === "string" ? "works" : String(id);
            });
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
                created,
                perfApi,
                webLogin: { ok: !!login.ok, err: login.err ?? null },
            };
        },
    },
];
