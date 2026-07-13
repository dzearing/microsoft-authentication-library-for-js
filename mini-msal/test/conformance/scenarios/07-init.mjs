/**
 * Area 10 — init & misc: initialize() requirements, getConfiguration, logger
 * wiring, exported enums/constants, version, redirect-bridge page contract.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
    storageDump,
    idp,
} from "../lib.mjs";

export const area = "init";

export const scenarios = [
    {
        id: "init.double-initialize",
        note: "initialize() twice is benign; events emitted",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const second = await tryEval(ctx, async () => {
                await globalThis.__msal.initialize();
                return "resolved";
            });
            const cap = await capture(ctx);
            return { second, events: cap.events };
        },
    },
    {
        id: "init.get-configuration",
        note: "getConfiguration(): resolved defaults (timeouts, offsets, flags)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const config = await tryEval(ctx, () => {
                const c = globalThis.__msal.getConfiguration?.();
                if (!c) return null;
                return {
                    topLevelKeys: Object.keys(c).sort(),
                    auth: {
                        clientId: c.auth.clientId,
                        authority: c.auth.authority,
                        navigateToLoginRequestUrl:
                            c.auth.navigateToLoginRequestUrl,
                        cloudDiscoveryMetadata:
                            c.auth.cloudDiscoveryMetadata === "" ? "" : "set",
                    },
                    cache: {
                        cacheLocation: c.cache.cacheLocation,
                        storeAuthStateInCookie:
                            c.cache.storeAuthStateInCookie,
                    },
                    system: {
                        iframeHashTimeout: c.system.iframeHashTimeout,
                        windowHashTimeout: c.system.windowHashTimeout,
                        loadFrameTimeout: c.system.loadFrameTimeout,
                        tokenRenewalOffsetSeconds:
                            c.system.tokenRenewalOffsetSeconds,
                        allowPlatformBroker: c.system.allowPlatformBroker,
                        nativeBrokerHandshakeTimeout:
                            c.system.nativeBrokerHandshakeTimeout,
                        redirectNavigationTimeout:
                            c.system.redirectNavigationTimeout,
                    },
                };
            });
            return { config };
        },
    },
    {
        id: "init.logger-callback",
        note: "loggerCallback wiring: does the library log through the configured callback?",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { captureLogs: true });
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const cap = await capture(ctx);
            const levels = [...new Set(cap.logs.map((l) => l.level))].sort();
            return {
                logged: cap.logs.length > 0,
                logCountOver10: cap.logs.length > 10,
                levelsSeen: levels,
            };
        },
    },
    {
        id: "init.exported-surface",
        note: "module exports: enums, error codes, version — the public API contract",
        async run(ctx) {
            await gotoHarness(ctx);
            return await tryEval(ctx, () => {
                const lib = globalThis.__lib;
                const digestEnum = (o) =>
                    o
                        ? Object.fromEntries(
                              Object.entries(o).filter(
                                  ([, v]) =>
                                      typeof v === "string" ||
                                      typeof v === "number"
                              )
                          )
                        : null;
                return {
                    version: lib.version ?? null,
                    hasCreateNestablePCA:
                        typeof lib.createNestablePublicClientApplication ===
                        "function",
                    hasCreateStandardPCA:
                        typeof lib.createStandardPublicClientApplication ===
                        "function",
                    hasIsPlatformBrokerAvailable:
                        typeof lib.isPlatformBrokerAvailable === "function",
                    eventType: digestEnum(lib.EventType),
                    interactionType: digestEnum(lib.InteractionType),
                    cacheLookupPolicy: digestEnum(lib.CacheLookupPolicy),
                    browserCacheLocation: digestEnum(lib.BrowserCacheLocation),
                    protocolMode: digestEnum(lib.ProtocolMode),
                    promptValue: digestEnum(lib.PromptValue),
                    errorClassesExported: {
                        AuthError: typeof lib.AuthError === "function",
                        BrowserAuthError:
                            typeof lib.BrowserAuthError === "function",
                        InteractionRequiredAuthError:
                            typeof lib.InteractionRequiredAuthError ===
                            "function",
                        ClientAuthError:
                            typeof lib.ClientAuthError === "function",
                        ServerError: typeof lib.ServerError === "function",
                    },
                    browserAuthErrorCodesCount: lib.BrowserAuthErrorCodes
                        ? Object.keys(lib.BrowserAuthErrorCodes).length
                        : null,
                    oidcDefaultScopes: lib.OIDC_DEFAULT_SCOPES ?? null,
                };
            });
        },
    },
    {
        id: "init.exported-surface-2",
        note: "module exports (C12): Logger/LogLevel/WrapperSKU, error-code namespaces, logger instance methods",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            return await tryEval(ctx, () => {
                const lib = globalThis.__lib;
                const digestEnum = (o) =>
                    o
                        ? Object.fromEntries(
                              Object.entries(o).filter(
                                  ([, v]) =>
                                      typeof v === "string" ||
                                      typeof v === "number"
                              )
                          )
                        : null;
                // Logger class: level gate + clone carry-over (no timestamps
                // in the digest — messages are only probed for the marker)
                const seen = [];
                const logger = new lib.Logger({
                    loggerCallback: (level, message) =>
                        seen.push({
                            level,
                            hasMarker: message.includes("c12-probe"),
                        }),
                    logLevel: lib.LogLevel.Verbose,
                });
                logger.info("c12-probe info");
                logger.trace("c12-probe trace (gated out at Verbose)");
                const clone = logger.clone("test-wrapper", "9.9.9");
                clone.verbose("c12-probe clone");
                const pca = globalThis.__msal;
                const instLogger = pca.getLogger?.();
                const instClone = instLogger?.clone?.("sku", "1.0");
                let setLoggerOk, wrapperVoid;
                try {
                    setLoggerOk =
                        pca.setLogger(new lib.Logger({})) === undefined;
                } catch (e) {
                    setLoggerOk = String(e);
                }
                try {
                    wrapperVoid =
                        pca.initializeWrapperLibrary(
                            "@azure/msal-react",
                            "3.0.0"
                        ) === undefined;
                } catch (e) {
                    wrapperVoid = String(e);
                }
                let bcaeShape;
                try {
                    const e = new lib.BrowserConfigurationAuthError(
                        "storage_not_supported"
                    );
                    bcaeShape = {
                        name: e.name,
                        errorCode: e.errorCode,
                        isAuthError: e instanceof lib.AuthError,
                        msgHasAka: e.message.includes(
                            "aka.ms/msal.js.errors"
                        ),
                    };
                } catch (e) {
                    bcaeShape = String(e);
                }
                return {
                    exports: {
                        logger: typeof lib.Logger,
                        logLevel: digestEnum(lib.LogLevel),
                        wrapperSKU: digestEnum(lib.WrapperSKU),
                        authErrorCodes: digestEnum(lib.AuthErrorCodes),
                        clientAuthErrorCodes: digestEnum(
                            lib.ClientAuthErrorCodes
                        ),
                        clientConfigurationErrorCodes: digestEnum(
                            lib.ClientConfigurationErrorCodes
                        ),
                        interactionRequiredAuthErrorCodes: digestEnum(
                            lib.InteractionRequiredAuthErrorCodes
                        ),
                        browserConfigurationAuthErrorCodes: digestEnum(
                            lib.BrowserConfigurationAuthErrorCodes
                        ),
                        browserConfigurationAuthError:
                            typeof lib.BrowserConfigurationAuthError,
                    },
                    bcaeShape,
                    loggerClass: {
                        callbackCalls: seen,
                        cloneIsLogger: clone instanceof lib.Logger,
                    },
                    instanceMethods: {
                        getLoggerReturnsLogger:
                            instLogger instanceof lib.Logger,
                        cloneHasMethods:
                            typeof instClone?.info === "function" &&
                            typeof instClone?.verbose === "function",
                        setLoggerOk,
                        initializeWrapperLibraryVoid: wrapperVoid,
                    },
                };
            });
        },
    },
    {
        id: "init.storage-before-login",
        note: "what the library writes to storage on bare initialize()",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            return { storage: await storageDump(ctx) };
        },
    },
    {
        id: "init.popup-without-bridge",
        note: "popup redirect page WITHOUT the redirect-bridge: does the popup flow complete?",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(
                ctx,
                stdConfig(ctx, {
                    system: {
                        popupBridgeTimeout: 3000,
                        windowHashTimeout: 3000,
                    },
                })
            );
            // /blank.html never runs the v5 redirect-bridge (it is mini's
            // normal popup page) — captures real's hard dependency on it
            const result = await tryResult(
                ctx,
                () =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: "http://localhost:4173/blank.html",
                    })
            );
            return { result };
        },
        timeout: 30000,
    },
];
