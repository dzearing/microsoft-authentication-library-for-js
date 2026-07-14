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
    {
        id: "init.exported-surface-full",
        note: "FULL module export surface (D5): every key + typeof pinned so the surface can never silently diverge, plus behavior probes for the long-tail exports (BrowserUtils, storage classes, SignedHttpRequest, stubbed PCA, EventHandler, perf helpers)",
        async run(ctx) {
            await gotoHarness(ctx);
            return await tryEval(ctx, async () => {
                const lib = globalThis.__lib;
                // documented mini-only convenience extras (D4 decision: they
                // stay); everything else must match real key-for-key
                const EXTRAS = [
                    "NativeAuthError",
                    "NestedAppAuthError",
                    "createAuth",
                ];
                const keys = Object.keys(lib)
                    .filter((k) => !EXTRAS.includes(k))
                    .sort();
                const types = Object.fromEntries(
                    keys.map((k) => [k, typeof lib[k]])
                );

                const errShape = (fn) => {
                    try {
                        fn();
                        return null;
                    } catch (e) {
                        return { name: e.name, code: e.errorCode };
                    }
                };
                const rejShape = (p) =>
                    p.then(
                        () => null,
                        (e) => ({ name: e.name, code: e.errorCode })
                    );

                // exact values of the plain-constant exports
                const constants = {
                    ApiId: lib.ApiId,
                    AzureCloudInstance: lib.AzureCloudInstance,
                    JsonWebTokenTypes: lib.JsonWebTokenTypes,
                    ResponseMode: lib.ResponseMode,
                    DEFAULT_IFRAME_TIMEOUT_MS: lib.DEFAULT_IFRAME_TIMEOUT_MS,
                    BrowserRootPerformanceEvents:
                        lib.BrowserRootPerformanceEvents,
                };

                const bu = lib.BrowserUtils;
                const browserUtils = {
                    keys: Object.keys(bu).sort(),
                    isInIframe: bu.isInIframe(),
                    isInPopup: bu.isInPopup(),
                    currentUri:
                        bu.getCurrentUri() ===
                        location.href.split(/[?#]/)[0],
                    homepage: bu.getHomepage() === `${location.origin}/`,
                    parseNoResponse: errShape(() =>
                        bu.parseAuthResponseFromUrl()
                    ),
                    blockNonBrowser:
                        errShape(() => bu.blockNonBrowserEnvironment()) ===
                        null,
                    blockBeforeInit: errShape(() =>
                        bu.blockAPICallsBeforeInitialize(false)
                    ),
                    blockReload: errShape(() =>
                        bu.blockReloadInHiddenIframes()
                    ),
                    preflightUninit: errShape(() => bu.preflightCheck(false)),
                    preflightInit: errShape(() => bu.preflightCheck(true)),
                    redirectPreflightMemory: errShape(() =>
                        bu.redirectPreflightCheck(true, {
                            system: { allowRedirectInIframe: false },
                            cache: { cacheLocation: "memoryStorage" },
                        })
                    ),
                    guidShape:
                        /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(
                            bu.createGuid()
                        ),
                    mergedClaims: JSON.parse(
                        bu.buildMergedClaims(undefined, ["cp1", "cp2"])
                    ),
                    mergedClaimsBad: errShape(() =>
                        bu.buildMergedClaims("not-json", [])
                    ),
                    invokeResult: bu.invoke(
                        (x) => x + 1,
                        "probe",
                        new lib.Logger({}),
                        new lib.StubPerformanceClient(),
                        "cid"
                    )(41),
                    invokeAsyncIsPromise:
                        typeof bu
                            .invokeAsync(
                                async (x) => x,
                                "probe",
                                new lib.Logger({}),
                                new lib.StubPerformanceClient(),
                                "cid"
                            )(1)
                            .then === "function",
                };

                const spca = lib.stubbedPublicClientApplication;
                const stubbedPCA = {
                    keys: Object.keys(spca).sort(),
                    login: await rejShape(spca.loginPopup()),
                    silent: await rejShape(spca.acquireTokenSilent()),
                    accounts: spca.getAllAccounts(),
                    account: spca.getAccount({}),
                    addEventCallback: spca.addEventCallback(() => {}),
                    addPerformanceCallback: spca.addPerformanceCallback(
                        () => {}
                    ),
                    removePerformanceCallback:
                        spca.removePerformanceCallback("x"),
                    getLogger: errShape(() => spca.getLogger()),
                    getActiveAccount: spca.getActiveAccount(),
                    setActiveAccount: spca.setActiveAccount(null) === undefined,
                };

                const AHP = lib.AuthenticationHeaderParser;
                const headerParser = {
                    fromAuthInfo: new AHP({
                        "Authentication-Info": 'nextnonce="abc123"',
                    }).getShrNonce(),
                    fromWww: new AHP({
                        "WWW-Authenticate": 'PoP nonce="n1", realm=""',
                    }).getShrNonce(),
                    missing: errShape(() => new AHP({}).getShrNonce()),
                    invalid: errShape(() =>
                        new AHP({
                            "Authentication-Info": 'foo="bar"',
                        }).getShrNonce()
                    ),
                    challenges: new AHP({}).parseChallenges(
                        'PoP nonce="n1", realm="r"'
                    ),
                };

                const EMU = lib.EventMessageUtils;
                const msg = (eventType, interactionType) => ({
                    eventType,
                    interactionType: interactionType ?? null,
                    payload: null,
                    error: null,
                    timestamp: 0,
                });
                const emu = {
                    atStartPopup: EMU.getInteractionStatusFromEvent(
                        msg("msal:acquireTokenStart", "popup")
                    ),
                    atStartSilent: EMU.getInteractionStatusFromEvent(
                        msg("msal:acquireTokenStart", "silent")
                    ),
                    hrStart: EMU.getInteractionStatusFromEvent(
                        msg("msal:handleRedirectStart")
                    ),
                    logoutStart: EMU.getInteractionStatusFromEvent(
                        msg("msal:logoutStart")
                    ),
                    logoutEndFromLogin: EMU.getInteractionStatusFromEvent(
                        msg("msal:logoutEnd"),
                        "login"
                    ),
                    logoutEndFromLogout: EMU.getInteractionStatusFromEvent(
                        msg("msal:logoutEnd"),
                        "logout"
                    ),
                    atSuccessFromAcquire: EMU.getInteractionStatusFromEvent(
                        msg("msal:acquireTokenSuccess", "redirect"),
                        "acquireToken"
                    ),
                    atSuccessSilent: EMU.getInteractionStatusFromEvent(
                        msg("msal:acquireTokenSuccess", "silent"),
                        "acquireToken"
                    ),
                    hrEndFromHR: EMU.getInteractionStatusFromEvent(
                        msg("msal:handleRedirectEnd"),
                        "handleRedirect"
                    ),
                    hrEndFromLogin: EMU.getInteractionStatusFromEvent(
                        msg("msal:handleRedirectEnd"),
                        "login"
                    ),
                };

                const eh = new lib.EventHandler();
                const seen = [];
                const ehId = eh.addEventCallback((m) => seen.push(m), [
                    "msal:loginSuccess",
                ]);
                eh.emitEvent("msal:loginStart", "cid1", "popup");
                eh.emitEvent("msal:loginSuccess", "cid2", "popup", {
                    probe: 1,
                });
                eh.removeEventCallback(ehId);
                eh.emitEvent("msal:loginSuccess", "cid3", "popup");
                const eventHandler = {
                    idIsString: typeof ehId === "string",
                    dupIdReturnsNull:
                        eh.addEventCallback(() => {}, [], "dup") === "dup" &&
                        eh.addEventCallback(() => {}, [], "dup") === null,
                    seenCount: seen.length,
                    message: seen[0] && {
                        keys: Object.keys(seen[0]).sort(),
                        eventType: seen[0].eventType,
                        interactionType: seen[0].interactionType,
                        correlationId: seen[0].correlationId,
                        payload: seen[0].payload,
                        error: seen[0].error,
                        tsIsNumber: typeof seen[0].timestamp === "number",
                    },
                };

                const ms = new lib.MemoryStorage();
                ms.setItem("a", "1");
                await ms.setUserData("b", "2");
                const memoryStorage = {
                    get: ms.getItem("a"),
                    user: ms.getUserData("b"),
                    miss: ms.getItem("zz"),
                    keys: ms.getKeys().sort(),
                    has: ms.containsKey("a"),
                    decrypt: await ms.decryptData(),
                    clearWorks: (ms.clear(), ms.getKeys().length === 0),
                };
                const ss = new lib.SessionStorage();
                ss.setItem("d5.probe", "v");
                const sessionStorageShape = {
                    roundtrip: ss.getItem("d5.probe"),
                    contains: ss.containsKey("d5.probe"),
                    user: ss.getUserData("d5.probe"),
                    removed: (ss.removeItem("d5.probe"),
                    ss.getItem("d5.probe")),
                    decrypt: await ss.decryptData(),
                };
                const ls = new lib.LocalStorage("client-id-d5");
                ls.setItem("d5.probe", "v");
                const localStorageShape = {
                    plainRoundtrip: ls.getItem("d5.probe"),
                    contains: ls.containsKey("d5.probe"),
                    userBeforeInit: errShape(() =>
                        ls.getUserData("d5.probe")
                    ),
                    setUserBeforeInit: await rejShape(
                        ls.setUserData("k", "v", "cid", "0")
                    ),
                    removed: (ls.removeItem("d5.probe"),
                    ls.getItem("d5.probe")),
                };

                const BPM = lib.BrowserPerformanceMeasurement;
                const inst = new BPM("n", "cid");
                const bpm = {
                    measureName: BPM.makeMeasureName("n", "cid"),
                    startMark: BPM.makeStartMark("n", "cid"),
                    endMark: BPM.makeEndMark("n", "cid"),
                    supported: BPM.supportsBrowserPerformance(),
                    instance: {
                        measureName: inst.measureName,
                        startMark: inst.startMark,
                        endMark: inst.endMark,
                    },
                };
                const spc = new lib.StubPerformanceClient();
                const sm = spc.startMeasurement("measure-x", "cid-x");
                const stubPerf = {
                    generateId: spc.generateId(),
                    startKeys: Object.keys(sm).sort(),
                    eventShape: {
                        ...sm.event,
                        eventId: typeof sm.event.eventId,
                        startTimeMs: typeof sm.event.startTimeMs,
                    },
                    end: sm.end(),
                    endMeasurement: spc.endMeasurement(),
                    addCb: spc.addPerformanceCallback(() => {}),
                    removeCb: spc.removePerformanceCallback("x"),
                };

                const erp = {
                    notMcp:
                        lib.enforceResourceParameter(false, {}) === undefined,
                    missingResource: errShape(() =>
                        lib.enforceResourceParameter(true, {
                            correlationId: "c",
                        })
                    ),
                    misplaced: errShape(() =>
                        lib.enforceResourceParameter(true, {
                            resource: "r",
                            extraQueryParameters: { resource: "x" },
                        })
                    ),
                    ok:
                        lib.enforceResourceParameter(true, {
                            resource: "https://r",
                        }) === undefined,
                };

                // SignedHttpRequest: keygen + SHR signing round-trip
                const shr = new lib.SignedHttpRequest({
                    resourceRequestMethod: "get",
                    resourceRequestUri:
                        "https://api.contoso.com/path/x?a=b",
                });
                const kid = await shr.generatePublicKeyThumbprint();
                const jwt = await shr.signRequest("the-at", kid, {
                    extra: "claim",
                });
                const dec = (s) => {
                    const b = s.replace(/-/g, "+").replace(/_/g, "/");
                    return JSON.parse(
                        atob(b + "=".repeat((4 - (b.length % 4)) % 4))
                    );
                };
                const [h, p] = jwt.split(".");
                const header = dec(h);
                const payload = dec(p);
                const signedHttpRequest = {
                    kidShape: typeof kid === "string" && kid.length === 43,
                    parts: jwt.split(".").length,
                    header: {
                        typ: header.typ,
                        alg: header.alg,
                        kidMatches: dec(header.kid).kid === kid,
                    },
                    payload: {
                        at: payload.at,
                        m: payload.m,
                        u: payload.u,
                        p: payload.p,
                        q: payload.q,
                        extra: payload.extra,
                        client_claims: payload.client_claims ?? null,
                        nonceIsString: typeof payload.nonce === "string",
                        tsIsNumber: typeof payload.ts === "number",
                        cnfJwkKeys: Object.keys(payload.cnf.jwk).sort(),
                    },
                    removeKeysUndef:
                        (await shr.removeKeys(kid, "cid")) === undefined,
                };

                return {
                    types,
                    constants,
                    browserUtils,
                    stubbedPCA,
                    headerParser,
                    emu,
                    eventHandler,
                    memoryStorage,
                    sessionStorageShape,
                    localStorageShape,
                    bpm,
                    stubPerf,
                    erp,
                    signedHttpRequest,
                };
            });
        },
        timeout: 30000,
    },
];
