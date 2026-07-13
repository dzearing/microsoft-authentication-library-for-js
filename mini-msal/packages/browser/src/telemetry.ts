/**
 * Telemetry feature: performance events, opt-in like real MSAL — events fire
 * only when the app passes `telemetry.client: new BrowserPerformanceClient()`
 * in config (real's default is a stub that swallows everything). Compose via
 * createClient(config, [popup, telemetry]) — telemetry LAST so it can wrap
 * the methods other features attach.
 *
 * Events replicate real's full BrowserPerformanceClient shape (C11): the
 * per-flow top-level fields plus the `ext` sub-measurement aggregate keyed by
 * real's internal operation names. Key sets and deterministic values are
 * snapshot-pinned per flow (init / popup / silent cache-hit / silent refresh /
 * ssoSilent / silent failure); durations, ids and network stats are live.
 */
import type {
    AccountInfo,
    AuthenticationResult,
    ClientContext,
    PerfClient,
    TokenEntity,
} from "./index.js";
import { LIB_NAME, version } from "./index.js";

/** One completed top-level operation, in real's full emitted shape. */
export interface PerformanceEvent {
    name: string;
    correlationId: string;
    durationMs: number;
    success: boolean;
    errorCode?: string;
    [field: string]: unknown;
}

type PerfCallback = (events: PerformanceEvent[]) => void;

/** The methods this feature attaches to the client. */
export interface TelemetryClient {
    addPerformanceCallback(cb: PerfCallback): string;
    removePerformanceCallback(id: string): boolean;
}

/** Opt-in perf client, config `telemetry.client` (like real's). */
export class BrowserPerformanceClient implements PerfClient {
    private cbs = new Map<string, PerfCallback>();
    constructor(_config?: unknown) {}
    addPerformanceCallback(cb: PerfCallback): string {
        const id = crypto.randomUUID();
        this.cbs.set(id, cb);
        return id;
    }
    removePerformanceCallback(id: string): boolean {
        return this.cbs.delete(id);
    }
    emitEvents(events: PerformanceEvent[]): void {
        this.cbs.forEach((cb) => cb(events));
    }
}

// ---- real's per-flow sub-measurement names ("ext" aggregate) ---------------
// Suffix "!" = CallCount key only, "+" = DurationMs key only, else both.
const DISC =
    "authorityFactoryCreateDiscoveredInstance authorityResolveEndpointsAsync authorityUpdateCloudDiscoveryMetadata authorityUpdateEndpointMetadata standardInteractionClientGetClientConfiguration standardInteractionClientGetDiscoveredAuthority";
const NETDISC = "authorityGetEndpointMetadataFromNetwork";
const PKCE =
    "generateCodeChallengeFromVerifier generateCodeVerifier generatePkceCodes getRandomValues sha256Digest";
const CODE =
    "authClientAcquireToken authClientCreateTokenRequestBody authClientExecuteTokenRequest authorizationCodeClientExecutePostToTokenEndpoint deserializeResponse getAuthCodeUrl getStandardParams handleCodeResponse handleCodeResponseFromServer handleResponseCode handleServerTokenResponse networkClientSendPostRequestAsync! setUserData";
const STD =
    "initializeBaseRequest standardInteractionClientCreateAuthCodeClient standardInteractionClientInitializeAuthorizationRequest";
const SIL =
    "acquireTokenFromCache acquireTokenSilentAsync initializeBaseRequest initializeSilentRequest silentCacheClientAcquireToken silentFlowClientAcquireCachedToken";
const RT =
    "acquireTokenByRefreshToken cacheManagerGetRefreshToken networkClientSendPostRequestAsync! refreshTokenClientAcquireToken refreshTokenClientAcquireTokenByRefreshToken refreshTokenClientAcquireTokenWithCachedRefreshToken refreshTokenClientCreateTokenRequestBody refreshTokenClientExecutePostToTokenEndpoint refreshTokenClientExecuteTokenRequest silentRefreshClientAcquireToken";
const IFRAME =
    "removeHiddenIframe silentHandlerInitiateAuthRequest silentHandlerMonitorIframeForHash! silentIframeClientTokenHelper";
// extras on the silent RT-fail -> iframe-fail path beyond the groups above
const FAILX =
    "acquireTokenBySilentIframe deserializeResponse getAuthCodeUrl getStandardParams handleCodeResponse handleResponseCode silentIframeClientAcquireToken standardInteractionClientCreateAuthCodeClient standardInteractionClientInitializeAuthorizationRequest";

const join = (...parts: (string | false)[]) =>
    parts.filter(Boolean).join(" ");

function subMeasurements(groups: string): Record<string, number> {
    const ext: Record<string, number> = {};
    for (const tok of groups.split(" ")) {
        const mark = tok.endsWith("!") ? 1 : tok.endsWith("+") ? 2 : 0;
        const name = mark ? tok.slice(0, -1) : tok;
        if (mark !== 2) ext[name + "CallCount"] = 1;
        if (mark !== 1) ext[name + "DurationMs"] = 0;
    }
    return ext;
}

// module-wide instance counters, like real's (reset per page load)
let instanceCount = 0;
const instancesByClientId: Record<string, number> = {};

const MSA_TID = "9188040d-6c67-4c5b-b112-36a304b66dad";
function accountType(account?: AccountInfo | null): string | undefined {
    const claims = account?.idTokenClaims as Record<string, unknown> | undefined;
    if (claims?.tfp || claims?.acr) return "B2C";
    if (!claims?.tid) return undefined;
    return claims.tid === MSA_TID ? "MSA" : "AAD";
}

export function telemetry(ctx: ClientContext): void {
    const c = ctx.client;
    const config = ctx.config;
    const pc = config.telemetry?.client;
    // registration works without an opt-in client, like real's stub — the
    // callback just never fires
    c.addPerformanceCallback = (cb: PerfCallback): string =>
        pc?.addPerformanceCallback(cb) ?? crypto.randomUUID();
    c.removePerformanceCallback = (id: string): boolean =>
        pc?.removePerformanceCallback(id) ?? false;
    if (!pc) return;

    const clientId = config.auth.clientId;
    instanceCount++;
    instancesByClientId[clientId] =
        (instancesByClientId[clientId] ?? 0) + 1;
    const conn = (navigator as any).connection ?? {};
    // real resolves the authority over the network once per instance; every
    // event after the first reports the in-memory metadata cache
    let firstResolve = true;
    // homeAccountId -> API that last wrote the account (real's accountCachedBy)
    const cachedBy = new Map<string, string>();

    const netSource = () => {
        const src = firstResolve ? "network" : "cache";
        firstResolve = false;
        return src;
    };
    const collect = (type: "accessToken" | "idToken" | "refreshToken") => {
        const all: TokenEntity[] = [];
        ctx.findToken(type, (t) => (all.push(t), false));
        return all;
    };
    const rtSize = (hid: string) =>
        ctx.findToken("refreshToken", (t) => t.homeAccountId === hid)?.secret
            .length ?? 0;
    const scopesIntersect = (target: string | undefined, scopes: string[]) => {
        const t = (target ?? "").toLowerCase().split(" ");
        return scopes.some((sc) => t.includes(sc.toLowerCase()));
    };
    /** pre-call cache snapshot for the counters real derives while saving */
    const snapshot = (req: any) => ({
        req,
        accounts: new Set(
            c.getAllAccounts().map((a: AccountInfo) => a.homeAccountId)
        ),
        ats: collect("accessToken"),
    });
    type Pre = ReturnType<typeof snapshot>;
    const removedCount = (pre: Pre, r: AuthenticationResult) =>
        pre.ats.filter(
            (t) =>
                t.homeAccountId === r.account.homeAccountId &&
                scopesIntersect(t.target, r.scopes)
        ).length;
    const successCommon = (r: AuthenticationResult, pre: Pre) => ({
        accessTokenSize: r.accessToken.length,
        idTokenSize: r.idToken.length,
        accountType: accountType(r.account),
        dataBoundary: r.account?.dataBoundary,
        // counted while saving the account — cache hits don't carry it
        ...(!r.fromCache && {
            cacheMatchedAccounts: pre.accounts.has(r.account.homeAccountId)
                ? 1
                : 0,
        }),
        cloudDiscoverySource: "config",
        scenarioId: pre.req?.scenarioId,
    });

    const emitEvent = (
        name: string,
        success: boolean,
        cid: string | undefined,
        started: number,
        startVis: string,
        ext: Record<string, number>,
        fields: Record<string, unknown>
    ) => {
        const durationMs = Math.round(performance.now() - started);
        pc.emitEvents([
            {
                eventId: crypto.randomUUID(),
                status: 2,
                authority:
                    config.auth.authority ??
                    "https://login.microsoftonline.com/common",
                libraryName: LIB_NAME,
                libraryVersion: version,
                clientId,
                name,
                startTimeMs: Date.now() - durationMs,
                correlationId: cid ?? crypto.randomUUID(),
                appName: "",
                appVersion: "",
                durationMs,
                success,
                startPageVisibility: startVis,
                startOnlineStatus: navigator.onLine,
                endPageVisibility: document.visibilityState,
                networkEffectiveType: conn.effectiveType,
                networkRtt: conn.rtt,
                incompleteSubMeasurements: undefined,
                incompleteSubsCount: 0,
                context: JSON.stringify({ [name]: { dur: durationMs } }),
                logs: "",
                ext,
                ...fields,
            },
        ]);
    };

    const origInit = c.initialize;
    c.initialize = async (...args: any[]) => {
        const t0 = performance.now();
        const vis0 = document.visibilityState;
        try {
            const r = await origInit(...args);
            const accounts = c.getAllAccounts().length;
            emitEvent(
                "initializeClientApplication",
                true,
                undefined,
                t0,
                vis0,
                subMeasurements("initializeCache"),
                {
                    allowPlatformBroker: !!config.system?.allowPlatformBroker,
                    cacheLocation:
                        config.cache?.cacheLocation ?? "sessionStorage",
                    cacheRetentionDays: 5,
                    isMcp: false,
                    msalInstanceCount: instanceCount,
                    sameClientIdInstanceCount: instancesByClientId[clientId],
                    preMigrateAcntCount: accounts,
                    preMigrateATCount: collect("accessToken").length,
                    preMigrateITCount: collect("idToken").length,
                    preMigrateRTCount: collect("refreshToken").length,
                    postMigrateAcntCount: accounts,
                    postMigrateATCount: collect("accessToken").length,
                    postMigrateITCount: collect("idToken").length,
                    postMigrateRTCount: collect("refreshToken").length,
                }
            );
            return r;
        } catch (e: any) {
            emitEvent(
                "initializeClientApplication",
                false,
                undefined,
                t0,
                vis0,
                {},
                { errorCode: e?.errorCode }
            );
            throw e;
        }
    };

    const origPopup = c.acquireTokenPopup;
    if (origPopup) {
        c.acquireTokenPopup = async (req: any) => {
            const t0 = performance.now();
            const vis0 = document.visibilityState;
            const pre = snapshot(req);
            try {
                const r = await origPopup(req);
                const src = netSource();
                emitEvent(
                    "acquireTokenPopup",
                    true,
                    r.correlationId,
                    t0,
                    vis0,
                    subMeasurements(
                        join(DISC, src === "network" && NETDISC, PKCE, CODE, STD)
                    ),
                    {
                        ...successCommon(r, pre),
                        authorityEndpointSource: src,
                        httpVerToken: "",
                        isAsyncPopup:
                            config.system?.navigatePopups === false,
                        kmsi: false,
                        lateResponseExperimentEnabled: false,
                        redirectBridgeMessageVersion: 1,
                        redirectBridgeTimeoutMs:
                            config.system?.popupBridgeTimeout ?? 60000,
                        refreshTokenSize: rtSize(r.account.homeAccountId),
                        requestId: crypto.randomUUID(),
                        usePreGeneratedPkce: false,
                    }
                );
                cachedBy.set(r.account.homeAccountId, "acquireTokenPopup");
                return r;
            } catch (e: any) {
                netSource();
                emitEvent(
                    "acquireTokenPopup",
                    false,
                    req?.correlationId,
                    t0,
                    vis0,
                    subMeasurements(join(DISC, PKCE, STD, "getAuthCodeUrl getStandardParams")),
                    {
                        errorCode: e?.errorCode,
                        subErrorCode: e?.subError ?? "",
                        scenarioId: req?.scenarioId,
                    }
                );
                throw e;
            }
        };
    }

    const origSilent = c.acquireTokenSilent;
    c.acquireTokenSilent = async (req: any) => {
        const t0 = performance.now();
        const vis0 = document.visibilityState;
        const pre = snapshot(req);
        try {
            const r = await origSilent(req);
            const src = netSource();
            const hid = r.account.homeAccountId;
            const silCommon = {
                ...successCommon(r, pre),
                ...(cachedBy.has(hid) && {
                    accountCachedBy: cachedBy.get(hid),
                }),
                authorityEndpointSource: src,
                cacheLookupPolicy: req?.cacheLookupPolicy,
                deduped: false,
                isNativeBroker: false,
                ssoCapable: undefined,
                visibilityChangeCount: 0,
                onlineStatusChangeCount: 0,
            };
            if (r.fromCache) {
                emitEvent(
                    "acquireTokenSilent",
                    true,
                    r.correlationId,
                    t0,
                    vis0,
                    subMeasurements(
                        join(
                            DISC,
                            src === "network" && NETDISC,
                            SIL,
                            "silentFlowClientGenerateResultFromCacheRecord"
                        )
                    ),
                    { ...silCommon, cacheOutcome: "0", fromCache: true }
                );
            } else {
                const removed = removedCount(pre, r);
                emitEvent(
                    "acquireTokenSilent",
                    true,
                    r.correlationId,
                    t0,
                    vis0,
                    subMeasurements(
                        join(
                            DISC,
                            src === "network" && NETDISC,
                            SIL,
                            RT,
                            "handleServerTokenResponse setUserData"
                        )
                    ),
                    {
                        ...silCommon,
                        ...(removed && { accessTokensRemoved: removed }),
                        cacheOutcome: "3",
                        errorCode: undefined,
                        subErrorCode: undefined,
                        fromCache: false,
                        httpVerToken: "",
                        kmsi: false,
                        refreshTokenSize: rtSize(hid),
                        requestId: crypto.randomUUID(),
                    }
                );
                cachedBy.set(hid, "acquireTokenSilent_silentFlow");
            }
            return r;
        } catch (e: any) {
            const src = netSource();
            emitEvent(
                "acquireTokenSilent",
                false,
                req?.correlationId,
                t0,
                vis0,
                subMeasurements(
                    join(
                        DISC,
                        src === "network" && NETDISC,
                        SIL,
                        RT,
                        PKCE,
                        IFRAME,
                        FAILX
                    )
                ),
                {
                    ...(req?.account && {
                        accountType: accountType(req.account),
                        dataBoundary: req.account.dataBoundary,
                    }),
                    authorityEndpointSource: src,
                    cacheLookupPolicy: req?.cacheLookupPolicy,
                    cacheOutcome: "3",
                    cloudDiscoverySource: "config",
                    deduped: false,
                    errorCode: e?.errorCode,
                    httpVerToken: "",
                    lateResponseExperimentEnabled: false,
                    loginHintFromUpn: true,
                    prompt: "none",
                    redirectBridgeMessageVersion: 1,
                    redirectBridgeTimeoutMs:
                        config.system?.iframeBridgeTimeout ?? 10000,
                    refreshTokenSize: 0,
                    requestId: crypto.randomUUID(),
                    scenarioId: req?.scenarioId,
                    ...(e?.silentRefreshReason && {
                        silentRefreshReason: e.silentRefreshReason,
                    }),
                    ssoCapable: undefined,
                    subErrorCode: e?.subError ?? "",
                    visibilityChangeCount: 0,
                    onlineStatusChangeCount: 0,
                }
            );
            throw e;
        }
    };

    const origSso = c.ssoSilent;
    c.ssoSilent = async (req: any) => {
        const t0 = performance.now();
        const vis0 = document.visibilityState;
        const pre = snapshot(req);
        try {
            const r = await origSso(req);
            const src = netSource();
            const hid = r.account.homeAccountId;
            const removed = removedCount(pre, r);
            emitEvent(
                "ssoSilent",
                true,
                r.correlationId,
                t0,
                vis0,
                subMeasurements(
                    join(
                        DISC,
                        src === "network" && NETDISC,
                        PKCE,
                        CODE,
                        STD,
                        IFRAME
                    )
                ),
                {
                    ...successCommon(r, pre),
                    ...(removed && { accessTokensRemoved: removed }),
                    ...(cachedBy.has(hid) && {
                        accountCachedBy: cachedBy.get(hid),
                    }),
                    authorityEndpointSource: src,
                    httpVerToken: "",
                    isNativeBroker: false,
                    kmsi: false,
                    lateResponseExperimentEnabled: false,
                    ...(req?.loginHint && { loginHintFromRequest: true }),
                    prompt: "none",
                    redirectBridgeMessageVersion: 1,
                    redirectBridgeTimeoutMs:
                        config.system?.iframeBridgeTimeout ?? 10000,
                    refreshTokenSize: rtSize(hid),
                    requestId: crypto.randomUUID(),
                    ssoCapable: undefined,
                    visibilityChangeCount: 0,
                    onlineStatusChangeCount: 0,
                }
            );
            cachedBy.set(hid, "ssoSilent_silentFlow");
            return r;
        } catch (e: any) {
            const src = netSource();
            emitEvent(
                "ssoSilent",
                false,
                req?.correlationId,
                t0,
                vis0,
                subMeasurements(
                    join(DISC, src === "network" && NETDISC, PKCE, STD, IFRAME)
                ),
                {
                    errorCode: e?.errorCode,
                    subErrorCode: e?.subError ?? "",
                    prompt: "none",
                    scenarioId: req?.scenarioId,
                }
            );
            throw e;
        }
    };
}
