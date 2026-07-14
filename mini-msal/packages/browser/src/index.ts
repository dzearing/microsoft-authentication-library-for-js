/**
 * mini-msal core: a from-scratch minimal implementation of the msal-browser
 * behaviors every app pays for, exposed as a composable, closure-based
 * client factory (pay-to-play architecture):
 *   - createClient(config, features) — core + tree-shakable feature modules
 *   - OIDC discovery (authority metadata fetched at runtime)
 *   - loginRedirect / acquireTokenRedirect (auth-code + PKCE S256)
 *   - handleRedirectPromise (code exchange at the token endpoint)
 *   - acquireTokenSilent: cache -> refresh-token grant -> hidden-iframe
 *     prompt=none fallback; ssoSilent
 *   - single-account sessionStorage cache + active account
 *   - getAllAccounts / getAccount / getActiveAccount / setActiveAccount
 *   - event callbacks (EventMessage-shaped, add/remove)
 *   - logoutRedirect
 *   - InteractionRequiredAuthError / BrowserAuthError classification
 *
 * Optional features ship as subpath exports (./popup, later ./broker, ./naa,
 * …): plain functions that receive the client's internal ClientContext and
 * attach methods to it — no classes, near-zero seam bytes. The
 * @mini-msal/compat package composes core + ALL features into the classic
 * PublicClientApplication drop-in.
 */

/**
 * Key set (including undefined-valued keys) matches real MSAL's AccountInfo —
 * event payloads expose Object.keys, so presence matters, not just values.
 */
export interface AccountInfo {
    homeAccountId: string;
    environment: string;
    username: string;
    localAccountId: string;
    tenantId: string;
    name?: string;
    idTokenClaims: Record<string, unknown>;
    idToken?: string;
    authorityType?: string;
    /** public shape is a Map keyed by tenantId, like real's AccountInfo */
    tenantProfiles?: Map<string, unknown>;
    nativeAccountId?: string;
    dataBoundary?: string;
    kmsi?: boolean;
    loginHint?: string;
    upn?: string;
}

/** Same key set as real MSAL's AuthenticationResult (see AccountInfo note). */
export interface AuthenticationResult {
    authority: string;
    uniqueId: string;
    tenantId: string;
    scopes: string[];
    account: AccountInfo;
    idToken: string;
    idTokenClaims: Record<string, unknown>;
    accessToken: string;
    fromCache: boolean;
    expiresOn: Date;
    extExpiresOn?: Date;
    refreshOn?: Date;
    correlationId: string;
    requestId: string;
    familyId: string;
    tokenType: string;
    /** custom request state (or "") on interactive flows; undefined on silent */
    state?: string;
    cloudGraphHostName?: string;
    msGraphHost?: string;
    code?: string;
    fromPlatformBroker: boolean;
}

/** popup size/position overrides (real clamps them to the parent window) */
export interface PopupWindowAttributes {
    popupSize?: { height?: number; width?: number };
    popupPosition?: { top?: number; left?: number };
}

export interface TokenRequest {
    scopes: string[];
    account?: AccountInfo;
    loginHint?: string;
    sid?: string;
    domainHint?: string;
    prompt?: string;
    redirectUri?: string;
    /** per-request authority override (no re-discovery, like real) */
    authority?: string;
    /** custom state, echoed on result.state (wire: `<libState>|<custom>`) */
    state?: string;
    /** claims JSON, merged with defaults + clientCapabilities xms_cc */
    claims?: string;
    extraQueryParameters?: Record<string, string>;
    /** accepted for compat; real 5.16 ignores it at runtime (its
     * extraQueryParameters ride the token-endpoint query instead) */
    tokenQueryParameters?: Record<string, string>;
    cacheLookupPolicy?: number;
    forceRefresh?: boolean;
    correlationId?: string;
    /** redirect flows: page to return to after the roundtrip (defaults to
     * the current page when navigateToLoginRequestUrl is on) */
    redirectStartPage?: string;
    popupWindowAttributes?: PopupWindowAttributes;
    popupWindowParent?: Window;
}

export interface Config {
    auth: {
        clientId: string;
        authority?: string;
        redirectUri?: string;
        postLogoutRedirectUri?: string;
        clientCapabilities?: string[];
        /** return to the initiating page after a redirect (default true) */
        navigateToLoginRequestUrl?: boolean;
        /** called with the URL before any redirect navigation; return false
         * to cancel the navigation */
        onRedirectNavigate?: (url: string) => boolean | void;
    };
    system?: {
        /** open popups synchronously on about:blank inside the user gesture,
         * then navigate them (real's default true); false defers the open
         * until the authorize URL is ready */
        navigatePopups?: boolean;
        popupBridgeTimeout?: number;
        iframeBridgeTimeout?: number;
        /** ms before a redirect navigation is considered failed (default 30s) */
        redirectNavigationTimeout?: number;
        /** custom navigation implementation for redirect flows */
        navigationClient?: NavigationClient;
        /** probe/use the platform broker (./broker feature) */
        allowPlatformBroker?: boolean;
        /** extension-transport Handshake timeout, ms (default 2000) */
        nativeBrokerHandshakeTimeout?: number;
        loggerOptions?: {
            loggerCallback?: (
                level: number,
                message: string,
                containsPii: boolean
            ) => void;
            /** real's LogLevel: Error 0, Warning 1, Info 2 (default), Verbose 3, Trace 4 */
            logLevel?: number;
            piiLoggingEnabled?: boolean;
        };
    };
    /** sessionStorage (default) built in; "localStorage" needs the
     * ./local-storage feature composed */
    cache?: { cacheLocation?: string; cacheRetentionDays?: number };
    experimental?: Record<string, unknown>;
    /** perf events opt-in: `client: new BrowserPerformanceClient()` (./telemetry) */
    telemetry?: {
        client?: PerfClient;
    };
}

/** Duck-typed seam for ./telemetry's BrowserPerformanceClient — the core
 * never imports the feature module. */
export interface PerfClient {
    addPerformanceCallback(cb: (events: any[]) => void): string;
    removePerformanceCallback(id: string): boolean;
    emitEvents(events: any[]): void;
}

/** default error prose, same as real MSAL's */
const AKA = (code: string) =>
    `See https://aka.ms/msal.js.errors#${code} for details`;

/**
 * Real MSAL's literal wire identity (Decision Log: impersonation) — keep in
 * one block so a future un-impersonation is a one-line change. Real browser
 * 5.16 uses StubServerTelemetryManager, so the telemetry params are sent as
 * EMPTY strings on every token request (verified in snapshots).
 */
const WIRE_ID = {
    "x-client-SKU": "msal.js.browser",
    "x-client-VER": "5.16.0",
} as const;
const TOKEN_TELEMETRY = {
    ...WIRE_ID,
    "x-client-current-telemetry": "",
    "x-client-last-telemetry": "",
    "x-ms-lib-capability": "retry-after, h429",
} as const;

/** exported version string = real's (Decision Log: impersonation) */
export const version = WIRE_ID["x-client-VER"];
/** libraryName as emitted on perf events (./telemetry) — same impersonation block */
export const LIB_NAME = "@azure/msal-browser";

/** real's LogLevel numeric enum, incl. TS reverse mappings */
export const LogLevel: Record<string | number, string | number> = {
    Error: 0,
    Warning: 1,
    Info: 2,
    Verbose: 3,
    Trace: 4,
    0: "Error",
    1: "Warning",
    2: "Info",
    3: "Verbose",
    4: "Trace",
};

/** wrapper-library SKUs for initializeWrapperLibrary (real's WrapperSKU) */
export const WrapperSKU = {
    React: "@azure/msal-react",
    Angular: "@azure/msal-angular",
} as const;

export interface NavigationOptions {
    apiId: number;
    timeout: number;
    noHistory: boolean;
}

/**
 * real's NavigationClient: every redirect navigation routes through a
 * pluggable client (config.system.navigationClient / setNavigationClient) so
 * apps can substitute SPA-router navigation. The default navigates the window
 * and returns a promise that only settles by rejecting after `timeout`.
 */
export class NavigationClient {
    navigateInternal(url: string, options: NavigationOptions) {
        return NavigationClient.defaultNavigateWindow(url, options);
    }
    navigateExternal(url: string, options: NavigationOptions) {
        return NavigationClient.defaultNavigateWindow(url, options);
    }
    static defaultNavigateWindow(
        url: string,
        options: NavigationOptions
    ): Promise<boolean> {
        if (options.noHistory) {
            location.replace(url);
        } else {
            location.assign(url);
        }
        return new Promise((_, reject) => {
            setTimeout(() => {
                reject(new BrowserAuthError("timed_out", "failed_to_redirect"));
            }, options.timeout);
        });
    }
}

/** real's normalizeUrlForComparison: drop hash, ensure trailing slash */
const normUrl = (u: string): string => {
    if (!u) return u;
    try {
        const x = new URL(u.split("#")[0]);
        if (!x.pathname.endsWith("/")) x.pathname += "/";
        return x.href;
    } catch {
        return u;
    }
};

type LoggerOptions = NonNullable<Config["system"]>["loggerOptions"];

/**
 * real's Logger: level-gated callback with real's message format
 * `[<UTC time>] : [<correlationId>] : <pkg>@<ver> : <Level> - <msg>`.
 * Wrapper libraries (msal-react/-angular) clone() it with their own SKU.
 */
export class Logger {
    private cb: NonNullable<
        NonNullable<LoggerOptions>["loggerCallback"]
    >;
    private level: number;
    private pii: boolean;
    private pkg: string;
    constructor(
        options?: LoggerOptions,
        packageName = "",
        packageVersion = ""
    ) {
        this.cb = options?.loggerCallback ?? (() => {});
        this.level = options?.logLevel ?? 2;
        this.pii = options?.piiLoggingEnabled ?? false;
        this.pkg = `${packageName}@${packageVersion}`;
    }
    clone(packageName: string, packageVersion: string): Logger {
        return new Logger(
            {
                loggerCallback: this.cb,
                logLevel: this.level,
                piiLoggingEnabled: this.pii,
            },
            packageName,
            packageVersion
        );
    }
    isPiiLoggingEnabled(): boolean {
        return this.pii;
    }
    /** same semi-public signature as real's Logger.logMessage */
    logMessage(
        msg: string,
        options: {
            logLevel: number;
            containsPii?: boolean;
            correlationId?: string;
        }
    ): void {
        const { logLevel, containsPii = false, correlationId = "" } = options;
        if (logLevel > this.level || (containsPii && !this.pii)) return;
        this.cb(
            logLevel,
            `[${new Date().toUTCString()}] : [${correlationId}] : ${this.pkg} : ${LogLevel[logLevel]} - ${msg}`,
            containsPii
        );
    }
    error(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 0, correlationId: cid });
    }
    errorPii(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 0, containsPii: true, correlationId: cid });
    }
    warning(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 1, correlationId: cid });
    }
    warningPii(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 1, containsPii: true, correlationId: cid });
    }
    info(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 2, correlationId: cid });
    }
    infoPii(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 2, containsPii: true, correlationId: cid });
    }
    verbose(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 3, correlationId: cid });
    }
    verbosePii(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 3, containsPii: true, correlationId: cid });
    }
    trace(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 4, correlationId: cid });
    }
    tracePii(m: string, cid?: string) {
        this.logMessage(m, { logLevel: 4, containsPii: true, correlationId: cid });
    }
}

export class AuthError extends Error {
    name = "AuthError";
    constructor(
        public errorCode: string,
        public errorMessage: string = AKA(errorCode),
        public subError: string = ""
    ) {
        super(`${errorCode}: ${errorMessage}`);
    }
}

export class InteractionRequiredAuthError extends AuthError {
    name = "InteractionRequiredAuthError";
}
export class ServerError extends AuthError {
    name = "ServerError";
}
export class ClientAuthError extends AuthError {
    name = "ClientAuthError";
}
export class ClientConfigurationError extends AuthError {
    name = "ClientConfigurationError";
}
export class NestedAppAuthError extends AuthError {
    name = "NestedAppAuthError";
}
/** browser-layer errors always carry the aka.ms message, like real MSAL */
export class BrowserAuthError extends AuthError {
    name = "BrowserAuthError";
    constructor(errorCode: string, subError?: string) {
        super(errorCode, undefined, subError);
    }
}

// same interaction-required detection as real MSAL (code / description /
// suberror, contains-match on the description)
const IR_CODES =
    /interaction_required|consent_required|login_required|bad_token|ui_not_allowed|interrupted_user/;
const IR_SUBS =
    /^(message_only|additional_action|basic_action|user_password_expired|consent_required|bad_token|ui_not_allowed|interrupted_user)$/;

function classify(code: string, desc = "", sub = ""): AuthError {
    return IR_CODES.test(code) || IR_CODES.test(desc) || IR_SUBS.test(sub)
        ? new InteractionRequiredAuthError(code, desc || undefined, sub)
        : new ServerError(code, desc || undefined, sub);
}

// exactly real msal-browser 5.16's EventType map (no accountAdded/
// accountRemoved/loginFailure — those left real's surface)
export const EventType = {
    INITIALIZE_START: "msal:initializeStart",
    INITIALIZE_END: "msal:initializeEnd",
    ACTIVE_ACCOUNT_CHANGED: "msal:activeAccountChanged",
    LOGIN_SUCCESS: "msal:loginSuccess",
    ACQUIRE_TOKEN_START: "msal:acquireTokenStart",
    BROKERED_REQUEST_START: "msal:brokeredRequestStart",
    ACQUIRE_TOKEN_SUCCESS: "msal:acquireTokenSuccess",
    BROKERED_REQUEST_SUCCESS: "msal:brokeredRequestSuccess",
    ACQUIRE_TOKEN_FAILURE: "msal:acquireTokenFailure",
    BROKERED_REQUEST_FAILURE: "msal:brokeredRequestFailure",
    ACQUIRE_TOKEN_NETWORK_START: "msal:acquireTokenFromNetworkStart",
    HANDLE_REDIRECT_START: "msal:handleRedirectStart",
    HANDLE_REDIRECT_END: "msal:handleRedirectEnd",
    POPUP_OPENED: "msal:popupOpened",
    LOGOUT_START: "msal:logoutStart",
    LOGOUT_SUCCESS: "msal:logoutSuccess",
    LOGOUT_FAILURE: "msal:logoutFailure",
    LOGOUT_END: "msal:logoutEnd",
    RESTORE_FROM_BFCACHE: "msal:restoreFromBFCache",
    BROKER_CONNECTION_ESTABLISHED: "msal:brokerConnectionEstablished",
} as const;

export const InteractionType = {
    Redirect: "redirect",
    Popup: "popup",
    Silent: "silent",
    None: "none",
} as const;
export type InteractionKind =
    (typeof InteractionType)[keyof typeof InteractionType];

/** real's InteractionStatus (wrapper libraries derive it from events) */
export const InteractionStatus = {
    Startup: "startup",
    Logout: "logout",
    AcquireToken: "acquireToken",
    HandleRedirect: "handleRedirect",
    None: "none",
} as const;

export const CacheLookupPolicy = {
    Default: 0,
    AccessToken: 1,
    AccessTokenAndRefreshToken: 2,
    RefreshToken: 3,
    RefreshTokenAndNetwork: 4,
    Skip: 5,
} as const;

export interface EventMessage {
    eventType: string;
    interactionType: string | null;
    payload: unknown;
    error: unknown;
    timestamp: number;
}

type EventCallback = (message: EventMessage) => void;

export interface AccountFilter {
    homeAccountId?: string;
    localAccountId?: string;
    username?: string;
}

// ---- small utils ----------------------------------------------------------

const enc = new TextEncoder();

function b64url(bytes: ArrayBuffer): string {
    return btoa(String.fromCharCode(...new Uint8Array(bytes)))
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");
}

function randomString(): string {
    return b64url(crypto.getRandomValues(new Uint8Array(32)).buffer);
}

async function pkce(): Promise<{ verifier: string; challenge: string }> {
    const verifier = randomString();
    const challenge = b64url(
        await crypto.subtle.digest("SHA-256", enc.encode(verifier))
    );
    return { verifier, challenge };
}

export function decodeJwt(token: string): Record<string, any> {
    return JSON.parse(
        atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );
}

/** signin_state claim contains kmsi/dvc_dmjd (real's AuthToken.isKmsi) */
export function isKmsi(claims: Record<string, any>): boolean {
    return !!(claims.signin_state as string[] | undefined)?.some((v) =>
        ["kmsi", "dvc_dmjd"].includes(v.trim().toLowerCase())
    );
}

/** never the ONLY scopes real considers when deduping intersecting ATs */
const OIDC_SCOPES = ["openid", "profile", "email", "offline_access"];

/** cache environment: real MSAL uses the cloud's preferred_cache host */
const preferredEnv = (authority: string): string => {
    const host = new URL(authority).host;
    return /login\.microsoftonline\.com|login\.microsoft\.com|sts\.windows\.net/.test(
        host
    )
        ? "login.windows.net"
        : host;
};

/** real's createAccountEntityFromAccountInfo (hydrateCache,
 * loadExternalTokens with request.account) */
const entityFromAccountInfo = (
    a: AccountInfo,
    apiId: number,
    graphHosts?: { cloudGraphHostName?: string; msGraphHost?: string }
): AccountEntity => ({
    homeAccountId: a.homeAccountId,
    environment: a.environment,
    realm: a.tenantId,
    localAccountId: a.localAccountId,
    username: a.username,
    authorityType: a.authorityType ?? "MSSTS",
    name: a.name,
    nativeAccountId: a.nativeAccountId,
    tenantProfiles: [...(a.tenantProfiles?.values() ?? [])],
    lastUpdatedAt: String(Date.now()),
    cachedByApiId: apiId,
    ...graphHosts,
});

async function post(
    url: string,
    body: Record<string, string | undefined>,
    store: Store,
    throttleKey: string
) {
    // real's ThrottlingUtils: a cached 429/5xx/Retry-After token response
    // blocks identical requests (same thumbprint) until throttleTime — the
    // retry re-throws the stored error with NO network call
    const t = JSON.parse(store.get(throttleKey) ?? "null");
    if (t) {
        if (t.throttleTime >= Date.now()) {
            throw new ServerError(
                t.errorCodes?.join(" ") || "",
                t.errorMessage,
                t.subError
            );
        }
        store.remove(throttleKey);
    }
    const res = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
        },
        body: new URLSearchParams(
            Object.entries(body).filter(
                (e): e is [string, string] => e[1] !== undefined
            )
        ).toString(),
    });
    const json = await res.json();
    if (
        res.status === 429 ||
        res.status >= 500 ||
        (res.headers.has("Retry-After") && !res.ok)
    ) {
        const sec = Date.now() / 1000;
        store.set(
            throttleKey,
            JSON.stringify({
                throttleTime: Math.floor(
                    Math.min(
                        sec +
                            (parseInt(res.headers.get("Retry-After")!) || 60),
                        sec + 3600
                    ) * 1000
                ),
                error: json.error,
                errorCodes: json.error_codes,
                errorMessage: json.error_description,
                subError: json.suberror,
            })
        );
    }
    if (json.error || json.error_description || json.suberror) {
        const err = classify(json.error ?? "", json.error_description, json.suberror);
        if (err instanceof ServerError) {
            // real MSAL formats token-endpoint server errors this way
            const N = "Not Available";
            err.errorMessage = err.message = `Error(s): ${json.error_codes || N} - Timestamp: ${json.timestamp || N} - Description: ${json.error_description || N} - Correlation ID: ${json.correlation_id || N} - Trace ID: ${json.trace_id || N}`;
        }
        throw err;
    }
    return json;
}

// ---- the client -----------------------------------------------------------

/**
 * Cache entities and keys match @azure/msal-browser 5.x ("msal.3" schema)
 * exactly, so mini-msal and real MSAL read each other's sign-in state —
 * drop-in cache interop, verified by the cross-stack E2E.
 */
const P = "msal.3";

export interface TokenEntity {
    credentialType: string;
    homeAccountId: string;
    environment: string;
    clientId: string;
    secret: string;
    realm?: string;
    target?: string;
    cachedAt?: string;
    expiresOn?: string;
    extendedExpiresOn?: string;
    refreshOn?: string;
    tokenType?: string;
    familyId?: string;
    lastUpdatedAt?: string;
}

export interface AccountEntity {
    homeAccountId: string;
    environment: string;
    realm: string;
    localAccountId: string;
    username: string;
    authorityType: string;
    name?: string;
    clientInfo?: string;
    tenantProfiles?: unknown[];
    nativeAccountId?: string;
    lastUpdatedAt?: string;
    cachedByApiId?: number;
    cloudGraphHostName?: string;
    msGraphHost?: string;
}

interface TokenKeys {
    idToken: string[];
    accessToken: string[];
    refreshToken: string[];
}

export interface AuthCodeResponse {
    code: string;
    /** PKCE verifier; absent on hybrid-spa redemptions (ApiId 866) */
    verifier?: string;
    scopes: string[];
    redirectUri?: string;
    correlationId?: string;
    /** authorize-request nonce, validated against the id_token's claim */
    nonce?: string;
    /** CCS routing hint ("Oid:<oid>@<tid>" or "UPN:<hint>") */
    ccs?: string;
    /** real MSAL ApiId of the calling flow, cached on the account entity */
    apiId?: number;
    /** custom request state, echoed on result.state */
    userState?: string;
    /** raw request claims JSON (merged with defaults at send time) */
    claims?: string;
    /** request extraQueryParameters (ride the token-endpoint query too) */
    eqp?: Record<string, string>;
    /** per-request authority override */
    authority?: string;
}

/**
 * Storage seam, mirroring real's browserStorage split: plaintext get/set/
 * remove for key indexes + metadata (msal.version, throttle entries), and
 * getUser/setUser for cache entities — ./local-storage swaps in an
 * encrypted-at-rest implementation. Temp state (interaction lock,
 * msal.request) always stays in sessionStorage, like real's
 * temporaryCacheStorage.
 */
export interface Store {
    /** async setup (key import, cache decrypt), awaited by initialize() */
    init?(): Promise<void>;
    get(key: string): string | null;
    set(key: string, value: string): void;
    remove(key: string): void;
    getUser(key: string): string | null;
    /** kmsi entities are persisted plaintext by ./local-storage, like real */
    setUser(key: string, value: string, kmsi?: boolean): void | Promise<void>;
}

/** The core client surface returned by createClient. */
export interface AuthClient {
    initialize(): Promise<void>;
    addEventCallback(cb: EventCallback): string | null;
    removeEventCallback(id: string): void;
    getAllAccounts(filter?: AccountFilter): AccountInfo[];
    getAccount(filter: AccountFilter): AccountInfo | null;
    /** resolved config incl. real's observable defaults */
    getConfiguration(): any;
    getActiveAccount(): AccountInfo | null;
    setActiveAccount(account: AccountInfo | null): void;
    getLogger(): Logger;
    setLogger(logger: Logger): void;
    /** store wrapper-library (react/angular) SKU + version */
    initializeWrapperLibrary(sku: string, version: string): void;
    /** swap the navigation implementation used by redirect flows */
    setNavigationClient(navigationClient: NavigationClient): void;
    loginRedirect(req: TokenRequest): Promise<void>;
    acquireTokenRedirect(req: TokenRequest): Promise<void>;
    handleRedirectPromise(): Promise<AuthenticationResult | null>;
    ssoSilent(req: TokenRequest): Promise<AuthenticationResult>;
    acquireTokenSilent(req: TokenRequest): Promise<AuthenticationResult>;
    logoutRedirect(req?: { account?: AccountInfo | null }): Promise<void>;
    /** local sign-out: purge the cache, no navigation, no end_session */
    clearCache(req?: { account?: AccountInfo | null }): Promise<void>;
    /** seed the cache from an externally-acquired AuthenticationResult
     * (hybrid SSR apps); writes account + id/access token entities */
    hydrateCache(
        result: AuthenticationResult,
        request?: { correlationId?: string }
    ): Promise<void>;
}

/**
 * Internal seam handed to feature modules (popup, broker, naa, …). A feature
 * is a plain function that attaches methods to ctx.client using the internal
 * plumbing exposed here.
 */
export interface ClientContext {
    config: Config;
    client: AuthClient & Record<string, any>;
    emit(
        eventType: string,
        interactionType?: string,
        payload?: unknown,
        error?: unknown
    ): void;
    preflight(): void;
    /** take the interaction lock; throws interaction_in_progress if held */
    lock(type?: string): void;
    unlock(): void;
    authorizeUrl(
        req: TokenRequest,
        interactionType?: string
    ): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
        correlationId: string;
        nonce: string;
        ccs?: string;
    }>;
    /** await the response the redirect-bridge page broadcasts for this state */
    waitForCode(state: string, timeoutMs: number): Promise<string>;
    redeem(res: AuthCodeResponse): Promise<AuthenticationResult>;
    clearAccount(account?: AccountInfo | null): void;
    /** replace the cache backend (./local-storage); call before initialize */
    setStore(store: Store): void;
    /** the live cache backend (resolved at call time, after setStore) */
    getStore(): Store;
    /** run a hook during initialize(), after the store is ready
     * (./cache-migration) */
    onInit(hook: () => Promise<void> | void): void;
    /** cache an account entity + index it (routed through the Store seam) */
    writeAccount(entity: AccountEntity, kmsi?: boolean): Promise<void>;
    /** read a cached credential entity from a token-key index (./naa) */
    findToken(
        type: "idToken" | "accessToken" | "refreshToken",
        match: (t: TokenEntity) => boolean
    ): TokenEntity | undefined;
    /** cache token entities + index them (./naa hydration,
     * loadExternalTokens); writes only the credentials present */
    writeTokens(t: {
        homeAccountId: string;
        environment: string;
        realm: string;
        idToken?: string;
        accessToken?: string;
        target?: string;
        /** epoch seconds */
        expiresOn?: number;
        /** epoch seconds; defaults to expiresOn */
        extendedExpiresOn?: number;
        refreshToken?: string;
        /** family id (foci) for the refresh token entity */
        foci?: string;
        /** KMSI entities persist plaintext via ./local-storage */
        kmsi?: boolean;
    }): Promise<void>;
    /** ./broker's silent interception: a promise routes the request to the
     * platform broker; undefined runs the web silent ladder */
    nativeSilent?: (
        req: TokenRequest,
        account: AccountInfo
    ) => Promise<AuthenticationResult> | undefined;
    logoutUrl(
        req?: { postLogoutRedirectUri?: string; correlationId?: string },
        interactionType?: string
    ): string;
    /** main-window navigation through the NavigationClient seam
     * (./popup's logoutPopup mainWindowRedirectUri) */
    navigate(url: string, apiId: number): Promise<boolean | void>;
}

export type Feature = (ctx: ClientContext) => void;

export function createClient(
    config: Config,
    features: Feature[] = []
): AuthClient {
    const clientId = config.auth.clientId;
    const authority = (
        config.auth.authority ?? "https://login.microsoftonline.com/common"
    ).replace(/\/$/, "");
    const redirectUri = new URL(config.auth.redirectUri ?? "/", location.href)
        .href;
    let metadata:
        | {
              authorization_endpoint: string;
              token_endpoint: string;
              end_session_endpoint: string;
          }
        | undefined;
    const listeners = new Map<string, EventCallback>();
    let nextListenerId = 0;
    let redirectResult: Promise<AuthenticationResult | null> | null = null;
    let initialized = false;
    const inFlight = new Map<string, Promise<AuthenticationResult>>();

    // cache backend; ./local-storage swaps in the encrypted implementation
    let store: Store = {
        get: (k) => sessionStorage.getItem(k),
        set: (k, v) => sessionStorage.setItem(k, v),
        remove: (k) => sessionStorage.removeItem(k),
        getUser: (k) => sessionStorage.getItem(k),
        setUser: (k, v) => sessionStorage.setItem(k, v),
    };

    // feature hooks run during initialize(), after the store is ready
    const initHooks: (() => Promise<void> | void)[] = [];

    const uninitialized = () =>
        new BrowserAuthError("uninitialized_public_client_application");

    // ---- logger (real's Logger gate: default volume Info(2)) ----
    let logger = new Logger(
        config.system?.loggerOptions,
        LIB_NAME,
        WIRE_ID["x-client-VER"]
    );
    const log = (level: number, msg: string) =>
        logger.logMessage(msg, { logLevel: level });
    // wrapper SKU/version (initializeWrapperLibrary) — real forwards these
    // to server telemetry; mini sends stub-empty telemetry headers (see
    // TOKEN_TELEMETRY), so the metadata is only stored (C19 revisits)
    let wrapperMeta: [sku: string, version: string] | null = null;

    // ---- redirect navigation seam (real's NavigationClient plumbing) ----
    let navClient = config.system?.navigationClient ?? new NavigationClient();
    const navOptions = (apiId: number, noHistory = false) => ({
        apiId,
        timeout: config.system?.redirectNavigationTimeout ?? 30_000,
        noHistory,
    });
    // temp keys real uses for the return-to-start-page replay
    const originKey = `msal.${clientId}.request.origin`;
    const hashKey = `msal.${clientId}.urlHash`;

    // ---- interaction lock (same storage entry as real MSAL) ----
    const lockKey = "msal.interaction.status";
    const lock = (type = "signin") => {
        if (sessionStorage.getItem(lockKey)) {
            throw new BrowserAuthError("interaction_in_progress");
        }
        sessionStorage.setItem(lockKey, JSON.stringify({ clientId, type }));
    };
    const unlock = () => sessionStorage.removeItem(lockKey);

    // ---- events (EventMessage-shaped like real's EventHandler.emitEvent) ----
    // real's EventHandler posts login/logout/activeAccountChanged to other
    // tabs/instances unconditionally; receiving is subscribed at initialize,
    // only in localStorage mode (StandardController)
    const eventBus = new BroadcastChannel("msal.broadcast.event");
    const emit = (
        eventType: string,
        interactionType?: string,
        payload?: unknown,
        error?: unknown
    ) => {
        log(3, `emitting event ${eventType}`);
        const m: EventMessage = {
            eventType,
            interactionType: interactionType ?? null,
            payload: payload ?? null,
            error: error ?? null,
            timestamp: Date.now(),
        };
        if (
            eventType === EventType.LOGIN_SUCCESS ||
            eventType === EventType.LOGOUT_SUCCESS ||
            eventType === EventType.ACTIVE_ACCOUNT_CHANGED
        ) {
            eventBus.postMessage(m);
        }
        listeners.forEach((l) => l(m));
    };

    // ---- cache (real-MSAL v5 schema) ----
    const env = () => preferredEnv(authority);

    const readJSON = <T,>(key: string): T | null => {
        const raw = store.get(key);
        return raw ? (JSON.parse(raw) as T) : null;
    };

    const writeJSON = (key: string, value: unknown) => {
        store.set(key, JSON.stringify(value));
    };

    // entities go through the user-data path (encrypted at rest by
    // ./local-storage); indexes/metadata stay plaintext like real
    const readUser = <T,>(key: string): T | null => {
        const raw = store.getUser(key);
        return raw ? (JSON.parse(raw) as T) : null;
    };

    const writeUser = (key: string, value: unknown, kmsi?: boolean) =>
        store.setUser(key, JSON.stringify(value), kmsi);

    const tokenKeysKey = `${P}.token.keys.${clientId}`;

    const tokenKeys = (): TokenKeys =>
        readJSON<TokenKeys>(tokenKeysKey) ?? {
            idToken: [],
            accessToken: [],
            refreshToken: [],
        };

    const accountKeys = (): string[] =>
        readJSON<string[]>(`${P}.account.keys`) ?? [];

    const findCred = (
        list: string[],
        match: (t: TokenEntity) => boolean
    ): TokenEntity | undefined => {
        for (const k of list) {
            const t = readUser<TokenEntity>(k);
            if (t && t.clientId === clientId && match(t)) {
                return t;
            }
        }
        return undefined;
    };

    // real expands each tenantProfile into its own AccountInfo, sourcing
    // per-tenant fields from that tenant's cached id token claims
    // (CacheManager.buildTenantProfiles + updateAccountTenantProfileData)
    const toAccountInfo = (e: AccountEntity, tenantId?: string): AccountInfo => {
        const profiles = (e.tenantProfiles ?? []) as Record<string, any>[];
        const tid = tenantId ?? e.realm;
        const profile = profiles.find((p) => p.tenantId === tid);
        const id = findCred(
            tokenKeys().idToken,
            (t) => t.homeAccountId === e.homeAccountId && t.realm === tid
        );
        const claims = id ? decodeJwt(id.secret) : undefined;
        return {
            authorityType: e.authorityType,
            dataBoundary: undefined,
            environment: e.environment,
            homeAccountId: e.homeAccountId,
            idToken: id?.secret,
            idTokenClaims: claims ?? {},
            kmsi: claims ? isKmsi(claims) : undefined,
            localAccountId: claims
                ? claims.oid ?? claims.sub ?? ""
                : profile?.localAccountId ?? e.localAccountId,
            loginHint: claims?.login_hint,
            name: claims ? claims.name : profile?.name ?? e.name,
            nativeAccountId: e.nativeAccountId,
            tenantId: claims ? claims.tid ?? claims.tfp ?? claims.acr ?? "" : tid,
            tenantProfiles: new Map(profiles.map((p) => [p.tenantId, p])),
            upn: claims?.upn,
            username: claims
                ? claims.preferred_username ?? claims.upn ?? ""
                : profile?.username ?? e.username,
        };
    };

    const matchesFilter = (a: AccountInfo, f: AccountFilter): boolean =>
        (!f.homeAccountId || a.homeAccountId === f.homeAccountId) &&
        (!f.localAccountId || a.localAccountId === f.localAccountId) &&
        (!f.username ||
            a.username.toLowerCase() === f.username.toLowerCase());

    const getAllAccounts = (filter?: AccountFilter): AccountInfo[] =>
        accountKeys()
            .map((k) => readUser<AccountEntity>(k))
            .filter((e): e is AccountEntity => !!e)
            .flatMap((e) =>
                ((e.tenantProfiles as { tenantId: string }[]) ?? [
                    { tenantId: e.realm },
                ]).map((p) => toAccountInfo(e, p.tenantId))
            )
            .filter((a) => !filter || matchesFilter(a, filter));

    // Real returns null on an empty/all-empty filter (CacheManager
    // .getAccountInfoFilteredBy) rather than the first account.
    const getAccount = (filter: AccountFilter): AccountInfo | null =>
        filter && Object.values(filter).some((v) => v)
            ? getAllAccounts(filter)[0] ?? null
            : null;

    const activeKey = `msal.${clientId}.active-account-filters`;

    const getActiveAccount = (): AccountInfo | null => {
        const f = readJSON<{ homeAccountId: string }>(activeKey);
        return f ? getAccount({ homeAccountId: f.homeAccountId }) : null;
    };

    /**
     * Real's buildAccountToCache: reuse the tenant-agnostic base account
     * entity (homeAccountId + environment) and append the incoming tenant
     * profiles to it — guest-tenant tokens never create a second entity.
     * Returns the (merged) entity actually written.
     */
    const mergeAccount = async (
        e: AccountEntity,
        kmsi?: boolean
    ): Promise<AccountEntity> => {
        const ks = accountKeys();
        const bases = ks
            .map((k) => readUser<AccountEntity>(k))
            .filter(
                (b): b is AccountEntity =>
                    !!b &&
                    b.homeAccountId === e.homeAccountId &&
                    b.environment === e.environment
            );
        // like real, >1 base match means a corrupt cache: ignore the hit
        const target = bases.length === 1 ? bases[0] : e;
        if (target !== e) {
            const profiles = (target.tenantProfiles ??= []) as {
                tenantId: string;
            }[];
            for (const p of (e.tenantProfiles ?? []) as { tenantId: string }[]) {
                if (!profiles.some((q) => q.tenantId === p.tenantId)) {
                    profiles.push(p);
                }
            }
            target.lastUpdatedAt = e.lastUpdatedAt;
            target.cachedByApiId = e.cachedByApiId;
        }
        const key = `${P}|${target.homeAccountId}|${target.environment}|${target.realm}`;
        await writeUser(key, target, kmsi);
        if (!ks.includes(key)) {
            writeJSON(`${P}.account.keys`, [...ks, key]);
        }
        return target;
    };

    /**
     * Real's saveAccessToken scope dedupe: remove every cached AT for the
     * same clientId/account/realm/tokenType whose (non-OIDC) scope set
     * intersects the new token's. Returns the pruned key list.
     */
    const dedupeATs = (
        list: string[],
        newKey: string,
        t: {
            homeAccountId: string;
            environment: string;
            realm: string;
            target: string;
        }
    ): string[] => {
        const scopes = t.target.toLowerCase().split(" ").filter(Boolean);
        const nonOidc = scopes.filter((s) => !OIDC_SCOPES.includes(s));
        const cmp = nonOidc.length ? nonOidc : scopes;
        return list.filter((k) => {
            if (k === newKey) {
                return true;
            }
            const c = readUser<TokenEntity>(k);
            if (
                c &&
                c.clientId === clientId &&
                c.homeAccountId === t.homeAccountId &&
                c.environment === t.environment &&
                c.realm === t.realm &&
                (c.tokenType ?? "Bearer") === "Bearer" &&
                (c.target ?? "")
                    .toLowerCase()
                    .split(" ")
                    .some((s) => cmp.includes(s))
            ) {
                store.remove(k);
                return false;
            }
            return true;
        });
    };

    /** write id/access/refresh token entities + index them (ctx.writeTokens
     * seam, hydrateCache, loadExternalTokens); writes what's present */
    const writeTokenEntities = async (t: {
        homeAccountId: string;
        environment: string;
        realm: string;
        idToken?: string;
        accessToken?: string;
        target?: string;
        expiresOn?: number;
        extendedExpiresOn?: number;
        refreshToken?: string;
        foci?: string;
        kmsi?: boolean;
    }): Promise<void> => {
        const ts = String(Date.now());
        const base = `${P}|${t.homeAccountId}|${t.environment}`;
        const shared = {
            homeAccountId: t.homeAccountId,
            environment: t.environment,
            clientId,
            lastUpdatedAt: ts,
        };
        const keys = tokenKeys();
        const add = (list: string[], k: string) =>
            list.includes(k) ? list : [...list, k];
        if (t.idToken) {
            const idKey = `${base}|idtoken|${clientId}|${t.realm}||`;
            await writeUser(
                idKey,
                {
                    ...shared,
                    credentialType: "IdToken",
                    secret: t.idToken,
                    realm: t.realm,
                } satisfies TokenEntity,
                t.kmsi
            );
            keys.idToken = add(keys.idToken, idKey);
        }
        if (t.accessToken) {
            const target = t.target ?? "";
            const atKey = `${base}|accesstoken|${clientId}|${t.realm}|${target.toLowerCase()}|`;
            await writeUser(
                atKey,
                {
                    ...shared,
                    credentialType: "AccessToken",
                    secret: t.accessToken,
                    realm: t.realm,
                    target,
                    cachedAt: String(Math.floor(Date.now() / 1000)),
                    expiresOn: String(t.expiresOn ?? 0),
                    extendedExpiresOn: String(
                        t.extendedExpiresOn ?? t.expiresOn ?? 0
                    ),
                    tokenType: "Bearer",
                } satisfies TokenEntity,
                t.kmsi
            );
            keys.accessToken = add(
                dedupeATs(keys.accessToken, atKey, {
                    homeAccountId: t.homeAccountId,
                    environment: t.environment,
                    realm: t.realm,
                    target,
                }),
                atKey
            );
        }
        if (t.refreshToken) {
            const rtKey = `${base}|refreshtoken|${clientId}|||`;
            await writeUser(
                rtKey,
                {
                    ...shared,
                    credentialType: "RefreshToken",
                    secret: t.refreshToken,
                    ...(t.foci && { familyId: t.foci }),
                } satisfies TokenEntity,
                t.kmsi
            );
            keys.refreshToken = add(keys.refreshToken, rtKey);
        }
        writeJSON(tokenKeysKey, keys);
    };

    /**
     * Same environment guard as real MSAL 5.16: when this app is re-booted
     * inside one of our own hidden iframes (auth response in the hash), auth
     * APIs refuse to run — the opener's poller owns the response. (Real 5.16
     * no longer blocks calls inside msal-named popup windows.)
     */
    const preflight = () => {
        if (!initialized) {
            throw uninitialized();
        }
        if (
            window !== window.parent &&
            /[#&](code|error)=/.test(location.hash)
        ) {
            throw new BrowserAuthError("block_iframe_reload");
        }
    };

    // like real MSAL's addScopes: request order, exact-case dedupe, OIDC
    // defaults appended once
    const normScopes = (scopes: string[]): string =>
        [...new Set([...scopes, "openid", "profile", "offline_access"])].join(
            " "
        );

    // real's buildMergedClaims: request claims + default id_token claims
    // (signin_state/login_hint) + xms_cc from auth.clientCapabilities, sent
    // on authorize AND token requests
    const mergedClaims = (claims?: string): string => {
        const c = claims ? JSON.parse(claims) : {};
        c.id_token = {
            signin_state: { essential: false },
            login_hint: { essential: false },
            ...c.id_token,
        };
        const caps = config.auth.clientCapabilities;
        if (caps?.length) {
            (c.access_token ??= {}).xms_cc = { values: caps };
        }
        return JSON.stringify(c);
    };

    /** canonical per-request authority (no trailing slash) */
    const authorityFor = (req?: { authority?: string }): string =>
        (req?.authority ?? authority).replace(/\/$/, "");

    // CCS routing hint, like real's ccsCredential: account wins over hint
    const ccsFrom = (req: TokenRequest): string | undefined =>
        req.account
            ? `Oid:${req.account.localAccountId}@${req.account.tenantId}`
            : req.loginHint
              ? `UPN:${req.loginHint}`
              : undefined;

    // ---- authorize-request plumbing ----
    const authorizeUrl = async (
        req: TokenRequest,
        interactionType = "redirect"
    ): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
        correlationId: string;
        nonce: string;
        ccs?: string;
    }> => {
        const { verifier, challenge } = await pkce();
        // real's wire state: base64 lib state {id, meta}, "|<custom>" appended
        // when the request carries one (result.state echoes only the custom
        // part). The redirect-bridge page broadcasts on the lib-state id.
        const state =
            btoa(
                JSON.stringify({
                    id: crypto.randomUUID(),
                    meta: { interactionType },
                })
            ) + (req.state ? `|${req.state}` : "");
        log(2, "building authorize url");
        const correlationId = req.correlationId ?? crypto.randomUUID();
        const nonce = crypto.randomUUID();
        // real's hint ladder: sid only on prompt=none (and it suppresses
        // login_hint); no account hints at all with prompt=select_account
        const sid = req.prompt === "none" ? req.sid : undefined;
        const skipHints = !!sid || req.prompt === "select_account";
        const hint = skipHints
            ? undefined
            : req.loginHint ?? req.account?.username;
        const ccs = skipHints ? undefined : ccsFrom(req);
        const reqRedirectUri = req.redirectUri
            ? new URL(req.redirectUri, location.href).href
            : redirectUri;
        const url = new URL(
            metadata!.authorization_endpoint.replace(
                authority,
                authorityFor(req)
            )
        );
        const p = url.searchParams;
        p.set("client_id", clientId);
        p.set("response_type", "code");
        p.set("redirect_uri", reqRedirectUri);
        p.set("scope", normScopes(req.scopes));
        p.set("state", state);
        p.set("nonce", nonce);
        p.set("code_challenge", challenge);
        p.set("code_challenge_method", "S256");
        p.set("response_mode", "fragment");
        p.set("client_info", "1");
        p.set("client-request-id", correlationId);
        p.set("claims", mergedClaims(req.claims));
        p.set("clidata", "1");
        p.set("x-client-SKU", WIRE_ID["x-client-SKU"]);
        p.set("x-client-VER", WIRE_ID["x-client-VER"]);
        if (sid) p.set("sid", sid);
        if (hint) p.set("login_hint", hint);
        if (req.domainHint) p.set("domain_hint", req.domainHint);
        if (ccs) p.set("X-AnchorMailbox", ccs);
        if (req.prompt) p.set("prompt", req.prompt);
        // like real's addExtraParameters: never overrides standard params
        for (const [k, v] of Object.entries(req.extraQueryParameters ?? {})) {
            if (!p.has(k) && v) p.set(k, v);
        }
        return {
            url: url.href,
            verifier,
            state,
            redirectUri: reqRedirectUri,
            correlationId,
            nonce,
            ccs,
        };
    };

    /**
     * Await the auth response the redirect-bridge page broadcasts on the
     * lib-state id, like real v5's waitForBridgeResponse: no URL polling and
     * no popup-close detection — a redirect page without the bridge (or a
     * closed popup) simply times out.
     */
    const waitForCode = (state: string, timeoutMs: number): Promise<string> =>
        new Promise((resolve, reject) => {
            log(3, "waiting for bridge response");
            const { id } = JSON.parse(atob(state.split("|")[0]));
            const channel = new BroadcastChannel(id);
            const settle = (fn: () => void) => {
                clearTimeout(timer);
                channel.close();
                fn();
            };
            const timer = setTimeout(
                () =>
                    settle(() =>
                        reject(
                            new BrowserAuthError(
                                "timed_out",
                                "redirect_bridge_timeout"
                            )
                        )
                    ),
                timeoutMs
            );
            channel.onmessage = (ev) =>
                settle(() => {
                    const params = new URLSearchParams(ev.data.payload);
                    const err = params.get("error");
                    const code = params.get("code");
                    if (err) {
                        reject(
                            classify(
                                err,
                                params.get("error_description") ?? ""
                            )
                        );
                    } else if (params.get("state") !== state) {
                        reject(new ClientAuthError("state_mismatch"));
                    } else if (code) {
                        resolve(code);
                    } else {
                        reject(
                            new BrowserAuthError(
                                "hash_does_not_contain_known_properties"
                            )
                        );
                    }
                });
        });

    // ---- token redemption ----
    const tokenRequest = async (
        scopes: string[],
        grant: Record<string, string | undefined>,
        // state present (custom or "") only on interactive/ssoSilent results,
        // like real; correlationId generated per request when not provided
        meta?: {
            correlationId?: string;
            state?: string;
            nonce?: string;
            ccs?: string;
            apiId?: number;
            claims?: string;
            eqp?: Record<string, string>;
            authority?: string;
            homeAccountId?: string;
        }
    ): Promise<AuthenticationResult> => {
        const correlationId = meta?.correlationId ?? crypto.randomUUID();
        const reqAuthority = authorityFor(meta);
        // throttle key mirrors real's RequestThumbprint (undefined fields
        // dropped by JSON.stringify)
        const throttleKey = `throttling.${JSON.stringify({
            clientId,
            authority: `${reqAuthority}/`,
            scopes,
            homeAccountIdentifier: meta?.homeAccountId,
            claims: meta?.claims,
            authenticationScheme: "Bearer",
        })}`;
        // extraQueryParameters + client-request-id ride the token endpoint
        // QUERY string (real 5.16's createTokenQueryParameters)
        const q = new URLSearchParams(meta?.eqp);
        q.set("client-request-id", correlationId);
        log(2, "sending token request");
        const json = await post(
            `${metadata!.token_endpoint.replace(authority, reqAuthority)}?${q}`,
            {
                redirect_uri: redirectUri,
                ...grant,
                client_id: clientId,
                scope: normScopes(scopes),
                client_info: "1",
                claims: mergedClaims(meta?.claims),
                ...(meta?.ccs && { "X-AnchorMailbox": meta.ccs }),
                ...TOKEN_TELEMETRY,
            },
            store,
            throttleKey
        );
        log(3, "token response received");
        const claims = decodeJwt(json.id_token);
        if (meta?.nonce && claims.nonce !== meta.nonce) {
            throw new ClientAuthError("nonce_mismatch");
        }
        // like real MSAL, homeAccountId comes from client_info when present
        let uid = claims.oid ?? claims.sub;
        let utid = claims.tid ?? "";
        if (json.client_info) {
            try {
                const ci = JSON.parse(
                    atob(
                        json.client_info.replace(/-/g, "+").replace(/_/g, "/")
                    )
                );
                uid = ci.uid;
                utid = ci.utid;
            } catch {
                /* fall back to claims */
            }
        }
        const homeAccountId = `${uid}.${utid}`;
        const localAccountId = claims.oid ?? claims.sub;
        const username = claims.preferred_username ?? claims.email ?? "";
        const grantedStr: string = json.scope ?? scopes.join(" ");
        const now = Math.floor(Date.now() / 1000);
        // real stamps every entity write with lastUpdatedAt (epoch ms string)
        const ts = String(Date.now());
        const expiresOn = now + json.expires_in;
        // refresh_in → refreshOn on the AT entity + result (proactive refresh)
        const refreshOn = json.refresh_in
            ? now + Number(json.refresh_in)
            : undefined;
        const environment = env();
        const realm = claims.tid ?? "";
        const base = `${P}|${homeAccountId}|${environment}`;

        const idKey = `${base}|idtoken|${clientId}|${realm}||`;
        const atKey = `${base}|accesstoken|${clientId}|${realm}|${grantedStr.toLowerCase()}|`;
        const rtKey = `${base}|refreshtoken|${clientId}|||`;
        const kmsi = isKmsi(claims);

        const entity = await mergeAccount(
            {
                homeAccountId,
                environment,
                realm,
                localAccountId,
                username,
                authorityType: "MSSTS",
                name: claims.name,
                clientInfo: json.client_info,
                tenantProfiles: [
                    {
                        tenantId: realm,
                        localAccountId,
                        name: claims.name,
                        username,
                        isHomeTenant: realm === homeAccountId.split(".")[1],
                    },
                ],
                lastUpdatedAt: ts,
                cachedByApiId: meta?.apiId,
            },
            kmsi
        );
        await writeUser(idKey, {
            credentialType: "IdToken",
            homeAccountId,
            environment,
            clientId,
            secret: json.id_token,
            realm,
            lastUpdatedAt: ts,
        } satisfies TokenEntity, kmsi);
        await writeUser(atKey, {
            credentialType: "AccessToken",
            homeAccountId,
            environment,
            clientId,
            secret: json.access_token,
            realm,
            target: grantedStr,
            cachedAt: String(now),
            expiresOn: String(expiresOn),
            extendedExpiresOn: String(expiresOn),
            ...(refreshOn ? { refreshOn: String(refreshOn) } : undefined),
            tokenType: "Bearer",
            lastUpdatedAt: ts,
        } satisfies TokenEntity, kmsi);
        if (json.refresh_token) {
            await writeUser(rtKey, {
                credentialType: "RefreshToken",
                homeAccountId,
                environment,
                clientId,
                secret: json.refresh_token,
                lastUpdatedAt: ts,
            } satisfies TokenEntity, kmsi);
        }

        const keys = tokenKeys();
        // real's saveAccessToken drops cached ATs (same account/realm/type)
        // whose scopes intersect the new token's before writing it
        keys.accessToken = dedupeATs(keys.accessToken, atKey, {
            homeAccountId,
            environment,
            realm,
            target: grantedStr,
        });
        const add = (list: string[], k: string) =>
            list.includes(k) ? list : [...list, k];
        writeJSON(tokenKeysKey, {
            idToken: add(keys.idToken, idKey),
            accessToken: add(keys.accessToken, atKey),
            refreshToken: json.refresh_token
                ? add(keys.refreshToken, rtKey)
                : keys.refreshToken,
        });
        // real emits no same-tab accountAdded event; cross-tab propagation
        // is the localStorage/BroadcastChannel feature's job
        return {
            authority: `${reqAuthority}/`,
            uniqueId: localAccountId,
            tenantId: realm,
            scopes: [...new Set(grantedStr.split(" "))],
            account: toAccountInfo(entity, realm),
            idToken: json.id_token,
            idTokenClaims: claims,
            accessToken: json.access_token,
            fromCache: false,
            expiresOn: new Date(expiresOn * 1000),
            extExpiresOn: new Date(
                (now + (json.ext_expires_in ?? json.expires_in)) * 1000
            ),
            refreshOn: refreshOn ? new Date(refreshOn * 1000) : undefined,
            correlationId,
            requestId: "",
            familyId: json.foci ?? "",
            tokenType: "Bearer",
            state: meta?.state,
            cloudGraphHostName: "",
            msGraphHost: "",
            code: undefined,
            fromPlatformBroker: false,
        };
    };

    const redeem = (res: AuthCodeResponse): Promise<AuthenticationResult> =>
        tokenRequest(
            res.scopes,
            {
                grant_type: "authorization_code",
                code: res.code,
                // hybrid-spa codes (ApiId 866) come from a confidential
                // client: the redemption carries no PKCE verifier and no
                // redirect_uri (real's HybridSpaAuthorizationCodeClient)
                ...(res.apiId === 866
                    ? { redirect_uri: undefined }
                    : {
                          code_verifier: res.verifier,
                          ...(res.redirectUri && {
                              redirect_uri: res.redirectUri,
                          }),
                      }),
            },
            {
                correlationId: res.correlationId,
                state: res.userState ?? "",
                nonce: res.nonce,
                ccs: res.ccs,
                apiId: res.apiId,
                claims: res.claims,
                eqp: res.eqp,
                authority: res.authority,
            }
        );

    const redeemRefresh = (
        req: TokenRequest,
        refreshToken: string
    ): Promise<AuthenticationResult> =>
        tokenRequest(
            req.scopes,
            {
                grant_type: "refresh_token",
                refresh_token: refreshToken,
                // real redeems against the request's redirectUri
                ...(req.redirectUri && {
                    redirect_uri: new URL(req.redirectUri, location.href).href,
                }),
            },
            // 61 = real's ApiId.acquireTokenSilent_silentFlow
            {
                correlationId: req.correlationId,
                ccs: ccsFrom(req),
                apiId: 61,
                claims: req.claims,
                eqp: req.extraQueryParameters,
                authority: req.authority,
                homeAccountId: req.account?.homeAccountId,
            }
        );

    // ---- interactive: redirect ----
    const processRedirect = async (): Promise<AuthenticationResult | null> => {
        if (window !== window.parent) {
            // app re-loaded inside our own hidden iframe: leave the hash for
            // the opener's poller (same behavior as real MSAL's iframe guard)
            return null;
        }
        let params = new URLSearchParams(location.hash.slice(1));
        let cachedHash = false;
        if (!params.get("code") && !params.get("error")) {
            // no response in the URL: a replay navigation may have cached the
            // hash for this load (real's getRedirectResponse urlHash pickup)
            const cached = sessionStorage.getItem(hashKey);
            if (cached) {
                sessionStorage.removeItem(hashKey);
                params = new URLSearchParams(cached.slice(1));
                cachedHash = true;
            }
        }
        const err = params.get("error");
        const stored = sessionStorage.getItem("msal.request");
        // clean load: real resolves null silently, no handleRedirect events;
        // the interaction lock set before navigating away is released
        if ((!params.get("code") && !err) || !stored) {
            unlock();
            return null;
        }
        const {
            verifier,
            state,
            scopes,
            correlationId,
            nonce,
            ccs,
            userState,
            claims: reqClaims,
            eqp,
            authority: reqAuthority,
        } = JSON.parse(stored);
        const had = accountKeys().length;
        emit(EventType.HANDLE_REDIRECT_START, "redirect");
        const title = document.title;
        document.title = "Microsoft Authentication";
        try {
            if (params.get("state") !== state) {
                // forged/unknown state: real treats the response as not ours
                // and resolves null (no failure event, no throw)
                unlock();
                sessionStorage.removeItem("msal.request");
                return null;
            }
            const navBack = config.auth.navigateToLoginRequestUrl ?? true;
            if (!cachedHash) {
                const origin = sessionStorage.getItem(originKey) ?? "";
                if (navBack && normUrl(origin) !== normUrl(location.href)) {
                    // returned somewhere other than the initiating page:
                    // cache the response and navigate back; the next load's
                    // handleRedirectPromise processes it (real's replay)
                    sessionStorage.setItem(hashKey, location.hash);
                    history.replaceState(
                        null,
                        "",
                        location.pathname + location.search
                    );
                    let target = origin;
                    if (!target) {
                        target = location.origin + "/";
                        sessionStorage.setItem(originKey, target);
                    }
                    if (
                        (await navClient.navigateInternal(
                            target,
                            navOptions(865, true) // ApiId.handleRedirectPromise
                        )) !== false
                    ) {
                        // interaction lock intentionally stays held mid-replay
                        return null;
                    }
                    // custom client declined to navigate: process in place
                    sessionStorage.removeItem(hashKey);
                } else {
                    history.replaceState(
                        null,
                        "",
                        location.pathname + location.search
                    );
                    const i = origin.indexOf("#");
                    if (navBack && i > -1) {
                        // restore the app's own pre-login hash (replaceHash)
                        location.hash = origin.slice(i + 1);
                    }
                }
            }
            unlock();
            sessionStorage.removeItem("msal.request");
            sessionStorage.removeItem(originKey);
            if (err) {
                // login was cancelled/denied at the IdP
                throw classify(err, params.get("error_description") ?? "");
            }
            const result = await redeem({
                code: params.get("code")!,
                verifier,
                scopes,
                correlationId,
                nonce,
                ccs,
                apiId: 865, // ApiId.handleRedirectPromise
                userState,
                claims: reqClaims,
                eqp,
                authority: reqAuthority,
            });
            emit(EventType.ACQUIRE_TOKEN_SUCCESS, "redirect", result);
            if (had < accountKeys().length) {
                emit(EventType.LOGIN_SUCCESS, "redirect", result.account);
            }
            return result;
        } catch (e) {
            emit(EventType.ACQUIRE_TOKEN_FAILURE, "redirect", undefined, e);
            throw e;
        } finally {
            document.title = title;
            emit(EventType.HANDLE_REDIRECT_END, "redirect");
        }
    };

    // ---- silent ----
    /** hidden-iframe prompt=none flow, no events (shared by ssoSilent + ladder) */
    const silentFrame = async (
        req: TokenRequest,
        apiId: number
    ): Promise<AuthenticationResult> => {
        const {
            url,
            verifier,
            state,
            redirectUri: ru,
            correlationId,
            nonce,
            ccs,
        } = await authorizeUrl(
            { ...req, prompt: req.prompt ?? "none" },
            "silent"
        );
        const frame = document.createElement("iframe");
        frame.style.display = "none";
        document.body.append(frame);
        try {
            frame.src = url;
            const code = await waitForCode(
                state,
                config.system?.iframeBridgeTimeout ?? 10_000
            );
            return await redeem({
                code,
                verifier,
                scopes: req.scopes,
                redirectUri: ru,
                correlationId,
                nonce,
                ccs,
                apiId,
                userState: req.state,
                claims: req.claims,
                eqp: req.extraQueryParameters,
                authority: req.authority,
            });
        } finally {
            frame.remove();
        }
    };

    const ssoSilent = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        preflight();
        // real emits acquireToken* with the validated request (correlationId
        // added), and loginSuccess(account) when the account count grew
        const validRequest = { correlationId: crypto.randomUUID(), ...req };
        const had = accountKeys().length;
        emit(EventType.ACQUIRE_TOKEN_START, "silent", validRequest);
        try {
            const result = await silentFrame(validRequest, 863); // ApiId.ssoSilent
            emit(EventType.ACQUIRE_TOKEN_SUCCESS, "silent", result);
            if (had < accountKeys().length) {
                emit(EventType.LOGIN_SUCCESS, "silent", result.account);
            }
            return result;
        } catch (e) {
            emit(EventType.ACQUIRE_TOKEN_FAILURE, "silent", undefined, e);
            throw e;
        }
    };

    // CacheLookupPolicy gates each rung of the silent ladder
    // (AT -> RT -> iframe); forceRefresh bypasses the AT rung.
    // Same semantics as real MSAL's acquireTokenSilentAsync.
    const silentLadder = async (
        req: TokenRequest,
        account: AccountInfo
    ): Promise<AuthenticationResult> => {
        const pol = req.cacheLookupPolicy ?? CacheLookupPolicy.Default;
        const reqAuthority = authorityFor(req);
        let frameReason: string | undefined;
        const useAT =
            !req.forceRefresh &&
            pol <= CacheLookupPolicy.AccessTokenAndRefreshToken;
        const useRT =
            pol !== CacheLookupPolicy.AccessToken &&
            pol !== CacheLookupPolicy.Skip;
        const useFrame =
            pol === CacheLookupPolicy.Default ||
            pol >= CacheLookupPolicy.RefreshTokenAndNetwork;
        const keys = tokenKeys();
        const wanted = req.scopes.map((sc) => sc.toLowerCase());
        let at: TokenEntity | undefined;
        if (useAT) {
            const matches = keys.accessToken.filter((k) => {
                const t = readUser<TokenEntity>(k);
                return (
                    !!t &&
                    t.clientId === clientId &&
                    t.homeAccountId === account.homeAccountId &&
                    t.realm === account.tenantId &&
                    wanted.every((sc) =>
                        (t.target ?? "").toLowerCase().split(" ").includes(sc)
                    )
                );
            });
            if (matches.length > 1) {
                // real's getAccessToken: >1 match clears them ALL and
                // refreshes over the network
                matches.forEach((k) => store.remove(k));
                keys.accessToken = keys.accessToken.filter(
                    (k) => !matches.includes(k)
                );
                writeJSON(tokenKeysKey, keys);
            } else if (matches.length === 1) {
                const t = readUser<TokenEntity>(matches[0])!;
                if (Number(t.expiresOn) - 300 > Date.now() / 1000) {
                    at = t;
                }
            }
        }
        if (at) {
            const id = findCred(
                keys.idToken,
                (t) => t.homeAccountId === account.homeAccountId
            );
            return {
                authority: `${reqAuthority}/`,
                uniqueId: account.localAccountId,
                tenantId: account.tenantId,
                scopes: (at.target ?? "").split(" "),
                account,
                idToken: id?.secret ?? "",
                idTokenClaims: account.idTokenClaims,
                accessToken: at.secret,
                fromCache: true,
                expiresOn: new Date(Number(at.expiresOn) * 1000),
                extExpiresOn: new Date(Number(at.extendedExpiresOn) * 1000),
                refreshOn: at.refreshOn
                    ? new Date(Number(at.refreshOn) * 1000)
                    : undefined,
                correlationId: req.correlationId ?? crypto.randomUUID(),
                requestId: "",
                familyId: "",
                tokenType: "Bearer",
                state: undefined,
                cloudGraphHostName: "",
                msGraphHost: "",
                code: undefined,
                fromPlatformBroker: false,
            };
        }
        if (pol === CacheLookupPolicy.AccessToken) {
            // AT-only policy: expired/missing AT is a hard error
            throw new ClientAuthError("token_refresh_required");
        }
        // real emits this once, before the RT/iframe network leg, with the
        // fully-initialized silent request as payload
        emit(EventType.ACQUIRE_TOKEN_NETWORK_START, "silent", {
            account,
            authenticationScheme: "Bearer",
            authority: `${reqAuthority}/`,
            correlationId: req.correlationId,
            forceRefresh: !!req.forceRefresh,
            redirectUri: req.redirectUri,
            scopes: req.scopes,
        });
        if (useRT) {
            const rt = findCred(
                keys.refreshToken,
                (t) => t.homeAccountId === account.homeAccountId
            );
            if (rt) {
                try {
                    return await redeemRefresh({ ...req, account }, rt.secret);
                } catch (e) {
                    // real's checkIfRefreshTokenErrorCanBeResolvedSilently:
                    // iframe renewal only for invalid_grant/
                    // token_refresh_required errors that don't require
                    // interaction (bad_token subError excepted), or
                    // no_tokens_found/refresh_token_expired
                    const err = e as AuthError;
                    const resolvable =
                        (!(
                            e instanceof InteractionRequiredAuthError &&
                            err.subError !== "bad_token"
                        ) &&
                            (err.errorCode === "invalid_grant" ||
                                err.errorCode === "token_refresh_required")) ||
                        err.errorCode === "no_tokens_found" ||
                        err.errorCode === "refresh_token_expired";
                    if (!useFrame || !resolvable) {
                        throw e;
                    }
                    frameReason = err.errorCode;
                }
            } else if (!useFrame) {
                throw new InteractionRequiredAuthError("no_tokens_found");
            }
        }
        // last resort: hidden iframe with prompt=none
        let result: AuthenticationResult;
        try {
            result = await silentFrame(
                { ...req, account, loginHint: account.username },
                864 // ApiId.acquireTokenSilent_authCode
            );
        } catch (e) {
            // ./telemetry reads this off the error: real's perf events carry
            // the RT error that triggered the iframe fallback
            if (frameReason) {
                (e as any).silentRefreshReason = frameReason;
            }
            throw e;
        }
        // real's acquireTokenSilent results carry state: undefined (the key
        // exists — event payloads expose it — but interactive/ssoSilent
        // results are the only ones with a value)
        result.state = undefined;
        return result;
    };

    // ---- logout ----
    const logoutUrl = (
        req?: { postLogoutRedirectUri?: string; correlationId?: string },
        interactionType = "redirect"
    ): string => {
        const url = new URL(metadata!.end_session_endpoint);
        const p = url.searchParams;
        p.set(
            "post_logout_redirect_uri",
            new URL(
                req?.postLogoutRedirectUri ??
                    config.auth.postLogoutRedirectUri ??
                    redirectUri,
                location.href
            ).href
        );
        p.set("client-request-id", req?.correlationId ?? crypto.randomUUID());
        // real's redirect bridge requires a state param (lib-state format:
        // base64 of {id, meta:{interactionType}}); the IdP echoes it back
        p.set(
            "state",
            btoa(
                JSON.stringify({
                    id: crypto.randomUUID(),
                    meta: { interactionType },
                })
            )
        );
        return url.href;
    };

    const clearAccount = (account?: AccountInfo | null) => {
        const stays = (k: string) =>
            !!account && !k.includes(`|${account.homeAccountId}|`);
        const keys = tokenKeys();
        const kept: TokenKeys = {
            idToken: [],
            accessToken: [],
            refreshToken: [],
        };
        for (const type of ["idToken", "accessToken", "refreshToken"] as const) {
            for (const k of keys[type]) {
                stays(k) ? kept[type].push(k) : store.remove(k);
            }
        }
        const keptAccounts: string[] = [];
        for (const k of accountKeys()) {
            stays(k) ? keptAccounts.push(k) : store.remove(k);
        }
        if (account) {
            writeJSON(tokenKeysKey, kept);
            writeJSON(`${P}.account.keys`, keptAccounts);
            const f = readJSON<{ homeAccountId: string }>(activeKey);
            if (f?.homeAccountId === account.homeAccountId) {
                store.remove(activeKey);
            }
        } else {
            store.remove(tokenKeysKey);
            store.remove(`${P}.account.keys`);
            store.remove(activeKey);
        }
    };

    const client: AuthClient = {
        async initialize() {
            // second initialize is a silent no-op (no events), like real
            if (initialized) return;
            log(2, "initializing");
            emit(EventType.INITIALIZE_START);
            // encryption-key + cache import for ./local-storage (real's
            // browserStorage.initialize), no-op for the default store
            await store.init?.();
            // feature init hooks (./cache-migration), like real's
            // BrowserCacheManager.initialize ordering
            for (const h of initHooks) {
                await h();
            }
            if (config.cache?.cacheLocation === "localStorage") {
                // real subscribes to cross-tab events only in localStorage
                // mode (StandardController.initialize)
                eventBus.onmessage = (ev) =>
                    listeners.forEach((l) => l(ev.data));
            }
            // per-instance memory only: real MSAL re-fetches discovery for
            // every new instance and leaves no such key in storage
            metadata ??= await (
                await fetch(
                    `${authority}/v2.0/.well-known/openid-configuration`
                )
            ).json();
            // real tracks lib up/downgrades via this key (trackVersionChanges)
            store.set("msal.version", WIRE_ID["x-client-VER"]);
            initialized = true;
            emit(EventType.INITIALIZE_END);
        },

        handleRedirectPromise(): Promise<AuthenticationResult | null> {
            return initialized
                ? (redirectResult ??= processRedirect())
                : Promise.reject(uninitialized());
        },

        addEventCallback(cb: EventCallback): string | null {
            const id = String(nextListenerId++);
            listeners.set(id, cb);
            return id;
        },

        removeEventCallback(id: string): void {
            listeners.delete(id);
        },

        getAllAccounts,
        getAccount,
        getActiveAccount,

        getLogger: () => logger,
        setLogger(l: Logger) {
            logger = l;
        },
        initializeWrapperLibrary(sku: string, version: string) {
            wrapperMeta = [sku, version];
        },
        setNavigationClient(c: NavigationClient) {
            navClient = c;
        },

        // resolved config: user input over real's observable defaults (only
        // keys snapshots compare; unset optionals stay undefined like real)
        getConfiguration: () => ({
            auth: {
                clientId,
                authority:
                    config.auth.authority ??
                    "https://login.microsoftonline.com/common",
                cloudDiscoveryMetadata: "",
                ...config.auth,
            },
            cache: { cacheLocation: "sessionStorage", ...config.cache },
            system: {
                allowPlatformBroker: false,
                nativeBrokerHandshakeTimeout: 2000,
                redirectNavigationTimeout: 30_000,
                tokenRenewalOffsetSeconds: 300,
                popupBridgeTimeout: 60_000,
                iframeBridgeTimeout: 10_000,
                ...config.system,
            },
            experimental: {
                iframeTimeoutTelemetry: false,
                allowPlatformBrokerWithDOM: false,
                ...config.experimental,
            },
            telemetry: { ...config.telemetry },
        }),

        setActiveAccount(account: AccountInfo | null) {
            if (account) {
                writeJSON(activeKey, {
                    homeAccountId: account.homeAccountId,
                    localAccountId: account.localAccountId,
                    tenantId: account.tenantId,
                });
            } else {
                store.remove(activeKey);
            }
            // real emits this event with no payload (BrowserCacheManager)
            emit(EventType.ACTIVE_ACCOUNT_CHANGED);
        },

        async loginRedirect(req: TokenRequest): Promise<void> {
            return client.acquireTokenRedirect(req);
        },

        async acquireTokenRedirect(req: TokenRequest): Promise<void> {
            preflight();
            if (window !== window.parent) {
                // same guard as real MSAL: no full-page redirects from iframes
                throw new BrowserAuthError("redirect_in_iframe");
            }
            lock();
            const { url, verifier, state, correlationId, nonce, ccs } =
                await authorizeUrl(req);
            // persisted so the post-redirect redemption reuses the same
            // correlationId/nonce/CCS hint, like real's temp request cache
            sessionStorage.setItem(
                "msal.request",
                JSON.stringify({
                    verifier,
                    state,
                    scopes: req.scopes,
                    correlationId,
                    nonce,
                    ccs,
                    userState: req.state,
                    claims: req.claims,
                    eqp: req.extraQueryParameters,
                    authority: req.authority,
                })
            );
            // start page cached so handleRedirectPromise can navigate back
            // (navigateToLoginRequestUrl, default true)
            sessionStorage.setItem(
                originKey,
                new URL(req.redirectStartPage ?? location.href, location.href)
                    .href
            );
            // app hook: returning false cancels the navigation (the promise
            // resolves and, like real, the interaction lock stays held)
            if (config.auth.onRedirectNavigate?.(url) === false) return;
            await navClient.navigateExternal(
                url,
                navOptions(861) // ApiId.acquireTokenRedirect
            );
        },

        ssoSilent,

        async acquireTokenSilent(
            req: TokenRequest
        ): Promise<AuthenticationResult> {
            preflight();
            const account = req.account ?? getActiveAccount();
            if (!account) {
                throw new BrowserAuthError("no_account_error");
            }
            if (!getAccount({ homeAccountId: account.homeAccountId })) {
                // real rejects unknown accounts as an authority mismatch
                throw new ClientConfigurationError("authority_mismatch");
            }
            // identical concurrent calls share one in-flight promise (real's
            // acquireTokenSilentDeduped; thumbprint has no policy/forceRefresh)
            const key = JSON.stringify([
                req.scopes,
                account.homeAccountId,
                req.authority,
                req.claims,
            ]);
            let shared = inFlight.get(key);
            if (!shared) {
                // events fire once per deduped request, inside the shared
                // promise, like real's acquireTokenSilentAsync
                const validRequest = {
                    ...req,
                    correlationId: req.correlationId ?? crypto.randomUUID(),
                };
                emit(EventType.ACQUIRE_TOKEN_START, "silent", validRequest);
                shared = (
                    ctx.nativeSilent?.(validRequest, account) ??
                    silentLadder(validRequest, account)
                )
                    .then(
                        (r) => {
                            emit(EventType.ACQUIRE_TOKEN_SUCCESS, "silent", r);
                            return r;
                        },
                        (e) => {
                            emit(
                                EventType.ACQUIRE_TOKEN_FAILURE,
                                "silent",
                                undefined,
                                e
                            );
                            throw e;
                        }
                    )
                    .finally(() => inFlight.delete(key));
                inFlight.set(key, shared);
            }
            return shared;
        },

        async logoutRedirect(req?: {
            account?: AccountInfo | null;
            postLogoutRedirectUri?: string;
            correlationId?: string;
        }): Promise<void> {
            preflight();
            lock("signout");
            const validRequest = {
                correlationId: crypto.randomUUID(),
                postLogoutRedirectUri: config.auth.postLogoutRedirectUri,
                ...req,
            };
            // the page navigates away, so logoutStart is usually the only
            // event a same-page listener sees; real passes the RAW request
            // (null payload for logoutRedirect())
            emit(EventType.LOGOUT_START, "redirect", req);
            clearAccount(req?.account);
            const url = logoutUrl(validRequest);
            if (config.auth.onRedirectNavigate?.(url) === false) {
                // cancelled logout: real releases the interaction lock and a
                // same-page listener does see logoutEnd
                unlock();
                emit(EventType.LOGOUT_END, "redirect");
                return;
            }
            await navClient.navigateExternal(
                url,
                navOptions(961) // ApiId.logout
            );
        },

        // real's PCA.clearCache -> SilentCacheClient.logout ->
        // clearCacheOnLogout: purely local purge, no navigation, no
        // end_session request, no initialize requirement
        async clearCache(req?: {
            account?: AccountInfo | null;
        }): Promise<void> {
            const account = req?.account;
            const active = readJSON<{ homeAccountId: string }>(activeKey);
            if (
                active &&
                (!account || active.homeAccountId === account.homeAccountId)
            ) {
                // real's removeAccount: clearing the active account goes
                // through setActiveAccount(null) (emits activeAccountChanged)
                client.setActiveAccount(null);
            }
            clearAccount(account);
            if (!account) {
                // real's browserStorage.clear(): every remaining msal /
                // client-id key in the cache location + temp storage goes too
                const areas = [sessionStorage];
                if (config.cache?.cacheLocation === "localStorage") {
                    areas.push(localStorage);
                }
                for (const s of areas) {
                    for (const k of Object.keys(s)) {
                        if (k.includes("msal") || k.includes(clientId)) {
                            store.remove(k); // keeps ./local-storage's mirror in sync
                            s.removeItem(k);
                        }
                    }
                }
            }
        },

        // real's PCA.hydrateCache: account entity from result.account, then
        // id + access token entities (never a refresh token)
        async hydrateCache(result: AuthenticationResult): Promise<void> {
            const kmsi = isKmsi(result.idTokenClaims ?? {});
            await mergeAccount(
                entityFromAccountInfo(result.account, 963, {
                    // ApiId.hydrateCache
                    cloudGraphHostName: result.cloudGraphHostName,
                    msGraphHost: result.msGraphHost,
                }),
                kmsi
            );
            await writeTokenEntities({
                homeAccountId: result.account.homeAccountId,
                environment: result.account.environment,
                realm: result.tenantId,
                idToken: result.idToken,
                accessToken: result.accessToken,
                target: result.scopes.join(" "),
                expiresOn: result.expiresOn
                    ? Math.floor(+new Date(result.expiresOn) / 1000)
                    : 0,
                extendedExpiresOn: result.extExpiresOn
                    ? Math.floor(+new Date(result.extExpiresOn) / 1000)
                    : 0,
                kmsi,
            });
        },
    };

    const ctx: ClientContext = {
        config,
        client,
        emit,
        preflight,
        lock,
        unlock,
        authorizeUrl,
        waitForCode,
        redeem,
        clearAccount,
        navigate: (url, apiId) =>
            navClient.navigateInternal(url, navOptions(apiId)),
        setStore: (s) => (store = s),
        getStore: () => store,
        onInit: (h) => initHooks.push(h),
        writeAccount: async (e, kmsi) => {
            await mergeAccount(e, kmsi);
        },
        findToken: (type, match) => findCred(tokenKeys()[type], match),
        writeTokens: writeTokenEntities,
        logoutUrl,
    };
    for (const f of features) f(ctx);
    return client;
}

/**
 * Real's top-level loadExternalTokens export (the getTokenCache/ITokenCache
 * successor): write an externally-acquired token response (e.g. from
 * msal-node in a hybrid app, or a test harness) into the cache and return
 * the AuthenticationResult it represents. `features` lets callers compose
 * cache backends (compat passes ./local-storage + ./cache-migration).
 */
export async function loadExternalTokens(
    config: Config,
    request: {
        scopes?: string[];
        authority?: string;
        account?: AccountInfo;
        correlationId?: string;
        state?: string;
    },
    response: Record<string, any>,
    options: {
        clientInfo?: string;
        /** epoch seconds; defaults to now + response.expires_in */
        expiresOn?: number;
        extendedExpiresOn?: number;
    } = {},
    features: Feature[] = []
): Promise<AuthenticationResult> {
    let ctx!: ClientContext;
    const client = createClient(config, [
        ...(Array.isArray(features) ? features : []),
        (c) => (ctx = c),
    ]);
    // storage init + authority discovery, like real's standalone
    // BrowserCacheManager/Authority setup
    await client.initialize();
    const claims = response.id_token
        ? decodeJwt(response.id_token)
        : undefined;
    const kmsi = isKmsi(claims ?? {});
    const authority = (
        request.authority ??
        config.auth.authority ??
        "https://login.microsoftonline.com/common"
    ).replace(/\/$/, "");
    const environment = preferredEnv(authority);
    let entity: AccountEntity;
    if (request.account) {
        entity = entityFromAccountInfo(request.account, 964); // ApiId.loadExternalTokens
    } else {
        const clientInfo: string =
            options.clientInfo || response.client_info || "";
        if (!clientInfo && !claims) {
            throw new BrowserAuthError("unable_to_load_token");
        }
        let uid = claims?.oid ?? claims?.sub;
        let utid = claims?.tid ?? "";
        try {
            const ci = JSON.parse(
                atob(clientInfo.replace(/-/g, "+").replace(/_/g, "/"))
            );
            uid = ci.uid;
            utid = ci.utid;
        } catch {
            /* fall back to claims */
        }
        const homeAccountId = `${uid}.${utid}`;
        const realm = claims?.tid ?? "";
        const localAccountId = claims?.oid ?? claims?.sub ?? "";
        const username =
            claims?.preferred_username ?? claims?.email ?? "";
        entity = {
            homeAccountId,
            environment,
            realm,
            localAccountId,
            username,
            authorityType: "MSSTS",
            name: claims?.name,
            clientInfo: clientInfo || undefined,
            tenantProfiles: [
                {
                    tenantId: realm,
                    localAccountId,
                    name: claims?.name,
                    username,
                    isHomeTenant: realm === homeAccountId.split(".")[1],
                },
            ],
            lastUpdatedAt: String(Date.now()),
            cachedByApiId: 964, // ApiId.loadExternalTokens
        };
    }
    await ctx.writeAccount(entity, kmsi);
    const now = Math.floor(Date.now() / 1000);
    const scopeStr: string =
        response.scope ?? (request.scopes ?? []).join(" ");
    const hasAT = !!(response.access_token && response.expires_in && scopeStr);
    const expiresOn =
        options.expiresOn ?? now + Number(response.expires_in ?? 0);
    const extExpiresOn =
        options.extendedExpiresOn ??
        now + Number(response.ext_expires_in ?? response.expires_in ?? 0);
    await ctx.writeTokens({
        homeAccountId: entity.homeAccountId,
        environment: entity.environment,
        realm: entity.realm,
        idToken: response.id_token,
        accessToken: hasAT ? response.access_token : undefined,
        target: scopeStr,
        expiresOn,
        extendedExpiresOn: extExpiresOn,
        refreshToken: response.refresh_token,
        foci: response.foci,
        kmsi,
    });
    // real's generateAuthenticationResult (fromCache: true)
    return {
        authority: `${authority}/`,
        uniqueId: entity.localAccountId,
        tenantId: entity.realm,
        scopes: hasAT ? scopeStr.split(" ") : [],
        account: client.getAccount({
            homeAccountId: entity.homeAccountId,
        })!,
        idToken: response.id_token || "",
        idTokenClaims: claims ?? {},
        accessToken: hasAT ? response.access_token : "",
        fromCache: true,
        expiresOn: hasAT ? new Date(expiresOn * 1000) : null,
        extExpiresOn: hasAT ? new Date(extExpiresOn * 1000) : undefined,
        correlationId: request.correlationId || "",
        requestId: "",
        familyId: response.foci || "",
        tokenType: hasAT ? "Bearer" : "",
        state: request.state || "",
        cloudGraphHostName: entity.cloudGraphHostName || "",
        msGraphHost: entity.msGraphHost || "",
        fromPlatformBroker: false,
    } as unknown as AuthenticationResult;
}
