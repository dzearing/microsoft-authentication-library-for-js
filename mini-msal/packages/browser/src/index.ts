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
    tenantProfiles?: unknown[];
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
}

export interface Config {
    auth: {
        clientId: string;
        authority?: string;
        redirectUri?: string;
        postLogoutRedirectUri?: string;
        clientCapabilities?: string[];
    };
    system?: {
        popupBridgeTimeout?: number;
        iframeBridgeTimeout?: number;
    };
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

function decodeJwt(token: string): Record<string, any> {
    return JSON.parse(
        atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );
}

async function post(
    url: string,
    body: Record<string, string>,
    throttleKey?: string
) {
    // real's ThrottlingUtils: a cached 429/5xx/Retry-After token response
    // blocks identical requests (same thumbprint) until throttleTime — the
    // retry re-throws the stored error with NO network call
    const t =
        throttleKey &&
        JSON.parse(sessionStorage.getItem(throttleKey) ?? "null");
    if (t) {
        if (t.throttleTime >= Date.now()) {
            throw new ServerError(
                t.errorCodes?.join(" ") || "",
                t.errorMessage,
                t.subError
            );
        }
        sessionStorage.removeItem(throttleKey!);
    }
    const res = await fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
        },
        body: new URLSearchParams(body).toString(),
    });
    const json = await res.json();
    if (
        throttleKey &&
        (res.status === 429 ||
            res.status >= 500 ||
            (res.headers.has("Retry-After") && !res.ok))
    ) {
        const sec = Date.now() / 1000;
        sessionStorage.setItem(
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

interface TokenEntity {
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
    lastUpdatedAt?: string;
}

interface AccountEntity {
    homeAccountId: string;
    environment: string;
    realm: string;
    localAccountId: string;
    username: string;
    authorityType: string;
    name?: string;
    clientInfo?: string;
    tenantProfiles?: unknown[];
    lastUpdatedAt?: string;
    cachedByApiId?: number;
}

interface TokenKeys {
    idToken: string[];
    accessToken: string[];
    refreshToken: string[];
}

export interface AuthCodeResponse {
    code: string;
    verifier: string;
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

/** The core client surface returned by createClient. */
export interface AuthClient {
    initialize(): Promise<void>;
    addEventCallback(cb: EventCallback): string | null;
    removeEventCallback(id: string): void;
    getAllAccounts(filter?: AccountFilter): AccountInfo[];
    getAccount(filter: AccountFilter): AccountInfo | null;
    getActiveAccount(): AccountInfo | null;
    setActiveAccount(account: AccountInfo | null): void;
    loginRedirect(req: TokenRequest): Promise<void>;
    acquireTokenRedirect(req: TokenRequest): Promise<void>;
    handleRedirectPromise(): Promise<AuthenticationResult | null>;
    ssoSilent(req: TokenRequest): Promise<AuthenticationResult>;
    acquireTokenSilent(req: TokenRequest): Promise<AuthenticationResult>;
    logoutRedirect(req?: { account?: AccountInfo | null }): Promise<void>;
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
    authorizeUrl(req: TokenRequest): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
        correlationId: string;
        nonce: string;
        ccs?: string;
    }>;
    pollForCode(
        win: { location: Location; closed?: boolean },
        state: string,
        timeoutMs: number
    ): Promise<string>;
    redeem(res: AuthCodeResponse): Promise<AuthenticationResult>;
    clearAccount(account?: AccountInfo | null): void;
    logoutUrl(
        req?: { postLogoutRedirectUri?: string; correlationId?: string },
        interactionType?: string
    ): string;
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

    const uninitialized = () =>
        new BrowserAuthError("uninitialized_public_client_application");

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
    const emit = (
        eventType: string,
        interactionType?: string,
        payload?: unknown,
        error?: unknown
    ) => {
        const m: EventMessage = {
            eventType,
            interactionType: interactionType ?? null,
            payload: payload ?? null,
            error: error ?? null,
            timestamp: Date.now(),
        };
        listeners.forEach((l) => l(m));
    };

    // ---- cache (real-MSAL v5 schema) ----
    /** cache environment: real MSAL uses the cloud's preferred_cache host */
    const env = () => {
        const host = new URL(authority).host;
        return /login\.microsoftonline\.com|login\.microsoft\.com|sts\.windows\.net/.test(
            host
        )
            ? "login.windows.net"
            : host;
    };

    const readJSON = <T,>(key: string): T | null => {
        const raw = sessionStorage.getItem(key);
        return raw ? (JSON.parse(raw) as T) : null;
    };

    const writeJSON = (key: string, value: unknown) => {
        sessionStorage.setItem(key, JSON.stringify(value));
    };

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
            const t = readJSON<TokenEntity>(k);
            if (t && t.clientId === clientId && match(t)) {
                return t;
            }
        }
        return undefined;
    };

    const toAccountInfo = (e: AccountEntity): AccountInfo => {
        const id = findCred(
            tokenKeys().idToken,
            (t) => t.homeAccountId === e.homeAccountId
        );
        return {
            authorityType: e.authorityType,
            dataBoundary: undefined,
            environment: e.environment,
            homeAccountId: e.homeAccountId,
            idToken: id?.secret,
            idTokenClaims: id ? decodeJwt(id.secret) : {},
            kmsi: undefined,
            localAccountId: e.localAccountId,
            loginHint: undefined,
            name: e.name,
            nativeAccountId: undefined,
            tenantId: e.realm,
            tenantProfiles: e.tenantProfiles,
            upn: undefined,
            username: e.username,
        };
    };

    const matchesFilter = (a: AccountInfo, f: AccountFilter): boolean =>
        (!f.homeAccountId || a.homeAccountId === f.homeAccountId) &&
        (!f.localAccountId || a.localAccountId === f.localAccountId) &&
        (!f.username ||
            a.username.toLowerCase() === f.username.toLowerCase());

    const getAllAccounts = (filter?: AccountFilter): AccountInfo[] =>
        accountKeys()
            .map((k) => readJSON<AccountEntity>(k))
            .filter((e): e is AccountEntity => !!e)
            .map(toAccountInfo)
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
        req: TokenRequest
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
        // real's wire state: base64 lib state, "|<custom>" appended when the
        // request carries one (result.state echoes only the custom part)
        const state =
            btoa(JSON.stringify({ id: crypto.randomUUID() })) +
            (req.state ? `|${req.state}` : "");
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

    /** poll a window/iframe we opened until it lands back on redirectUri with a code */
    const pollForCode = (
        win: { location: Location; closed?: boolean },
        state: string,
        timeoutMs: number
    ): Promise<string> =>
        new Promise((resolve, reject) => {
            const started = Date.now();
            const timer = setInterval(() => {
                // no popup-close detection, like real 5.16: a closed window
                // simply never delivers a response and the bridge times out
                if (Date.now() - started > timeoutMs) {
                    clearInterval(timer);
                    return reject(
                        new BrowserAuthError(
                            "timed_out",
                            "redirect_bridge_timeout"
                        )
                    );
                }
                let hash = "";
                try {
                    hash = win.location.hash;
                } catch {
                    return; // still cross-origin: keep waiting
                }
                const params = new URLSearchParams(hash.slice(1));
                const err = params.get("error");
                const code = params.get("code");
                if (err) {
                    clearInterval(timer);
                    reject(
                        classify(err, params.get("error_description") ?? "")
                    );
                } else if (code) {
                    clearInterval(timer);
                    params.get("state") === state
                        ? resolve(code)
                        : reject(new ClientAuthError("state_mismatch"));
                }
            }, 50);
        });

    // ---- token redemption ----
    const tokenRequest = async (
        scopes: string[],
        grant: Record<string, string>,
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
            throttleKey
        );
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

        const accountKey = `${base}|${realm}`;
        const idKey = `${base}|idtoken|${clientId}|${realm}||`;
        const atKey = `${base}|accesstoken|${clientId}|${realm}|${grantedStr.toLowerCase()}|`;
        const rtKey = `${base}|refreshtoken|${clientId}|||`;

        const entity: AccountEntity = {
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
                    isHomeTenant: true,
                },
            ],
            lastUpdatedAt: ts,
            cachedByApiId: meta?.apiId,
        };
        writeJSON(accountKey, entity);
        writeJSON(idKey, {
            credentialType: "IdToken",
            homeAccountId,
            environment,
            clientId,
            secret: json.id_token,
            realm,
            lastUpdatedAt: ts,
        } satisfies TokenEntity);
        writeJSON(atKey, {
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
        } satisfies TokenEntity);
        if (json.refresh_token) {
            writeJSON(rtKey, {
                credentialType: "RefreshToken",
                homeAccountId,
                environment,
                clientId,
                secret: json.refresh_token,
                lastUpdatedAt: ts,
            } satisfies TokenEntity);
        }

        const keys = tokenKeys();
        const add = (list: string[], k: string) =>
            list.includes(k) ? list : [...list, k];
        writeJSON(tokenKeysKey, {
            idToken: add(keys.idToken, idKey),
            accessToken: add(keys.accessToken, atKey),
            refreshToken: json.refresh_token
                ? add(keys.refreshToken, rtKey)
                : keys.refreshToken,
        });
        const acctKeys = accountKeys();
        if (!acctKeys.includes(accountKey)) {
            writeJSON(`${P}.account.keys`, [...acctKeys, accountKey]);
        }
        // real emits no same-tab accountAdded event; cross-tab propagation
        // is the localStorage/BroadcastChannel feature's job
        return {
            authority: `${reqAuthority}/`,
            uniqueId: localAccountId,
            tenantId: realm,
            scopes: [...new Set(grantedStr.split(" "))],
            account: toAccountInfo(entity),
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
                code_verifier: res.verifier,
                ...(res.redirectUri && { redirect_uri: res.redirectUri }),
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
        // back from a redirect (or clean load): release the interaction lock
        // set before navigating away, like real's handleRedirectPromise
        unlock();
        const params = new URLSearchParams(location.hash.slice(1));
        const code = params.get("code");
        const err = params.get("error");
        const stored = sessionStorage.getItem("msal.request");
        // clean load: real resolves null silently, no handleRedirect events
        if ((!code && !err) || !stored) return null;
        sessionStorage.removeItem("msal.request");
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
        try {
            if (params.get("state") !== state) {
                // forged/unknown state: real treats the response as not ours
                // and resolves null (no failure event, no throw)
                return null;
            }
            history.replaceState(null, "", location.pathname + location.search);
            if (err) {
                // login was cancelled/denied at the IdP
                throw classify(err, params.get("error_description") ?? "");
            }
            const result = await redeem({
                code: code!,
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
        } = await authorizeUrl({ ...req, prompt: req.prompt ?? "none" });
        const frame = document.createElement("iframe");
        frame.style.display = "none";
        document.body.append(frame);
        try {
            frame.src = url;
            const code = await pollForCode(
                frame.contentWindow! as Window,
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
        const at = useAT
            ? findCred(keys.accessToken, (t) => {
                  const target = (t.target ?? "").toLowerCase().split(" ");
                  return (
                      t.homeAccountId === account.homeAccountId &&
                      wanted.every((sc) => target.includes(sc)) &&
                      Number(t.expiresOn) - 300 > Date.now() / 1000
                  );
              })
            : undefined;
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
                }
            } else if (!useFrame) {
                throw new InteractionRequiredAuthError("no_tokens_found");
            }
        }
        // last resort: hidden iframe with prompt=none
        const result = await silentFrame(
            { ...req, account, loginHint: account.username },
            864 // ApiId.acquireTokenSilent_authCode
        );
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
                stays(k) ? kept[type].push(k) : sessionStorage.removeItem(k);
            }
        }
        const keptAccounts: string[] = [];
        for (const k of accountKeys()) {
            stays(k) ? keptAccounts.push(k) : sessionStorage.removeItem(k);
        }
        if (account) {
            writeJSON(tokenKeysKey, kept);
            writeJSON(`${P}.account.keys`, keptAccounts);
            const f = readJSON<{ homeAccountId: string }>(activeKey);
            if (f?.homeAccountId === account.homeAccountId) {
                sessionStorage.removeItem(activeKey);
            }
        } else {
            sessionStorage.removeItem(tokenKeysKey);
            sessionStorage.removeItem(`${P}.account.keys`);
            sessionStorage.removeItem(activeKey);
        }
    };

    const client: AuthClient = {
        async initialize() {
            // second initialize is a silent no-op (no events), like real
            if (initialized) return;
            emit(EventType.INITIALIZE_START);
            // per-instance memory only: real MSAL re-fetches discovery for
            // every new instance and leaves no such key in storage
            metadata ??= await (
                await fetch(
                    `${authority}/v2.0/.well-known/openid-configuration`
                )
            ).json();
            // real tracks lib up/downgrades via this key (trackVersionChanges)
            sessionStorage.setItem("msal.version", WIRE_ID["x-client-VER"]);
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

        setActiveAccount(account: AccountInfo | null) {
            if (account) {
                writeJSON(activeKey, {
                    homeAccountId: account.homeAccountId,
                    localAccountId: account.localAccountId,
                    tenantId: account.tenantId,
                });
            } else {
                sessionStorage.removeItem(activeKey);
            }
            emit(EventType.ACTIVE_ACCOUNT_CHANGED, undefined, account);
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
            location.assign(url);
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
                shared = silentLadder(validRequest, account)
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
            // the page navigates away, so logoutStart is the only event a
            // same-page listener can see (like real's RedirectClient.logout)
            emit(EventType.LOGOUT_START, "redirect", validRequest);
            clearAccount(req?.account);
            location.assign(logoutUrl(validRequest));
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
        pollForCode,
        redeem,
        clearAccount,
        logoutUrl,
    };
    for (const f of features) f(ctx);
    return client;
}
