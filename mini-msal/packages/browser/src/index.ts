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

export interface AccountInfo {
    homeAccountId: string;
    username: string;
    localAccountId: string;
    tenantId: string;
    name?: string;
    idTokenClaims: Record<string, unknown>;
}

export interface AuthenticationResult {
    accessToken: string;
    idToken: string;
    scopes: string[];
    expiresOn: Date;
    account: AccountInfo;
    fromCache: boolean;
}

export interface TokenRequest {
    scopes: string[];
    account?: AccountInfo;
    loginHint?: string;
    prompt?: string;
    redirectUri?: string;
    cacheLookupPolicy?: number;
}

export interface Config {
    auth: {
        clientId: string;
        authority?: string;
        redirectUri?: string;
        postLogoutRedirectUri?: string;
    };
    system?: {
        popupBridgeTimeout?: number;
        iframeBridgeTimeout?: number;
    };
}

/** default error prose, same as real MSAL's */
const AKA = (code: string) =>
    `See https://aka.ms/msal.js.errors#${code} for details`;

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

export const EventType = {
    LOGIN_SUCCESS: "msal:loginSuccess",
    LOGIN_FAILURE: "msal:loginFailure",
    LOGOUT_SUCCESS: "msal:logoutSuccess",
    HANDLE_REDIRECT_END: "msal:handleRedirectEnd",
    ACQUIRE_TOKEN_SUCCESS: "msal:acquireTokenSuccess",
    ACCOUNT_ADDED: "msal:accountAdded",
    ACCOUNT_REMOVED: "msal:accountRemoved",
    ACTIVE_ACCOUNT_CHANGED: "msal:activeAccountChanged",
} as const;

export const InteractionType = {
    Redirect: "redirect",
    Popup: "popup",
    Silent: "silent",
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
    payload?: unknown;
    error?: unknown;
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

async function post(url: string, body: Record<string, string>) {
    const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams(body).toString(),
    });
    const json = await res.json();
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
    tokenType?: string;
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
    emit(eventType: string, payload?: unknown, error?: unknown): void;
    preflight(): void;
    authorizeUrl(
        req: TokenRequest,
        extra?: Record<string, string>
    ): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
    }>;
    pollForCode(
        win: { location: Location; closed?: boolean },
        state: string,
        timeoutMs: number
    ): Promise<string>;
    redeem(res: AuthCodeResponse): Promise<AuthenticationResult>;
    clearAccount(account?: AccountInfo | null): void;
    logoutUrl(): string;
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

    const uninitialized = () =>
        new BrowserAuthError("uninitialized_public_client_application");

    // ---- events ----
    const emit = (eventType: string, payload?: unknown, error?: unknown) => {
        listeners.forEach((l) => l({ eventType, payload, error }));
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
            homeAccountId: e.homeAccountId,
            username: e.username,
            localAccountId: e.localAccountId,
            tenantId: e.realm,
            name: e.name,
            idTokenClaims: id ? decodeJwt(id.secret) : {},
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

    // ---- authorize-request plumbing ----
    const authorizeUrl = async (
        req: TokenRequest,
        extra?: Record<string, string>
    ): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
    }> => {
        const { verifier, challenge } = await pkce();
        const state = randomString();
        const reqRedirectUri = req.redirectUri
            ? new URL(req.redirectUri, location.href).href
            : redirectUri;
        const url = new URL(metadata!.authorization_endpoint);
        const p = url.searchParams;
        p.set("client_id", clientId);
        p.set("response_type", "code");
        p.set("redirect_uri", reqRedirectUri);
        p.set(
            "scope",
            `openid profile offline_access ${req.scopes.join(" ")}`
        );
        p.set("state", state);
        p.set("code_challenge", challenge);
        p.set("code_challenge_method", "S256");
        p.set("response_mode", "fragment");
        if (req.loginHint) p.set("login_hint", req.loginHint);
        if (req.prompt) p.set("prompt", req.prompt);
        for (const k in extra) p.set(k, extra[k]);
        return { url: url.href, verifier, state, redirectUri: reqRedirectUri };
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
        grant: Record<string, string>
    ): Promise<AuthenticationResult> => {
        const json = await post(metadata!.token_endpoint, {
            redirect_uri: redirectUri,
            ...grant,
            client_id: clientId,
            scope: `openid profile offline_access ${scopes.join(" ")}`,
        });
        const claims = decodeJwt(json.id_token);
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
        const account: AccountInfo = {
            homeAccountId,
            username: claims.preferred_username ?? claims.email ?? "",
            localAccountId: claims.oid ?? claims.sub,
            tenantId: claims.tid ?? "",
            name: claims.name,
            idTokenClaims: claims,
        };
        const grantedStr: string = json.scope ?? scopes.join(" ");
        const now = Math.floor(Date.now() / 1000);
        const expiresOn = now + json.expires_in;
        const environment = env();
        const realm = account.tenantId;
        const base = `${P}|${homeAccountId}|${environment}`;

        const accountKey = `${base}|${realm}`;
        const idKey = `${base}|idtoken|${clientId}|${realm}||`;
        const atKey = `${base}|accesstoken|${clientId}|${realm}|${grantedStr.toLowerCase()}|`;
        const rtKey = `${base}|refreshtoken|${clientId}|||`;

        writeJSON(accountKey, {
            homeAccountId,
            environment,
            realm,
            localAccountId: account.localAccountId,
            username: account.username,
            authorityType: "MSSTS",
            name: account.name,
            clientInfo: json.client_info,
            tenantProfiles: [
                {
                    tenantId: realm,
                    localAccountId: account.localAccountId,
                    name: account.name,
                    username: account.username,
                    isHomeTenant: true,
                },
            ],
        } satisfies AccountEntity);
        writeJSON(idKey, {
            credentialType: "IdToken",
            homeAccountId,
            environment,
            clientId,
            secret: json.id_token,
            realm,
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
            tokenType: "Bearer",
        } satisfies TokenEntity);
        if (json.refresh_token) {
            writeJSON(rtKey, {
                credentialType: "RefreshToken",
                homeAccountId,
                environment,
                clientId,
                secret: json.refresh_token,
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
        const isNewAccount = !acctKeys.includes(accountKey);
        if (isNewAccount) {
            writeJSON(`${P}.account.keys`, [...acctKeys, accountKey]);
        }
        if (!readJSON(activeKey)) {
            writeJSON(activeKey, {
                homeAccountId,
                localAccountId: account.localAccountId,
                tenantId: account.tenantId,
            });
        }
        if (isNewAccount) {
            emit(EventType.ACCOUNT_ADDED, account);
        }
        return {
            accessToken: json.access_token,
            idToken: json.id_token,
            scopes: grantedStr.toLowerCase().split(" "),
            expiresOn: new Date(expiresOn * 1000),
            account,
            fromCache: false,
        };
    };

    const redeem = (res: AuthCodeResponse): Promise<AuthenticationResult> =>
        tokenRequest(res.scopes, {
            grant_type: "authorization_code",
            code: res.code,
            code_verifier: res.verifier,
            ...(res.redirectUri && { redirect_uri: res.redirectUri }),
        });

    const redeemRefresh = (
        scopes: string[],
        refreshToken: string
    ): Promise<AuthenticationResult> =>
        tokenRequest(scopes, {
            grant_type: "refresh_token",
            refresh_token: refreshToken,
        });

    // ---- interactive: redirect ----
    const processRedirect = async (): Promise<AuthenticationResult | null> => {
        if (window !== window.parent) {
            // app re-loaded inside our own hidden iframe: leave the hash for
            // the opener's poller (same behavior as real MSAL's iframe guard)
            return null;
        }
        const params = new URLSearchParams(location.hash.slice(1));
        const code = params.get("code");
        const err = params.get("error");
        const stored = sessionStorage.getItem("msal.request");
        try {
            if ((!code && !err) || !stored) return null;
            sessionStorage.removeItem("msal.request");
            const { verifier, state, scopes } = JSON.parse(stored);
            if (params.get("state") !== state) {
                throw new ClientAuthError("state_mismatch");
            }
            history.replaceState(null, "", location.pathname + location.search);
            if (err) {
                // login was cancelled/denied at the IdP
                throw classify(err, params.get("error_description") ?? "");
            }
            const result = await redeem({ code: code!, verifier, scopes });
            emit(EventType.LOGIN_SUCCESS, result);
            return result;
        } catch (e) {
            emit(EventType.LOGIN_FAILURE, undefined, e);
            throw e;
        } finally {
            emit(EventType.HANDLE_REDIRECT_END);
        }
    };

    // ---- silent ----
    const ssoSilent = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        preflight();
        const { url, verifier, state, redirectUri: ru } = await authorizeUrl(
            req,
            { prompt: "none" }
        );
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
            });
        } finally {
            frame.remove();
        }
    };

    // ---- logout ----
    const logoutUrl = (): string => {
        const url = new URL(metadata!.end_session_endpoint);
        url.searchParams.set(
            "post_logout_redirect_uri",
            new URL(
                config.auth.postLogoutRedirectUri ?? redirectUri,
                location.href
            ).href
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
        emit(EventType.ACCOUNT_REMOVED, account);
        emit(EventType.LOGOUT_SUCCESS);
    };

    const client: AuthClient = {
        async initialize() {
            initialized = true;
            const key = `msal.meta.${authority}`;
            const cached = sessionStorage.getItem(key);
            if (cached) {
                metadata = JSON.parse(cached);
                return;
            }
            metadata = await (
                await fetch(
                    `${authority}/v2.0/.well-known/openid-configuration`
                )
            ).json();
            sessionStorage.setItem(key, JSON.stringify(metadata));
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
            emit(EventType.ACTIVE_ACCOUNT_CHANGED, account);
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
            const { url, verifier, state } = await authorizeUrl(req);
            sessionStorage.setItem(
                "msal.request",
                JSON.stringify({ verifier, state, scopes: req.scopes })
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
            const keys = tokenKeys();
            const wanted = req.scopes.map((sc) => sc.toLowerCase());
            const at =
                req.cacheLookupPolicy === CacheLookupPolicy.Skip
                    ? undefined
                    : findCred(keys.accessToken, (t) => {
                          const target = (t.target ?? "")
                              .toLowerCase()
                              .split(" ");
                          return (
                              t.homeAccountId === account.homeAccountId &&
                              wanted.every((sc) => target.includes(sc)) &&
                              Number(t.expiresOn) - 300 > Date.now() / 1000
                          );
                      });
            if (at) {
                const id = findCred(
                    keys.idToken,
                    (t) => t.homeAccountId === account.homeAccountId
                );
                const result = {
                    accessToken: at.secret,
                    idToken: id?.secret ?? "",
                    scopes: (at.target ?? "").toLowerCase().split(" "),
                    expiresOn: new Date(Number(at.expiresOn) * 1000),
                    account,
                    fromCache: true,
                };
                emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
                return result;
            }
            const rt = findCred(
                keys.refreshToken,
                (t) => t.homeAccountId === account.homeAccountId
            );
            if (rt) {
                try {
                    const result = await redeemRefresh(req.scopes, rt.secret);
                    emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
                    return result;
                } catch (e) {
                    if (!(e instanceof InteractionRequiredAuthError)) throw e;
                }
            }
            // last resort: hidden iframe with prompt=none
            const result = await ssoSilent({
                ...req,
                loginHint: account.username,
            });
            emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
            return result;
        },

        async logoutRedirect(req?: {
            account?: AccountInfo | null;
        }): Promise<void> {
            preflight();
            clearAccount(req?.account);
            location.assign(logoutUrl());
        },
    };

    const ctx: ClientContext = {
        config,
        client,
        emit,
        preflight,
        authorizeUrl,
        pollForCode,
        redeem,
        clearAccount,
        logoutUrl,
    };
    for (const f of features) f(ctx);
    return client;
}
