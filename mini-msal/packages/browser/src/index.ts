/**
 * mini-msal: a from-scratch minimal implementation of the public API surface
 * exercised by the full-flow app (src/app.tsx):
 *   - OIDC discovery (authority metadata fetched at runtime)
 *   - loginRedirect / loginPopup / ssoSilent (auth-code + PKCE S256)
 *   - handleRedirectPromise (code exchange at the token endpoint)
 *   - acquireTokenSilent: cache -> refresh-token grant -> hidden-iframe
 *     prompt=none fallback; acquireTokenPopup / acquireTokenRedirect
 *   - single-account sessionStorage cache + active account
 *   - getAllAccounts / getAccountByHomeId / getAccountByUsername
 *   - event callbacks (EventMessage-shaped, add/remove) + account storage events
 *   - logoutRedirect / logoutPopup
 *   - InteractionRequiredAuthError / BrowserAuthError classification
 *
 * NOT implemented (see report): platform broker/WAM, nested app auth,
 * multi-account cache, localStorage/cookie/IndexedDB storage, telemetry,
 * logger, CIAM/B2C/ADFS authority variants, claims/CAE, PoP tokens.
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
}

export class AuthError extends Error {
    constructor(public errorCode: string, public errorMessage: string) {
        super(`${errorCode}: ${errorMessage}`);
    }
}

export class InteractionRequiredAuthError extends AuthError {}
export class BrowserAuthError extends AuthError {}

const INTERACTION_CODES =
    /^(interaction_required|consent_required|login_required|no_tokens_found|no_account|monitor_window_timeout)$/;

function classify(code: string, desc: string): AuthError {
    return INTERACTION_CODES.test(code)
        ? new InteractionRequiredAuthError(code, desc)
        : new AuthError(code, desc);
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
    if (json.error) {
        throw classify(json.error, json.error_description);
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

interface AuthCodeResponse {
    code: string;
    verifier: string;
    scopes: string[];
    redirectUri?: string;
}

export class PublicClientApplication {
    private authority: string;
    private redirectUri: string;
    private clientId: string;
    private metadata?: {
        authorization_endpoint: string;
        token_endpoint: string;
        end_session_endpoint: string;
    };
    private listeners = new Map<string, EventCallback>();
    private nextListenerId = 0;
    private redirectResult: Promise<AuthenticationResult | null> | null = null;

    constructor(private config: Config) {
        this.clientId = config.auth.clientId;
        this.authority = (
            config.auth.authority ?? "https://login.microsoftonline.com/common"
        ).replace(/\/$/, "");
        this.redirectUri = new URL(
            config.auth.redirectUri ?? "/",
            location.href
        ).href;
    }

    async initialize(): Promise<void> {
        const key = `msal.meta.${this.authority}`;
        const cached = sessionStorage.getItem(key);
        if (cached) {
            this.metadata = JSON.parse(cached);
            return;
        }
        this.metadata = await (
            await fetch(
                `${this.authority}/v2.0/.well-known/openid-configuration`
            )
        ).json();
        sessionStorage.setItem(key, JSON.stringify(this.metadata));
    }

    // ---- events ----
    addEventCallback(cb: EventCallback): string | null {
        const id = String(this.nextListenerId++);
        this.listeners.set(id, cb);
        return id;
    }

    removeEventCallback(id: string): void {
        this.listeners.delete(id);
    }

    private emit(eventType: string, payload?: unknown, error?: unknown) {
        this.listeners.forEach((l) => l({ eventType, payload, error }));
    }

    // ---- cache (real-MSAL v5 schema) ----
    /** cache environment: real MSAL uses the cloud's preferred_cache host */
    private get env(): string {
        const host = new URL(this.authority).host;
        return /login\.microsoftonline\.com|login\.microsoft\.com|sts\.windows\.net/.test(
            host
        )
            ? "login.windows.net"
            : host;
    }

    private readJSON<T>(key: string): T | null {
        const raw = sessionStorage.getItem(key);
        return raw ? (JSON.parse(raw) as T) : null;
    }

    private writeJSON(key: string, value: unknown) {
        sessionStorage.setItem(key, JSON.stringify(value));
    }

    private get tokenKeysKey() {
        return `${P}.token.keys.${this.clientId}`;
    }

    private tokenKeys(): TokenKeys {
        return (
            this.readJSON<TokenKeys>(this.tokenKeysKey) ?? {
                idToken: [],
                accessToken: [],
                refreshToken: [],
            }
        );
    }

    private accountKeys(): string[] {
        return this.readJSON<string[]>(`${P}.account.keys`) ?? [];
    }

    private findCred(
        list: string[],
        match: (t: TokenEntity) => boolean
    ): TokenEntity | undefined {
        for (const k of list) {
            const t = this.readJSON<TokenEntity>(k);
            if (t && t.clientId === this.clientId && match(t)) {
                return t;
            }
        }
        return undefined;
    }

    private toAccountInfo(e: AccountEntity): AccountInfo {
        const id = this.findCred(
            this.tokenKeys().idToken,
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
    }

    getAllAccounts(): AccountInfo[] {
        return this.accountKeys()
            .map((k) => this.readJSON<AccountEntity>(k))
            .filter((e): e is AccountEntity => !!e)
            .map((e) => this.toAccountInfo(e));
    }

    getAccount(filter: {
        homeAccountId?: string;
        localAccountId?: string;
        username?: string;
    }): AccountInfo | null {
        return (
            this.getAllAccounts().find(
                (a) =>
                    (!filter.homeAccountId ||
                        a.homeAccountId === filter.homeAccountId) &&
                    (!filter.localAccountId ||
                        a.localAccountId === filter.localAccountId) &&
                    (!filter.username ||
                        a.username.toLowerCase() ===
                            filter.username.toLowerCase())
            ) ?? null
        );
    }

    private get activeKey() {
        return `msal.${this.clientId}.active-account-filters`;
    }

    getActiveAccount(): AccountInfo | null {
        const f = this.readJSON<{ homeAccountId: string }>(this.activeKey);
        return f ? this.getAccount({ homeAccountId: f.homeAccountId }) : null;
    }

    setActiveAccount(account: AccountInfo | null) {
        if (account) {
            this.writeJSON(this.activeKey, {
                homeAccountId: account.homeAccountId,
                localAccountId: account.localAccountId,
                tenantId: account.tenantId,
            });
        } else {
            sessionStorage.removeItem(this.activeKey);
        }
        this.emit(EventType.ACTIVE_ACCOUNT_CHANGED, account);
    }

    /**
     * Same environment guards as real MSAL: when this app is re-booted inside
     * one of our own hidden iframes (auth response in the hash) or popups
     * (window named "msal.*"), auth APIs refuse to run — the opener's poller
     * owns the response.
     */
    private preflight() {
        if (
            window !== window.parent &&
            /[#&](code|error)=/.test(location.hash)
        ) {
            throw new BrowserAuthError(
                "block_iframe_reload",
                "Auth response in iframe"
            );
        }
        if (window.name.startsWith("msal.")) {
            throw new BrowserAuthError(
                "block_nested_popups",
                "Auth APIs blocked inside MSAL-opened popups"
            );
        }
    }

    // ---- authorize-request plumbing ----
    private async authorizeUrl(
        req: TokenRequest,
        extra?: Record<string, string>
    ): Promise<{
        url: string;
        verifier: string;
        state: string;
        redirectUri: string;
    }> {
        const { verifier, challenge } = await pkce();
        const state = randomString();
        const redirectUri = req.redirectUri
            ? new URL(req.redirectUri, location.href).href
            : this.redirectUri;
        const url = new URL(this.metadata!.authorization_endpoint);
        const p = url.searchParams;
        p.set("client_id", this.clientId);
        p.set("response_type", "code");
        p.set("redirect_uri", redirectUri);
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
        return { url: url.href, verifier, state, redirectUri };
    }

    /** poll a window/iframe we opened until it lands back on redirectUri with a code */
    private pollForCode(
        win: { location: Location; closed?: boolean },
        state: string,
        timeoutMs: number
    ): Promise<string> {
        return new Promise((resolve, reject) => {
            const started = Date.now();
            const timer = setInterval(() => {
                if (win.closed) {
                    clearInterval(timer);
                    return reject(
                        new BrowserAuthError(
                            "user_cancelled",
                            "Window was closed"
                        )
                    );
                }
                if (Date.now() - started > timeoutMs) {
                    clearInterval(timer);
                    return reject(
                        new InteractionRequiredAuthError(
                            "monitor_window_timeout",
                            "Token acquisition timed out"
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
                        classify(err, params.get("error_description") ?? err)
                    );
                } else if (code) {
                    clearInterval(timer);
                    params.get("state") === state
                        ? resolve(code)
                        : reject(
                              new AuthError(
                                  "state_mismatch",
                                  "State does not match"
                              )
                          );
                }
            }, 50);
        });
    }

    // ---- interactive: redirect ----
    async loginRedirect(req: TokenRequest): Promise<void> {
        return this.acquireTokenRedirect(req);
    }

    async acquireTokenRedirect(req: TokenRequest): Promise<void> {
        this.preflight();
        if (window !== window.parent) {
            // same guard as real MSAL: no full-page redirects from iframes
            throw new BrowserAuthError(
                "redirect_in_iframe",
                "Redirect interaction is not allowed in an iframe"
            );
        }
        const { url, verifier, state } = await this.authorizeUrl(req);
        sessionStorage.setItem(
            "msal.request",
            JSON.stringify({ verifier, state, scopes: req.scopes })
        );
        location.assign(url);
    }

    handleRedirectPromise(): Promise<AuthenticationResult | null> {
        return (this.redirectResult ??= this.processRedirect());
    }

    private async processRedirect(): Promise<AuthenticationResult | null> {
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
                throw new AuthError("state_mismatch", "State does not match");
            }
            history.replaceState(null, "", location.pathname + location.search);
            if (err) {
                // login was cancelled/denied at the IdP
                throw classify(err, params.get("error_description") ?? err);
            }
            const result = await this.redeem({ code: code!, verifier, scopes });
            this.emit(EventType.LOGIN_SUCCESS, result);
            return result;
        } catch (e) {
            this.emit(EventType.LOGIN_FAILURE, undefined, e);
            throw e;
        } finally {
            this.emit(EventType.HANDLE_REDIRECT_END);
        }
    }

    // ---- interactive: popup ----
    async loginPopup(req: TokenRequest): Promise<AuthenticationResult> {
        try {
            const result = await this.acquireTokenPopup(req);
            this.emit(EventType.LOGIN_SUCCESS, result);
            return result;
        } catch (e) {
            this.emit(EventType.LOGIN_FAILURE, undefined, e);
            throw e;
        }
    }

    async acquireTokenPopup(req: TokenRequest): Promise<AuthenticationResult> {
        this.preflight();
        const { url, verifier, state, redirectUri } = await this.authorizeUrl(
            req
        );
        const popup = open(
            url,
            "msal.popup",
            "width=483,height=600,popup=yes"
        );
        if (!popup) {
            throw new BrowserAuthError(
                "popup_window_error",
                "Popup was blocked"
            );
        }
        try {
            const code = await this.pollForCode(popup, state, 60_000);
            return await this.redeem({
                code,
                verifier,
                scopes: req.scopes,
                redirectUri,
            });
        } finally {
            popup.close();
        }
    }

    // ---- silent ----
    async ssoSilent(req: TokenRequest): Promise<AuthenticationResult> {
        this.preflight();
        const { url, verifier, state, redirectUri } = await this.authorizeUrl(
            req,
            { prompt: "none" }
        );
        const frame = document.createElement("iframe");
        frame.style.display = "none";
        document.body.append(frame);
        try {
            frame.src = url;
            const code = await this.pollForCode(
                frame.contentWindow! as Window,
                state,
                10_000
            );
            return await this.redeem({
                code,
                verifier,
                scopes: req.scopes,
                redirectUri,
            });
        } finally {
            frame.remove();
        }
    }

    async acquireTokenSilent(req: TokenRequest): Promise<AuthenticationResult> {
        this.preflight();
        const account = req.account ?? this.getActiveAccount();
        if (
            !account ||
            !this.getAccount({ homeAccountId: account.homeAccountId })
        ) {
            throw new InteractionRequiredAuthError(
                "no_account",
                "Sign in first"
            );
        }
        const keys = this.tokenKeys();
        const wanted = req.scopes.map((sc) => sc.toLowerCase());
        const at =
            req.cacheLookupPolicy === CacheLookupPolicy.Skip
                ? undefined
                : this.findCred(keys.accessToken, (t) => {
                      const target = (t.target ?? "").toLowerCase().split(" ");
                      return (
                          t.homeAccountId === account.homeAccountId &&
                          wanted.every((sc) => target.includes(sc)) &&
                          Number(t.expiresOn) - 300 > Date.now() / 1000
                      );
                  });
        if (at) {
            const id = this.findCred(
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
            this.emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
            return result;
        }
        const rt = this.findCred(
            keys.refreshToken,
            (t) => t.homeAccountId === account.homeAccountId
        );
        if (rt) {
            try {
                const result = await this.redeemRefresh(req.scopes, rt.secret);
                this.emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
                return result;
            } catch (e) {
                if (!(e instanceof InteractionRequiredAuthError)) throw e;
            }
        }
        // last resort: hidden iframe with prompt=none
        const result = await this.ssoSilent({
            ...req,
            loginHint: account.username,
        });
        this.emit(EventType.ACQUIRE_TOKEN_SUCCESS, result);
        return result;
    }

    // ---- token redemption ----
    private redeem(res: AuthCodeResponse): Promise<AuthenticationResult> {
        return this.tokenRequest(res.scopes, {
            grant_type: "authorization_code",
            code: res.code,
            code_verifier: res.verifier,
            ...(res.redirectUri && { redirect_uri: res.redirectUri }),
        });
    }

    private redeemRefresh(
        scopes: string[],
        refreshToken: string
    ): Promise<AuthenticationResult> {
        return this.tokenRequest(scopes, {
            grant_type: "refresh_token",
            refresh_token: refreshToken,
        });
    }

    private async tokenRequest(
        scopes: string[],
        grant: Record<string, string>
    ): Promise<AuthenticationResult> {
        const json = await post(this.metadata!.token_endpoint, {
            redirect_uri: this.redirectUri,
            ...grant,
            client_id: this.clientId,
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
        const env = this.env;
        const realm = account.tenantId;
        const base = `${P}|${homeAccountId}|${env}`;

        const accountKey = `${base}|${realm}`;
        const idKey = `${base}|idtoken|${this.clientId}|${realm}||`;
        const atKey = `${base}|accesstoken|${this.clientId}|${realm}|${grantedStr.toLowerCase()}|`;
        const rtKey = `${base}|refreshtoken|${this.clientId}|||`;

        this.writeJSON(accountKey, {
            homeAccountId,
            environment: env,
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
        this.writeJSON(idKey, {
            credentialType: "IdToken",
            homeAccountId,
            environment: env,
            clientId: this.clientId,
            secret: json.id_token,
            realm,
        } satisfies TokenEntity);
        this.writeJSON(atKey, {
            credentialType: "AccessToken",
            homeAccountId,
            environment: env,
            clientId: this.clientId,
            secret: json.access_token,
            realm,
            target: grantedStr,
            cachedAt: String(now),
            expiresOn: String(expiresOn),
            extendedExpiresOn: String(expiresOn),
            tokenType: "Bearer",
        } satisfies TokenEntity);
        if (json.refresh_token) {
            this.writeJSON(rtKey, {
                credentialType: "RefreshToken",
                homeAccountId,
                environment: env,
                clientId: this.clientId,
                secret: json.refresh_token,
            } satisfies TokenEntity);
        }

        const keys = this.tokenKeys();
        const add = (list: string[], k: string) =>
            list.includes(k) ? list : [...list, k];
        this.writeJSON(this.tokenKeysKey, {
            idToken: add(keys.idToken, idKey),
            accessToken: add(keys.accessToken, atKey),
            refreshToken: json.refresh_token
                ? add(keys.refreshToken, rtKey)
                : keys.refreshToken,
        });
        const acctKeys = this.accountKeys();
        const isNewAccount = !acctKeys.includes(accountKey);
        if (isNewAccount) {
            this.writeJSON(`${P}.account.keys`, [...acctKeys, accountKey]);
        }
        if (!this.readJSON(this.activeKey)) {
            this.writeJSON(this.activeKey, {
                homeAccountId,
                localAccountId: account.localAccountId,
                tenantId: account.tenantId,
            });
        }
        if (isNewAccount) {
            this.emit(EventType.ACCOUNT_ADDED, account);
        }
        return {
            accessToken: json.access_token,
            idToken: json.id_token,
            scopes: grantedStr.toLowerCase().split(" "),
            expiresOn: new Date(expiresOn * 1000),
            account,
            fromCache: false,
        };
    }

    // ---- logout ----
    private logoutUrl(): string {
        const url = new URL(this.metadata!.end_session_endpoint);
        url.searchParams.set(
            "post_logout_redirect_uri",
            new URL(
                this.config.auth.postLogoutRedirectUri ?? this.redirectUri,
                location.href
            ).href
        );
        return url.href;
    }

    private clearAccount(account?: AccountInfo | null) {
        const stays = (k: string) =>
            !!account && !k.includes(`|${account.homeAccountId}|`);
        const keys = this.tokenKeys();
        const kept: TokenKeys = { idToken: [], accessToken: [], refreshToken: [] };
        for (const type of ["idToken", "accessToken", "refreshToken"] as const) {
            for (const k of keys[type]) {
                stays(k) ? kept[type].push(k) : sessionStorage.removeItem(k);
            }
        }
        const keptAccounts: string[] = [];
        for (const k of this.accountKeys()) {
            stays(k) ? keptAccounts.push(k) : sessionStorage.removeItem(k);
        }
        if (account) {
            this.writeJSON(this.tokenKeysKey, kept);
            this.writeJSON(`${P}.account.keys`, keptAccounts);
            const f = this.readJSON<{ homeAccountId: string }>(this.activeKey);
            if (f?.homeAccountId === account.homeAccountId) {
                sessionStorage.removeItem(this.activeKey);
            }
        } else {
            sessionStorage.removeItem(this.tokenKeysKey);
            sessionStorage.removeItem(`${P}.account.keys`);
            sessionStorage.removeItem(this.activeKey);
        }
        this.emit(EventType.ACCOUNT_REMOVED, account);
        this.emit(EventType.LOGOUT_SUCCESS);
    }

    async logoutRedirect(req?: {
        account?: AccountInfo | null;
    }): Promise<void> {
        this.preflight();
        this.clearAccount(req?.account);
        location.assign(this.logoutUrl());
    }

    async logoutPopup(req?: { account?: AccountInfo | null }): Promise<void> {
        this.preflight();
        const popup = open(
            this.logoutUrl(),
            "msal.popup",
            "width=483,height=600,popup=yes"
        );
        if (!popup) {
            throw new BrowserAuthError(
                "popup_window_error",
                "Popup was blocked"
            );
        }
        // wait for the popup to land back on the post-logout page (server
        // session cleared), then close; events fire only after completion,
        // matching real MSAL's ordering
        const started = Date.now();
        await new Promise<void>((resolve) => {
            const timer = setInterval(() => {
                let done = popup.closed || Date.now() - started > 5000;
                try {
                    done ||= popup.location.origin === location.origin;
                } catch {
                    /* still on the IdP: keep waiting */
                }
                if (done) {
                    clearInterval(timer);
                    resolve();
                }
            }, 50);
        });
        popup.close();
        this.clearAccount(req?.account);
    }
}

/** Shorthand factory: creates and initializes a client in one call. */
export async function createAuth(
    config: Config
): Promise<PublicClientApplication> {
    const auth = new PublicClientApplication(config);
    await auth.initialize();
    return auth;
}
