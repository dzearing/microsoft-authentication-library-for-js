/**
 * Nested app auth (NAA): createNestableClient. When the host page provides
 * window.nestedAppAuthBridge (Teams/Office embedded apps), tokens come from
 * the host over its JSON postMessage protocol — GetInitContext handshake at
 * create time, GetTokenPopup/GetToken per request, browser-cache-first on
 * the silent path — and the web/redirect APIs throw unsupported_method.
 * Without a bridge (or a failed handshake) the provided fallback factory
 * builds a standard web client, like real's
 * createNestablePublicClientApplication.
 * Protocol notes: docs/design/naa-protocol.md.
 */
import {
    createClient,
    AuthError,
    ClientAuthError,
    InteractionRequiredAuthError,
    NestedAppAuthError,
    ServerError,
    EventType,
    type AccountInfo,
    type AuthenticationResult,
    type ClientContext,
    type Config,
    type TokenRequest,
} from "./index.js";

const OIDC = ["openid", "profile", "offline_access"];

interface Bridge {
    initContext: {
        accountContext?: {
            homeAccountId: string;
            environment: string;
            tenantId: string;
        } | null;
    };
    send(method: string, params?: Record<string, unknown>): Promise<any>;
}

/** real's validateBridgeResultOrThrow: an absent payload piece means the
 * host can't do nested app auth */
const required = <T,>(v: T | undefined): T => {
    if (v === undefined) {
        throw { status: "NESTED_APP_AUTH_UNAVAILABLE" };
    }
    return v;
};

const decodeJwt = (token: string): Record<string, any> =>
    JSON.parse(
        atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );

/** real's NestedAppAuthAdapter.fromBridgeError: bridge status -> msal error */
const mapError = (e: any): AuthError => {
    if (e instanceof AuthError) {
        return e;
    }
    switch (e?.status) {
        case "USER_CANCEL":
            return new ClientAuthError("user_canceled");
        case "NO_NETWORK":
            return new ClientAuthError("no_network_connectivity");
        case "ACCOUNT_UNAVAILABLE":
            return new ClientAuthError("no_account_found");
        case "DISABLED":
            return new ClientAuthError("nested_app_auth_bridge_disabled");
        case "NESTED_APP_AUTH_UNAVAILABLE":
            return new ClientAuthError(
                e.code || "nested_app_auth_bridge_disabled",
                e.description || undefined
            );
        case "TRANSIENT_ERROR":
        case "PERSISTENT_ERROR":
            return new ServerError(e.code || "", e.description || undefined);
        case "USER_INTERACTION_REQUIRED":
            return new InteractionRequiredAuthError(
                e.code || "",
                e.description || undefined
            );
        default:
            return e?.status
                ? new AuthError(e.code || "", e.description || undefined)
                : new AuthError("unknown_error", "An unknown error occurred");
    }
};

/** GetInitContext handshake; throws when no bridge is present or the host
 * withholds an initContext (real swallows this and falls back) */
const connect = async (): Promise<Bridge> => {
    // optional host hook, awaited before the handshake like real
    await (window as any).__initializeNestedAppAuth?.();
    const bridge = (window as any).nestedAppAuthBridge;
    const pending = new Map<
        string,
        { resolve: (v: any) => void; reject: (e: unknown) => void }
    >();
    bridge.addEventListener(
        "message",
        (response: string | { data: string }) => {
            const envelope = JSON.parse(
                typeof response === "string" ? response : response.data
            );
            const p = pending.get(envelope.requestId);
            if (p) {
                pending.delete(envelope.requestId);
                envelope.success
                    ? p.resolve(envelope)
                    : p.reject(envelope.error);
            }
        }
    );
    const send = (method: string, params?: Record<string, unknown>) =>
        new Promise<any>((resolve, reject) => {
            const requestId = crypto.randomUUID();
            pending.set(requestId, { resolve, reject });
            bridge.postMessage(
                JSON.stringify({
                    messageType: "NestedAppAuthRequest",
                    method,
                    requestId,
                    sendTime: Date.now(),
                    // real's wire identity (Decision Log: impersonation)
                    clientLibrary: "msal.js.browser",
                    clientLibraryVersion: "5.16.0",
                    ...params,
                })
            );
        });
    return {
        initContext: required((await send("GetInitContext")).initContext),
        send,
    };
};

const naa =
    (bridge: Bridge) =>
    (ctx: ClientContext): void => {
        const c = ctx.client;
        const config = ctx.config;
        // real prioritizes the bridge's account context, then the last token
        // response's, for cache lookups — request.account is never consulted
        let accountCtx = bridge.initContext.accountContext ?? null;

        const unsupported = (): never => {
            throw new NestedAppAuthError("unsupported_method");
        };

        // same shape real's buildMergedClaims produces (rides tokenParams)
        const mergedClaims = (claims?: string): string => {
            const cl = claims ? JSON.parse(claims) : {};
            cl.id_token = {
                signin_state: { essential: false },
                login_hint: { essential: false },
                ...cl.id_token,
            };
            const caps = config.auth.clientCapabilities;
            if (caps?.length) {
                (cl.access_token ??= {}).xms_cc = { values: caps };
            }
            return JSON.stringify(cl);
        };

        // real's toNaaTokenRequest. undefined-valued keys vanish in the
        // bridge JSON; extraParameters is a Map in real, so it serializes as
        // {} — even request extraQueryParameters never hit the wire.
        const toNaaRequest = (
            req: TokenRequest & { resource?: string }
        ): Record<string, any> => ({
            platformBrokerId: req.account?.homeAccountId,
            clientId: config.auth.clientId,
            authority: req.authority,
            resource: req.resource,
            scope: (req.scopes ?? OIDC).join(" "),
            correlationId: req.correlationId,
            claims: mergedClaims(req.claims),
            state: req.state,
            authenticationScheme: "Bearer",
            extraParameters: {},
        });

        // real's fromNaaAccountInfo: id_token claims fill any gaps
        const toAccount = (
            a: Record<string, any>,
            idToken: string,
            claims: Record<string, any>
        ): AccountInfo => {
            const localAccountId =
                a.localAccountId || claims.oid || claims.sub || "";
            const tenantId = a.tenantId || claims.tid || "";
            if (!a.environment) {
                throw new ClientAuthError("invalid_cache_environment");
            }
            const username =
                a.username ||
                claims.preferred_username ||
                claims.upn ||
                claims.emails?.[0] ||
                "";
            const name = a.name || claims.name || "";
            return {
                homeAccountId:
                    a.homeAccountId || `${localAccountId}.${tenantId}`,
                environment: a.environment,
                tenantId,
                username,
                localAccountId,
                name,
                loginHint: a.loginHint || claims.login_hint,
                idToken,
                idTokenClaims: claims,
                nativeAccountId: a.platformBrokerId,
                tenantProfiles: new Map([
                    [
                        tenantId,
                        {
                            tenantId,
                            localAccountId,
                            name,
                            username,
                            isHomeTenant:
                                tenantId ===
                                (
                                    a.homeAccountId ||
                                    `${localAccountId}.${tenantId}`
                                ).split(".")[1],
                        },
                    ],
                ]),
            };
        };

        // real's fromNaaTokenResponse
        const fromResponse = (
            naaReq: Record<string, any>,
            r: { token: Record<string, any>; account: Record<string, any> },
            reqTimestamp: number
        ): AuthenticationResult => {
            if (!r.token.id_token || !r.token.access_token) {
                throw new ClientAuthError("null_or_empty_token");
            }
            const expiresOn = new Date(
                (reqTimestamp + (r.token.expires_in || 0)) * 1000
            );
            const claims = decodeJwt(r.token.id_token);
            const account = toAccount(r.account, r.token.id_token, claims);
            return {
                authority: r.token.authority || account.environment,
                uniqueId: account.localAccountId,
                tenantId: account.tenantId,
                scopes: (r.token.scope || naaReq.scope).split(" "),
                account,
                idToken: r.token.id_token,
                idTokenClaims: claims,
                accessToken: r.token.access_token,
                fromCache: false,
                expiresOn,
                extExpiresOn: expiresOn,
                tokenType: naaReq.authenticationScheme,
                correlationId: naaReq.correlationId,
                state: naaReq.state,
            } as AuthenticationResult;
        };

        const getToken = async (
            method: "GetToken" | "GetTokenPopup",
            tokenParams: Record<string, unknown>
        ) => {
            const r = await bridge.send(method, { tokenParams });
            return { token: required(r.token), account: required(r.account) };
        };

        // real's acquireTokenFromCache: only Default/AccessToken/
        // AccessTokenAndRefreshToken policies, never with claims or
        // forceRefresh; cached results report the AT entity's environment as
        // the authority (toAuthenticationResultFromCache)
        const fromCache = (req: TokenRequest): AuthenticationResult | null => {
            if (
                req.claims ||
                req.forceRefresh ||
                (req.cacheLookupPolicy ?? 0) > 2
            ) {
                return null;
            }
            const account =
                accountCtx &&
                c.getAccount({ homeAccountId: accountCtx.homeAccountId });
            if (!account) {
                return null;
            }
            const scopes = req.scopes?.length ? req.scopes : OIDC;
            const wanted = scopes.map((s) => s.toLowerCase());
            const at = ctx.findToken(
                "accessToken",
                (t) =>
                    t.homeAccountId === account.homeAccountId &&
                    wanted.every((sc) =>
                        (t.target ?? "").toLowerCase().split(" ").includes(sc)
                    ) &&
                    Number(t.expiresOn) - 300 > Date.now() / 1000
            );
            const id = ctx.findToken(
                "idToken",
                (t) => t.homeAccountId === account.homeAccountId
            );
            if (!at || !id) {
                return null;
            }
            return {
                authority: at.environment || account.environment,
                uniqueId: account.localAccountId,
                tenantId: account.tenantId,
                scopes: (at.target || scopes.join(" ")).split(" "),
                account,
                idToken: id.secret,
                idTokenClaims: account.idTokenClaims,
                accessToken: at.secret,
                fromCache: true,
                expiresOn: new Date(Number(at.expiresOn) * 1000),
                extExpiresOn: new Date(Number(at.extendedExpiresOn) * 1000),
                tokenType: "Bearer",
                correlationId: req.correlationId,
                state: req.state,
            } as AuthenticationResult;
        };

        // real's hydrateCache: account + id/access token entities, so later
        // silent calls resolve from the browser cache (failures swallowed)
        const hydrate = async (r: AuthenticationResult) => {
            const a = r.account;
            await ctx.writeAccount({
                homeAccountId: a.homeAccountId,
                environment: a.environment,
                realm: a.tenantId,
                localAccountId: a.localAccountId,
                username: a.username,
                authorityType: "MSSTS",
                name: a.name,
                nativeAccountId: a.nativeAccountId,
                tenantProfiles: a.tenantProfiles
                    ? [...a.tenantProfiles.values()]
                    : undefined,
                lastUpdatedAt: String(Date.now()),
                cachedByApiId: 963, // ApiId.hydrateCache
            });
            await ctx.writeTokens({
                homeAccountId: a.homeAccountId,
                environment: a.environment,
                realm: a.tenantId,
                idToken: r.idToken,
                accessToken: r.accessToken,
                target: r.scopes.join(" "),
                expiresOn: Math.floor(+r.expiresOn / 1000),
            });
        };

        const acquire = async (
            method: "GetToken" | "GetTokenPopup",
            interactionType: string,
            req: TokenRequest
        ): Promise<AuthenticationResult> => {
            const validRequest = {
                ...req,
                correlationId: req.correlationId || crypto.randomUUID(),
            };
            ctx.emit(
                EventType.ACQUIRE_TOKEN_START,
                interactionType,
                validRequest
            );
            try {
                if (method === "GetToken") {
                    const cached = fromCache(validRequest);
                    if (cached) {
                        ctx.emit(
                            EventType.ACQUIRE_TOKEN_SUCCESS,
                            interactionType,
                            cached
                        );
                        return cached;
                    }
                }
                const naaReq = toNaaRequest(validRequest);
                if (method === "GetToken") {
                    naaReq.forceRefresh = validRequest.forceRefresh;
                }
                const reqTimestamp = Math.floor(Date.now() / 1000);
                const result = fromResponse(
                    naaReq,
                    await getToken(method, naaReq),
                    reqTimestamp
                );
                try {
                    await hydrate(result);
                } catch {
                    /* real only logs hydration failures */
                }
                accountCtx = {
                    homeAccountId: result.account.homeAccountId,
                    environment: result.account.environment,
                    tenantId: result.account.tenantId,
                };
                ctx.emit(
                    EventType.ACQUIRE_TOKEN_SUCCESS,
                    interactionType,
                    result
                );
                return result;
            } catch (e) {
                const error = mapError(e);
                ctx.emit(
                    EventType.ACQUIRE_TOKEN_FAILURE,
                    interactionType,
                    undefined,
                    e
                );
                throw error;
            }
        };

        c.acquireTokenPopup = (req: TokenRequest) =>
            acquire("GetTokenPopup", "popup", req);
        c.loginPopup = (req?: TokenRequest) =>
            acquire("GetTokenPopup", "popup", req ?? { scopes: [...OIDC] });
        c.acquireTokenSilent = (req: TokenRequest) =>
            acquire("GetToken", "silent", req);
        c.ssoSilent = (req: TokenRequest) => acquire("GetToken", "silent", req);
        c.handleRedirectPromise = () => Promise.resolve(null);
        c.acquireTokenRedirect = unsupported;
        c.loginRedirect = unsupported;
        c.logoutRedirect = unsupported;
        c.logoutPopup = unsupported;
        c.acquireTokenByCode = unsupported;
        c.addPerformanceCallback = unsupported;
        c.removePerformanceCallback = unsupported;
        c.getPerformanceClient = unsupported;
        c.getRedirectResponse = unsupported;
        c.clearCache = unsupported;
    };

/**
 * real's createNestablePublicClientApplication: bridge handshake first —
 * success yields the NAA client, any failure yields fallback(config) (the
 * caller supplies its standard-client factory, so this module stays
 * tree-shakable).
 */
export async function createNestableClient<T>(
    config: Config,
    fallback: (config: Config) => Promise<T>
): Promise<T> {
    let bridge: Bridge;
    try {
        bridge = await connect();
    } catch {
        // no (or failed) bridge: standard web client, like real
        return fallback(config);
    }
    const client = createClient(config, [naa(bridge)]);
    await client.initialize();
    return client as unknown as T;
}
