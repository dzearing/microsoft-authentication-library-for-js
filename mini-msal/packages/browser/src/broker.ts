/**
 * Platform-broker feature (WAM), both transports. Probed at initialize()
 * when system.allowPlatformBroker is set: the DOM path
 * (navigator.platformAuthentication, needs
 * experimental.allowPlatformBrokerWithDOM too) is tried first, then the
 * browser-extension path (Handshake over window.postMessage +
 * MessageChannel — preferred extension id, retried with any). Attaches
 * acquireTokenByCode({nativeAccountId}) — the hybrid entry point into the
 * broker — and routes acquireTokenSilent through the broker when the
 * account carries a nativeAccountId (ctx.nativeSilent seam).
 * Protocol notes: docs/design/broker-protocol.md.
 */
import {
    AuthError,
    BrowserAuthError,
    InteractionRequiredAuthError,
    EventType,
    isKmsi,
    type AccountEntity,
    type AccountInfo,
    type AuthenticationResult,
    type ClientContext,
    type TokenRequest,
} from "./index.js";

const AKA = (code: string) =>
    `See https://aka.ms/msal.js.errors#${code} for details`;

const CHANNEL_ID = "53ee284d-920a-4b59-9d30-a60315b26836";
const PREFERRED_EXTENSION_ID = "ppnbnpeolgkicgegkbkbjmhlideopiji";

/** transport handle: per-transport x-client-xtra-sku value + a GetToken
 * sender resolving to the camelCase shape handleResponse consumes */
interface Provider {
    sku: string;
    send(req: Record<string, any>): Promise<Record<string, any>>;
}

/** real's NativeAuthError: raw broker code/description + ext status blob */
export class NativeAuthError extends AuthError {
    name = "NativeAuthError";
    constructor(
        errorCode: string,
        description?: string,
        public ext?: { error?: number; status?: string }
    ) {
        super(errorCode, description || undefined);
    }
}

/** fatal broker errors drop the provider for the rest of the session */
const isFatal = (e: unknown): boolean =>
    e instanceof NativeAuthError &&
    (e.ext?.status === "DISABLED" || e.ext?.error === -2147186943);

/** real's createNativeAuthError: broker ext.status -> msal error class */
const mapError = (e: {
    code: string;
    errorCode: string;
    description: string;
    status?: string;
}): AuthError => {
    switch (e.status) {
        case "ACCOUNT_UNAVAILABLE":
            return new InteractionRequiredAuthError(
                "native_account_unavailable",
                AKA(e.code)
            );
        case "USER_INTERACTION_REQUIRED":
            return new InteractionRequiredAuthError(
                e.code,
                e.description || undefined
            );
        case "USER_CANCEL":
            return new BrowserAuthError("user_cancelled");
        case "NO_NETWORK":
            return new BrowserAuthError("no_network_connectivity");
        case "UI_NOT_ALLOWED":
            return new InteractionRequiredAuthError("ui_not_allowed");
        default:
            return new NativeAuthError(e.code, e.description, {
                error: parseInt(e.errorCode),
                status: e.status,
            });
    }
};

const decodeJwt = (token: string): Record<string, any> =>
    JSON.parse(
        atob(token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))
    );

/** The methods this feature attaches to the client. */
export interface BrokerClient {
    acquireTokenByCode(
        req: Partial<TokenRequest> & { nativeAccountId?: string; code?: string }
    ): Promise<AuthenticationResult>;
}

export function broker(ctx: ClientContext): void {
    const c = ctx.client;
    const config = ctx.config;
    // the winning transport once a probe succeeds
    let provider: Provider | undefined;

    const origInit = c.initialize;
    c.initialize = async () => {
        await origInit();
        // real probes at initialize when allowPlatformBroker is set: DOM
        // handler first (needs the experimental flag too), then the browser
        // extension — preferred extension id, retried with undefined. All
        // probe errors are swallowed — the app falls back to web flows.
        if (!config.system?.allowPlatformBroker) return;
        if (config.experimental?.allowPlatformBrokerWithDOM) {
            try {
                const pa = (navigator as any).platformAuthentication;
                const contracts =
                    await pa?.getSupportedContracts("MicrosoftEntra");
                if (contracts?.includes("get-token-and-sign-out")) {
                    provider = domProvider(pa);
                }
            } catch {
                /* try the extension next, like real */
            }
        }
        if (!provider) {
            try {
                provider = await extensionProvider(PREFERRED_EXTENSION_ID);
            } catch {
                try {
                    provider = await extensionProvider(undefined);
                } catch {
                    /* web-flow fallback, like real */
                }
            }
        }
    };

    /** extension error blob ({code, description, ext:{status,error}}) →
     * the same mapping the DOM error shape goes through */
    const extError = (e: {
        code?: string;
        description?: string;
        ext?: { status?: string; error?: number };
    }) =>
        mapError({
            code: e.code ?? "",
            errorCode: String(e.ext?.error ?? ""),
            description: e.description ?? "",
            status: e.ext?.status,
        });

    // real's PlatformAuthExtensionHandler: Handshake posted to the window
    // with a MessageChannel port; our own Handshake bouncing back on the
    // window means no extension intercepted it (nativeExtensionNotInstalled),
    // silence until the timeout means a handler swallowed it. The port then
    // carries GetToken/Response for the session.
    const extensionProvider = (extensionId?: string): Promise<Provider> =>
        new Promise((resolve, reject) => {
            const mc = new MessageChannel();
            const responseId = crypto.randomUUID();
            const pending = new Map<
                string,
                { resolve: (v: any) => void; reject: (e: unknown) => void }
            >();
            const fail = (code: string) => {
                clearTimeout(timer);
                window.removeEventListener("message", onWindow, false);
                mc.port1.close();
                mc.port2.close();
                reject(new BrowserAuthError(code));
            };
            const onWindow = (event: MessageEvent) => {
                const d = event.data;
                if (
                    event.source === window &&
                    d?.channel === CHANNEL_ID &&
                    (!d.extensionId || d.extensionId === extensionId) &&
                    d.responseId === responseId &&
                    d.body?.method === "Handshake"
                ) {
                    fail("native_extension_not_installed");
                }
            };
            const timer = setTimeout(
                () => fail("native_handshake_timeout"),
                config.system?.nativeBrokerHandshakeTimeout ?? 2000
            );
            mc.port1.onmessage = (m) => {
                const d = m.data;
                if (
                    d?.responseId === responseId &&
                    d.body?.method === "HandshakeResponse"
                ) {
                    clearTimeout(timer);
                    window.removeEventListener("message", onWindow, false);
                    extensionId = d.extensionId;
                    // real's makeExtraSkuString: extension slot filled only
                    // when the handshake reported a version; name "chrome"
                    // for the preferred extension, else "unknown"
                    const name =
                        extensionId === PREFERRED_EXTENSION_ID
                            ? "chrome"
                            : extensionId
                              ? "unknown"
                              : "";
                    const extSku =
                        name && d.body.version
                            ? `${name}|${d.body.version}`
                            : "|";
                    resolve({
                        sku: `msal.js.browser|5.16.0,|,${extSku},|`,
                        send: (req) =>
                            new Promise((res, rej) => {
                                const rid = crypto.randomUUID();
                                pending.set(rid, { resolve: res, reject: rej });
                                mc.port1.postMessage({
                                    channel: CHANNEL_ID,
                                    extensionId,
                                    responseId: rid,
                                    body: { method: "GetToken", request: req },
                                });
                            }),
                    });
                } else if (d?.body?.method === "Response") {
                    const p = pending.get(d.responseId);
                    if (!p) return;
                    pending.delete(d.responseId);
                    const r = d.body.response;
                    const result = r?.result;
                    if (r?.status !== "Success") {
                        p.reject(extError(r ?? {}));
                    } else if (result?.code && result.description) {
                        p.reject(extError(result));
                    } else if (
                        !result ||
                        [
                            "access_token",
                            "id_token",
                            "client_info",
                            "account",
                            "scope",
                            "expires_in",
                        ].some((k) => !(k in result))
                    ) {
                        p.reject(
                            new AuthError(
                                "unexpected_error",
                                "Response missing expected properties."
                            )
                        );
                    } else {
                        // snake_case wire result → handleResponse's shape
                        p.resolve({
                            accessToken: result.access_token,
                            idToken: result.id_token,
                            clientInfo: result.client_info,
                            account: result.account,
                            scopes: result.scope,
                            expiresIn: result.expires_in,
                            state: result.state,
                        });
                    }
                }
            };
            window.addEventListener("message", onWindow, false);
            window.postMessage(
                {
                    channel: CHANNEL_ID,
                    extensionId,
                    responseId,
                    body: { method: "Handshake" },
                },
                window.origin,
                [mc.port2]
            );
        });

    // real's initializePlatformRequest: request minus scopes/claims, plus the
    // broker protocol fields; leftovers become stringified extraParameters at
    // the DOM layer
    const initRequest = (req: Record<string, any>, accountId: string) => {
        const { scopes, claims, ...rest } = req;
        const cfgAuthority = (
            config.auth.authority ?? "https://login.microsoftonline.com/common"
        ).replace(/\/$/, "");
        return {
            ...rest,
            claims: mergedClaims(claims),
            accountId,
            clientId: config.auth.clientId,
            authority: `${(req.authority ?? cfgAuthority).replace(/\/$/, "")}/`,
            scope: [
                ...new Set([
                    ...(scopes ?? []),
                    "openid",
                    "profile",
                    "offline_access",
                ]),
            ].join(" "),
            redirectUri: new URL(
                req.redirectUri ?? config.auth.redirectUri ?? "/",
                location.href
            ).href,
            // silent requests always prompt=none, like real's getPrompt
            prompt: rest.authenticationScheme ? "none" : req.prompt,
            correlationId: req.correlationId,
            tokenType: rest.authenticationScheme,
            windowTitleSubstring: document.title,
            extraParameters: {
                ...req.extraParameters,
                telemetry: "MATS",
                "x-client-xtra-sku": provider!.sku,
            },
            extendedExpiryToken: false,
            keyId: undefined,
        };
    };

    // same shape real's buildMergedClaims produces (value rides the broker
    // request's extraParameters)
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

    // real's PlatformAuthDOMHandler: named protocol fields stay top-level,
    // every remaining truthy property is stringified into extraParameters
    const domProvider = (pa: any): Provider => ({
        sku: "msal.js.browser|5.16.0,|,DOM API|",
        async send(r) {
            const {
                accountId,
                clientId,
                authority,
                scope,
                redirectUri,
                correlationId,
                state,
                storeInCache,
                embeddedClientId,
                extraParameters,
                ...rest
            } = r;
            const extra: Record<string, string> = { ...extraParameters };
            for (const [k, v] of Object.entries(rest)) {
                if (v) {
                    extra[k] =
                        typeof v === "object" ? JSON.stringify(v) : String(v);
                }
            }
            const response = await pa.executeGetToken({
                accountId,
                brokerId: "MicrosoftEntra",
                authority,
                clientId,
                correlationId,
                extraParameters: extra,
                isSecurityTokenService: false,
                redirectUri,
                scope,
                state,
                storeInCache,
                embeddedClientId,
            });
            if (response.isSuccess === false && response.error?.code) {
                throw mapError(response.error);
            }
            if (
                ["accessToken", "idToken", "clientInfo", "account", "scopes",
                    "expiresIn"].some((k) => !(k in response))
            ) {
                throw new AuthError(
                    "unexpected_error",
                    "Response missing expected properties."
                );
            }
            return response;
        },
    });

    // real's handleNativeResponse: cache the account (with nativeAccountId),
    // build the broker-shaped 14-key AuthenticationResult
    const handleResponse = async (
        response: Record<string, any>,
        request: Record<string, any>,
        apiId: number
    ): Promise<AuthenticationResult> => {
        if (response.account.id !== request.accountId) {
            // user switch in the broker prompt is not supported
            throw new NativeAuthError("user_switch");
        }
        const claims = decodeJwt(response.idToken);
        let uid = claims.oid ?? claims.sub;
        let utid = claims.tid ?? "";
        try {
            const ci = JSON.parse(
                atob(
                    response.clientInfo.replace(/-/g, "+").replace(/_/g, "/")
                )
            );
            uid = ci.uid;
            utid = ci.utid;
        } catch {
            /* fall back to claims */
        }
        const homeAccountId = `${uid}.${utid}`;
        const host = new URL(request.authority).host;
        const environment =
            /login\.microsoftonline\.com|login\.microsoft\.com|sts\.windows\.net/.test(
                host
            )
                ? "login.windows.net"
                : host;
        const realm = claims.tid ?? "";
        const localAccountId = claims.oid ?? claims.sub;
        const username = claims.preferred_username ?? claims.email ?? "";
        const tenantProfiles = [
            {
                tenantId: realm,
                localAccountId,
                name: claims.name,
                username,
                isHomeTenant: realm === homeAccountId.split(".")[1],
            },
        ];
        const entity: AccountEntity = {
            homeAccountId,
            environment,
            realm,
            localAccountId,
            username,
            authorityType: "MSSTS",
            name: claims.name,
            clientInfo: response.clientInfo,
            nativeAccountId: response.account.id,
            tenantProfiles,
            lastUpdatedAt: String(Date.now()),
            cachedByApiId: apiId,
        };
        await ctx.writeAccount(entity);
        const props = response.account.properties || {};
        const account: AccountInfo = {
            authorityType: "MSSTS",
            dataBoundary: undefined,
            environment,
            homeAccountId,
            idToken: response.idToken,
            idTokenClaims: claims,
            kmsi: isKmsi(claims),
            localAccountId,
            loginHint: claims.login_hint,
            name: claims.name,
            nativeAccountId: response.account.id,
            tenantId: realm,
            tenantProfiles: new Map(
                tenantProfiles.map((p) => [p.tenantId, p])
            ),
            upn: claims.upn,
            username,
        };
        return {
            authority: request.authority,
            uniqueId: props.UID || claims.oid || claims.sub || "",
            tenantId: props.TenantId || claims.tid || "",
            scopes: (response.scopes || request.scope).split(" "),
            account,
            idToken: response.idToken,
            idTokenClaims: claims,
            accessToken: response.accessToken,
            fromCache: false,
            expiresOn: new Date(
                (Math.floor(Date.now() / 1000) + Number(response.expiresIn)) *
                    1000
            ),
            tokenType: "Bearer",
            correlationId: request.correlationId,
            state: response.state || "",
            fromPlatformBroker: true,
        } as AuthenticationResult;
    };

    const acquireNative = async (
        req: Record<string, any>,
        accountId: string,
        apiId: number
    ): Promise<AuthenticationResult> => {
        const nativeReq = initRequest(req, accountId);
        try {
            const response = await provider!.send(nativeReq);
            return await handleResponse(response, nativeReq, apiId);
        } catch (e) {
            if (isFatal(e)) {
                // e.g. broker disabled by policy: drop it for the session
                provider = undefined;
            }
            throw e;
        }
    };

    c.acquireTokenByCode = async (
        req: Partial<TokenRequest> & { nativeAccountId?: string }
    ): Promise<AuthenticationResult> => {
        ctx.preflight();
        const correlationId = req.correlationId ?? crypto.randomUUID();
        // real emits start with the request as passed (no correlationId
        // stamp) and, on the native path, NO success event
        ctx.emit(EventType.ACQUIRE_TOKEN_START, "silent", req);
        try {
            if (!req.nativeAccountId) {
                throw new BrowserAuthError(
                    "auth_code_or_nativeAccountId_required"
                );
            }
            if (!provider) {
                throw new BrowserAuthError(
                    "unable_to_acquire_token_from_native_platform"
                );
            }
            return await acquireNative(
                { ...req, correlationId },
                req.nativeAccountId,
                866 // ApiId.acquireTokenByCode
            );
        } catch (e) {
            ctx.emit(EventType.ACQUIRE_TOKEN_FAILURE, "silent", undefined, e);
            throw e;
        }
    };

    // silent requests route through the broker when the account is
    // broker-joined (real's acquireTokenSilentNoIframe)
    ctx.nativeSilent = (req: TokenRequest, account: AccountInfo) =>
        provider && account.nativeAccountId
            ? acquireNative(
                  {
                      ...req,
                      account,
                      authenticationScheme: "Bearer",
                      forceRefresh: !!req.forceRefresh,
                  },
                  account.nativeAccountId,
                  61 // ApiId.acquireTokenSilent_silentFlow
              ).then((r) => {
                  // silent results carry state: undefined (key present)
                  r.state = undefined;
                  return r;
              })
            : undefined;
}
