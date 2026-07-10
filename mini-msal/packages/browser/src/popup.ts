/**
 * Popup feature: loginPopup / acquireTokenPopup / logoutPopup, built on the
 * core's authorize-URL plumbing and window polling. Compose via
 * createClient(config, [popup]) — apps that never open popups don't pay for
 * this module.
 */
import {
    BrowserAuthError,
    EventType,
    type AccountInfo,
    type AuthenticationResult,
    type ClientContext,
    type TokenRequest,
} from "./index.js";

/** The methods this feature attaches to the client. */
export interface PopupClient {
    loginPopup(req: TokenRequest): Promise<AuthenticationResult>;
    acquireTokenPopup(req: TokenRequest): Promise<AuthenticationResult>;
    logoutPopup(req?: {
        account?: AccountInfo | null;
        postLogoutRedirectUri?: string;
        correlationId?: string;
    }): Promise<void>;
}

const FEATURES = "width=483,height=600,popup=yes";

function openPopup(url: string): Window {
    // unique per request, like real's generatePopupName — concurrent
    // popups must never clobber each other's window
    const win = open(url, `msal.${crypto.randomUUID()}`, FEATURES);
    if (!win) {
        throw new BrowserAuthError("popup_window_error");
    }
    return win;
}

export function popup(ctx: ClientContext): void {
    const c = ctx.client;

    c.acquireTokenPopup = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        // like real: preflight/lock failures reject BEFORE any event fires
        ctx.preflight();
        ctx.lock();
        const had = c.getAllAccounts().length;
        ctx.emit(EventType.ACQUIRE_TOKEN_START, "popup", req);
        try {
            const { url, verifier, state, redirectUri, correlationId, nonce, ccs } =
                await ctx.authorizeUrl(req);
            const win = openPopup(url);
            ctx.emit(EventType.POPUP_OPENED, "popup", { popupWindow: win });
            try {
                const code = await ctx.pollForCode(
                    win,
                    state,
                    ctx.config.system?.popupBridgeTimeout ?? 60_000
                );
                const result = await ctx.redeem({
                    code,
                    verifier,
                    scopes: req.scopes,
                    redirectUri,
                    correlationId,
                    nonce,
                    ccs,
                    apiId: 862, // ApiId.acquireTokenPopup
                });
                ctx.emit(EventType.ACQUIRE_TOKEN_SUCCESS, "popup", result);
                if (had < c.getAllAccounts().length) {
                    // loginSuccess carries the account, not the result
                    ctx.emit(EventType.LOGIN_SUCCESS, "popup", result.account);
                }
                return result;
            } finally {
                win.close();
            }
        } catch (e) {
            ctx.emit(EventType.ACQUIRE_TOKEN_FAILURE, "popup", undefined, e);
            throw e;
        } finally {
            ctx.unlock();
        }
    };

    // real's loginPopup is just acquireTokenPopup with a correlationId
    // stamped on the request (visible in the acquireTokenStart payload);
    // the login vs acquire event split happens on account-count change
    c.loginPopup = (req: TokenRequest): Promise<AuthenticationResult> =>
        c.acquireTokenPopup({ correlationId: crypto.randomUUID(), ...req });

    c.logoutPopup = async (req?: {
        account?: AccountInfo | null;
        postLogoutRedirectUri?: string;
        correlationId?: string;
    }): Promise<void> => {
        ctx.preflight();
        ctx.lock("signout");
        const validRequest: {
            correlationId: string;
            postLogoutRedirectUri?: string;
            account?: AccountInfo | null;
            state?: string;
        } = {
            correlationId: crypto.randomUUID(),
            postLogoutRedirectUri: ctx.config.auth.postLogoutRedirectUri,
            ...req,
        };
        ctx.emit(EventType.LOGOUT_START, "popup", validRequest);
        try {
            // real clears the cache and emits logoutSuccess BEFORE the popup
            // opens; the popup roundtrip only clears the server session
            ctx.clearAccount(req?.account);
            validRequest.state = crypto.randomUUID();
            ctx.emit(EventType.LOGOUT_SUCCESS, "popup", validRequest);
            const win = openPopup(ctx.logoutUrl(validRequest, "popup"));
            ctx.emit(EventType.POPUP_OPENED, "popup", { popupWindow: win });
            // wait for the popup to land back on the post-logout page (server
            // session cleared), then close
            const started = Date.now();
            await new Promise<void>((resolve) => {
                const timer = setInterval(() => {
                    let done = win.closed || Date.now() - started > 5000;
                    try {
                        done ||= win.location.origin === location.origin;
                    } catch {
                        /* still on the IdP: keep waiting */
                    }
                    if (done) {
                        clearInterval(timer);
                        resolve();
                    }
                }, 50);
            });
            win.close();
        } catch (e) {
            ctx.emit(EventType.LOGOUT_FAILURE, "popup", undefined, e);
            throw e;
        } finally {
            ctx.emit(EventType.LOGOUT_END, "popup");
            ctx.unlock();
        }
    };
}
