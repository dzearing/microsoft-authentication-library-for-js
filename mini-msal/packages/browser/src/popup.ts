/**
 * Popup feature: loginPopup / acquireTokenPopup / logoutPopup, built on the
 * core's authorize-URL plumbing and redirect-bridge wait (the popup must land
 * on a page running ./redirect-bridge, like real v5). Compose via
 * createClient(config, [popup]) — apps that never open popups don't pay for
 * this module.
 */
import {
    BrowserAuthError,
    EventType,
    type AccountInfo,
    type AuthenticationResult,
    type ClientContext,
    type LogoutRequest,
    type PopupWindowAttributes,
    type TokenRequest,
} from "./index.js";

/** logoutPopup request: core LogoutRequest + popup-window controls */
export interface LogoutPopupRequest extends LogoutRequest {
    mainWindowRedirectUri?: string;
    popupWindowAttributes?: PopupWindowAttributes;
    popupWindowParent?: Window;
}

/** The methods this feature attaches to the client. */
export interface PopupClient {
    loginPopup(req: TokenRequest): Promise<AuthenticationResult>;
    acquireTokenPopup(req: TokenRequest): Promise<AuthenticationResult>;
    logoutPopup(req?: LogoutPopupRequest): Promise<void>;
}

const POPUP_W = 483;
const POPUP_H = 600;

interface PopupParams {
    name: string;
    attrs: PopupWindowAttributes;
    parent: Window;
    /** pre-opened about:blank popup (navigatePopups); null = open was
     * blocked, undefined = deferred open */
    popup?: Window | null;
}

/** real's openSizedPopup: request attrs clamped to the parent window,
 * centered defaults (identical features-string format, incl. spaces) */
function openSized(url: string, p: PopupParams): Window | null {
    const parent = p.parent;
    const winLeft = parent.screenLeft ? parent.screenLeft : parent.screenX;
    const winTop = parent.screenTop ? parent.screenTop : parent.screenY;
    const winWidth =
        parent.innerWidth ||
        document.documentElement.clientWidth ||
        document.body.clientWidth;
    const winHeight =
        parent.innerHeight ||
        document.documentElement.clientHeight ||
        document.body.clientHeight;
    let { width, height } = p.attrs.popupSize ?? {};
    let { top, left } = p.attrs.popupPosition ?? {};
    if (!width || width < 0 || width > winWidth) {
        width = POPUP_W;
    }
    if (!height || height < 0 || height > winHeight) {
        height = POPUP_H;
    }
    if (!top || top < 0 || top > winHeight) {
        top = Math.max(0, winHeight / 2 - POPUP_H / 2 + winTop);
    }
    if (!left || left < 0 || left > winWidth) {
        left = Math.max(0, winWidth / 2 - POPUP_W / 2 + winLeft);
    }
    return parent.open(
        url,
        p.name,
        `width=${width}, height=${height}, top=${top}, left=${left}, scrollbars=yes`
    );
}

/** real's openPopup: navigate the pre-opened popup or open one now; a null
 * pre-open (blocked) surfaces as popup_window_error, like real */
function navigatePopup(url: string, p: PopupParams): Window {
    try {
        let win = p.popup;
        if (win) {
            win.location.assign(url);
        } else if (win === undefined) {
            win = openSized(url, p);
        }
        if (!win) {
            throw new Error();
        }
        try {
            win.document.title = "Microsoft Authentication";
        } catch {
            /* cross-origin: title cannot be set */
        }
        win.focus?.();
        return win;
    } catch {
        throw new BrowserAuthError("popup_window_error");
    }
}

export function popup(ctx: ClientContext): void {
    const c = ctx.client;
    const clientId = ctx.config.auth.clientId;
    // real's default navigatePopups=true: open about:blank synchronously in
    // the caller's stack so popup blockers see the user gesture
    const syncOpen = ctx.config.system?.navigatePopups !== false;

    c.acquireTokenPopup = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        // like real: preflight/lock failures reject BEFORE any event fires
        ctx.preflight();
        ctx.lock();
        const had = c.getAllAccounts().length;
        ctx.emit(EventType.ACQUIRE_TOKEN_START, "popup", req);
        const cid = req.correlationId ?? crypto.randomUUID();
        const p: PopupParams = {
            // real's generatePopupName
            name: `msal.${clientId}.${(
                req.scopes ?? ["openid", "profile", "offline_access"]
            ).join("-")}.${req.authority ?? ctx.config.auth.authority ?? ""}.${cid}`,
            attrs: req.popupWindowAttributes ?? {},
            parent: req.popupWindowParent ?? window,
        };
        if (syncOpen) {
            p.popup = openSized("about:blank", p);
        }
        try {
            const { url, verifier, state, redirectUri, correlationId, nonce, ccs } =
                await ctx.authorizeUrl({ ...req, correlationId: cid }, "popup");
            const win = navigatePopup(url, p);
            ctx.emit(EventType.POPUP_OPENED, "popup", { popupWindow: win });
            try {
                const auth = await ctx.waitForCode(
                    state,
                    ctx.config.system?.popupBridgeTimeout ?? 60_000
                );
                const result = await ctx.redeem({
                    ...auth,
                    verifier,
                    scopes: req.scopes,
                    redirectUri,
                    correlationId,
                    nonce,
                    ccs,
                    apiId: 862, // ApiId.acquireTokenPopup
                    userState: req.state,
                    claims: req.claims,
                    eqp: req.extraQueryParameters,
                    authority: req.authority,
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
            // real closes the synchronously-opened popup on failure
            p.popup?.close();
            ctx.stFail(862, cid, e); // real PopupClient's failure cache
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

    c.logoutPopup = async (req?: LogoutPopupRequest): Promise<void> => {
        ctx.preflight();
        ctx.lock("signout");
        const validRequest: LogoutPopupRequest & {
            correlationId: string;
            state?: string;
        } = {
            correlationId: crypto.randomUUID(),
            postLogoutRedirectUri: ctx.config.auth.postLogoutRedirectUri,
            ...req,
        };
        const p: PopupParams = {
            // real's generateLogoutPopupName
            name: `msal.${clientId}.${
                req?.account && req.account.homeAccountId
            }.${validRequest.correlationId}`,
            attrs: req?.popupWindowAttributes ?? {},
            parent: req?.popupWindowParent ?? window,
        };
        if (syncOpen) {
            p.popup = openSized("about:blank", p);
        }
        ctx.emit(EventType.LOGOUT_START, "popup", validRequest);
        try {
            // real clears the cache and emits logoutSuccess BEFORE the popup
            // opens; the popup roundtrip only clears the server session
            ctx.clearAccount(req?.account);
            validRequest.state = crypto.randomUUID();
            ctx.emit(EventType.LOGOUT_SUCCESS, "popup", validRequest);
            const win = navigatePopup(
                await ctx.logoutUrl(validRequest, "popup"),
                p
            );
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
            if (req?.mainWindowRedirectUri) {
                // real navigates the MAIN window after the popup roundtrip,
                // through the NavigationClient seam; the page unloads before
                // the navigation promise settles, so logoutEnd/unlock never
                // run (the lock survives into the next page, like real)
                await ctx.navigate(
                    new URL(req.mainWindowRedirectUri, location.href).href,
                    962 // ApiId.logoutPopup
                );
            }
        } catch (e) {
            // real closes the synchronously-opened popup on failure
            p.popup?.close();
            ctx.emit(EventType.LOGOUT_FAILURE, "popup", undefined, e);
            throw e;
        } finally {
            ctx.emit(EventType.LOGOUT_END, "popup");
            ctx.unlock();
        }
    };
}
