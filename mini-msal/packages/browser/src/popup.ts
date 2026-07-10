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
    logoutPopup(req?: { account?: AccountInfo | null }): Promise<void>;
}

const FEATURES = "width=483,height=600,popup=yes";

function openPopup(url: string): Window {
    const win = open(url, "msal.popup", FEATURES);
    if (!win) {
        throw new BrowserAuthError("popup_window_error", "Popup was blocked");
    }
    return win;
}

export function popup(ctx: ClientContext): void {
    const c = ctx.client;

    c.acquireTokenPopup = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        ctx.preflight();
        const { url, verifier, state, redirectUri } = await ctx.authorizeUrl(
            req
        );
        const win = openPopup(url);
        try {
            const code = await ctx.pollForCode(win, state, 60_000);
            return await ctx.redeem({
                code,
                verifier,
                scopes: req.scopes,
                redirectUri,
            });
        } finally {
            win.close();
        }
    };

    c.loginPopup = async (
        req: TokenRequest
    ): Promise<AuthenticationResult> => {
        try {
            const result = await c.acquireTokenPopup(req);
            ctx.emit(EventType.LOGIN_SUCCESS, result);
            return result;
        } catch (e) {
            ctx.emit(EventType.LOGIN_FAILURE, undefined, e);
            throw e;
        }
    };

    c.logoutPopup = async (req?: {
        account?: AccountInfo | null;
    }): Promise<void> => {
        ctx.preflight();
        const win = openPopup(ctx.logoutUrl());
        // wait for the popup to land back on the post-logout page (server
        // session cleared), then close; events fire only after completion,
        // matching real MSAL's ordering
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
        ctx.clearAccount(req?.account);
    };
}
