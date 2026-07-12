// GENERATED from msal-only.ts by sync-variants.mjs — do not edit
/**
 * Full auth flow with no React at all — isolates the cost of
 * @azure/msal-browser + @azure/msal-common. References the same public API
 * surface as app.tsx so tree shaking cannot drop real-world code paths.
 */
import {
    PublicClientApplication,
    EventType,
    CacheLookupPolicy,
    InteractionRequiredAuthError,
    BrowserAuthError,
    type EventMessage,
    type AuthenticationResult,
    type AccountInfo,
} from "@mini-msal/compat";
import { msalConfig, loginRequest } from "./authConfig.js";

const msalInstance = new PublicClientApplication(msalConfig);

async function getToken(account: AccountInfo): Promise<string | null> {
    try {
        const result = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account,
            cacheLookupPolicy: CacheLookupPolicy.Default,
        });
        return result.accessToken;
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            try {
                const result = await msalInstance.acquireTokenPopup(
                    loginRequest
                );
                return result.accessToken;
            } catch (popupError) {
                if (popupError instanceof BrowserAuthError) {
                    await msalInstance.acquireTokenRedirect(loginRequest);
                    return null;
                }
                throw popupError;
            }
        }
        throw e;
    }
}

async function main() {
    await msalInstance.initialize();
    const callbackId = msalInstance.addEventCallback(
        (message: EventMessage) => {
            if (message.eventType === EventType.LOGIN_SUCCESS) {
                const payload = message.payload as AuthenticationResult;
                const account =
                    msalInstance.getAccount({
                        homeAccountId: payload.account.homeAccountId,
                    }) ??
                    msalInstance.getAccount({
                        username: payload.account.username,
                    });
                msalInstance.setActiveAccount(account);
            }
        }
    );
    window.addEventListener("unload", () => {
        if (callbackId) {
            msalInstance.removeEventCallback(callbackId);
        }
    });

    const response = await msalInstance.handleRedirectPromise();
    if (response) {
        console.log("logged in", response.account.username);
    }

    let account = msalInstance.getActiveAccount();
    if (!account) {
        [account] = msalInstance.getAllAccounts();
    }
    if (!account) {
        // try silent SSO first, then popup, then full redirect
        try {
            const result = await msalInstance.ssoSilent({
                ...loginRequest,
                loginHint: "user@contoso.com",
            });
            account = result.account;
        } catch {
            try {
                const result = await msalInstance.loginPopup(loginRequest);
                account = result.account;
            } catch {
                await msalInstance.loginRedirect(loginRequest);
                return;
            }
        }
        msalInstance.setActiveAccount(account);
    }

    const token = await getToken(account);
    console.log("token length", token?.length);

    document.getElementById("logout")?.addEventListener("click", () => {
        msalInstance
            .logoutPopup({ account })
            .catch(() => msalInstance.logoutRedirect({ account }));
    });
}

main().catch(console.error);
