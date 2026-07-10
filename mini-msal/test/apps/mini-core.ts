/**
 * Core-only mini-msal usage: redirect + silent flows, accounts, events —
 * no popup feature, no React, no compat layer. This is the pay-to-play
 * floor: what an app pays for createClient alone. Mirrors the API pressure
 * of msal-only.ts minus the popup paths.
 */
import {
    createClient,
    EventType,
    CacheLookupPolicy,
    InteractionRequiredAuthError,
    type EventMessage,
    type AuthenticationResult,
} from "@mini-msal/browser";
import { msalConfig, loginRequest } from "./authConfig.js";

const client = createClient(msalConfig);

async function main() {
    await client.initialize();
    const callbackId = client.addEventCallback((message: EventMessage) => {
        if (message.eventType === EventType.LOGIN_SUCCESS) {
            const payload = message.payload as AuthenticationResult;
            const account =
                client.getAccount({
                    homeAccountId: payload.account.homeAccountId,
                }) ??
                client.getAccount({ username: payload.account.username });
            client.setActiveAccount(account);
        }
    });
    window.addEventListener("unload", () => {
        if (callbackId) {
            client.removeEventCallback(callbackId);
        }
    });

    const response = await client.handleRedirectPromise();
    if (response) {
        console.log("logged in", response.account.username);
    }

    let account = client.getActiveAccount();
    if (!account) {
        [account] = client.getAllAccounts();
    }
    if (!account) {
        // try silent SSO first, then full redirect
        try {
            const result = await client.ssoSilent({
                ...loginRequest,
                loginHint: "user@contoso.com",
            });
            account = result.account;
            client.setActiveAccount(account);
        } catch {
            await client.loginRedirect(loginRequest);
            return;
        }
    }

    try {
        const result = await client.acquireTokenSilent({
            ...loginRequest,
            account,
            cacheLookupPolicy: CacheLookupPolicy.Default,
        });
        console.log("token length", result.accessToken.length);
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            await client.acquireTokenRedirect(loginRequest);
            return;
        }
        throw e;
    }

    document.getElementById("logout")?.addEventListener("click", () => {
        client.logoutRedirect({ account });
    });
}

main().catch(console.error);
