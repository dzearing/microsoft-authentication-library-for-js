/**
 * Redirect-only SPA on @mini-msal/browser core — the pay-to-play floor.
 * No feature modules composed: sign-in, silent acquisition, and sign-out
 * all ride full-page redirects, so no redirect-bridge page is needed.
 */
import {
    createClient,
    InteractionRequiredAuthError,
} from "@mini-msal/browser";
import { msalConfig, loginRequest } from "./authConfig.js";

const client = createClient(msalConfig);
const root = document.getElementById("root")!;

async function main() {
    await client.initialize();

    // completes the roundtrip when the IdP redirects back to this page
    const response = await client.handleRedirectPromise();
    if (response) {
        client.setActiveAccount(response.account);
    }

    const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
    if (!account) {
        root.innerHTML = `<button id="signin">Sign in</button>`;
        document.getElementById("signin")!.onclick = () => {
            client.loginRedirect(loginRequest).catch(console.error);
        };
        return;
    }

    let accessToken: string;
    try {
        ({ accessToken } = await client.acquireTokenSilent({
            ...loginRequest,
            account,
        }));
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            await client.acquireTokenRedirect({ ...loginRequest, account });
            return;
        }
        throw e;
    }

    root.innerHTML =
        `<p>Hello ${account.username}</p>` +
        `<p>token: ${accessToken.length} chars</p>` +
        `<button id="signout">Sign out</button>`;
    document.getElementById("signout")!.onclick = () => {
        client.logoutRedirect({ account }).catch(console.error);
    };
}

main().catch((e) => {
    root.textContent = String(e);
});
