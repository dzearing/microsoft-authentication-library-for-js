/**
 * Popup SPA: @mini-msal/browser core + the popup feature (loginPopup /
 * acquireTokenPopup / logoutPopup). Popups complete on a dedicated
 * redirect-bridge page served next to the app — see README.md (popup.html
 * runs @mini-msal/browser/redirect-bridge).
 */
import {
    createClient,
    InteractionRequiredAuthError,
    type AuthClient,
} from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";
import { msalConfig, loginRequest } from "./authConfig.js";

const client = createClient(msalConfig, [popup]) as AuthClient & PopupClient;
const root = document.getElementById("root")!;

// popups redirect to the bridge page next to this app's index.html
const popupRequest = {
    ...loginRequest,
    redirectUri: new URL("popup.html", window.location.href).href,
};

const showError = (e: unknown) => {
    root.textContent = String(e);
};

async function render() {
    const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
    if (!account) {
        root.innerHTML = `<button id="signin">Sign in</button>`;
        document.getElementById("signin")!.onclick = () => {
            client
                .loginPopup(popupRequest)
                .then((result) => {
                    client.setActiveAccount(result.account);
                    return render();
                })
                .catch(showError);
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
        if (!(e instanceof InteractionRequiredAuthError)) throw e;
        ({ accessToken } = await client.acquireTokenPopup(popupRequest));
    }

    root.innerHTML =
        `<p>Hello ${account.username}</p>` +
        `<p>token: ${accessToken.length} chars</p>` +
        `<button id="signout">Sign out</button>`;
    document.getElementById("signout")!.onclick = () => {
        client.logoutPopup({ account }).then(render).catch(showError);
    };
}

client.initialize().then(render).catch(showError);
