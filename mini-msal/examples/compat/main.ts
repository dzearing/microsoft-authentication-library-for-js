/**
 * Drop-in usage of @mini-msal/compat: the same PublicClientApplication
 * surface as @azure/msal-browser with every feature composed. Migrating a
 * real msal-browser app is two steps — swap the import, and serve mini's
 * redirect-bridge on your popup page (see README.md and docs/UPGRADING.md).
 */
import {
    PublicClientApplication,
    InteractionRequiredAuthError,
} from "@mini-msal/compat";
import { msalConfig, loginRequest } from "./authConfig.js";

const pca = new PublicClientApplication(msalConfig);
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
    const account = pca.getActiveAccount() ?? pca.getAllAccounts()[0];
    if (!account) {
        root.innerHTML = `<button id="signin">Sign in</button>`;
        document.getElementById("signin")!.onclick = () => {
            pca.loginPopup(popupRequest)
                .then((result) => {
                    pca.setActiveAccount(result.account);
                    return render();
                })
                .catch(showError);
        };
        return;
    }

    let accessToken: string;
    try {
        ({ accessToken } = await pca.acquireTokenSilent({
            ...loginRequest,
            account,
        }));
    } catch (e) {
        if (!(e instanceof InteractionRequiredAuthError)) throw e;
        ({ accessToken } = await pca.acquireTokenPopup(popupRequest));
    }

    root.innerHTML =
        `<p>Hello ${account.username}</p>` +
        `<p>token: ${accessToken.length} chars</p>` +
        `<button id="signout">Sign out</button>`;
    document.getElementById("signout")!.onclick = () => {
        pca.logoutPopup({ account }).then(render).catch(showError);
    };
}

pca.initialize()
    // completes a redirect roundtrip if one is pending (harmless otherwise)
    .then(() => pca.handleRedirectPromise())
    .then(render)
    .catch(showError);
