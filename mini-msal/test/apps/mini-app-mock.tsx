// GENERATED from app.tsx by sync-variants.mjs — do not edit
/**
 * Full-authentication-flow usage of @azure/msal-react + @azure/msal-browser.
 * Deliberately references the public API surface a real production app uses,
 * so tree shaking cannot drop code paths a real consumer would ship:
 *   - redirect, popup, and ssoSilent login
 *   - acquireTokenSilent (with CacheLookupPolicy) + popup/redirect fallbacks
 *   - event callbacks (add/remove), account storage events, active account
 *   - account lookup by homeAccountId/username
 *   - logoutRedirect and logoutPopup
 *   - MsalProvider, Authenticated/Unauthenticated/MsalAuthenticationTemplate,
 *     useMsal, useAccount, useIsAuthenticated, useMsalAuthentication, withMsal
 */
import {
    PublicClientApplication,
    EventType,
    InteractionType,
    CacheLookupPolicy,
    InteractionRequiredAuthError,
    BrowserAuthError,
    type AccountInfo,
    type EventMessage,
    type AuthenticationResult,
} from "@mini-msal/compat";
import {
    MsalProvider,
    AuthenticatedTemplate,
    UnauthenticatedTemplate,
    MsalAuthenticationTemplate,
    useMsal,
    useAccount,
    useIsAuthenticated,
    useMsalAuthentication,
    withMsal,
    type IMsalContext,
} from "@mini-msal/react";
import * as React from "react";
import { createRoot } from "react-dom/client";
import { msalConfig, loginRequest } from "./authConfig.mock.js";

const msalInstance = new PublicClientApplication(msalConfig);

/**
 * v5 guidance: popup and hidden-iframe flows redirect to a dedicated page
 * (served at /popup.html, running the msal redirect-bridge) rather than the
 * app itself. For real AAD, register http://localhost:4173/popup.html (and
 * /mini-popup.html) as additional SPA redirect URIs. sync-variants rewrites
 * this to /mini-popup.html (mini's own bridge page) for the mini stack.
 */
const popupRequest = { ...loginRequest, redirectUri: "/mini-popup.html" };

function LoginButtons() {
    const { instance, inProgress } = useMsal();
    const busy = inProgress !== "none";
    return (
        <>
            <button
                disabled={busy}
                onClick={() => instance.loginRedirect(loginRequest)}
            >
                Sign in (redirect)
            </button>
            <button
                disabled={busy}
                onClick={() =>
                    instance
                        .loginPopup(popupRequest)
                        .then((result) =>
                            instance.setActiveAccount(result.account)
                        )
                        .catch((e) => {
                            if (
                                e instanceof BrowserAuthError ||
                                e instanceof InteractionRequiredAuthError
                            ) {
                                console.warn(e.errorCode);
                            }
                        })
                }
            >
                Sign in (popup)
            </button>
            <button
                disabled={busy}
                onClick={() =>
                    instance
                        .ssoSilent({
                            ...popupRequest,
                            loginHint: "user@contoso.com",
                        })
                        .catch(() => instance.loginRedirect(loginRequest))
                }
            >
                Try SSO
            </button>
        </>
    );
}

function LogoutButtons() {
    const { instance } = useMsal();
    const account = instance.getActiveAccount();
    return (
        <>
            <button onClick={() => instance.logoutRedirect({ account })}>
                Sign out (redirect)
            </button>
            <button onClick={() => instance.logoutPopup({ account })}>
                Sign out (popup)
            </button>
        </>
    );
}

function acquireWithFallback(
    instance: IMsalContext["instance"],
    account: AccountInfo
): Promise<AuthenticationResult | null> {
    return instance
        .acquireTokenSilent({
            ...loginRequest,
            account,
            cacheLookupPolicy: CacheLookupPolicy.Default,
        })
        .catch((e) => {
            if (e instanceof InteractionRequiredAuthError) {
                return instance
                    .acquireTokenPopup(popupRequest)
                    .catch((popupError) => {
                        if (popupError instanceof BrowserAuthError) {
                            // popup blocked -> fall back to full redirect
                            return instance
                                .acquireTokenRedirect(loginRequest)
                                .then(() => null);
                        }
                        throw popupError;
                    });
            }
            throw e;
        });
}

function Profile() {
    const { instance } = useMsal();
    const account = useAccount(); // no filter -> the active account

    const [data, setData] = React.useState<string | null>(null);

    React.useEffect(() => {
        if (!account) {
            return;
        }
        acquireWithFallback(instance, account)
            .then((result) => {
                if (result) {
                    return fetch("https://graph.microsoft.com/v1.0/me", {
                        headers: {
                            Authorization: `Bearer ${result.accessToken}`,
                        },
                    })
                        .then((r) => r.json())
                        .then((json) =>
                            setData(json.displayName ?? "unknown")
                        );
                }
            })
            .catch(console.error);
    }, [instance, account]);

    return (
        <div>
            Hello {data ?? account?.username} ({account?.username})
        </div>
    );
}

function AccountSwitcher() {
    const { instance, accounts } = useMsal();
    const active = useAccount(); // subscribes to ACTIVE_ACCOUNT_CHANGED
    return (
        <div>
            {accounts.map((a) => (
                <button
                    key={a.homeAccountId}
                    onClick={() => instance.setActiveAccount(a)}
                >
                    {(active?.homeAccountId === a.homeAccountId ? "\u25cf " : "") +
                        a.username}
                </button>
            ))}
            <button
                onClick={() =>
                    instance
                        .loginPopup({
                            ...popupRequest,
                            prompt: "select_account",
                        })
                        .then((result) =>
                            instance.setActiveAccount(result.account)
                        )
                        .catch(console.warn)
                }
            >
                Add account
            </button>
        </div>
    );
}

// exercises useMsalAuthentication: auto-attempt a silent login, expose
// manual login + error to the UI
function AutoLogin() {
    const { login, error } = useMsalAuthentication(
        InteractionType.Silent,
        popupRequest
    );
    return error ? (
        <button onClick={() => login(InteractionType.Popup, popupRequest)}>
            Retry sign in
        </button>
    ) : null;
}

// exercises withMsal (class-component consumption)
class StatusBanner extends React.Component<{ msalContext: IMsalContext }> {
    render() {
        const { accounts, inProgress } = this.props.msalContext;
        return (
            <small>
                {accounts.length} account(s), interaction: {inProgress}
            </small>
        );
    }
}
const WrappedStatusBanner = withMsal(StatusBanner);

function ErrorScreen(props: { error: unknown }) {
    return <p>Authentication failed: {String(props.error)}</p>;
}

function LoadingScreen() {
    return <p>Signing in…</p>;
}

function App() {
    return (
        <MsalProvider instance={msalInstance}>
            <WrappedStatusBanner />
            <AuthenticatedTemplate>
                <Profile />
                <AccountSwitcher />
                <LogoutButtons />
            </AuthenticatedTemplate>
            <UnauthenticatedTemplate>
                <p>Please sign in.</p>
                <LoginButtons />
                <AutoLogin />
            </UnauthenticatedTemplate>
            <MsalAuthenticationTemplate
                interactionType={InteractionType.Redirect}
                authenticationRequest={loginRequest}
                errorComponent={ErrorScreen}
                loadingComponent={LoadingScreen}
            >
                <p>Protected content</p>
            </MsalAuthenticationTemplate>
        </MsalProvider>
    );
}

msalInstance.initialize().then(() => {
    // exposed for the E2E test driver (e2e.mjs)
    (globalThis as Record<string, unknown>).__msal = msalInstance;
    // full-flow wiring: event callbacks driving active-account bookkeeping
    const callbackId = msalInstance.addEventCallback((message: EventMessage) => {
        if (message.eventType === EventType.LOGIN_SUCCESS) {
            // payload may lack an account on some popup/bridge paths
            const payload = message.payload as AuthenticationResult | null;
            const account =
                (payload?.account &&
                    msalInstance.getAccount({
                        homeAccountId: payload.account.homeAccountId,
                    })) ||
                (msalInstance.getActiveAccount()
                    ? null // already have one; leave it
                    : msalInstance.getAllAccounts()[0]);
            if (account) {
                msalInstance.setActiveAccount(account);
            }
        } else if (message.eventType === EventType.ACCOUNT_REMOVED) {
            console.log("account removed in another tab");
        }
    });
    window.addEventListener("unload", () => {
        if (callbackId) {
            msalInstance.removeEventCallback(callbackId);
        }
    });
    createRoot(document.getElementById("root")!).render(<App />);
});
