/**
 * mini-msal React bindings, mirroring the @azure/msal-react surface the
 * full-flow app uses: MsalProvider, useMsal, useAccount, useIsAuthenticated,
 * useMsalAuthentication, AuthenticatedTemplate, UnauthenticatedTemplate,
 * MsalAuthenticationTemplate, withMsal.
 */
import * as React from "react";
import {
    PublicClientApplication,
    AccountInfo,
    AuthenticationResult,
    TokenRequest,
    EventType,
    InteractionType,
    InteractionKind,
} from "@mini-msal/browser";

export interface IMsalContext {
    instance: PublicClientApplication;
    accounts: AccountInfo[];
    inProgress: "startup" | "handleRedirect" | "none";
    /** internal: lets the provider re-render on active-account switches */
    activeId?: string;
}

const MsalContext = React.createContext<IMsalContext | null>(null);

export function MsalProvider(props: {
    instance: PublicClientApplication;
    children?: React.ReactNode;
}) {
    const { instance } = props;
    const [state, setState] = React.useState<IMsalContext>(() => ({
        instance,
        accounts: instance.getAllAccounts(),
        inProgress: "startup",
        activeId: instance.getActiveAccount()?.homeAccountId,
    }));

    React.useEffect(() => {
        // like real msal-react: only publish a new context value when the
        // accounts actually changed (stable references prevent effect loops
        // in consumers that depend on an account object)
        const update = (inProgress: IMsalContext["inProgress"]) =>
            setState((prev) => {
                const accounts = instance.getAllAccounts();
                const activeId = instance.getActiveAccount()?.homeAccountId;
                const sameAccounts =
                    JSON.stringify(prev.accounts) === JSON.stringify(accounts);
                if (
                    sameAccounts &&
                    prev.inProgress === inProgress &&
                    prev.instance === instance &&
                    prev.activeId === activeId
                ) {
                    return prev;
                }
                return {
                    instance,
                    accounts: sameAccounts ? prev.accounts : accounts,
                    inProgress,
                    activeId,
                };
            });
        const id = instance.addEventCallback(({ eventType }) => {
            if (
                eventType === EventType.LOGIN_SUCCESS ||
                eventType === EventType.LOGOUT_SUCCESS ||
                eventType === EventType.ACQUIRE_TOKEN_SUCCESS ||
                eventType === EventType.ACCOUNT_ADDED ||
                eventType === EventType.ACCOUNT_REMOVED ||
                eventType === EventType.ACTIVE_ACCOUNT_CHANGED
            ) {
                update("none");
            }
        });
        update("handleRedirect");
        instance
            .handleRedirectPromise()
            .catch(() => {})
            .then(() => update("none"));
        return () => {
            if (id) instance.removeEventCallback(id);
        };
    }, [instance]);

    return (
        <MsalContext.Provider value={state}>
            {props.children}
        </MsalContext.Provider>
    );
}

export function useMsal(): IMsalContext {
    const ctx = React.useContext(MsalContext);
    if (!ctx) throw new Error("useMsal must be used inside MsalProvider");
    return ctx;
}

export function useIsAuthenticated(): boolean {
    return useMsal().accounts.length > 0;
}

export function useAccount(filter?: {
    homeAccountId?: string;
    localAccountId?: string;
    username?: string;
}): AccountInfo | null {
    const { accounts, activeId } = useMsal();
    if (!filter) {
        // resolve from the context's stable accounts array so consumers get
        // a stable reference AND re-render when the active account switches
        return accounts.find((a) => a.homeAccountId === activeId) ?? null;
    }
    return (
        accounts.find(
            (a) =>
                (!filter.homeAccountId ||
                    a.homeAccountId === filter.homeAccountId) &&
                (!filter.localAccountId ||
                    a.localAccountId === filter.localAccountId) &&
                (!filter.username ||
                    a.username.toLowerCase() ===
                        filter.username.toLowerCase())
        ) ?? null
    );
}

export interface MsalAuthenticationResult {
    login: (
        interactionType?: InteractionKind,
        request?: TokenRequest
    ) => Promise<AuthenticationResult | null>;
    result: AuthenticationResult | null;
    error: unknown;
}

export function useMsalAuthentication(
    interactionType: InteractionKind,
    request: TokenRequest
): MsalAuthenticationResult {
    const { instance, inProgress } = useMsal();
    const isAuthenticated = useIsAuthenticated();
    const [result, setResult] =
        React.useState<AuthenticationResult | null>(null);
    const [error, setError] = React.useState<unknown>(null);
    const attempted = React.useRef(false);

    React.useEffect(() => {
        // a failed redirect roundtrip (e.g. user cancelled at the IdP) is
        // this hook's error — captured via the synchronous LOGIN_FAILURE
        // event (like real msal-react) so it is guaranteed to land before
        // inProgress flips to "none" and the auto-login effect runs
        const id = instance.addEventCallback((m) => {
            if (m.eventType === EventType.LOGIN_FAILURE) {
                setError(m.error);
            } else if (m.eventType === EventType.LOGIN_SUCCESS && m.payload) {
                setResult(m.payload as AuthenticationResult);
            }
        });
        return () => {
            if (id) instance.removeEventCallback(id);
        };
    }, [instance]);

    const login = React.useCallback(
        async (type: InteractionKind = interactionType, req = request) => {
            try {
                let res: AuthenticationResult | null = null;
                if (type === InteractionType.Popup) {
                    res = await instance.loginPopup(req);
                } else if (type === InteractionType.Silent) {
                    res = await instance
                        .acquireTokenSilent(req)
                        .catch(() => instance.ssoSilent(req));
                } else {
                    await instance.loginRedirect(req);
                }
                setResult(res);
                setError(null);
                return res;
            } catch (e) {
                setError(e);
                throw e;
            }
        },
        [instance, interactionType, request]
    );

    React.useEffect(() => {
        if (
            isAuthenticated ||
            error ||
            inProgress !== "none" ||
            attempted.current
        ) {
            return;
        }
        attempted.current = true;
        login().catch(() => {});
    }, [isAuthenticated, error, inProgress, login]);

    return { login, result, error };
}

export function AuthenticatedTemplate(props: { children?: React.ReactNode }) {
    const { inProgress } = useMsal();
    return useIsAuthenticated() && inProgress !== "startup" ? (
        <>{props.children}</>
    ) : null;
}

export function UnauthenticatedTemplate(props: {
    children?: React.ReactNode;
}) {
    const { inProgress } = useMsal();
    return !useIsAuthenticated() && inProgress === "none" ? (
        <>{props.children}</>
    ) : null;
}

export function MsalAuthenticationTemplate(props: {
    interactionType: InteractionKind;
    authenticationRequest: TokenRequest;
    errorComponent?: React.ComponentType<{ error: unknown }>;
    loadingComponent?: React.ComponentType;
    children?: React.ReactNode;
}) {
    const isAuthenticated = useIsAuthenticated();
    const { error } = useMsalAuthentication(
        props.interactionType,
        props.authenticationRequest
    );
    if (isAuthenticated) {
        return <>{props.children}</>;
    }
    if (error && props.errorComponent) {
        return <props.errorComponent error={error} />;
    }
    return props.loadingComponent ? <props.loadingComponent /> : null;
}

export function withMsal<P extends { msalContext: IMsalContext }>(
    Component: React.ComponentType<P>
) {
    return (props: Omit<P, "msalContext">) => {
        const ctx = useMsal();
        return <Component {...(props as P)} msalContext={ctx} />;
    };
}
