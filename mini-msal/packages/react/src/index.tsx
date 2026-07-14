/**
 * mini-msal React bindings — a byte-sized port of @azure/msal-react 5.5.1:
 * MsalProvider (initialize + the full InteractionStatus state machine),
 * useMsal, useAccount, useIsAuthenticated, useMsalAuthentication,
 * AuthenticatedTemplate, UnauthenticatedTemplate, MsalAuthenticationTemplate,
 * withMsal.
 */
import * as React from "react";
import {
    AccountInfo,
    AuthenticationResult,
    TokenRequest,
    AuthError,
    EventType,
    InteractionType,
    InteractionStatus,
    InteractionRequiredAuthError,
    WrapperSKU,
    Logger,
    type EventMessage,
    type InteractionKind,
    type AuthClient,
} from "@mini-msal/browser";
import type { PopupClient } from "@mini-msal/browser/popup";

/**
 * The provider works against any composed client instance that has the core +
 * popup surface — e.g. @mini-msal/compat's PublicClientApplication or
 * createClient(config, [popup]).
 */
export type IPublicClientApplication = AuthClient & PopupClient;

/** real @azure/msal-react's identity (impersonated like the core WIRE_ID) */
const NAME = "@azure/msal-react";
export const version = "5.5.1";
const OIDC_DEFAULT_SCOPES = ["openid", "profile", "offline_access"];

type Status = (typeof InteractionStatus)[keyof typeof InteractionStatus];

export interface IMsalContext {
    instance: IPublicClientApplication;
    accounts: AccountInfo[];
    inProgress: Status;
    logger: Logger;
}

export interface AccountIdentifiers {
    homeAccountId?: string;
    localAccountId?: string;
    username?: string;
}

export class ReactAuthError extends AuthError {
    name = "ReactAuthError";
}

// ---- account helpers (real msal-react's utils/utilities.ts) ----------------

const sameIds = (a: AccountInfo, b: AccountInfo) =>
    a.homeAccountId === b.homeAccountId &&
    a.localAccountId === b.localAccountId &&
    a.username === b.username;

const accountArraysAreEqual = (a: AccountInfo[], b: AccountInfo[]) =>
    a.length === b.length && a.every((x, i) => x && b[i] && sameIds(x, b[i]));

const hasIds = (ids?: AccountIdentifiers) =>
    !!ids && !!(ids.homeAccountId || ids.localAccountId || ids.username);

function getAccountByIdentifiers(
    all: AccountInfo[],
    ids: AccountIdentifiers
): AccountInfo | null {
    if (!all.length || !hasIds(ids)) return null;
    const eq = (want: string | undefined, have: string) =>
        !want || want.toLowerCase() === have.toLowerCase();
    return (
        all.find(
            (a) =>
                eq(ids.username, a.username) &&
                eq(ids.homeAccountId, a.homeAccountId) &&
                eq(ids.localAccountId, a.localAccountId)
        ) ?? null
    );
}

function getChildrenOrFunction<T>(
    children: React.ReactNode | ((args: T) => React.ReactNode),
    args: T
): React.ReactNode {
    return typeof children === "function" ? children(args) : children;
}

// ---- provider ---------------------------------------------------------------

export const MsalContext = React.createContext<IMsalContext | null>(null);
export const MsalConsumer = MsalContext.Consumer;

/** real EventMessageUtils.getInteractionStatusFromEvent (incl. clear-guards) */
function statusFromEvent(m: EventMessage, current: Status): Status | null {
    const interactive =
        m.interactionType === InteractionType.Redirect ||
        m.interactionType === InteractionType.Popup;
    switch (m.eventType) {
        case EventType.ACQUIRE_TOKEN_START:
            return interactive ? InteractionStatus.AcquireToken : null;
        case EventType.HANDLE_REDIRECT_START:
            return InteractionStatus.HandleRedirect;
        case EventType.LOGOUT_START:
            return InteractionStatus.Logout;
        case EventType.LOGOUT_END:
            return current === InteractionStatus.Logout
                ? InteractionStatus.None
                : null;
        case EventType.HANDLE_REDIRECT_END:
            return current === InteractionStatus.HandleRedirect
                ? InteractionStatus.None
                : null;
        case EventType.ACQUIRE_TOKEN_SUCCESS:
        case EventType.ACQUIRE_TOKEN_FAILURE:
        case EventType.RESTORE_FROM_BFCACHE:
            return interactive && current === InteractionStatus.AcquireToken
                ? InteractionStatus.None
                : null;
    }
    return null;
}

export function MsalProvider(props: {
    instance: IPublicClientApplication;
    children?: React.ReactNode;
}) {
    const { instance } = props;
    React.useEffect(() => {
        instance.initializeWrapperLibrary(WrapperSKU.React, version);
    }, [instance]);
    const logger = React.useMemo(
        () => instance.getLogger().clone(NAME, version),
        [instance]
    );
    // real MsalProvider's reducer: every event may move inProgress via the
    // status mapping; accounts only start updating once startup completes.
    // A null action is the UNBLOCK_INPROGRESS fallback after
    // handleRedirectPromise settles (cached promises fire no events).
    const [state, dispatch] = React.useReducer(
        (
            prev: { inProgress: Status; accounts: AccountInfo[] },
            message: EventMessage | null
        ) => {
            let inProgress = prev.inProgress;
            if (message) {
                inProgress =
                    statusFromEvent(message, prev.inProgress) ?? inProgress;
            } else if (prev.inProgress === InteractionStatus.Startup) {
                inProgress = InteractionStatus.None;
            }
            if (inProgress === InteractionStatus.Startup) return prev;
            const accounts = instance.getAllAccounts();
            const same = accountArraysAreEqual(accounts, prev.accounts);
            return inProgress === prev.inProgress && same
                ? prev
                : { inProgress, accounts: same ? prev.accounts : accounts };
        },
        { inProgress: InteractionStatus.Startup, accounts: [] }
    );
    React.useEffect(() => {
        const id = instance.addEventCallback(dispatch);
        instance
            .initialize()
            .then(() =>
                instance
                    .handleRedirectPromise()
                    .catch(() => {})
                    .finally(() => dispatch(null))
            )
            .catch(() => {});
        return () => {
            if (id) instance.removeEventCallback(id);
        };
    }, [instance]);
    const value = React.useMemo(
        () => ({ instance, logger, ...state }),
        [instance, logger, state]
    );
    return (
        <MsalContext.Provider value={value}>
            {props.children}
        </MsalContext.Provider>
    );
}

// ---- hooks -------------------------------------------------------------------

export function useMsal(): IMsalContext {
    const ctx = React.useContext(MsalContext);
    if (!ctx) throw new Error("useMsal must be used inside MsalProvider");
    return ctx;
}

export function useIsAuthenticated(
    matchAccount?: AccountIdentifiers
): boolean {
    const { accounts, inProgress } = useMsal();
    if (inProgress === InteractionStatus.Startup) return false;
    return hasIds(matchAccount)
        ? !!getAccountByIdentifiers(accounts, matchAccount!)
        : accounts.length > 0;
}

function getAccount(
    instance: IPublicClientApplication,
    ids?: AccountIdentifiers
): AccountInfo | null {
    // no identifiers at all -> the active account, like real
    if (!hasIds(ids)) return instance.getActiveAccount();
    return getAccountByIdentifiers(instance.getAllAccounts(), ids!);
}

function accountInfoIsEqual(
    a: AccountInfo | null,
    b: AccountInfo | null
): boolean {
    if (!a || !b) return false;
    const ca = (a.idTokenClaims ?? {}) as Record<string, unknown>;
    const cb = (b.idTokenClaims ?? {}) as Record<string, unknown>;
    return (
        sameIds(a, b) &&
        a.tenantId === b.tenantId &&
        a.environment === b.environment &&
        a.nativeAccountId === b.nativeAccountId &&
        // iat/nonce change with every new id token — consumers re-render
        ca.iat === cb.iat &&
        ca.nonce === cb.nonce
    );
}

export function useAccount(
    accountIdentifiers?: AccountIdentifiers
): AccountInfo | null {
    const { instance, inProgress } = useMsal();
    const [account, setAccount] = React.useState<AccountInfo | null>(() =>
        inProgress === InteractionStatus.Startup
            ? null
            : getAccount(instance, accountIdentifiers)
    );
    const updateAccount = React.useCallback(() => {
        const next = getAccount(instance, accountIdentifiers);
        setAccount((cur) => (accountInfoIsEqual(cur, next) ? cur : next));
    }, [instance, accountIdentifiers]);
    React.useEffect(() => {
        if (inProgress !== InteractionStatus.Startup) updateAccount();
        const id = instance.addEventCallback((m: EventMessage) => {
            if (
                m.eventType === EventType.ACTIVE_ACCOUNT_CHANGED ||
                (m.eventType === EventType.LOGIN_SUCCESS &&
                    m.interactionType === InteractionType.Silent)
            ) {
                updateAccount();
            }
        });
        return () => {
            if (id) instance.removeEventCallback(id);
        };
    }, [updateAccount, inProgress, instance]);
    return account;
}

export interface MsalAuthenticationResult {
    login: (
        interactionType?: InteractionKind,
        request?: TokenRequest
    ) => Promise<AuthenticationResult | null>;
    acquireToken: (
        interactionType?: InteractionKind,
        request?: TokenRequest
    ) => Promise<AuthenticationResult | null>;
    result: AuthenticationResult | null;
    error: unknown;
}

export function useMsalAuthentication(
    interactionType: InteractionKind,
    authenticationRequest?: TokenRequest,
    accountIdentifiers?: AccountIdentifiers
): MsalAuthenticationResult {
    const { instance, inProgress } = useMsal();
    const isAuthenticated = useIsAuthenticated(accountIdentifiers);
    const account = useAccount(accountIdentifiers);
    const [[result, error], setResponse] = React.useState<
        [AuthenticationResult | null, unknown]
    >([null, null]);
    // prevent state updates after unmount, like real
    const mounted = React.useRef(true);
    React.useEffect(() => {
        mounted.current = true;
        return () => {
            mounted.current = false;
        };
    }, []);
    const interactionInProgress = React.useRef(
        inProgress !== InteractionStatus.None
    );
    React.useEffect(() => {
        interactionInProgress.current = inProgress !== InteractionStatus.None;
    }, [inProgress]);
    // controls the one automatic login/acquireToken this hook performs
    const shouldAcquireToken = React.useRef(true);
    React.useEffect(() => {
        if (error || result) shouldAcquireToken.current = false;
    }, [error, result]);

    const login = React.useCallback(
        async (
            type: InteractionKind = interactionType,
            req = authenticationRequest
        ) => {
            const getToken = async () => {
                if (type === InteractionType.Popup) {
                    return instance.loginPopup(req);
                }
                if (type === InteractionType.Redirect) {
                    // this promise is not expected to resolve (full-frame nav)
                    await instance.handleRedirectPromise();
                    return instance.loginRedirect(req).then(() => null);
                }
                if (type === InteractionType.Silent) {
                    return instance.ssoSilent((req ?? {}) as TokenRequest);
                }
                const err = new ReactAuthError(
                    "invalid_interaction_type",
                    "The provided interaction type is invalid."
                );
                if (mounted.current) setResponse([null, err]);
                throw err;
            };
            return getToken().then(
                (res) => {
                    if (res && mounted.current) setResponse([res, null]);
                    return res;
                },
                (e) => {
                    if (mounted.current) setResponse([null, e]);
                    throw e;
                }
            );
        },
        [instance, interactionType, authenticationRequest]
    );

    const acquireToken = React.useCallback(
        async (
            type: InteractionKind = interactionType,
            req?: TokenRequest
        ) => {
            const tokenRequest: TokenRequest = req
                ? { ...req }
                : {
                      ...authenticationRequest,
                      scopes:
                          authenticationRequest?.scopes ??
                          OIDC_DEFAULT_SCOPES,
                  };
            tokenRequest.correlationId ??= crypto.randomUUID();
            if (!tokenRequest.account && account) {
                tokenRequest.account = account;
            }
            return instance.acquireTokenSilent(tokenRequest).then(
                (res) => {
                    if (mounted.current) setResponse([res, null]);
                    return res;
                },
                (e: unknown) => {
                    if (e instanceof InteractionRequiredAuthError) {
                        if (!interactionInProgress.current) {
                            return login(type, tokenRequest);
                        }
                        e = new ReactAuthError(
                            "unable_to_fallback_to_interaction",
                            "Interaction is required but another interaction is already in progress. Please try again when the current interaction is complete."
                        );
                    }
                    if (mounted.current) setResponse([null, e]);
                    throw e;
                }
            );
        },
        [instance, interactionType, authenticationRequest, account, login]
    );

    React.useEffect(() => {
        if (
            shouldAcquireToken.current &&
            inProgress === InteractionStatus.None
        ) {
            if (!isAuthenticated) {
                shouldAcquireToken.current = false;
                login().catch(() => {});
            } else if (account) {
                shouldAcquireToken.current = false;
                acquireToken().catch(() => {});
            }
            // if logged out, clear the result so the component doesn't show
            // authenticated content for a user that is gone
        } else if (!account && result) {
            setResponse([null, null]);
        }
    }, [isAuthenticated, account, inProgress, login, acquireToken]);

    return { login, acquireToken, result, error };
}

// ---- components ---------------------------------------------------------------

export interface TemplateProps extends AccountIdentifiers {
    children?: React.ReactNode | ((ctx: IMsalContext) => React.ReactNode);
}

export function AuthenticatedTemplate({
    username,
    homeAccountId,
    localAccountId,
    children,
}: TemplateProps) {
    const context = useMsal();
    const isAuthenticated = useIsAuthenticated({
        username,
        homeAccountId,
        localAccountId,
    });
    return isAuthenticated &&
        context.inProgress !== InteractionStatus.Startup ? (
        <>{getChildrenOrFunction(children, context)}</>
    ) : null;
}

export function UnauthenticatedTemplate({
    username,
    homeAccountId,
    localAccountId,
    children,
}: TemplateProps) {
    const context = useMsal();
    const isAuthenticated = useIsAuthenticated({
        username,
        homeAccountId,
        localAccountId,
    });
    return !isAuthenticated &&
        context.inProgress !== InteractionStatus.Startup &&
        context.inProgress !== InteractionStatus.HandleRedirect ? (
        <>{getChildrenOrFunction(children, context)}</>
    ) : null;
}

export function MsalAuthenticationTemplate(props: {
    interactionType: InteractionKind;
    authenticationRequest?: TokenRequest;
    username?: string;
    homeAccountId?: string;
    localAccountId?: string;
    errorComponent?: React.ElementType;
    loadingComponent?: React.ElementType;
    children?:
        | React.ReactNode
        | ((result: MsalAuthenticationResult) => React.ReactNode);
}) {
    const { username, homeAccountId, localAccountId } = props;
    const ids = React.useMemo(
        () => ({ username, homeAccountId, localAccountId }),
        [username, homeAccountId, localAccountId]
    );
    const context = useMsal();
    const msalAuthResult = useMsalAuthentication(
        props.interactionType,
        props.authenticationRequest,
        ids
    );
    const isAuthenticated = useIsAuthenticated(ids);
    const ErrorComp = props.errorComponent;
    const LoadingComp = props.loadingComponent;
    if (
        msalAuthResult.error &&
        context.inProgress === InteractionStatus.None
    ) {
        if (ErrorComp) return <ErrorComp {...msalAuthResult} />;
        // no error component: throw so an app error boundary catches it
        throw msalAuthResult.error;
    }
    if (isAuthenticated) {
        return <>{getChildrenOrFunction(props.children, msalAuthResult)}</>;
    }
    if (LoadingComp && context.inProgress !== InteractionStatus.None) {
        return <LoadingComp {...context} />;
    }
    return null;
}

export function withMsal<P extends { msalContext: IMsalContext }>(
    Component: React.ComponentType<P>
) {
    return (props: Omit<P, "msalContext">) => {
        const ctx = useMsal();
        return <Component {...(props as P)} msalContext={ctx} />;
    };
}
