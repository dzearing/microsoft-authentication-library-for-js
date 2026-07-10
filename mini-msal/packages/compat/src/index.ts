/**
 * @mini-msal/compat: the classic drop-in surface. Exports a
 * PublicClientApplication (constructor + initialize()) that composes the
 * @mini-msal/browser core with ALL feature modules, and re-exports the error
 * classes / enums / types — so swapping "@azure/msal-browser" for
 * "@mini-msal/compat" is a one-import change. Apps that want to pay only for
 * what they use compose createClient + features directly instead.
 */
import {
    createClient,
    type AuthClient,
    type Config,
} from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";

export {
    AuthError,
    InteractionRequiredAuthError,
    BrowserAuthError,
    ClientAuthError,
    ClientConfigurationError,
    ServerError,
    NestedAppAuthError,
    EventType,
    InteractionType,
    CacheLookupPolicy,
    type AccountInfo,
    type AccountFilter,
    type AuthenticationResult,
    type Config,
    type EventMessage,
    type InteractionKind,
    type TokenRequest,
} from "@mini-msal/browser";

export interface PublicClientApplication extends AuthClient, PopupClient {}
export class PublicClientApplication {
    constructor(config: Config) {
        // the composed closure-based client IS the instance (constructor
        // return override) — compat adds no wrapper layer
        return createClient(config, [popup]) as PublicClientApplication;
    }
}

/** Shorthand factory: creates and initializes a client in one call. */
export async function createAuth(
    config: Config
): Promise<PublicClientApplication> {
    const auth = new PublicClientApplication(config);
    await auth.initialize();
    return auth;
}
