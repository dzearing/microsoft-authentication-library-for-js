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
import { localStorageCache } from "@mini-msal/browser/local-storage";
import {
    telemetry,
    type TelemetryClient,
} from "@mini-msal/browser/telemetry";

export { BrowserPerformanceClient } from "@mini-msal/browser/telemetry";

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
    version,
    type AccountInfo,
    type AccountFilter,
    type AuthenticationResult,
    type Config,
    type EventMessage,
    type InteractionKind,
    type TokenRequest,
} from "@mini-msal/browser";

// ---- real msal-browser's exported enums/constants (surface parity) ----

export const BrowserCacheLocation = {
    LocalStorage: "localStorage",
    SessionStorage: "sessionStorage",
    MemoryStorage: "memoryStorage",
} as const;

export const ProtocolMode = {
    AAD: "AAD",
    OIDC: "OIDC",
    EAR: "EAR",
} as const;

export const PromptValue = {
    LOGIN: "login",
    SELECT_ACCOUNT: "select_account",
    CONSENT: "consent",
    CREATE: "create",
    NONE: "none",
    NO_SESSION: "no_session",
} as const;

export const OIDC_DEFAULT_SCOPES = ["openid", "profile", "offline_access"];

/** real's 52 BrowserAuthErrorCodes, packed (camelCase key <- snake code) */
export const BrowserAuthErrorCodes: Record<string, string> =
    Object.fromEntries(
        (
            "pkce_not_created ear_jwk_empty ear_jwe_empty crypto_nonexistent " +
            "empty_navigate_uri hash_empty_error no_state_in_hash " +
            "hash_does_not_contain_known_properties unable_to_parse_state " +
            "state_interaction_type_mismatch interaction_in_progress " +
            "interaction_in_progress_cancelled popup_window_error " +
            "empty_window_error user_cancelled redirect_bridge_empty_response " +
            "redirect_in_iframe block_iframe_reload block_nested_popups " +
            "iframe_closed_prematurely silent_logout_unsupported " +
            "no_account_error silent_prompt_value_error " +
            "no_token_request_cache_error unable_to_parse_token_request_cache_error " +
            "auth_request_not_set_error invalid_cache_type non_browser_environment " +
            "database_not_open no_network_connectivity post_request_failed " +
            "get_request_failed failed_to_parse_response unable_to_load_token " +
            "crypto_key_not_found auth_code_required " +
            "auth_code_or_nativeAccountId_required spa_code_and_nativeAccountId_present " +
            "database_unavailable unable_to_acquire_token_from_native_platform " +
            "native_handshake_timeout native_extension_not_installed " +
            "native_connection_not_established uninitialized_public_client_application " +
            "native_prompt_not_supported invalid_base64_string " +
            "invalid_pop_token_request failed_to_build_headers " +
            "failed_to_parse_headers failed_to_decrypt_ear_response " +
            "timed_out empty_response"
        )
            .split(" ")
            .map((c) => [c.replace(/_(.)/g, (_, ch) => ch.toUpperCase()), c])
    );

export interface PublicClientApplication
    extends AuthClient,
        PopupClient,
        TelemetryClient {}
export class PublicClientApplication {
    constructor(config: Config) {
        // the composed closure-based client IS the instance (constructor
        // return override) — compat adds no wrapper layer; telemetry last
        // so it can wrap the methods other features attach
        return createClient(config, [
            localStorageCache,
            popup,
            telemetry,
        ]) as PublicClientApplication;
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

/** real's factory surface: construct + initialize */
export const createStandardPublicClientApplication = createAuth;
/** NAA factory — no bridge support yet (C9); like real with no bridge
 * present, falls back to a standard PCA */
export const createNestablePublicClientApplication = createAuth;
/** platform-broker probe — mini has no broker support yet (C7) */
export const isPlatformBrokerAvailable = async (): Promise<boolean> => false;
