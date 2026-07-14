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
    loadExternalTokens as loadExternalTokensCore,
    type AuthClient,
    type Config,
} from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";
import { pop } from "@mini-msal/browser/pop";
import { broker, type BrokerClient } from "@mini-msal/browser/broker";
import { localStorageCache } from "@mini-msal/browser/local-storage";
import { cacheMigration } from "@mini-msal/browser/cache-migration";
import { createNestableClient } from "@mini-msal/browser/naa";
import {
    telemetry,
    type TelemetryClient,
} from "@mini-msal/browser/telemetry";

export { BrowserPerformanceClient } from "@mini-msal/browser/telemetry";
export { NativeAuthError } from "@mini-msal/browser/broker";

export {
    AuthError,
    InteractionRequiredAuthError,
    Logger,
    LogLevel,
    WrapperSKU,
    BrowserAuthError,
    ClientAuthError,
    ClientConfigurationError,
    ServerError,
    NestedAppAuthError,
    EventType,
    InteractionType,
    InteractionStatus,
    CacheLookupPolicy,
    NavigationClient,
    type NavigationOptions,
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

export const AuthenticationScheme = {
    BEARER: "Bearer",
    POP: "pop",
    SSH: "ssh-cert",
} as const;

/** unpack "snake_code" (key camelized from code) / "irregularKey=snake_code"
 * entries into real's *ErrorCodes namespace shape */
const pack = (s: string): Record<string, string> =>
    Object.fromEntries(
        s.split(" ").map((e) => {
            const [k, v] = e.split("=");
            return v
                ? [k, v]
                : [k.replace(/_(.)/g, (_, ch) => ch.toUpperCase()), k];
        })
    );

/** real's 52 BrowserAuthErrorCodes, packed (camelCase key <- snake code) */
export const BrowserAuthErrorCodes: Record<string, string> = pack(
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
);

/** real's AuthErrorCodes (msal-common auth-layer) */
export const AuthErrorCodes = pack("post_request_failed unexpected_error");

/** real's 37 ClientAuthErrorCodes */
export const ClientAuthErrorCodes = pack(
    "authorization_code_missing_from_server_response binding_key_not_removed " +
        "cannotAppendScopeSet=cannot_append_scopeset cannot_remove_empty_scope " +
        "client_info_decoding_error client_info_empty_error " +
        "emptyInputScopeSet=empty_input_scopeset " +
        "end_session_endpoint_not_supported " +
        "endpointResolutionError=endpoints_resolution_error " +
        "hash_not_deserialized invalid_cache_environment invalid_cache_record " +
        "invalid_state key_id_missing method_not_implemented " +
        "misplacedResourceParam=misplaced_resource_parameter " +
        "multiple_matching_appMetadata multiple_matching_tokens " +
        "nested_app_auth_bridge_disabled network_error no_account_found " +
        "no_account_in_silent_request no_crypto_object no_network_connectivity " +
        "nonce_mismatch null_or_empty_token " +
        "openIdConfigError=openid_config_error platform_broker_error " +
        "request_cannot_be_made resource_parameter_required state_mismatch " +
        "state_not_found " +
        "tokenClaimsCnfRequiredForSignedJwt=token_claims_cnf_required_for_signedjwt " +
        "token_parsing_error token_refresh_required " +
        "unexpected_credential_type user_canceled"
);

/** real's 24 ClientConfigurationErrorCodes */
export const ClientConfigurationErrorCodes = pack(
    "authority_mismatch authority_uri_insecure cannot_allow_platform_broker " +
        "cannot_set_OIDCOptions claims_request_parsing_error " +
        "empty_input_scopes_error invalid_authentication_header " +
        "invalid_authority_metadata invalid_claims " +
        "invalid_cloud_discovery_metadata invalid_code_challenge_method " +
        "invalid_platform_broker_configuration invalid_request_method_for_EAR " +
        "issuer_validation_failed logout_request_empty " +
        "missing_nonce_authentication_header missing_ssh_jwk missing_ssh_kid " +
        "pkce_params_missing redirect_uri_empty token_request_empty " +
        "untrusted_authority urlEmptyError=empty_url_error url_parse_error"
);

/** real's 9 InteractionRequiredAuthErrorCodes */
export const InteractionRequiredAuthErrorCodes = pack(
    "bad_token consent_required interaction_required interrupted_user " +
        "login_required native_account_unavailable no_tokens_found " +
        "refresh_token_expired ui_not_allowed"
);

/** real's 3 BrowserConfigurationAuthErrorCodes */
export const BrowserConfigurationAuthErrorCodes = pack(
    "in_mem_redirect_unavailable storage_not_supported " +
        "stubbed_public_client_application_called"
);

// the long tail of real's export surface (D5): BrowserConfigurationAuthError,
// enums/constants, BrowserUtils, storage classes, EventHandler,
// EventMessageUtils, perf helpers, SignedHttpRequest,
// stubbedPublicClientApplication, enforceResourceParameter
export * from "./surface.js";

export interface PublicClientApplication
    extends AuthClient,
        PopupClient,
        BrokerClient,
        TelemetryClient {}
export class PublicClientApplication {
    constructor(config: Config) {
        // the composed closure-based client IS the instance (constructor
        // return override) — compat adds no wrapper layer; telemetry last
        // so it can wrap the methods other features attach
        return createClient(config, [
            localStorageCache,
            cacheMigration,
            popup,
            pop,
            broker,
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

/** real's top-level loadExternalTokens (ITokenCache successor), composed
 * with the same cache backends the compat PCA uses */
export const loadExternalTokens = (
    config: Config,
    request: Parameters<typeof loadExternalTokensCore>[1],
    response: Record<string, any>,
    options?: Parameters<typeof loadExternalTokensCore>[3]
) =>
    loadExternalTokensCore(config, request, response, options, [
        localStorageCache,
        cacheMigration,
    ]);

/** real's factory surface: construct + initialize */
export const createStandardPublicClientApplication = createAuth;
/** NAA factory: bridge-backed client when the host provides
 * window.nestedAppAuthBridge, else a standard PCA (like real) */
export const createNestablePublicClientApplication = (
    config: Config
): Promise<PublicClientApplication> =>
    createNestableClient(config, createAuth);
/** platform-broker probe — mini has no broker support yet (C7) */
export const isPlatformBrokerAvailable = async (): Promise<boolean> => false;
