/**
 * Generates docs/GAP_REPORT.md from results/mini.json plus the
 * authored per-scenario classifications below.
 *
 * Classes:
 *   pass            — observably identical (after normalization)
 *   behavioral-diff — mini implements the capability but differs observably
 *   missing-feature — capability absent from mini by design
 *   bug             — mini misbehaves on something it claims to support
 */
import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const results = JSON.parse(
    readFileSync(path.join(__dirname, "results", "mini.json"), "utf8")
);

// ---- authored classifications ----------------------------------------------
// real/mini: concise observed behavior. cost: est. KB (minified) to close.
const C = {
    // ---- core ----
    "core.login-redirect-roundtrip": {
        class: "behavioral-diff",
        real: "Full event sequence (initializeStart/End, handleRedirectStart, acquireTokenSuccess, loginSuccess, handleRedirectEnd w/ interactionType); authorize carries nonce, client-request-id, client_info=1, claims(login_hint/signin_state), x-client-SKU/VER, X-AnchorMailbox; result has authority/correlationId/tokenType/state; scopes keep request casing.",
        mini: "Only accountAdded+loginSuccess+handleRedirectEnd events; authorize lacks nonce (no id_token nonce validation — security-relevant), telemetry and routing params; result lacks authority/correlationId/tokenType/state; scopes lowercased, different order.",
        cost: 1.0,
    },
    "core.handle-redirect-clean-load": {
        class: "behavioral-diff",
        real: "Clean load emits initializeStart/End only; handleRedirectPromise resolves null with NO handleRedirect events.",
        mini: "No initialize events; emits handleRedirectEnd even on a clean load.",
        cost: 0.1,
    },
    "core.login-popup-roundtrip": {
        class: "behavioral-diff",
        real: "acquireTokenStart/popupOpened/acquireTokenSuccess/loginSuccess (interactionType=popup); token POST carries x-client-SKU/VER, telemetry params, x-ms-lib-capability, claims, client-request-id query.",
        mini: "loginSuccess+accountAdded only; token POST carries only the OAuth basics; same PKCE correctness (S256 verified on both).",
        cost: 0.8,
    },
    "core.acquire-token-popup": {
        class: "behavioral-diff",
        real: "Popup acquisition for new scopes: full event stream, result metadata.",
        mini: "Works (token acquired); shape/event gaps as above.",
        cost: 0,
    },
    "core.sso-silent-cold": {
        class: "behavioral-diff",
        real: "InteractionRequiredAuthError named class, acquireTokenStart/Failure events.",
        mini: "Correct login_required code and IRAE classification, but error.name is 'Error' (classes never set .name) and no events on failure.",
        cost: 0.1,
    },
    "core.sso-silent-warm": {
        class: "behavioral-diff",
        real: "Succeeds via hidden iframe; emits acquireTokenSuccess AND loginSuccess for ssoSilent.",
        mini: "Succeeds; emits no events for ssoSilent (only cache-level accountAdded on first write).",
        cost: 0.2,
    },
    "core.acquire-token-redirect": {
        class: "behavioral-diff",
        real: "Derives login_hint + X-AnchorMailbox from the account on redirect acquisition (sticky account routing).",
        mini: "No account-derived hints on the authorize request — on a shared IdP session the wrong user could be silently picked.",
        cost: 0.2,
    },
    "core.redirect-state-tampered": {
        class: "behavioral-diff",
        real: "Forged state: handleRedirectPromise resolves NULL silently (response treated as not-ours; no failure event).",
        mini: "Throws state_mismatch + emits loginFailure. Stricter than real; apps double-handling errors would behave differently.",
        cost: 0.2,
    },
    "core.storage-shape-after-login": {
        class: "behavioral-diff",
        real: "Entities carry lastUpdatedAt + cachedByApiId metadata; writes msal.version; active account NOT auto-set.",
        mini: "Core msal.3 schema matches (interop verified) but lacks metadata fields; adds its own msal.meta.* discovery cache; AUTO-SETS active account on first login (real never does).",
        cost: 0.3,
    },
    "core.popup-open-timing": {
        class: "behavioral-diff",
        real: "Default navigatePopups=true: window.open('about:blank') fires SYNCHRONOUSLY in the caller's stack (popup blockers see the user gesture) with real's name (msal.{clientId}.{scopes}.{authority}.{cid}) and sized/centered features string, then the popup is navigated to the authorize URL.",
        mini: "Pre-C14: opened the final authorize URL only AFTER async PKCE work (blocker-visible gap), fixed 'width=483,height=600,popup=yes' features, random msal.{uuid} name; navigatePopups ignored.",
        cost: 0.5,
    },
    "core.popup-open-timing-async": {
        class: "behavioral-diff",
        real: "system.navigatePopups=false: the open is deferred and fires directly at the authorize URL (async, real's non-default mode) with the same name/features format.",
        mini: "Pre-C14: identical deferred timing by accident, but name/features format differed.",
        cost: 0,
    },
    "core.popup-window-attributes": {
        class: "missing-feature",
        real: "request.popupWindowAttributes popupSize/popupPosition drive the features string (clamped to the parent window; centered defaults).",
        mini: "Pre-C14: attributes ignored — fixed features string.",
        cost: 0.2,
    },

    // ---- silent ----
    "silent.cache-hit-fresh": {
        class: "behavioral-diff",
        real: "Cache hit; still fetches authority metadata (1 discovery request) per fresh client.",
        mini: "Cache hit with zero network (own discovery cache). Result-shape gaps only.",
        cost: 0,
    },
    "silent.expiry-window-refresh": {
        class: "behavioral-diff",
        real: "AT inside 300s renewal window → refresh_token grant.",
        mini: "Same refresh behavior (300s window matches); shape gaps only.",
        cost: 0,
    },
    "silent.refresh-token-grant": {
        class: "behavioral-diff",
        real: "RT grant carries telemetry/claims/anchor params; redirect_uri = request redirectUri.",
        mini: "RT grant works; sends config redirectUri instead of request's; missing protocol extras.",
        cost: 0.1,
    },
    "silent.iframe-fallback-no-rt": {
        class: "behavioral-diff",
        real: "prompt=none iframe fallback with account-derived hints.",
        mini: "Same fallback chain (authorize prompt=none → code exchange); param gaps only.",
        cost: 0,
    },
    "silent.policy-access-token-valid": {
        class: "behavioral-diff",
        real: "CacheLookupPolicy.AccessToken + valid AT → cache hit.",
        mini: "Same outcome; shape gaps only.",
        cost: 0,
    },
    "silent.policy-access-token-expired": {
        class: "bug",
        real: "CacheLookupPolicy.AccessToken + expired AT → throws ClientAuthError token_refresh_required (no fallback, as documented).",
        mini: "Ignores the policy: silently falls through to the RT grant and RETURNS A TOKEN where real throws.",
        cost: 0.3,
    },
    "silent.policy-at-and-rt": {
        class: "behavioral-diff",
        real: "Expired AT → RT grant (no iframe).",
        mini: "Policy enforced (A5); result-shape gaps only (B1).",
        cost: 0,
    },
    "silent.policy-refresh-token": {
        class: "behavioral-diff",
        real: "CacheLookupPolicy.RefreshToken IGNORES the valid AT and uses the RT grant (fromCache=false).",
        mini: "Policy enforced (A5): same RT grant, fromCache=false; result-shape gaps only (B1).",
        cost: 0,
    },
    "silent.policy-rt-and-network": {
        class: "behavioral-diff",
        real: "RT grant fails (invalid_grant) → falls back to prompt=none iframe → new code exchange.",
        mini: "Policy enforced (A5): same RT→iframe network sequence; result-shape gaps only (B1).",
        cost: 0,
    },
    "silent.policy-skip": {
        class: "behavioral-diff",
        real: "Skip → straight to prompt=none iframe (skips AT cache AND RT).",
        mini: "Policy enforced (A5): straight to the iframe like real; result-shape gaps only (B1).",
        cost: 0,
    },
    "silent.concurrent-dedupe": {
        class: "bug",
        real: "Two parallel identical acquireTokenSilent calls → ONE token-endpoint request (in-flight dedupe), both callers get the same token.",
        mini: "Two parallel calls → TWO refresh requests (no dedupe). Same tokens returned, but doubled IdP load and RT-rotation hazard against real AAD.",
        cost: 0.3,
    },
    "silent.force-refresh": {
        class: "behavioral-diff",
        real: "forceRefresh:true bypasses a valid cached AT and refreshes over the network; emits acquireTokenFromNetworkStart.",
        mini: "forceRefresh honored (A5): same RT-grant network refresh; remaining diffs are event-stream (B2) + result shape (B1).",
        cost: 0,
    },

    // ---- accounts ----
    "accounts.multi-account": {
        class: "behavioral-diff",
        real: "Two accounts listed with environment field; active account remains NULL until setActiveAccount is called.",
        mini: "Accounts match (no environment field); auto-sets the FIRST logged-in account as active.",
        cost: 0.2,
    },
    "accounts.get-account-filters": {
        class: "bug",
        real: "getAccount({}) → null; getAllAccounts(filter) filters (e.g. by username → 1 account). Case-insensitive username match.",
        mini: "getAccount({}) returns the first account; getAllAccounts(filter) ignores the filter and returns everything.",
        cost: 0.2,
    },
    "accounts.active-account-persistence": {
        class: "behavioral-diff",
        real: "Active account persists across reload + fresh client (identical in both); real also emits initialize events after reload.",
        mini: "Persistence identical; only the initialize events are missing.",
        cost: 0,
    },
    "accounts.logout-popup-per-account": {
        class: "behavioral-diff",
        real: "logoutStart/Success/End events (interactionType=popup); end_session gets client-request-id + state; per-request postLogoutRedirectUri honored.",
        mini: "Only that account's entities removed (matches); events differ (accountRemoved/logoutSuccess, no start/end); per-request postLogoutRedirectUri IGNORED (uses config); no state on end_session.",
        cost: 0.3,
    },
    "accounts.logout-popup-main-window-redirect": {
        class: "missing-feature",
        real: "logoutPopup({mainWindowRedirectUri}) navigates the MAIN window (NavigationClient, ApiId 962) after the popup roundtrip; popupWindowAttributes honored on the logout popup; name msal.{clientId}.{homeAccountId}.{cid}; interaction lock survives into the next page.",
        mini: "Pre-C14: mainWindowRedirectUri ignored — main window stayed on the authenticated page over a cleared cache; attributes/name diverged.",
        cost: 0.4,
    },
    "accounts.logout-redirect": {
        class: "behavioral-diff",
        real: "Sets msal.interaction.status (signout) temp cache; end_session carries state + client-request-id; storage fully cleared after roundtrip.",
        mini: "Storage cleared (matches); no temp interaction state, no state param on end_session.",
        cost: 0.2,
    },
    "accounts.local-storage": {
        class: "missing-feature",
        real: "cacheLocation:'localStorage' honored — and v5 ENCRYPTS localStorage entities ({id, nonce, data:<AES ciphertext>}, per-session key in cookie), only key indexes stay plaintext.",
        mini: "cacheLocation ignored; everything stays in sessionStorage. NOTE: real's encrypted localStorage means cross-stack cache interop is fundamentally impossible in localStorage mode.",
        cost: 2.5,
    },
    "accounts.cross-tab-events": {
        class: "missing-feature",
        real: "Second tab sees the account via localStorage; logout in tab 1 raises events in tab 2 (storage listener).",
        mini: "No localStorage support → second tab sees nothing.",
        cost: 0.5,
    },

    // ---- errors ----
    "errors.cancelled-login-redirect": {
        class: "behavioral-diff",
        real: "access_denied → ServerError (named class), loginFailure + handleRedirectStart/End events.",
        mini: "Same access_denied code; class name 'Error', fewer events.",
        cost: 0.1,
    },
    "errors.popup-blocked": {
        class: "behavioral-diff",
        real: "BrowserAuthError popup_window_error + acquireTokenStart/Failure events.",
        mini: "Same popup_window_error code; name/message/events differ.",
        cost: 0,
    },
    "errors.popup-closed-by-user": {
        class: "behavioral-diff",
        real: "v5.16 has NO popup-close detection: closing the popup leaves the promise pending until popupBridgeTimeout, then BrowserAuthError timed_out (subError redirect_bridge_timeout).",
        mini: "Detects the closed window within 50ms → user_cancelled immediately. Mini is faster/friendlier but observably different (different code, no 60s hang).",
        cost: 0,
    },
    "errors.iframe-timeout": {
        class: "behavioral-diff",
        real: "iframeBridgeTimeout configurable (set to 1.5s) → timed_out/redirect_bridge_timeout.",
        mini: "Timeout hardcoded at 10s (config ignored) — with an 8s-delayed IdP mini got the late login_required response instead of timing out.",
        cost: 0.1,
    },
    "errors.interaction-required-variants": {
        class: "behavioral-diff",
        real: "interaction_required/consent_required/login_required → InteractionRequiredAuthError (named).",
        mini: "Identical codes and IRAE classification; only .name missing ('Error').",
        cost: 0.05,
    },
    "errors.uninitialized-client": {
        class: "bug",
        real: "Uniform BrowserAuthError uninitialized_public_client_application from loginPopup/acquireTokenSilent/handleRedirectPromise/logoutRedirect before initialize().",
        mini: "No guard: loginPopup/logoutRedirect CRASH with TypeError (metadata undefined); handleRedirectPromise resolves; silent throws no_account_error.",
        cost: 0.15,
    },
    "errors.redirect-in-iframe": {
        class: "behavioral-diff",
        real: "BrowserAuthError redirect_in_iframe.",
        mini: "Same code; name/message differ.",
        cost: 0,
    },
    "errors.nested-popup-guard": {
        class: "behavioral-diff",
        real: "SURPRISE: real v5.16 ALLOWS loginPopup from an msal.*-named window (completes fine); silent fails only with no_account_error. The window-name block documented for v3/v4 is gone in the bridge era.",
        mini: "Guard removed (matches real): popup completes but result shape differs (B1) and silent then SUCCEEDS because mini auto-set the active account (A8).",
        cost: 0.1,
    },
    "errors.interaction-in-progress": {
        class: "bug",
        real: "Second interactive call while a popup is pending → BrowserAuthError interaction_in_progress; the FIRST call completes normally.",
        mini: "No interaction lock + fixed 'msal.popup' window name: the second popup NAVIGATES the first one — first call fails state_mismatch, SECOND call wins the token. Callers race.",
        cost: 0.3,
    },
    "errors.silent-unknown-account": {
        class: "behavioral-diff",
        real: "Unknown account object → ClientConfigurationError authority_mismatch; no account at all → BrowserAuthError no_account_error.",
        mini: "Both cases → no_account (IRAE-classified). Coarser taxonomy.",
        cost: 0.1,
    },

    // ---- params ----
    "params.prompt-passthrough": {
        class: "pass",
        real: "prompt login/consent/select_account forwarded verbatim.",
        mini: "Identical.",
        cost: 0,
    },
    "params.login-hint-domain-hint": {
        class: "missing-feature",
        real: "login_hint + domain_hint + derived X-AnchorMailbox (UPN routing) on authorize.",
        mini: "login_hint works; domainHint and X-AnchorMailbox unsupported.",
        cost: 0.1,
    },
    "params.sid-passthrough": {
        class: "missing-feature",
        real: "sid forwarded on prompt=none requests.",
        mini: "sid not supported.",
        cost: 0.05,
    },
    "params.extra-query-parameters": {
        class: "missing-feature",
        real: "extraQueryParameters land on authorize; tokenQueryParameters land on the token endpoint QUERY string (plus client-request-id).",
        mini: "Neither supported.",
        cost: 0.15,
    },
    "params.claims-and-cae": {
        class: "missing-feature",
        real: "claims + clientCapabilities:['cp1'] merged into one claims JSON (xms_cc added under access_token) on BOTH authorize and token requests.",
        mini: "claims/CAE unsupported.",
        cost: 0.3,
    },
    "params.custom-state": {
        class: "missing-feature",
        real: "Custom state encoded inside the wire state (libState|userState), returned verbatim on result.state.",
        mini: "Custom state ignored; result.state absent.",
        cost: 0.2,
    },
    "params.per-request-redirect-uri": {
        class: "pass",
        real: "request.redirectUri overrides config for authorize + token redemption.",
        mini: "Identical.",
        cost: 0,
    },
    "params.authority-override-per-request": {
        class: "missing-feature",
        real: "Per-request authority triggers discovery for that authority; result.authority reflects it; token refreshed against it.",
        mini: "authority + forceRefresh ignored → returned the cached token from the default authority.",
        cost: 0.3,
    },
    "params.scopes-normalization": {
        class: "behavioral-diff",
        real: "Dedupes exact-duplicate scopes (keeps distinct casings), appends OIDC defaults once; empty scopes → 'openid profile offline_access'.",
        mini: "Scope wire format matches (A4); remaining diffs are B1 result-shape fields (authority, correlationId, tokenType, state, fromPlatformBroker).",
        cost: 0.1,
    },

    // ---- resilience ----
    "resilience.throttle-429-retry-after": {
        class: "missing-feature",
        real: "429+Retry-After → ServerError; writes a throttling cache entry; an immediate retry is served the SAME error from the throttle cache with NO network request.",
        mini: "Error surfaced (generic), no throttle cache — immediate retry hits the IdP again (and succeeded once the injection expired).",
        cost: 0.8,
    },
    "resilience.5xx-server-error": {
        class: "behavioral-diff",
        real: "503 → ServerError with telemetry-formatted message; no automatic retry (1 hit).",
        mini: "Same single-attempt behavior; error is a bare 'Error' with the raw description.",
        cost: 0.05,
    },
    "resilience.network-drop": {
        class: "behavioral-diff",
        real: "Dropped socket transparently retried by the browser fetch stack — request succeeded on both stacks; only result-shape diffs.",
        mini: "Same recovery; shape gaps only.",
        cost: 0,
    },
    "resilience.proactive-refresh": {
        class: "missing-feature",
        real: "refresh_in from the token response is persisted as refreshOn on the AT entity (basis for proactive refresh).",
        mini: "refresh_in ignored; no refreshOn.",
        cost: 0.2,
    },

    // ---- init ----
    "init.double-initialize": {
        class: "behavioral-diff",
        real: "Second initialize() is a no-op; initializeStart/End emitted once.",
        mini: "Also a no-op; no initialize events exist.",
        cost: 0,
    },
    "init.get-configuration": {
        class: "missing-feature",
        real: "getConfiguration() returns fully-resolved defaults (iframeHashTimeout 10000/6000, windowHashTimeout 60000, tokenRenewalOffsetSeconds 300, allowPlatformBroker false, redirectNavigationTimeout 30000 …).",
        mini: "No getConfiguration.",
        cost: 0.2,
    },
    "init.logger-callback": {
        class: "missing-feature",
        real: "loggerCallback receives verbose pipeline logs (levels Info+Verbose seen, >10 entries for one popup login).",
        mini: "No logger at all.",
        cost: 0.4,
    },
    "init.exported-surface": {
        class: "missing-feature",
        real: "Exports version, createStandard/NestablePublicClientApplication, isPlatformBrokerAvailable, ClientAuthError/ServerError classes, 52 BrowserAuthErrorCodes, full EventType (incl. broker events), PromptValue/ProtocolMode/BrowserCacheLocation enums, OIDC_DEFAULT_SCOPES.",
        mini: "Small subset: PublicClientApplication, 8 EventTypes, 3 error classes, CacheLookupPolicy, InteractionType.",
        cost: 1.0,
    },
    "init.exported-surface-2": {
        class: "missing-feature",
        real: "Exports Logger (clone/level gate/format), LogLevel, WrapperSKU, AuthErrorCodes/ClientAuthErrorCodes/ClientConfigurationErrorCodes/InteractionRequiredAuthErrorCodes/BrowserConfigurationAuthErrorCodes namespaces + BrowserConfigurationAuthError class; PCA exposes getLogger/setLogger/initializeWrapperLibrary.",
        mini: "None of these existed pre-C12 — import of LogLevel or any *ErrorCodes namespace failed at module load; getLogger() was undefined.",
        cost: 0.5,
    },
    "init.exported-surface-full": {
        class: "missing-feature",
        real: "Exports the FULL 51-key module surface: ApiId, AzureCloudInstance, JsonWebTokenTypes, ResponseMode, DEFAULT_IFRAME_TIMEOUT_MS, BrowserRootPerformanceEvents, BrowserUtils namespace (22 fns), LocalStorage/SessionStorage/MemoryStorage, EventHandler, EventMessageUtils, AuthenticationHeaderParser, SignedHttpRequest, BrowserPerformanceMeasurement, StubPerformanceClient, enforceResourceParameter, stubbedPublicClientApplication.",
        mini: "Pre-D5: 18 of these exports missing entirely (import failed at module load).",
        cost: 0.5,
    },
    "init.storage-before-login": {
        class: "behavioral-diff",
        real: "Bare initialize() writes only msal.version.",
        mini: "Writes its msal.meta.<authority> discovery cache instead (this is also how mini avoids re-fetching discovery).",
        cost: 0,
    },
    "init.popup-without-bridge": {
        class: "behavioral-diff",
        real: "HARD dependency on the redirect-bridge page: a popup landing on a blank (bridge-less) page never resolves → timed_out/redirect_bridge_timeout. Apps MUST serve the bridge at every popup redirectUri.",
        mini: "URL-polls the popup — works fine with a blank page (that is mini's normal mode).",
        cost: 0,
    },

    // ---- telemetry ----
    "telemetry.perf-events-popup-login": {
        class: "missing-feature",
        real: "With telemetry.client=BrowserPerformanceClient: initializeClientApplication + acquireTokenPopup events (success, durations, correlationId). NOTE: default client is a stub — apps must opt in.",
        mini: "No performance client / addPerformanceCallback.",
        cost: 2.0,
    },
    "telemetry.perf-events-silent": {
        class: "missing-feature",
        real: "acquireTokenSilent perf event for cache hit AND network refresh (cache hit event has success+duration).",
        mini: "None.",
        cost: 0,
    },
    "telemetry.perf-event-shape": {
        class: "missing-feature",
        real: "Full emitted perf-event shape (C11): init + popup events carry real's complete field set (sizes, cache/migration counters, instance counts, redirect-bridge config) plus the ext sub-measurement aggregate keyed by real's internal operation names.",
        mini: "Events carried only {name, correlationId, durationMs, success, errorCode}.",
        cost: 1.0,
    },
    "telemetry.perf-event-shape-silent": {
        class: "missing-feature",
        real: "Full perf-event shape for acquireTokenSilent cache hit / network refresh / ssoSilent / RT+iframe failure: per-flow field sets incl. cacheOutcome, accountCachedBy, accessTokensRemoved, silentRefreshReason and flow-specific ext key sets.",
        mini: "Same 5-field minimal events for every silent flow.",
        cost: 1.0,
    },
    "telemetry.token-request-headers": {
        class: "missing-feature",
        real: "Server telemetry travels as token BODY params (not HTTP headers): x-client-SKU=msal.js.browser, x-client-VER, x-client-current-telemetry, x-client-last-telemetry, x-ms-lib-capability='retry-after, h429'.",
        mini: "None of the telemetry params.",
        cost: 0.5,
    },
    "telemetry.correlation-id-propagation": {
        class: "missing-feature",
        real: "request.correlationId → authorize client-request-id param, result.correlationId, perf events.",
        mini: "correlationId not supported anywhere.",
        cost: 0.2,
    },
    "telemetry.last-telemetry-after-failure": {
        class: "missing-feature",
        real: "After a failed RT grant real retried via iframe fallback (2 token calls) and the next request carries x-client-last-telemetry (empty summary here; failure cache exists).",
        mini: "Failure surfaced directly; no retry chain, no last-telemetry.",
        cost: 0,
    },

    // ---- broker ----
    "broker.dom-probe-and-interactive": {
        class: "missing-feature",
        real: "DOM broker probed via navigator.platformAuthentication.getSupportedContracts('MicrosoftEntra'); acquireTokenByCode({nativeAccountId}) routes to executeGetToken (request: accountId/clientId/scope/redirectUri/correlationId/extraParameters incl. telemetry:MATS); result has fromPlatformBroker=true; account caches nativeAccountId; subsequent SILENT calls also broker-route.",
        mini: "No acquireTokenByCode, no broker (TypeError).",
        cost: 10.0,
    },
    "broker.dom-first-login": {
        class: "missing-feature",
        real: "KEY CONTRACT: even with a live broker, a first-time interactive login does NOT use it (no cached nativeAccountId) — web flow + spa_accountid/hybrid seeding is required first.",
        mini: "Same observable outcome (web flow) — but only because mini has no broker at all.",
        cost: 0,
    },
    "broker.dom-error-mapping": {
        class: "missing-feature",
        real: "Broker statuses map exactly: USER_CANCEL→BrowserAuthError user_cancelled; ACCOUNT_UNAVAILABLE→IRAE native_account_unavailable; NO_NETWORK→BrowserAuthError no_network_connectivity.",
        mini: "n/a (no broker).",
        cost: 0,
    },
    "broker.dom-disabled-fallback": {
        class: "missing-feature",
        real: "DISABLED is fatal: NativeAuthError surfaced, broker dropped — the second call doesn't touch the broker and throws unable_to_acquire_token_from_native_platform.",
        mini: "n/a.",
        cost: 0,
    },
    "broker.extension-handshake-capture": {
        class: "missing-feature",
        real: "With allowPlatformBroker and no extension: TWO Handshake postMessages (channel 53ee284d-…, preferred extensionId ppnbnpeolgkicgegkbkbjmhlideopiji then undefined, MessagePort transferred); failure swallowed; init + web flows unaffected.",
        mini: "Config ignored; no handshake attempted (also no failure).",
        cost: 0,
    },
    "broker.extension-fake-e2e": {
        class: "missing-feature",
        real: "Full extension protocol works against a fake: HandshakeResponse over the transferred port, GetToken request {accountId, scope, tokenType, windowTitleSubstring, extraParameters:{telemetry:MATS}}, Response/Success/result → AuthenticationResult with fromPlatformBroker=true.",
        mini: "n/a.",
        cost: 0,
    },
    "broker.is-platform-broker-available": {
        class: "missing-feature",
        real: "isPlatformBrokerAvailable() exported; returns false in this environment either way (needs full extension handshake, not just the DOM object).",
        mini: "Not exported.",
        cost: 0,
    },

    // ---- naa ----
    "naa.init-handshake": {
        class: "missing-feature",
        real: "createNestablePublicClientApplication + window.nestedAppAuthBridge: GetInitContext envelope (messageType NestedAppAuthRequest, clientLibrary msal.js.browser 5.16.0, requestId/sendTime); NAA controller active (acquireTokenRedirect → NestedAppAuthError unsupported_method).",
        mini: "createNestablePublicClientApplication not exported.",
        cost: 5.0,
    },
    "naa.get-token-popup": {
        class: "missing-feature",
        real: "acquireTokenPopup → bridge GetTokenPopup (tokenParams: clientId, scope 'openid profile offline_access User.Read', authenticationScheme Bearer, correlationId); token+account response mapped into AuthenticationResult.",
        mini: "n/a.",
        cost: 0,
    },
    "naa.silent-cache-then-bridge": {
        class: "missing-feature",
        real: "First silent → bridge GetToken; SECOND silent served from the browser cache (bridge call count stays 1).",
        mini: "n/a.",
        cost: 0,
    },
    "naa.error-mapping": {
        class: "missing-feature",
        real: "USER_INTERACTION_REQUIRED→IRAE(code); USER_CANCEL→ClientAuthError user_canceled; PERSISTENT_ERROR→ServerError(code); NO_NETWORK→ClientAuthError no_network_connectivity.",
        mini: "n/a.",
        cost: 0,
    },
    "naa.unsupported-apis": {
        class: "missing-feature",
        real: "loginRedirect/logoutRedirect/logoutPopup/acquireTokenByCode/addPerformanceCallback → NestedAppAuthError unsupported_method; handleRedirectPromise → null.",
        mini: "n/a.",
        cost: 0,
    },
    "naa.no-bridge-fallback": {
        class: "missing-feature",
        real: "No bridge present → silently falls back to a standard PCA (perf API works, web popup login succeeds).",
        mini: "n/a.",
        cost: 0,
    },
    "navigation.deep-link-replay": {
        class: "missing-feature",
        real: "navigateToLoginRequestUrl (default true): start page cached under msal.{cid}.request.origin; on return at a different URL the response hash is cached under msal.{cid}.urlHash and the window navigates back to the deep link (lock held mid-replay); handleRedirectPromise delivers the result there on the next load.",
        mini: "Pre-C13: always processed the hash in place on the redirectUri page — deep-link users landed on the wrong page; no request.origin/urlHash keys.",
        cost: 1.0,
    },
    "navigation.on-redirect-navigate-cancel-login": {
        class: "missing-feature",
        real: "auth.onRedirectNavigate(authorizeUrl) fires before navigating; returning false cancels navigation, loginRedirect resolves, interaction lock stays held, request.origin stays cached.",
        mini: "Pre-C13: hook never read; page always hard-navigated via location.assign.",
        cost: 0.3,
    },
    "navigation.on-redirect-navigate-cancel-logout": {
        class: "missing-feature",
        real: "onRedirectNavigate(endSessionUrl) on logoutRedirect: false cancels navigation, releases the interaction lock, logoutStart (null payload) + logoutEnd fire.",
        mini: "Pre-C13: hook never read; also emitted logoutStart with a synthesized payload where real passes the raw (undefined) request.",
        cost: 0.2,
    },
    "navigation.navigation-client-config": {
        class: "missing-feature",
        real: "system.navigationClient routes redirect navigation through navigateExternal(url, {apiId: 861, timeout: 30000, noHistory: false}) instead of navigating.",
        mini: "Pre-C13: no navigationClient seam — location.assign directly.",
        cost: 0.3,
    },
    "navigation.navigation-client-setter": {
        class: "missing-feature",
        real: "NavigationClient exported (subclassable); setNavigationClient() swaps the client consulted by redirect flows at runtime.",
        mini: "Pre-C13: no NavigationClient export; setNavigationClient threw TypeError.",
        cost: 0.3,
    },
    "react.provider-initializes-instance": {
        class: "missing-feature",
        real: "MsalProvider initialize()s an un-initialized instance, then handleRedirectPromise processes the response hash (accounts populate, code redeemed); initializeWrapperLibrary on mount.",
        mini: "Pre-C15: provider never called initialize(); handleRedirectPromise rejected uninitialized_public_client_application (swallowed) — signed-out UI, redirect response lost.",
        cost: 0.5,
    },
    "react.inprogress-status-sequence": {
        class: "missing-feature",
        real: "inProgress walks the full InteractionStatus machine per event: startup → none, acquireToken during popup interaction, logout during logoutPopup, with real's clear-guards.",
        mini: "Pre-C15: 3-value union stuck at 'none' during popup interaction and logout — `disabled={inProgress !== 'none'}` gating never engaged.",
        cost: 0.5,
    },
    "react.use-msal-authentication": {
        class: "missing-feature",
        real: "useMsalAuthentication returns {login, acquireToken, result, error}; auto-acquires silently for a signed-in user (active account attached) and resets result when the account disappears after logout.",
        mini: "Pre-C15: no acquireToken callback (undefined), no auto-acquire for signed-in users (result stayed null), stale result kept after logout.",
        cost: 0.5,
    },
    "react.template-render-props-identifiers": {
        class: "missing-feature",
        real: "Templates invoke function-as-children with the msal context and scope on username/homeAccountId/localAccountId props (case-insensitive).",
        mini: "Pre-C15: function children rendered as nothing (React warning); identifier props silently ignored — any signed-in account matched.",
        cost: 0.3,
    },
    "react.auth-template-error-contract": {
        class: "missing-feature",
        real: "MsalAuthenticationTemplate renders ErrorComponent with the ENTIRE auth result spread ({login, acquireToken, result, error}); without one the error is THROWN so an app error boundary catches it; LoadingComponent gets the msal context as props.",
        mini: "Pre-C15: ErrorComponent got only {error}; errors without an ErrorComponent were swallowed (loading/null rendered); LoadingComponent got no props.",
        cost: 0.3,
    },
    "react.account-identifier-matching": {
        class: "missing-feature",
        real: "useIsAuthenticated(identifiers) matches a specific user (case-insensitive on all three ids, false during startup); useAccount({}) falls back to the active account.",
        mini: "Pre-C15: useIsAuthenticated took no argument (any account = authenticated, true during startup); homeAccountId/localAccountId compared case-sensitively; empty filter returned accounts[0] instead of the active account.",
        cost: 0.4,
    },
    "cache.at-scope-dedupe": {
        class: "behavior-diff",
        real: "saveAccessToken removes every cached AT for the same account/realm/type whose (non-OIDC) scope set intersects the new token's — at most one AT per scope family; a narrower silent request then serves the NEW token.",
        mini: "Pre-C16: ATs accumulated per scope-set key and the silent lookup returned the FIRST (older, narrower) match — a stale token the server had superseded.",
        cost: 0.5,
    },
    "cache.at-multi-match-clear": {
        class: "behavior-diff",
        real: "getAccessToken with >1 matching cached ATs removes them ALL and refreshes over the network.",
        mini: "Pre-C16: served the first matching AT from cache, leaving the ambiguous duplicates in place.",
        cost: 0.2,
    },
    "cache.tenant-profile-merge": {
        class: "behavior-diff",
        real: "Guest-tenant tokens merge into ONE base account entity (homeAccountId+environment) with an appended tenantProfile (isHomeTenant computed); getAllAccounts expands profiles into per-tenant AccountInfo objects.",
        mini: "Pre-C16: wrote a second account entity per realm, each falsely claiming isHomeTenant:true, and never expanded profiles.",
        cost: 0.6,
    },
    "cache.schema-migration": {
        class: "missing-feature",
        real: "initialize migrates msal.0/1/2-schema accounts+tokens into msal.3 (stamping lastUpdatedAt, pruning entries older than cacheRetentionDays=5d, expired or invalid) so pre-v5 users keep silent SSO.",
        mini: "Pre-C16: old-schema entries were ignored forever — upgraded users were forced to re-authenticate. Now a compat-composed /cache-migration feature.",
        cost: 1.0,
    },
    "cache.kmsi-plaintext-localstorage": {
        class: "behavior-diff",
        real: "KMSI entities (signin_state kmsi/dvc_dmjd) are persisted PLAINTEXT in localStorage mode so sign-in survives losing the per-session encryption cookie; AccountInfo.kmsi reflects the claim.",
        mini: "Pre-C16: everything was encrypted under the session cookie key (KMSI users lost sign-in on browser restart) and account.kmsi was hardcoded undefined.",
        cost: 0.4,
    },
    "token-apis.clear-cache": {
        class: "missing-feature",
        real: "clearCache(logoutRequest) is the documented local sign-out: purges every msal/client-id cache key (accounts, tokens, msal.version), emits activeAccountChanged when the active account goes, no navigation and no end_session request.",
        mini: "Pre-C17: no clearCache on the client — drop-in consumers crashed with TypeError and had no way to purge the cache without navigating.",
        cost: 0.5,
    },
    "token-apis.hydrate-cache": {
        class: "missing-feature",
        real: "hydrateCache(result, request) seeds browser storage from an externally-acquired AuthenticationResult (hybrid SSR apps): account entity (ApiId 963, cloudGraphHostName/msGraphHost stamped) + id/access token entities, never a refresh token; acquireTokenSilent then serves fromCache with zero network.",
        mini: "Pre-C17: method missing (TypeError); the only hydration logic was internal to the NAA feature.",
        cost: 0.4,
    },
    "token-apis.load-external-tokens": {
        class: "missing-feature",
        real: "Top-level loadExternalTokens(config, request, response, options) export (the getTokenCache/ITokenCache successor): builds its own cache, writes account (ApiId 964) + id/access/refresh token entities from a raw token response, returns a fromCache AuthenticationResult.",
        mini: "Pre-C17: export did not exist (undefined import) — no supported path for e2e harnesses or hybrid apps to inject externally-acquired tokens.",
        cost: 0.4,
    },
    "token-apis.acquire-token-by-code": {
        class: "missing-feature",
        real: "acquireTokenByCode({code}) redeems a confidential-client-acquired spa code at the token endpoint (no PKCE code_verifier, no redirect_uri), dedupes concurrent same-code calls onto one POST/promise, and emits acquireTokenStart per call + one acquireTokenSuccess.",
        mini: "Pre-C17: threw auth_code_or_nativeAccountId_required whenever only code was supplied — the documented hybrid-SPA flow failed client-side with zero IdP traffic.",
        cost: 0.5,
    },
    "token-apis.acquire-token-by-code-errors": {
        class: "behavior-diff",
        real: "acquireTokenByCode with BOTH code and nativeAccountId throws spa_code_and_nativeAccountId_present; with neither it throws auth_code_or_nativeAccountId_required — each with start+failure events and no wire traffic.",
        mini: "Pre-C17: the both-present case silently took the platform-broker path on nativeAccountId instead of throwing the dedicated error.",
        cost: 0.2,
    },
    "authority.lazy-discovery-default-path": {
        class: "behavior-diff",
        real: "initialize() issues zero network requests; the openid-configuration fetch happens lazily at the first token flow (AAD-mode path includes /v2.0/).",
        mini: "Pre-C18: initialize() eagerly awaited the discovery fetch — an extra startup request, and initialize() rejected outright when the IdP was unreachable (real defers failures to per-request errors).",
        cost: 0.4,
    },
    "authority.oidc-discovery-endpoint-path": {
        class: "behavior-diff",
        real: "system.protocolMode 'OIDC' + non-Microsoft host: the discovery URL omits /v2.0/ (<authority>/.well-known/openid-configuration) — the documented mode for Auth0/Okta/IdentityServer authorities. NOTE protocolMode lives under config.system; auth.protocolMode is ignored.",
        mini: "Pre-C18: always inserted /v2.0/ regardless of protocolMode — generic OIDC IdPs 404 that path, so initialize() failed where real works.",
        cost: 0.5,
    },
    "authority.authority-metadata-config": {
        class: "missing-feature",
        real: "auth.authorityMetadata inline endpoint JSON resolves endpoints from config — the openid-configuration request is never sent (documented startup-latency optimization).",
        mini: "Pre-C18: the option was never read; the discovery GET fired on every instance regardless.",
        cost: 0.4,
    },
    "authority.known-authorities-validation": {
        class: "missing-feature",
        real: "Default (AAD) mode with an unknown host: AAD instance-discovery GET to login.microsoftonline.com must vouch for the authority; a network error / invalid_instance rejects the flow with ClientAuthError endpoints_resolution_error before any IdP contact.",
        mini: "Pre-C18: no trust validation at all — mini silently completed flows against any authority host, and never sent the instance-discovery probe.",
        cost: 0.6,
    },
    "authority.hardcoded-cloud-metadata": {
        class: "behavior-diff",
        real: "Known Microsoft cloud hosts (login.microsoftonline.com et al.) resolve endpoints from hardcoded metadata — a default-authority app never sends a well-known request in its lifetime; logout URLs build offline.",
        mini: "Pre-C18: fetched openid-configuration from login.microsoftonline.com at initialize — extra request, and bootstrap broke offline where real works.",
        cost: 0.4,
    },
    "authority.instance-aware-cloud-instance": {
        class: "missing-feature",
        real: "instance_aware flows: cloud_instance_host_name in the authorize response switches the token-redemption host; cloud_graph_host_name/msgraph_host are cached on the account entity and surfaced as result.cloudGraphHostName/msGraphHost on interactive AND cache-hit results.",
        mini: "Pre-C18: the fragment fields were dropped (hash parse read only code/state/error) and both result fields were hardcoded '' — multi-cloud guest apps would call the wrong Graph cloud.",
        cost: 0.5,
    },
    "config.allow-redirect-in-iframe": {
        class: "missing-feature",
        real: "system.allowRedirectInIframe=true lets an embedded app (Teams tab, portal iframe) run full redirect flows: loginRedirect navigates the iframe to the IdP and handleRedirectPromise processes the returned code in place.",
        mini: "Pre-C19: the flag was never read — acquireTokenRedirect unconditionally threw redirect_in_iframe inside any iframe, and processRedirect bailed out.",
        cost: 0.5,
    },
    "config.token-renewal-offset": {
        class: "behavior-diff",
        real: "system.tokenRenewalOffsetSeconds tunes the cached-AT expiry buffer: a large offset forces proactive network refreshes, 0 serves near-expiry tokens from cache.",
        mini: "Pre-C19: the buffer was hardcoded to 300s and getConfiguration() still reported the user's value — apps tuning proactive refresh saw the opposite cache/network behavior.",
        cost: 0.4,
    },
    "config.logout-hint-params": {
        class: "missing-feature",
        real: "The end_session URL carries logout_hint (explicit request.logoutHint, else derived from the account's login_hint claim), id_token_hint, and extraQueryParameters — the IdP skips its account picker on logout.",
        mini: "Pre-C19: logout URLs carried only post_logout_redirect_uri/client-request-id/state; the hint fields weren't even accepted on the request, so real's silent picker-free logout became interactive.",
        cost: 0.4,
    },
    "config.server-telemetry-enabled": {
        class: "missing-feature",
        real: "system.serverTelemetryEnabled=true populates x-client-current/last-telemetry ('5|apiId,…' schema) on token POSTs and persists failures to a server-telemetry-<clientId> cache entry, flushed on the next request and cleared on success.",
        mini: "Pre-C19: the params were hardcoded empty and no entry was ever written — ESTS-side diagnostics lost the failure history when apps opted in.",
        cost: 0.2,
    },
    "telemetry.perf-redirect-event": {
        class: "missing-feature",
        real: "handleRedirectPromise emits ONE root acquireTokenRedirect perf event per processed redirect response (cached-request cid, redemption-half ext, previousLibraryVersion); clean loads and repeat (memoized) calls emit nothing.",
        mini: "Pre-C20: the redirect flow was never instrumented — dashboards keyed on name==='acquireTokenRedirect' went dark.",
        cost: 0.4,
    },
    "telemetry.perf-failure-correlation": {
        class: "behavior-diff",
        real: "A failed acquireTokenSilent without an app correlationId yields event cid === error.correlationId === the token request's client-request-id QUERY param — client events join to caught errors and server logs.",
        mini: "Pre-C20: errors carried no correlationId at all and the failure event got a fresh random UUID — failure events were uncorrelatable.",
        cost: 0.3,
    },
    "telemetry.perf-callback-dedupe": {
        class: "behavior-diff",
        real: "addPerformanceCallback dedupes registrations by callback source text (returns the existing id, single delivery); without a perf client the returned id is ''.",
        mini: "Pre-C20: every registration got a fresh UUID — re-registered callbacks (React effects) double-counted every metric; the stub path returned a random UUID instead of ''.",
        cost: 0.2,
    },
    "telemetry.perf-preflight-failures": {
        class: "behavior-diff",
        real: "acquireTokenSilent with no account ABANDONS the started measurement (zero events reach callbacks); uninitialized preflight failures end the measurement and DO emit success:false events.",
        mini: "Pre-C20: the wrapper emitted a success:false event for no_account_error too — phantom failures real never reports.",
        cost: 0.2,
    },
    "telemetry.perf-init-once": {
        class: "behavior-diff",
        real: "initialize() is measured at most once — repeat calls return early before the measurement starts, so exactly one initializeClientApplication event.",
        mini: "Pre-C20: every outer initialize() call emitted an event even though the inner call no-ops — duplicate startup events inflating init counts.",
        cost: 0.1,
    },
    "telemetry.performance-marks": {
        class: "missing-feature",
        real: "sessionStorage msal.browser.performance.enabled='1' + an opt-in perf client writes msal.start/end/measure.<op>.<cid> performance-timeline entries for the root and every completed sub-measurement (DevTools-visible spans).",
        mini: "Pre-C20: the Performance API was never touched — the diagnostics flag produced nothing.",
        cost: 0.1,
    },
    "pop.silent-shr": {
        class: "missing-feature",
        real: "authenticationScheme 'pop': fresh RSA-2048 binding keypair per request, token_type=pop + req_cnf (b64url {kid, xms_ksl:'sw'}) on the token endpoint, AT cached as AccessToken_With_AuthScheme with keyId + a 'pop' cache-key suffix (bearer ATs never match a pop lookup), result.tokenType 'pop' with the AT wrapped as a SignedHttpRequest JWT — re-signed with a fresh nonce on every cache hit; request.popKid reuses a key and skips signing.",
        mini: "Pre-C21: the field was silently ignored — the IdP never saw req_cnf, the app got an unbound bearer token that PoP-protected resources (ARM/Graph SHR) reject.",
        cost: 0.8,
    },
    "pop.ssh-scheme-and-errors": {
        class: "missing-feature",
        real: "authenticationScheme 'ssh-cert': the request's sshJwk rides req_cnf verbatim with token_type=ssh-cert, the entity keyId comes from the response key_id, results carry the raw AT (no signing); missing sshJwk/sshKid throw missing_ssh_jwk/missing_ssh_kid ClientConfigurationErrors.",
        mini: "Pre-C21: no validation, no wire params, bearer-shaped cache/result.",
        cost: 0.2,
    },
};

// ---- generate ---------------------------------------------------------------
const areaOrder = [
    "core",
    "silent",
    "accounts",
    "errors",
    "broker",
    "naa",
    "telemetry",
    "params",
    "resilience",
    "init",
    "navigation",
    "react",
    "cache",
    "token-apis",
    "authority",
    "config",
    "pop",
];
const areaTitles = {
    core: "1. Core flows",
    silent: "2. Silent acquisition",
    accounts: "3. Accounts & cache",
    errors: "4. Errors & guards",
    broker: "5. Platform broker / WAM",
    naa: "6. Nested app auth (NAA)",
    telemetry: "7. Telemetry",
    params: "8. Request passthrough",
    resilience: "9. Resilience",
    init: "10. Init & misc",
    navigation: "11. Redirect navigation",
    react: "12. React bindings",
    cache: "13. Cache entity semantics",
    "token-apis": "14. Programmatic token APIs",
    authority: "15. Authority modes & discovery",
    config: "16. Config knobs & logout params",
    pop: "17. Proof-of-Possession schemes",
};

const rows = results.map((r) => {
    const c = C[r.id];
    if (!c) throw new Error(`unclassified scenario: ${r.id}`);
    // a runner "pass" always wins; otherwise the authored class stands
    const cls = r.status === "pass" ? "pass" : c.class;
    return { ...r, ...c, class: cls };
});

const byArea = new Map(areaOrder.map((a) => [a, []]));
for (const r of rows) byArea.get(r.area).push(r);

const count = (list, cls) => list.filter((r) => r.class === cls).length;

let md = `# mini-msal conformance gap report

Generated by \`npm run conformance\` — ${rows.length} black-box scenarios captured
against **@azure/msal-browser 5.16.0 + msal-react 5.5.1** (snapshots in
\`test/conformance/snapshots/real/\`, deterministic across two consecutive runs), then
replayed against **mini-msal**. Diff details: \`test/conformance/results/mini.json\`.

Classes: **pass** = observably identical · **behavioral-diff** = capability
exists but differs observably · **missing-feature** = absent by design ·
**bug** = mini misbehaves on something it claims to support.

## Summary

| Area | Scenarios | Pass | Behavioral diff | Missing feature | Bug |
|---|---|---|---|---|---|
`;
for (const a of areaOrder) {
    const list = byArea.get(a);
    md += `| ${areaTitles[a]} | ${list.length} | ${count(list, "pass")} | ${count(list, "behavioral-diff")} | ${count(list, "missing-feature")} | ${count(list, "bug")} |\n`;
}
md += `| **Total** | **${rows.length}** | **${count(rows, "pass")}** | **${count(rows, "behavioral-diff")}** | **${count(rows, "missing-feature")}** | **${count(rows, "bug")}** |\n`;

md += `
## Systemic gaps (appear across most scenarios; counted once)

These cross-cutting diffs are the bulk of the raw diff noise and are listed
here instead of being repeated per scenario:

1. **Result shape** — mini's AuthenticationResult lacks \`authority\`,
   \`correlationId\`, \`tokenType\`, \`state\`, \`fromPlatformBroker\`,
   \`account.environment\`; scopes are lowercased and ordered differently
   (real preserves request casing/order). *behavioral-diff, ~0.3 KB.*
2. **Error classes don't set \`.name\`** — every mini error reports
   \`name: "Error"\` (real: BrowserAuthError / InteractionRequiredAuthError /
   ServerError / ClientAuthError / ClientConfigurationError). Codes are right
   far more often than names. *behavioral-diff, ~0.05 KB (one line per class).*
3. **Event stream** — mini has no initializeStart/End, acquireTokenStart/
   Failure, popupOpened, handleRedirectStart, logoutStart/End; no
   \`interactionType\` on any event; payload shapes differ; ordering differs
   (mini: accountAdded→loginSuccess; real: acquireTokenSuccess→loginSuccess).
   *behavioral-diff, ~0.5 KB.*
4. **Protocol params** — mini never sends \`nonce\` (also never validates the
   id_token nonce — the one security-relevant gap found), \`client-request-id\`,
   \`client_info=1\`, \`claims\` (login_hint/signin_state), \`x-client-SKU/VER\`,
   \`X-AnchorMailbox\`, \`clidata\`; scope param ordering differs. *behavioral-diff
   (nonce: arguably bug), ~0.4 KB.*
5. **Storage metadata** — mini doesn't write \`msal.version\` or per-entity
   \`lastUpdatedAt\`/\`cachedByApiId\`; adds its own \`msal.meta.*\` key; auto-sets
   the active account on first login (real never does). Core \`msal.3\` entity
   schema itself matches — cross-stack cache interop still verified by e2e.
   *behavioral-diff, ~0.2 KB.*

## Per-scenario findings

`;

for (const a of areaOrder) {
    const list = byArea.get(a);
    md += `### ${areaTitles[a]}\n\n`;
    for (const r of list) {
        const badge =
            r.class === "pass"
                ? "✅ pass"
                : r.class === "bug"
                  ? "🐞 **bug**"
                  : r.class === "missing-feature"
                    ? "🚫 missing-feature"
                    : "↔️ behavioral-diff";
        md += `#### \`${r.id}\` — ${badge}`;
        if (r.cost) md += ` · est. ${r.cost} KB to close`;
        md += `\n\n- **real**: ${r.real}\n- **mini**: ${r.mini}\n`;
        if (r.status === "error") {
            md += `- ⚠️ scenario errored on mini: ${String(r.fatal).slice(0, 160)}\n`;
        }
        md += "\n";
    }
}

md += `## Cost-to-parity estimate (minified, rough)

| Feature bucket | Est. KB |
|---|---|
| Systemic polish (result shape, error names, events, nonce+protocol params, storage metadata) | ~1.5 |
| CacheLookupPolicy semantics + forceRefresh + silent dedupe + interaction lock + uninit guard + filters (all 🐞 bugs) | ~1.5 |
| Request passthrough (sid, domainHint, eQP, claims/CAE, custom state, authority override) | ~1.1 |
| localStorage + cross-tab events | ~3.0 |
| Throttling + refreshOn | ~1.0 |
| Telemetry (perf client + server telemetry params + correlationId) | ~2.7 |
| Logger + getConfiguration + export surface | ~1.6 |
| Platform broker (DOM + extension transports) | ~10 |
| Nested app auth | ~5 |
| **Full parity total (vs mini's current ~20 KB, real's 248 KB)** | **~27 KB → ~47 KB** |

The first three buckets (~4 KB) close every **bug** and most app-visible
behavioral diffs; broker + NAA account for over half the remaining distance
and only matter for WAM/Teams-host scenarios.
`;

const out = path.join(__dirname, "../../docs/GAP_REPORT.md");
writeFileSync(out, md);
console.log(
    `GAP_REPORT.md written: ${rows.length} scenarios — ` +
        ["pass", "behavioral-diff", "missing-feature", "bug"]
            .map((c) => `${c}: ${count(rows, c)}`)
            .join(", ")
);
