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
