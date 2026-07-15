# Parity work — history & decision log (A0–C10)

Companion to [PARITY_STATE.md](./PARITY_STATE.md). This file is HISTORY:
per-task decision-log entries, the completed A0–C10 task descriptions, and
the progress table. Task sessions should NOT read this wholesale — grep for
a task id (e.g. `B3`, `C6`) when a pending task needs background on a past
judgment call.

## Decision Log — dated per-task entries

Entries (append as you go):
- **A1 (2026-07-10)**: error taxonomy now mirrors real: `AuthError` base w/
  `.name`, aka.ms default message (`${code}: ${message}` on Error.message),
  `subError`; new ServerError/ClientAuthError/ClientConfigurationError/
  NestedAppAuthError; BrowserAuthError always carries the aka.ms message.
  `classify()` = real's interaction-required detection (code/desc/suberror);
  token-endpoint non-IRAE errors get real's `Error(s): … - Timestamp: …`
  ServerError format. Match-real REGRESSIONS made here: (1) removed the
  msal.*-window-name nested-popup guard (gone in real 5.16); (2) removed
  popup-close detection — closed popup now just times out
  (`timed_out`/`redirect_bridge_timeout` BrowserAuthError, was
  `user_cancelled`); (3) poll timeouts now honor `system.popupBridgeTimeout`
  (60s default) / `system.iframeBridgeTimeout` (10s default). `no_account` →
  `no_account_error` BrowserAuthError; unknown account → `authority_mismatch`
  ClientConfigurationError; state_mismatch is a ClientAuthError. NOTE for the
  loop: TWO concurrent Claude sessions were found running this SOP
  simultaneously (their `lsof … kill` steps killed each other's servers,
  causing phantom ECONNREFUSED crashes and results/mini.json clobbering) —
  ensure only ONE loop session is alive before trusting run output.
- **A0 (2026-07-10)**: core is `createClient(config, features)` in
  `packages/browser/src/index.ts`; internal seam = `ClientContext`
  (emit/preflight/authorizeUrl/pollForCode/redeem/clearAccount/logoutUrl +
  `client` to attach methods to). First feature: `@mini-msal/browser/popup`.
  `@mini-msal/compat`'s `PublicClientApplication` uses the constructor-return
  trick (`return createClient(config, [popup])`) — instances are the composed
  closure object, so `instanceof PublicClientApplication` is false (nothing
  relied on it). React types against `AuthClient & PopupClient`
  (`IPublicClientApplication`). Closure refactor shrank the compat stack
  20.1 → 18.7 KB min for free; core-only floor is 10.9 KB min / 4.2 KB gz
  (`mini-core` variant in `npm run measure`).

- **A7 (2026-07-10)**: interaction lock lives in core (`ctx.lock(type)/unlock()`)
  writing real's exact entry (`msal.interaction.status` =
  `{"clientId":…,"type":"signin"|"signout"}` — B4's signout entry now exists).
  Judgment call: mini's processRedirect releases the lock UNCONDITIONALLY at
  the top (real only clears own-clientId / signout locks in
  handleRedirectPromise); observably identical for single-client apps and all
  scenarios — revisit only if a multi-client scenario ever lands. loginPopup
  still emits LOGIN_FAILURE on interaction_in_progress (real emits nothing
  there) — events are B2's job. — "pass count went UP" validation is
  waived when a task's scenario ALSO depends on a later task (here B1's
  result-shape fields); the bar is then "task's expected diffs eliminated +
  zero regressions". Discovered: the conformance harness SORTS result.scopes
  (lib.mjs) — snapshot scope order is alphabetical, not wire order; real
  result scopes = ScopeSet.fromString(AT target).asArray(), casing preserved.

- **B1 (2026-07-10)**: result.state semantics per snapshots (NOT the naive
  "always custom-or-empty"): interactive flows + ssoSilent return state ""
  (or custom, C1); acquireTokenSilent results (ALL rungs, incl. the hidden
  iframe) have state undefined — msal-common sets `state: ""` in
  generateAuthenticationResult, but real silent responses observably carry
  none, so mini strips it on the silentLadder iframe rung. correlationId for
  redirect flows is generated at acquireTokenRedirect time and persisted in
  the msal.request temp entry so the post-redirect result reuses it (real
  stores it in the temp request the same way).

- **A5 (2026-07-10)**: silent ladder now policy-gated: useAT (pol 0/1/2 and
  !forceRefresh) → useRT (pol ≠ 1/5) → useFrame (pol 0/4/5). AccessToken(1)
  miss → ClientAuthError `token_refresh_required`; RT missing with no iframe
  rung → IRAE `no_tokens_found` (real's RefreshTokenClient behavior; not
  scenario-covered but kept for parity). ALSO front-ran part of B4: dropped
  the `msal.meta.*` sessionStorage discovery cache — real re-fetches
  discovery per instance (every policy snapshot has `networkCalls[0] =
  "discovery"` on the second instance), so discovery is now per-instance
  memory only. This is what eliminated the networkCalls diffs across all
  silent.* scenarios.

- **B2 (2026-07-10)**: event streams mirror real 5.16's emit sites exactly.
  Key discoveries locked in: (1) event payloadKeys expose Object.keys of real's
  objects, so mini's AuthenticationResult now carries real's FULL 21-key shape
  (incl. undefined-valued cloudGraphHostName/code/refreshOn/state-on-silent,
  uniqueId/tenantId/idTokenClaims/extExpiresOn/familyId/requestId/msGraphHost)
  and AccountInfo the full 15-key shape (authorityType/dataBoundary/idToken/
  kmsi/loginHint/nativeAccountId/tenantProfiles/upn); silent results have
  `state: undefined` (key PRESENT), not a deleted key. (2) real emits NO
  same-tab accountAdded/accountRemoved, and its EventType map has NO
  LOGIN_FAILURE/ACCOUNT_ADDED/ACCOUNT_REMOVED — mini's EventType now equals
  real's 20-key map (also serves C5); loginSuccess payload is the ACCOUNT (not
  the result) and fires only when the account count grew; failures emit only
  acquireTokenFailure. (3) loginPopup = acquireTokenPopup with correlationId
  stamped on the request; silent events fire once per deduped request;
  acquireTokenFromNetworkStart fires after the AT rung misses (any non-
  AccessToken policy) with the initialized request (account/
  authenticationScheme/authority/correlationId/forceRefresh/redirectUri/
  scopes). (4) logoutPopup order = logoutStart → clear cache → +state on
  request → logoutSuccess → popupOpened → logoutEnd (cache now cleared BEFORE
  the popup roundtrip, like real). (5) FRONT-RAN one B4 line: forged-state
  redirect now resolves null silently (real emits only handleRedirectStart/
  End) — core.redirect-state-tampered passes. (6) React bindings updated:
  redirect-result/error capture now keys off acquireTokenSuccess/Failure with
  interactionType redirect (LOGIN_FAILURE is gone; loginSuccess payload is an
  account). EventMessage = {eventType, interactionType, payload, error,
  timestamp} — no correlationId field yet (unobserved by harness; add if a
  scenario ever compares it). InteractionType gained None:"none" (C5 item).

- **B3 (2026-07-10)**: protocol params. KEY DISCOVERY: real browser 5.16 sends
  `x-client-current-telemetry: ""` / `x-client-last-telemetry: ""` — LITERAL
  empty strings, even after failures (StubServerTelemetryManager is what
  StandardController uses; verified in ServerTelemetryManager.mjs + every
  snapshot incl. last-telemetry-after-failure). C3's "5|apiId,…" format never
  hits the wire — telemetry.token-request-headers now passes with "" values.
  Implementation: one `WIRE_ID`/`TOKEN_TELEMETRY` constants block (Decision
  Log impersonation). Authorize query adds nonce (GUID), client_info=1,
  client-request-id (request correlationId; authorizeUrl generates when
  absent and returns it so callers reuse ONE cid per request), claims
  (default literal, signin_state before login_hint — string-compared),
  clidata=1, x-client-SKU/VER, login_hint falls back to account.username,
  X-AnchorMailbox ccs = account ? `Oid:<localAccountId>@<tenantId>` :
  loginHint ? `UPN:<hint>` : absent (account wins). Token requests add the
  same claims/client_info/ccs + TOKEN_TELEMETRY in the BODY,
  client-request-id on the QUERY string, content-type gains `;charset=utf-8`.
  ccs/nonce thread via AuthCodeResponse (popup/iframe) and the msal.request
  temp entry (redirect). RT grant redeems against the REQUEST's redirectUri.
  id_token nonce now VALIDATED on auth-code redemption (`nonce_mismatch`
  ClientAuthError, only when we sent one — mock IdP echoes nonce; RT grants
  skip it like real). Remaining telemetry.last-telemetry-after-failure diff
  is NOT about telemetry params: real's failedCall SUCCEEDED (tokenCalls 2)
  — real appears to fall back to iframe when the RT grant fails with the
  injected invalid_grant, mini surfaces the error; investigate in C3.

- **B4 (2026-07-10)**: storage metadata + logout params. `msal.version`
  written at initialize (sessionStorage — C6 must route it to the configured
  cacheLocation; accounts.local-storage gained 1 diff from this, 13→14, still
  a C6 item). Every entity write stamps `lastUpdatedAt` (epoch-ms string);
  account entity additionally gets `cachedByApiId` = real's ApiId of the
  redeeming flow (popup 862, handleRedirectPromise 865, ssoSilent 863, silent
  iframe rung 864, RT grant 61 — only 862 is snapshot-observable, rest set
  for true parity). end_session now carries `client-request-id` (request
  correlationId) + `state` (real's lib-state format: base64 `{id,meta:
  {interactionType}}`); logoutRedirect/logoutPopup honor per-request
  postLogoutRedirectUri + correlationId. INFRA FIX (determinism): mock IdP's
  id_token nonce was a single global `lastNonce` — a stray late /authorize
  (abandoned iframe from a previous scenario) landing mid-roundtrip crossed
  wires and made mini's nonce validation throw `nonce_mismatch`
  (params.scopes-normalization flaked in full runs, passed in isolation).
  Now the nonce is keyed by PKCE code_challenge at /authorize and resolved
  from SHA256(code_verifier) at /token — per-request exact, no ordering
  sensitivity. Snapshots unaffected (nonce/JWT normalized); real re-checked
  green on roundtrip/sso/scopes scenarios after the change.

- **C1 (2026-07-10)**: request passthrough. KEY DISCOVERY: real browser 5.16
  IGNORES `tokenQueryParameters` at runtime (the option exists only in
  typings — zero grep hits in msal-common/msal-browser dist); instead
  `extraQueryParameters` ride the token-endpoint QUERY string too
  (createTokenQueryParameters: extras first, then client-request-id) — the
  snapshot proves it (tokenQuery has dc/slice, `tokenslice` never hits the
  wire). Mini keeps `tokenQueryParameters` in the TokenRequest type for API
  compat but sends nothing. Claims: mergedClaims() = real's buildMergedClaims
  (request claims parsed, default id_token claims appended without overwrite,
  clientCapabilities → access_token.xms_cc.values), byte-identical
  JSON.stringify output to the old DEFAULT_CLAIMS literal in the no-input
  case; sent on authorize + token body. Hints: sid only when prompt=none and
  it suppresses login_hint; prompt=select_account suppresses all hints AND
  the X-AnchorMailbox ccs (real's Authorize.mjs hint ladder); silentFrame now
  passes prompt via the request (`prompt: req.prompt ?? "none"`) so the sid
  gate sees it — authorizeUrl's separate `extra` param is gone from
  ClientContext. Custom state: wire = `btoa({id:<uuid>})` + `|<custom>` when
  present (real's setRequestState; wireStateEqualsCustom false), result.state
  echoes only the custom part — threaded as `userState` through
  AuthCodeResponse + msal.request (redirect leg). Per-request authority: NO
  re-discovery (real reuses cached metadata — snapshot discoveryPaths is []);
  endpoints get the config-authority prefix swapped for the request authority
  (no-op on the mock IdP whose endpoints are tenant-less; correct for AAD
  tenant-in-path URLs); result.authority + acquireTokenNetworkStart payload +
  AT-cache-hit results all report the canonical request authority.
  claims/eqp/authority thread through AuthCodeResponse (popup/iframe),
  msal.request (redirect), and redeemRefresh meta (RT grant).

- **C2 (2026-07-10)**: resilience. Throttle cache = real's ThrottlingUtils in
  mini's `post()`: key `throttling.<JSON RequestThumbprint>` (clientId,
  authority-with-slash, scopes, homeAccountIdentifier, claims,
  authenticationScheme — undefined dropped), written on 429 / 5xx /
  Retry-After+non-2xx with `{throttleTime,error,errorCodes,errorMessage,
  subError}`; throttleTime = min(now+RetryAfter||60s, now+3600s) in ms.
  Blocked retry re-throws `ServerError(errorCodes.join(" ")||"",
  errorMessage, subError)` — errorCode is EMPTY STRING (real preProcess),
  not the original error code, and the message is the RAW error_description
  (no "Error(s): …" formatting — that only wraps live token responses).
  Expired entries removed at pre-check. `refresh_in` → `refreshOn` stored on
  the AT entity (epoch-s string) + result.refreshOn Date (network + cache-hit
  paths). NOTE: real 5.16 fires NO observable background refresh when
  refreshOn has passed (snapshot backgroundTokenRequests=0, fromCache:true) —
  mini implements none. Real's Authorize-side removeThrottle-on-interactive-
  success NOT implemented (unobservable in scenarios; add in C10 sweep only
  if something compares it). Throttle store is sessionStorage directly —
  C6 must route it through the configured cacheLocation like real.

- **C3 (2026-07-10)**: silent iframe-fallback breadth now = real's
  checkIfRefreshTokenErrorCanBeResolvedSilently (StandardController): after an
  RT-rung failure, fall back to the hidden iframe ONLY when the policy allows
  iframe renewal (Default/RefreshTokenAndNetwork/Skip — mini's `useFrame`
  already encoded this) AND the error is resolvable: (non-IRAE OR IRAE with
  subError `bad_token`) with errorCode `invalid_grant`/`token_refresh_required`,
  OR errorCode `no_tokens_found`/`refresh_token_expired`. Two-sided change vs
  mini's old "fall back on any IRAE": a ServerError invalid_grant RT failure
  now RECOVERS via iframe (telemetry.last-telemetry-after-failure — real's
  tokenCalls=2), while a true interaction_required IRAE from the RT grant now
  surfaces immediately instead of burning an iframe attempt.

- **C4 (2026-07-10)**: perf events live in a new feature module
  `@mini-msal/browser/telemetry` (pay-to-play): `telemetry(ctx)` wraps the
  client's initialize/acquireTokenSilent/ssoSilent/acquireTokenPopup AFTER
  other features attach (compat composes `[popup, telemetry]` — order
  matters), emitting one event per top-level call
  ({name, correlationId, durationMs, success, errorCode?}) via the opt-in
  `config.telemetry.client` (`BrowserPerformanceClient`, exported from the
  subpath + re-exported by compat). Like real, no opt-in client → 
  addPerformanceCallback registers but nothing ever fires. Event
  correlationId = result's on success / request's on failure (correlation-id
  scenario requires the request CID to surface). Names kept to real's
  observed set; acquireTokenRedirect NOT wrapped (real's redirect perf event
  is "acquireTokenPreRedirect", unobserved — add in C10 only if compared).
  Core paid zero runtime bytes (only a `PerfClient` type + `telemetry?`
  config key; mini-core size unchanged at 17.4 KB min).

- **C5 (2026-07-10)**: init & surface + THE BRIDGE MIGRATION. Mini's popup and
  hidden-iframe flows now complete via the redirect-bridge (BroadcastChannel
  keyed by the lib-state id, `{v:1,payload}` message), NOT URL polling — real
  v5's actual mechanism. Wire state is now real's full lib-state
  `btoa({id,meta:{interactionType}})` (+`|custom`). New subpath
  `@mini-msal/browser/redirect-bridge` (`broadcastResponseToMainFrame()`);
  new rspack variant `mini-redirect-bridge` (0.6 KB min vs real's 6.5 KB);
  postbuild serves it at `/mini-popup.html`, `dist/{conformance-mini,
  mini-mock-app}/popup{,2}.html`; sync-variants rewrites mini apps'
  `/popup.html` → `/mini-popup.html`; `/blank.html` stays bridge-less ON
  PURPOSE (init.popup-without-bridge depends on it). A redirect page without
  the bridge (or a closed popup) now just times out
  (`timed_out`/`redirect_bridge_timeout`). Logger: implemented IN CORE, not
  as the planned `/logger` feature module — the call sites must live in core
  anyway, and the sink is ~10 lines (`log(level,msg)` gated by
  `system.loggerOptions.logLevel`, default Info(2)); real logs obfuscated
  short codes, message text is free-form. getConfiguration(): only
  snapshot-observed defaults + user passthrough; unset optionals stay
  undefined (real's 5.16 defaults have NO navigateToLoginRequestUrl/
  storeAuthStateInCookie/hash-timeout keys). Surface: version +
  BrowserCacheLocation/ProtocolMode/PromptValue/OIDC_DEFAULT_SCOPES/
  BrowserAuthErrorCodes (52 keys, packed as a split-string in compat) +
  createStandardPublicClientApplication (= createAuth) +
  createNestablePublicClientApplication (= createAuth until C9) +
  isPlatformBrokerAvailable (async false until C7). Free extra passes:
  broker.is-platform-broker-available, naa.no-bridge-fallback. KNOWN COST:
  naa.get-token-popup + naa.error-mapping now ERR at ~60s each (scenario
  installs a fake NAA bridge; the nestable factory returns a standard PCA
  whose popup waits out the 60s bridge timeout) — full runs pay +2 min until
  C9 lands; not a pass regression.

- **C6 (2026-07-10)**: localStorage + cross-tab. Core gained a `Store` seam
  (plaintext get/set/remove for indexes/msal.version/throttle entries +
  getUser/setUser for entities; default = sessionStorage) — new feature
  `@mini-msal/browser/local-storage` swaps in real's encrypted store when
  `cacheLocation: "localStorage"`: HKDF base key in the `msal.cache.encryption`
  session cookie ({id,key}, Secure, SameSite=None), per-write AES-GCM
  (HKDF salt = random 16-byte nonce, info = clientId-if-key-contains-it, zero
  IV — real's exact scheme), `{id,nonce,data,lastUpdatedAt}` wrappers,
  in-memory plaintext mirror imported at initialize (foreign-key entries
  pruned + indexes rewritten, real's importExistingCache) and synced across
  tabs via the `msal.broadcast.cache` channel. Temp state (interaction lock,
  msal.request) stays in sessionStorage ALWAYS, like real's
  temporaryCacheStorage — in LS mode sessionStorage dumps empty. Cross-tab
  EVENTS live in core: real's EventHandler posts loginSuccess/logoutSuccess/
  activeAccountChanged on `msal.broadcast.event` UNCONDITIONALLY (all cache
  modes) but subscribes only in localStorage mode
  (StandardController.initialize) — mini mirrors both sides. Entity writes are
  now awaited (`writeUser`) since encryption is async. Throttle +
  active-account + msal.version now route through the configured location
  (closing the B4/C2 routing notes). Core cost: +0.4 KB min (seam + event
  bus); the crypto lives in the feature module (compat +2.8 KB min total).

- **C7 (2026-07-10)**: DOM broker = new feature `@mini-msal/browser/broker`
  (compat composes [localStorageCache, popup, broker, telemetry]). Core seams
  added: `ctx.writeAccount(entity)` (entity + account.keys index through the
  Store seam) and `ctx.nativeSilent?(req, account)` — acquireTokenSilent runs
  `nativeSilent ?? silentLadder` inside the deduped promise, so broker silent
  keeps start/success events and dedupe but emits NO acquireTokenFromNetworkStart
  (matches real); toAccountInfo now surfaces the entity's nativeAccountId.
  KEY DISCOVERIES from real 5.16: (1) acquireTokenByCode({nativeAccountId})
  emits acquireTokenStart with the RAW request (no correlationId stamp) and NO
  success event on the native path (success only exists in the hybrid
  spa-code branch); (2) brokered AuthenticationResult is a 14-key shape
  (authority/uniqueId/tenantId/scopes/account/idToken/idTokenClaims/
  accessToken/fromCache/expiresOn/tokenType/correlationId/state/
  fromPlatformBroker) — NOT the web 21-key shape; byCode result.state =
  response.state || "" (string), silent overrides state to undefined (key
  present); scopes INCLUDE offline_access (response.scope verbatim);
  (3) ACCOUNT_UNAVAILABLE maps to IRAE with errorCode
  `native_account_unavailable` but message aka.ms#account_unavailable (code
  and anchor differ — real passes the broker's code to getDefaultErrorMessage);
  DISABLED falls to the default branch → NativeAuthError(code, raw
  description), fatal → provider dropped, next call
  `unable_to_acquire_token_from_native_platform`; (4) DOM request = named
  protocol fields + every remaining TRUTHY request prop stringified into
  extraParameters (falsy dropped — that's why prompt/tokenType appear only on
  silent and extendedExpiryToken:false never shows) + telemetry:"MATS" +
  x-client-xtra-sku. Simplifications (all scenario-unobservable): no native
  in-memory token cache (real caches broker ATs in nativeInternalStorage —
  mini re-calls executeGetToken every time); popup/ssoSilent/redirect broker
  routing not implemented (dom-first-login proves cold-cache popup goes web
  anyway); fatal-on-silent doesn't convert to token_refresh_required+iframe
  fallback; isPlatformBrokerAvailable stays async false (real gates the DOM
  probe on its domConfig ARGUMENT, which callers don't pass — snapshot is
  false/false even with a DOM broker installed). NativeAuthError exported
  from ./broker + re-exported by compat (exported-surface checks specific
  keys only — verified no diff).

- **C8 (2026-07-12)**: extension transport joined ./broker behind a `Provider`
  seam ({sku, send(nativeReq)}) — DOM keeps its filtered executeGetToken shape,
  extension posts the WHOLE initRequest object verbatim as body.request over
  the MessageChannel port (that's why the snapshot's requestKeys include
  undefined-valued keyId/prompt/tokenType — structured clone preserves them;
  never JSON.stringify that request). Probe order = real's
  getPlatformAuthProvider: DOM (if experimental flag) → extension preferred id
  → extension undefined id, all errors swallowed. Bounce detection: our own
  Handshake seen back on the window (bubble listener, source===window,
  matching responseId) → not-installed, fail fast; else
  system.nativeBrokerHandshakeTimeout (default 2000ms, type added to core
  config). x-client-xtra-sku is per-transport (makeExtraSkuString): extension
  = `msal.js.browser|5.16.0,|,<name>|<version>,|` with name "chrome" iff the
  preferred extension id answered ("unknown" otherwise), version from the
  HandshakeResponse; DOM keeps its literal. Extension results are snake_case
  (access_token/scope/expires_in…) and get mapped to the DOM camelCase shape
  before the shared handleResponse; error blobs ({code, description,
  ext:{status,error}}) route through the same mapError as DOM (fatal DISABLED
  drop included). Both extension scenarios passed first try; DOM scenarios
  untouched.

- **C9 (2026-07-12)**: NAA = new subpath `@mini-msal/browser/naa` exporting
  `createNestableClient(config, fallback)` — the fallback factory keeps the
  module tree-shakable (compat passes `createAuth`). Bridge handshake at
  create; success → `createClient(config, [naa(bridge)])` with the naa
  feature OVERRIDING the client's methods (popup/silent → bridge, redirect/
  logout/byCode/perf-callbacks/clearCache → sync-throw NestedAppAuthError
  `unsupported_method`, handleRedirectPromise → null); any handshake failure
  → fallback. Core gained two seams for hydration: `ctx.findToken(type,
  match)` + `ctx.writeTokens({...})` (+0.7 KB min on mini-core — the naa
  read/write path routes through the Store seam so localStorage mode works).
  KEY DISCOVERIES from real 5.16: (1) tokenParams.extraParameters is a
  Map in real, so it JSON-serializes to `{}` — request extraQueryParameters
  NEVER hit the bridge wire (mini sends a literal `{}`); (2) cache lookups
  use the bridge's accountContext (or the last token response's), NEVER
  request.account — that's why the scenario's first silent call (raw account
  object) goes to the bridge; (3) cached silent results report the AT
  entity's ENVIRONMENT as result.authority (`toAuthenticationResultFromCache`
  — snapshot shows "localhost:4599", not the config authority); (4) NAA
  results have NO fromPlatformBroker key and tokenParams scope is
  request.scopes verbatim (no OIDC append); NO_NETWORK/USER_CANCEL mappings
  IGNORE the bridge's code (fixed codes), IRAE/ServerError use code +
  description-as-message. Simplifications (unobservable): mini's NAA client
  reuses core initialize, so it fetches discovery (real NAA never contacts
  the authority); real's double ACQUIRE_TOKEN_SUCCESS on cache hits and
  FAILURE-with-null-error on cache miss are not mirrored (single
  START/SUCCESS/FAILURE stream). All 6 naa.* passed first try; the two
  ~60s ERR scenarios from C5 are gone (full run ~2 min faster).

- **C10 (2026-07-12)**: final sweep, all green: full conformance mini 75/75,
  determinism `conformance:check` 75/75, e2e 25/25, GAP_REPORT regenerated
  all-pass. Size matrix gained the missing apples-to-apples row: new rspack
  variant `mini-compat` (32.5 KB min / 11.0 gz / 9.9 br) built from
  `test/apps/mini-only.ts`, GENERATED from `msal-only.ts` by sync-variants
  (import swapped to `@mini-msal/compat`) — the no-React counterpart of
  msal-browser-core (220.5 KB → 6.8× smaller). Verified `mini-core.ts` still
  exercises sign-out (logoutRedirect), silent + ssoSilent + redirect fallback,
  and multi-account APIs (getAllAccounts/getAccount filters/active account).
  Docs refreshed: docs/README.md (headline sizes, compat in the package tree,
  bridge-based architecture prose, implemented list now includes
  broker/NAA/LS/telemetry, final size matrix in Status), root README
  (75/75 + new sizes), bundle-size-experiment.md (final results table +
  variant rows; old 2026-07-01 numbers kept as historical note).
- **C11 (2026-07-13)**: perf-event field parity. Two new scenarios
  (`telemetry.perf-event-shape`, `-shape-silent`) pin real's FULL emitted
  event shape for six flows: initialize, popup login, silent cache-hit,
  silent RT-refresh, ssoSilent, silent failure (RT invalid_grant → iframe
  login_required — chosen over the first capture's 10s
  redirect_bridge_timeout for speed/determinism). Digest rules: harness now
  keeps raw clones in `__cap.perfRaw` (undefined-valued keys preserved as
  null); scenario-side normalizer types-out volatile values (ids, clocks,
  networkRtt/effectiveType, logs, errorStack), compares `ext` by KEY SET
  (values → `<number>`), and `context` by string presence only — decision:
  real's context is a JSON dump of its internal call tree; replicating that
  tree in mini is bytes with no consumer value (pipelines key on fields,
  not the debug blob). Everything else — ~40 top-level keys AND values per
  flow — must match exactly. Gotcha for future ext work: the runner's
  B64_43_RE normalizes ANY 43-char key name (`generateCodeChallengeFrom-
  VerifierDurationMs`, `silentFlowClientAcquireCachedTokenCallCount`,
  `silentHandlerMonitorIframeForHashDurationMs` are all 43 chars →
  `<b64-43>`, colliding keys collapse — first capture read misled exactly
  this way). Implementation (telemetry.ts rewrite): per-flow shape tables —
  ext name groups as compact strings (DISC/PKCE/CODE/STD/SIL/RT/IFRAME +
  `!` = CallCount-only marker), field sets per flow with live values: token
  sizes from result + cache (`ctx.findToken`, incl. RT secret length),
  account type from claims (B2C/MSA/AAD like real's getAccountType),
  cacheMatchedAccounts/accessTokensRemoved from pre-call cache snapshots
  (absent on cache hits / when 0, like real), accountCachedBy via in-memory
  homeAccountId→API map (popup → `acquireTokenPopup`, silent refresh →
  `acquireTokenSilent_silentFlow`), authorityEndpointSource via a
  first-resolve-per-instance flag (first acquire event = `network` + ext
  gains authorityGetEndpointMetadataFromNetwork, then `cache`), migration/
  instance counters on initialize, `navigator.connection` for
  networkRtt/effectiveType. New core seam: silent ladder attaches
  `silentRefreshReason` (the RT error code that triggered the iframe
  fallback) to the error the iframe leg throws. Identity: `LIB_NAME =
  "@azure/msal-browser"` exported from the index.ts impersonation block.
  Approximations where the suite can't observe (recorded, not tested):
  visibilityChangeCount/onlineStatusChangeCount constant 0, deduped
  constant false, popup/sso failure shapes + non-iframe silent failures
  approximated, context string is mini-shaped. acquireTokenPreRedirect NOT
  emitted: it fires only via the onRedirectNavigate hook pre-navigation
  (unobservable until C13's seam exists; C20 owns redirect perf events).
  Size: compat 32.5 → 39.7 KB min (+7.2 KB = the shape tables; pay-to-play
  holds — mini-core unchanged at 19.5, only /telemetry composers pay).
- **C12 (2026-07-13)**: logger + error-code namespace API surface (audit
  findings `get-logger`, `error-code-namespace-exports`). New scenario
  `init.exported-surface-2` (suite 77→78) pins: full LogLevel enum incl. TS
  reverse mappings, WrapperSKU, all five *ErrorCodes namespaces
  KEY-AND-VALUE exact (37/24/9/2/3 entries), BrowserConfigurationAuthError
  shape, Logger level-gate + clone behavior, and
  getLogger/setLogger/initializeWrapperLibrary on the instance — mini green
  on FIRST run after implementation. Placement decisions: Logger, LogLevel,
  WrapperSKU + the three instance methods live in CORE (getLogger is core
  client surface; internal `log()` now routes through a swappable Logger so
  setLogger really redirects library logs; message format = real's
  `[UTC] : [cid] : pkg@ver : Level - msg`, level default Info(2)). The five
  *ErrorCodes namespaces + BrowserConfigurationAuthError are COMPAT-only
  (drop-in surface packing, same as BrowserAuthErrorCodes precedent) via a
  shared `pack()` helper — irregular keys that don't camelize from their
  code are written `key=code`: cannotAppendScopeSet, emptyInputScopeSet,
  endpointResolutionError, misplacedResourceParam, openIdConfigError,
  tokenClaimsCnfRequiredForSignedJwt (ClientAuth) and urlEmptyError
  (ClientConfiguration). initializeWrapperLibrary only STORES the wrapper
  sku/version: real forwards it to server-telemetry headers, but browser
  5.16 stubs those to empty strings on the wire (finding's
  verdict.corrections) — C19's server-telemetry task owns populating them.
  Size: compat 39.7 → 42.9 KB min (+3.2 = Logger class + namespace
  tables), mini-core 19.5 → 20.9 (+1.4, the core-resident Logger — first
  core growth since C8; accepted as core API surface, not a feature).
- **C13 (2026-07-13)**: redirect navigation seams (audit findings
  `navigate-to-login-request-url`, `on-redirect-navigate`,
  `on-redirect-navigate-hook`, `navigation-client`,
  `set-navigation-client`). Five new scenarios in NEW area 11
  `11-navigation.mjs` (suite 78→83); ALL green on first mini run after
  implementation. One seam, real's shape: `NavigationClient` class in CORE
  (default navigateInternal/External = location.replace when noHistory else
  assign, returning a promise that only REJECTS after `timeout` — so
  loginRedirect now stays pending during a real navigation, exactly like
  real), consulted for every redirect navigation with real's options
  ({apiId 861/865/961, timeout system.redirectNavigationTimeout(30s),
  noHistory}); swappable via system.navigationClient AND
  setNavigationClient(). auth.onRedirectNavigate(url) fires before
  acquire/logout navigation; false cancels — acquire leaves the interaction
  lock HELD (real does), logout releases it and emits logoutEnd.
  navigateToLoginRequestUrl (default true): acquireTokenRedirect caches the
  start page under real's temp key `msal.{cid}.request.origin`
  (request.redirectStartPage honored); handleRedirectPromise compares
  normalized URLs (normUrl = drop hash + trailing-slash pathname), and on
  mismatch caches the response hash under `msal.{cid}.urlHash`, clears the
  URL hash, and navigateInternal(origin||homepage, noHistory:true) — result
  is delivered by the NEXT load's handleRedirectPromise from the cached
  hash; navigateInternal returning false processes in place (real's
  contract). In-place branch restores the initiating page's own #hash
  (real's replaceHash). document.title = "Microsoft Authentication" during
  processing, restored in finally. Match-real event fix: redirect
  logoutStart now carries the RAW logout request (null payload for
  `logoutRedirect()`) — real only synthesizes the payload on the popup
  path; the popup snapshot (account/correlationId/postLogoutRedirectUri)
  still passes. unlock() moved from top-of-processRedirect into the
  process/clean/forged legs so the lock survives the mid-replay load
  (pinned by the deep-link scenario's afterReplay probe). e2e fix (not a
  regression): the synthetic cancelled-login check seeded only
  `msal.request`; a genuine mini flow now also caches request.origin, so
  the seed gained that key — without it mini (like real) replays to the
  homepage. mini's other temp key `msal.request` (params blob) still
  differs from real's `msal.{cid}.request.params` schema — the new probes
  compare only the two replay keys; full temp-key schema parity was NOT in
  the audit findings and stays a non-goal. Size: compat 42.9 → 44.6 KB min
  (+1.7), mini-core 20.9 → 22.5 (+1.6 — NavigationClient + replay live in
  core since redirect is core surface).

- **C14 (2026-07-13)**: popup behaviors (audit findings `navigate-popups`,
  `logout-popup-main-window-redirect`). Four new scenarios (suite 83→87):
  `core.popup-open-timing` / `-async` / `core.popup-window-attributes` in
  01-core, `accounts.logout-popup-main-window-redirect` in 03-accounts —
  all green on first mini run after implementation. New lib helpers
  `armOpenRecorder`/`openCalls` wrap window.open and persist calls in
  sessionStorage (survives the mainWindowRedirectUri navigation); `sync`
  flag = call landed inside the scenario's `__inApiCall` span, i.e. inside
  the user-gesture window that popup blockers honor. Behavior, real's
  shape: `system.navigatePopups` (default TRUE) sync-opens `about:blank`
  in the caller's stack for acquireTokenPopup/loginPopup AND logoutPopup,
  then navigates it via `location.assign` (real's openPopup incl.
  document.title = "Microsoft Authentication" + focus); false defers the
  open to the final URL. openSizedPopup parity: features string with
  spaces (`width=483, height=600, top=…, left=…, scrollbars=yes`),
  request `popupWindowAttributes` popupSize/popupPosition clamped to the
  parent window, centered defaults, `popupWindowParent` honored. Popup
  names: token `msal.{clientId}.{scopes.join("-")}.{authority}.{cid}`,
  logout `msal.{clientId}.{homeAccountId}.{cid}` (correlationId now
  pre-generated in popup.ts and passed into authorizeUrl). A blocked sync
  open (null) is carried like real: no immediate throw — the flow fails
  at navigate time with popup_window_error (errors.popup-blocked
  unchanged), and the pre-opened popup is closed on any failure.
  logoutPopup `mainWindowRedirectUri`: after the popup roundtrip the MAIN
  window navigates through the NavigationClient seam — new core
  `ctx.navigate(url, apiId)` (ApiId.logoutPopup 962, noHistory:false);
  snapshot pins that the interaction lock SURVIVES into the next page
  (real never reaches its unlock either) — matched by awaiting the
  never-settling navigation inside try. telemetry `isAsyncPopup` now
  `navigatePopups === false` instead of hardcoded. Size: compat 44.6 →
  45.9 KB min (+1.3), mini-core 22.5 → 22.6 (+0.1 — ctx.navigate seam).

- **C15 (2026-07-13)**: React bindings parity (audit findings
  `msalprovider-initialize`, `inprogress-interaction-statuses`,
  `usemsalauthentication-acquiretoken`,
  `template-render-prop-children-and-account-props`,
  `msal-authentication-template-error-contract`,
  `account-identifier-matching`). NEW test surface: react harness pages
  `conformance-react-{real,mini}` (React BUNDLED — not size targets),
  shared glue `test/apps/harness/react-setup.tsx` on top of setup.ts.
  Because page.evaluate can't pass components, fixtures live in the
  harness and scenarios `__mount(name, props)`: an always-mounted
  `<Status>` recorder renders `|status:<inProgress>;accounts:<n>|` into
  #root and appends distinct inProgress values to `__renderLog`; other
  fixtures expose `__probeLog`/`__hook`. Scenarios (new area 12-react)
  retarget ctx at `ctx.reactHarnessUrl` and wait on #root textContent
  with a tolerate-timeout helper so a diverging stack yields a stable
  diff, not a scenario error. Six scenarios, deterministic (real --check
  ×2), ALL green on the first mini run after the rewrite.
  `packages/react/src/index.tsx` rewritten as a compact port of
  @azure/msal-react 5.5.1 (exports `version` "5.5.1", clones the logger
  as "@azure/msal-react" — identity-impersonation decision): provider
  `initialize()` → `handleRedirectPromise()` → UNBLOCK fallback,
  `initializeWrapperLibrary(WrapperSKU.React, version)`, reducer ports
  EventMessageUtils.getInteractionStatusFromEvent (clear-guards,
  interactive-only AcquireToken, RESTORE_FROM_BFCACHE) and keeps accounts
  `[]` until startup completes (kills the cached-user startup flash);
  hooks/templates as in real (case-insensitive id matching, empty
  filter → getActiveAccount, accountInfoIsEqual iat/nonce, acquireToken
  callback with OIDC default scopes + correlationId + IRAE fallback +
  ReactAuthError codes, result reset when the account disappears, throw
  error without ErrorComponent, spread contract into Error/Loading
  components, function-as-children). Core gained the 5-value
  `InteractionStatus` export (re-exported by compat). Gotchas learned:
  (1) logout POPUPS complete via the bridge — `logoutPopup` in a react
  scenario must pass `postLogoutRedirectUri: ctx.popupUrl` or it never
  resolves; (2) NEITHER stack auto-sets the active account after
  loginPopup, so useMsalAuthentication auto-acquire tests must
  setActiveAccount first; (3) hook requests need `redirectUri:
  ctx.popupUrl` too, else real's silent iframe lands bridge-less and
  fails timed_out (nondeterministic) instead of login_required; (4)
  scrubString needed a `conformance-react-(real|mini)` rule BEFORE the
  existing one (substring wouldn't match). Size: mini-msal-stack 53.0 →
  55.8 KB min (+2.8 — the react port), compat UNCHANGED 45.9, mini-core
  UNCHANGED 22.6 (InteractionStatus tree-shakes out of non-react builds).

### C16 — 2026-07-13 — cache entity semantics (dedupe, tenant merge, migration, KMSI)

Suite 93→98 (new area 13-cache, 5 scenarios, all deterministic and all
green on the first mini run after implementation). Mock IdP gained
`/config?claims.<name>=<json-or-string>` overrides merged into id_token
claims (cleared by /reset) — used for the guest `tid` and
`signin_state:["kmsi"]`; token-response `scope`/`access_token` overrides
via the existing /config drive the superset-scope acquisition.

Decisions/details:
- **AT scope dedupe** (`dedupeATs`, used by tokenRequest AND
  ctx.writeTokens): matches real's saveAccessToken filter (clientId,
  homeAccountId, environment, realm, Bearer) + ScopeSet
  .intersectingScopeSets semantics — OIDC scopes
  (openid/profile/email/offline_access) are stripped from the NEW token's
  comparison set unless it is OIDC-only; any remaining overlap removes the
  cached AT before the new one is indexed. Lookup side: >1 request-matching
  ATs (scope-superset match, expiry ignored) are ALL removed and the silent
  ladder falls through to the network, exactly real's getAccessToken
  multi-match clear. The silent AT rung also gained real's realm filter
  (t.realm === account.tenantId) — needed so a guest-tenant AT can't
  satisfy a home-tenant request after the merge work.
- **Tenant-profile merge** (`mergeAccount` = real's buildAccountToCache):
  base entity looked up tenant-agnostically by homeAccountId+environment
  (>1 matches → ignore the hit, like real); incoming tenantProfiles
  appended if their tenantId is new; lastUpdatedAt + cachedByApiId
  refreshed on reuse (real stamps the CURRENT ApiId — snapshot shows 61
  after a silent merge over a popup login's 861). isHomeTenant is now
  COMPUTED (tid === homeAccountId utid segment) in all three writers —
  web (tokenRequest), broker, naa — the latter two previously hardcoded
  `true`. ctx.writeAccount routes through mergeAccount so features share
  the semantics. Public AccountInfo.tenantProfiles is now a Map keyed by
  tenantId (real's shape) and getAllAccounts flatMaps entities × profiles
  into per-tenant AccountInfos whose idToken/claims/username/name/
  localAccountId/tenantId/kmsi/loginHint/upn are sourced from that
  tenant's cached id token claims (real's updateAccountTenantProfileData
  precedence: claims > profile > entity).
- **KMSI** — Store seam extended: `setUser(key, value, kmsi?)`;
  ./local-storage writes KMSI entities PLAINTEXT (no AES-GCM wrapper), so
  they survive losing the per-session encryption cookie — verified by the
  scenario's cookie-delete + reload leg. kmsi computed once per token
  response (signin_state contains kmsi/dvc_dmjd, real's AuthToken.isKmsi)
  and threaded to all four entity writes. AccountInfo.kmsi now true/false
  when claims exist (real returns false for non-KMSI users, not
  undefined) — no snapshot pinned the old `undefined`, confirmed by the
  full-suite run.
- **Schema migration** is a NEW pay-to-play feature
  (`@mini-msal/browser/cache-migration`, composed by compat), not core —
  consumers without pre-v5 users don't pay for it. Core gained two tiny
  seams for it: `ctx.onInit(hook)` (runs during initialize after
  store.init, real's BrowserCacheManager.initialize ordering) and
  `ctx.getStore()`. The feature ports migrateExistingCache: schemas 0-2
  (msal.account.keys / msal.1.* / msal.2.*), stale-account removal past
  cacheRetentionDays (default 5, now a Config knob), lastUpdatedAt
  stamping, expired-AT/RT pruning (300s offset), id-tokens-first ordering
  so the KMSI map exists for AT/RT writes, tenantProfile reconstruction
  from claims, old entries left in place (real keeps them; only indexes
  are pruned). Simplification vs real: old entries that are encrypted
  wrappers are removed outright (real attempts decryption with the live
  cookie first) — unobservable unless a pre-v5 ENCRYPTED cache coexists
  with a live session cookie, which real itself only handles for
  same-session migrations.
- **Gotcha**: the new-AT write happens before dedupe runs, so dedupeATs
  must skip the just-written key (identical superset re-acquisitions
  share the key).
- Size: mini-msal-stack 55.8 → 60.6 KB min (+4.8), compat 45.9 → 50.6
  (+4.7 — merge/dedupe/expansion are core+compat, migration ~2.3 of it),
  mini-core 22.6 → 24.1 (+1.5). e2e 25/25 (v5 sessionStorage interop
  unaffected).

### C17 — 2026-07-13 — programmatic token APIs (clearCache, hydrateCache, loadExternalTokens, hybrid acquireTokenByCode)

Suite 98→103 (new area 14-token-apis, 5 scenarios, deterministic — three
identical real captures — and all green on the first mini run after
implementation). No mock-IdP changes needed: its /token endpoint redeems
any `mock-code-<idx>` without a prior /authorize, so hybrid-spa codes are
mintable "out-of-band" by just naming one. The area's network digest
EXCLUDES discovery requests (eager-vs-lazy discovery timing is C18's
concern).

Decisions/details:
- **clearCache(req?)** — core method (real has it on every PCA;
  SilentCacheClient.logout → clearCacheOnLogout). Account-scoped → existing
  clearAccount; no-account → clearAccount() PLUS real's
  browserStorage.clear() semantics: every remaining key containing "msal"
  or the clientId (incl. msal.version) is removed from sessionStorage
  (+localStorage only in localStorage mode, mirroring real's
  cache-location + temp-storage split; store.remove keeps ./local-storage's
  memory mirror in sync). Clearing the active account routes through
  setActiveAccount(null) → ONE activeAccountChanged, like real's
  removeAccount. No initialize requirement, no events otherwise, no
  navigation, no end_session (snapshot-pinned).
- **setActiveAccount event payload fix**: real's
  BrowserCacheManager.setActiveAccount emits activeAccountChanged with NO
  payload; mini was attaching the account. Only the new clear-cache
  scenario pins this (no other snapshot captures the event's shape);
  packages/react only matches on eventType.
- **hydrateCache(result, request?)** — core method. Account entity from
  result.account via new module-level entityFromAccountInfo (real's
  createAccountEntityFromAccountInfo): NO clientInfo, cloudGraphHostName/
  msGraphHost stamped (empty strings included — real assigns
  unconditionally), cachedByApiId 963, tenantProfiles Map → array. Then
  id+access token entities only (never a refresh token), KMSI-aware,
  expiresOn/extExpiresOn seconds from the result's Dates.
- **loadExternalTokens(config, request, response, options, features?)** —
  top-level core export (real's ITokenCache successor, ApiId 964). Builds
  a throwaway core client (captures ctx via a one-off feature) and
  initialize()s it — storage init + discovery, like real's standalone
  BrowserCacheManager/Authority setup (also writes msal.version, which
  real's standalone path does not; unobservable once any client
  initializes). Account from request.account OR client_info/claims
  (clientInfo kept on the entity, like buildAccountToCache); id/access/
  refresh entities written per presence (AT requires access_token +
  expires_in + scope, real's guards); returns real's
  generateAuthenticationResult shape (fromCache true, correlationId
  request's or "", state request's or ""). 5th param diverges from real
  (real: performanceClient; mini: features array so compat can compose
  cache backends — compat's re-export passes [localStorageCache,
  cacheMigration]). Guarded with Array.isArray so a drop-in consumer
  passing a real perf client is harmlessly ignored.
- **ctx.writeTokens generalized** (shared writeTokenEntities, also used by
  hydrateCache): all credentials optional (writes what's present), +
  extendedExpiresOn, refreshToken/foci (TokenEntity gained familyId), and
  kmsi plumbed to the Store seam. ./naa's call sites unchanged.
- **Hybrid acquireTokenByCode({code})** — stays in ./broker (task pointer;
  the compat surface owns the method either way). code+nativeAccountId →
  spa_code_and_nativeAccountId_present (mini previously took the broker
  path silently); code-only → ctx.redeem with ApiId 866; concurrent
  same-code calls share ONE promise (real's hybridAuthCodeResponses map,
  entry deleted on settle → a later call POSTs again). Events: START per
  call, ONE SUCCESS per redemption (deduped pair → 2×START + 1×SUCCESS,
  snapshot-pinned); failures emit FAILURE in the shared promise handler
  AND per-call outer catch, like real's double emission.
- **Wire**: real's hybrid redemption has NO code_verifier ("PKCE not
  needed") and NO redirect_uri (HybridSpaAuthorizationCodeClient
  includeRedirectUri=false) — redeem() branches on apiId 866 and post()
  now drops undefined body values (URLSearchParams would stringify them).
  Everything else rides the normal tokenRequest path: default claims,
  client_info=1, telemetry params, client-request-id query.
- Size: mini-msal-stack 60.6 → 62.5 KB min (+1.9), compat 50.6 → 52.5
  (+1.9), mini-core 24.1 → 25.6 (+1.5 — clearCache/hydrateCache/
  writeTokenEntities are core client surface; loadExternalTokens
  tree-shakes out of mini-core). e2e 25/25.

## Completed task list (Phases A0, A, B, C — all done)

## Task list (execute strictly top-to-bottom)

Statuses: `pending` | `in-progress` | `done <date> — pass X/75, mini-stack Y KB`

### Phase A0 — architecture (do this FIRST)

- [x] **A0** `done 2026-07-10 — pass 2/75, mini-stack 18.7 KB min / 6.9 gz` —
  Pay-to-play refactor (see Decision Log). The earlier "1 ERROR" was a
  first-load flake (harness page timeout on the run's first goto), not a
  regression: re-run gave 2 pass / 73 diff with per-scenario statuses AND
  diff counts identical to the pre-refactor baseline; e2e 25/25.

  1. Split `packages/browser/src/index.ts` into a core (`createClient`:
     config, discovery, redirect login + handleRedirectPromise, silent
     ladder cache→RT→iframe, accounts/active-account, events, errors,
     cache read/write in msal.3 schema) and a first feature module
     `@mini-msal/browser/popup` (loginPopup/acquireTokenPopup/logoutPopup +
     popup polling). Feature = function receiving the client's internal
     context (closure-based; no classes for seams).
  2. New `packages/compat` → `@mini-msal/compat`: exports
     `PublicClientApplication` (classic constructor + initialize()) that
     composes core + all current features, API-identical to today's class.
     Also re-export error classes/enums/EventType from browser so compat is
     a one-import drop-in.
  3. Repoint consumers at compat: `test/infra/sync-variants.mjs` replacement
     `"@azure/msal-browser"` → `"@mini-msal/compat"`;
     `test/apps/harness/harness-mini.ts` imports `@mini-msal/compat`;
     `@mini-msal/react` keeps working against the composed instance
     (it only uses the public instance API — verify).
  4. Size matrix: add a minimal core-only entry (redirect+silent, no popup,
     no react) as rspack variant `mini-core` and include it in
     `npm run measure` output.
  5. Validate: build clean; `npm run conformance:mini` — pass/diff counts
     IDENTICAL to baseline (2 pass / 73 diff, same per-scenario statuses:
     this is a pure refactor); e2e 25/25; record compat + core sizes below.
  All later tasks implement features as modules in their final homes; compat
  composes each new module as it lands.

### Phase A — bugs (mini misbehaves on claimed features)

- [x] **A1** `done 2026-07-10 — pass 8/75, mini-stack 19.2 KB min / 7.1 gz` — Error classes: set `.name` on AuthError/
  InteractionRequiredAuthError/BrowserAuthError; add ClientAuthError,
  ClientConfigurationError, ServerError, NestedAppAuthError classes with
  correct names; align error codes surfaced in scenarios (`no_account` →
  `no_account_error` BrowserAuthError, unknown-account → `authority_mismatch`
  ClientConfigurationError, 5xx/oauth-error → ServerError, aka.ms-style
  messages where snapshots expect them). Scenarios: core.sso-silent-cold,
  errors.* (most), resilience.5xx-server-error.
- [x] **A2** `done 2026-07-10 — pass 9/75, mini-stack 19.3 KB min / 7.2 gz` — Uninitialized guard: all public async APIs before
  `initialize()` throw BrowserAuthError
  `uninitialized_public_client_application`; handleRedirectPromise rejects
  too; getAllAccounts still works. Scenario: errors.uninitialized-client.
- [x] **A3** `done 2026-07-10 — pass 10/75, mini-stack 19.4 KB min / 7.2 gz` — Account filters: `getAccount({})` → null;
  `getAllAccounts(filter)` honors filter; username matching case-insensitive.
  Scenario: accounts.get-account-filters.
- [x] **A4** `done 2026-07-10 — pass 10/75, mini-stack 19.4 KB min / 7.2 gz` —
  Scope normalization: dedupe exact duplicates (preserve casing variants like
  real), no trailing space on empty scopes, OIDC defaults appended once,
  order = request order then defaults. All 8 scope diffs on
  params.scopes-normalization eliminated (18→10); the scenario can't PASS
  until B1 supplies result-shape fields (authority/correlationId/tokenType/
  state/fromPlatformBroker) — pass count unchanged by design, zero
  regressions. Result scopes now preserve granted-string casing
  (harness sorts them; real never lowercases).
- [x] **A5** `done 2026-07-10 — pass 11/75, mini-stack 19.5 KB min / 7.3 gz` — CacheLookupPolicy full semantics + forceRefresh:
  AccessToken(1): AT-or-throw ClientAuthError `token_refresh_required`;
  AccessTokenAndRefreshToken(2): AT→RT, no iframe; RefreshToken(3): RT only
  (skip valid AT), no iframe; RefreshTokenAndNetwork(4): RT→iframe;
  Skip(5): straight to iframe (no RT); forceRefresh bypasses AT (RT→iframe).
  Scenarios: silent.policy-*, silent.force-refresh.
- [x] **A6** `done 2026-07-10 — pass 12/75, mini-stack 19.7 KB min / 7.3 gz` —
  In-flight dedupe: concurrent identical acquireTokenSilent share one promise
  → one network call. Key matches real's thumbprint
  (scopes|homeAccountId|authority|claims — real includes NO policy or
  forceRefresh; StandardController.acquireTokenSilentDeduped), map entry
  deleted on settle. Scenario: silent.concurrent-dedupe.
- [x] **A7** `done 2026-07-10 — pass 13/75, mini-stack 20.0 KB min / 7.4 gz` —
  Interaction lock + unique popup names: second
  interactive call while one is pending → BrowserAuthError
  `interaction_in_progress`; popup window names unique per request
  (`msal.<guid>`-style) so popups never clobber; first call completes.
  Scenario: errors.interaction-in-progress.
- [x] **A8** `done 2026-07-10 — pass 13/75, mini-stack 19.9 KB min / 7.4 gz` — Don't auto-set active account on login (real never
  does); active account only via setActiveAccount. Check e2e still 25/25
  (the demo app sets active explicitly on login result — verify; if e2e
  depended on auto-active, fix the APP not the library... but app.tsx is
  shared with real MSAL and passes on real, so it must already handle it).
  Scenarios: accounts.multi-account (activeAfterSecondLogin),
  accounts.logout-popup-per-account (active), core.storage-shape-after-login.

### Phase B — systemic behavioral diffs

- [x] **B1** `done 2026-07-10 — pass 23/75, mini-stack 20.4 KB min / 7.5 gz` — Result shape parity: AuthenticationResult gains
  `authority` (canonical, trailing slash), `correlationId` (generate per
  request if not provided), `tokenType: "Bearer"`, `state` (custom state or
  ""), `fromPlatformBroker: false`, `expiresOn` semantics unchanged;
  `account.environment`; scopes keep REQUEST casing/order (real returns
  granted scopes minus offline_access, request-cased — study snapshots).
  Touches nearly every scenario's `result` block.
- [x] **B2** `done 2026-07-10 — pass 32/75, mini-stack 22.9 KB min / 8.2 gz` —
  Event stream parity: initializeStart/End, acquireTokenStart/Success/Failure
  (interactionType + real's payload keys), popupOpened, handleRedirectStart/
  End semantics (NO handleRedirect events on clean load — result null
  silently), logoutStart/Success/End for logout flows,
  acquireTokenFromNetworkStart on forced refresh, ordering matches snapshots.
  EventMessage gained interactionType. React layer updated (see Decision Log).
- [x] **B3** `done 2026-07-10 — pass 40/75, mini-stack 24.0 KB min / 8.6 gz` — Protocol params on authorize/token requests: `nonce`
  (send + VALIDATE id_token nonce claim on redemption), `client-request-id`
  (= correlationId) on authorize query AND token query string, `client_info=1`,
  `claims` default `{"id_token":{"signin_state":…,"login_hint":…}}` merged
  like real, `X-AnchorMailbox` (UPN: for loginHint, Oid: for account),
  `clidata=1`, scope param ordering `<resource scopes> openid profile
  offline_access`, account-derived login_hint on acquireTokenRedirect,
  token redemption `redirect_uri` = request's redirectUri. Compare snapshot
  `idp[*].query/body` blocks exactly (volatile values normalize away).
- [x] **B4** `done 2026-07-10 — pass 44/75, mini-stack 24.4 KB min / 8.7 gz` — Storage + misc behavior parity: write `msal.version`,
  per-entity `lastUpdatedAt` + `cachedByApiId`; ~~drop mini's `msal.meta.*`
  key~~ DONE in A5 (discovery now cached in per-instance memory only — the
  key is gone and every new instance re-fetches discovery like real); forged-state redirect →
  resolve null silently (no throw, no loginFailure event);
  logout: end_session gets `state` + `client-request-id` params, temp
  `msal.interaction.status` signout entry, honor per-request
  postLogoutRedirectUri; msal.request temp-cache naming if visible in dumps.
  Scenarios: core.storage-shape-after-login, core.redirect-state-tampered,
  accounts.logout-*, init.storage-before-login.

### Phase C — missing features

- [x] **C1** `done 2026-07-10 — pass 50/75, mini-stack 25.4 KB min / 9.0 gz` — Request passthrough: `sid`, `domainHint`
  (`domain_hint`), `extraQueryParameters` (authorize), `tokenQueryParameters`
  (token endpoint query), custom `state` (wire format `<libState>|<custom>`,
  echo on result.state), `claims` + `clientCapabilities:["cp1"]` → xms_cc
  merge on authorize+token, per-request `authority` override (discovery for
  that authority, result.authority reflects it, forceRefresh honored with it).
  Scenarios: params.* (all remaining).
- [x] **C2** `done 2026-07-10 — pass 52/75, mini-stack 26.2 KB min / 9.3 gz` — Resilience: 429 throttle cache (storage entry shaped
  like real's `throttling.*`, immediate retry re-throws same ServerError with
  NO network until Retry-After passes); ServerError formatting for 5xx
  (message format per snapshot); `refresh_in` → `refreshOn` on the AT entity
  + proactive-refresh observable behavior per snapshot
  (resilience.proactive-refresh: what real did — study snapshot first).
  Scenarios: resilience.*.
- [x] **C3** `done 2026-07-10 — pass 53/75, mini-stack 26.4 KB min / 9.3 gz` —
  MOSTLY DONE BY B3 (identity/telemetry/lib-capability body params,
  client-request-id query param, correlationId threading —
  telemetry.token-request-headers passes; real sends EMPTY telemetry values,
  see B3 log entry). Finished here: RT-failure → iframe fallback breadth now
  matches real's checkIfRefreshTokenErrorCanBeResolvedSilently (see Decision
  Log) — telemetry.last-telemetry-after-failure passes.
  telemetry.correlation-id-propagation's last diff is the perf event (C4).
- [x] **C4** `done 2026-07-10 — pass 56/75, mini-stack 27.1 KB min / 9.6 gz` — Perf events: `addPerformanceCallback` +
  BrowserPerformanceClient-equivalent opt-in via `telemetry.client` config;
  emit `initializeClientApplication`, `acquireTokenPopup`,
  `acquireTokenSilent` (+ cache-hit variant) events with
  name/success/durationMs/correlationId. Harness passes perfClient:true only
  when lib.BrowserPerformanceClient exists — export a compatible class.
  Scenarios: telemetry.perf-events-*.
- [x] **C5** `done 2026-07-10 — pass 62/75, mini-stack 29.3 KB min / 10.4 gz` — Init & surface: `getConfiguration()` returning
  real-default-shaped config (values per init.get-configuration snapshot);
  logger (`system.loggerOptions.loggerCallback`, Info+Verbose volume);
  initializeStart/End already from B2; export surface per
  init.exported-surface snapshot: version "5.16.0" (Decision Log:
  impersonation), createStandardPublicClientApplication,
  isPlatformBrokerAvailable, full EventType map incl. broker/bfcache strings,
  InteractionType.None, PromptValue, ProtocolMode, BrowserCacheLocation,
  OIDC_DEFAULT_SCOPES, BrowserAuthErrorCodes (52 keys), error classes.
  Scenarios: init.* remaining.
- [x] **C6** `done 2026-07-10 — pass 64/75, mini-stack 32.1 KB min / 11.3 gz` — localStorage + cross-tab: `cacheLocation:
  "localStorage"` with real's observable shape (encrypted entities
  `{id:<guid>, nonce:<b64>, data:<ciphertext>}` via AES-GCM, session key in
  a cookie like real — study real's LocalStorage/CookieStorage source),
  plaintext key indexes + msal.version in localStorage, account/logout events
  propagating to other tabs (storage listener). sessionStorage remains
  default; interop e2e must stay green. Scenarios: accounts.local-storage,
  accounts.cross-tab-events.
- [x] **C7** `done 2026-07-10 — pass 68/75, mini-stack 37.0 KB min / 12.6 gz` — Platform broker, DOM transport: config
  `system.allowPlatformBroker` + `experimental.allowPlatformBrokerWithDOM`;
  probe via navigator.platformAuthentication.getSupportedContracts
  ("MicrosoftEntra" → "get-token-and-sign-out"); `acquireTokenByCode({
  nativeAccountId })` → executeGetToken with real's request shape; response
  mapping → AuthenticationResult with fromPlatformBroker:true; cache
  nativeAccountId on the account; silent routes via broker when account has
  nativeAccountId; error mapping (USER_CANCEL→user_cancelled BrowserAuthError,
  ACCOUNT_UNAVAILABLE→native_account_unavailable IRAE, NO_NETWORK→
  no_network_connectivity, DISABLED→fatal NativeAuthError + drop broker +
  second call `unable_to_acquire_token_from_native_platform`). Protocol:
  docs/design/broker-protocol.md. Scenarios: broker.dom-*,
  broker.is-platform-broker-available.
- [x] **C8** `done 2026-07-12 — pass 70/75, mini-stack 38.8 KB min / 13.3 gz` — Platform broker, extension transport: Handshake via
  window.postMessage + MessageChannel (channel id, preferred extension id,
  2s default timeout `nativeBrokerHandshakeTimeout`, retry with undefined
  extensionId, swallow failures), GetToken over the port, Response/Success
  mapping identical to C7. First-login gating parity (no nativeAccountId →
  web flow). Scenarios: broker.extension-*, broker.dom-first-login.
- [x] **C9** `done 2026-07-12 — pass 75/75, mini-stack 39.5 KB min / 13.4 gz` — NAA: `createNestablePublicClientApplication` +
  window.nestedAppAuthBridge (GetInitContext handshake envelope with
  clientLibrary `msal.js.browser` / clientLibraryVersion `5.16.0` per
  Decision Log), GetToken/GetTokenPopup mapping, browser-cache-first silent, error
  mapping per docs/design/naa-protocol.md, unsupported APIs →
  NestedAppAuthError `unsupported_method`, no-bridge → standard PCA
  fallback. Scenarios: naa.*.
- [x] **C10** `done 2026-07-12 — pass 75/75, mini-stack 39.5 KB min / 13.4 gz` — Final sweep: full `npm run conformance` (expect
  75/75 mini pass), determinism `conformance:check` 75/75, e2e 25/25,
  `npm run measure` final sizes. **Final bundle-size retest (user-requested
  2026-07-10)**: publish a fresh measured matrix with at minimum (a)
  minimum-size basic-auth build — sign-in/sign-out (redirect) + silent +
  multi-account via `createClient` core only (`mini-core` variant; verify it
  still exercises sign-out and multi-account after all parity work), (b) full
  `@mini-msal/compat` build (all features composed), (c) compat + react
  stack, each min/gzip/brotli, side-by-side with real msal-stack numbers; regenerate GAP_REPORT (should show all
  pass); update docs/README.md status section + size claims, root README,
  bundle-size-experiment results table if size changed materially; final
  commit; write user-facing summary (do NOT reset-context after this one).

### C18 — 2026-07-13 — authority modes & discovery (lazy discovery, trust validation, metadata sources, instance-aware)

Suite 103→109 (new area 15-authority, 6 scenarios, deterministic — two
identical real captures — all green on the first mini run after
implementation, plus one follow-up fix for 3 regressed silent snapshots).
Mock IdP gained an `/authorize` instance-aware extension: when the request
carries `instance_aware`, the success fragment appends
`cloud_instance_host_name=localhost:4599`, `cloud_graph_host_name=
graph.cloud.test`, `msgraph_host=graph.test`.

Decisions/details:
- **protocolMode lives under config.SYSTEM, not auth** (real 5.16's
  buildConfiguration: `system.protocolMode`, default "AAD"; the
  `auth.protocolMode: "OIDC"` the suite's stdConfig always passed is
  silently ignored by real — so the whole existing suite actually ran in
  AAD protocol mode). Mini reads `config.system.protocolMode` only, and
  getConfiguration() now reports the "AAD" default. Snapshot-pinned by
  authority.oidc-discovery-endpoint-path (system.protocolMode OIDC +
  non-Microsoft host → discovery URL WITHOUT /v2.0/; AAD default → with).
- **Lazy discovery**: initialize() no longer fetches openid-configuration
  (real's initialize issues zero network requests). New per-instance
  memoized resolveEndpoints() runs at every flow entry: authorizeUrl,
  tokenRequest, logoutUrl, AND silentLadder — real's SilentCacheClient
  resolves the authority even for pure cache hits (three silent.*
  snapshots pin a discovery call before a cache-hit result; mini regressed
  on exactly those until silentLadder awaited it). Failed discovery clears
  the memo (next request retries, like real's uncached failure).
- **Every endpoint-resolution failure surfaces as ClientAuthError
  `endpoints_resolution_error`** — real's AuthorityFactory
  .createDiscoveredInstance catches everything from resolveEndpointsAsync
  (untrusted_authority, openid_config_error, invalid metadata JSON…) and
  rethrows that one wrap. The audit finding predicted
  ClientConfigurationError untrusted_authority; the capture proved the
  wrap. Mini keeps the inner untrusted_authority throw for fidelity but
  the observable error is always the wrap.
- **Trust validation (real's updateCloudDiscoveryMetadata order)**:
  cloudDiscoveryMetadata config (host-substring check on the raw JSON
  string — real parses + matches aliases; equivalent in practice, bytes
  matter) → knownAuthorities (host compare, URL or bare-host entries) →
  hardcoded MS_CLOUD_ALIASES (all 13 alias hosts from real's
  InstanceDiscoveryMetadataAliases) → `.ciamlogin.com` suffix (CIAM
  short-circuit) → network AAD instance discovery GET
  `login.microsoftonline.com/common/discovery/instance?api-version=1.1&
  authorization_endpoint=<authority>/oauth2/v2.0/authorize`. Untrusted iff
  the response has no `metadata` and no error other than
  `invalid_instance` (a non-invalid_instance error response is real's
  trusted custom-domain path), or the fetch throws. Pinned by
  authority.known-authorities-validation (Playwright route blocks the
  probe host; real: 1 probe, exact query, zero IdP contact,
  endpoints_resolution_error).
- **Endpoint metadata sources (real's order)**: auth.authorityMetadata
  JSON (skips discovery entirely — authority.authority-metadata-config
  pins zero discovery requests) → hardcoded endpoint templates for the 6
  HARDCODED_ENDPOINT_HOSTS (`https://<host>/<tenant>/oauth2/v2.0/…` —
  authority.hardcoded-cloud-metadata pins logoutRedirect building the
  end_session URL with login.microsoftonline.com fully blocked and zero
  probes) → network openid-configuration with real's
  defaultOpenIdConfigurationEndpoint path rule (/v2.0/ inserted except:
  authority already ends in /v2.0, first path segment "adfs", or OIDC
  protocolMode + non-Microsoft host).
- **Instance-aware multi-cloud fields**: waitForCode now resolves
  `{code, cloudInstanceHostName?, cloudGraphHostName?, msGraphHost?}`
  (ctx seam type change; popup.ts + silentFrame spread it into redeem;
  processRedirect reads the same fragment params). cloud_instance_host_name
  swaps the token-endpoint HOST before redemption (real's
  updateTokenEndpointAuthority; no re-discovery — real serves the swapped
  authority from its host-keyed metadata cache). cloud_graph_host_name/
  msgraph_host are cached on the account entity ONLY when the base account
  is newly created (real's buildAccountToCache; mergeAccount already keeps
  the cached entity's fields) and surfaced on results from the CACHED
  entity — interactive, RT-refresh, and AT-cache-hit paths all read the
  entity now (hardcoded "" removed). Pinned end-to-end by
  authority.instance-aware-cloud-instance (result + entity + silent
  cache-hit + EQP passthrough on the authorize wire).
- **auth.instanceAware config + EQP domain-replace**: implemented in
  authorizeUrl (request account + instance_aware → authority domain
  replaced with account.environment, real's getDiscoveredAuthority). The
  domain-replace leg is UNOBSERVABLE under the mock (account.environment
  == authority host already) — implemented for fidelity, not
  snapshot-pinned; revisit only if a real-world report flags it.
- **logoutUrl is now async** (awaits resolveEndpoints) — ctx seam type
  change, logoutPopup awaits it after the sync about:blank open, so popup
  timing semantics are unchanged.

### C19 — 2026-07-14 — config knobs + logout params (allowRedirectInIframe, tokenRenewalOffsetSeconds, logout hints, serverTelemetryEnabled)

Suite 109→113 (new area 16-config, 4 scenarios, deterministic — capture +
--check both clean — all green on the first mini run after
implementation). Audit findings: `allow-redirect-in-iframe`,
`token-renewal-offset-seconds`, `logout-hint-param`,
`server-telemetry-enabled`.

Decisions/details:
- **navigateToLoginRequestUrl is a handleRedirectPromise OPTION in real
  5.16** (`options?.navigateToLoginRequestUrl ?? true` in
  RedirectClient/StandardController) — the `auth.navigateToLoginRequestUrl`
  config flag is NOT read on the hRP side, so a scenario can't suppress
  the replay via config. The iframe scenario instead makes the iframe src
  EXACTLY equal the redirectUri: after the roundtrip the current URL
  matches the stored login-request URL and real processes the hash in
  place (no replay navigation destroying Playwright's evaluate context).
  Mini's processRedirect keeps reading the config flag (its replay
  semantics were pinned by area 11 where both stacks agree).
- **allowRedirectInIframe gates two places**: the acquireTokenRedirect
  `redirect_in_iframe` guard AND processRedirect's leave-the-hash-for-
  the-opener bail-out (real's RedirectClient replay branch checks
  `!isInIframe() || allowRedirectInIframe`; the in-place branches never
  check). Scenario completes a FULL redirect login inside an iframe.
- **tokenRenewalOffsetSeconds**: real's check is `now + offset >
  expiresOn` — offset 3500 against a fresh 3600s AT is still a cache HIT
  (the audit's suggested test was wrong there); the scenario patches the
  AT to ~600s remaining and uses offset 1800 (forces refresh) + offset 0
  against ~120s remaining (still cache-hit where the old hardcoded 300
  would refresh). Mini: one `config.system?.tokenRenewalOffsetSeconds ??
  300` in the silent ladder.
- **Logout hint params**: new exported `LogoutRequest` (account,
  postLogoutRedirectUri, correlationId, logoutHint, idTokenHint,
  extraQueryParameters); logoutUrl sets id_token_hint, logout_hint
  (explicit request.logoutHint, else account.loginHint, else the
  account's idTokenClaims.login_hint — real's initializeLogoutRequest
  derivation), then eQP appended LAST and never overriding protocol
  params (real's addExtraQueryParameters). Shared by logoutRedirect and
  ./popup's logoutPopup (LogoutPopupRequest extends it). Mock IdP id
  tokens get a login_hint claim via the C16 `/config?claims.*` override.
- **Server telemetry (ported ServerTelemetryManager)**: when
  `system.serverTelemetryEnabled` (default false = real's stub → empty
  strings, no entry): token POSTs carry
  `x-client-current-telemetry: 5|<apiId>,0,,,|<wrapperSku>,<wrapperVer>`
  (cacheOutcome effectively always 0 on the wire — each real interaction
  client gets a fresh manager, and cache hits never reach the network)
  and `x-client-last-telemetry:
  5|<cacheHits>|<apiId,cid pairs>|<errorCodes>|<count>,<overflow>` from
  the `server-telemetry-<clientId>` entry (real's exact entity JSON:
  failedRequests/errors/cacheHits). Failure hooks mirror real's
  cacheFailedRequest sites: RT refresh = 61, silent iframe = **863
  ALWAYS** (real's createSilentIframeClient passes ApiId.ssoSilent even
  for the acquireTokenSilent ladder — snapshot-pinned; mini's silentFrame
  keeps 864 for its perf-event internals but stFail records 863),
  processRedirect = 865, popup = 862 via the new `ctx.stFail` seam.
  Entry cleared after a successful token response (real clears in
  handleServerTokenResponse), with real's 330-byte flush cap + partial
  retention and 50-error FIFO; AT cache hits increment cacheHits
  (SilentFlowClient). initializeWrapperLibrary's SKU/version now feed the
  current-telemetry platform fields.
- Sizes: compat 56.4 KB min (+1.4), mini-core 29.4 (+1.4 — the ST
  manager, hint params, and knobs are all core; stFail rides the existing
  ctx seam for ./popup).

### C20 — 2026-07-14 — perf-event emission semantics (redirect root event, failure cid joins, callback dedupe, preflight/init-once, performance marks)

Suite 113→119 (6 new 08-telemetry scenarios, deterministic — captured
twice, byte-identical — all green on the first mini run after
implementation). Audit findings: `handle-redirect-perf-event`,
`failure-event-correlation-id`, `duplicate-perf-callback-dedupe`,
`preflight-failure-no-perf-event`, `init-perf-event-once`,
`performance-marks-session-flag`.

Decisions/details (several audit claims corrected by capture):
- **Redirect root event** (telemetry.perf-redirect-event): one
  `acquireTokenRedirect` event per processed redirect response, emitted by
  handleRedirectPromise with the CACHED request's correlationId; clean
  loads emit nothing (real returns before startMeasurement when no
  interaction is in progress) and repeat calls reuse the memoized promise.
  Ext is the REDEMPTION HALF only — no PKCE / getAuthCodeUrl /
  getStandardParams / deserializeResponse (those ran on the pre-redirect
  page) — new REDIR group + DISC + NETDISC. Here
  networkClientSendPostRequestAsync carries BOTH CallCount and DurationMs
  (unlike popup's CallCount-only); the DurationMs key is exactly 43 chars
  so it normalizes as `<b64-43>` (C11's collision, now load-bearing).
  Event carries `previousLibraryVersion` because the return page's cache
  already holds `msal.version` from the pre-redirect load: mini reads the
  store key at telemetry-feature setup (pre-initialize) and rides it on
  every event when present. Mini wrapper dedupes by promise identity and
  skips emission for `uninitialized_public_client_application` (real
  blocks before the measurement). accountType/cacheMatchedAccounts/
  refreshTokenSize/kmsi/httpVerToken/requestId like popup's success shape;
  no accessTokenSize/idTokenSize/scenarioId on this event (snapshot-pinned).
- **Failure-event correlation join** (telemetry.perf-failure-correlation):
  KEY CAPTURE FACT — real 5.16's RT token POST body has NO
  client-request-id (RefreshTokenClient.createTokenRequestBody never calls
  addCorrelationId); the cid rides the token-endpoint QUERY string
  (createTokenQueryParameters), which mini already matched. The scenario
  compares the QUERY param. AuthError gains a `correlationId` property;
  the silent deduped-promise rejection handler stamps
  `validRequest.correlationId` (real StandardController:1266) and
  processRedirect's catch stamps the stored request cid; the telemetry
  failure path uses `e.correlationId ?? req.correlationId` so event cid
  === error.correlationId === wire cid without an app-supplied cid.
- **Callback dedupe** (telemetry.perf-callback-dedupe):
  BrowserPerformanceClient.addPerformanceCallback compares
  `cb.toString()` against every registered callback and returns the
  EXISTING id on match (single delivery). KEY CAPTURE FACT: with NO
  telemetry.client configured, real returns `""` from
  addPerformanceCallback (NOT the audit-suggested "callback-id") — mini's
  stub path returns "".
- **Preflight failures** (telemetry.perf-preflight-failures): confirmed
  the audit verdict's correction — `no_account_error` emits NOTHING (real
  throws after startMeasurement with no end/catch attached; mini's silent
  wrapper rethrows without emitting), but uninitialized preflight
  failures DO emit success:false events for both acquireTokenSilent and
  acquireTokenPopup (real's preflightCheck wrapper ends the measurement).
- **Init-once** (telemetry.perf-init-once): repeat initialize() calls
  return before real starts the measurement — mini's wrapper gates
  emission behind a flag set on first success.
- **Performance marks** (telemetry.performance-marks): with
  `sessionStorage["msal.browser.performance.enabled"]="1"` AND an opt-in
  perf client, real writes msal.start/end/measure.<op>.<cid> timeline
  entries; after root end, the surviving measures for a silent cache-hit
  are EXACTLY the C11 ext DurationMs-key set + the root name — validating
  the ext tables. Mini synthesizes mark/mark/measure triplets at emit time
  from its ext table (0-duration; the scenario compares names with cids
  stripped). Flag off → zero entries in both.
- Sizes: compat 58.0 KB min (+1.6), mini-core 29.5 (+0.1 — only the
  AuthError property + two cid stamps are core; everything else is the
  telemetry feature). e2e 25/25.

### C21 — 2026-07-14 — authenticationScheme "pop"/"ssh-cert" (PoP binding, SHR signing)

Suite 119→121 (new area 17-pop: pop.silent-shr, pop.ssh-scheme-and-errors —
deterministic, captured twice byte-identical; both green on the first mini
run after implementation + one extra fix below). Audit finding:
`authentication-scheme-pop`.

Decisions/details:

- **Architecture**: crypto lives in a NEW pay-to-play `./pop` feature
  (RSA-2048 RS256 keygen — real uses RSASSA-PKCS1-v1_5, NOT the audit
  summary's "ECDSA" — kid = b64url(sha256(sorted {e,kty,n})), SHR signing,
  keypairs in real's IndexedDB keystore `msal.db`/`msal.db.keys` with the
  private key re-imported unextractable, in-memory map first). Core carries
  the scheme plumbing: `ctx.pop` seam, req_cnf/token_type body params,
  scheme-aware AT lookup + save-dedupe (pop/ssh ATs never evict bearer ones
  and vice versa), `AccessToken_With_AuthScheme` entity + `keyId` + scheme
  cache-key suffix, result tokenType/SHR, ssh config validation. Scheme/SHR
  request fields ride an `shr` member on AuthCodeResponse (popup/iframe) and
  inside the cached `msal.request` (redirect roundtrip). A "pop" request on
  a client without `./pop` composed throws
  BrowserAuthError `feature_not_configured` (D1 will generalize).
- **Capture facts**: SHR header key order is `typ, alg, kid` (JoseHeader
  class-field order — kid itself is b64url({kid})); payload order
  `at, ts, m, u, nonce, p, q, [client_claims,] cnf` with undefineds
  dropped; cnf.jwk is the FULL exported public JWK, keys sorted
  (alg,e,ext,key_ops,kty,n); m is uppercased; q = [[], "<query>"]. Entity
  keyId comes from the SERVER AT's own cnf.kid claim (missing →
  ClientAuthError token_claims_cnf_required_for_signedjwt); ssh keyId from
  response `key_id`; entity tokenType from response `token_type`. Cache
  hits RE-SIGN a fresh SHR (new nonce/ts, same embedded AT) unless
  request.popKid — popKid also skips keygen (req_cnf = b64url({kid}) only)
  and returns the raw secret. ssh-cert without sshJwk/sshKid throws
  ClientConfigurationError missing_ssh_jwk/missing_ssh_kid. The silent
  dedupe thumbprint gained the scheme/SHR fields (real's includes them).
- **Extra parity fix exposed by the new full-body digest**: real's RT grant
  carries `redirect_uri` ONLY when the request passes one; mini was always
  sending the config default on the refresh leg (existing scenarios only
  digested grant names, so 119/119 never saw it). redeemRefresh now
  overrides the base with undefined (post() drops undefined values).
- **Mock IdP**: pop/ssh extension — echoes the requested token_type; for
  pop mints the AT as an alg-none JWT embedding req_cnf's kid as cnf.kid
  (AAD's contract, required by real's cache write); for ssh echoes the
  req_cnf JWK's kid as `key_id`.
- **Compat**: composes `pop` + exports the AuthenticationScheme constant
  (Bearer/pop/ssh-cert).
- **Size call (borderline vs the task's ~3 KB gate)**: compat 61.2 KB min
  (+3.2 vs C20's 58.0) — +1.4 is core scheme plumbing (incl. the RT
  redirect_uri fix), ~+1.8 the ./pop feature. Judged within the "~3 KB"
  tolerance rather than stopping the loop: the gap was audit-confirmed
  observable, everything is scenario-pinned, and descoping is a one-line
  revert (drop `pop` from the compat compose list) if the bytes matter
  more. mini-core 30.9 (+1.4). e2e 25/25.

### D1 — 2026-07-14 — seam hardening + consumer packaging

- **Seam hardening (core)**: `createClient` installs throwing stubs for every
  feature-owned public API before features run (features overwrite them):
  loginPopup / acquireTokenPopup / logoutPopup → `./popup`,
  acquireTokenByCode → `./broker`, addPerformanceCallback /
  removePerformanceCallback → `./telemetry`. Each throws
  `BrowserAuthError("feature_not_configured")` with
  `errorMessage = '<api> requires composing <feature> from
  "@mini-msal/browser/<feature>"'` (new exported `featureError()` helper;
  the pre-existing pop seam guard now uses it too:
  `authenticationScheme "pop" requires composing pop from
  "@mini-msal/browser/pop"`). This error is mini-specific by design — real
  MSAL has no non-composed state; compat composes everything so conformance
  never observes it.
- **Unit checks**: `npm run seams` (test/unit/seams.mjs) drives a new
  core-only harness bundle (`conformance-mini-core`,
  test/apps/harness/harness-mini-core.ts — exposes `__core` + `__features`
  globals; a test vehicle, NOT a size target although it appears in the
  measure table). 10 checks: all 6 stubs throw the documented error on a
  core-only client; core APIs intact; composing popup/broker/telemetry
  replaces exactly its own stubs while OTHER features' stubs keep guarding
  (partial compositions).
- **Consumer packaging**: packages previously exported raw `.ts` — unusable
  outside this repo. Each package now has a tsc dist build
  (`tsc -p packages/<p>` — typescript@7 devDep; JS + .d.ts, ES2022, module
  ESNext) and a three-condition exports map per entry:
  `"types"` → dist/*.d.ts, `"mini-msal-src"` → src/*.ts, `"default"` →
  dist/*.js. Internal rspack builds resolve `mini-msal-src` first
  (resolve.conditionNames in test/infra/rspack.config.mjs), so ALL measured
  bundles + harnesses still compile straight from src exactly as before —
  consumers get dist. `files: ["dist", "src"]` added for npm pack.
  Five real type errors surfaced and fixed (types-only, zero runtime bytes):
  3× `Uint8Array<ArrayBuffer>` BufferSource generics in local-storage.ts,
  2× spread-args casts in telemetry.ts, 1× TokenRequest cast in react.
- **pack:check** (test/packaging/check.mjs, network for react+types
  install): tsc-builds all three packages, `npm pack`s them, installs the
  tarballs into a throwaway consumer (os.tmpdir, file: deps), then verifies
  (1) STRICT-tsconfig `tsc --noEmit` over consumer sources exercising the
  root + every subpath export, and (2) rspack (default conditions → dist)
  tree-shakes: core-only 29.9 KB min / core+popup 32.5 / compat 59.5 /
  react-bindings 65.1 (react external) — matching the src-built matrix
  (30.9/61.2) within default-minifier tolerance; gates assert the ranges +
  ordering. compat's `"@mini-msal/browser": "*"` dependency is satisfied by
  the sibling tarball install (npm does not hit the registry for it).
- **Size cost of the stubs**: mini-core 31.3 KB min (+0.4), compat 61.7
  (+0.5), stack 71.7 (+0.5) — the guard surface is core bytes by design
  (that's the point: the errors exist BEFORE any feature is composed).
- **SOP change**: `npm run seams` added to the validation gates in
  PARITY_STATE (cheap, catches seam regressions); `npm run pack:check` on
  demand for packaging-affecting tasks (needs network).

### D2 — 2026-07-14 — consumer docs (UPGRADING, ALACARTE, README rewrites)

- **New standalone guides**: `docs/UPGRADING.md` (drop-in migration from
  `@azure/msal-browser`/`@azure/msal-react`: import swap, bridge-page swap,
  cache carry-over/KMSI/migration, post-switch verification checklist,
  non-goals) and `docs/ALACARTE.md` (step-down guide compat 61.7 →
  explicit-composition 56.3 → core+popup 32.6 → core 29.7 KB min, feature
  catalog with measured per-feature deltas, bridge-page rule, methodology).
  `README.md` + `docs/README.md` rewritten as consumer guides: quick-start
  per profile (compat / core-redirect / core+popup / react), feature-catalog
  table, bridge requirement — and stale 75/75-era numbers (compat 32.5,
  core 19.5) updated to the current 121/121 matrix.
- **Per-feature size deltas measured** (not previously known): one minimal
  app bundled with a single feature composed at a time, exact
  test/infra/rspack.config.mjs minify settings (swc passes:3, es2022,
  mini-msal-src condition). Base 29.7 KB min; deltas: popup +2.9, pop +2.0,
  local-storage +2.6, cache-migration +3.0, broker +6.9, telemetry +9.4,
  naa +8.4 (standalone `createNestableClient` entry incl. bridge handshake);
  full 6-feature explicit composition 56.3 (compat's constant/error-code
  namespaces account for most of the remaining 5.4 to 61.7); react +10.0
  (stack 71.7 − compat 61.7, React external). Docs note deltas are
  approximately additive. Measurement script was session-scratch (not
  committed) — D3's examples will make per-profile sizes a permanent
  `npm run measure` fixture.
- **Copy-paste-runnable enforced mechanically**: all 14 `ts`/`tsx` fenced
  blocks across the four docs extracted and type-checked (strict,
  moduleResolution bundler, jsx react-jsx) against the packages' built
  `dist` types — clean. Two real sample bugs this caught: `MsalProvider`'s
  `instance` prop is `IPublicClientApplication = AuthClient & PopupClient`,
  so an à-la-carte client needs `[popup]` + a cast (docs now say so), and an
  undeclared `config` in the full-composition sample.
- **Doc-content decisions**: à-la-carte typing pattern documented as
  `createClient(...) as AuthClient & PopupClient` (pack:check's
  `as unknown as PopupClient` works but reads worse); bridge page documented
  as required for popup/ssoSilent/silent-iframe and NOT for pure-redirect
  apps; prototype/not-on-npm caveat stated in both guides; B2C/ADFS stay
  the only authority non-goals (CIAM is implemented since C18, PoP since
  C21 — the old README wrongly listed PoP as unimplemented).

### D3 — 2026-07-14 — runnable examples with measured sizes

- **NEW `examples/`** — one minimal, self-contained app per consumption
  profile, each a `main.ts(x)` + placeholder `authConfig.ts` + README
  (what it composes, measured size, bridge rule, how to run):
  core-redirect (`createClient` only, 30.3 KB min / 10.6 gz), core-popup
  (core + `./popup`, popup.html bridge, 33.2 / 11.5), compat
  (`PublicClientApplication` drop-in, 61.1 / 19.9), react (compat under
  `MsalProvider` + templates + hooks, 66.4 / 21.8 — React external, same
  methodology as the matrix). `examples/README.md` indexes them.
- **Permanent size fixture**: the examples build as rspack variants
  `example-*` with the exact size-matrix settings, so `npm run measure`
  now reports per-profile consumer sizes permanently (D2 had noted these
  numbers were session-scratch).
- **Smoke check**: NEW `npm run examples:smoke`
  (`test/examples/smoke.mjs`, e2e-harness pattern: spawns mock IdP +
  static server, headless system Chrome). Each example also builds an
  `example-*-smoke` variant whose `./authConfig.js` import is swapped for
  `test/apps/authConfig.mock.ts` via rspack
  `NormalModuleReplacementPlugin` (React bundled so the react smoke runs
  standalone); postbuild writes their index.html + bridge popup.html
  pages. The check drives the real UI — click `#signin`, complete the
  redirect or popup roundtrip — and asserts the greeting plus a
  silently-acquired token render. **8/8 first run.**
- **Decisions**: measured builds bake the placeholder AAD config (that's
  what consumers copy); smoke variants exist because config is baked at
  bundle time. Popup-using examples pass an explicit `popup.html`
  redirectUri (the documented bridge pattern); core-redirect ships no
  bridge. The react example composes compat under the provider (the
  documented default; the à-la-carte cast is ALACARTE's job to teach).
- Zero package-code changes — sizes unchanged (compat 61.7, core 31.3,
  stack 71.7). Docs link the examples (README table, docs/README
  consuming/map/status sections).

### D4 — 2026-07-14 — requirements audit vs Mission (skeptic pass)

Every validation gate re-run fresh this session (not trusted from docs):
build clean, conformance:mini **121/121**, e2e **25/25**, seams **10/10**,
examples:smoke **8/8**, pack:check OK (types + tree-shake gates 29.9 /
32.5 / 59.5 / 65.1), report regenerated (0 behavioral-diff / 0
missing-feature / 0 bug), measure matches every recorded number (compat
61.7, core 31.3, stack 71.7, examples 30.3/33.2/61.1/66.4).

Verdict per Mission bullet:

1. **100% compat — GAPS FOUND (filed D5).** The suite proves behavioral
   parity, but an exhaustive module-export diff (real ESM keys vs
   compat dist) found **18 exports real has and compat lacks**: ApiId,
   AuthenticationHeaderParser, AzureCloudInstance,
   BrowserPerformanceMeasurement, BrowserRootPerformanceEvents,
   BrowserUtils, DEFAULT_IFRAME_TIMEOUT_MS, EventHandler,
   EventMessageUtils, JsonWebTokenTypes, LocalStorage, MemoryStorage,
   SessionStorage, ResponseMode, SignedHttpRequest,
   StubPerformanceClient, enforceResourceParameter,
   stubbedPublicClientApplication. Compat extras (NativeAuthError,
   NestedAppAuthError, createAuth) are harmless. Also: real's PCA
   **constructs fine in plain Node** (SSR/Next.js drop-in — verified:
   `new PublicClientApplication({auth:{clientId}})` succeeds in node,
   ops fail later); mini's compat throws raw
   `ReferenceError: location is not defined` from createClient at
   construction time. React surface is complete (msal-react keys ⊆ mini
   react; extra ReactAuthError harmless). PCA instance methods: only
   runtime-extras `waitForIframeResponse`/`waitForPopupResponse` differ —
   they are NOT in real's .d.ts (internal, untyped), descoped unless
   trivial during D5.
2. **Pay-to-play — PASS.** pack:check tree-shake gates prove uncomposed
   features cost zero bytes (core 29.9 vs compat 59.5); every subpath
   export strict-type-checks from a real npm-pack consumer; seams 10/10
   cover partial compositions.
3. **Effortless consumption — PASS.** feature_not_configured stubs name
   the exact import (seams-verified); ALACARTE step-down + per-feature
   measured deltas; consumer POV proven by pack:check's throwaway
   consumer.
4. **Docs & examples — GAP FOUND (filed D6).** Deliverables exist, sizes
   verified fresh against `npm run measure` (all match). But D2's "14 doc
   samples type-checked" was a **session-scratch one-off** — nothing in
   the repo re-checks doc samples (confirmed: D2 commit a542eb133 is
   docs-only; no docs:check script exists), so samples can rot silently.

Filed: **D5** (compat export-surface completion + SSR-safe construction,
test-first via an exhaustive exported-surface-full scenario + node
construction check), **D6** (committed `npm run docs:check`), **D7**
(re-audit close-out — final summary gate). Audit is evidence-only: zero
package-code changes, sizes unchanged.

### D5 — 2026-07-14 — compat export-surface completion + SSR-safe construction

Test-first, both gates red before impl:

- **init.exported-surface-full** (suite 121→122): pins the FULL sorted
  `Object.keys(lib)` + per-key typeof (real snapshot captured twice,
  deterministic), plus behavior probes for every long-tail export
  (BrowserUtils fns, storage classes, SHR sign/verify shape,
  stubbed-PCA rejections, EventHandler dispatch, EMU status mapping,
  perf helpers, enforceResourceParameter, header parser). Mini-only
  extras `NativeAuthError`/`NestedAppAuthError`/`createAuth` are
  filtered by an explicit allowlist IN the scenario — any new
  undocumented export will fail the pin. Mini went from a module-load
  TypeError to green first run after implementation.
- **seams 10→12**: node-side SSR checks import the BUILT dist into the
  seams process itself (no browser globals): `new
  PublicClientApplication({auth:{clientId}})` and core `createClient()`
  must construct. Pre-fix mini threw `ReferenceError: location is not
  defined`.

Implementation decisions:

- **SSR fix is construction-only, like the task spec**: createClient's
  eager `new URL(config.auth.redirectUri ?? "/", location.href)` became
  a lazy closure invoked at the three flow-time call sites. Verified
  against real in plain Node: real's PCA constructs AND initialize()s
  and getAllAccounts() returns [] there; mini guarantees construction
  only (ops touch sessionStorage/location and throw). Recorded as an
  accepted divergence — Next.js/SSR drop-ins construct at module scope
  but only operate in the browser.
- **surface.ts placement**: all 18 exports live in NEW
  `packages/compat/src/surface.ts` (compat is where
  BrowserConfigurationAuthError already lived — the class moved there,
  index re-exports `*`). Nothing was added to core, and the surface
  tree-shakes away for consumers who don't import it: mini-core 31.3
  unchanged, pack:check gates core 29.9 / popup 32.5 (unchanged) /
  compat 59.9 (+0.4) / react 65.4 (+0.3).
- **Reuse over duplication**: SignedHttpRequest is built on ./pop's
  machinery — pop.ts now exports `makeBoundKeyPair`/`signPop`/`keystore`
  (feature closure rewritten on top of them; `signPop` gained real's
  claims-override param, spread before cnf so claims can never override
  the jwk). LocalStorage (the exported class) is built on
  ./local-storage's newly exported `loadEncryptionCookie`/
  `encryptEntry`/`decryptEntry` — same msal.cache.encryption cookie,
  HKDF/AES-GCM scheme and broadcast channel, so the class interoperates
  with the feature's at-rest entries.
- **Faithful-lean ports**: constants byte-exact (ApiId, ResponseMode,
  JsonWebTokenTypes, AzureCloudInstance, DEFAULT_IFRAME_TIMEOUT_MS,
  BrowserRootPerformanceEvents); stubbedPublicClientApplication matches
  real's 26 keys and per-method return/reject behavior; EventHandler /
  EventMessageUtils / AuthenticationHeaderParser /
  BrowserPerformanceMeasurement / StubPerformanceClient /
  enforceResourceParameter are direct ports; BrowserUtils has real's 22
  keys with lean internals (createGuid is UUIDv4 not real's v7 —
  scenario pins format only; blockReloadInHiddenIframes uses core's
  code|error hash test rather than real's full deserializer).
- **DESCOPED**: PCA instance runtime-extras
  `waitForIframeResponse`/`waitForPopupResponse` — absent from real's
  .d.ts (untyped internals), left out. `BrowserUtils.waitForBridgeResponse`
  (which IS exported by real) was implemented.
- Type-only: core's `EventMessage` gained optional `correlationId`
  (real's message shape; core still doesn't set it — event-stream
  snapshots unchanged).
- Consumer-doc size/count refresh: 121→122, seams 12/12, compat
  61.7→62.1, stack 71.7→72.0, examples 30.4/33.3/61.5/66.8 (README,
  docs/README, UPGRADING, ALACARTE, example READMEs).

All gates: conformance:mini **122/122**, e2e **25/25**, seams **12/12**,
examples:smoke **8/8**, pack:check OK, GAP_REPORT 122-scenario all-pass.

### D6 — 2026-07-14 — committed doc-sample check (npm run docs:check)

- **NEW `npm run docs:check`** (`test/docs/check.mjs`): re-derives D2's
  session-scratch sample check as a committed gate. tsc-builds the three
  packages' dist (browser first), extracts every fenced ```ts / ```tsx
  block from `README.md`, `docs/README.md`, `docs/UPGRADING.md`,
  `docs/ALACARTE.md` into `test/docs/.samples/` (one file per sample,
  named `<doc>-L<line>`, `export {}` appended so samples are modules and
  can't collide), then strict-tscs the lot (moduleResolution bundler,
  jsx react-jsx, DOM libs, noEmit). No network, no npm pack: the scratch
  dir lives inside the repo, so `@mini-msal/*` resolves through the
  workspace symlinks whose exports `"types"` condition points at the
  dist d.ts — the same types a consumer gets.
- **14 samples, 0 opt-outs.** Opt-out marker for intentionally-partial
  samples: `<!-- docs-check:skip <reason> -->` on the nearest non-empty
  line above the fence (logged as skipped). Zero samples extracted →
  hard fail (extraction-regression guard). On tsc failure the scratch
  dir is kept for debugging (gitignored: `test/docs/.samples/`).
- **Self-test run as specced**: temporarily reintroduced D2's two caught
  sample bugs — dropped the `as AuthClient & PopupClient` cast on the
  ALACARTE react sample (→ TS2322 `AuthClient` not assignable to
  `IPublicClientApplication`) and the `const config` declaration in the
  full-composition sample (→ TS2304 cannot find name `config`). The check
  failed on exactly both, then passed clean after revert.
- SOP validate list gains `npm run docs:check` (docs-affecting tasks
  only — it's ~15s, dominated by the three tsc dist builds).
- Session gotcha (already covered by the stale-servers note): `npm run
  seams` chained immediately after e2e crashed with a TimeoutError
  (port contention) that a `| tail` pipe masked to exit 0 — standalone
  rerun after killing 4173/4599 passed 12/12. Don't pipe gate commands
  through tail when the exit code matters.
- Zero package-code changes — all sizes unchanged.

All gates: conformance:mini **122/122**, e2e **25/25**, seams **12/12**,
docs:check **14 samples OK**, measure matches D5's matrix.

### D6b — 2026-07-14 — size-reduction reference doc (docs/SIZE.md)

- **NEW `docs/SIZE.md`** (user-requested 2026-07-14): plain-language
  reference on HOW mini-msal got small, for readers who never opened this
  repo. Headline table (per-profile real-vs-mini with the honest ratios:
  drop-in 3.5×, popup SPA 6.8×, core 7.4×, bridge 10.8× — real's key
  property being that it costs 220.5 KB REGARDLESS of profile), a
  where-real's-bytes-go attribution, then 8 technique sections, each with
  a concrete example: (1) pay-to-play composition vs StandardController's
  static everything-imports (real excerpt) + per-feature price list;
  (2) closures vs class layers — StandardOperatingContext boilerplate
  excerpt + the minifier argument (property names can't mangle, locals
  can); (3) one browser-only package vs the ~80 KB msal-common layer;
  (4) one error family + AKA() default message vs per-subsystem
  class/codes/factory modules; (5) telemetry event shapes as string
  tables vs method-wrapping instrumentation; (6) the 0.6 KB bridge
  (full source shown) vs real's 6.5 KB; (7) small habits (WIRE_ID
  single-source constants, zero TS enums — verified by grep, es2022,
  bytes-over-abstraction rule); (8) sideEffects:false + per-feature
  subpath exports + pack:check size gates. Closes with what was NOT
  traded away (parity gates) + non-goals link.
- **Every claim verified against source before writing**: attribution
  numbers re-measured via `npm run analyze -- msal-stack` (msal-common
  rows sum ≈79.1 KB → "~80 KB"; interaction_client 44.9; cache
  31.5+18.2; controllers 25.6); real excerpts copied from dist (the
  StandardOperatingContext.initialize body corrected against the real
  file rather than paraphrased); enum grep = 0 hits; AKA string matches
  real's getDefaultErrorMessage; UPGRADING anchor fixed to
  #known-non-goals. Softened two overclaims during self-review: bridge
  described as "same {v:1,payload} contract" (interchangeability isn't
  e2e-proven), msal-common described by its Node+browser generality
  (not specific back-end counts).
- **Doc-sample policy decision**: real-msal internals are shown in
  ```js fences (docs:check only extracts ts/tsx — they're not consumer
  code); mini-side samples are ```ts and MUST type-check — the
  composition, feature-shape (ClientContext), error-family, and bridge
  samples all pass strict tsc against dist types. ONE justified skip:
  the telemetry table excerpt is abridged with "…"
  (`<!-- docs-check:skip -->`). check.mjs DOCS list += docs/SIZE.md →
  **18 samples + 1 skip**.
- Linked from README.md (doc table), docs/README.md (guides list),
  ALACARTE.md (closing pointer).
- **bundle-size-experiment.md Results refreshed** (task's stale-header
  clause): 75/75-era table (32.5/39.5/19.5) replaced with the
  2026-07-14 122/122 matrix (62.1/72.0/31.3, same real numbers), the
  old milestone kept as history prose; library-only accounting updated
  ~27.6→~57 KB, ~8×→~4.3×; pointers to ALACARTE/SIZE added.
- Zero package-code changes — measure re-run: all sizes unchanged
  (72.0 stack / 62.1 compat / 31.3 core; examples 30.4/33.3/61.5/66.8
  match the docs).

All gates: conformance:mini **122/122**, e2e **25/25**, seams **12/12**,
docs:check **18 samples OK (1 skip)**, measure matches D5's matrix.

## Progress log

| Task | Status | Pass | mini-stack size | Notes |
|---|---|---|---|---|
| baseline | — | 2/75 | 20.1 KB min / 6.9 gz | e2e 25/25, check 75/75 |
| A0 | done 2026-07-10 | 2/75 | 18.7 KB min / 6.9 gz | pure refactor, statuses+diff counts identical to baseline; e2e 25/25; NEW mini-core variant 10.9 KB min / 4.2 gz |
| A1 | done 2026-07-10 | 8/75 | 19.2 KB min / 7.1 gz | +6 pass (popup-closed-by-user, iframe-timeout, interaction-required-variants, redirect-in-iframe, silent-unknown-account, 5xx-server-error); e2e 25/25; mini-core 11.4 min / 4.4 gz |
| A2 | done 2026-07-10 | 9/75 | 19.3 KB min / 7.2 gz | +1 pass (errors.uninitialized-client): `initialized` flag, guard first in preflight(), handleRedirectPromise rejects pre-init (own check — full preflight would break iframe null path); e2e 25/25; mini-core 11.5 min / 4.5 gz |
| A3 | done 2026-07-10 | 10/75 | 19.4 KB min / 7.2 gz | +1 pass (accounts.get-account-filters): getAllAccounts(filter?) honors filter; getAccount({}) / all-empty filter → null (matches real CacheManager.getAccountInfoFilteredBy); shared matchesFilter predicate; e2e 25/25; mini-core 11.5 min / 4.5 gz |
| A4 | done 2026-07-10 | 10/75 | 19.4 KB min / 7.2 gz | +0 pass by design: all 8 scope diffs on params.scopes-normalization gone (18→10 diffs), remainder is B1 result-shape. normScopes = real addScopes ([...req, ...defaults] → Set → join, exact-case dedupe) on authorize+token; result scopes keep granted casing (dropped toLowerCase). No regressions (same 10 pass, 591 total diffs); e2e 25/25; mini-core 11.5 min / 4.5 gz |
| A5 | done 2026-07-10 | 11/75 | 19.5 KB min / 7.3 gz | +1 pass (silent.policy-access-token-expired). Policy ladder gates AT/RT/iframe rungs; forceRefresh skips AT; AT-only miss → token_refresh_required ClientAuthError. Discovery cache moved to per-instance memory (msal.meta.* key dropped — B4 item front-ran): networkCalls now match on ALL silent.* scenarios; remaining policy-scenario diffs are pure B1 result-shape (authority/correlationId/tokenType/fromPlatformBroker), force-refresh remainder is B2 events. Total diffs 591→568; e2e 25/25; mini-core 11.7 min / 4.6 gz |
| A6 | done 2026-07-10 | 12/75 | 19.7 KB min / 7.3 gz | +1 pass (silent.concurrent-dedupe): inFlight Map keyed like real's thumbprint (scopes/homeAccountId/authority/claims, no policy/forceRefresh), entry deleted on settle; silent ladder hoisted to silentLadder() closure. Total diffs 568→566; e2e 25/25; mini-core 11.8 min / 4.7 gz |
| A7 | done 2026-07-10 | 13/75 | 20.0 KB min / 7.4 gz | +1 pass (errors.interaction-in-progress): core lock()/unlock() on `msal.interaction.status` (real's entry shape `{clientId,type}`, type signin/signout); popup + redirect + logout take the lock, popup/logoutPopup release in finally, processRedirect releases on return (covers redirect + signout return legs); popup window names now `msal.<uuid>` (real names are per-request unique via correlationId). Total diffs 566→561; e2e 25/25; mini-core 12.1 min / 4.8 gz |
| A8 | done 2026-07-10 | 13/75 | 19.9 KB min / 7.4 gz | +0 pass by design (scenarios also need B1/B2): removed the auto-write of `active-account-filters` on login — active account ONLY via setActiveAccount, logout still clears it when it matches. All active-account diffs gone (activeAfterSecondLogin, logout.ok.active, storage active-account-filters keys). e2e 25/25 unchanged — demo app already calls setActiveAccount explicitly. Total diffs 561→554; mini-core 12.0 min / 4.7 gz |
| B1 | done 2026-07-10 | 23/75 | 20.4 KB min / 7.5 gz | +10 pass (silent.cache-hit-fresh, expiry-window-refresh, policy-access-token-valid/at-and-rt/refresh-token/rt-and-network/skip, params.per-request-redirect-uri, params.scopes-normalization, resilience.network-drop). Result gains authority (`<authority>/`), correlationId (request's or per-request uuid; threaded redirect via msal.request, popup/ssoSilent via AuthCodeResponse, RT via redeemRefresh arg), tokenType "Bearer", fromPlatformBroker false, state "" on interactive+ssoSilent ONLY (absent on acquireTokenSilent — snapshots show real silent results carry NO state; silentLadder deletes it from the iframe rung). AccountInfo gains environment (cache entity's). telemetry.correlation-id-propagation resultMatches now true (rest of that scenario is C3). Total diffs 554→448; e2e 25/25; mini-core 12.5 min / 4.9 gz |
| B2 | done 2026-07-10 | 32/75 | 22.9 KB min / 8.2 gz | +9 pass (handle-redirect-clean-load, acquire-token-popup, redirect-state-tampered, force-refresh, multi-account, active-account-persistence, cancelled-login-redirect, popup-blocked, double-initialize). Full real event streams + full result/account key shapes (payloadKeys = Object.keys). Total diffs 448→265; remaining event diffs only in cross-tab (C6), perf (C4), broker (C7). e2e 25/25; mini-core 14.0 min / 5.2 gz |
| B3 | done 2026-07-10 | 40/75 | 24.0 KB min / 8.6 gz | +8 pass (login-redirect-roundtrip, login-popup-roundtrip, sso-silent-cold/warm, acquire-token-redirect, refresh-token-grant, iframe-fallback-no-rt, telemetry.token-request-headers). Authorize+token requests now carry real's full param set (nonce+validation, client-request-id, client_info, default claims, clidata, X-AnchorMailbox ccs, x-client-SKU/VER, empty telemetry params, lib-capability, charset content-type, RT redirect_uri). GAP_REPORT: 0 bugs remaining. e2e 25/25; mini-core 15.0 min / 5.6 gz |
| B4 | done 2026-07-10 | 44/75 | 24.4 KB min / 8.7 gz | +4 pass (core.storage-shape-after-login, init.storage-before-login, accounts.logout-redirect, accounts.logout-popup-per-account; redirect-state-tampered was already done by B2). msal.version on initialize; lastUpdatedAt on all entities + cachedByApiId (real ApiId per flow) on the account; end_session gains client-request-id + state (lib-state format), per-request postLogoutRedirectUri honored on both logout flows. Mock-IdP determinism fix: nonce now keyed per PKCE challenge (see Decision Log — was a global that a stray late authorize could clobber → phantom nonce_mismatch flake). Total diffs 145→143 (accounts.local-storage +1, a C6 item). e2e 25/25; mini-core 15.4 min / 5.8 gz |
| C1 | done 2026-07-10 | 50/75 | 25.4 KB min / 9.0 gz | +6 pass (params.login-hint-domain-hint, sid-passthrough, extra-query-parameters, claims-and-cae, custom-state, authority-override-per-request — all 9 params.* now green). sid gated on prompt=none, select_account suppresses hints+ccs; eQP on authorize + token QUERY (real ignores tokenQueryParameters — see Decision Log); mergedClaims w/ clientCapabilities xms_cc; custom state `b64(libState)\|custom` echoed on result; per-request authority w/o re-discovery. Total diffs 143→132; remaining diffs all C2–C9 areas. e2e 25/25; mini-core 16.3 min / 6.1 gz |
| C2 | done 2026-07-10 | 52/75 | 26.2 KB min / 9.3 gz | +2 pass (resilience.throttle-429-retry-after, proactive-refresh — all 4 resilience.* green). Throttle cache in post() keyed by real-shaped thumbprint; blocked retry re-throws stored ServerError w/ errorCode "" + raw error_description, zero network; refresh_in → refreshOn on AT entity + result. Total diffs 132→127; remaining 23 diff scenarios all C3–C9 areas. e2e 25/25; mini-core 17.2 min / 6.3 gz |
| C3 | done 2026-07-10 | 53/75 | 26.4 KB min / 9.3 gz | +1 pass (telemetry.last-telemetry-after-failure). Iframe fallback eligibility = real's checkIfRefreshTokenErrorCanBeResolvedSilently: invalid_grant/token_refresh_required non-IRAE (or IRAE w/ subError bad_token) + no_tokens_found/refresh_token_expired — RT invalid_grant now recovers via iframe (tokenCalls 2); plain IRAE no longer falls back. Total diffs 127→124; remaining telemetry diffs are C4 perf events. e2e 25/25; mini-core 17.4 min / 6.4 gz |
| C4 | done 2026-07-10 | 56/75 | 27.1 KB min / 9.6 gz | +3 pass (telemetry.perf-events-popup-login, perf-events-silent, correlation-id-propagation — ALL telemetry.* green). New `@mini-msal/browser/telemetry` feature: BrowserPerformanceClient (opt-in via config.telemetry.client) + method-wrapping telemetry(ctx) composed last in compat. Total diffs 124→118; remaining 19 scenarios all C5–C9 (init/localStorage/broker/naa). GAP_REPORT: 0 bugs, 1 behavioral-diff (init.popup-without-bridge, a C5 item). e2e 25/25; mini-core 17.4 min / 6.4 gz (core unchanged — types only) |
| C5 | done 2026-07-10 | 62/75 | 29.3 KB min / 10.4 gz | +6 pass (init.get-configuration, logger-callback, exported-surface, popup-without-bridge + free: broker.is-platform-broker-available, naa.no-bridge-fallback — ALL init.* green). Popup/iframe completion migrated to real's redirect-bridge (BroadcastChannel, see Decision Log) — mini bridge page bundle 0.6 KB min (real's: 6.5 KB); logger in core; getConfiguration; full export surface incl. 52 BrowserAuthErrorCodes. GAP_REPORT: 0 bugs, 0 behavioral-diffs, 13 missing-feature (C6–C9 only). NOTE: naa.get-token-popup/error-mapping now ERR at 60s each (+2 min per full run until C9 — see Decision Log). e2e 25/25; mini-core 18.2 min / 6.7 gz |
| C6 | done 2026-07-10 | 64/75 | 32.1 KB min / 11.3 gz | +2 pass (accounts.local-storage, cross-tab-events — ALL accounts.* green). New `@mini-msal/browser/local-storage` feature (encrypted entities via cookie-keyed HKDF/AES-GCM, plaintext indexes, memory mirror, msal.broadcast.cache sync); core `Store` seam + msal.broadcast.event bus (post always, subscribe in LS mode). GAP_REPORT: 0 bugs, 0 behavioral-diffs, 11 missing-feature (C7–C9 only). e2e 25/25 (sessionStorage interop untouched); mini-core 18.6 min / 6.9 gz |
| C7 | done 2026-07-10 | 68/75 | 37.0 KB min / 12.6 gz | +4 pass (broker.dom-probe-and-interactive, dom-first-login, dom-error-mapping, dom-disabled-fallback — all 4 first try). New `@mini-msal/browser/broker` feature (DOM transport): initialize probe, acquireTokenByCode({nativeAccountId}), brokered silent via new ctx.nativeSilent seam, real's error mapping + fatal-DISABLED provider drop. Core +ctx.writeAccount, toAccountInfo surfaces nativeAccountId (mini-core 18.8 min / 7.0 gz, +0.2). GAP_REPORT: 68 pass, 7 missing-feature (C8 extension ×2, C9 naa ×5). e2e 25/25 |
| C8 | done 2026-07-12 | 70/75 | 38.8 KB min / 13.3 gz | +2 pass (broker.extension-handshake-capture, extension-fake-e2e — both first try; ALL broker.* green). Extension transport in ./broker behind a Provider seam: Handshake via window.postMessage + MessageChannel (bounce-back detection + nativeBrokerHandshakeTimeout), GetToken sends the full initRequest verbatim over the port, snake_case result mapped to the shared handleResponse, per-transport x-client-xtra-sku (chrome\|<version> from HandshakeResponse). GAP_REPORT: 70 pass, 5 missing-feature (C9 naa only). e2e 25/25; mini-core 18.8 min / 7.0 gz (unchanged — feature-module bytes only) |
| C9 | done 2026-07-12 | 75/75 | 39.5 KB min / 13.4 gz | +5 pass (ALL naa.* green, first try — 75/75 TOTAL, GAP_REPORT all-pass). New `@mini-msal/browser/naa`: createNestableClient(config, fallback) — GetInitContext handshake, GetTokenPopup/GetToken over the bridge, browser-cache-first silent via new core seams findToken/writeTokens, real's error-status mapping, unsupported APIs sync-throw unsupported_method. Compat's createNestablePublicClientApplication now = createNestableClient(config, createAuth). e2e 25/25; mini-core 19.5 min / 7.1 gz (+0.7 seams) |
| C10 | done 2026-07-12 | 75/75 | 39.5 KB min / 13.4 gz | FINAL SWEEP all green: conformance mini 75/75, check 75/75, e2e 25/25, GAP_REPORT all-pass. New `mini-compat` variant (compat, no React) 32.5 min / 11.0 gz / 9.9 br vs msal-browser-core 220.5/55.4/46.4 (6.8×); stack 39.5/13.4/12.0 vs 248.8/64.9/53.9 (6.3×); mini-core 19.5/7.1/6.3 (verified: still exercises sign-out + multi-account). READMEs + bundle-size-experiment.md updated to final matrix. C-series COMPLETE — D-series next |
| C11 | done 2026-07-13 | 77/77 | 46.8 KB min / 15.5 gz | +2 scenarios (suite 75→77: telemetry.perf-event-shape, -shape-silent — both green first mini run after 2 table fixes). Full per-flow perf-event shapes (6 flows), ext sub-measurement key sets, live sizes/counters; core seam err.silentRefreshReason; harness __cap.perfRaw. compat 39.7 min (+7.2 — shape tables), mini-core UNCHANGED 19.5 (pay-to-play holds). e2e 25/25 |
| C12 | done 2026-07-13 | 78/78 | 50.0 KB min / 16.6 gz | +1 scenario (suite 77→78: init.exported-surface-2 — mini green first run). Core: Logger class (real clone/gate/format), LogLevel + reverse mappings, WrapperSKU, getLogger/setLogger/initializeWrapperLibrary (wrapper meta stored; headers stay stub-empty until C19). Compat: 5 *ErrorCodes namespaces (exact keys+values) + BrowserConfigurationAuthError via shared pack() w/ key=code irregulars. compat 42.9 min (+3.2), mini-core 20.9 (+1.4 core Logger). e2e 25/25 |
| C13 | done 2026-07-13 | 83/83 | 51.6 KB min / 17.0 gz | +5 scenarios (suite 78→83, new area 11-navigation — ALL green first mini run). NavigationClient class + setNavigationClient + system.navigationClient; auth.onRedirectNavigate cancel hook (acquire keeps lock, logout releases + logoutEnd); navigateToLoginRequestUrl deep-link replay via real's msal.{cid}.request.origin / urlHash temp keys, redirectStartPage, in-place #hash restore, doc-title swap; redirect logoutStart payload = raw request (real). e2e cancelled-login seed gained request.origin. compat 44.6 min (+1.7), mini-core 22.5 (+1.6 — redirect is core). e2e 25/25 |
| C14 | done 2026-07-13 | 87/87 | 53.0 KB min / 17.5 gz | +4 scenarios (suite 83→87, all green first mini run). system.navigatePopups default-true sync about:blank open in the user-gesture stack + location.assign navigate; openSizedPopup features/geometry + popupWindowAttributes/popupWindowParent; real popup name formats (token + logout); blocked sync open fails late as popup_window_error like real; logoutPopup mainWindowRedirectUri via new core ctx.navigate seam (ApiId 962, lock survives navigation); telemetry isAsyncPopup wired. compat 45.9 min (+1.3), mini-core 22.6 (+0.1). e2e 25/25 |
| C15 | done 2026-07-13 | 93/93 | 55.8 KB min / 18.4 gz | +6 scenarios (suite 87→93, NEW area 12-react on NEW dual react harness pages conformance-react-{real,mini}; all green first mini run). packages/react rewritten as a port of msal-react 5.5.1: provider initialize() + initializeWrapperLibrary + full InteractionStatus reducer (real's event mapping w/ clear-guards; accounts [] during startup); useMsalAuthentication {login, acquireToken, result, error} w/ auto-acquire + IRAE fallback + logout reset; useIsAuthenticated(ids)/useAccount case-insensitive + active-account fallback; templates w/ identifier props + function children; MsalAuthenticationTemplate spread/throw error contract. Core +InteractionStatus export. compat UNCHANGED 45.9, mini-core UNCHANGED 22.6. e2e 25/25 |
| C16 | done 2026-07-13 | 98/98 | 60.6 KB min / 19.9 gz | +5 scenarios (suite 93→98, NEW area 13-cache — all green first mini run). AT intersecting-scope dedupe on save + multi-match clear on lookup (+realm filter on the silent AT rung); mergeAccount = buildAccountToCache (one base entity per homeAccountId+env, profiles appended, isHomeTenant computed, shared by web/broker/naa); getAllAccounts expands tenantProfiles (Map) into per-tenant AccountInfos; Store.setUser kmsi seam → plaintext KMSI entities in localStorage mode + real AccountInfo.kmsi; NEW /cache-migration feature (msal.0/1/2→3, 5-day TTL) via new ctx.onInit/getStore seams; mock IdP claims.* overrides. compat 50.6 min (+4.7), mini-core 24.1 (+1.5). e2e 25/25 |
| C17 | done 2026-07-13 | 103/103 | 62.5 KB min / 20.4 gz | +5 scenarios (suite 98→103, NEW area 14-token-apis — all green first mini run). Core clearCache (local sign-out: clearAccount + real's clear-all-msal-keys, activeAccountChanged via setActiveAccount(null)) + hydrateCache (entityFromAccountInfo ApiId 963, id+AT only, KMSI-aware) + top-level loadExternalTokens export (ApiId 964, id/AT/RT per presence, compat re-export composes local-storage+migration); ./broker hybrid acquireTokenByCode({code}) — ApiId 866 redemption w/o code_verifier/redirect_uri, same-code promise dedupe, spa_code_and_nativeAccountId_present; ctx.writeTokens generalized (optional creds, extExpiresOn, RT+foci, kmsi); setActiveAccount event payload dropped (real emits none); post() drops undefined body values. compat 52.5 min (+1.9), mini-core 25.6 (+1.5). e2e 25/25 |
| C18 | done 2026-07-13 | 109/109 | 64.9 KB min / 21.1 gz | +6 scenarios (suite 103→109, NEW area 15-authority — all green first mini run; 3 silent.* snapshots then pinned that real discovers even on cache hits → silentLadder awaits resolveEndpoints). Lazy discovery (initialize does zero network), trust validation (knownAuthorities/cloudDiscoveryMetadata/hardcoded clouds/CIAM/AAD instance-discovery probe → endpoints_resolution_error wrap), endpoint sources config→hardcoded→network with real's /v2.0/ path rule, system.protocolMode (NOT auth — real ignores auth.protocolMode!), instance-aware cloud_instance_host_name token-host swap + cloud_graph_host_name/msgraph_host entity+result fields. Mock IdP: instance_aware fragment extension. compat 55.0 min (+2.5), mini-core 28.0 (+2.4 — discovery is core). e2e 25/25 |
| C19 | done 2026-07-14 | 113/113 | 66.4 KB min / 21.6 gz | +4 scenarios (suite 109→113, NEW area 16-config — all green first mini run). system.allowRedirectInIframe gates acquireTokenRedirect guard + processRedirect bail-out (scenario completes a FULL redirect login in an iframe; iframe src === redirectUri pins real's in-place hRP — navigateToLoginRequestUrl is an hRP OPTION in real 5.16, not config); system.tokenRenewalOffsetSeconds replaces the hardcoded 300s AT buffer (real: now+offset>expiresOn); LogoutRequest logout_hint (explicit/derived login_hint claim) + id_token_hint + eQP on end_session URLs (redirect + popup); serverTelemetryEnabled ports ServerTelemetryManager (current/last wire format, server-telemetry-<clientId> entry, 330B flush cap, 50-error FIFO, clear-on-success, cacheHits; stFail hooks 61/863-always/865/862 via new ctx.stFail). compat 56.4 min (+1.4), mini-core 29.4 (+1.4). e2e 25/25 |
| C20 | done 2026-07-14 | 119/119 | 68.0 KB min / 22.0 gz | +6 scenarios (suite 113→119, 08-telemetry — all green first mini run). Root acquireTokenRedirect event from handleRedirectPromise (cached-request cid, redemption-half ext incl. both networkClientSendPostRequestAsync keys, previousLibraryVersion from pre-init msal.version; clean loads + memoized re-calls emit nothing); failure-event cid joins (AuthError.correlationId stamped by silent/redirect flows — KEY FACT: real's RT token cid is a QUERY param, absent from the body); addPerformanceCallback toString-dedupe + "" stub id (NOT callback-id); no_account_error abandons the measurement (no event) while uninitialized preflight fails DO emit; initialize measured at most once; msal.browser.performance.enabled=1 → mark/measure timeline entries synthesized from the ext tables (real's measure set === C11 ext DurationMs keys + root). compat 58.0 min (+1.6), mini-core 29.5 (+0.1). e2e 25/25 |
| C21 | done 2026-07-14 | 121/121 | 71.2 KB min / 23.3 gz | +2 scenarios (suite 119→121, NEW area 17-pop — green first mini run + 1 fix). NEW ./pop feature (RSA-2048 RS256 keypair — real is RSASSA-PKCS1-v1_5 NOT ECDSA; kid = b64url(sha256(sorted {e,kty,n})); IndexedDB msal.db keystore, unextractable private key) + core scheme plumbing: token_type/req_cnf on auth-code + RT grants, AccessToken_With_AuthScheme entity w/ keyId (pop: from the server AT's cnf.kid, required; ssh: response key_id) + scheme cache-key suffix, scheme-aware AT lookup/save-dedupe, SHR result signing incl. cache-hit RE-sign (header typ,alg,kid; payload at,ts,m,u,nonce,p,q,cnf w/ full sorted public JWK), popKid skips keygen+signing, ssh-cert missing_ssh_jwk/missing_ssh_kid config errors, scheme in throttle + silent-dedupe thumbprints. EXTRA FIX exposed by full-body digest: RT grant redirect_uri only when the request has one (mini sent the config default). Compat +AuthenticationScheme export. compat 61.2 min (+3.2 — judged within the task's ~3 KB gate; descope = drop pop from compose), mini-core 30.9 (+1.4). e2e 25/25 |
| D1 | done 2026-07-14 | 121/121 | 71.7 KB min / 23.4 gz | Seam hardening + consumer packaging. Core stubs for all 6 feature-owned APIs → BrowserAuthError feature_not_configured naming the "@mini-msal/browser/<feature>" import (features overwrite; pop seam guard shares featureError()); NEW `npm run seams` 10/10 vs new conformance-mini-core harness. Packages get tsc dist (JS+d.ts, typescript@7) + 3-condition exports (types/mini-msal-src/default) — internal builds keep compiling from src via resolve.conditionNames; `npm run pack:check` proves npm-pack → throwaway consumer: strict tsc types across all subpaths + rspack tree-shake core 29.9 / popup 32.5 / compat 59.5 / react 65.1 KB min. Stub cost: mini-core 31.3 (+0.4), compat 61.7 (+0.5). e2e 25/25 |
| D2 | done 2026-07-14 | 121/121 | 71.7 KB min / 23.4 gz | Consumer docs. NEW docs/UPGRADING.md (drop-in migration: import swap, 0.6 KB bridge-page swap, cache carry-over incl. msal.0/1/2 migration + KMSI, verification checklist, non-goals) + NEW docs/ALACARTE.md (step-down 61.7 compat → 56.3 explicit → 32.6 core+popup → 29.7 core; per-feature deltas measured: popup +2.9 / pop +2.0 / local-storage +2.6 / cache-migration +3.0 / broker +6.9 / telemetry +9.4 / naa +8.4 / react +10.0). README.md + docs/README.md rewritten as consumer guides (per-profile quick-starts, feature catalog, bridge rule) — stale 75/75-era numbers fixed. All 14 doc code samples mechanically type-checked against dist types (caught 2 sample bugs: MsalProvider instance needs AuthClient & PopupClient; undeclared config). Zero package-code changes: sizes unchanged. e2e 25/25, seams 10/10 |
| D3 | done 2026-07-14 | 121/121 | 71.7 KB min / 23.4 gz | Runnable examples. NEW examples/{core-redirect,core-popup,compat,react} — minimal app + placeholder authConfig + README each, with measured sizes 30.3 / 33.2 / 61.1 / 66.4 KB min (react = React-external, matrix methodology); built as permanent `npm run measure` variants (size regression fixture) plus *-smoke twins (authConfig swapped to the mock IdP via NormalModuleReplacementPlugin, React bundled). NEW `npm run examples:smoke`: headless sign-in per example via real UI clicks (redirect + popup roundtrips + silent token render) — 8/8 first run. READMEs link examples. Zero package-code changes; sizes unchanged. e2e 25/25, seams 10/10 |
| D4 | done 2026-07-14 | 121/121 | 71.7 KB min / 23.4 gz | Requirements audit (skeptic pass). All gates re-run fresh: 121/121, e2e 25/25, seams 10/10, smoke 8/8, pack:check OK, report 0-gap, sizes verified. Bullet 2 + 3 PASS. Bullet 1 GAPS → D5: 18 module exports missing from compat (SignedHttpRequest, stubbedPublicClientApplication, BrowserUtils, ResponseMode, AzureCloudInstance, storage classes, perf/event utils…) found by exhaustive export diff; PCA construction throws raw ReferenceError in Node while real supports SSR construction. Bullet 4 GAP → D6: doc-sample type-check was session-scratch, no committed docs:check. D7 filed as re-audit close-out. Zero code changes |
| D5 | done 2026-07-14 | 122/122 | 72.0 KB min / 23.5 gz | Export-surface completion + SSR construction. Suite 121→122 (init.exported-surface-full pins FULL sorted key+typeof map w/ documented extras allowlist + behavior probes per export; green first mini run). seams 10→12 (node-side dist import: compat PCA + core createClient construct in plain Node; fix = lazy redirectUri closure). 18 exports in NEW compat surface.ts: exact constants, stubbedPublicClientApplication, AuthenticationHeaderParser, EventMessageUtils, EventHandler, Memory/Session/LocalStorage (reuses ./local-storage's exported cookie/AES-GCM helpers — interoperable at-rest), BrowserPerformanceMeasurement, StubPerformanceClient, enforceResourceParameter, BrowserUtils (22 fns), SignedHttpRequest (reuses ./pop's exported makeBoundKeyPair/signPop/keystore + claims-override param). waitForIframe/PopupResponse DESCOPED (absent from real .d.ts). compat 62.1 min (+0.4, tree-shakes away when unused; pack gates 29.9/32.5/59.9/65.4), mini-core 31.3 unchanged. e2e 25/25, smoke 8/8 |
| D6 | done 2026-07-14 | 122/122 | 72.0 KB min / 23.5 gz | Committed doc-sample check. NEW `npm run docs:check` (test/docs/check.mjs): tsc-builds the 3 package dists, extracts every fenced ts/tsx block from README / docs/README / UPGRADING / ALACARTE into test/docs/.samples (one module per sample, `<doc>-L<line>`) and strict-tscs them against dist types via the workspace symlinks' exports "types" condition — 14 samples, 0 opt-outs (`<!-- docs-check:skip -->` supported; 0-samples-extracted hard-fails). Self-tested by reintroducing D2's two sample bugs (dropped MsalProvider cast → TS2322, undeclared config → TS2304): both caught, clean after revert. Scratch kept on failure (gitignored). SOP validate list gains docs:check for docs-affecting tasks. Zero package-code changes; sizes unchanged. e2e 25/25, seams 12/12 |
| D6b | done 2026-07-14 | 122/122 | 72.0 KB min / 23.5 gz | Size-reduction reference doc (user-requested). NEW docs/SIZE.md: plain-language how-it-got-small for repo-outsiders — honest headline ratios (drop-in 3.5×, popup SPA 6.8×, core 7.4×, bridge 10.8×; real costs 220.5 KB regardless of profile), real-bundle byte attribution (analyze: msal-common ~80 KB, interaction_client 44.9, cache 49.7, controllers 25.6), 8 technique sections w/ verified examples (pay-to-play vs StandardController static imports; closures vs class layers + minifier property-name argument; no msal-common tier; one error family + AKA(); telemetry shape tables; 0.6 KB bridge source; WIRE_ID/zero-enums/es2022 habits; sideEffects+subpath exports+pack gates). Linked from README, docs/README, ALACARTE. docs:check DOCS += SIZE.md → 18 samples + 1 justified skip (abridged table excerpt); real excerpts are ```js by policy. bundle-size-experiment.md Results refreshed 75/75-era → 122/122 matrix (~8×→~4.3× library-only). Zero package-code changes; sizes unchanged. e2e 25/25, seams 12/12 |
