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
