# Parity work state — mini-msal → full msal-browser 5.16.0 parity

**This is the driving document for an iterative, context-resetting work loop.**
Each Claude session executes exactly ONE pending task below, then resets its
context and continues with the next. Everything a fresh session needs is in
this file plus the docs it links.

## Mission

Bring `@mini-msal/browser` + `@mini-msal/react` (in `mini-msal/packages/`) to
**full behavioral parity** with real `@azure/msal-browser` 5.16.0 +
`@azure/msal-react` 5.5.1 — **75/75 conformance scenarios passing, zero
bugs** — while keeping the bundle as small as possible. Track size every task.

Evidence baseline (2026-07-07): 2 pass / 33 behavioral-diff / 32
missing-feature / 8 bugs. See `docs/GAP_REPORT.md` for per-scenario detail and
`docs/design/conformance-suite-notes.md` for suite internals + hard-won
gotchas (READ BOTH before your first fix).

## Standard Operating Procedure (every cycle)

Work from `mini-msal/` on branch `dzearing/mini-msal`.

1. **Orient**: read this file. Pick the FIRST task with status `pending`.
   Mark it `in-progress` (edit this file).
2. **Test first**: run the task's listed scenarios against mini to see the
   current diffs yourself:
   `node test/conformance/run.mjs --target=mini --grep=<id-or-area>`
   Read the diffs in `test/conformance/results/mini.json`.
3. **Fix**: edit `packages/browser/src/index.ts` / `packages/react/src/index.tsx`.
   Rules:
   - NEVER edit `node_modules/@azure/*`. Reading its source is encouraged.
   - Snapshots in `test/conformance/snapshots/real/` are ground truth —
     do not edit them. If you believe a SCENARIO measures the wrong thing,
     fix the scenario, re-capture real (`--target=real --grep=<id>`), verify
     determinism, and note it in the Decision Log below.
   - Parity means matching real EVEN WHERE REAL IS WORSE (see Decision Log
     defaults). Note such regressions in the log.
   - Keep code style of the existing source; keep it SMALL. Prefer bytes
     over abstraction.
4. **Validate** (all must hold before marking done):
   - `npm run build` clean
   - Task's scenarios: `node test/conformance/run.mjs --target=mini --grep=…`
     → expected diffs eliminated
   - Full run `npm run conformance:mini` → pass count went UP, no scenario
     regressed vs the "pass count" recorded for the previous task below
   - `npm run e2e` → 25/25
   - `npm run measure` → record mini-msal-stack size in the task row
5. **Update this file**: set task status `done` with date, pass count
   (`X/75`), size; append anything newly discovered as new `pending` tasks
   (insert in sensible order); update the Decision Log if you made a judgment
   call.
6. **Regenerate the report**: `node test/conformance/report.mjs` (runner
   "pass" auto-overrides stale classifications). If a classification's
   real/mini prose is now wrong, update it in `test/conformance/report.mjs`.
7. **Commit** everything on `dzearing/mini-msal` with a message
   `parity(<task-id>): <summary>` + the Claude trailer.
8. **Continue the loop**: run `/reset-context` with EXACTLY this continuation
   prompt:

   > Read mini-msal/docs/design/PARITY_STATE.md and follow its Standard
   > Operating Procedure: execute the next pending task (test → fix →
   > validate → update the state doc → commit), then run /reset-context with
   > this same prompt to continue the loop. Stop instead of resetting only
   > when: all tasks are done (75/75 pass, e2e 25/25), you are blocked on a
   > decision only the user can make, or the same task has failed twice.

   If all tasks are done, do NOT reset — write a final summary for the user.

## Environment gotchas (cost hours before — do not rediscover)

- **Stale servers**: if e2e/conformance suddenly 404s or fails en masse,
  a stray `node` from a crashed run is squatting ports 4173/4599:
  `lsof -ti:4173 -ti:4599 | xargs kill`
- **Rebuild required** after touching `packages/*/src` or `test/apps/`:
  the harness/apps are bundled (`npm run build`).
- Everything else (redirect-bridge realities, timeout config names,
  broker/NAA protocols, Playwright popup-event timing trap, IdP inject/config
  endpoints): `docs/design/conformance-suite-notes.md`.
- `DEBUG=1` on any runner spawns servers with visible logs.
- Real MSAL source for reference: `node_modules/@azure/msal-browser/dist/`.

## Decision Log

Defaults already decided (do not re-litigate without user input):
- **Match real even where real is worse** — e.g. real v5.16 has NO popup
  close detection (only bridge timeout). Parity wins; mini's nicer behavior
  goes away. Record each such regression here when you make it.
- Cache stays real-v5 `msal.3` schema; sessionStorage interop must keep
  working (e2e interop checks).
- localStorage mode: implement to match real's OBSERVABLE behavior per the
  snapshots (encrypted `{id,nonce,data}` entities, plaintext key indexes,
  cross-tab events). Ciphertext bytes are normalized `<dynamic>` so any
  AES-GCM w/ per-session cookie key matching real's entry SHAPE passes.
- Scenario timeouts/size: keep suite deterministic; never rely on wall-clock
  ordering.

- **Identity impersonation (user-decided 2026-07-10)**: mini sends real
  MSAL's literal identity values everywhere snapshots compare them —
  `x-client-SKU: msal.js.browser`, `x-client-VER: 5.16.0`, exported
  `version: "5.16.0"`, NAA `clientLibrary: msal.js.browser` /
  `clientLibraryVersion: 5.16.0`. Byte-for-byte parity wins over honest wire
  identity. Keep these in ONE constants block in packages/browser so a future
  un-impersonation is a one-line change.

- **Pay-to-play architecture (user-decided 2026-07-10)**: refactor BEFORE the
  parity work. `@mini-msal/browser` becomes a small composable core
  (`createClient(config, features)`) with features as tree-shakable subpath
  exports added as they're built (`/popup`, `/broker`, `/naa`, `/telemetry`,
  `/local-storage`, `/resilience`, `/logger`); a new `@mini-msal/compat`
  package exports the classic `PublicClientApplication` composing ALL
  features — the drop-in target. **The 75-scenario conformance suite and the
  e2e run against compat.** Features are plain functions/closures (not
  classes) to keep seam bytes near zero and minification strong. Size story
  becomes a matrix (core-only / core+popup / compat) measured continuously.

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

- **A4 (2026-07-10)**: judgment call — "pass count went UP" validation is
  waived when a task's scenario ALSO depends on a later task (here B1's
  result-shape fields); the bar is then "task's expected diffs eliminated +
  zero regressions". Discovered: the conformance harness SORTS result.scopes
  (lib.mjs) — snapshot scope order is alphabetical, not wire order; real
  result scopes = ScopeSet.fromString(AT target).asArray(), casing preserved.

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
- [ ] **A6** `pending` — In-flight dedupe: concurrent identical
  acquireTokenSilent (key: scopes|homeAccountId|authority|policy) share one
  promise → one network call. Scenario: silent.concurrent-dedupe.
- [ ] **A7** `pending` — Interaction lock + unique popup names: second
  interactive call while one is pending → BrowserAuthError
  `interaction_in_progress`; popup window names unique per request
  (`msal.<guid>`-style) so popups never clobber; first call completes.
  Scenario: errors.interaction-in-progress.
- [ ] **A8** `pending` — Don't auto-set active account on login (real never
  does); active account only via setActiveAccount. Check e2e still 25/25
  (the demo app sets active explicitly on login result — verify; if e2e
  depended on auto-active, fix the APP not the library... but app.tsx is
  shared with real MSAL and passes on real, so it must already handle it).
  Scenarios: accounts.multi-account (activeAfterSecondLogin),
  accounts.logout-popup-per-account (active), core.storage-shape-after-login.

### Phase B — systemic behavioral diffs

- [ ] **B1** `pending` — Result shape parity: AuthenticationResult gains
  `authority` (canonical, trailing slash), `correlationId` (generate per
  request if not provided), `tokenType: "Bearer"`, `state` (custom state or
  ""), `fromPlatformBroker: false`, `expiresOn` semantics unchanged;
  `account.environment`; scopes keep REQUEST casing/order (real returns
  granted scopes minus offline_access, request-cased — study snapshots).
  Touches nearly every scenario's `result` block.
- [ ] **B2** `pending` — Event stream parity: add initializeStart/End,
  acquireTokenStart/Success/Failure (with interactionType + real's payload
  keys), popupOpened, handleRedirectStart/End semantics (NO handleRedirect
  events on clean load — real emits none; result null silently),
  logoutStart/Success/End for logout flows, acquireTokenFromNetworkStart on
  forced refresh, event ordering to match snapshots (acquireTokenSuccess
  BEFORE loginSuccess, accountAdded placement, ACTIVE_ACCOUNT_CHANGED).
  EventMessage gains interactionType field. Check react layer still re-renders.
- [ ] **B3** `pending` — Protocol params on authorize/token requests: `nonce`
  (send + VALIDATE id_token nonce claim on redemption), `client-request-id`
  (= correlationId) on authorize query AND token query string, `client_info=1`,
  `claims` default `{"id_token":{"signin_state":…,"login_hint":…}}` merged
  like real, `X-AnchorMailbox` (UPN: for loginHint, Oid: for account),
  `clidata=1`, scope param ordering `<resource scopes> openid profile
  offline_access`, account-derived login_hint on acquireTokenRedirect,
  token redemption `redirect_uri` = request's redirectUri. Compare snapshot
  `idp[*].query/body` blocks exactly (volatile values normalize away).
- [ ] **B4** `pending` — Storage + misc behavior parity: write `msal.version`,
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

- [ ] **C1** `pending` — Request passthrough: `sid`, `domainHint`
  (`domain_hint`), `extraQueryParameters` (authorize), `tokenQueryParameters`
  (token endpoint query), custom `state` (wire format `<libState>|<custom>`,
  echo on result.state), `claims` + `clientCapabilities:["cp1"]` → xms_cc
  merge on authorize+token, per-request `authority` override (discovery for
  that authority, result.authority reflects it, forceRefresh honored with it).
  Scenarios: params.* (all remaining).
- [ ] **C2** `pending` — Resilience: 429 throttle cache (storage entry shaped
  like real's `throttling.*`, immediate retry re-throws same ServerError with
  NO network until Retry-After passes); ServerError formatting for 5xx
  (message format per snapshot); `refresh_in` → `refreshOn` on the AT entity
  + proactive-refresh observable behavior per snapshot
  (resilience.proactive-refresh: what real did — study snapshot first).
  Scenarios: resilience.*.
- [ ] **C3** `pending` — Server telemetry + correlationId: token BODY params
  with real's literal identity values per the Decision Log
  (`x-client-SKU: msal.js.browser`, `x-client-VER: 5.16.0`),
  x-client-current/last-telemetry (real format: `5|apiId,cacheStatus,,,|,`
  — study snapshots incl. last-telemetry-after-failure), x-ms-lib-capability
  `retry-after, h429`, client-request-id token QUERY param;
  request.correlationId accepted + surfaced (result, events, headers).
  Scenarios: telemetry.token-request-headers,
  telemetry.correlation-id-propagation, telemetry.last-telemetry-after-failure.
- [ ] **C4** `pending` — Perf events: `addPerformanceCallback` +
  BrowserPerformanceClient-equivalent opt-in via `telemetry.client` config;
  emit `initializeClientApplication`, `acquireTokenPopup`,
  `acquireTokenSilent` (+ cache-hit variant) events with
  name/success/durationMs/correlationId. Harness passes perfClient:true only
  when lib.BrowserPerformanceClient exists — export a compatible class.
  Scenarios: telemetry.perf-events-*.
- [ ] **C5** `pending` — Init & surface: `getConfiguration()` returning
  real-default-shaped config (values per init.get-configuration snapshot);
  logger (`system.loggerOptions.loggerCallback`, Info+Verbose volume);
  initializeStart/End already from B2; export surface per
  init.exported-surface snapshot: version "5.16.0" (Decision Log:
  impersonation), createStandardPublicClientApplication,
  isPlatformBrokerAvailable, full EventType map incl. broker/bfcache strings,
  InteractionType.None, PromptValue, ProtocolMode, BrowserCacheLocation,
  OIDC_DEFAULT_SCOPES, BrowserAuthErrorCodes (52 keys), error classes.
  Scenarios: init.* remaining.
- [ ] **C6** `pending` — localStorage + cross-tab: `cacheLocation:
  "localStorage"` with real's observable shape (encrypted entities
  `{id:<guid>, nonce:<b64>, data:<ciphertext>}` via AES-GCM, session key in
  a cookie like real — study real's LocalStorage/CookieStorage source),
  plaintext key indexes + msal.version in localStorage, account/logout events
  propagating to other tabs (storage listener). sessionStorage remains
  default; interop e2e must stay green. Scenarios: accounts.local-storage,
  accounts.cross-tab-events.
- [ ] **C7** `pending` — Platform broker, DOM transport: config
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
- [ ] **C8** `pending` — Platform broker, extension transport: Handshake via
  window.postMessage + MessageChannel (channel id, preferred extension id,
  2s default timeout `nativeBrokerHandshakeTimeout`, retry with undefined
  extensionId, swallow failures), GetToken over the port, Response/Success
  mapping identical to C7. First-login gating parity (no nativeAccountId →
  web flow). Scenarios: broker.extension-*, broker.dom-first-login.
- [ ] **C9** `pending` — NAA: `createNestablePublicClientApplication` +
  window.nestedAppAuthBridge (GetInitContext handshake envelope with
  clientLibrary `msal.js.browser` / clientLibraryVersion `5.16.0` per
  Decision Log), GetToken/GetTokenPopup mapping, browser-cache-first silent, error
  mapping per docs/design/naa-protocol.md, unsupported APIs →
  NestedAppAuthError `unsupported_method`, no-bridge → standard PCA
  fallback. Scenarios: naa.*.
- [ ] **C10** `pending` — Final sweep: full `npm run conformance` (expect
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
