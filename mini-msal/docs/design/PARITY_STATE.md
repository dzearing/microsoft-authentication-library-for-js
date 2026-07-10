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
