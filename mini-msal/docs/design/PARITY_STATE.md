# Parity work state — mini-msal → full msal-browser 5.16.0 parity

**This is the driving document for an iterative, context-resetting work loop.**
Each Claude session executes exactly ONE pending task below, then resets its
context and continues with the next. Everything a fresh session needs is in
this file plus the per-task Context pointers.

## Mission

Deliver a drop-in msal replacement that is also à-la-carte consumable
(user requirements, restated 2026-07-12). Definition of done — ALL of:

1. **100% compat**: `@mini-msal/compat` is a one-import drop-in for
   `@azure/msal-browser` — all conformance scenarios passing, zero bugs,
   e2e 25/25.
2. **Pay-to-play**: every feature is an independently consumable,
   tree-shakable module composed into `createClient`; nobody pays bytes for
   features they don't compose.
3. **Effortless consumption**: it must be VERY straightforward to consume
   just the parts you need — verified from a consumer's point of view, with
   clear error behavior when a non-composed feature is invoked (D-series).
4. **Docs & examples**: consumer-facing documentation and runnable examples
   for each consumption profile, each with its measured size (D-series).

Track bundle size on every task (core + compat).

**Status (2026-07-14)**: Phases A0–C21 + D1–D5 COMPLETE — conformance
122/122, e2e 25/25, seams 12/12, examples:smoke 8/8, pack:check green,
GAP_REPORT all-pass. Size matrix: compat-no-react 62.1 KB min (real:
220.5), compat+react stack 72.0 (real: 248.8), core-only 31.3. Packages
consumer-ready (tsc dist + types, npm-pack-verified); consumer docs +
runnable size-tracked examples shipped. D5 closed the D4 audit's bullet-1
gaps: full 51-key export surface pinned by scenario, 18 missing exports
implemented, PCA constructs in plain Node (SSR). Remaining: D6 (committed
docs:check), D7 (close-out re-audit).
History and per-task decisions: [PARITY_LOG.md](./PARITY_LOG.md).

## Context budget (user-required 2026-07-13)

Keep each task session below **~250k tokens**. Rules:

- Read ONLY: this file, the task's `Context:` pointers, and the files you
  edit. Everything else is on-demand: **grep, don't read wholesale** —
  especially `PARITY_LOG.md`, `GAP_REPORT.md`, real MSAL dist files, and
  snapshots (read single scenarios' snapshots, never the directory).
- `PARITY_LOG.md` holds all A0–C10 history/decisions. When a pending task
  needs background on a past judgment call, grep it by task id (`C6`, `B3`…).
- If a session balloons toward the budget mid-task, finish the smallest safe
  checkpoint (build green + no regressions), update this doc honestly
  (status stays `in-progress`, note what remains), commit, and reset.

## Standard Operating Procedure (every cycle)

Work from `mini-msal/` on branch `dzearing/mini-msal`.

1. **Orient**: read this file. Pick the FIRST task with status `pending`.
   Mark it `in-progress` (edit this file). Read the task's `Context:`
   pointers — nothing more.
2. **Test first**: run the task's listed scenarios against mini to see the
   current diffs yourself:
   `node test/conformance/run.mjs --target=mini --grep=<id-or-area>`
   Read the diffs in `test/conformance/results/mini.json`. For NEW
   scenarios (C-gaps tasks): write the scenario, capture real
   (`--target=real --grep=<id>`), verify determinism (run twice), then run
   mini against it. Suite internals + hard-won gotchas:
   `docs/design/conformance-suite-notes.md` (grep it, e.g. for "timeout",
   "popup", "idp").
3. **Fix**: edit `packages/browser/src/*.ts` / `packages/compat/src/` /
   `packages/react/src/index.tsx`. Rules:
   - NEVER edit `node_modules/@azure/*`. Reading its source is encouraged
     (grep/targeted reads in `node_modules/@azure/msal-browser/dist/`).
   - Snapshots in `test/conformance/snapshots/real/` are ground truth —
     do not edit them. If a SCENARIO measures the wrong thing, fix the
     scenario, re-capture real, verify determinism, note it in PARITY_LOG.
   - Parity means matching real EVEN WHERE REAL IS WORSE. Record such
     regressions in PARITY_LOG's Decision Log.
   - Keep code style of the existing source; keep it SMALL. Prefer bytes
     over abstraction.
4. **Validate** (all must hold before marking done):
   - `npm run build` clean
   - Task's scenarios green; full `npm run conformance:mini` → pass count
     did not regress vs the previous task's recorded count (suite may GROW
     when a task adds scenarios — record the new total as `X/<total>`)
   - `npm run e2e` → 25/25
   - `npm run seams` → 12/12 (seam-hardening + node SSR checks, cheap; D1/D5)
   - `npm run measure` → record mini-msal-stack size in the task row
   - `npm run pack:check` (needs network) only when a task touches
     package.json exports, tsconfigs, or public API surface
5. **Update docs**: set task status `done` with date, pass count, size in
   this file; append newly discovered work as new `pending` tasks here;
   append the task's decision entry to `PARITY_LOG.md` (Decision Log
   section) and a row to its progress table. Regenerate the report:
   `node test/conformance/report.mjs`.
6. **Commit** everything on `dzearing/mini-msal` with a message
   `parity(<task-id>): <summary>` + the Claude trailer. Push:
   `git push fork dzearing/mini-msal` (origin rejects pushes).
7. **Continue the loop**: run `/reset-context` with EXACTLY this
   continuation prompt:

   > Read mini-msal/docs/design/PARITY_STATE.md and follow its Standard
   > Operating Procedure: execute the next pending task (test → fix →
   > validate → update the state doc → commit + push), then run
   > /reset-context with this same prompt to continue the loop. Stop instead
   > of resetting only when: every task in the task list is done, you are
   > blocked on a decision only the user can make, or the same task has
   > failed twice.

   If all tasks are done, do NOT reset — write a final summary for the user.

## Environment gotchas (cost hours before — do not rediscover)

- **Stale servers**: if e2e/conformance suddenly 404s or fails en masse,
  a stray `node` from a crashed run is squatting ports 4173/4599:
  `lsof -ti:4173 -ti:4599 | xargs kill`
- **Rebuild required** after touching `packages/*/src` or `test/apps/`:
  the harness/apps are bundled (`npm run build`).
- Only ONE loop session may run at a time (two concurrent sessions kill
  each other's servers — see A1 entry in PARITY_LOG).
- `DEBUG=1` on any runner spawns servers with visible logs.
- Real MSAL source for reference: `node_modules/@azure/msal-browser/dist/`.
- Everything else (redirect-bridge realities, timeout config names,
  broker/NAA protocols, Playwright popup-event timing trap, IdP
  inject/config endpoints): `docs/design/conformance-suite-notes.md`.

## Decision defaults (do not re-litigate without user input)

- **Match real even where real is worse** — parity wins over nicer behavior;
  record each such regression in PARITY_LOG.
- Cache stays real-v5 `msal.3` schema; sessionStorage interop must keep
  working (e2e interop checks).
- Scenario timeouts/size: keep suite deterministic; never rely on wall-clock
  ordering.
- **Identity impersonation (user-decided 2026-07-10)**: mini sends real
  MSAL's literal identity values everywhere snapshots compare them —
  `x-client-SKU: msal.js.browser`, `x-client-VER: 5.16.0`, exported
  `version: "5.16.0"`, NAA `clientLibrary`/`clientLibraryVersion`. Kept in
  ONE constants block in packages/browser.
- **Pay-to-play architecture (user-decided 2026-07-10)**: core =
  `createClient(config, features)`; features are plain closures as subpath
  exports (`/popup`, `/broker`, `/naa`, `/local-storage`, `/telemetry`,
  `/redirect-bridge`); `@mini-msal/compat` composes ALL features into the
  classic `PublicClientApplication` — conformance + e2e run against compat.
- Dated per-task decision entries (A1–C10): grep
  [PARITY_LOG.md](./PARITY_LOG.md).

## Reference docs (read on demand, never wholesale)

- `PARITY_LOG.md` — A0–C10 history: decision entries, completed task
  descriptions, progress table with sizes.
- `conformance-suite-notes.md` — suite internals, harness gotchas.
- `../GAP_REPORT.md` — generated per-scenario report (currently all-pass).
- `broker-protocol.md`, `naa-protocol.md` — wire protocols for fakes.
- `bundle-size-experiment.md` — size methodology + final matrix.

## Task list (execute strictly top-to-bottom)

Statuses: `pending` | `in-progress` | `done <date> — pass X/<total>, mini-stack Y KB`

### Phases A0–C10 — COMPLETE (baseline 2/75 → 75/75)

All done 2026-07-10 → 2026-07-12; full descriptions, per-task decisions,
pass counts and sizes: [PARITY_LOG.md](./PARITY_LOG.md).

### Phase C-gaps — parity gaps beyond the 75 scenarios (adversarial audit 2026-07-13)

C10 proved parity **as observed by the suite**; the user flagged that
unobserved surface (e.g. perf-event fields) can still break real consumers.
A 47-agent adversarial audit (2026-07-13) confirmed 40 gaps (0 refuted).
**Evidence file: `audit-findings-2026-07-13.json` (same dir) — grep it by
the finding ids listed in each task; each finding carries verified
file:line evidence on both sides (`verdict.reason`) and factual
corrections (`verdict.corrections`). Read the task's findings BEFORE
implementing; never read the file wholesale (~44k tokens).**

Each task is TEST-FIRST: extend the conformance suite (new scenario or
widened comparison), re-capture real snapshots
(`--target=real --grep=<id>`), verify determinism (run twice), THEN
implement mini to match. Suite growth means pass counts read
`X/<new total>`. Keep new scenarios deterministic against the mock IdP
(see conformance-suite-notes.md). If a task's feature turns out
substantially larger than its session budget, split it: land the
scenarios + a partial implementation, add a follow-up task, note it here.

- [x] **C11** `done 2026-07-13 — pass 77/77, mini-stack 46.8 KB` — Perf-event
  field parity. Suite grew 75→77 (telemetry.perf-event-shape,
  -shape-silent): full per-flow event shapes pinned for initialize / popup /
  silent cache-hit / silent refresh / ssoSilent / silent failure — top-level
  keys AND deterministic values exact, ext compared by key set, context by
  string presence (decision: real's internal call-tree blob not replicated).
  telemetry.ts rewritten table-driven; core seam `err.silentRefreshReason`;
  compat 39.7 KB min (+7.2), mini-core unchanged 19.5. acquireTokenPreRedirect
  not observable pre-C13; redirect perf events are C20's. Details/gotchas
  (43-char ext-key normalization collision!): PARITY_LOG C11 entry.

- [x] **C12** `done 2026-07-13 — pass 78/78, mini-stack 50.0 KB` — API
  surface: logger + error-code namespaces. Suite grew 77→78 (new
  init.exported-surface-2, deterministic, mini green first run). Browser
  core: `Logger` class (real's clone/level-gate/message format), `LogLevel`
  (incl. reverse mappings), `WrapperSKU`, and getLogger/setLogger/
  initializeWrapperLibrary on AuthClient (wrapper meta stored only —
  telemetry headers stay stub-empty until C19). Compat: 5 *ErrorCodes
  namespaces via a shared `pack()` (irregular keys as `key=code` entries —
  5 in ClientAuthErrorCodes, 1 in ClientConfigurationErrorCodes) +
  BrowserConfigurationAuthError. compat 42.9 KB min (+3.2), core 20.9
  (+1.4 — Logger is core since getLogger is core surface). e2e 25/25.
- [x] **C13** `done 2026-07-13 — pass 83/83, mini-stack 51.6 KB` — Redirect
  navigation seams. Suite grew 78→83 (new area 11-navigation: deep-link
  replay, onRedirectNavigate cancel login/logout, navigationClient
  config/setter — all green first mini run). One seam like real:
  NavigationClient class (core export + compat re-export, default = real's
  replace/assign + reject-after-timeout promise), system.navigationClient +
  setNavigationClient(), auth.onRedirectNavigate cancel hook (acquire keeps
  the lock, logout releases + logoutEnd), navigateToLoginRequestUrl replay
  via real's msal.{cid}.request.origin / urlHash keys (redirectStartPage,
  homepage fallback, in-place #hash restore, navigateInternal-false =
  process in place). Redirect logoutStart payload now the raw request.
  compat 44.6 KB min (+1.7), core 22.5 (+1.6). e2e 25/25 (cancelled-login
  seed gained request.origin). Details: PARITY_LOG C13 entry.
- [x] **C14** `done 2026-07-13 — pass 87/87, mini-stack 53.0 KB` — Popup
  behaviors (findings `navigate-popups`, `logout-popup-main-window-
  redirect`). Suite grew 83→87 (core.popup-open-timing[-async],
  core.popup-window-attributes, accounts.logout-popup-main-window-redirect
  — all green first mini run after impl). system.navigatePopups (default
  true): about:blank opened SYNCHRONOUSLY in the caller's stack then
  location.assign'd; real's openSizedPopup geometry/features string +
  generate[Logout]PopupName; request.popupWindowAttributes/
  popupWindowParent; logoutPopup mainWindowRedirectUri navigates the main
  window via new ctx.navigate seam (ApiId 962; lock survives navigation,
  like real); telemetry isAsyncPopup wired. compat 45.9 KB min (+1.3),
  core 22.6 (+0.1). Details: PARITY_LOG C14 entry.
- [x] **C15** `done 2026-07-13 — pass 93/93, mini-stack 55.8 KB` — React
  bindings parity. Suite grew 87→93 (new area 12-react, run against NEW
  dual-built react harness pages `conformance-react-{real,mini}` — React
  bundled, fixtures + `__mount` in `test/apps/harness/react-setup.tsx`;
  all 6 green first mini run). `packages/react` rewritten as a port of
  msal-react 5.5.1: provider initialize()s the instance +
  initializeWrapperLibrary + full 5-value InteractionStatus reducer
  (real's event mapping incl. clear-guards + RESTORE_FROM_BFCACHE, accounts
  frozen [] during startup), context logger, useMsalAuthentication
  {login, acquireToken, result, error} (auto-acquire for signed-in active
  account, IRAE→interaction fallback, result reset on logout),
  useIsAuthenticated(identifiers)/useAccount (case-insensitive ids,
  empty-filter→active account, iat/nonce equality), templates with
  identifier props + function-as-children, MsalAuthenticationTemplate
  full error contract (spread result into ErrorComponent, THROW without
  one, LoadingComponent gets context). Core/compat: + InteractionStatus
  export. compat unchanged 45.9, core 22.6. Gotchas (logout popups need
  bridge postLogoutRedirectUri; neither stack auto-sets active account):
  PARITY_LOG C15 entry.
- [x] **C16** `done 2026-07-13 — pass 98/98, mini-stack 60.6 KB` — Cache
  entity semantics. Suite grew 93→98 (new area 13-cache: at-scope-dedupe,
  at-multi-match-clear, tenant-profile-merge, schema-migration,
  kmsi-plaintext-localstorage — all green first mini run after impl).
  Mock IdP gained `/config?claims.<name>=<json>` id_token-claim overrides
  (drives guest-tid + signin_state). Core: saveAccessToken intersecting-
  scope dedupe (OIDC-stripped) + >1-match clear-on-lookup; base-account
  merge (mergeAccount = real's buildAccountToCache: one entity per
  homeAccountId+environment, tenantProfiles appended, isHomeTenant
  computed) shared by web/broker/naa writers; getAllAccounts expands
  profiles into per-tenant AccountInfos (tenantProfiles now a Map,
  per-tenant idToken/claims/kmsi/loginHint/upn sourced from claims). KMSI:
  Store.setUser(key, value, kmsi) seam — ./local-storage persists KMSI
  entities PLAINTEXT (survive cookie loss); AccountInfo.kmsi real values.
  NEW /cache-migration feature (compat-composed; core got onInit/getStore
  seams): msal.0/1/2→msal.3 migration + 5-day retention TTL. compat 50.6
  min (+4.7), core 24.1 (+1.5). e2e 25/25. Details: PARITY_LOG C16 entry.
- [x] **C17** `done 2026-07-13 — pass 103/103, mini-stack 62.5 KB` —
  Programmatic token APIs. Suite grew 98→103 (new area 14-token-apis:
  clear-cache, hydrate-cache, load-external-tokens, acquire-token-by-code,
  acquire-token-by-code-errors — all green first mini run; area's network
  digest excludes discovery, C18's concern). Core: clearCache (local
  sign-out — clearAccount + real's clear-all-msal-keys incl. msal.version,
  activeAccountChanged via setActiveAccount(null)), hydrateCache
  (entityFromAccountInfo ApiId 963, id+AT only, KMSI-aware), top-level
  loadExternalTokens export (ApiId 964; 5th param = features, compat
  re-export composes local-storage + cache-migration), setActiveAccount
  now emits NO payload (real), post() drops undefined body values,
  ctx.writeTokens generalized (optional creds, extExpiresOn, RT+foci,
  kmsi). ./broker: hybrid acquireTokenByCode({code}) — ApiId 866
  redemption without code_verifier/redirect_uri, same-code promise
  dedupe, spa_code_and_nativeAccountId_present. compat 52.5 min (+1.9),
  core 25.6 (+1.5). e2e 25/25. Details: PARITY_LOG C17 entry.
- [x] **C18** `done 2026-07-13 — pass 109/109, mini-stack 64.9 KB` —
  Authority modes & discovery. Suite grew 103→109 (new area 15-authority:
  lazy-discovery-default-path, oidc-discovery-endpoint-path,
  authority-metadata-config, known-authorities-validation,
  hardcoded-cloud-metadata, instance-aware-cloud-instance — all green
  first mini run; 3 silent.* snapshots then forced resolveEndpoints into
  silentLadder — real discovers even on cache hits). Lazy discovery
  (initialize = zero network), trust chain (cloudDiscoveryMetadata /
  knownAuthorities / 13 hardcoded cloud aliases / .ciamlogin.com / AAD
  instance-discovery probe), endpoint sources auth.authorityMetadata →
  hardcoded 6-host templates → network with real's /v2.0/ path rule.
  KEY CAPTURE FACTS: real reads protocolMode from config.SYSTEM (ignores
  stdConfig's auth.protocolMode — suite always ran AAD mode), and ALL
  resolution failures surface as ClientAuthError
  `endpoints_resolution_error` (createDiscoveredInstance wraps
  untrusted_authority etc.). Instance-aware: waitForCode → {code,
  cloud*} object (ctx seam change), token-endpoint host swap, cloud
  graph fields cached on new base accounts + read from entity on ALL
  result paths; logoutUrl now async. Mock IdP: instance_aware fragment
  extension. compat 55.0 min (+2.5), core 28.0 (+2.4 — discovery is
  core). e2e 25/25. Details: PARITY_LOG C18 entry.
- [x] **C19** `done 2026-07-14 — pass 113/113, mini-stack 66.4 KB` — Config
  knobs + logout params. Suite grew 109→113 (new area 16-config:
  allow-redirect-in-iframe, token-renewal-offset, logout-hint-params,
  server-telemetry-enabled — all green first mini run after impl).
  system.allowRedirectInIframe gates BOTH the acquireTokenRedirect guard
  and processRedirect's iframe bail-out; system.tokenRenewalOffsetSeconds
  replaces the hardcoded 300s AT buffer; logoutUrl + new exported
  LogoutRequest carry logout_hint (explicit or derived account.loginHint /
  login_hint claim), id_token_hint, eQP (appended last, non-overriding) —
  popup + redirect logout share the builder. serverTelemetryEnabled:
  real's ServerTelemetryManager ported (5|apiId,0,,,|sku,ver current;
  5|hits|fails|errors|n,overflow last; server-telemetry-<clientId> entry,
  330-byte flush cap, FIFO at 50, cleared-on-success, cacheHits++ on AT
  hits); stFail hooks: RT=61, iframe=863 ALWAYS (real's
  SilentIframeClient is created with ApiId.ssoSilent even on the silent
  ladder), redirect hRP=865, popup=862 via new ctx.stFail seam. KEY
  CAPTURE FACTS: navigateToLoginRequestUrl is a hRP OPTION in real 5.16
  (config.auth flag ignored there — iframe scenario pins the in-place
  path by making iframe src === redirectUri); real's offset check is
  now+offset>expiresOn, so offset<lifetime still cache-hits. compat 56.4
  min (+1.4), core 29.4 (+1.4). e2e 25/25. Details: PARITY_LOG C19 entry.
- [x] **C20** `done 2026-07-14 — pass 119/119, mini-stack 68.0 KB` —
  Perf-event emission semantics. Suite grew 113→119 (6 new 08-telemetry
  scenarios, all green first mini run after impl): root
  acquireTokenRedirect event from handleRedirectPromise (cached-request
  cid, redemption-half ext, previousLibraryVersion from pre-init
  msal.version; clean loads + memoized re-calls emit nothing); failed
  silent w/o app cid: event cid === new AuthError.correlationId === wire
  cid (KEY CAPTURE FACT: real 5.16's RT token cid rides the QUERY string,
  the POST body has none); addPerformanceCallback toString-dedupe
  returning the existing id, stub (no perf client) id is "" NOT the
  audit's "callback-id"; no_account_error abandons the measurement (zero
  events) while uninitialized preflight failures DO emit success:false
  (audit correction confirmed); initialize measured at most once;
  msal.browser.performance.enabled='1' + perf client → mark/measure
  timeline entries — real's surviving measure set for a silent cache-hit
  is EXACTLY C11's ext DurationMs keys + root, so mini synthesizes from
  its ext tables at emit time. compat 58.0 min (+1.6), core 29.5 (+0.1).
  e2e 25/25. Details: PARITY_LOG C20 entry.
- [x] **C21** `done 2026-07-14 — pass 121/121, mini-stack 71.2 KB` —
  authenticationScheme "pop"/"ssh-cert". Suite grew 119→121 (new area
  17-pop: silent-shr, ssh-scheme-and-errors — green first mini run + one
  fix). NEW `./pop` feature (compat-composed): RSA-2048 RS256 keypair
  (real is RSASSA-PKCS1-v1_5, NOT the audit's "ECDSA"), kid =
  b64url(sha256(sorted {e,kty,n})), keys persist in real's IndexedDB
  msal.db keystore (private key unextractable). Core: token_type/req_cnf
  on both grants, AccessToken_With_AuthScheme entity + keyId (pop: server
  AT's cnf.kid, required; ssh: response key_id) + scheme cache-key suffix,
  scheme-aware AT lookup/save-dedupe, SHR signing incl. cache-hit RE-sign
  (fresh nonce/ts), popKid skips keygen+signing, missing_ssh_jwk/kid
  config errors, scheme rides throttle + silent-dedupe thumbprints and
  survives the redirect roundtrip. EXTRA FIX (exposed by the new
  full-body wire digest): RT grants carry redirect_uri ONLY when the
  request passes one. SIZE CALL: compat 61.2 min (+3.2, marginally over
  the task's "~3 KB" gate — judged within tolerance since it's all
  scenario-pinned observable parity; descoping stays a one-line revert:
  drop `pop` from compat's compose list). core 30.9 (+1.4). e2e 25/25.
  Details: PARITY_LOG C21 entry.

### Phase D — à-la-carte DX, docs & examples (after C-gaps)

The C-series proves compat correctness; the D-series makes the à-la-carte
story real for consumers. Same SOP applies (validate → update doc → commit +
push → reset). "Consumer" below means someone who has never read this repo.

- [x] **D1** `done 2026-07-14 — pass 121/121, mini-stack 71.7 KB` — Seam
  hardening + consumer packaging. (a) Core installs throwing stubs for all 6
  feature-owned APIs (loginPopup/acquireTokenPopup/logoutPopup → ./popup,
  acquireTokenByCode → ./broker, add/removePerformanceCallback →
  ./telemetry): BrowserAuthError `feature_not_configured` whose message
  names the exact import (shared exported `featureError()`; the ./pop seam
  guard uses it too). NEW `npm run seams` (test/unit/seams.mjs, 10/10)
  against new `conformance-mini-core` harness — covers core-only + partial
  compositions. (b) Packages now ship tsc-built dist (JS + .d.ts;
  typescript@7 devDep; 5 type-only fixes) with exports conditions
  types/mini-msal-src/default — internal rspack builds still compile from
  src (conditionNames), consumers get dist. NEW `npm run pack:check`
  (network): npm pack → throwaway consumer → strict-tsc across every
  subpath export + rspack tree-shake gates (core 29.9 / popup 32.5 /
  compat 59.5 / react 65.1 KB min). Stub bytes: core 31.3 (+0.4), compat
  61.7 (+0.5). e2e 25/25. Details: PARITY_LOG D1 entry.
- [x] **D2** `done 2026-07-14 — pass 121/121, mini-stack 71.7 KB` — Consumer
  docs (user-required deliverables named 2026-07-13). Shipped: (a)
  `docs/UPGRADING.md` — drop-in migration (import swap, bridge swap, cache
  carry-over/KMSI/msal.0-2 migration, verification checklist, non-goals:
  B2C/ADFS only — PoP + CIAM are in); (b) `docs/ALACARTE.md` — step-down
  61.7 compat → 56.3 explicit-composition → 32.6 core+popup → 29.7 core,
  feature catalog with NEWLY MEASURED per-feature deltas (popup +2.9 /
  pop +2.0 / local-storage +2.6 / cache-migration +3.0 / broker +6.9 /
  telemetry +9.4 / naa +8.4 / react +10.0; same minify settings as the
  matrix), bridge-page rule (needed for popup/ssoSilent/silent-iframe,
  not pure-redirect); (c) README.md + docs/README.md rewritten as consumer
  guides (per-profile quick-starts, feature-catalog table, bridge
  requirement) with the stale 75/75-era numbers corrected. All 14 doc code
  samples extracted + strict-type-checked against the packages' dist types
  (caught: MsalProvider instance needs AuthClient & PopupClient cast for
  à-la-carte). Zero package-code changes — sizes unchanged. e2e 25/25,
  seams 10/10. Original spec:
  (a) `docs/UPGRADING.md` — basic drop-in usage/upgrade doc: migrating from
  `@azure/msal-browser`/`@azure/msal-react` to `@mini-msal/compat` +
  `@mini-msal/react` (change one import; cache carries over — signed-in
  users stay signed in; the redirect-bridge page swap `/popup.html` →
  mini's 0.6 KB bridge; what to verify after switching; known
  non-goals like B2C — note PoP IS supported since C21 via ./pop).
  (b) `docs/ALACARTE.md` — "lower your bundle cost" guide: start from
  compat, then step down profile-by-profile (compat 32.5 KB → core+popup →
  core-only 19.5 KB), what each feature module adds in measured KB and which
  API calls require it, with before/after size numbers per step and the
  exact createClient composition for each.
  (c) Rewrite `mini-msal/README.md` + `docs/README.md` "How you consume it"
  into a real consumer guide linking both: quick-start per profile (compat
  drop-in; core-only redirect SPA; core+popup; +react), a feature-catalog
  table (import path · what it adds · measured KB cost from the size
  matrix), and the redirect-bridge page requirement. Every code sample in
  all three docs must be copy-paste runnable against the packages as they
  exist.
  Context: `docs/README.md`, `README.md`, size numbers from
  `bundle-size-experiment.md` (final matrix) or `npm run measure`.
- [x] **D3** `done 2026-07-14 — pass 121/121, mini-stack 71.7 KB` — Runnable
  examples with measured sizes. NEW `examples/{core-redirect,core-popup,
  compat,react}` — one minimal self-contained app per profile (main +
  placeholder authConfig + README stating composition/size/bridge rule) —
  measured 30.3 / 33.2 / 61.1 / 66.4 KB min (react = React-external, matrix
  methodology); each is a permanent `npm run measure` variant (size
  regression fixture) with an `example-*-smoke` twin (authConfig swapped to
  the mock IdP via rspack NormalModuleReplacementPlugin, React bundled).
  NEW `npm run examples:smoke` (test/examples/smoke.mjs, e2e-harness
  pattern): drives each example's real UI — #signin click, redirect or
  popup+bridge roundtrip — asserts greeting + silently-acquired token;
  8/8 first run. README/docs link examples. Zero package-code changes —
  sizes unchanged. e2e 25/25, seams 10/10. Details: PARITY_LOG D3 entry.
- [x] **D4** `done 2026-07-14 — pass 121/121, mini-stack 71.7 KB` —
  Requirements audit vs this Mission. All gates re-run fresh this session
  (conformance 121/121, e2e 25/25, seams 10/10, examples:smoke 8/8,
  pack:check OK, report 0-gap, measure matches all recorded sizes).
  Bullets 2 (pay-to-play) + 3 (effortless consumption) PASS with evidence.
  Bullet 1 GAPS → filed D5: exhaustive module-export diff found 18 real
  exports missing from compat, and mini's PCA construction throws a raw
  ReferenceError in Node where real supports SSR construction. Bullet 4
  GAP → filed D6: doc-sample type-checking was D2 session-scratch, not a
  committed check. D7 filed as the close-out re-audit. Zero code changes.
  Details/evidence: PARITY_LOG D4 entry.

- [x] **D5** `done 2026-07-14 — pass 122/122, mini-stack 72.0 KB` — Compat
  export-surface completion + SSR-safe construction. Suite grew 121→122
  (init.exported-surface-full: FULL sorted key+typeof map — mini extras
  NativeAuthError/NestedAppAuthError/createAuth filtered by a documented
  allowlist — plus behavior probes for every new export; deterministic,
  green first mini run after impl). seams grew 10→12: node-side SSR
  checks import the built dist in-process — compat PCA + core
  createClient construct in plain Node (fix: createClient's redirectUri
  is now a lazy closure; ops still require a browser, like real). All 18
  exports implemented in NEW packages/compat/src/surface.ts as
  faithful-lean ports: 6 exact constants (ApiId, AzureCloudInstance,
  JsonWebTokenTypes, ResponseMode, DEFAULT_IFRAME_TIMEOUT_MS,
  BrowserRootPerformanceEvents), stubbedPublicClientApplication (26 keys,
  BrowserConfigurationAuthError rejections), AuthenticationHeaderParser,
  EventMessageUtils, EventHandler, MemoryStorage/SessionStorage +
  LocalStorage (reuses ./local-storage's newly exported cookie/AES-GCM
  helpers — interoperates with the feature's at-rest format),
  BrowserPerformanceMeasurement, StubPerformanceClient,
  enforceResourceParameter, BrowserUtils namespace (22 fns),
  SignedHttpRequest (reuses ./pop's newly exported
  makeBoundKeyPair/signPop/keystore; signPop gained real's
  claims-override param). PCA instance extras
  waitForIframeResponse/waitForPopupResponse DESCOPED (not in real's
  .d.ts). compat 62.1 min (+0.4 — surface tree-shakes away for consumers
  who don't import it, pack:check gates 29.9/32.5/59.9/65.4), core 31.3
  (unchanged). e2e 25/25, smoke 8/8. Doc sizes/counts refreshed. Details:
  PARITY_LOG D5 entry. Original spec:
  (a) NEW scenario `init.exported-surface-full` — in the harness, pin the
  FULL sorted `Object.keys(lib)` list (real snapshot is ground truth) so
  the export surface can never silently diverge again; (b) NEW node-side
  check (extend `test/unit/seams.mjs` or a small `test/unit/ssr.mjs` wired
  into `npm run seams`): `new PublicClientApplication({auth:{clientId}})`
  in plain Node must construct without throwing (real does — SSR/Next.js
  drop-in; ops may fail later), currently mini throws
  `ReferenceError: location is not defined` from createClient's eager
  `new URL(..., location.href)`. Then implement the 18 missing compat
  exports: ApiId, AuthenticationHeaderParser, AzureCloudInstance,
  BrowserPerformanceMeasurement, BrowserRootPerformanceEvents,
  BrowserUtils, DEFAULT_IFRAME_TIMEOUT_MS, EventHandler,
  EventMessageUtils, JsonWebTokenTypes, LocalStorage, MemoryStorage,
  SessionStorage, ResponseMode, SignedHttpRequest (reuse ./pop's SHR
  machinery), StubPerformanceClient, enforceResourceParameter,
  stubbedPublicClientApplication. Match real's observable behavior for
  each (grep real dist per symbol); enums/constants exact; classes may be
  thin ports as long as scenario-observable behavior matches. Runtime-only
  `waitForIframeResponse`/`waitForPopupResponse` (absent from real .d.ts)
  are descoped unless trivial — record the decision either way. Compat
  extras (NativeAuthError, NestedAppAuthError, createAuth) stay. Track the
  compat size delta; if implementation balloons past the session budget,
  land scenarios + partial impl and file a follow-up.
  Context: PARITY_LOG D4 entry; diff PCA construction in real:
  `node_modules/@azure/msal-browser/dist/app/PublicClientApplication.mjs`
  + `controllers/StandardController.mjs`; per-symbol source under
  `node_modules/@azure/msal-browser/dist/`.

- [ ] **D6** `pending` — Committed doc-sample check (D4 audit, Mission
  bullet 4). NEW `npm run docs:check` (e.g. `test/docs/check.mjs`):
  extract fenced ts/tsx code samples from `README.md`, `docs/README.md`,
  `docs/UPGRADING.md`, `docs/ALACARTE.md` and strict-tsc them against the
  packages' built dist types (same approach D2 ran as session scratch —
  re-derive it; D2's log entry notes the two sample bugs it caught, use
  them as the check's self-test by temporarily breaking a sample). Samples
  that are intentionally partial can opt out via an HTML comment marker —
  keep the marker count low and justified. Wire it into the SOP validate
  list (cheap, docs-affecting tasks only). Zero package-code changes
  expected.
  Context: the four doc files; PARITY_LOG D2 entry (what was checked and
  how); `test/packaging/check.mjs` (existing strict-tsc consumer pattern
  to crib).

- [ ] **D7** `pending` — Close-out re-audit + final summary. Re-run the D4
  skeptic pass over the four Mission bullets now that D5/D6 landed
  (re-run all gates fresh; re-run the export/method/SSR diffs from the D4
  entry; spot-check docs:check). File any new gap as a task before this
  one; when clean, write the final summary for the user and do NOT
  reset-context.
  Context: PARITY_LOG D4 entry (the exact diffs/checks to repeat), this
  file's Mission section.
