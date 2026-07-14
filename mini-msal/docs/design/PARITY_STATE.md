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

**Status (2026-07-13)**: Phases A0–C18 COMPLETE — conformance 109/109,
e2e 25/25, GAP_REPORT all-pass. Size matrix: compat-no-react 55.0 KB min
(real: 220.5), compat+react stack 64.9 KB (real: 248.8), core-only
28.0 KB. Remaining: C-gaps C19–C21 (post-audit parity gaps beyond the
suite) then D-series (à-la-carte DX/docs/examples).
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
   - `npm run measure` → record mini-msal-stack size in the task row
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
- [ ] **C19** `pending` — Config knobs + logout params. Findings:
  `allow-redirect-in-iframe` (honor the flag inside iframes),
  `token-renewal-offset-seconds` (expiry buffer configurable, not
  hardcoded 300s), `logout-hint-param` (logoutHint / id_token-derived
  logout_hint / idTokenHint / eQP on end_session URL),
  `server-telemetry-enabled` (LOW: x-client-current/last-telemetry
  populated + server-telemetry cache entry when enabled). Scenarios per
  findings. Context: audit json; `packages/browser/src/index.ts`.
- [ ] **C20** `pending` — Perf-event emission semantics (C11's sibling —
  do AFTER C11). Findings: `handle-redirect-perf-event` (root
  acquireTokenRedirect event from handleRedirectPromise; NOTE
  verdict.corrections: clean loads never start the measurement),
  `failure-event-correlation-id` (failure event cid = library-generated
  request cid = error.correlationId = wire client-request-id),
  `duplicate-perf-callback-dedupe` (dedupe registrations by callback
  source text), `preflight-failure-no-perf-event` (preflight failures
  emit NOTHING), `init-perf-event-once` (initializeClientApplication at
  most once), `performance-marks-session-flag` (LOW:
  msal.browser.performance.enabled='1' → performance.mark/measure
  timeline entries). Scenarios per findings. Context: audit json;
  `packages/browser/src/telemetry.ts`, scenarios 08-telemetry.mjs.
- [ ] **C21** `pending` — authenticationScheme "pop"/"ssh" (PoP binding).
  Finding: `authentication-scheme-pop` — real sends req_cnf (JWK
  thumbprint), caches with kid in the key, returns tokenType "pop" +
  SignedHttpRequest; mini silently downgrades to Bearer. This was
  previously listed as a non-goal in docs — the audit confirmed it's an
  observable wire + result difference. Implement OBSERVABLE parity
  (req_cnf param, pop tokenType/cache shape, SHR signing via WebCrypto)
  if it fits a session; if the size cost is disproportionate (>~3 KB
  min on compat), STOP and ask the user whether to descope to a
  documented non-goal instead — that is a user decision. Context: audit
  json; `node_modules/@azure/msal-common/dist/crypto/PopTokenGenerator.mjs`,
  `packages/browser/src/index.ts`.

### Phase D — à-la-carte DX, docs & examples (after C-gaps)

The C-series proves compat correctness; the D-series makes the à-la-carte
story real for consumers. Same SOP applies (validate → update doc → commit +
push → reset). "Consumer" below means someone who has never read this repo.

- [ ] **D1** `pending` — Seam hardening + consumer packaging. (a) Calling a
  non-composed feature's API must fail with a clear, documented
  BrowserAuthError (e.g. `feature_not_configured: loginPopup requires
  composing popup from "@mini-msal/browser/popup"`) — never
  undefined-is-not-a-function; add conformance-style unit checks for every
  feature-owned public API on a core-only client. (b) Verify real-world
  packaging: `npm pack` each package, install into a throwaway consumer app
  (temp dir, file: deps), confirm subpath exports + types resolve and a
  minimal build tree-shakes to the expected size. Fix exports maps as needed.
  Context: `packages/*/package.json`, `packages/browser/src/index.ts`
  (seam surface), scratch dir for the throwaway consumer.
- [ ] **D2** `pending` — Consumer docs (user-required deliverables named
  2026-07-13). Two NEW standalone guides plus the README rewrite:
  (a) `docs/UPGRADING.md` — basic drop-in usage/upgrade doc: migrating from
  `@azure/msal-browser`/`@azure/msal-react` to `@mini-msal/compat` +
  `@mini-msal/react` (change one import; cache carries over — signed-in
  users stay signed in; the redirect-bridge page swap `/popup.html` →
  mini's 0.6 KB bridge; what to verify after switching; known
  non-goals like B2C/PoP).
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
- [ ] **D3** `pending` — Runnable examples with measured sizes. `examples/`
  with one minimal app per profile (core-redirect, core+popup, compat
  drop-in, react), built via the existing rspack infra and added to
  `npm run measure` so each example doubles as a size regression check;
  per-example README states what it composes and its measured size. Wire a
  smoke check (reuse e2e harness) proving each example signs in against the
  mock IdP.
  Context: `test/infra/rspack.config.mjs`, `test/apps/mini-core.ts` (model
  entry), `test/e2e/e2e.mjs` (harness to reuse).
- [ ] **D4** `pending` — Requirements audit vs this Mission. Walk the four
  definition-of-done bullets as a skeptic; for each, cite the evidence
  (conformance counts, size matrix, packaging check, docs/examples). File any
  gap found as a new task before this one is marked done. When no gaps
  remain: final summary for the user, do NOT reset-context.
  Context: this file's Mission section, `PARITY_LOG.md` progress table,
  `docs/README.md`, `examples/`.
