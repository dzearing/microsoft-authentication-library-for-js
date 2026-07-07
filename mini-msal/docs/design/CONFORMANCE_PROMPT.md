# Task: MSAL behavior-conformance suite + mini-msal gap report

## Goal

Build a comprehensive black-box conformance test suite that captures the
observable behavior of **real @azure/msal-browser 5.16.0 + @azure/msal-react
5.5.1** (do NOT modify these packages — they are ground truth), then run the
identical suite against **mini-msal** (`src/mini-msal/`) and produce a gap
report. Do not fix mini-msal gaps in this pass — report only, with each gap
classified. Broker/WAM, NAA, and telemetry are explicitly in scope even though
mini doesn't implement them: the suite documents what "full parity" means.

## Where you are

Everything lives in `bundle-size-test/` at the repo root (npm project, deps
installed). Existing assets you should reuse, not rebuild:

- `src/mini-msal/core.ts` + `react.tsx` — the mini implementation under test.
  It already implements: OIDC discovery, redirect/popup/ssoSilent (auth-code +
  PKCE S256), acquireTokenSilent (cache → refresh token → prompt=none iframe),
  multi-account + active account, per-account logout (redirect/popup),
  EventMessage callbacks, InteractionRequired/Browser error classification,
  iframe/popup reload guards (`block_iframe_reload`, `block_nested_popups`,
  `redirect_in_iframe`), per-request redirectUri/prompt/loginHint, and — key —
  it reads/writes real MSAL's exact v5 cache schema ("msal.3|…" entities,
  token.keys/account.keys indexes, active-account-filters), verified
  bidirectionally.
- `mock-idp.mjs` — HTTPS mock OIDC IdP on :4599 (self-signed cert in `.cert/`,
  regenerate with openssl if missing). Two identities (ada/grace), honors
  prompt=none (login_required when session cold) and prompt=select_account,
  login_hint routing, per-user codes/RTs/client_info, `/reset` endpoint for
  test isolation. Extend it as needed (e.g. error injection, claims,
  Retry-After throttling responses) — it's test infra, fair game.
- `serve.mjs` — static server for `dist/` on :4173 (long keep-alive on both
  servers is deliberate: node's 5s default caused connection-reset flakes).
- `e2e.mjs` — existing 25-check E2E driving BOTH stacks through login /
  multi-account / interop / cancelled-login via playwright-core + system
  Chrome (`chromium.launch({ channel: "chrome", headless: true })`,
  `newContext({ ignoreHTTPSErrors: true })`,
  `process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0"` for node→IdP fetches).
  It must still pass 25/25 when you're done.
- `rspack.config.mjs` + `sync-variants.mjs` + `postbuild.mjs` — build. `npm
  run build` regenerates everything; mock app pages are `dist/mini-mock-app/`
  and `dist/real-mock-app/` (same generated app source, different import
  target). The apps expose the client instance as `globalThis.__msal`.

Hard-won gotchas (respect these or you'll rediscover them slowly):
- Real MSAL requires `https` authorities and `jwks_uri` in discovery metadata.
- v5 popups/silent-iframes deliver responses via the **redirect-bridge**
  (`@azure/msal-browser/redirect-bridge`, BroadcastChannel), not URL polling:
  popup redirect pages must run the bridge for real MSAL (see
  `dist/*/popup.html` written by postbuild). mini polls the window URL and
  needs a blank page instead.
- Real MSAL blocks auth APIs when `window.name` starts with `"msal."` or when
  an auth response hash is present in an iframe.
- Real v5 emits some LOGIN_SUCCESS events with a payload lacking `account`.
- Reset the mock IdP (`https://localhost:4599/reset`) between scenarios —
  a warm IdP session makes silent SSO succeed and changes outcomes.

## Task A — conformance suite against real MSAL

Create `bundle-size-test/conformance/` with a runner (playwright-core, reuse
the e2e patterns) that executes **scenarios against a live page running real
msal-browser** and records observable behavior only (public API returns/throws,
error types/codes, events emitted in order, sessionStorage/localStorage writes,
network requests to the IdP, navigations, headers on token requests). Each
scenario produces a structured result object; store the captured real-MSAL
results as JSON snapshots (`conformance/snapshots/real/*.json`) so the mini run
can be diffed mechanically. Build a dedicated test page (a thin harness
exposing `__msal` — you may build new rspack entries; do not reuse the demo app
if a bare harness is cleaner).

Cover at minimum, as separately-runnable areas:

1. **Core flows**: loginRedirect/handleRedirectPromise, loginPopup,
   acquireTokenPopup, ssoSilent (success + login_required), acquireTokenRedirect;
   state validation, PKCE params, response_mode, hash cleanup after processing.
2. **Silent acquisition**: cache hit (fromCache, expiresOn semantics),
   expiry-window refresh, refresh-token grant, iframe fallback, every
   `CacheLookupPolicy` value, concurrent duplicate requests (does real dedupe
   parallel identical acquireTokenSilent calls? capture it), forceRefresh.
3. **Accounts & cache**: multi-account, getAccount filters, active account
   persistence, logout variants incl. per-account, cache entity shapes/keys
   (already have a reference dump in this repo's history — re-capture), cache
   with `cacheLocation: "localStorage"` incl. cross-tab account events.
4. **Errors & guards**: full error taxonomy for each failure (class,
   errorCode) — cancelled login, popup blocked, popup closed by user, iframe
   timeout, interaction_required family, uninitialized client, redirect in
   iframe, nested popups, interaction_in_progress (two concurrent interactive
   calls).
5. **Broker / WAM (platform auth)**: enable `system.allowPlatformBroker`.
   The DOM handler talks to the browser extension via CustomEvents/postMessage
   (see `dist/broker/nativeBroker/PlatformAuthDOMHandler.mjs` and
   `PlatformAuthExtensionHandler.mjs` in the real package for the handshake
   shape). Write a **fake extension** in the test page that answers the
   handshake, and capture: how msal probes for the broker, what request it
   sends, how responses/errors map to results, and fallback behavior when no
   broker answers. Mini is expected to fail all of these — the point is
   recording real's contract.
6. **NAA (nested app auth)**: `createNestablePublicClientApplication` with a
   mocked bridge host (postMessage protocol — see `dist/naa/BridgeProxy.mjs`).
   Capture initialization handshake, token request/response mapping, error
   mapping.
7. **Telemetry**: `addPerformanceCallback` event shapes for the main flows
   (names, success/failure, durations present); server telemetry: capture
   `x-client-SKU`, `x-client-VER`, `client-request-id`, and the
   `x-client-last-*`/telemetry headers on token-endpoint requests (extend
   mock-idp to record request headers and expose them for assertions);
   correlationId propagation from request into events and headers.
8. **Requests passthrough**: prompt, loginHint, sid, domainHint,
   extraQueryParameters, claims (+ CAE `clientCapabilities: ["cp1"]` → `xms_cc`
   in claims/params), custom `state`, per-request redirectUri, `authority`
   override per request, scopes normalization (openid/profile/offline_access).
9. **Resilience**: token endpoint 429/Retry-After throttling behavior (extend
   mock-idp to inject), 5xx retry behavior, network failure error mapping,
   proactive refresh (`refreshOn`/`shouldRefresh` if observable).
10. **Init & misc**: initialize() requirements (API call before init throws
    uninitialized), getConfiguration, logger callback wiring, bfcache/pageshow
    handling if observable, `msal.version` key, redirect-bridge page contract.

Scenario = one JSON-serializable spec: setup (IdP state, storage seed, config),
actions (via `page.evaluate` against `__msal`), captured observations. Keep
observations normalized (sort keys, strip timestamps/GUIDs into placeholders
like `<guid>`) so diffs are stable.

## Task B — run against mini, produce the gap report

Same runner, same scenarios, `target=mini`. Diff against the real snapshots.
Produce `conformance/GAP_REPORT.md`:

- Summary table: area | scenarios | pass | behavioral-diff | missing-feature.
- Per-gap entry: scenario, real behavior (concise), mini behavior, class:
  **missing-feature** (broker, NAA, telemetry, localStorage, throttling…),
  **behavioral-diff** (works but differs — e.g. error code strings, event
  ordering, dedupe), or **bug** (mini misbehaves on something it claims to
  support), plus an estimated cost-to-fix in KB for missing features where
  reasonable.
- Do NOT modify `src/mini-msal/` in this pass, even for obvious bugs — the
  report is the deliverable. Exception: if a mini crash prevents *other*
  scenarios from running, you may add the smallest guard needed and must note
  it in the report.

## Ground rules

- Never edit `node_modules/@azure/*` (reading its source to understand
  protocols is encouraged and expected, especially for broker/NAA mocking).
- Deterministic: `/reset` the IdP between scenarios; fresh browser context per
  scenario; no reliance on wall-clock ordering. Suite must pass twice in a row.
- `npm run e2e` must still pass 25/25 at the end.
- Wire up `npm run conformance` (capture+diff both stacks) and
  `npm run conformance:real` / `npm run conformance:mini`.
- Finish with: counts (scenarios, real-captured, mini pass/diff/missing), the
  GAP_REPORT.md, and your top-5 most interesting findings.
