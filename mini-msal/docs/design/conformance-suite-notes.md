# Conformance suite — implementation notes & session handoff

Complete state dump of the conformance-suite build (2026-07-06/07). Everything
needed to continue, re-run, or extend the suite without prior context.

## Status: COMPLETE ✅

- 75 scenarios, 10 areas, all captured against real @azure/msal-browser 5.16.0
- Determinism verified: `npm run conformance:check` → 75/75 pass (two runs in a row)
- Mini run: 2 pass / 33 behavioral-diff / 32 missing-feature / 8 bug (see GAP_REPORT.md)
- `npm run e2e` still passes 25/25
- `packages/browser/src/ + packages/react/src/` was NOT modified (per task rules — report only)

## How everything fits together

```
test/conformance/
  run.mjs              runner: --target=real|mini [--check] [--area=a,b] [--grep=s]
  lib.mjs              scenario helpers + normalization + diff engine
  scenarios/01-10*.mjs 75 scenarios (each exports {area, scenarios:[{id, note, run(ctx)}]})
  snapshots/real/*.json normalized real-MSAL observations (one per scenario id)
  results/{real,mini,real-check}.json   per-run outcomes (+diff paths for mini)
  results/diff-overview.txt             compact human-readable diff dump
  report.mjs           docs/GAP_REPORT.md generator: merges results/mini.json with the
                       authored CLASSIFICATIONS map (class/real/mini/cost per id).
                       Edit classifications THERE, then `node test/conformance/report.mjs`.
  docs/design/broker-protocol.md, docs/design/naa-protocol.md   wire-protocol specs (from
                       reading real's dist source) used to build the in-page fakes
  ../../docs/GAP_REPORT.md   the deliverable (written by report.mjs)
test/apps/harness/       harness pages (setup.ts + harness-{real,mini}.ts) built by
                       rspack into dist/conformance-{real,mini}/ (postbuild writes
                       index.html + popup.html + popup2.html per harness)
```

Harness page exposes: `__lib` (module namespace), `__create(config, opts)`
(opts: `noInit`, `captureLogs`, `perfClient`), `__msal`, `__cap`
{events, perf, logs}, `__serializeError`. It does NOTHING on load — scenarios
drive via `page.evaluate`. Scenario helpers in lib.mjs: `gotoHarness`, `create`,
`tryEval`/`tryResult` (uniform {ok}/{err} capture), `fire`/`pending`
(fire-and-forget for concurrent-interaction tests), `loginViaRedirect`,
`loginViaPopup`, `patchAccessTokens`, `removeCreds`, `storageDump`, `capture`,
`idp.{reset,requests,inject,config,session}`, `digestIdpLog`.

Normalization (lib.mjs `normalize`): volatile keys → `<dynamic>` (state, nonce,
code_challenge/verifier, client-request-id, correlationId, cachedAt/expiresOn/
refreshOn, sendTime, `data` [encrypted localStorage ciphertext], header noise);
GUIDs → `<guid>`, JWTs → `<jwt>`, 43-char b64url → `<b64-43>`, epoch numbers →
`<epoch-s>/<epoch-ms>`; `conformance-(real|mini)` → `conformance-app` so both
targets share snapshots; recursive key sort. Diffs are path-level (lib.mjs `diff`).

Exit codes: real capture fails on scenario errors; real --check fails on any
diff (nondeterminism); mini always exits 0 (diffs ARE the deliverable).

## Mock IdP extensions (test/infra/mock-idp.mjs)

All cleared by `/reset`:
- `GET /requests` — ordered log of discovery/authorize/token/logout requests
  (query, POST body, headers). Token requests also log their query string.
- `GET /inject?endpoint=token|authorize&status=429&retryAfter=2&error=x&errorDescription=y&suberror=z&drop=1&delay=6000&count=1`
  — queued one-shot injections. `delay` defers the normal response (works for
  both endpoints); `drop` destroys the socket; delay-only injections continue
  the normal flow after sleeping.
- `GET /config?expires_in=7200&refresh_in=1&spa_accountid=native-acc-1` —
  merged into every token response until reset (numbers auto-coerced).
- `GET /session?active=1&user=0` — set IdP SSO session state directly.
- `/logout` now echoes `state` back to post_logout_redirect_uri (OIDC
  RP-initiated logout) — REQUIRED for real v5 logoutPopup to ever resolve.

## Hard-won gotchas (cost real debugging time — do not rediscover)

1. **Everything popup/iframe in real v5.16 completes via the redirect-bridge**
   (BroadcastChannel keyed by the libraryState id inside `state`). There is NO
   URL polling and NO popup.closed polling (`waitForBridgeResponse` in
   utils/BrowserUtils.mjs). Consequences:
   - Silent/popup requests in scenarios must pass `redirectUri: ctx.popupUrl`
     (bridge page for real, blank for mini).
   - Logout popups need `postLogoutRedirectUri: ctx.popupUrl` AND the IdP
     state echo, else they hang until timeout.
   - User closing the popup = hang until `popupBridgeTimeout` (default 60s),
     then `timed_out`/`redirect_bridge_timeout` — NOT `user_cancelled`.
2. **Timeout config keys in 5.16** are `popupBridgeTimeout` (60s default) and
   `iframeBridgeTimeout` (10s default) — the old `windowHashTimeout`/
   `iframeHashTimeout` names do nothing for bridge waits. Set them small in
   failure scenarios to keep runs fast.
3. **Broker gating**: `canUsePlatformBroker` requires a cached
   `nativeAccountId` for EVERY brokered call including interactive. First-time
   login never uses the broker. Two ways in: token response `spa_accountid`
   (hybrid; ResponseHandler stores it as nativeAccountId) or
   `acquireTokenByCode({nativeAccountId})` (the designed broker entry — used
   by the broker scenarios). The fake must echo `request.accountId` as
   `response.account.id` or real throws `userSwitch`.
   DOM path requires BOTH `system.allowPlatformBroker: true` AND
   `experimental.allowPlatformBrokerWithDOM: true`.
4. **Extension-path fake**: capture-phase window "message" listener +
   `event.stopImmediatePropagation()` on the Handshake message (else real's
   own bubble listener sees its bounce and concludes nativeExtensionNotInstalled),
   reply on `event.ports[0]` with HandshakeResponse echoing responseId, then
   answer GetToken on that port. Full field spec: docs/design/broker-protocol.md.
5. **NAA fake**: `window.nestedAppAuthBridge = {addEventListener, postMessage}`,
   responses are JSON STRINGS passed directly to the stored callback; echo
   `requestId`; `initContext` required; token responses need decodable JWT
   id_token + `account.environment`. Full spec: docs/design/naa-protocol.md.
6. **Perf events need opt-in**: default telemetry client is StubPerformanceClient.
   Harness `opts.perfClient: true` constructs `new lib.BrowserPerformanceClient(config)`.
7. **Server telemetry is in the token BODY** (x-client-SKU/VER,
   x-client-current/last-telemetry, x-ms-lib-capability) — not HTTP headers.
   client-request-id rides the token endpoint QUERY string.
8. **Real v5 encrypts localStorage entities** ({id, nonce, data: ciphertext},
   per-session key) — hence `data` in the volatile-keys list, and localStorage
   cache interop between stacks is architecturally impossible.
9. **Playwright popup-event timing trap**: for real (opens about:blank then
   navigates) `context.waitForEvent("page")` fires immediately; for mini
   (window.open directly at the authorize URL) it fires only after navigation
   commits — so with a delayed IdP the event arrives AFTER login completed.
   For close-the-popup tests, wrap `window.open` in-page to capture the handle
   and close it at a fixed delay after the CALL (see errors.popup-closed-by-user).
10. **Same-URL navigations** (logoutRedirect roundtrip lands back on the exact
    page URL): `waitForURL` matches instantly against the CURRENT url — use
    `page.waitForEvent("framenavigated")` armed BEFORE triggering.
11. Server 302 redirect chains never commit intermediate URLs — you cannot
    observe the IdP /logout hop from `page.url()`; assert via `idp.requests()`.
12. mini's `PublicClientApplication` lacks several APIs (acquireTokenByCode,
    getConfiguration, addPerformanceCallback…) — scenarios use tryEval/tryResult
    everywhere so mini's TypeErrors are captured as observations, never crashes.
13. The runner's `B64_43_RE` normalizes ANY 43-char `[\w-]` token — including
    OBJECT KEY NAMES. Three real perf-event ext keys are exactly 43 chars
    (`generateCodeChallengeFromVerifierDurationMs`,
    `silentFlowClientAcquireCachedTokenCallCount`,
    `silentHandlerMonitorIframeForHashDurationMs`) and show up as `<b64-43>`
    in snapshots; two such keys in one object COLLAPSE to a single key. Do
    not conclude a key is missing because you only see `<b64-43>` (this
    misled C11's first read). Perf-event raw shapes: `__cap.perfRaw` in the
    harness (undefined-valued keys preserved as null).

## Key real-MSAL behavioral facts captured (basis for GAP_REPORT)

- Silent: real dedupes concurrent identical acquireTokenSilent (1 network call);
  CacheLookupPolicy fully enforced (AccessToken+expired → token_refresh_required;
  RefreshToken ignores valid AT; RefreshTokenAndNetwork falls back to iframe;
  Skip goes straight to iframe); forceRefresh honored; refresh_in → refreshOn.
- 429: ServerError + throttle cache entry; immediate retry re-throws WITHOUT
  network. 5xx: no retry. Dropped socket: browser fetch retries transparently.
- Forged state on redirect response: real resolves handleRedirectPromise NULL
  silently; clean load emits initialize events only (no handleRedirect events).
- Real never auto-sets active account; getAccount({}) → null; getAllAccounts
  accepts a filter.
- msal.* window-name popup block NO LONGER EXISTS in 5.16 (loginPopup succeeds
  inside an msal-named window).
- interaction_in_progress enforced via temp cache; second popup rejected, first
  completes.
- Errors before initialize(): uniform uninitialized_public_client_application.
- prompt only forwarded for login/consent/select_account/none (broker gate also
  filters prompts); custom state is embedded as `libState|userState` and
  returned on result.state; claims + clientCapabilities merge xms_cc.

## If asked to fix mini next (the 8 bugs, cheapest first)

1. Set `.name` on AuthError subclasses (~1 line each).
2. Uninitialized guard (throw uninitialized_public_client_application).
3. getAccount({}) → null; getAllAccounts(filter) honor filter.
4. Scope dedupe + no trailing space on empty scopes.
5. forceRefresh + CacheLookupPolicy semantics in acquireTokenSilent.
6. In-flight dedupe map keyed by (scopes|account|authority).
7. Interaction lock + unique popup window names (fixes the clobbering).
8. Don't auto-set active account on first login (match real).
GAP_REPORT.md has per-item KB estimates; systemic result-shape/event/nonce gaps
are listed in its "Systemic gaps" section.

## Re-run cheatsheet

```
npm run build              # required after touching test/apps/harness or packages/*/src
npm run conformance        # real capture → mini run → report
npm run conformance:check  # determinism (must stay 75/75)
node test/conformance/run.mjs --target=mini --grep=silent.force-refresh   # one scenario
DEBUG=1 node test/conformance/run.mjs ...   # server logs + console passthrough
npm run e2e                # must stay 25/25
```
