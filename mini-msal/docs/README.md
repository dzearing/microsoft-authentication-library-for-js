# mini-msal

**A prototype that asks: how small can MSAL for the browser be if we rebuild
it from scratch — while behaving the same way?**

mini-msal is a from-scratch reimplementation of the parts of
`@azure/msal-browser` and `@azure/msal-react` that a typical single-page app
actually uses. It signs users in with the same protocol (auth code + PKCE),
stores tokens in the same cache format, exposes the same API shapes — and
ships in about **20 KB minified / 7 KB gzip**, where the real stack costs
about **249 KB minified / 65 KB gzip**.

It is a **prototype**, not a supported product. Its job is to prove, with
evidence, what a minimal MSAL could look like and exactly where it still
differs from the real thing.

## The goal: parity, but tiny

Two constraints drive everything here:

1. **Behave like real MSAL.** Same API surface, same wire protocol, same
   error codes, same cache entries on disk. "Looks similar" isn't good
   enough — we measure parity with a 75-scenario black-box conformance suite
   that runs identical scenarios against both libraries and diffs every
   observable detail (see [GAP_REPORT.md](./GAP_REPORT.md)).
2. **Stay small.** The real `@azure/msal-browser` costs ~220 KB minified even
   after aggressive tree shaking, because its main controller statically pulls
   in every feature (broker, telemetry, IndexedDB, …) whether you use it or
   not. mini-msal only includes what the app calls.

Where those goals conflict, we document the difference rather than silently
diverge — that's what the conformance suite and gap report are for.

## How you consume it

Two packages, mirroring the real ones:

| Package | Mirrors | What's in it |
|---|---|---|
| `@mini-msal/browser` | `@azure/msal-browser` | `PublicClientApplication`, auth flows, cache, errors, events |
| `@mini-msal/react` | `@azure/msal-react` | `MsalProvider`, hooks, templates, `withMsal` |

You import them directly. The API mirrors MSAL's, so for the supported
surface, code looks exactly like MSAL code — only the import specifier
changes:

```tsx
import { PublicClientApplication } from "@mini-msal/browser";
import { MsalProvider, useMsal, MsalAuthenticationTemplate } from "@mini-msal/react";

const pca = new PublicClientApplication({
    auth: {
        clientId: "your-client-id",
        authority: "https://login.microsoftonline.com/your-tenant",
        redirectUri: window.location.origin,
    },
});
await pca.initialize();

// ...the calls you already know:
await pca.loginPopup({ scopes: ["User.Read"] });
const result = await pca.acquireTokenSilent({ scopes: ["User.Read"], account });
```

Because mini-msal reads and writes **real MSAL's exact v5 cache schema**
(`msal.3` entities, token-key indexes, active-account filters — verified in
both directions by the e2e suite), you can even migrate a signed-in app
between the two libraries without users losing their sessions, as long as the
cache stays in `sessionStorage`.

Two practical differences from real MSAL to know up front:

- **No redirect-bridge page.** Real MSAL v5 requires you to serve a special
  bridge page at every popup/iframe redirect URI. mini-msal polls the popup
  URL instead, so a **blank page** works. (Register a blank page like
  `/blank.html` as a redirect URI and pass it as `redirectUri` on popup and
  silent requests.)
- **One storage mode.** Only `cacheLocation: "sessionStorage"` is
  implemented today.

## Architecture

The whole library is intentionally small enough to read in one sitting:

```
packages/
├── browser/src/index.ts    ~900 lines: everything protocol + cache
└── react/src/index.tsx     ~250 lines: React bindings over the browser package
```

**`@mini-msal/browser`** is one class (`PublicClientApplication`) plus small
helpers. The flow logic is deliberately simple:

- **Interactive login** builds an authorize URL (auth code + PKCE S256,
  `response_mode=fragment`), then either navigates the page (redirect flow)
  or opens a popup and polls its URL until the code arrives (popup flow).
  `handleRedirectPromise()` finishes the redirect flow after the round trip.
- **Silent acquisition** (`acquireTokenSilent`) walks a three-step ladder:
  return a cached, unexpired access token → otherwise redeem the refresh
  token → otherwise retry sign-in invisibly in a hidden iframe with
  `prompt=none`.
- **The cache is real MSAL's cache.** Every account, id/access/refresh token
  is written under the same keys with the same JSON shape real MSAL v5 uses.
  This is the load-bearing design decision: it makes parity externally
  checkable and migration free.
- **Errors** are classified the way MSAL classifies them
  (`InteractionRequiredAuthError` vs `BrowserAuthError` vs generic
  `AuthError`) with MSAL's error-code strings, so `catch` blocks written for
  MSAL keep working.
- **Events** use MSAL's `EventMessage` shape and event-type strings
  (`msal:loginSuccess`, …) via `addEventCallback`/`removeEventCallback`.

**`@mini-msal/react`** mirrors msal-react's component model: `MsalProvider`
holds the instance and re-renders on auth events; `AuthenticatedTemplate`,
`UnauthenticatedTemplate`, and `MsalAuthenticationTemplate` gate UI;
`useMsal`, `useAccount`, `useIsAuthenticated`, `useMsalAuthentication`, and
`withMsal` match the real hooks' contracts (including the stable-account
reference behavior that prevents effect loops).

## What's implemented, and what isn't

Implemented and verified end-to-end:

- OIDC discovery; redirect, popup, and `ssoSilent` login (auth code + PKCE)
- `acquireTokenSilent` (cache → refresh token → `prompt=none` iframe),
  `acquireTokenPopup`, `acquireTokenRedirect`
- Multi-account cache, active account, `getAccount*` lookups
- Logout (redirect + popup, per-account or all)
- Event callbacks, MSAL-compatible error classification
- Iframe/popup environment guards; per-request `redirectUri`, `prompt`,
  `loginHint`; the React layer

Not implemented (by design, for now): platform broker (WAM), nested app auth
(Teams/Office hosts), telemetry/perf events, logger, `localStorage` +
cross-tab sync, 429 throttling cache, proactive refresh (`refresh_in`),
claims/CAE, custom state, `sid`/`domainHint`/`extraQueryParameters`,
per-request authority override, B2C/CIAM/ADFS authorities, PoP tokens.

The precise, evidence-backed list — including **8 known bugs** in what mini
does claim to support, each with a cost-to-fix estimate — lives in
[GAP_REPORT.md](./GAP_REPORT.md). Headline estimate: **~4 KB** of additions
close every bug and most app-visible behavioral differences; full parity
including broker + NAA is ~27 KB more (still ~5× smaller than real).

## How we know it behaves the same

Three independent test layers, all runnable locally with no Azure setup
(a mock OIDC identity provider runs on localhost):

1. **E2E parity suite** (`npm run e2e`, 25 checks): drives a real browser
   through the identical full app — login, popups, multi-account, logout —
   built once against real MSAL and once against mini-msal, including the
   cache-interop proof in both directions.
2. **Conformance suite** (`npm run conformance`, 75 scenarios): captures real
   MSAL's observable behavior — API results, error types and codes, events in
   order, storage writes, every request the IdP sees — as normalized JSON
   snapshots, then replays the same scenarios against mini-msal and diffs
   mechanically. Deterministic across runs. Even the platform-broker and
   nested-app-auth protocols are exercised, using in-page fakes built from
   the real wire protocols (see [design/broker-protocol.md](./design/broker-protocol.md)
   and [design/naa-protocol.md](./design/naa-protocol.md)).
3. **Bundle-size harness** (`npm run measure`): the number the whole project
   exists to keep honest (see
   [design/bundle-size-experiment.md](./design/bundle-size-experiment.md)).

## Repository map

```
mini-msal/
├── docs/                 you are here (README, GAP_REPORT, design/)
│   └── design/           deep-dives: conformance suite internals, wire
│                         protocols, bundle-size experiment, original spec
├── packages/
│   ├── browser/          @mini-msal/browser  (the core library)
│   └── react/            @mini-msal/react    (React bindings)
└── test/
    ├── infra/            mock OIDC IdP, static server, build config
    ├── apps/             the shared demo/e2e app + conformance harness pages
    ├── e2e/              25-check dual-stack E2E
    ├── conformance/      75-scenario conformance runner + snapshots
    └── bundle-size/      size measurement + source-map attribution
```

## Running it

```sh
npm install
npm run build              # generate app variants + build all bundles
npm run e2e                # 25-check E2E against the local mock IdP
npm run conformance        # capture real → replay mini → regenerate GAP_REPORT.md
npm run conformance:check  # determinism check (real vs its own snapshots)
npm run measure            # bundle sizes (min / gzip / brotli)
```

To try it against a real Entra ID tenant, see
[design/bundle-size-experiment.md](./design/bundle-size-experiment.md#testing-with-real-entra-id).

## Status and next steps

Current conformance standing (75 scenarios): **2 identical · 33 behavioral
differences · 32 missing features · 8 bugs.** Most behavioral differences are
cosmetic (event coverage, error `.name`, result metadata); the bugs are small
and enumerated with fixes in the gap report. The suggested order of work is
in [design/conformance-suite-notes.md](./design/conformance-suite-notes.md#if-asked-to-fix-mini-next-the-8-bugs-cheapest-first).
