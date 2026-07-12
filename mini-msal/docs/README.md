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

Two styles, one library (pay-to-play architecture):

| Style | Import | What you get |
|---|---|---|
| **Drop-in (compat)** | `@mini-msal/compat` | The classic `PublicClientApplication` with **every** feature composed — swap the import from `@azure/msal-browser` and nothing else changes |
| **À la carte** | `@mini-msal/browser` + feature subpaths | `createClient(config, [features])` — you bundle only the features you compose |
| React bindings | `@mini-msal/react` | `MsalProvider`, hooks, templates, `withMsal` — works over either style |

Drop-in:

```ts
import { PublicClientApplication } from "@mini-msal/compat";

const pca = new PublicClientApplication({
    auth: { clientId: "...", authority: "...", redirectUri: window.location.origin },
});
await pca.initialize();
// ...the calls you already know:
await pca.loginPopup({ scopes: ["User.Read"] });
const result = await pca.acquireTokenSilent({ scopes: ["User.Read"], account });
```

À la carte — a redirect-only SPA that doesn't want to pay for popup, broker,
localStorage encryption, or telemetry bytes:

```ts
import { createClient } from "@mini-msal/browser";

const pca = createClient({ auth: { clientId: "..." } }); // core only
await pca.initialize();
await pca.loginRedirect({ scopes: ["User.Read"] });
```

...and features compose in as plain functions when you need them:

```ts
import { createClient } from "@mini-msal/browser";
import { popup } from "@mini-msal/browser/popup";
import { telemetry } from "@mini-msal/browser/telemetry";

const pca = createClient(config, [popup, telemetry]);
```

Feature modules today: `./popup`, `./broker` (platform broker / WAM),
`./local-storage` (encrypted localStorage + cross-tab sync), `./telemetry`
(performance events), `./redirect-bridge` (the popup/iframe bridge-page
bundle, 0.6 KB vs real's 6.5 KB).

Because mini-msal reads and writes **real MSAL's exact v5 cache schema**
(`msal.3` entities, token-key indexes, active-account filters — verified in
both directions by the e2e suite), you can even migrate a signed-in app
between the two libraries without users losing their sessions, as long as the
cache stays in `sessionStorage`.

Like real MSAL v5, popup/iframe responses complete via a redirect-bridge
page served at the popup redirect URI — mini ships a 0.6 KB bridge bundle
(`@mini-msal/browser/redirect-bridge`). Both `sessionStorage` (default) and
`localStorage` (encrypted entities + cross-tab events, matching real's
observable shape) cache locations are implemented.

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

Current conformance standing (75 scenarios): **68 identical · 7 missing
features · 0 behavioral diffs · 0 bugs.** The work is driven task-by-task
from [design/PARITY_STATE.md](./design/PARITY_STATE.md); remaining:
platform-broker extension transport (C8), nested app auth (C9), final sweep +
published size matrix (C10), then an à-la-carte DX / docs / examples phase
(D-series).

Size today (min / gzip): à-la-carte core ~18.8 / 7.0 KB · `@mini-msal/compat`
with every feature ~37.0 / 12.6 KB · real msal stack ~249 / 65 KB. Every
parity task records its size cost in the PARITY_STATE progress log — full
parity is being bought at roughly 1/7th of real's size, and consumers who
compose less pay less.
