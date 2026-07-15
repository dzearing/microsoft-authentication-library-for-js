# mini-msal

**A prototype that asks: how small can MSAL for the browser be if we rebuild
it from scratch — while behaving the same way?**

mini-msal is a from-scratch reimplementation of `@azure/msal-browser` and
`@azure/msal-react`. It signs users in with the same protocol (auth code +
PKCE), stores tokens in the same cache format, exposes the same API shapes —
and the full drop-in with **every** feature composed ships in about
**62.1 KB minified / 20.1 KB gzip**, where tree-shaken `@azure/msal-browser`
costs about **220.5 KB minified / 55.4 KB gzip**. Compose less and pay less,
down to a 29.7 KB redirect-only floor.

It is a **prototype**, not a supported product. Its job is to prove, with
evidence, what a minimal MSAL could look like and exactly where it still
differs from the real thing.

## The goal: parity, but tiny

Two constraints drive everything here:

1. **Behave like real MSAL.** Same API surface, same wire protocol, same
   error codes, same cache entries on disk. "Looks similar" isn't good
   enough — we measure parity with a 122-scenario black-box conformance suite
   that runs identical scenarios against both libraries and diffs every
   observable detail (see [GAP_REPORT.md](./GAP_REPORT.md) — currently
   **122/122 identical**).
2. **Stay small.** The real `@azure/msal-browser` costs ~220 KB minified even
   after aggressive tree shaking, because its main controller statically pulls
   in every feature (broker, telemetry, IndexedDB, …) whether you use it or
   not. mini-msal only includes what the app composes.

Where those goals conflict, we document the difference rather than silently
diverge — that's what the conformance suite and gap report are for.

## Consuming it

Two styles, one library (pay-to-play architecture), plus optional React
bindings. Two standalone guides cover the details:

- **[UPGRADING.md](./UPGRADING.md)** — migrate an existing
  `@azure/msal-browser` / `@azure/msal-react` app to the drop-in (one import
  swap; signed-in users stay signed in).
- **[ALACARTE.md](./ALACARTE.md)** — step the bundle down profile by profile
  with measured sizes and the exact composition for each.
- **[SIZE.md](./SIZE.md)** — how mini-msal got small: each size-reduction
  technique explained in plain language, with examples.
- **[../examples/](../examples/)** — a runnable minimal app per profile
  (core-redirect, core+popup, compat drop-in, react), each size-tracked in
  `npm run measure` and smoke-tested against the mock IdP
  (`npm run examples:smoke`).

> **Prototype note**: the packages are not published to npm — consume them
> from this repo as workspace dependencies or via `npm pack` (that path is
> CI-verified: `npm run pack:check` builds a throwaway consumer against the
> packed tarballs, type-checks every subpath export under strict TS, and
> tree-shake-gates the bundle sizes).

### Profile 1 — drop-in compat (62.1 KB min)

The classic `PublicClientApplication` with every feature composed. Swap the
import from `@azure/msal-browser` and nothing else changes — config, request
objects, `catch` blocks, and event callbacks all keep working:

```ts
import { PublicClientApplication } from "@mini-msal/compat";

const pca = new PublicClientApplication({
    auth: {
        clientId: "11111111-2222-3333-4444-555555555555",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin,
    },
});
await pca.initialize();
await pca.handleRedirectPromise();
await pca.loginPopup({ scopes: ["User.Read"] });
const account = pca.getAllAccounts()[0];
const result = await pca.acquireTokenSilent({ scopes: ["User.Read"], account });
```

### Profile 2 — core-only redirect SPA (29.7 KB min)

`createClient` with no feature modules: redirect sign-in, the full silent
ladder (cache → refresh token → `prompt=none` iframe), multi-account,
events, sign-out.

```ts
import { createClient, InteractionRequiredAuthError } from "@mini-msal/browser";

const client = createClient({
    auth: { clientId: "11111111-2222-3333-4444-555555555555" },
});
await client.initialize();
const response = await client.handleRedirectPromise();
if (response) client.setActiveAccount(response.account);

const account = client.getActiveAccount() ?? client.getAllAccounts()[0];
if (!account) {
    await client.loginRedirect({ scopes: ["User.Read"] });
} else {
    try {
        const result = await client.acquireTokenSilent({ scopes: ["User.Read"], account });
        console.log(result.accessToken);
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            await client.acquireTokenRedirect({ scopes: ["User.Read"] });
        } else throw e;
    }
}
```

### Profile 3 — core + the features you call (32.6 KB with popup)

Features are plain functions composed into `createClient`; only what you
import ends up in the bundle. Calling a feature-owned API you didn't compose
throws a documented `feature_not_configured` error naming the exact import.

```ts
import { createClient, type AuthClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";

const client = createClient(
    { auth: { clientId: "11111111-2222-3333-4444-555555555555" } },
    [popup]
) as AuthClient & PopupClient;

await client.initialize();
const result = await client.loginPopup({ scopes: ["User.Read"] });
```

### Profile 4 — React bindings (+10.0 KB, over either style)

```tsx
import { PublicClientApplication } from "@mini-msal/compat";
import { MsalProvider, AuthenticatedTemplate, UnauthenticatedTemplate } from "@mini-msal/react";
import { createRoot } from "react-dom/client";

const pca = new PublicClientApplication({ auth: { clientId: "…" } });

createRoot(document.getElementById("root")!).render(
    <MsalProvider instance={pca}>
        <AuthenticatedTemplate>signed in</AuthenticatedTemplate>
        <UnauthenticatedTemplate>signed out</UnauthenticatedTemplate>
    </MsalProvider>
);
```

`MsalProvider`, `useMsal`, `useAccount`, `useIsAuthenticated`,
`useMsalAuthentication`, the three templates, and `withMsal` match
msal-react 5.5.1's contracts (verified by the conformance suite's react
area). The provider accepts a compat instance or an à-la-carte client —
its `instance` prop is typed as msal-react's `IPublicClientApplication`
(`AuthClient & PopupClient`; see [ALACARTE.md](./ALACARTE.md)).

### Feature catalog

Measured single-feature deltas over the 29.7 KB core baseline
(2026-07-14; see [ALACARTE.md](./ALACARTE.md) for methodology and recipes):

| Import path | Adds (min) | What it enables |
|---|---:|---|
| `@mini-msal/browser/popup` | +2.9 KB | `loginPopup`, `acquireTokenPopup`, `logoutPopup` |
| `@mini-msal/browser/pop` | +2.0 KB | `authenticationScheme` "pop"/"ssh-cert" (signed HTTP requests) |
| `@mini-msal/browser/broker` | +6.9 KB | Platform broker (WAM), hybrid `acquireTokenByCode` |
| `@mini-msal/browser/local-storage` | +2.6 KB | `cacheLocation: "localStorage"` — encrypted entities, cross-tab sync, KMSI |
| `@mini-msal/browser/cache-migration` | +3.0 KB | msal.0/1/2 → msal.3 cache migration + retention TTL |
| `@mini-msal/browser/telemetry` | +9.4 KB | `addPerformanceCallback` perf events, performance marks |
| `@mini-msal/browser/naa` | +8.4 KB | Nested app auth for Teams/Office hosts (`createNestableClient`) |
| `@mini-msal/browser/redirect-bridge` | 0.6 KB page | The bridge page bundle (separate page, not app bytes) |
| `@mini-msal/react` | +10.0 KB | Provider, hooks, templates (React itself not counted) |

### The redirect-bridge page

Like real MSAL v5, popup, `ssoSilent`, and silent-iframe responses complete
via a page served at the request's `redirectUri`. mini ships the page script
as a 0.6 KB bundle (vs real's 6.5 KB):

```ts
// bridge.ts — the entire bridge page script
import { broadcastResponseToMainFrame } from "@mini-msal/browser/redirect-bridge";

broadcastResponseToMainFrame().catch(console.error);
```

Pure redirect flows don't need it.

### Cache interop

Because mini-msal reads and writes **real MSAL's exact v5 cache schema**
(`msal.3` entities, token-key indexes, active-account filters — verified in
both directions by the e2e suite), you can migrate a signed-in app between
the two libraries without users losing their sessions. Both `sessionStorage`
(default) and `localStorage` (encrypted entities + cross-tab events, matching
real's observable shape) are supported; older `msal.0/1/2` caches are
migrated on `initialize()` like real 5.x.

## Architecture

The whole library is intentionally small enough to read in one sitting:

```
packages/
├── browser/src/            @mini-msal/browser: createClient core + feature
│                           modules (popup, pop, broker, naa, local-storage,
│                           cache-migration, telemetry, redirect-bridge) as
│                           subpath exports
├── compat/src/             @mini-msal/compat: classic PublicClientApplication
│                           composing every feature — the drop-in
└── react/src/index.tsx     @mini-msal/react: React bindings over either style
```

**`@mini-msal/browser`** is a composable core — `createClient(config,
[features])` — where each feature is a plain function receiving the client's
internal context (closures, no classes, so seam bytes stay near zero and
tree shaking works). The flow logic is deliberately simple:

- **Interactive login** builds an authorize URL (auth code + PKCE S256,
  `response_mode=fragment`), then either navigates the page (redirect flow)
  or opens a popup (popup flow); popup/iframe responses complete via real
  v5's redirect-bridge mechanism (a BroadcastChannel message posted by the
  bridge page at the redirect URI). `handleRedirectPromise()` finishes the
  redirect flow after the round trip.
- **Silent acquisition** (`acquireTokenSilent`) walks a three-step ladder:
  return a cached, unexpired access token → otherwise redeem the refresh
  token → otherwise retry sign-in invisibly in a hidden iframe with
  `prompt=none`.
- **Authorities** resolve lazily per flow like real 5.x: trust validation
  (cloud discovery metadata, `knownAuthorities`, hardcoded cloud aliases,
  CIAM), then endpoint resolution from config → hardcoded templates →
  network discovery. `initialize()` performs zero network calls.
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

Packages ship tsc-built `dist/` (JS + `.d.ts`) to consumers via exports
conditions; the repo's internal builds compile straight from `src/`.

## What's implemented, and what isn't

Implemented and verified end-to-end (all 122 conformance scenarios pass):

- OIDC discovery; redirect, popup, and `ssoSilent` login (auth code + PKCE);
  authority modes (AAD, CIAM, `knownAuthorities`, instance-aware multi-cloud,
  configured/hardcoded/network metadata sources)
- `acquireTokenSilent` (cache → refresh token → `prompt=none` iframe, full
  `CacheLookupPolicy` semantics, in-flight dedupe, forceRefresh,
  `tokenRenewalOffsetSeconds`), `acquireTokenPopup`, `acquireTokenRedirect`,
  `acquireTokenByCode` (hybrid SPA)
- PoP and ssh-cert authentication schemes (SHR signing, IndexedDB keystore)
- Multi-account cache with tenant profiles, active account, `getAccount*`
  lookups, KMSI, schema migration from older majors
- Programmatic cache APIs: `clearCache`, `hydrateCache`, `loadExternalTokens`
- Logout (redirect + popup, per-account or all, `logoutHint`,
  `mainWindowRedirectUri`, `onRedirectNavigate`)
- Event callbacks (real's exact event streams + payload shapes),
  MSAL-compatible error classes/codes, interaction lock, `Logger`/`LogLevel`,
  `NavigationClient` seam, deep-link replay (`navigateToLoginRequestUrl`)
- Platform broker (WAM) — DOM and extension transports; nested app auth
  (Teams/Office hosts); perf events (`addPerformanceCallback`, real's
  per-flow event shapes and emission semantics, performance marks); server
  telemetry headers; `localStorage` (encrypted entities) + cross-tab sync;
  429 throttling cache; `refresh_in`/`refreshOn`; claims/CAE; custom state;
  `sid`/`domainHint`/`extraQueryParameters`; per-request authority override;
  the React layer

Not implemented (by design, for now): B2C/ADFS authorities, and anything
else the 122 scenarios don't observe (e.g. real's native in-memory broker
token cache — mini re-asks the broker each call).

The evidence lives in [GAP_REPORT.md](./GAP_REPORT.md): **122/122 scenarios
identical — 0 behavioral diffs, 0 bugs.**

## How we know it behaves the same

Independent test layers, all runnable locally with no Azure setup
(a mock OIDC identity provider runs on localhost):

1. **E2E parity suite** (`npm run e2e`, 25 checks): drives a real browser
   through the identical full app — login, popups, multi-account, logout —
   built once against real MSAL and once against mini-msal, including the
   cache-interop proof in both directions.
2. **Conformance suite** (`npm run conformance`, 122 scenarios): captures real
   MSAL's observable behavior — API results, error types and codes, events in
   order, storage writes, every request the IdP sees — as normalized JSON
   snapshots, then replays the same scenarios against mini-msal and diffs
   mechanically. Deterministic across runs. Even the platform-broker and
   nested-app-auth protocols are exercised, using in-page fakes built from
   the real wire protocols (see [design/broker-protocol.md](./design/broker-protocol.md)
   and [design/naa-protocol.md](./design/naa-protocol.md)).
3. **Seam checks** (`npm run seams`, 10 checks): core-only and partial
   compositions — every non-composed feature API throws the documented
   `feature_not_configured` error.
4. **Packaging check** (`npm run pack:check`): `npm pack` → throwaway
   consumer → strict-TS type-check of every subpath export + tree-shake size
   gates against the built `dist/`.
5. **Bundle-size harness** (`npm run measure`): the number the whole project
   exists to keep honest (see
   [design/bundle-size-experiment.md](./design/bundle-size-experiment.md)).

## Repository map

```
mini-msal/
├── docs/                 you are here (README, UPGRADING, ALACARTE,
│   │                     GAP_REPORT, design/)
│   └── design/           deep-dives: conformance suite internals, wire
│                         protocols, bundle-size experiment, original spec
├── examples/             one minimal consumer app per profile
│                         (size-tracked + smoke-tested)
├── packages/
│   ├── browser/          @mini-msal/browser  (core + feature subpaths)
│   ├── compat/           @mini-msal/compat   (the drop-in)
│   └── react/            @mini-msal/react    (React bindings)
└── test/
    ├── infra/            mock OIDC IdP, static server, build config
    ├── apps/             the shared demo/e2e app + conformance harness pages
    ├── e2e/              25-check dual-stack E2E
    ├── examples/         examples smoke check (sign-in per example)
    ├── unit/             seam-hardening checks
    ├── packaging/        npm-pack consumer check
    ├── conformance/      122-scenario conformance runner + snapshots
    └── bundle-size/      size measurement + source-map attribution
```

## Running it

```sh
npm install
npm run build              # generate app variants + build all bundles
npm run e2e                # 25-check E2E against the local mock IdP
npm run examples:smoke     # each examples/ profile signs in vs the mock IdP
npm run conformance        # capture real → replay mini → regenerate GAP_REPORT.md
npm run conformance:check  # determinism check (real vs its own snapshots)
npm run seams              # feature-seam checks (core-only + partial compositions)
npm run pack:check         # npm-pack → consumer type-check + tree-shake gates
npm run measure            # bundle sizes (min / gzip / brotli)
```

To try it against a real Entra ID tenant, see
[design/bundle-size-experiment.md](./design/bundle-size-experiment.md#testing-with-real-entra-id).

## Status and next steps

Current conformance standing (122 scenarios): **122 identical · 0 missing
features · 0 behavioral diffs · 0 bugs** — the C-series (full compat parity,
including the post-audit C11–C21 gap closure) is complete; the D-series
(consumer packaging, docs) is underway. e2e 25/25, seams 12/12, packaging
check green, and each [examples/](../examples/) profile is size-tracked and
smoke-tested. The work is driven task-by-task from
[design/PARITY_STATE.md](./design/PARITY_STATE.md); remaining: the final
requirements audit (D4).

Measured size matrix (2026-07-14, msal-browser 5.16.0, `npm run measure`):

| build | minified | gzip -9 | brotli |
|---|---:|---:|---:|
| real: msal-react + msal-browser stack | 248.8 KB | 64.9 KB | 53.9 KB |
| real: msal-browser only (no React) | 220.5 KB | 55.4 KB | 46.4 KB |
| real: redirect-bridge page | 6.5 KB | 2.7 KB | 2.4 KB |
| **mini: compat + react stack** | **72.0 KB** | **23.5 KB** | **20.9 KB** |
| **mini: `@mini-msal/compat`, every feature (no React)** | **62.1 KB** | **20.1 KB** | **17.9 KB** |
| **mini: `createClient` core only (redirect + silent + multi-account + sign-out)** | **31.3 KB** | **10.8 KB** | **9.8 KB** |
| mini: redirect-bridge page | 0.6 KB | 0.4 KB | 0.3 KB |

Full drop-in parity costs ~28% of real's bytes (3.6× smaller browser-only,
3.5× smaller with React); consumers who compose less pay less, down to the
~30 KB core ([ALACARTE.md](./ALACARTE.md) has the per-feature price list).
Every parity task recorded its size cost in the PARITY_STATE progress log.
