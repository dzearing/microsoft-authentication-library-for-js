# À la carte: lower your bundle cost

`@mini-msal/compat` composes **every** feature so migration is one import
swap. But mini-msal's core is pay-to-play: `createClient(config, features)`
plus tree-shakable feature modules under `@mini-msal/browser/*`. If your app
doesn't call popups, the broker, or telemetry, don't ship them — this guide
steps the bundle down profile by profile, with measured sizes at each step.

All numbers are minified KB from this repo's build settings (rspack, swc
minifier with 3 compress passes, es2022, full tree shaking) against
msal-browser 5.16.0's 220.5 KB. Regenerate the headline matrix with
`npm run measure`; per-feature deltas were measured 2026-07-14 by bundling
the same minimal app with one feature composed at a time. Consumer builds
from the published `dist` (default swc settings, verified by
`npm run pack:check`) land within ~1 KB of these numbers.

## How composition works

A feature is a plain function; you pass the ones you want to `createClient`:

```ts
import { createClient, type AuthClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";
import { telemetry } from "@mini-msal/browser/telemetry";

const client = createClient(
    { auth: { clientId: "11111111-2222-3333-4444-555555555555" } },
    [popup, telemetry]
) as AuthClient & PopupClient;
```

Only the composed modules end up in your bundle. There are no silent
failures: calling a feature-owned API you didn't compose throws a
`BrowserAuthError` with code `feature_not_configured` whose message names the
exact missing import, e.g.:

```
feature_not_configured: loginPopup requires composing popup from "@mini-msal/browser/popup"
```

## The feature catalog

| Import | Adds (min) | What it enables — compose it if you call… |
|---|---:|---|
| `@mini-msal/browser` (core) | 29.7 KB base | `initialize`, `loginRedirect`/`acquireTokenRedirect`/`handleRedirectPromise`, `acquireTokenSilent` (cache → refresh token → `prompt=none` iframe), `ssoSilent`, accounts/active account, `logoutRedirect`, `clearCache`, `hydrateCache`, events, logger, error classes, authority validation/discovery |
| `popup` from `…/popup` | +2.9 KB | `loginPopup`, `acquireTokenPopup`, `logoutPopup` (incl. `mainWindowRedirectUri`) |
| `pop` from `…/pop` | +2.0 KB | `authenticationScheme: "pop"` / `"ssh-cert"` requests — RSA keypair in IndexedDB, signed HTTP requests (SHR), scheme-aware cache |
| `broker` from `…/broker` | +6.9 KB | Platform broker (WAM) via DOM/extension transports; hybrid `acquireTokenByCode` |
| `localStorageCache` from `…/local-storage` | +2.6 KB | `cache.cacheLocation: "localStorage"` — encrypted-at-rest entities, cross-tab sync, plaintext KMSI persistence (required for that cacheLocation; no-op otherwise) |
| `cacheMigration` from `…/cache-migration` | +3.0 KB | Reading caches written by older msal-browser majors (`msal.0/1/2` → `msal.3`) + retention TTL |
| `telemetry` from `…/telemetry` | +9.4 KB | `addPerformanceCallback` perf events, `BrowserPerformanceClient`, performance marks, wrapper SKU headers |
| `createNestableClient` from `…/naa` | +8.4 KB | Nested app auth (Teams/Office host bridge) — a factory entrypoint rather than a composed feature |
| `…/redirect-bridge` | 0.6 KB page | The bridge page script for popup/`ssoSilent`/silent-iframe completion (separate page bundle, not part of your app bundle) |
| `@mini-msal/react` | +10.0 KB | `MsalProvider`, hooks, templates, `withMsal` (React itself not counted) |

Deltas compose approximately additively (shared plumbing lives in core).

## Step down, profile by profile

### Step 0 — where you start: compat drop-in, 62.1 KB

```ts
import { PublicClientApplication } from "@mini-msal/compat";

const pca = new PublicClientApplication({
    auth: { clientId: "11111111-2222-3333-4444-555555555555" },
});
```

Everything composed (popup, pop, broker, localStorage, cache migration,
telemetry), plus real msal-browser's exported constant/error-code namespaces.
Already 3.6× smaller than real's 220.5 KB — this is the right place to stay
while validating the migration (see [UPGRADING.md](./UPGRADING.md)).

### Step 1 — same features, explicit composition: 56.3 KB

Swap the compat class for `createClient` and list what compat composed —
then delete the lines your app never calls. The `*ErrorCodes` namespaces and
enum objects compat re-exports for surface parity drop out too.

```ts
import { createClient, type AuthClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";
import { pop } from "@mini-msal/browser/pop";
import { broker, type BrokerClient } from "@mini-msal/browser/broker";
import { localStorageCache } from "@mini-msal/browser/local-storage";
import { cacheMigration } from "@mini-msal/browser/cache-migration";
import { telemetry, type TelemetryClient } from "@mini-msal/browser/telemetry";

const config = { auth: { clientId: "11111111-2222-3333-4444-555555555555" } };
const client = createClient(config, [
    localStorageCache,
    cacheMigration,
    popup,
    pop,
    broker,
    telemetry, // keep telemetry last: it wraps the other features' methods
]) as AuthClient & PopupClient & BrokerClient & TelemetryClient;
```

### Step 2 — popup SPA, sessionStorage cache: 32.6 KB

The common interactive SPA: popups plus the silent ladder, default
sessionStorage cache, no broker/telemetry/PoP.

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

### Step 3 — redirect-only SPA, the floor: 29.7 KB

Redirect sign-in, silent acquisition, multi-account, sign-out — no feature
modules at all. **7.4× smaller** than real msal-browser doing the same job.

```ts
import { createClient, InteractionRequiredAuthError } from "@mini-msal/browser";

const client = createClient({
    auth: {
        clientId: "11111111-2222-3333-4444-555555555555",
        redirectUri: window.location.origin,
    },
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
        } else {
            throw e;
        }
    }
}
```

(The repo's own `mini-core` measurement entry, which exercises a slightly
wider API surface incl. `ssoSilent` and events, lands at 31.3 KB — that's the
number the headline matrix reports for core-only.)

### React at any step: +10.0 KB

`@mini-msal/react` works over compat **and** à-la-carte clients — its
`instance` prop is typed as msal-react's `IPublicClientApplication`
(`AuthClient & PopupClient`), which a `[popup]` composition satisfies:

```tsx
import { createClient, type AuthClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";
import { MsalProvider, AuthenticatedTemplate, useMsal } from "@mini-msal/react";
import { createRoot } from "react-dom/client";

const client = createClient(
    { auth: { clientId: "11111111-2222-3333-4444-555555555555" } },
    [popup]
) as AuthClient & PopupClient;

function Home() {
    const { accounts } = useMsal();
    return <p>signed in as {accounts[0]?.username}</p>;
}

createRoot(document.getElementById("root")!).render(
    <MsalProvider instance={client}>
        <AuthenticatedTemplate>
            <Home />
        </AuthenticatedTemplate>
    </MsalProvider>
);
```

## Don't forget the bridge page

Popup, `ssoSilent`, and `acquireTokenSilent`'s `prompt=none` iframe fallback
complete via a page served at the request's `redirectUri` (same mechanism as
real MSAL v5). If you use any of those, serve this 0.6 KB page there:

```ts
// bridge.ts — the entire bridge page script
import { broadcastResponseToMainFrame } from "@mini-msal/browser/redirect-bridge";

broadcastResponseToMainFrame().catch(console.error);
```

A pure redirect app (step 3, with `acquireTokenRedirect` as the interaction
fallback) never loads it.

## The full picture

| Profile | Composition | min | vs real |
|---|---|---:|---:|
| real msal-browser (tree-shaken) | — | 220.5 KB | 1× |
| compat drop-in | everything | 62.1 KB | 3.5× smaller |
| explicit full composition | core + 6 features | 56.3 KB | 3.9× smaller |
| popup SPA | core + `popup` | 32.6 KB | 6.8× smaller |
| redirect-only SPA | core | 29.7 KB | 7.4× smaller |

Add `@mini-msal/react` (+10.0) and/or the 0.6 KB bridge page per your flows;
real's equivalents are +28.3 (msal-react) and 6.5 KB (bridge).
