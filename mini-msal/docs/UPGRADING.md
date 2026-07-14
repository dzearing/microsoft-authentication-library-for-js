# Migrating from `@azure/msal-browser` / `@azure/msal-react`

`@mini-msal/compat` is a drop-in replacement for `@azure/msal-browser` 5.x:
same `PublicClientApplication` API, same wire protocol, same error classes and
codes, same event streams, same on-disk cache. `@mini-msal/react` is the same
for `@azure/msal-react` 5.x. Parity is measured, not claimed: a 121-scenario
conformance suite diffs every observable detail against the real libraries
(see [GAP_REPORT.md](./GAP_REPORT.md) — currently **121/121 identical**).

What you save (minified, tree-shaken, measured 2026-07-14):

| | `@azure/msal-*` | `@mini-msal/*` |
|---|---:|---:|
| browser library only | 220.5 KB | 61.7 KB |
| browser + React bindings | 248.8 KB | 71.7 KB |
| redirect-bridge page | 6.5 KB | 0.6 KB |

> **Prototype note**: these packages are not published to npm. Consume them
> from this repo (workspace dependency, or `npm pack` the `packages/*`
> directories — that packaging path is CI-verified by `npm run pack:check`).

## Step 1 — swap the imports

```diff
-import { PublicClientApplication, InteractionRequiredAuthError } from "@azure/msal-browser";
+import { PublicClientApplication, InteractionRequiredAuthError } from "@mini-msal/compat";
```

```diff
-import { MsalProvider, useMsal, AuthenticatedTemplate } from "@azure/msal-react";
+import { MsalProvider, useMsal, AuthenticatedTemplate } from "@mini-msal/react";
```

Nothing else changes. Your existing configuration object, request objects,
`catch` blocks, and event callbacks keep working:

```ts
import { PublicClientApplication, EventType } from "@mini-msal/compat";

const pca = new PublicClientApplication({
    auth: {
        clientId: "11111111-2222-3333-4444-555555555555",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin,
    },
    cache: { cacheLocation: "sessionStorage" },
});
await pca.initialize();
await pca.handleRedirectPromise();
pca.addEventCallback((message) => {
    if (message.eventType === EventType.LOGIN_SUCCESS) {
        console.log("signed in");
    }
});
```

The compat surface includes the full classic API: `loginRedirect`/`loginPopup`,
`acquireTokenSilent`/`Popup`/`Redirect`/`ByCode`, `ssoSilent`, logout (popup,
redirect, `mainWindowRedirectUri`), accounts and active-account APIs,
`clearCache`, `hydrateCache`, top-level `loadExternalTokens`,
`createStandardPublicClientApplication`,
`createNestablePublicClientApplication` (nested app auth for Teams/Office
hosts), `addPerformanceCallback`, `getLogger`/`setLogger`/`LogLevel`,
`NavigationClient`/`setNavigationClient`, the error classes
(`AuthError`, `BrowserAuthError`, `ClientAuthError`,
`InteractionRequiredAuthError`, `ServerError`, …) and their `*ErrorCodes`
namespaces, `EventType`, `InteractionStatus`, `CacheLookupPolicy`, and the
enums (`BrowserCacheLocation`, `ProtocolMode`, `PromptValue`,
`AuthenticationScheme` — pop and ssh-cert schemes are implemented).

## Step 2 — swap the redirect-bridge page

Like real MSAL v5, popup, `ssoSilent`, and silent-iframe responses complete
via a small page served at the request's `redirectUri` that relays the auth
response to the main frame. If your app serves real MSAL's bridge bundle
(commonly at `/popup.html` or `/redirect.html`), rebuild that page's script
from mini's bridge module — same page contract, 0.6 KB instead of 6.5 KB:

```ts
// bridge.ts — the entire bridge page script
import { broadcastResponseToMainFrame } from "@mini-msal/browser/redirect-bridge";

broadcastResponseToMainFrame().catch(console.error);
```

Serve it from the same URL your popup/silent `redirectUri` points at. Pure
redirect flows (`loginRedirect` + `handleRedirectPromise`) don't use the
bridge page.

## Signed-in users stay signed in

mini-msal reads and writes real MSAL's exact v5 cache schema (`msal.3`
entities, token-key index, active-account filters), in both `sessionStorage`
(default) and `localStorage` (encrypted entities + cross-tab sync, matching
real's observable shape). Deploying the swap does not sign anyone out:

- Sessions cached by `@azure/msal-browser` 5.x are picked up as-is. The e2e
  suite proves interop in both directions (real → mini and mini → real).
- Caches written by older msal-browser majors (`msal.0/1/2` schemas) are
  migrated to `msal.3` on `initialize()`, exactly as real 5.x does.
- "Keep me signed in" (KMSI) accounts persist in plaintext `localStorage`
  entities like real, so they survive cookie loss.

## What to verify after switching

A one-time smoke pass over the flows your app actually uses:

1. **Redirect round trip** — `loginRedirect` → sign-in → `handleRedirectPromise`
   resolves with the account; deep links return to the original URL when
   `navigateToLoginRequestUrl` applies.
2. **Popup flows** — `loginPopup`/`acquireTokenPopup` complete (this exercises
   your swapped bridge page).
3. **Silent ladder** — `acquireTokenSilent` returns cached tokens instantly,
   refreshes on expiry, and throws `InteractionRequiredAuthError` when
   interaction is genuinely required (your existing fallback should catch it).
4. **Existing sessions** — deploy over a signed-in session and confirm the
   user stays signed in (cache interop).
5. **Logout** — `logoutRedirect`/`logoutPopup` end the session and fire
   `msal:logoutSuccess`.
6. If you consume **events**, **perf telemetry** (`addPerformanceCallback`),
   or the **platform broker / nested app auth**, exercise one call of each —
   all are implemented and conformance-pinned, but they're the least commonly
   used surfaces.

## Known non-goals

Documented differences rather than silent divergence — the full evidence is
[GAP_REPORT.md](./GAP_REPORT.md):

- **B2C and ADFS authorities** are untested and out of scope (AAD, CIAM
  `*.ciamlogin.com`, and `knownAuthorities`/OIDC `protocolMode` validation are
  implemented).
- Real's **native broker in-memory token cache** is not replicated: mini
  re-asks the platform broker on each call (the broker does its own caching).
- Anything the 121 conformance scenarios don't observe is not guaranteed. If
  you find a behavioral difference, the conformance suite is the arbiter —
  add a scenario, capture real, and diff (see
  [design/conformance-suite-notes.md](./design/conformance-suite-notes.md)).

## Next step: shrink further

Compat composes **every** feature so the drop-in is zero-effort. If your app
doesn't call popups, the broker, telemetry, or localStorage, you can step down
to an à-la-carte composition and pay roughly half the bytes — see
[ALACARTE.md](./ALACARTE.md).
