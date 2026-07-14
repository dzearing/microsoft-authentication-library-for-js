# compat — drop-in `PublicClientApplication`

**Composes:** everything. `@mini-msal/compat` exposes the same
`PublicClientApplication` surface as `@azure/msal-browser` with all features
(popup, PoP, broker, localStorage, cache migration, telemetry) pre-composed.

**Measured (2026-07-14):** 61.1 KB minified / 19.9 KB gzip for the entire
app bundle (`example-compat` in `npm run measure`) — 3.6× smaller than real
msal-browser's 220.5 KB for the same code.

Migrating an existing msal-browser app is two steps
([docs/UPGRADING.md](../../docs/UPGRADING.md)):

1. Swap the import: `@azure/msal-browser` → `@mini-msal/compat`.
2. Serve mini's redirect-bridge on your popup page (see the snippet in
   [../core-popup/README.md](../core-popup/README.md)).

The cache schema is real v5's `msal.3`, so signed-in users stay signed in
across the swap.

## Files

- `main.ts` — classic PCA boilerplate: `initialize` →
  `handleRedirectPromise` → `loginPopup` or greeting + silent token +
  `logoutPopup`.
- `authConfig.ts` — replace the placeholders with your registration's
  values; register both the app page and `popup.html` as redirect URIs.

## Run it

From the repo root: `npm run build`, then `npm run examples:smoke` proves
this example completes a real popup sign-in headlessly against the local
mock IdP (`dist/example-compat-smoke`).
