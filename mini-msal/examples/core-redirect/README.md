# core-redirect — redirect-only SPA on core

**Composes:** nothing — `createClient(msalConfig)` from `@mini-msal/browser`
is the whole stack. This is the pay-to-play floor.

**Measured (2026-07-14):** 30.3 KB minified / 10.6 KB gzip for the entire
app bundle (`example-core-redirect` in `npm run measure`). Real
msal-browser doing the same job ships 220.5 KB.

Sign-in (`loginRedirect`), silent acquisition (`acquireTokenSilent`: cache →
refresh token → `prompt=none` iframe), and `logoutRedirect` all ride
full-page redirects. No redirect-bridge page is required for this profile —
if you later add popups or `ssoSilent`, see
[docs/ALACARTE.md](../../docs/ALACARTE.md) for the bridge-page rule.

## Files

- `main.ts` — the app: initialize → `handleRedirectPromise` → sign-in
  button or greeting + token + sign-out.
- `authConfig.ts` — replace the placeholder `clientId`/`authority`/
  `redirectUri` with your app registration's values.

## Run it

From the repo root: `npm run build`, then `npm run examples:smoke` proves
this example signs in headlessly against the local mock IdP
(`dist/example-core-redirect-smoke`). To adopt it, copy both files into
your bundler of choice and serve the bundle from a page with
`<div id="root">`.
