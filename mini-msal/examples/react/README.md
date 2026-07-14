# react — `MsalProvider` + templates + hooks

**Composes:** `@mini-msal/compat` (the drop-in `PublicClientApplication`)
under `@mini-msal/react`, which matches msal-react 5.5.1's contracts
(`MsalProvider`, the templates, `useMsal`/`useAccount`/`useIsAuthenticated`/
`useMsalAuthentication`, `withMsal`). The provider initializes the instance
and completes redirect roundtrips itself — no manual
`handleRedirectPromise` needed.

**Measured (2026-07-14):** 66.4 KB minified / 21.8 KB gzip
(`example-react` in `npm run measure`), with React external like the rest
of the size matrix — this is what the auth stack adds to an app that
already ships React. The real msal-react + msal-browser stack measured the
same way is 248.8 KB.

To shrink further, pass an à-la-carte client instead of compat: the
provider's `instance` prop accepts `AuthClient & PopupClient` (see
[docs/ALACARTE.md](../../docs/ALACARTE.md)).

## Files

- `main.tsx` — `MsalProvider` wrapping `AuthenticatedTemplate` (greeting +
  silently acquired token + sign-out) and `UnauthenticatedTemplate`
  (redirect sign-in button).
- `authConfig.ts` — replace the placeholders with your registration's
  values.

## Run it

From the repo root: `npm run build`, then `npm run examples:smoke` proves
this example signs in headlessly against the local mock IdP
(`dist/example-react-smoke`, which bundles React so it runs standalone).
