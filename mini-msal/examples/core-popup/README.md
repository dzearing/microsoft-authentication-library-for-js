# core-popup — popup SPA on core + the popup feature

**Composes:** `createClient(msalConfig, [popup])` — core plus `popup` from
`@mini-msal/browser/popup` (`loginPopup` / `acquireTokenPopup` /
`logoutPopup`).

**Measured (2026-07-14):** 33.3 KB minified / 11.5 KB gzip for the entire
app bundle (`example-core-popup` in `npm run measure`) — the popup feature
costs ~3 KB over the core-redirect example.

## The bridge page

Popups complete on a dedicated redirect page (real msal-browser v5 works the
same way). Serve a `popup.html` next to your app whose only job is running
the 0.6 KB bridge, and register it as a SPA redirect URI:

```html
<!doctype html><title>auth</title><script type="module" src="bridge.js"></script>
```

where `bridge.js` bundles one import: `import "@mini-msal/browser/redirect-bridge";`
In this repo the built example's `popup.html` points at the shared
`dist/mini-redirect-bridge` bundle. The app passes that page as the popup
`redirectUri` (see `popupRequest` in `main.ts`).

## Files

- `main.ts` — initialize → sign-in button (`loginPopup`) or greeting +
  silently acquired token + `logoutPopup`.
- `authConfig.ts` — replace the placeholders with your registration's
  values; register both the app page and `popup.html` as redirect URIs.

## Run it

From the repo root: `npm run build`, then `npm run examples:smoke` proves
this example completes a real popup sign-in headlessly against the local
mock IdP (`dist/example-core-popup-smoke`).
