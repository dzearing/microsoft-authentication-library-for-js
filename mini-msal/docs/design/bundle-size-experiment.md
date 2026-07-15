# Bundle-size experiment (where mini-msal started)

Measures the real-world min+gzip cost of a **full-authentication-flow**
`@azure/msal-react` app using rspack with aggressive production settings (SWC
minifier w/ 3 compress passes, toplevel mangling, scope hoisting, full tree
shaking, es2022 output), and compares it against `@mini-msal/*` — a
from-scratch reimplementation of the same exercised behavior — to establish
the achievable size floor. React/react-dom are externals in all measured
variants: we measure what the MSAL stack *adds* to an app that already ships
React.

The test app (`test/apps/app.tsx`) deliberately references the full public API
a production app uses, so tree shaking cannot drop real code paths: redirect,
popup, and ssoSilent login; acquireTokenSilent (+CacheLookupPolicy) with
popup→redirect fallback; event callbacks (add/remove); active-account
bookkeeping; getAccount lookups; logoutRedirect + logoutPopup; MsalProvider,
Authenticated/Unauthenticated/MsalAuthenticationTemplate, useMsal, useAccount,
useIsAuthenticated, useMsalAuthentication, withMsal.

## Usage

```sh
npm install
npm run build      # sync variants, build all bundles, emit e2e html
npm run measure    # min / gzip -9 / brotli -11 per variant
npm run analyze -- msal-stack [--files]   # per-package byte attribution via source maps
npm run e2e        # full-flow E2E: BOTH stacks vs the local mock IdP (headless Chrome)
```

`test/apps/mini-app.tsx`, `test/apps/mini-app-mock.tsx`, and
`test/apps/app-mock.tsx` are GENERATED from `test/apps/app.tsx` by
`test/infra/sync-variants.mjs` (only import specifiers / auth config differ),
so all variants exercise byte-identical usage.

## Variants

| variant | what it is |
|---|---|
| `msal-stack` | full-flow app: msal-react + msal-browser + msal-common (react external) |
| `msal-browser-core` | same flow, no React layer (`test/apps/msal-only.ts`) |
| `msal-import-all` | tree-shaking ceiling: whole msal-browser surface kept alive |
| `msal-redirect-bridge` | the extra redirect page v5 apps must ship for popup/ssoSilent flows (`@azure/msal-browser/redirect-bridge`) |
| `mini-msal-stack` | identical app on `@mini-msal/compat` + `@mini-msal/react` (react external) |
| `mini-compat` | same no-React flow as `msal-browser-core` on `@mini-msal/compat`, every feature composed (`test/apps/mini-only.ts`, generated from `msal-only.ts`) |
| `mini-core` | pay-to-play floor: `createClient` core only — redirect sign-in/sign-out, silent, ssoSilent, multi-account; no popup/broker/naa/telemetry (`test/apps/mini-core.ts`) |
| `mini-redirect-bridge` | mini's counterpart of `msal-redirect-bridge` (`@mini-msal/browser/redirect-bridge`) |
| `mini-mock-app`, `real-mock-app` | E2E builds against the mock IdP (react bundled; excluded from measurement) |

## Results (2026-07-14 — full 122/122 conformance parity; msal-browser 5.16.0, msal-react 5.5.1)

| variant | minified | gzip -9 | brotli |
|---|---:|---:|---:|
| msal-import-all | 257.0 KB | 65.1 KB | 54.6 KB |
| msal-stack | 248.8 KB | 64.9 KB | 53.9 KB |
| msal-browser-core | 220.5 KB | 55.4 KB | 46.4 KB |
| msal-redirect-bridge | 6.5 KB | 2.7 KB | 2.4 KB |
| **mini-msal-stack** (compat + react) | **72.0 KB** | **23.5 KB** | **20.9 KB** |
| **mini-compat** (every feature, no React) | **62.1 KB** | **20.1 KB** | **17.9 KB** |
| **mini-core** (createClient core only) | **31.3 KB** | **10.8 KB** | **9.8 KB** |
| mini-redirect-bridge | 0.6 KB | 0.4 KB | 0.3 KB |

At 100% conformance (122/122 scenarios identical, e2e 25/25), the full
drop-in is **3.5× smaller** than tree-shaken msal-browser (62.1 vs 220.5 KB
min) and the full React stack is **3.5× smaller** (72.0 vs 248.8 KB); a
redirect-only SPA composing nothing pays 31.3 KB (**7.0× smaller**). For
history: at the 75-scenario milestone (2026-07-12) compat measured 32.5 /
stack 39.5 / core 19.5 — the C11–C21 unobserved-surface parity push (perf
events, error/logger namespaces, navigation seams, cache entity semantics,
authority discovery, config knobs, PoP) plus the D-series export-surface
completion roughly doubled mini while real stayed constant. The original
pre-parity prototype (2026-07-01 baseline, ~30% of the feature surface)
measured 20.1 KB min / 6.9 KB gz for the whole react stack. Per-profile
consumption sizes and the plain-language technique write-up live in
[../ALACARTE.md](../ALACARTE.md) and [../SIZE.md](../SIZE.md).

Tree shaking is nearly powerless against MSAL: a full-flow app's tree-shaken
msal-browser (220.5 KB) is only ~14% smaller than the entire library surface
kept alive (257.0 KB), because `PublicClientApplication` → `StandardController`
statically imports every flow (popup, broker/WAM, IndexedDB, telemetry, …)
regardless of what the app uses.

Net of app code (~2.6 KB) and bundler runtime (~2.3 KB), the MSAL libraries
cost ~243 KB min / ~63 KB gz (+6.5 KB bridge page for popup apps); mini's
compat drop-in costs ~57 KB min for the same exercised behavior (+0.6 KB
bridge page) — **~4.3×** at library-only accounting, at 100% conformance.

## E2E verification

`npm run e2e` starts a mock OIDC IdP (https, self-signed cert in `.cert/`,
honors `prompt=none` session state, echoes nonce, returns `client_info`) and a
static server, then drives headless system Chrome (playwright-core) through
the SAME scenario for **both** the real MSAL app and the mini-msal app:

1. page load → silent SSO auto-attempt fails with `login_required` →
   `MsalAuthenticationTemplate` auto-triggers a redirect login → PKCE(S256)
   code exchange → authenticated UI renders the profile
2. `withMsal` status banner + `MsalAuthenticationTemplate` protected content
3. `getAllAccounts` returns the signed-in account
4. `acquireTokenSilent` served from cache
5. `acquireTokenPopup` full popup roundtrip (both libraries complete via
   their redirect-bridge page, real v5's actual mechanism)
6. add a second account (`loginPopup` + `prompt=select_account`), switcher UI
   lists both, switch the active account both ways, per-account token
   isolation (distinct mock users get distinct tokens)
7. `logoutPopup({account})` removes only that account; final logout clears all
8. **cache interop proof**: sign in on one stack, navigate the same
   tab to the other — it reads the first stack's cache (accounts, tokens,
   active account) with zero IdP roundtrips; verified in BOTH directions.
   mini-msal reads/writes real MSAL's exact v5 cache schema ("msal.3"
   entities, token-key indexes, active-account-filters).
9. cancelled-login regression (mini): an `#error=access_denied` redirect
   response surfaces as an error without an auto-login retry loop

All 25 checks pass. Fidelity notes discovered while building
this: v5 popups REQUIRE the redirect-bridge page (a blank popup redirect page
times out); real MSAL rejects `http://` authorities; mini-msal replicates real
MSAL's iframe/popup reload guards (`block_iframe_reload`, `block_nested_popups`,
`redirect_in_iframe`) and msal-react's stable-account-reference behavior
(without it, effect loops occur — same class of bug real msal-react guards
against).

## Testing with real Entra ID

The mock-IdP E2E requires no registration. To test against real AAD:
1. Register an app (free): entra.microsoft.com → App registrations → New →
   platform **Single-page application**, redirect URIs `http://localhost:4173/`
   plus (for popup/SSO flows) `http://localhost:4173/popup.html` and
   `http://localhost:4173/blank.html`.
2. Put its Application (client) ID in `test/apps/authConfig.ts` (clientId),
   and optionally your tenant authority.
3. `npm run build && node test/infra/serve.mjs`, open http://localhost:4173/
   and sign in. The banner shows which stack is loaded (blue = real MSAL,
   green = mini) and toggles between them; both run the identical app against
   real Entra ID.
