# Examples — one minimal app per consumption profile

Each folder is a complete, minimal sign-in app (a `main.ts(x)` plus an
`authConfig.ts` you fill in with your registration values). Every example is
built by the repo's rspack infra with the same settings as the size matrix,
so its measured size is a tracked regression check, and each one is
smoke-tested end-to-end (real headless sign-in against the local mock IdP).

| Example | Imports | Minified | Gzip | Bridge page needed? |
|---|---|---:|---:|---|
| [core-redirect](./core-redirect/) | `@mini-msal/browser` | 30.4 KB | 10.6 KB | no |
| [core-popup](./core-popup/) | core + `…/popup` | 33.3 KB | 11.5 KB | yes |
| [compat](./compat/) | `@mini-msal/compat` | 61.5 KB | 20.0 KB | yes |
| [react](./react/) | compat + `@mini-msal/react` | 66.8 KB | 22.0 KB | for popups/ssoSilent |

Sizes measured 2026-07-14 (`npm run measure`); react is measured with React
external, like the rest of the matrix — it's what the auth stack adds to an
app that already ships React. Real `@azure/msal-browser` doing the compat
example's job is 220.5 KB; the real react stack is 248.8 KB.

## Commands (repo root)

- `npm run build` — builds each example twice: `dist/example-<name>`
  (measured, placeholder config) and `dist/example-<name>-smoke` (identical
  code with `authConfig.ts` swapped for the mock-IdP config).
- `npm run examples:smoke` — headless Chrome signs in through every smoke
  build and asserts a silently acquired token renders.
- `npm run measure` — prints the size table; the `example-*` rows are the
  regression check.

## Picking a profile

Start from [compat](./compat/) for a drop-in migration
([docs/UPGRADING.md](../docs/UPGRADING.md)), then step down to à-la-carte
composition to shed bytes ([docs/ALACARTE.md](../docs/ALACARTE.md)).
