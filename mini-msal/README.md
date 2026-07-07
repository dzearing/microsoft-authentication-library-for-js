# mini-msal (prototype)

A from-scratch, ~20 KB reimplementation of the `@azure/msal-browser` +
`@azure/msal-react` surface a typical SPA uses — same protocol, same cache
format, same API shapes — with a conformance suite that measures exactly how
close it is to the real thing.

**Start here → [docs/README.md](./docs/README.md)** (goals, architecture,
how to consume it, current parity status).

| Doc | What it covers |
|---|---|
| [docs/README.md](./docs/README.md) | The prototype: goals, architecture, API surface, consumption |
| [docs/GAP_REPORT.md](./docs/GAP_REPORT.md) | Generated 75-scenario parity report vs msal-browser 5.16.0 |
| [docs/design/](./docs/design/) | Deep dives: conformance suite internals, broker/NAA wire protocols, bundle-size experiment, original spec |

## Quick start

```sh
npm install
npm run build              # generate app variants + build all bundles
npm run e2e                # 25-check dual-stack E2E vs local mock IdP
npm run conformance        # capture real → replay mini → regenerate docs/GAP_REPORT.md
npm run measure            # bundle sizes: real stack ~249 KB min vs mini ~20 KB
```

Layout: library code in `packages/` (`@mini-msal/browser`, `@mini-msal/react`),
all test infrastructure under `test/` (infra, apps, e2e, conformance,
bundle-size), documentation under `docs/`.
