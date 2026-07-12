# mini-msal (prototype)

A from-scratch, pay-to-play reimplementation of `@azure/msal-browser` +
`@azure/msal-react` — same protocol, same cache format, same API shapes.
Consume it as a one-import drop-in (`@mini-msal/compat`, ~37 KB min with
every feature composed) or à la carte (`createClient` core ~19 KB min +
tree-shakable feature modules), vs ~249 KB min for the real stack. A
75-scenario conformance suite measures exactly how close it is to the real
thing (currently 68/75 identical, 0 bugs).

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
npm run measure            # bundle sizes: real ~249 KB vs compat ~37 KB vs core ~19 KB
```

Layout: library code in `packages/` (`@mini-msal/browser` core + feature
subpaths, `@mini-msal/compat` drop-in, `@mini-msal/react`),
all test infrastructure under `test/` (infra, apps, e2e, conformance,
bundle-size), documentation under `docs/`.
