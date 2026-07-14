# mini-msal (prototype)

A from-scratch, pay-to-play reimplementation of `@azure/msal-browser` +
`@azure/msal-react` — same protocol, same cache format, same API shapes.
Consume it as a one-import drop-in (`@mini-msal/compat`, 61.7 KB min with
every feature composed, vs 220.5 KB for tree-shaken `@azure/msal-browser`)
or à la carte (`createClient` core, 29.7 KB min + tree-shakable feature
modules). A 121-scenario conformance suite measures exactly how close it is
to the real thing: **121/121 identical** (0 behavioral diffs, 0 bugs).

**Start here → [docs/README.md](./docs/README.md)** (goals, architecture,
consumption profiles, current parity status).

| Doc | What it covers |
|---|---|
| [docs/README.md](./docs/README.md) | The prototype: goals, architecture, quick-start per consumption profile |
| [docs/UPGRADING.md](./docs/UPGRADING.md) | Migrating from `@azure/msal-browser`/`@azure/msal-react` — the one-import drop-in |
| [docs/ALACARTE.md](./docs/ALACARTE.md) | Lowering your bundle cost profile-by-profile, with measured sizes |
| [examples/](./examples/) | Runnable minimal app per profile (core-redirect 30.3 KB, core+popup 33.2, compat 61.1, react 66.4), each smoke-tested end-to-end |
| [docs/GAP_REPORT.md](./docs/GAP_REPORT.md) | Generated 121-scenario parity report vs msal-browser 5.16.0 |
| [docs/design/](./docs/design/) | Deep dives: conformance suite internals, broker/NAA wire protocols, bundle-size experiment, original spec |

## Quick start

```sh
npm install
npm run build              # generate app variants + build all bundles
npm run e2e                # 25-check dual-stack E2E vs local mock IdP
npm run examples:smoke     # each examples/ profile signs in vs the mock IdP
npm run conformance        # capture real → replay mini → regenerate docs/GAP_REPORT.md
npm run seams              # feature-seam checks (core-only + partial compositions)
npm run pack:check         # npm-pack → consumer type-check + tree-shake gates
npm run measure            # bundle sizes: real ~220 KB vs compat ~62 KB vs core ~31 KB
```

Layout: library code in `packages/` (`@mini-msal/browser` core + feature
subpaths, `@mini-msal/compat` drop-in, `@mini-msal/react`), consumer
examples in `examples/` (one per profile, size-tracked + smoke-tested),
all test infrastructure under `test/` (infra, apps, e2e, examples, unit,
packaging, conformance, bundle-size), documentation under `docs/`.
