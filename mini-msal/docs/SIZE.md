# How mini-msal got small

mini-msal does the same job as `@azure/msal-browser` 5.16.0 — verified by a
122-scenario conformance suite that byte-compares wire traffic, cache
contents, events, errors, and telemetry against the real library (see
[GAP_REPORT.md](./GAP_REPORT.md)) — at a fraction of the bytes. This document
explains, in plain language, **what was done differently** to get there. Each
section is one technique, with a concrete example where that helps.

All sizes are minified KB measured by this repo's `npm run measure`
(rspack, swc minifier, es2022, full tree shaking; methodology in
[design/bundle-size-experiment.md](./design/bundle-size-experiment.md)).

| What you ship | real msal | mini-msal | smaller |
|---|---:|---:|---:|
| Drop-in, every feature (no React) | 220.5 KB | 62.1 KB | 3.5× |
| Full React stack (React external) | 248.8 KB | 72.0 KB | 3.5× |
| Popup SPA (core + `popup` only) | 220.5 KB | 32.6 KB | 6.8× |
| Redirect-only SPA (core only) | 220.5 KB | 29.7 KB | 7.4× |
| Redirect-bridge page (popup/silent flows) | 6.5 KB | 0.6 KB | 10.8× |

Notice the left column: **real msal costs 220.5 KB no matter which profile
you use.** That constant is the single biggest thing mini-msal changed — with
mini you pay for what you compose, so the "smaller" ratio grows as your needs
shrink. The per-profile recipes are in [ALACARTE.md](./ALACARTE.md).

## Where real msal-browser's bytes go

Source-map attribution of a real full-flow app bundle
(`npm run analyze -- msal-stack`) shows the shape of the problem:

- **~45 KB**: ten `interaction_client` classes (popup, redirect, three
  silent variants, platform broker, …) sharing two base-class layers.
- **~50 KB**: cache management, split across `msal-browser`'s cache layer
  *and* `msal-common`'s cache layer underneath it.
- **~26 KB**: `controllers` — a facade/controller layer between the
  `PublicClientApplication` class and the interaction clients.
- **~80 KB total**: the `@azure/msal-common` package — a
  platform-neutral core that browser code wraps, re-exports, and partially
  duplicates.

None of that is waste in real's terms — it is the cost of a layered,
multi-platform, everything-always-linked architecture. The techniques below
are all, one way or another, about not paying for layers.

## 1. Pay-to-play features instead of a static monolith

The dominant technique. In real msal-browser, constructing
`PublicClientApplication` builds a `StandardController`, and that module
statically imports **every flow the library supports** — whether or not the
app uses it. From real's `controllers/StandardController.mjs`:

```js
import { PopupClient } from '../interaction_client/PopupClient.mjs';
import { RedirectClient } from '../interaction_client/RedirectClient.mjs';
import { SilentIframeClient } from '../interaction_client/SilentIframeClient.mjs';
import { SilentRefreshClient } from '../interaction_client/SilentRefreshClient.mjs';
import { PlatformAuthInteractionClient } from '../interaction_client/PlatformAuthInteractionClient.mjs'; // WAM broker
import { SilentCacheClient } from '../interaction_client/SilentCacheClient.mjs';
import { SilentAuthCodeClient } from '../interaction_client/SilentAuthCodeClient.mjs';
import { CryptoOps } from '../crypto/CryptoOps.mjs';                      // incl. IndexedDB keystore
import { InitializeCache, /* …30+ telemetry event imports… */ } from '../telemetry/BrowserPerformanceEvents.mjs';
// …and ~20 more
```

Because everything is reachable from the constructor, tree shaking is nearly
powerless: a full-flow app's tree-shaken msal-browser (220.5 KB) is only ~14%
smaller than keeping the *entire* library surface alive (257.0 KB).

mini-msal inverts this. The core ships redirect + silent flows only; every
other capability is a separate module — a plain function you pass to
`createClient`. The bundler never sees what you don't import:

```ts
import { createClient, type AuthClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";

const client = createClient(
    { auth: { clientId: "11111111-2222-3333-4444-555555555555" } },
    [popup] // no broker, no telemetry, no PoP ⇒ zero bytes for them
) as AuthClient & PopupClient;
```

Each feature's measured price (from [ALACARTE.md](./ALACARTE.md)): popup
+2.9 KB, PoP/SSH schemes +2.0, localStorage cache +2.6, legacy-cache
migration +3.0, WAM broker +6.9, nested-app auth +8.4, perf telemetry +9.4.
A non-composed feature's API doesn't fail silently — it throws
`feature_not_configured` naming the exact import to add.

`@mini-msal/compat` is simply the composition of every feature (nested-app
auth is a separate factory entrypoint) behind real's
`PublicClientApplication` class shape — that's the 62.1 KB drop-in row.

## 2. Plain closures instead of class-and-layer towers

Real routes every call through a stack of objects: the
`PublicClientApplication` facade delegates to a controller chosen by an
"operating context", the controller delegates to per-flow interaction-client
classes, which share two levels of base classes. Each layer is real code:
constructors, fields, delegating methods, and boilerplate like this
(`operatingcontext/StandardOperatingContext.mjs`):

```js
class StandardOperatingContext extends BaseOperatingContext {
    getModuleName() { return StandardOperatingContext.MODULE_NAME; }
    getId() { return StandardOperatingContext.ID; }
    async initialize(correlationId) {
        this.available = typeof window !== "undefined";
        return this.available;
    }
}
```

mini-msal's core is **one function**. `createClient` closes over its state
(config, cache, in-flight dedupe maps), defines the flows as local
functions, and returns an object literal with the public methods. A feature
module is just another function handed the same context:

```ts
import type { ClientContext } from "@mini-msal/browser";

export function myFeature(ctx: ClientContext): void {
    // attach methods to ctx.client, reuse ctx's URL/cache/network seams
}
```

This is worth more than the deleted delegation code: minifiers can rename
`let metadataPromise` and local helper functions to one letter, but they must
preserve every `class` **method and property name** (`getModuleName`,
`initialize`, …) because JavaScript objects are dynamically inspectable.
Closure-heavy code minifies dramatically better than class-heavy code.

## 3. One browser-only package — no msal-common layer

Real msal-browser sits on `@azure/msal-common`, a platform-neutral core
shared with Node. The browser package wraps common's cache manager with a
browser cache manager, wraps its network module, re-exports its errors, and
carries interfaces sized for platforms the browser never runs. That layer
alone is ~80 KB minified of a real app bundle (see the attribution above) —
and its generality is the reason: it is written to serve the Node and browser
libraries alike, so the browser ships abstractions sized for both.

mini-msal has no "common" tier. It is one package written directly against
browser APIs (`fetch`, `crypto.subtle`, `sessionStorage`, DOM). The entire
browser package is nine source files; the same msal.3 cache schema real uses
is read and written directly (that's what keeps signed-in users signed in
across the migration), in a few hundred lines instead of two cache layers.

## 4. One error family instead of an error module per subsystem

Real ships, per subsystem: an error class module, a codes module exporting
every code as a named constant, and a `create*Error` factory —
`BrowserAuthError`, `BrowserConfigurationAuthError`, `ClientAuthError`,
`ClientConfigurationError`, `InteractionRequiredAuthError`, `NativeAuthError`,
`NestedAppAuthError`, … plus utility functions per class.

mini keeps the same **observable** contract (class names, `errorCode`,
`errorMessage`, aka.ms message format — all conformance-pinned) from one tiny
family: a single base class computes the default message, and each subclass
is two lines that set the name:

```ts
const AKA = (code: string) =>
    `See https://aka.ms/msal.js.errors#${code} for details`;

export class AuthError extends Error {
    correlationId?: string;
    constructor(
        public errorCode: string,
        public errorMessage: string = AKA(errorCode),
        public subError: string = ""
    ) {
        super(`${errorCode}: ${errorMessage}`);
    }
    name = "AuthError";
}
export class InteractionRequiredAuthError extends AuthError {
    name = "InteractionRequiredAuthError";
}
export class ServerError extends AuthError {
    name = "ServerError";
}
```

Error codes are plain string literals at the throw site, not imported
constants. The `ClientAuthErrorCodes`-style namespace objects real exports
are generated in `@mini-msal/compat` by one shared 10-line packer — à-la-carte
users never ship them.

## 5. Telemetry event shapes as data tables, not instrumentation everywhere

Real telemetry works by wrapping hundreds of internal methods in measurement
helpers (`invokeAsync(fn, PerformanceEvents.X, …)`) across the whole
codebase; the ~40-field performance events your callback receives *emerge*
from whichever wrapped methods ran. That instrumentation is woven through
every module — you ship it even before counting the perf client itself.

mini pins the **output** instead of replicating the machinery: for each flow,
the exact set of sub-measurement keys real would produce is stored as a
compact space-separated string table, and the event is synthesized at emit
time. From `packages/browser/src/telemetry.ts`:

<!-- docs-check:skip illustrative internal excerpt (abridged tables) -->
```ts
// Suffix "!" = CallCount key only, "+" = DurationMs key only, else both.
const PKCE =
    "generateCodeChallengeFromVerifier generateCodeVerifier generatePkceCodes getRandomValues sha256Digest";
const RT =
    "acquireTokenByRefreshToken cacheManagerGetRefreshToken networkClientSendPostRequestAsync! refreshTokenClientAcquireToken …";
```

One string per flow beats a wrapper call at every function. And because
telemetry is itself a composable feature (+9.4 KB), apps that never call
`addPerformanceCallback` ship none of it — real apps carry the
instrumentation unconditionally.

## 6. A bridge page that only does its one job (0.6 KB vs 6.5 KB)

msal v5 popups and silent iframes finish on a page at your `redirectUri`
that relays the auth response back to the opener. Real's
`@azure/msal-browser/redirect-bridge` entry bundles library plumbing and
lands at 6.5 KB. The relay job is actually tiny — parse the response
fragment, scrub the URL, post the payload on a `BroadcastChannel`, close.
mini's bridge is those four steps and nothing else — 0.6 KB, honoring the
same `{v: 1, payload}` page contract as real v5's bridge:

```ts
export async function broadcastResponseToMainFrame(): Promise<void> {
    const payload = (location.hash || location.search).slice(1);
    const state = new URLSearchParams(payload).get("state") ?? "";
    const { id } = JSON.parse(atob(state.split("|")[0]));
    history.replaceState(null, "", location.origin + location.pathname);
    const channel = new BroadcastChannel(id);
    channel.postMessage({ v: 1, payload });
    channel.close();
    try {
        window.close();
    } catch {}
}
```

## 7. Small habits that add up

- **Every constant lives once.** Wire identity headers, the exported
  `version`, telemetry SKU strings — one `WIRE_ID` block feeds them all.
  No duplicated literals across modules, and policy changes are one-line.
- **Const objects, not TypeScript `enum`s** (zero `enum` keywords in the
  package): enums compile to double-keyed runtime objects; const maps and
  string-literal unions compile to nothing or to exactly what's used.
- **es2022 output, no polyfills or legacy branches**: modern browser APIs
  (`crypto.subtle`, `BroadcastChannel`, `fetch`) are used directly.
- **Bytes over abstraction as a code-review rule**: helpers earn their place
  by being smaller than the code they replace, not by being tidier. Size is
  measured and recorded on every change (`npm run measure`).

## 8. Packaging that lets bundlers do their rest

All three packages declare `"sideEffects": false` and expose each feature as
its own subpath export (`@mini-msal/browser/popup`, `…/telemetry`, …) mapping
to a separate compiled file — so unused features are unresolved imports, not
dead code the bundler must prove unreachable. Even inside `@mini-msal/compat`,
the extra export surface (constant namespaces, utility classes) tree-shakes
away for consumers who only construct `PublicClientApplication`. The
published `dist` gates are enforced by `npm run pack:check`, which builds a
throwaway consumer per profile and fails if any profile's size regresses.

## What was *not* traded away

Small did not mean approximate. The same repo pins, against real
msal-browser 5.16.0 running side by side: 122 conformance scenarios
(byte-level wire/cache/event/error/telemetry comparison, including the full
51-key export surface), a 25-check end-to-end suite incl. bidirectional
live-cache interop between the two libraries, and per-profile size gates.
Known non-goals (B2C, ADFS) are listed in
[UPGRADING.md](./UPGRADING.md#known-non-goals). Where real's behavior is
suboptimal, mini matches real anyway — parity beats improvement in a
drop-in.
