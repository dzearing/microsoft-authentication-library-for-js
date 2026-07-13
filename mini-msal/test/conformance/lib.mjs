/**
 * Shared helpers for conformance scenarios: scenario context, page/eval
 * utilities, IdP control, and observation normalization.
 */

export const BASE = "http://localhost:4173";
export const IDP = "https://localhost:4599";

export const CLIENT_ID = "11111111-2222-3333-4444-555555555555";

/** Standard harness config (JSON-serializable — no functions). */
export function stdConfig(ctx, overrides = {}) {
    const { auth = {}, cache = {}, system = {}, telemetry, experimental } =
        overrides;
    return {
        auth: {
            clientId: CLIENT_ID,
            authority: `${IDP}/tenant`,
            knownAuthorities: ["localhost:4599"],
            protocolMode: "OIDC",
            redirectUri: ctx.harnessUrl,
            postLogoutRedirectUri: ctx.harnessUrl,
            ...auth,
        },
        cache: { cacheLocation: "sessionStorage", ...cache },
        system: { ...system },
        ...(telemetry && { telemetry }),
        ...(experimental && { experimental }),
    };
}

// ---- IdP control ------------------------------------------------------------

export const idp = {
    reset: () => fetch(`${IDP}/reset`).then((r) => r.text()),
    requests: () => fetch(`${IDP}/requests`).then((r) => r.json()),
    inject: (params) =>
        fetch(`${IDP}/inject?` + new URLSearchParams(params)).then((r) =>
            r.text()
        ),
    config: (params) =>
        fetch(`${IDP}/config?` + new URLSearchParams(params)).then((r) =>
            r.text()
        ),
    session: (params) =>
        fetch(`${IDP}/session?` + new URLSearchParams(params)).then((r) =>
            r.text()
        ),
};

// ---- page helpers -----------------------------------------------------------

/** Navigate to the harness page (fresh load). */
export async function gotoHarness(ctx, hash = "") {
    await ctx.page.goto(ctx.harnessUrl + hash);
    await ctx.page.waitForFunction(() => !!globalThis.__create);
}

/** Create (and by default initialize) a client in the page. */
export function create(ctx, config, opts = {}) {
    return ctx.page.evaluate(
        ([c, o]) => globalThis.__create(c, o),
        [config, opts]
    );
}

/**
 * Run an async function in the page with uniform error capture.
 * Returns { ok: <value> } or { err: <serialized error> }.
 */
export function tryEval(ctx, fn, arg) {
    return ctx.page.evaluate(
        async ([src, a]) => {
            try {
                const f = (0, eval)(`(${src})`);
                return { ok: (await f(a)) ?? null };
            } catch (e) {
                return { err: globalThis.__serializeError(e) };
            }
        },
        [fn.toString(), arg]
    );
}

/** Digest an AuthenticationResult in-page into stable fields. */
export const RESULT_DIGEST = `(r) => r && {
    accessToken: r.accessToken,
    idToken: typeof r.idToken === "string" && r.idToken.length > 0 ? "<jwt>" : r.idToken,
    tokenType: r.tokenType ?? null,
    scopes: [...(r.scopes ?? [])].sort(),
    account: r.account ? {
        username: r.account.username,
        homeAccountId: r.account.homeAccountId,
        tenantId: r.account.tenantId ?? r.account.realm ?? null,
    } : null,
    fromCache: r.fromCache ?? null,
    expiresOnType: r.expiresOn instanceof Date ? "Date" : typeof r.expiresOn,
    expiresInFuture: r.expiresOn ? +new Date(r.expiresOn) > Date.now() : null,
    state: r.state === undefined ? null : r.state,
    authority: r.authority ?? null,
    correlationIdPresent: typeof r.correlationId === "string" && r.correlationId.length > 0,
    fromPlatformBroker: r.fromPlatformBroker ?? null,
}`;

/** tryEval wrapper that digests an AuthenticationResult. */
export function tryResult(ctx, fn, arg) {
    return ctx.page.evaluate(
        async ([src, digestSrc, a]) => {
            try {
                const f = (0, eval)(`(${src})`);
                const digest = (0, eval)(digestSrc);
                const r = await f(a);
                return { ok: digest(r) };
            } catch (e) {
                return { err: globalThis.__serializeError(e) };
            }
        },
        [fn.toString(), RESULT_DIGEST, arg]
    );
}

/**
 * Fire an async page function WITHOUT awaiting it; result lands in
 * globalThis.__pending as {ok}/{err} (digested). Await with pending(ctx).
 */
export function fire(ctx, fn, arg) {
    return ctx.page.evaluate(
        ([src, a]) => {
            const f = (0, eval)(`(${src})`);
            globalThis.__pending = Promise.resolve()
                .then(() => f(a))
                .then(
                    (v) => ({
                        ok: v
                            ? {
                                  username: v.account?.username ?? null,
                                  accessToken: v.accessToken ?? null,
                              }
                            : null,
                    }),
                    (e) => ({ err: globalThis.__serializeError(e) })
                );
            return true;
        },
        [fn.toString(), arg]
    );
}

export function pending(ctx) {
    return ctx.page.evaluate(() => globalThis.__pending);
}

/** Captured events/perf/logs from the harness. */
export function capture(ctx) {
    return ctx.page.evaluate(() => globalThis.__cap);
}

/**
 * Wrap window.open to record {url, name, features, sync} per call. Calls are
 * persisted in sessionStorage so they survive a main-window navigation
 * (logoutPopup mainWindowRedirectUri). `sync` is true when the call happened
 * while globalThis.__inApiCall was set — scenarios flag the synchronous span
 * of the API call with it to observe popup-blocker-visible timing.
 */
export function armOpenRecorder(ctx) {
    return ctx.page.evaluate(() => {
        sessionStorage.setItem("__openCalls", "[]");
        const orig = window.open.bind(window);
        window.open = (url, name, features) => {
            const calls = JSON.parse(sessionStorage.getItem("__openCalls"));
            calls.push({
                url: String(url ?? ""),
                name: String(name ?? ""),
                features: String(features ?? ""),
                sync: !!globalThis.__inApiCall,
            });
            sessionStorage.setItem("__openCalls", JSON.stringify(calls));
            return orig(url, name, features);
        };
    });
}

/**
 * Read recorded window.open calls. URLs other than about:blank are digested
 * to origin+path plus SORTED QUERY KEY NAMES (values are volatile).
 */
export async function openCalls(ctx) {
    const calls = JSON.parse(
        await ctx.page.evaluate(
            () => sessionStorage.getItem("__openCalls") ?? "[]"
        )
    );
    return calls.map((c) => {
        let url = c.url;
        if (url && url !== "about:blank") {
            const u = new URL(url);
            url = `${u.origin}${u.pathname}?${[...u.searchParams.keys()]
                .sort()
                .join(",")}`;
        }
        return { ...c, url };
    });
}

/** Sorted sessionStorage + localStorage dump. */
export function storageDump(ctx) {
    return ctx.page.evaluate(() => {
        const dump = (s) => {
            const out = {};
            for (const k of [...Array(s.length).keys()]
                .map((i) => s.key(i))
                .sort()) {
                const v = s.getItem(k);
                try {
                    out[k] = JSON.parse(v);
                } catch {
                    out[k] = v;
                }
            }
            return out;
        };
        return {
            sessionStorage: dump(sessionStorage),
            localStorage: dump(localStorage),
        };
    });
}

/**
 * Perform a full interactive login via redirect: fires loginRedirect, waits
 * for the round trip, re-creates the client, returns handleRedirectPromise's
 * digested result. Leaves the page signed in.
 */
export async function loginViaRedirect(ctx, config, request = {}) {
    await create(ctx, config);
    await ctx.page.evaluate((req) => {
        globalThis.__msal
            .loginRedirect(req)
            .catch((e) => (globalThis.__redirectErr = globalThis.__serializeError(e)));
    }, { scopes: ["User.Read"], ...request });
    await ctx.page.waitForURL((u) => /[#&](code|error)=/.test(u.hash), {
        timeout: 15000,
    });
    await ctx.page.waitForFunction(() => !!globalThis.__create);
    await create(ctx, config);
    return tryResult(ctx, () => globalThis.__msal.handleRedirectPromise());
}

/** Login via popup (no page reload). Leaves the page signed in. */
export async function loginViaPopup(ctx, config, request = {}) {
    await create(ctx, config);
    return tryResult(
        ctx,
        async (req) => globalThis.__msal.loginPopup(req),
        {
            scopes: ["User.Read"],
            redirectUri: ctx.harnessUrl + "popup.html",
            ...request,
        }
    );
}

/**
 * Patch every cached AccessToken entity (e.g. force expiry). Returns count.
 * Fields are merged into the msal.3 entity JSON; values must be strings to
 * match the schema (epoch-seconds as strings).
 */
export function patchAccessTokens(ctx, fields, storage = "sessionStorage") {
    return ctx.page.evaluate(
        ([f, storageName]) => {
            const store = window[storageName];
            const keysKey = Object.keys(store).find((k) =>
                k.startsWith("msal.3.token.keys")
            );
            if (!keysKey) return 0;
            const keys = JSON.parse(store.getItem(keysKey));
            for (const k of keys.accessToken) {
                const t = JSON.parse(store.getItem(k));
                Object.assign(t, f);
                store.setItem(k, JSON.stringify(t));
            }
            return keys.accessToken.length;
        },
        [fields, storage]
    );
}

/** Remove all cached credentials of a type ("refreshToken" | "accessToken"). */
export function removeCreds(ctx, type, storage = "sessionStorage") {
    return ctx.page.evaluate(
        ([t, storageName]) => {
            const store = window[storageName];
            const keysKey = Object.keys(store).find((k) =>
                k.startsWith("msal.3.token.keys")
            );
            if (!keysKey) return 0;
            const keys = JSON.parse(store.getItem(keysKey));
            const removed = keys[t].length;
            for (const k of keys[t]) store.removeItem(k);
            keys[t] = [];
            store.setItem(keysKey, JSON.stringify(keys));
            return removed;
        },
        [type, storage]
    );
}

// ---- normalization ----------------------------------------------------------

const GUID_RE =
    /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;
const JWT_RE =
    /\beyJ[\w-]+\.[\w-]+\.[\w-]*/g;
// 43-char base64url blobs: PKCE verifiers/challenges, states, nonces
const B64_43_RE = /\b[\w-]{43}\b/g;

/** Keys whose values are inherently per-run; replaced with "<dynamic>". */
const VOLATILE_KEYS = new Set([
    "state",
    "nonce",
    "code_challenge",
    "code_verifier",
    "client-request-id",
    "client_info", // stable in mock but b64 blob — compare decoded elsewhere
    "correlationId",
    "correlation_id",
    "requestId",
    "sendTime",
    "cachedAt",
    "expiresOn",
    "extendedExpiresOn",
    "refreshOn",
    "x-ms-request-id",
    "date",
    "host",
    "user-agent",
    "sec-ch-ua",
    "sec-ch-ua-platform",
    "sec-ch-ua-mobile",
    "content-length",
    "accept-encoding",
    "accept-language",
    "connection",
    "traceparent",
    // real v5 encrypts localStorage cache entries ({data: <ciphertext>})
    // with a per-session key — ciphertext differs every run
    "data",
]);

function scrubString(s) {
    return s
        .replace(/conformance-react-(real|mini)/g, "conformance-react-app")
        .replace(/conformance-(real|mini)/g, "conformance-app")
        .replace(GUID_RE, "<guid>")
        .replace(JWT_RE, "<jwt>")
        .replace(B64_43_RE, "<b64-43>");
}

/**
 * Digest the mock-IdP request log into stable observations: full query/body
 * params (normalize() handles volatile values) and a whitelist of headers
 * that carry protocol meaning (x-client-*, client-request-id, origin, ...).
 */
export function digestIdpLog(reqs) {
    return reqs.map((r) => {
        const headers = {};
        for (const [k, v] of Object.entries(r.headers ?? {})) {
            if (/^(x-client-|x-ms-|client-request-id|origin$|content-type$)/.test(k)) {
                headers[k] = v;
            }
        }
        return {
            endpoint: r.endpoint,
            ...(r.query && { query: r.query }),
            ...(r.body && { body: r.body }),
            headers,
        };
    });
}

/**
 * Deep-normalize an observation: volatile keys -> "<dynamic>", GUID/JWT/epoch
 * value scrubbing, recursive key sort. Storage keys containing GUIDs etc. are
 * scrubbed too so cache-shape comparisons stay stable.
 */
export function normalize(value, keyHint = "") {
    if (value === undefined) return null;
    if (value === null) return null;
    if (typeof value === "number") {
        // epoch timestamps (seconds or ms) are per-run
        if (value > 1.5e9 && value < 2.5e9) return "<epoch-s>";
        if (value > 1.5e12) return "<epoch-ms>";
        return value;
    }
    if (typeof value === "string") {
        if (VOLATILE_KEYS.has(keyHint)) return "<dynamic>";
        if (/^\d{10}$/.test(value)) return "<epoch-s>";
        if (/^\d{13}$/.test(value)) return "<epoch-ms>";
        return scrubString(value);
    }
    if (Array.isArray(value)) {
        return value.map((v) => normalize(v));
    }
    if (typeof value === "object") {
        const out = {};
        for (const k of Object.keys(value).sort()) {
            if (VOLATILE_KEYS.has(k)) {
                out[scrubString(k)] =
                    value[k] === undefined || value[k] === null
                        ? null
                        : "<dynamic>";
            } else {
                out[scrubString(k)] = normalize(value[k], k);
            }
        }
        return out;
    }
    return value;
}

/** Deep diff of two normalized values; returns list of {path, real, mini}. */
export function diff(real, mini, path = "", out = []) {
    if (JSON.stringify(real) === JSON.stringify(mini)) return out;
    const isObj = (v) => v && typeof v === "object" && !Array.isArray(v);
    if (isObj(real) && isObj(mini)) {
        for (const k of new Set([...Object.keys(real), ...Object.keys(mini)])) {
            diff(
                k in real ? real[k] : "<missing>",
                k in mini ? mini[k] : "<missing>",
                path ? `${path}.${k}` : k,
                out
            );
        }
        return out;
    }
    if (Array.isArray(real) && Array.isArray(mini)) {
        const n = Math.max(real.length, mini.length);
        for (let i = 0; i < n; i++) {
            diff(
                i < real.length ? real[i] : "<missing>",
                i < mini.length ? mini[i] : "<missing>",
                `${path}[${i}]`,
                out
            );
        }
        return out;
    }
    const trunc = (v) => {
        const s = JSON.stringify(v);
        return s && s.length > 300 ? s.slice(0, 300) + "…" : s;
    };
    out.push({ path, real: trunc(real), mini: trunc(mini) });
    return out;
}
