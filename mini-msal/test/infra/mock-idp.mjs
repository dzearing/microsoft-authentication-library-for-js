/**
 * Mock OIDC identity provider to E2E-test the auth-code+PKCE flow without
 * real AAD credentials. Realistic enough for both real MSAL and mini-msal:
 *   - discovery, authorize, token, end_session endpoints
 *   - server-side session: prompt=none fails with login_required until an
 *     interactive authorize has happened (exercises ssoSilent error paths)
 *   - echoes the authorize nonce into id_token claims, returns client_info
 *   - CORS incl. OPTIONS preflight (real MSAL sends x-client-* headers)
 *
 * Conformance-test extensions (all cleared by /reset):
 *   - GET /requests — JSON log of every /authorize|/token|/logout request
 *     (query params, POST body params, headers) for header/param assertions
 *   - GET /inject?endpoint=token&status=429&retryAfter=2&count=1 — queue
 *     error injections; also error=<oauth_error>&errorDescription=...,
 *     drop=1 (destroy socket), suberror=..., status=5xx
 *   - GET /config?expires_in=7200&refresh_in=60 — override token response
 *     fields (persist until /reset)
 *   - GET /session?active=1&user=0 — set IdP SSO session state directly
 */
import { createServer } from "node:https";
import { readFileSync } from "node:fs";

// https because real MSAL rejects http:// authorities (authority_uri_insecure);
// self-signed cert in .cert/ (generated via openssl, browser launched with
// ignoreHTTPSErrors)
const tls = {
    key: readFileSync(new URL("../../.cert/key.pem", import.meta.url)),
    cert: readFileSync(new URL("../../.cert/cert.pem", import.meta.url)),
};

const PORT = 4599;
const b64url = (s) =>
    Buffer.from(JSON.stringify(s))
        .toString("base64")
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");

// two identities so account-switching can be exercised end-to-end
const users = [
    {
        sub: "sub-123",
        oid: "oid-123",
        tid: "tenant-123",
        preferred_username: "ada@contoso.com",
        name: "Ada Lovelace",
    },
    {
        sub: "sub-456",
        oid: "oid-456",
        tid: "tenant-123",
        preferred_username: "grace@contoso.com",
        name: "Grace Hopper",
    },
];
let sessionUser = 0; // which identity the IdP session currently holds
let sessionActive = false; // "is the user signed in at the IdP"
let lastNonce;

// ---- conformance-test state (cleared by /reset) ----
let requestLog = []; // every authorize/token/logout request, in order
let injections = []; // queued error injections consumed by /token|/authorize
let tokenOverrides = {}; // extra/overridden fields merged into /token responses

function logRequest(entry) {
    requestLog.push({ seq: requestLog.length, ...entry });
}

function takeInjection(endpoint) {
    const i = injections.findIndex((inj) => inj.endpoint === endpoint);
    if (i === -1) return null;
    const inj = injections[i];
    if (--inj.count <= 0) injections.splice(i, 1);
    return inj;
}

function makeIdToken(idx) {
    const claims = {
        aud: "11111111-2222-3333-4444-555555555555",
        iss: `https://localhost:${PORT}/tenant/v2.0`,
        ...users[idx],
        nonce: lastNonce,
        exp: Math.floor(Date.now() / 1000) + 3600,
        iat: Math.floor(Date.now() / 1000),
    };
    return `${b64url({ alg: "none" })}.${b64url(claims)}.sig`;
}

const server = createServer(tls, (req, res) => {
    const url = new URL(req.url, `https://localhost:${PORT}`);
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Headers", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
    console.log(req.method, url.pathname + url.search.slice(0, 120));

    if (req.method === "OPTIONS") {
        res.writeHead(204);
        return res.end();
    }

    if (url.pathname.endsWith("/openid-configuration")) {
        logRequest({
            endpoint: "discovery",
            method: req.method,
            path: url.pathname,
            headers: req.headers,
        });
        res.setHeader("Content-Type", "application/json");
        res.end(
            JSON.stringify({
                issuer: `https://localhost:${PORT}/tenant/v2.0`,
                jwks_uri: `https://localhost:${PORT}/keys`,
                authorization_endpoint: `https://localhost:${PORT}/authorize`,
                token_endpoint: `https://localhost:${PORT}/token`,
                end_session_endpoint: `https://localhost:${PORT}/logout`,
            })
        );
    } else if (url.pathname === "/authorize") {
        const q = url.searchParams;
        logRequest({
            endpoint: "authorize",
            method: req.method,
            query: Object.fromEntries(q),
            headers: req.headers,
        });
        const inj = takeInjection("authorize");
        const respondAuthorize = () => {
            if (inj?.error) {
                const r = new URL(q.get("redirect_uri"));
                r.hash =
                    `error=${encodeURIComponent(inj.error)}` +
                    `&error_description=${encodeURIComponent(inj.errorDescription ?? inj.error)}` +
                    (inj.suberror ? `&suberror=${encodeURIComponent(inj.suberror)}` : "") +
                    `&state=${encodeURIComponent(q.get("state") ?? "")}`;
                res.writeHead(302, { Location: r.href });
                return res.end();
            }
            lastNonce = q.get("nonce") ?? lastNonce;
            const redirect = new URL(q.get("redirect_uri"));
            if (q.get("prompt") === "select_account") {
                // simulate the user picking the other account in the chooser
                sessionUser = (sessionUser + 1) % users.length;
            }
            const hinted = users.findIndex(
                (u) => u.preferred_username === q.get("login_hint")
            );
            const idx = hinted >= 0 ? hinted : sessionUser;
            if (q.get("prompt") === "none" && !sessionActive) {
                redirect.hash = `error=login_required&error_description=no+session&state=${encodeURIComponent(
                    q.get("state")
                )}`;
            } else {
                sessionActive = true;
                redirect.hash = `code=mock-code-${idx}&state=${encodeURIComponent(
                    q.get("state")
                )}`;
            }
            res.writeHead(302, { Location: redirect.href });
            res.end();
        };
        if (inj?.delay) {
            setTimeout(respondAuthorize, inj.delay);
        } else {
            respondAuthorize();
        }
    } else if (url.pathname === "/token") {
        let body = "";
        req.on("data", (c) => (body += c));
        req.on("end", () => {
            const params = new URLSearchParams(body);
            logRequest({
                endpoint: "token",
                method: req.method,
                ...(url.search && {
                    query: Object.fromEntries(url.searchParams),
                }),
                body: Object.fromEntries(params),
                headers: req.headers,
            });
            const inj = takeInjection("token");
            const respondToken = () => {
                if (inj && (inj.drop || inj.status || inj.error)) {
                    if (inj.drop) return req.socket.destroy();
                    const status = inj.status ?? 400;
                    const headers = { "Content-Type": "application/json" };
                    if (inj.retryAfter) headers["Retry-After"] = inj.retryAfter;
                    // keep CORS headers already set above
                    res.writeHead(status, headers);
                    return res.end(
                        JSON.stringify({
                            error:
                                inj.error ??
                                (status === 429
                                    ? "temporarily_throttled"
                                    : "server_error"),
                            error_description:
                                inj.errorDescription ?? `injected ${status}`,
                            ...(inj.suberror && { suberror: inj.suberror }),
                        })
                    );
                }
                finishTokenImpl();
            };
            if (inj?.delay) {
                setTimeout(respondToken, inj.delay);
            } else {
                respondToken();
            }
            function finishTokenImpl() {
            // identity travels with the code / refresh token
            const grantValue =
                params.get("code") ?? params.get("refresh_token") ?? "";
            const idx = Number(grantValue.slice(-1)) || 0;
            console.log("  grant:", params.get("grant_type"), "user:", idx);
            res.setHeader("Content-Type", "application/json");
            res.end(
                JSON.stringify({
                    token_type: "Bearer",
                    scope: "openid profile User.Read",
                    expires_in: 3600,
                    access_token: `mock-access-token-${idx}`,
                    refresh_token: `mock-rt-${idx}`,
                    id_token: makeIdToken(idx),
                    client_info: b64url({
                        uid: users[idx].oid,
                        utid: users[idx].tid,
                    }),
                    ...tokenOverrides,
                })
            );
            }
        });
    } else if (url.pathname === "/reset") {
        // test isolation: reset IdP session state between e2e app runs
        sessionUser = 0;
        sessionActive = false;
        requestLog = [];
        injections = [];
        tokenOverrides = {};
        res.end("reset");
    } else if (url.pathname === "/requests") {
        res.setHeader("Content-Type", "application/json");
        res.end(JSON.stringify(requestLog));
    } else if (url.pathname === "/inject") {
        const q = url.searchParams;
        injections.push({
            endpoint: q.get("endpoint") ?? "token",
            status: q.get("status") ? Number(q.get("status")) : undefined,
            retryAfter: q.get("retryAfter") ?? undefined,
            error: q.get("error") ?? undefined,
            errorDescription: q.get("errorDescription") ?? undefined,
            suberror: q.get("suberror") ?? undefined,
            drop: q.get("drop") === "1",
            delay: q.get("delay") ? Number(q.get("delay")) : undefined,
            count: Number(q.get("count") ?? 1),
        });
        res.end("injected");
    } else if (url.pathname === "/config") {
        for (const [k, v] of url.searchParams) {
            tokenOverrides[k] = /^-?\d+$/.test(v) ? Number(v) : v;
        }
        res.end("configured");
    } else if (url.pathname === "/session") {
        const q = url.searchParams;
        if (q.has("active")) sessionActive = q.get("active") === "1";
        if (q.has("user")) sessionUser = Number(q.get("user"));
        res.end("session set");
    } else if (url.pathname === "/logout") {
        logRequest({
            endpoint: "logout",
            method: req.method,
            query: Object.fromEntries(url.searchParams),
            headers: req.headers,
        });
        sessionActive = false;
        const back = url.searchParams.get("post_logout_redirect_uri");
        if (back) {
            // OIDC RP-initiated logout: echo state back to the RP (real
            // MSAL's popup logout completes via this echoed state)
            const backUrl = new URL(back);
            const state = url.searchParams.get("state");
            if (state) backUrl.searchParams.set("state", state);
            res.writeHead(302, { Location: backUrl.href });
            res.end();
        } else {
            res.end("logged out");
        }
    } else {
        res.writeHead(404);
        res.end();
    }
});
// long keep-alive: Chrome reuses pooled connections; node's 5s default causes
// races where a POST rides a connection the server just closed
server.keepAliveTimeout = 120_000;
server.headersTimeout = 125_000;
server.listen(PORT, () => console.log(`mock idp on :${PORT}`));
