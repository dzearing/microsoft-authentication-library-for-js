/**
 * Area 8 — request passthrough: prompt, loginHint, sid, domainHint,
 * extraQueryParameters, claims/CAE, custom state, per-request redirectUri,
 * authority override, scopes normalization.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryResult,
    loginViaPopup,
    digestIdpLog,
    idp,
} from "../lib.mjs";

export const area = "params";

const authorizeReqs = (reqs) => reqs.filter((r) => r.endpoint === "authorize");
const tokenReqs = (reqs) => reqs.filter((r) => r.endpoint === "token");

export const scenarios = [
    {
        id: "params.prompt-passthrough",
        note: "prompt values (login/consent/select_account) forwarded to /authorize",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await create(ctx, config);
            const out = {};
            for (const prompt of ["login", "consent", "select_account"]) {
                await idp.reset();
                const r = await tryResult(
                    ctx,
                    ([p, popupUrl]) =>
                        globalThis.__msal.loginPopup({
                            scopes: ["User.Read"],
                            prompt: p,
                            redirectUri: popupUrl,
                        }),
                    [prompt, ctx.popupUrl]
                );
                const reqs = await idp.requests();
                out[prompt] = {
                    promptSent: authorizeReqs(reqs)[0]?.query?.prompt ?? null,
                    user: r.ok?.account?.username ?? r.err,
                };
            }
            return out;
        },
    },
    {
        id: "params.login-hint-domain-hint",
        note: "loginHint + domainHint forwarded to /authorize",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        loginHint: "grace@contoso.com",
                        domainHint: "contoso.com",
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const q = authorizeReqs(await idp.requests())[0]?.query ?? {};
            return {
                login_hint: q.login_hint ?? null,
                domain_hint: q.domain_hint ?? null,
                // real also derives an X-AnchorMailbox routing hint
                anchorMailbox: q["X-AnchorMailbox"] ?? null,
            };
        },
    },
    {
        id: "params.sid-passthrough",
        note: "sid forwarded on silent (prompt=none) requests",
        async run(ctx) {
            await idp.session({ active: "1" });
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.ssoSilent({
                        scopes: ["User.Read"],
                        sid: "session-id-abc123",
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const q = authorizeReqs(await idp.requests())[0]?.query ?? {};
            return {
                sid: q.sid ?? null,
                prompt: q.prompt ?? null,
                login_hint: q.login_hint ?? null,
            };
        },
    },
    {
        id: "params.extra-query-parameters",
        note: "extraQueryParameters on authorize; tokenQueryParameters on token endpoint",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        extraQueryParameters: {
                            slice: "testslice",
                            dc: "ESTS-PUB",
                        },
                        tokenQueryParameters: { tokenslice: "tq1" },
                    }),
                ctx.popupUrl
            );
            const reqs = await idp.requests();
            const q = authorizeReqs(reqs)[0]?.query ?? {};
            const t = tokenReqs(reqs)[0] ?? {};
            return {
                authorizeExtras: { slice: q.slice ?? null, dc: q.dc ?? null },
                tokenQuery: t.query ?? null,
                tokenBodyHasExtras: {
                    slice: t.body?.slice ?? null,
                    tokenslice: t.body?.tokenslice ?? null,
                },
            };
        },
    },
    {
        id: "params.claims-and-cae",
        note: "clientCapabilities cp1 (CAE) merges xms_cc into claims on authorize + token",
        async run(ctx) {
            await gotoHarness(ctx);
            const config = stdConfig(ctx, {
                auth: { clientCapabilities: ["cp1"] },
            });
            await create(ctx, config);
            await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        claims: JSON.stringify({
                            access_token: {
                                nbf: { essential: true, value: "1701000000" },
                            },
                        }),
                    }),
                ctx.popupUrl
            );
            const reqs = await idp.requests();
            const authorizeClaims =
                authorizeReqs(reqs)[0]?.query?.claims ?? null;
            const tokenClaims = tokenReqs(reqs)[0]?.body?.claims ?? null;
            const parse = (c) => {
                try {
                    return JSON.parse(c);
                } catch {
                    return c;
                }
            };
            return {
                authorizeClaims: parse(authorizeClaims),
                tokenClaims: parse(tokenClaims),
            };
        },
    },
    {
        id: "params.custom-state",
        note: "custom request state is round-tripped and surfaced on the result",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const result = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        state: "app-state-123",
                    }),
                ctx.popupUrl
            );
            const rawState =
                authorizeReqs(await idp.requests())[0]?.query?.state ?? "";
            return {
                customStateEcho: result.ok?.state ?? result.err ?? null,
                wireStateContainsCustom: rawState.includes("app-state-123"),
                wireStateEqualsCustom: rawState === "app-state-123",
            };
        },
    },
    {
        id: "params.per-request-redirect-uri",
        note: "request.redirectUri overrides the configured redirect for popup flows",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const alt = ctx.harnessUrl + "popup2.html";
            await tryResult(
                ctx,
                (redirectUri) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri,
                    }),
                alt
            );
            const reqs = await idp.requests();
            return {
                authorizeRedirectUri:
                    authorizeReqs(reqs)[0]?.query?.redirect_uri ?? null,
                tokenRedirectUri:
                    tokenReqs(reqs)[0]?.body?.redirect_uri ?? null,
            };
        },
    },
    {
        id: "params.authority-override-per-request",
        note: "per-request authority triggers discovery for that authority; result/cache realm",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            const login = await loginViaPopup(ctx, config);
            await idp.reset();
            await idp.session({ active: "1" });
            const result = await tryResult(
                ctx,
                (popupUrl) => {
                    const account = globalThis.__msal.getAllAccounts()[0];
                    return globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account,
                        authority: "https://localhost:4599/other-tenant",
                        forceRefresh: true,
                        redirectUri: popupUrl,
                    });
                },
                ctx.popupUrl
            );
            const reqs = await idp.requests();
            return {
                login: { ok: !!login.ok },
                result,
                discoveryPaths: reqs
                    .filter((r) => r.endpoint === "discovery")
                    .map((r) => r.path),
            };
        },
    },
    {
        id: "params.scopes-normalization",
        note: "scope dedupe/casing and default OIDC scopes on empty request",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const dupes = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read", "USER.READ", "openid", "User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const dupesScopeParam =
                authorizeReqs(await idp.requests())[0]?.query?.scope ?? null;
            await idp.reset();
            const empty = await tryResult(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: [],
                        redirectUri: popupUrl,
                        prompt: "login",
                    }),
                ctx.popupUrl
            );
            const emptyScopeParam =
                authorizeReqs(await idp.requests())[0]?.query?.scope ?? null;
            return {
                dupes: { result: dupes, scopeParam: dupesScopeParam },
                empty: { result: empty, scopeParam: emptyScopeParam },
            };
        },
    },
];
