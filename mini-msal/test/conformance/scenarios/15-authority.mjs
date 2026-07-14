/**
 * Area 15 — authority modes & discovery (C18): lazy endpoint discovery,
 * system.protocolMode OIDC discovery path, auth.authorityMetadata inline
 * metadata, knownAuthorities / untrusted-authority validation, hardcoded
 * known-cloud metadata, instance-aware cloud response fields.
 *
 * NOTE real 5.16 reads protocolMode from config.SYSTEM (not auth) — the
 * auth.protocolMode in stdConfig is inert on both stacks.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    idp,
    IDP,
    CLIENT_ID,
} from "../lib.mjs";

export const area = "authority";

const discoveryPaths = (reqs) =>
    reqs.filter((r) => r.endpoint === "discovery").map((r) => r.path);

/** stdConfig minus knownAuthorities/protocolMode: library-default (AAD)
 * authority validation against the mock host. */
const aadConfig = (ctx, authority = `${IDP}/tenant`) => ({
    auth: {
        clientId: CLIENT_ID,
        authority,
        redirectUri: ctx.harnessUrl,
        postLogoutRedirectUri: ctx.harnessUrl,
    },
    cache: { cacheLocation: "sessionStorage" },
});

/** Block + record every request the page makes to login.microsoftonline.com
 * (real's AAD instance-discovery host) so trust probes stay deterministic. */
async function armAadProbeRecorder(ctx, probes) {
    await ctx.context.route("https://login.microsoftonline.com/**", (route) => {
        const u = new URL(route.request().url());
        probes.push({
            url: u.origin + u.pathname,
            query: Object.fromEntries(u.searchParams),
        });
        return route.abort();
    });
}

const loginPopupFn = async (popupUrl) => {
    const r = await globalThis.__msal.loginPopup({
        scopes: ["User.Read"],
        redirectUri: popupUrl,
    });
    return { username: r.account?.username ?? null };
};

export const scenarios = [
    {
        id: "authority.lazy-discovery-default-path",
        note: "initialize() issues no discovery request; first token flow lazily fetches openid-configuration (AAD-mode path includes /v2.0/)",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const afterInit = discoveryPaths(await idp.requests());
            const login = await tryEval(ctx, loginPopupFn, ctx.popupUrl);
            const afterLogin = discoveryPaths(await idp.requests());
            return { afterInit, afterLogin, login };
        },
    },
    {
        id: "authority.oidc-discovery-endpoint-path",
        note: "system.protocolMode OIDC + non-Microsoft host: discovery URL omits /v2.0/",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(
                ctx,
                stdConfig(ctx, { system: { protocolMode: "OIDC" } })
            );
            const afterInit = discoveryPaths(await idp.requests());
            const login = await tryEval(ctx, loginPopupFn, ctx.popupUrl);
            const afterLogin = discoveryPaths(await idp.requests());
            return { afterInit, afterLogin, login };
        },
    },
    {
        id: "authority.authority-metadata-config",
        note: "auth.authorityMetadata inline endpoint JSON skips the openid-configuration request entirely",
        async run(ctx) {
            const md = {
                issuer: `${IDP}/tenant/v2.0`,
                jwks_uri: `${IDP}/keys`,
                authorization_endpoint: `${IDP}/authorize`,
                token_endpoint: `${IDP}/token`,
                end_session_endpoint: `${IDP}/logout`,
            };
            await gotoHarness(ctx);
            await create(
                ctx,
                stdConfig(ctx, {
                    auth: { authorityMetadata: JSON.stringify(md) },
                })
            );
            const login = await tryEval(ctx, loginPopupFn, ctx.popupUrl);
            const endpoints = (await idp.requests()).map((r) => r.endpoint);
            return { login, endpoints };
        },
    },
    {
        id: "authority.known-authorities-validation",
        note: "default (AAD) mode, host not known/hardcoded: instance-discovery GET to login.microsoftonline.com; blocked probe => ClientConfigurationError untrusted_authority, IdP never contacted for authorize",
        async run(ctx) {
            const probes = [];
            await armAadProbeRecorder(ctx, probes);
            await gotoHarness(ctx);
            await create(ctx, aadConfig(ctx));
            const err = await tryEval(ctx, () =>
                globalThis.__msal.ssoSilent({
                    scopes: ["User.Read"],
                    loginHint: "ada@contoso.com",
                })
            );
            const reqs = await idp.requests();
            return {
                err,
                probes,
                idpEndpoints: reqs.map((r) => r.endpoint),
            };
        },
    },
    {
        id: "authority.hardcoded-cloud-metadata",
        note: "login.microsoftonline.com authority: endpoints come from hardcoded metadata — zero network to the cloud host (blocked anyway); logout URL built from it",
        async run(ctx) {
            const probes = [];
            await armAadProbeRecorder(ctx, probes);
            await gotoHarness(ctx);
            await create(
                ctx,
                aadConfig(ctx, "https://login.microsoftonline.com/common"),
                { onRedirectNavigate: "record-cancel" }
            );
            const logout = await tryEval(ctx, () =>
                globalThis.__msal.logoutRedirect()
            );
            const navCalls = await tryEval(
                ctx,
                () => globalThis.__navCalls
            );
            return { logout, navCalls, probes };
        },
    },
    {
        id: "authority.instance-aware-cloud-instance",
        note: "instance_aware EQP: authorize response carries cloud_instance_host_name/cloud_graph_host_name/msgraph_host — cached on the account entity, surfaced on interactive AND cache-hit results",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            const login = await tryEval(
                ctx,
                async (popupUrl) => {
                    const r = await globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                        extraQueryParameters: { instance_aware: "true" },
                    });
                    return {
                        username: r.account?.username ?? null,
                        cloudGraphHostName: r.cloudGraphHostName,
                        msGraphHost: r.msGraphHost,
                    };
                },
                ctx.popupUrl
            );
            const silent = await tryEval(ctx, async () => {
                const account = globalThis.__msal.getAllAccounts()[0];
                const r = await globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account,
                });
                return {
                    fromCache: r.fromCache,
                    cloudGraphHostName: r.cloudGraphHostName,
                    msGraphHost: r.msGraphHost,
                };
            });
            const entity = await ctx.page.evaluate(() => {
                const keysKey = Object.keys(sessionStorage).find((k) =>
                    k.endsWith(".account.keys")
                );
                const keys = keysKey
                    ? JSON.parse(sessionStorage.getItem(keysKey))
                    : [];
                const e = keys[0]
                    ? JSON.parse(sessionStorage.getItem(keys[0]))
                    : null;
                return (
                    e && {
                        cloudGraphHostName: e.cloudGraphHostName ?? null,
                        msGraphHost: e.msGraphHost ?? null,
                    }
                );
            });
            const endpoints = (await idp.requests()).map((r) => ({
                endpoint: r.endpoint,
                ...(r.endpoint === "authorize" && {
                    instanceAware: r.query.instance_aware ?? null,
                }),
            }));
            return { login, silent, entity, endpoints };
        },
    },
];
