/**
 * Area 12 — React bindings parity (msal-react vs @mini-msal/react), driven
 * against the react harness pages (conformance-react-<target>/). Findings:
 * msalprovider-initialize, inprogress-interaction-statuses,
 * usemsalauthentication-acquiretoken, template-render-prop-children-and-
 * account-props, msal-authentication-template-error-contract,
 * account-identifier-matching.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    loginViaPopup,
    idp,
} from "../lib.mjs";

export const area = "react";

/** scenario ctx retargeted at the react harness pages */
const rctx = (ctx) => ({
    ...ctx,
    harnessUrl: ctx.reactHarnessUrl,
    popupUrl: ctx.reactPopupUrl,
});

function mount(ctx, fixture, props = {}) {
    return ctx.page.evaluate(
        ([f, p]) => globalThis.__mount(f, p),
        [fixture, props]
    );
}

const rendered = (ctx) =>
    ctx.page.evaluate(
        () => document.getElementById("root")?.textContent ?? ""
    );

/**
 * Wait until the root's textContent includes substr; observations record the
 * boolean so a stack that never reaches the state produces a stable diff
 * instead of a scenario error.
 */
async function waitRendered(ctx, substr, timeout = 15000) {
    return ctx.page
        .waitForFunction(
            (s) =>
                (
                    document.getElementById("root")?.textContent ?? ""
                ).includes(s),
            substr,
            { timeout }
        )
        .then(
            () => true,
            () => false
        );
}

export const scenarios = [
    {
        id: "react.provider-initializes-instance",
        note: "MsalProvider initialize()s an un-initialized instance and processes the redirect hash (finding msalprovider-initialize)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            // outbound leg with an initialized instance
            await create(ctx, config);
            await ctx.page.evaluate(() => {
                globalThis.__msal
                    .loginRedirect({ scopes: ["User.Read"] })
                    .catch(
                        (e) =>
                            (globalThis.__redirectErr =
                                globalThis.__serializeError(e))
                    );
            });
            await ctx.page.waitForURL((u) => /[#&](code|error)=/.test(u.hash), {
                timeout: 15000,
            });
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            await idp.reset(); // only the return-leg requests matter
            // return leg: hand the provider a NON-initialized instance
            await create(ctx, config, { noInit: true });
            await mount(ctx, "none");
            const sawSignedIn = await waitRendered(ctx, "accounts:1");
            const reqs = await idp.requests();
            return {
                sawSignedIn,
                rendered: await rendered(ctx),
                inProgressLog: await ctx.page.evaluate(
                    () => globalThis.__renderLog
                ),
                idpEndpoints: reqs.map((r) => r.endpoint),
                accounts: await tryEval(ctx, () =>
                    globalThis.__msal.getAllAccounts().map((a) => a.username)
                ),
                hashAfter: new URL(ctx.page.url()).hash,
            };
        },
    },
    {
        id: "react.inprogress-status-sequence",
        note: "inProgress transitions across popup login, acquireTokenPopup and logoutPopup (finding inprogress-interaction-statuses)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            await mount(ctx, "none");
            await waitRendered(ctx, "status:none");
            const login = await tryResult(
                ctx,
                async (popupUrl) =>
                    globalThis.__msal.loginPopup({
                        scopes: ["User.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const sawLoggedIn = await waitRendered(ctx, "accounts:1");
            const at = await tryResult(
                ctx,
                async (popupUrl) =>
                    globalThis.__msal.acquireTokenPopup({
                        scopes: ["Mail.Read"],
                        redirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            await tryEval(
                ctx,
                // v5 popup completion is bridge-based: the logout popup must
                // land on the bridge page
                (popupUrl) =>
                    globalThis.__msal.logoutPopup({
                        account: globalThis.__msal.getAllAccounts()[0],
                        postLogoutRedirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const sawLoggedOut = await waitRendered(ctx, "accounts:0");
            return {
                loginOk: !!login.ok,
                acquireOk: !!at.ok,
                sawLoggedIn,
                sawLoggedOut,
                inProgressLog: await ctx.page.evaluate(
                    () => globalThis.__renderLog
                ),
                rendered: await rendered(ctx),
            };
        },
    },
    {
        id: "react.use-msal-authentication",
        note: "signed-in user: auto silent acquire + acquireToken callback + result reset after logout (finding usemsalauthentication-acquiretoken)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx), {
                redirectUri: ctx.popupUrl,
            });
            // the hook resolves its account via getActiveAccount(); neither
            // stack auto-sets one after loginPopup
            await tryEval(ctx, () =>
                globalThis.__msal.setActiveAccount(
                    globalThis.__msal.getAllAccounts()[0]
                )
            );
            await mount(ctx, "authHook", {
                interactionType: "silent",
                request: { scopes: ["User.Read"] },
            });
            const sawResult = await waitRendered(ctx, "result=yes");
            const hookSignedIn = await ctx.page.evaluate(
                () => globalThis.__hook
            );
            await tryEval(
                ctx,
                (popupUrl) =>
                    globalThis.__msal.logoutPopup({
                        account: globalThis.__msal.getAllAccounts()[0],
                        postLogoutRedirectUri: popupUrl,
                    }),
                ctx.popupUrl
            );
            const sawReset = await waitRendered(ctx, "result=no");
            return {
                sawResult,
                hookSignedIn,
                sawReset,
                hookAfterLogout: await ctx.page.evaluate(
                    () => globalThis.__hook
                ),
                rendered: await rendered(ctx),
            };
        },
    },
    {
        id: "react.template-render-props-identifiers",
        note: "templates: function-as-children + username/homeAccountId scoping props (finding template-render-prop-children-and-account-props)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            await loginViaPopup(ctx, stdConfig(ctx), {
                redirectUri: ctx.popupUrl,
            });
            const home = await ctx.page.evaluate(
                () => globalThis.__msal.getAllAccounts()[0].homeAccountId
            );
            const out = {};
            await mount(ctx, "templates", { fnChildren: true });
            await waitRendered(ctx, "status:none");
            out.fnChildren = await rendered(ctx);
            await mount(ctx, "templates", {
                authProps: { homeAccountId: "no-such-account-id" },
            });
            await waitRendered(ctx, "status:none");
            out.wrongId = await rendered(ctx);
            await mount(ctx, "templates", {
                authProps: { homeAccountId: home.toUpperCase() },
            });
            await waitRendered(ctx, "status:none");
            out.matchingIdUppercased = await rendered(ctx);
            return out;
        },
    },
    {
        id: "react.auth-template-error-contract",
        note: "MsalAuthenticationTemplate: ErrorComponent gets the full auth result; without one the error is thrown to a boundary (finding msal-authentication-template-error-contract)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx));
            // cold IdP: the hook's automatic silent login fails login_required
            // (redirectUri = bridge page so the iframe response is relayed)
            const request = { scopes: ["User.Read"], redirectUri: ctx.popupUrl };
            await mount(ctx, "errTemplate", {
                interactionType: "silent",
                request,
                withErrorComponent: true,
                withLoading: true,
            });
            const sawErrorComp = await waitRendered(ctx, "|error-comp:");
            const withErrorComponent = await rendered(ctx);
            await mount(ctx, "errTemplate", {
                interactionType: "silent",
                request,
                withErrorComponent: false,
                withLoading: true,
            });
            const sawBoundary = await waitRendered(ctx, "|boundary:");
            return {
                sawErrorComp,
                withErrorComponent,
                sawBoundary,
                withoutErrorComponent: await rendered(ctx),
            };
        },
    },
    {
        id: "react.account-identifier-matching",
        note: "useIsAuthenticated(identifiers) + useAccount casing/empty-filter/startup semantics (finding account-identifier-matching)",
        async run(rawCtx) {
            const ctx = rctx(rawCtx);
            await gotoHarness(ctx);
            const config = stdConfig(ctx);
            await loginViaPopup(ctx, config, { redirectUri: ctx.popupUrl });
            await loginViaPopup(ctx, config, {
                redirectUri: ctx.popupUrl,
                prompt: "select_account",
            });
            const ids = await ctx.page.evaluate(() => {
                const all = globalThis.__msal.getAllAccounts();
                const ada = all.find((a) => a.username === "ada@contoso.com");
                const grace = all.find(
                    (a) => a.username === "grace@contoso.com"
                );
                globalThis.__msal.setActiveAccount(grace);
                return {
                    order: all.map((a) => a.username),
                    adaHome: ada.homeAccountId,
                };
            });
            const out = { accountOrder: ids.order };
            // casing: identifiers uppercased must still match (real lowercases)
            await mount(ctx, "probe", {
                matchAccount: { homeAccountId: ids.adaHome.toUpperCase() },
                filter: { homeAccountId: ids.adaHome.toUpperCase() },
            });
            await waitRendered(ctx, "status:none");
            out.uppercasedId = await rendered(ctx);
            // startup flash: first probe render must NOT read authenticated
            out.isAuthLog = await ctx.page.evaluate(
                () => globalThis.__probeLog
            );
            // specific-user miss + empty filter -> active account (grace),
            // not accounts[0] (ada)
            await mount(ctx, "probe", {
                matchAccount: { homeAccountId: "no-such-account-id" },
                filter: {},
            });
            await waitRendered(ctx, "status:none");
            out.missAndEmptyFilter = await rendered(ctx);
            return out;
        },
    },
];
