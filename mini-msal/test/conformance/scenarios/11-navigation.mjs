/**
 * Area 11 — redirect navigation seams (C13): navigateToLoginRequestUrl
 * deep-link replay, auth.onRedirectNavigate cancel hook, custom
 * NavigationClient (system.navigationClient + setNavigationClient).
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    capture,
} from "../lib.mjs";

export const area = "navigation";

/** Stable digest of a page URL: path + query + sorted hash param names. */
function urlDigest(href) {
    const u = new URL(href);
    return {
        page: u.origin + u.pathname,
        query: u.search,
        hashKeys: [...new URLSearchParams(u.hash.slice(1)).keys()].sort(),
    };
}

/** C13 probe: replay temp keys (request.origin / urlHash) + interaction lock. */
function navProbe(ctx) {
    return ctx.page.evaluate(() => {
        const out = { origin: null, urlHashKeys: null, lock: null };
        for (let i = 0; i < sessionStorage.length; i++) {
            const k = sessionStorage.key(i);
            if (k.endsWith(".request.origin")) {
                out.origin = sessionStorage.getItem(k);
            } else if (k.endsWith(".urlHash")) {
                out.urlHashKeys = [
                    ...new URLSearchParams(
                        sessionStorage.getItem(k).replace(/^#/, "")
                    ).keys(),
                ].sort();
            } else if (k.includes("interaction.status")) {
                out.lock = "set";
            }
        }
        return out;
    });
}

export const scenarios = [
    {
        id: "navigation.deep-link-replay",
        note: "default navigateToLoginRequestUrl: deep-link start page cached (request.origin), response hash cached (urlHash) + replayed back to the deep link, result delivered there on the next load",
        async run(ctx) {
            const deepLink = ctx.harnessUrl + "?tab=2&x=1";
            await ctx.page.goto(deepLink);
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            const config = stdConfig(ctx);
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
            await create(ctx, config);
            const beforeReplay = {
                url: urlDigest(ctx.page.url()),
                ...(await navProbe(ctx)),
            };
            // fire handleRedirectPromise; the page navigates back to the
            // deep link mid-call (evaluate may die with the context)
            await ctx.page
                .evaluate(() => {
                    globalThis.__msal.handleRedirectPromise().catch(() => {});
                })
                .catch(() => {});
            await ctx.page.waitForURL((u) => u.searchParams.get("tab") === "2", {
                timeout: 15000,
            });
            await ctx.page.waitForFunction(() => !!globalThis.__create);
            const afterReplay = {
                url: urlDigest(ctx.page.url()),
                ...(await navProbe(ctx)),
            };
            await create(ctx, config);
            const result = await tryResult(ctx, () =>
                globalThis.__msal.handleRedirectPromise()
            );
            const cap = await capture(ctx);
            return {
                beforeReplay,
                afterReplay,
                result,
                finalUrl: urlDigest(ctx.page.url()),
                after: await navProbe(ctx),
                events: cap.events,
                accounts: await tryEval(ctx, () =>
                    globalThis.__msal.getAllAccounts().map((a) => a.username)
                ),
            };
        },
    },
    {
        id: "navigation.on-redirect-navigate-cancel-login",
        note: "auth.onRedirectNavigate gets the authorize URL; returning false cancels navigation, loginRedirect resolves, interaction lock stays held",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), {
                onRedirectNavigate: "record-cancel",
            });
            const result = await tryEval(ctx, () =>
                globalThis.__msal.loginRedirect({ scopes: ["User.Read"] })
            );
            return {
                result,
                navCalls: await tryEval(ctx, () => globalThis.__navCalls),
                url: urlDigest(ctx.page.url()),
                ...(await navProbe(ctx)),
            };
        },
    },
    {
        id: "navigation.on-redirect-navigate-cancel-logout",
        note: "onRedirectNavigate on logoutRedirect: false cancels navigation, releases the interaction lock, logoutEnd fires",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), {
                onRedirectNavigate: "record-cancel",
            });
            const result = await tryEval(ctx, () =>
                globalThis.__msal.logoutRedirect()
            );
            const cap = await capture(ctx);
            return {
                result,
                navCalls: await tryEval(ctx, () => globalThis.__navCalls),
                url: urlDigest(ctx.page.url()),
                events: cap.events,
                ...(await navProbe(ctx)),
            };
        },
    },
    {
        id: "navigation.navigation-client-config",
        note: "system.navigationClient: loginRedirect routes through navigateExternal (apiId/timeout/noHistory) instead of navigating",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { navigationClient: "config" });
            const result = await tryEval(ctx, () =>
                globalThis.__msal.loginRedirect({ scopes: ["User.Read"] })
            );
            return {
                result,
                navCalls: await tryEval(ctx, () => globalThis.__navCalls),
                url: urlDigest(ctx.page.url()),
                ...(await navProbe(ctx)),
            };
        },
    },
    {
        id: "navigation.navigation-client-setter",
        note: "NavigationClient export + setNavigationClient(): swapped client is consulted for redirect navigation",
        async run(ctx) {
            await gotoHarness(ctx);
            await create(ctx, stdConfig(ctx), { navigationClient: "setter" });
            const navClientExport = await tryEval(ctx, () => ({
                type: typeof globalThis.__lib.NavigationClient,
                subclassable:
                    typeof globalThis.__lib.NavigationClient === "function" &&
                    new (class extends globalThis.__lib.NavigationClient {})() instanceof
                        globalThis.__lib.NavigationClient,
            }));
            const result = await tryEval(ctx, () =>
                globalThis.__msal.loginRedirect({ scopes: ["User.Read"] })
            );
            return {
                navClientExport,
                result,
                navCalls: await tryEval(ctx, () => globalThis.__navCalls),
                url: urlDigest(ctx.page.url()),
                ...(await navProbe(ctx)),
            };
        },
    },
];
