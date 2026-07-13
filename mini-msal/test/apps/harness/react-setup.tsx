/**
 * React-harness glue on top of setup.ts (12-react scenarios). Adds:
 *   __mount(fixture, props) — render a named fixture under MsalProvider
 *                             (instance = the __create'd __msal); resolves
 *                             after a settle (2 rAF + macrotask)
 *   __unmount()             — tear the root down
 *   #root textContent       — live observation surface; a <Status> recorder
 *                             is always mounted ("|status:<inProgress>;
 *                             accounts:<n>;logger:<type>|")
 *   __renderLog             — ordered distinct inProgress values seen
 *   __probeLog              — ordered distinct isAuthenticated values seen
 *   __hook                  — latest useMsalAuthentication observation
 *
 * Fixtures live here because page.evaluate cannot pass components/functions.
 */
import * as React from "react";
import { createRoot, type Root } from "react-dom/client";
import { setupHarness } from "./setup.js";

export function setupReactHarness(
    lib: any,
    reactLib: any,
    target: string
): void {
    setupHarness(lib, target);
    const g = globalThis as any;
    const {
        MsalProvider,
        useMsal,
        useIsAuthenticated,
        useAccount,
        useMsalAuthentication,
        AuthenticatedTemplate,
        UnauthenticatedTemplate,
        MsalAuthenticationTemplate,
    } = reactLib;

    const rootEl = document.createElement("div");
    rootEl.id = "root";
    document.body.appendChild(rootEl);
    let root: Root | null = null;

    const pushDistinct = (log: string[], entry: string) => {
        if (log[log.length - 1] !== entry) log.push(entry);
    };

    /** always-mounted recorder: inProgress transitions + live context */
    function Status() {
        const ctx = useMsal();
        pushDistinct(g.__renderLog, String(ctx.inProgress));
        return (
            <div>
                {`|status:${ctx.inProgress};accounts:${ctx.accounts.length};logger:${typeof ctx.logger}|`}
            </div>
        );
    }

    function Probe(props: any) {
        const isAuth = useIsAuthenticated(props.matchAccount ?? undefined);
        const account = useAccount(props.filter ?? undefined);
        pushDistinct(g.__probeLog, String(isAuth));
        return (
            <div>
                {`|probe:isAuth=${isAuth};account=${account?.username ?? "none"}|`}
            </div>
        );
    }

    function Templates(props: any) {
        return (
            <>
                <AuthenticatedTemplate {...(props.authProps ?? {})}>
                    {props.fnChildren
                        ? (c: any) => `[auth-fn:${c.accounts.length}]`
                        : "[auth-plain]"}
                </AuthenticatedTemplate>
                <UnauthenticatedTemplate {...(props.unauthProps ?? {})}>
                    {props.fnChildren
                        ? (c: any) => `[unauth-fn:${c.accounts.length}]`
                        : "[unauth-plain]"}
                </UnauthenticatedTemplate>
            </>
        );
    }

    function AuthHook(props: any) {
        const res = useMsalAuthentication(
            props.interactionType,
            props.request,
            props.accountIdentifiers ?? undefined
        );
        g.__hook = {
            typeofLogin: typeof res.login,
            typeofAcquireToken: typeof res.acquireToken,
            hasAccessToken: res.result ? !!res.result.accessToken : null,
            resultScopes: res.result?.scopes
                ? [...res.result.scopes].sort()
                : null,
            error: res.error ? g.__serializeError(res.error) : null,
        };
        return (
            <div>
                {`|hook:result=${res.result ? "yes" : "no"};error=${res.error ? "yes" : "no"}|`}
            </div>
        );
    }

    class Boundary extends React.Component<any, { err: any }> {
        state = { err: null as any };
        static getDerivedStateFromError(e: any) {
            return { err: e };
        }
        render() {
            const e = this.state.err;
            if (e) {
                return (
                    <div>{`|boundary:${e?.errorCode ?? e?.name ?? "err"}|`}</div>
                );
            }
            return this.props.children;
        }
    }

    function ErrTemplate(props: any) {
        const ErrorComp = (p: any) => (
            <div>
                {`|error-comp:login=${typeof p.login};acquireToken=${typeof p.acquireToken};hasResultProp=${"result" in p};error=${p.error?.errorCode ?? p.error?.name ?? "none"}|`}
            </div>
        );
        const LoadingComp = (p: any) => (
            <div>{`|loading:inProgress=${p?.inProgress ?? "no-props"}|`}</div>
        );
        return (
            <Boundary>
                <MsalAuthenticationTemplate
                    interactionType={props.interactionType}
                    authenticationRequest={props.request}
                    errorComponent={
                        props.withErrorComponent ? ErrorComp : undefined
                    }
                    loadingComponent={props.withLoading ? LoadingComp : undefined}
                >
                    {props.fnChildren
                        ? (r: any) =>
                              `[child-fn:acquireToken=${typeof r.acquireToken}]`
                        : "[child-plain]"}
                </MsalAuthenticationTemplate>
            </Boundary>
        );
    }

    const fixtures: Record<string, any> = {
        none: () => null,
        probe: Probe,
        templates: Templates,
        authHook: AuthHook,
        errTemplate: ErrTemplate,
    };

    g.__mount = (name: string, props: any = {}) => {
        const Fixture = fixtures[name];
        if (root) root.unmount();
        g.__renderLog = [];
        g.__probeLog = [];
        g.__hook = null;
        root = createRoot(rootEl);
        root.render(
            <MsalProvider instance={g.__msal}>
                <Status />
                <Fixture {...props} />
            </MsalProvider>
        );
        return new Promise((resolve) =>
            requestAnimationFrame(() =>
                requestAnimationFrame(() => setTimeout(resolve, 0))
            )
        );
    };
    g.__unmount = () => {
        if (root) root.unmount();
        root = null;
    };
}
