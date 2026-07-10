/**
 * Redirect-bridge page module (@mini-msal/browser/redirect-bridge): popup and
 * hidden-iframe flows redirect to a page running this, which broadcasts the
 * auth response back to the main frame over a BroadcastChannel keyed by the
 * lib-state id — the same page contract as real msal-browser v5's
 * redirect_bridge entry ({v: 1, payload}). The core's waitForCode listens on
 * that channel; a redirect page without the bridge just times out, like real.
 * (Redirect flows land on the app page itself and never need this.)
 */
export async function broadcastResponseToMainFrame(): Promise<void> {
    document.title = "Microsoft Authentication";
    const payload = (location.hash || location.search).slice(1);
    const state = new URLSearchParams(payload).get("state") ?? "";
    const { id } = JSON.parse(atob(state.split("|")[0]));
    // strip the auth response from the URL before anything can record it
    history.replaceState(null, "", location.origin + location.pathname);
    const channel = new BroadcastChannel(id);
    channel.postMessage({ v: 1, payload });
    channel.close();
    try {
        window.close();
    } catch {
        /* iframe or user-opened tab: nothing to close */
    }
}
