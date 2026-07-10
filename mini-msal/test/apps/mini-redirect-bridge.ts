/**
 * mini-msal popup/silent redirect page: broadcasts the auth response back to
 * the main frame over BroadcastChannel — same page contract as real v5's
 * redirect-bridge. Apps using popup or ssoSilent flows must serve a page like
 * this as the request redirectUri (its bundle size is part of the stack).
 */
import { broadcastResponseToMainFrame } from "@mini-msal/browser/redirect-bridge";

broadcastResponseToMainFrame().catch(console.error);
