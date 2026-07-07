/**
 * msal-browser v5 popup/silent redirect page: broadcasts the auth response
 * from the popup/iframe back to the main frame over BroadcastChannel.
 * v5 apps that use popup or ssoSilent flows must serve a page like this as
 * the request redirectUri (its bundle size is measured as part of the stack).
 */
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

broadcastResponseToMainFrame().catch(console.error);
