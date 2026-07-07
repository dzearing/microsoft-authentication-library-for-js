# NAA bridge protocol (@azure/msal-browser 5.16.0) — captured from source

Bridge object: `window.nestedAppAuthBridge` with `addEventListener("message", cb)` (cb invoked
directly with a JSON string or `{data: string}`) and `postMessage(jsonString)`.
Optional async hook `window.__initializeNestedAppAuth()` awaited before handshake.
Handshake failures are swallowed -> falls back to standard PCA.

Request envelope (all methods): `{messageType: "NestedAppAuthRequest", method, requestId: <guid>,
sendTime, clientLibrary: "msal.js.browser", clientLibraryVersion: "5.16.0", ...}`

- GetInitContext -> respond `{requestId, success: true, initContext: {sdkName, sdkVersion,
  accountContext: {homeAccountId, environment, tenantId} | null, capabilities: {}}}`.
  initContext REQUIRED else NESTED_APP_AUTH_UNAVAILABLE.
- GetTokenPopup (interactive) / GetToken (silent): tokenParams = {platformBrokerId?, clientId,
  authority?, resource?, scope (space-joined, default "openid profile offline_access"),
  correlationId, claims?, state?, authenticationScheme: "Bearer", extraParameters, forceRefresh (silent only)}.
  Respond `{requestId, success: true, token: {access_token REQ, id_token REQ (decodable JWT),
  expires_in?, scope?, authority?}, account: {environment REQ, homeAccountId?, tenantId?, username?,
  localAccountId?, name?, idTokenClaims?, platformBrokerId?, loginHint?}}`.
  token+account both required (undefined -> NESTED_APP_AUTH_UNAVAILABLE). Missing id_token/access_token
  -> ClientAuthError null_or_empty_token; missing account.environment -> invalid_cache_environment.
- Errors: `{requestId, success: false, error: {status REQ, code?, description?}}` with status in:
  USER_INTERACTION_REQUIRED -> InteractionRequiredAuthError(code||"")
  USER_CANCEL -> ClientAuthError user_canceled
  NO_NETWORK -> ClientAuthError no_network_connectivity
  TRANSIENT_ERROR / PERSISTENT_ERROR -> ServerError(code||"")
  DISABLED -> ClientAuthError nested_app_auth_bridge_disabled
  ACCOUNT_UNAVAILABLE -> ClientAuthError no_account_found
  NESTED_APP_AUTH_UNAVAILABLE -> ClientAuthError code||nested_app_auth_bridge_disabled
  other/absent status -> AuthError unknown_error
- Unsupported on NAA controller (throw NestedAppAuthError "unsupported_method"):
  acquireTokenRedirect, acquireTokenByCode, add/removePerformanceCallback, loginRedirect,
  logoutRedirect, logoutPopup, getPerformanceClient, getRedirectResponse, clearCache.
  handleRedirectPromise resolves null; loginPopup -> GetTokenPopup; ssoSilent -> GetToken.
- Silent path checks browser cache first; forceRefresh/claims bypass cache.
