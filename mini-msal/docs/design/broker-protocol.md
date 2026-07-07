# Platform broker (WAM) protocol (@azure/msal-browser 5.16.0) — captured from source

Two transports; DOM tried first ONLY if experimental.allowPlatformBrokerWithDOM=true
(requires system.allowPlatformBroker=true too, else invalidPlatformBrokerConfiguration).
Default path = extension handler (MessageChannel + window.postMessage).
All probe errors swallowed at initialize() -> web-flow fallback, INITIALIZE_END still fires.

Constants: CHANNEL_ID "53ee284d-920a-4b59-9d30-a60315b26836";
PREFERRED_EXTENSION_ID "ppnbnpeolgkicgegkbkbjmhlideopiji"; handshake timeout default 2000ms.

## Extension path
Probe: window.postMessage({channel: CHANNEL_ID, extensionId, responseId: <guid>,
body: {method: "Handshake"}}, window.origin, [port2]). If MSAL's own bubble-phase window
listener sees its own Handshake bounce back -> nativeExtensionNotInstalled; timeout 2000ms
-> nativeHandshakeTimeout. Tries PREFERRED_EXTENSION_ID first, retries with undefined.
Fake: capture-phase window "message" listener, stopImmediatePropagation() on Handshake,
reply on event.ports[0]: {channel, extensionId, responseId (echo), body: {method:
"HandshakeResponse", version: "1.0.0"}}. Keep the port.
GetToken arrives on the port: {channel, extensionId, responseId, body: {method: "GetToken",
request: {accountId, clientId, authority, scope, redirectUri, prompt?, correlationId,
tokenType, windowTitleSubstring, extraParameters: {telemetry: "MATS", ...}, claims?, state?, ...}}}.
Success reply on port: {channel, extensionId, responseId (echo), body: {method: "Response",
response: {status: "Success", result: {access_token, id_token (JWT), client_info (b64url
{uid,utid}), account: {id: <MUST equal request.accountId>, properties?}, scope, expires_in,
state, properties: {} (required), shr?, extendedLifetimeToken?}}}}.
Result missing any of access_token/id_token/client_info/account/scope/expires_in -> unexpectedError.
account.id !== request.accountId -> userSwitch. Result sets fromPlatformBroker: true.
Errors: response.status != "Success" with {code, description, ext} -> createNativeAuthError;
ext.status mapping: ACCOUNT_UNAVAILABLE -> IRAE nativeAccountUnavailable; USER_INTERACTION_REQUIRED
-> IRAE; USER_CANCEL -> BrowserAuthError userCancelled; NO_NETWORK -> noNetworkConnectivity;
UI_NOT_ALLOWED -> IRAE uiNotAllowed; DISABLED -> fatal NativeAuthError (drops broker, silent ->
tokenRefreshRequired).

## DOM path
Probe: navigator.platformAuthentication.getSupportedContracts("MicrosoftEntra") includes
"get-token-and-sign-out". executeGetToken(req) where req = {accountId, brokerId:
"MicrosoftEntra", authority, clientId, correlationId, extraParameters (stringified),
isSecurityTokenService: false, redirectUri, scope, state, storeInCache, embeddedClientId}.
Success return (camelCase): {isSuccess: true, accessToken, idToken, clientInfo, account:
{id, properties}, scopes (string), expiresIn, state?, properties?, extendedLifetimeToken?,
proofOfPossessionPayload?}. Error: {isSuccess: false, error: {code, errorCode, description,
protocolError, status, properties}}.

Silent broker routing requires account.nativeAccountId; response.account.id must match.
No public "broker found" getter; observable via AuthenticationResult.fromPlatformBroker.
