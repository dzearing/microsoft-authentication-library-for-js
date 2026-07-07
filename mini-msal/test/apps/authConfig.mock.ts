/**
 * Auth config pointing at the local mock IdP (mock-idp.mjs) for E2E tests.
 * Works for both real MSAL (needs knownAuthorities + OIDC protocol mode to
 * trust a non-AAD authority) and mini-msal (ignores the extra fields).
 */
const pageUrl = location.href.split(/[?#]/)[0];

export const msalConfig = {
    auth: {
        clientId: "11111111-2222-3333-4444-555555555555",
        authority: "https://localhost:4599/tenant",
        knownAuthorities: ["localhost:4599"],
        protocolMode: "OIDC" as const,
        redirectUri: pageUrl,
        postLogoutRedirectUri: pageUrl,
    },
    cache: {
        cacheLocation: "sessionStorage" as const,
    },
};

export const loginRequest = { scopes: ["User.Read"] };
