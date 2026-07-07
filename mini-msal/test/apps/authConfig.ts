/**
 * Central auth config for the real-msal variants. To test against real
 * Entra ID, replace clientId with your app registration's ID (SPA platform,
 * redirect URI http://localhost:4173/) — see README "Testing with real auth".
 */
export const msalConfig = {
    auth: {
        clientId: "eb54393f-391b-4d47-aa51-cedc8ecd143d",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: "/",
        postLogoutRedirectUri: "/",
    },
    cache: {
        cacheLocation: "sessionStorage" as const,
    },
};

export const loginRequest = { scopes: ["User.Read"] };
