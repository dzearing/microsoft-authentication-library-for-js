/**
 * Replace with your app registration's values:
 * https://learn.microsoft.com/entra/identity-platform/scenario-spa-app-registration
 */
export const msalConfig = {
    auth: {
        clientId: "11111111-2222-3333-4444-555555555555",
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin + "/",
    },
};

export const loginRequest = { scopes: ["User.Read"] };
