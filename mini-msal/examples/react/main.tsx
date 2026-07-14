/**
 * React bindings over the compat drop-in: MsalProvider + templates + hooks
 * (@mini-msal/react matches msal-react 5.5.1's contracts). The provider
 * initializes the instance and completes redirect roundtrips itself.
 */
import { PublicClientApplication } from "@mini-msal/compat";
import {
    MsalProvider,
    AuthenticatedTemplate,
    UnauthenticatedTemplate,
    useMsal,
} from "@mini-msal/react";
import * as React from "react";
import { createRoot } from "react-dom/client";
import { msalConfig, loginRequest } from "./authConfig.js";

const pca = new PublicClientApplication(msalConfig);

function Profile() {
    const { instance, accounts } = useMsal();
    const account = accounts[0];
    const [tokenLength, setTokenLength] = React.useState<number | null>(null);

    React.useEffect(() => {
        instance
            .acquireTokenSilent({ ...loginRequest, account })
            .then((result) => setTokenLength(result.accessToken.length))
            .catch(console.error);
    }, [instance, account]);

    return (
        <>
            <p>Hello {account.username}</p>
            {tokenLength !== null && <p>token: {tokenLength} chars</p>}
            <button onClick={() => instance.logoutRedirect({ account })}>
                Sign out
            </button>
        </>
    );
}

function SignIn() {
    const { instance, inProgress } = useMsal();
    return (
        <button
            id="signin"
            disabled={inProgress !== "none"}
            onClick={() => instance.loginRedirect(loginRequest)}
        >
            Sign in
        </button>
    );
}

createRoot(document.getElementById("root")!).render(
    <MsalProvider instance={pca}>
        <AuthenticatedTemplate>
            <Profile />
        </AuthenticatedTemplate>
        <UnauthenticatedTemplate>
            <SignIn />
        </UnauthenticatedTemplate>
    </MsalProvider>
);
