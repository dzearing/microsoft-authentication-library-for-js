
import { rspack } from "@rspack/core";
import path from "node:path";
import { fileURLToPath } from "node:url";
const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default {
    name: "mangleprops",
    mode: "production",
    context: path.resolve(__dirname, "../.."),
    entry: { bundle: "./test/apps/mini-app.tsx" },
    target: ["web", "es2022"],
    externals: {
        react: "module react",
        "react-dom": "module react-dom",
        "react-dom/client": "module react-dom/client",
        "react/jsx-runtime": "module react/jsx-runtime",
    },
    output: { path: path.resolve(__dirname, "../../dist-experiment-mini"), filename: "[name].js", clean: true },
    resolve: { extensions: [".ts", ".tsx", ".js"], extensionAlias: { ".js": [".ts", ".tsx", ".js"] } },
    module: { rules: [{ test: /\.tsx?$/, loader: "builtin:swc-loader", options: { jsc: { parser: { syntax: "typescript", tsx: true }, transform: { react: { runtime: "automatic" } }, target: "es2022" } } }] },
    optimization: {
        minimize: true,
        minimizer: [new rspack.SwcJsMinimizerRspackPlugin({
            minimizerOptions: {
                compress: { passes: 3, pure_getters: true, ecma: 2020 },
                mangle: {
                    toplevel: true,
                    props: { regex: "^(homeAccountId|redirectUri|account|localAccountId|ACQUIRE_TOKEN_SUCCESS|username|getAllAccounts|preflight|LOGIN_SUCCESS|metadata|emit|scopes|ACCOUNT_REMOVED|readCache|activeAccountId|ACTIVE_ACCOUNT_CHANGED|acquireTokenRedirect|clientId|loginRedirect|removeEventCallback|acquireTokenSilent|authority|authorizeUrl|refreshToken|acquireTokenPopup|setActiveAccount|getActiveAccount|addEventCallback|loadingComponent|microsoftonline|LOGOUT_SUCCESS|errorComponent|ssoSilent|listeners|LOGIN_FAILURE|refresh_token|ACCOUNT_ADDED|cacheKey|accessTokens|tokenRequest|id_token|access_token|clearAccount|accounts|authorization_endpoint|pollForCode|request|postLogoutRedirectUri|handleRedirectPromise|authenticationRequest|loginPopup|getAccount|writeCache|end_session_endpoint|HANDLE_REDIRECT_END|errorCode|eventType|parent|loginHint|redeem|expiresAt|preferred_username|logoutUrl|error_description|cacheLookupPolicy|auth|processRedirect|interactionType|logoutRedirect|nextListenerId|redirectResult|idToken|token_endpoint|redeemRefresh|com|Silent|fromCharCode|errorMessage|config|logoutPopup|accessToken|displayName|msalContext|Popup|initialize|popup|expires_in|inProgress|microsoft|Redirect|verifier|instance|Default|payload|__msal|buffer|encode|oid|sub|tid|token|email|scope|Read|html|meta|well|Skip|code|ruid|new|all)$" },
                },
                format: { comments: false },
            },
        })],
        concatenateModules: true,
    },
    performance: { hints: false },
    stats: "errors-warnings",
};
