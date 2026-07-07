
import { rspack } from "@rspack/core";
import path from "node:path";
import { fileURLToPath } from "node:url";
const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default {
    name: "mangleprops",
    mode: "production",
    context: path.resolve(__dirname, "../.."),
    entry: { bundle: "./test/apps/app.tsx" },
    target: ["web", "es2022"],
    externals: {
        react: "module react",
        "react-dom": "module react-dom",
        "react-dom/client": "module react-dom/client",
        "react/jsx-runtime": "module react/jsx-runtime",
    },
    output: { path: path.resolve(__dirname, "../../dist-experiment"), filename: "[name].js", clean: true },
    resolve: { extensions: [".ts", ".tsx", ".js"], extensionAlias: { ".js": [".ts", ".tsx", ".js"] } },
    module: { rules: [{ test: /\.tsx?$/, loader: "builtin:swc-loader", options: { jsc: { parser: { syntax: "typescript", tsx: true }, transform: { react: { runtime: "automatic" } }, target: "es2022" } } }] },
    optimization: {
        minimize: true,
        minimizer: [new rspack.SwcJsMinimizerRspackPlugin({
            minimizerOptions: {
                compress: { passes: 3, pure_getters: true, ecma: 2020 },
                mangle: {
                    toplevel: true,
                    props: { regex: "^(correlationId|performanceClient|logger|browserStorage|config|verbose|homeAccountId|eventHandler|clientId|platformAuthProvider|account|incrementFields|trace|authority|accessToken|nativeAccountId|addFields|extraParameters|getRequestCorrelationId|extraQueryParameters|emitEvent|authenticationScheme|localAccountId|browserCrypto|cacheManager|navigationClient|environment|serverTelemetryManager|platformAuthType|serverTelemetryEnabled|startMeasurement|system|microsoftonline|refreshToken|lastUpdatedAt|errorCode|controller|auth|authOptions|redirectUri|idToken|setInteractionInProgress|idTokenClaims|validateAndParseJson|postLogoutRedirectUri|canonicalAuthorityUrlComponents|getTokenKeys|canonicalAuthority|credentialType|tenantId|authorityOptions|getUrlComponents|username|commonLogger|hostnameAndPort|protocolMode|acquireToken|acquireTokenRedirect|Redirect|redirectNavigationTimeout|isBrowserEnvironment|loginHint|HostNameAndPort|embeddedClientId|PathSegments|encryptionCookie|azureCloudOptions|error_description|initialized|cryptoUtils|reject|trackStateChangeWithMeasurement|tenantProfiles|getAllAccounts|tokenType|createAuthCodeClient|ssoSilentMeasurement|generateCacheKey|warning|ACCESS_TOKEN_WITH_AUTH_SCHEME|removeTemporaryItem|acquireTokenPopup|acquireTokenByCodeAsyncMeasurement|generateCredentialKey|handleRedirectPromise|nativeInternalStorage|ACQUIRE_TOKEN_FAILURE|closeConnection|scopes|regionDiscoveryMetadata|networkClient|setTokenKeys|nativeStorage|state|OIDCOptions|handleDatabaseAccessError|acquireTokenSilent_silentFlow|resource|networkInterface|metadata|getInteractionInProgress|cacheFailedRequest|extensionId|end_session_endpoint|clientCredentials|telemetryCacheKey|internalStorage|authorization_endpoint|base64Decode|end|authorityType|getActiveAccount|_canonicalAuthorityUrlComponents|expiresOn|ACQUIRE_TOKEN_SUCCESS|setUserData|getAccessTokenCredential|waitForPopupResponse|interactionType|inProgress)$" },
                },
                format: { comments: false },
            },
        })],
        concatenateModules: true,
    },
    performance: { hints: false },
    stats: "errors-warnings",
};
