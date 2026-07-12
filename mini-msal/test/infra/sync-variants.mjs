/**
 * Generates the mini-msal and mock-IdP app variants from test/apps/app.tsx so
 * all variants exercise byte-for-byte identical usage — only the import
 * targets (and, for the mock, the auth config) differ.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const appsDir = fileURLToPath(new URL("../apps/", import.meta.url));
const app = readFileSync(appsDir + "app.tsx", "utf8");

const mini = app
    .replace(/"@azure\/msal-browser"/g, '"@mini-msal/compat"')
    .replace(/"@azure\/msal-react"/g, '"@mini-msal/react"')
    // mini completes popup/silent via its own redirect-bridge page
    .replace(/"\/popup\.html"/g, '"/mini-popup.html"');
writeFileSync(
    appsDir + "mini-app.tsx",
    "// GENERATED from app.tsx by sync-variants.mjs — do not edit\n" + mini
);

writeFileSync(
    appsDir + "mini-app-mock.tsx",
    "// GENERATED from app.tsx by sync-variants.mjs — do not edit\n" +
        mini.replace(/"\.\/authConfig\.js"/g, '"./authConfig.mock.js"')
);

writeFileSync(
    appsDir + "app-mock.tsx",
    "// GENERATED from app.tsx by sync-variants.mjs — do not edit\n" +
        app.replace(/"\.\/authConfig\.js"/g, '"./authConfig.mock.js"')
);

// Compat-only (no React) counterpart of msal-only.ts — measures the full
// @mini-msal/compat drop-in against msal-browser-core apples-to-apples.
const only = readFileSync(appsDir + "msal-only.ts", "utf8");
writeFileSync(
    appsDir + "mini-only.ts",
    "// GENERATED from msal-only.ts by sync-variants.mjs — do not edit\n" +
        only.replace(/"@azure\/msal-browser"/g, '"@mini-msal/compat"')
);

console.log("generated test/apps/{mini-app,mini-app-mock,app-mock}.tsx + mini-only.ts");
