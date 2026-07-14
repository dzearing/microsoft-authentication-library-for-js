/**
 * Consumer-packaging check (D1): proves the packages work for someone who has
 * never seen this repo. Builds each package's dist (tsc), `npm pack`s the
 * three tarballs, installs them into a throwaway consumer app (temp dir,
 * file: deps), then verifies from the consumer's side:
 *   1. types resolve under a STRICT tsconfig (root + every subpath export)
 *   2. a real bundler (rspack, default conditions -> dist) tree-shakes a
 *      core-only composition to roughly the size matrix's floor
 *
 *   npm run pack:check          (network needed: installs react + types)
 */
import { rspack } from "@rspack/core";
import { execFileSync } from "node:child_process";
import { mkdtempSync, mkdirSync, writeFileSync, statSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const TSC = path.join(ROOT, "node_modules", ".bin", "tsc");
const PKGS = ["browser", "compat", "react"];
const run = (cmd, args, cwd) =>
    execFileSync(cmd, args, { cwd, stdio: process.env.DEBUG ? "inherit" : "pipe" });

// 1. fresh dist for every package (browser first: compat/react need its d.ts)
for (const p of PKGS) {
    console.log(`tsc: packages/${p}`);
    run(TSC, ["-p", "tsconfig.json"], path.join(ROOT, "packages", p));
}

// 2. pack + throwaway consumer
const dir = mkdtempSync(path.join(tmpdir(), "mini-msal-consumer-"));
console.log(`consumer: ${dir}`);
const tarballs = {};
for (const p of PKGS) {
    const out = execFileSync("npm", ["pack", "--pack-destination", dir], {
        cwd: path.join(ROOT, "packages", p),
    })
        .toString()
        .trim()
        .split("\n")
        .pop();
    tarballs[p] = out;
}

writeFileSync(
    path.join(dir, "package.json"),
    JSON.stringify(
        {
            name: "mini-msal-consumer",
            private: true,
            type: "module",
            dependencies: {
                "@mini-msal/browser": `file:${tarballs.browser}`,
                "@mini-msal/compat": `file:${tarballs.compat}`,
                "@mini-msal/react": `file:${tarballs.react}`,
                react: "^19.2.1",
                "react-dom": "^19.2.1",
            },
            devDependencies: {
                "@types/react": "^19.2.0",
                "@types/react-dom": "^19.2.0",
            },
        },
        null,
        2
    )
);
writeFileSync(
    path.join(dir, "tsconfig.json"),
    JSON.stringify(
        {
            compilerOptions: {
                target: "ES2022",
                module: "ESNext",
                moduleResolution: "bundler",
                lib: ["ES2022", "DOM", "DOM.Iterable"],
                jsx: "react-jsx",
                strict: true,
                skipLibCheck: true,
                noEmit: true,
            },
            include: ["src"],
        },
        null,
        2
    )
);

mkdirSync(path.join(dir, "src"), { recursive: true });
const write = (name, content) => writeFileSync(path.join(dir, "src", name), content);

// core-only: the pay-to-play floor (mirrors test/apps/mini-core.ts pressure)
write(
    "core-only.ts",
    `import {
    createClient,
    EventType,
    InteractionRequiredAuthError,
    type AccountInfo,
    type AuthenticationResult,
    type EventMessage,
} from "@mini-msal/browser";

const client = createClient({ auth: { clientId: "client-id" } });
async function main(): Promise<void> {
    await client.initialize();
    client.addEventCallback((m: EventMessage) => {
        if (m.eventType === EventType.LOGIN_SUCCESS) {
            const r = m.payload as AuthenticationResult;
            client.setActiveAccount(r.account);
        }
    });
    const res = await client.handleRedirectPromise();
    const account: AccountInfo | null =
        res?.account ?? client.getActiveAccount();
    if (!account) {
        await client.loginRedirect({ scopes: ["User.Read"] });
        return;
    }
    try {
        const t = await client.acquireTokenSilent({ scopes: ["User.Read"], account });
        console.log(t.accessToken.length);
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            await client.acquireTokenRedirect({ scopes: ["User.Read"] });
        }
        await client.logoutRedirect({ account });
    }
}
main().catch(console.error);
`
);
// core + popup: one composed feature via a subpath export
write(
    "core-popup.ts",
    `import { createClient } from "@mini-msal/browser";
import { popup, type PopupClient } from "@mini-msal/browser/popup";

const client = createClient({ auth: { clientId: "client-id" } }, [popup]);
async function main(): Promise<void> {
    await client.initialize();
    const r = await (client as unknown as PopupClient).loginPopup({
        scopes: ["User.Read"],
    });
    console.log(r.account.username);
}
main().catch(console.error);
`
);
// every remaining subpath export must resolve (types + runtime)
write(
    "subpaths.ts",
    `import { pop } from "@mini-msal/browser/pop";
import { broker } from "@mini-msal/browser/broker";
import { localStorageCache } from "@mini-msal/browser/local-storage";
import { cacheMigration } from "@mini-msal/browser/cache-migration";
import { createNestableClient } from "@mini-msal/browser/naa";
import { telemetry } from "@mini-msal/browser/telemetry";
import { broadcastResponseToMainFrame } from "@mini-msal/browser/redirect-bridge";

console.log(
    typeof pop, typeof broker, typeof localStorageCache, typeof cacheMigration,
    typeof createNestableClient, typeof telemetry,
    typeof broadcastResponseToMainFrame
);
`
);
// compat drop-in
write(
    "compat-app.ts",
    `import { PublicClientApplication, InteractionRequiredAuthError } from "@mini-msal/compat";

const pca = new PublicClientApplication({ auth: { clientId: "client-id" } });
async function main(): Promise<void> {
    await pca.initialize();
    try {
        const r = await pca.loginPopup({ scopes: ["User.Read"] });
        console.log(r.account.username);
    } catch (e) {
        if (e instanceof InteractionRequiredAuthError) {
            await pca.loginRedirect({ scopes: ["User.Read"] });
        }
    }
}
main().catch(console.error);
`
);
// react bindings
write(
    "react-app.tsx",
    `import { PublicClientApplication } from "@mini-msal/compat";
import { MsalProvider, useMsal, AuthenticatedTemplate } from "@mini-msal/react";
import { createRoot } from "react-dom/client";

const pca = new PublicClientApplication({ auth: { clientId: "client-id" } });
function Home() {
    const { accounts } = useMsal();
    return <div>{accounts[0]?.username}</div>;
}
createRoot(document.getElementById("root")!).render(
    <MsalProvider instance={pca}>
        <AuthenticatedTemplate>
            <Home />
        </AuthenticatedTemplate>
    </MsalProvider>
);
`
);

console.log("npm install (consumer)…");
run("npm", ["install", "--no-audit", "--no-fund"], dir);

// 3. types resolve for a strict consumer
console.log("tsc --noEmit (strict consumer)…");
run(TSC, ["-p", "tsconfig.json"], dir);
console.log("  types OK");

// 4. real-bundler tree-shake check (default conditions -> dist, NOT src)
const bundle = (name, entry, externals) =>
    new Promise((resolve, reject) =>
        rspack(
            {
                mode: "production",
                context: dir,
                entry: { bundle: entry },
                target: ["web", "es2022"],
                externals,
                output: { path: path.join(dir, "out", name), clean: true },
                resolve: { extensions: [".ts", ".tsx", ".js"] },
                module: {
                    rules: [
                        {
                            test: /\.tsx?$/,
                            loader: "builtin:swc-loader",
                            options: {
                                jsc: {
                                    parser: { syntax: "typescript", tsx: true },
                                    transform: { react: { runtime: "automatic" } },
                                    target: "es2022",
                                },
                            },
                        },
                    ],
                },
                stats: "errors-warnings",
            },
            (err, stats) => {
                if (err || stats.hasErrors()) {
                    reject(err ?? new Error(stats.toString("errors-only")));
                } else {
                    resolve(
                        statSync(path.join(dir, "out", name, "bundle.js")).size / 1024
                    );
                }
            }
        )
    );

const sizes = {};
sizes.core = await bundle("core-only", "./src/core-only.ts");
sizes.popup = await bundle("core-popup", "./src/core-popup.ts");
sizes.compat = await bundle("compat", "./src/compat-app.ts");
sizes.react = await bundle("react", "./src/react-app.tsx", {
    react: "module react",
    "react-dom/client": "module react-dom/client",
    "react/jsx-runtime": "module react/jsx-runtime",
});
for (const [k, v] of Object.entries(sizes)) {
    console.log(`  ${k}: ${v.toFixed(1)} KB min`);
}

// sanity gates vs the size matrix (min, default swc minify settings run a
// little larger than test/infra's passes:3 numbers: core 30.9 / compat 61.2)
const fail = [];
const between = (k, lo, hi) =>
    (sizes[k] < lo || sizes[k] > hi) &&
    fail.push(`${k} ${sizes[k].toFixed(1)} KB outside [${lo}, ${hi}]`);
between("core", 20, 42);
between("compat", 45, 80);
sizes.core >= sizes.popup && fail.push("popup feature added no bytes over core");
sizes.popup >= sizes.compat && fail.push("compat not larger than core+popup");

if (fail.length) {
    console.error("PACKAGING CHECK FAILED:\n  " + fail.join("\n  "));
    process.exit(1);
}
rmSync(dir, { recursive: true, force: true });
console.log("packaging check OK (consumer dir cleaned)");
