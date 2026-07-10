/** Emits the static HTML pages for the E2E mock apps and the real-AAD test. */
import { writeFileSync as writeFs } from "node:fs";
import { fileURLToPath } from "node:url";

// dist/ lives at the package root; this file lives in test/infra/
const ROOT = fileURLToPath(new URL("../../", import.meta.url));
const writeFileSync = (p, content) => writeFs(ROOT + p, content);

const appHtml = (title) =>
    `<!doctype html><meta charset="utf-8"><title>${title}</title><div id="root"></div><script type="module" src="bundle.js"></script>`;

const BRIDGE_HTML =
    '<!doctype html><title>auth</title><script type="module" src="/msal-redirect-bridge/bundle.js"></script>';
const MINI_BRIDGE_HTML =
    '<!doctype html><title>auth</title><script type="module" src="/mini-redirect-bridge/bundle.js"></script>';
const BLANK_HTML = "<!doctype html><title>auth</title>";

for (const app of [
    "mini-mock-app",
    "real-mock-app",
    "conformance-real",
    "conformance-mini",
]) {
    writeFileSync(`dist/${app}/index.html`, appHtml(app));
}
// per-harness popup redirect pages (both stacks complete popup/silent via
// their redirect-bridge page); popup2 is a second registered redirect page
// for per-request redirectUri tests
writeFileSync("dist/conformance-real/popup.html", BRIDGE_HTML);
writeFileSync("dist/conformance-mini/popup.html", MINI_BRIDGE_HTML);
writeFileSync("dist/conformance-real/popup2.html", BRIDGE_HTML);
writeFileSync("dist/conformance-mini/popup2.html", MINI_BRIDGE_HTML);
// popup redirect pages used by the e2e's explicit per-app override
writeFileSync("dist/real-mock-app/popup.html", BRIDGE_HTML);
writeFileSync("dist/mini-mock-app/popup.html", MINI_BRIDGE_HTML);

// root-level pages: the apps' internal popup/silent flows redirect here
// (/popup.html = v5 redirect-bridge for real msal, /mini-popup.html = mini's
// bridge; for real AAD register both as SPA redirect URIs). /blank.html has
// NO bridge on purpose — init.popup-without-bridge depends on it.
writeFileSync("dist/popup.html", BRIDGE_HTML);
writeFileSync("dist/mini-popup.html", MINI_BRIDGE_HTML);
writeFileSync("dist/blank.html", BLANK_HTML);

// real Entra ID test page at the site root (matches the registered SPA
// redirect URI http://localhost:4173/). Footer toggles which stack to load.
writeFileSync(
    "dist/index.html",
    `<!doctype html><meta charset="utf-8"><title>msal bundle-size test (real AAD)</title>
<div id="banner" style="padding:10px 14px;font:bold 15px sans-serif;color:#fff"></div>
<div id="root" style="padding:8px;font:14px sans-serif"></div>
<script type="module">
const stack = localStorage.getItem("stack") || "real";
const other = stack === "real" ? "mini" : "real";
const banner = document.getElementById("banner");
banner.style.background = stack === "mini" ? "#1a7f37" : "#0969da";
banner.innerHTML = (stack === "mini"
    ? "MINI-MSAL rewrite — 20 KB min / 6.9 KB gzip"
    : "REAL MSAL (@azure/msal-react 5.5.1) — 248 KB min / 65 KB gzip") +
    ' &nbsp; <a style="color:#fff" href="#" id="swap">switch to ' + other + " →</a>";
document.getElementById("swap").onclick = (e) => {
    e.preventDefault();
    localStorage.setItem("stack", other);
    location.href = "/"; // sign-in state carries over: shared cache schema
};
import(stack === "mini" ? "/mini-aad-app/bundle.js" : "/real-aad-app/bundle.js");
</script>`
);
console.log("postbuild: html pages written");
