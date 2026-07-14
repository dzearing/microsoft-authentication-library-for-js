/**
 * Seam-hardening harness (test/unit/seams.mjs): exposes the core module and
 * every feature closure so checks can compose arbitrary subsets and assert
 * that non-composed feature APIs fail with feature_not_configured. This
 * bundle is a test vehicle, NOT a size-measurement target.
 */
import * as core from "@mini-msal/browser";
import { popup } from "@mini-msal/browser/popup";
import { pop } from "@mini-msal/browser/pop";
import { broker } from "@mini-msal/browser/broker";
import { telemetry } from "@mini-msal/browser/telemetry";
import { localStorageCache } from "@mini-msal/browser/local-storage";
import { cacheMigration } from "@mini-msal/browser/cache-migration";

const g = globalThis as any;
g.__core = core;
g.__features = { popup, pop, broker, telemetry, localStorageCache, cacheMigration };

document.body.textContent = "mini-core seam harness";
