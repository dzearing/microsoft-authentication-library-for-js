/**
 * Tree-shaking ceiling: imports the entire msal-browser public surface and
 * keeps every export alive, so nothing can be dropped. The delta between this
 * and msal-browser-core (a real app's tree-shaken usage) is all that tree
 * shaking can ever recover from the library.
 */
import * as msal from "@azure/msal-browser";

// keep every export reachable without invoking anything
(globalThis as Record<string, unknown>).__msalAll = msal;
console.log(Object.keys(msal).length, "exports retained");
