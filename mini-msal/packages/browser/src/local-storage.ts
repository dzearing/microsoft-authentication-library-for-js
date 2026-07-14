/**
 * localStorage cache feature: real MSAL v5's encrypted localStorage mode.
 * Entities are AES-GCM encrypted at rest ({id, nonce, data, lastUpdatedAt}
 * wrappers) under an HKDF base key kept in the msal.cache.encryption session
 * cookie; key indexes and metadata stay plaintext. Reads come from an
 * in-memory plaintext mirror imported at initialize and kept in sync across
 * tabs over the msal.broadcast.cache channel — exactly real's LocalStorage
 * class. Compose via createClient(config, [localStorageCache]); no-op unless
 * cache.cacheLocation === "localStorage". The cookie/crypto helpers are
 * exported for compat's LocalStorage class.
 */
import type { ClientContext, Store } from "./index.js";

const COOKIE = "msal.cache.encryption";

const b64url = (bytes: Uint8Array): string =>
    btoa(String.fromCharCode(...bytes))
        .replace(/\+/g, "-")
        .replace(/\//g, "_")
        .replace(/=+$/, "");

const b64dec = (s: string): Uint8Array<ArrayBuffer> =>
    Uint8Array.from(atob(s.replace(/-/g, "+").replace(/_/g, "/")), (c) =>
        c.charCodeAt(0)
    );

const ZERO_IV = { name: "AES-GCM", iv: new Uint8Array(12) };

// real derives a fresh AES-GCM key per operation: HKDF(salt = the random
// per-write nonce, info = context), then a zero IV
const deriveKey = (
    baseKey: CryptoKey,
    nonce: Uint8Array<ArrayBuffer>,
    context: string
) =>
    crypto.subtle.deriveKey(
        {
            name: "HKDF",
            salt: nonce,
            hash: "SHA-256",
            info: new TextEncoder().encode(context),
        },
        baseKey,
        { name: "AES-GCM", length: 256 },
        false,
        ["encrypt", "decrypt"]
    );

/** read (or create) the msal.cache.encryption session cookie {id, key} and
 * import the HKDF base key — real's LocalStorage.initialize key setup */
export async function loadEncryptionCookie(): Promise<{
    keyId: string;
    baseKey: CryptoKey;
}> {
    const raw = document.cookie
        .split(";")
        .map((c) => c.trim())
        .find((c) => c.startsWith(`${COOKIE}=`));
    let cookie: { id?: string; key?: string } = {};
    try {
        cookie = JSON.parse(decodeURIComponent(raw!.slice(COOKIE.length + 1)));
    } catch {
        /* absent or unparsable: generate a fresh key below */
    }
    let keyId: string;
    let rawKey: Uint8Array<ArrayBuffer>;
    if (cookie.id && cookie.key) {
        keyId = cookie.id;
        rawKey = b64dec(cookie.key);
    } else {
        keyId = crypto.randomUUID();
        rawKey = crypto.getRandomValues(new Uint8Array(32));
        document.cookie = `${COOKIE}=${encodeURIComponent(
            JSON.stringify({ id: keyId, key: b64url(rawKey) })
        )};path=/;SameSite=None;Secure;`;
    }
    const baseKey = await crypto.subtle.importKey("raw", rawKey, "HKDF", false, [
        "deriveKey",
    ]);
    return { keyId, baseKey };
}

/** encrypt one entity into real's at-rest wrapper JSON */
export async function encryptEntry(
    baseKey: CryptoKey,
    keyId: string,
    context: string,
    value: string,
    timestamp: string
): Promise<string> {
    const nonce = crypto.getRandomValues(new Uint8Array(16));
    const data = new Uint8Array(
        await crypto.subtle.encrypt(
            ZERO_IV,
            await deriveKey(baseKey, nonce, context),
            new TextEncoder().encode(value)
        )
    );
    return JSON.stringify({
        id: keyId,
        nonce: b64url(nonce),
        data: b64url(data),
        lastUpdatedAt: timestamp,
    });
}

/** decrypt one at-rest entity; unencrypted data comes back as-is, foreign/
 * corrupt entries as null (caller removes them) — real's
 * getItemFromEncryptedCache semantics */
export async function decryptEntry(
    baseKey: CryptoKey,
    keyId: string,
    context: string,
    raw: string | null
): Promise<string | null> {
    if (!raw) {
        return null;
    }
    let o: { id?: string; nonce?: string; data?: string };
    try {
        o = JSON.parse(raw);
    } catch {
        return null;
    }
    if (!o?.id || !o.nonce || !o.data) {
        // unencrypted data is kept as-is, like real
        return raw;
    }
    if (o.id !== keyId) {
        // encrypted by a previous session's key
        return null;
    }
    try {
        return new TextDecoder().decode(
            await crypto.subtle.decrypt(
                ZERO_IV,
                await deriveKey(baseKey, b64dec(o.nonce), context),
                b64dec(o.data)
            )
        );
    } catch {
        return null;
    }
}

export function localStorageCache(ctx: ClientContext): void {
    if (ctx.config.cache?.cacheLocation !== "localStorage") {
        return;
    }
    const clientId = ctx.config.auth.clientId;
    const mem = new Map<string, string>();
    const channel = new BroadcastChannel("msal.broadcast.cache");
    let keyId = "";
    let baseKey: CryptoKey;

    /** encryption context binds app-specific entries to this clientId */
    const context = (key: string) => (key.includes(clientId) ? clientId : "");

    /** decrypt one entity; null means invalid/foreign (caller removes it) */
    const decrypt = (key: string): Promise<string | null> =>
        decryptEntry(baseKey, keyId, context(key), localStorage.getItem(key));

    const store: Store = {
        async init() {
            // session cookie {id, key}: id marks entries this key can read,
            // key is the raw HKDF base key (real's CookieStorage entry)
            ({ keyId, baseKey } = await loadEncryptionCookie());
            // import existing entities into the mirror, pruning entries this
            // session can't read, and write the surviving indexes back
            // (real's importExistingCache)
            const importKeys = async (keys: string[]) => {
                const kept: string[] = [];
                for (const k of keys) {
                    const v = await decrypt(k);
                    if (v) {
                        mem.set(k, v);
                        kept.push(k);
                    } else {
                        localStorage.removeItem(k);
                    }
                }
                return kept;
            };
            const readIdx = <T,>(k: string): T | null => {
                try {
                    return JSON.parse(localStorage.getItem(k) ?? "null");
                } catch {
                    return null;
                }
            };
            const writeIdx = (k: string, live: boolean, v: unknown) =>
                live
                    ? localStorage.setItem(k, JSON.stringify(v))
                    : localStorage.removeItem(k);
            const acctIdx = "msal.3.account.keys";
            const accounts = await importKeys(readIdx<string[]>(acctIdx) ?? []);
            writeIdx(acctIdx, accounts.length > 0, accounts);
            const tokIdx = `msal.3.token.keys.${clientId}`;
            const t = readIdx<Record<string, string[]>>(tokIdx);
            const tokens = {
                idToken: await importKeys(t?.idToken ?? []),
                accessToken: await importKeys(t?.accessToken ?? []),
                refreshToken: await importKeys(t?.refreshToken ?? []),
            };
            writeIdx(
                tokIdx,
                tokens.idToken.length +
                    tokens.accessToken.length +
                    tokens.refreshToken.length >
                    0,
                tokens
            );
            // other tabs broadcast mirror updates (real's updateCache)
            channel.onmessage = (ev) => {
                const { key, value, context: c } = ev.data;
                if (!key || (c && c !== clientId)) {
                    return;
                }
                value ? mem.set(key, value) : mem.delete(key);
            };
        },
        get: (k) => localStorage.getItem(k),
        set: (k, v) => localStorage.setItem(k, v),
        remove(k) {
            if (mem.has(k)) {
                mem.delete(k);
                channel.postMessage({ key: k, value: null, context: context(k) });
            }
            localStorage.removeItem(k);
        },
        getUser: (k) => mem.get(k) ?? null,
        async setUser(k, v, kmsi) {
            if (kmsi) {
                // real persists KMSI entities PLAINTEXT so sign-in survives
                // losing the per-session encryption cookie (browser restart)
                localStorage.setItem(k, v);
            } else {
                localStorage.setItem(
                    k,
                    await encryptEntry(
                        baseKey,
                        keyId,
                        context(k),
                        v,
                        String(Date.now())
                    )
                );
            }
            mem.set(k, v);
            channel.postMessage({ key: k, value: v, context: context(k) });
        },
    };
    ctx.setStore(store);
}
