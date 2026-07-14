/**
 * Cache schema-migration feature: real MSAL v5's
 * BrowserCacheManager.migrateExistingCache. At initialize, accounts and
 * tokens written by older msal-browser versions (schemas 0-2: keys
 * msal.account.keys / msal.1.* / msal.2.*) are pruned (older than
 * cache.cacheRetentionDays — default 5 — or expired, or invalid) and the
 * survivors copied into the current msal.3 schema, so existing users keep
 * silent SSO after an upgrade. Old-schema entries that survive stay in
 * place, exactly like real. Compose via createClient(config,
 * [cacheMigration]); encrypted old entries (a previous localStorage
 * session's key) are removed, like real when decryption fails.
 */
import type { AccountEntity, ClientContext, TokenEntity } from "./index.js";
import { decodeJwt, isKmsi } from "./index.js";

const P = "msal.3";

interface OldTokenKeys {
    idToken: string[];
    accessToken: string[];
    refreshToken: string[];
}

export function cacheMigration(ctx: ClientContext): void {
    const clientId = ctx.config.auth.clientId;
    const days = ctx.config.cache?.cacheRetentionDays ?? 5;
    ctx.onInit(async () => {
        const store = ctx.getStore();
        const acctIdxKey = (s: number) =>
            s < 1 ? "msal.account.keys" : `msal.${s}.account.keys`;
        const tokIdxKey = (s: number) =>
            s < 1
                ? `msal.token.keys.${clientId}`
                : `msal.${s}.token.keys.${clientId}`;
        const parse = <T,>(raw: string | null): T | null => {
            try {
                return raw ? (JSON.parse(raw) as T) : null;
            } catch {
                return null;
            }
        };
        const readIdx = <T,>(k: string) => parse<T>(store.get(k));
        const writeIdx = (k: string, v: unknown) =>
            store.set(k, JSON.stringify(v));
        const ttlExpired = (lastUpdatedAt: string) =>
            Date.now() > Number(lastUpdatedAt) + days * 864e5;
        const nowSec = Math.floor(Date.now() / 1000);

        /**
         * Real's updateOldEntry for token entities: stamp a missing
         * lastUpdatedAt, remove if invalid / encrypted-by-a-lost-key /
         * TTL-stale / expired. Returns null when removed.
         */
        const oldToken = (key: string): TokenEntity | null => {
            const t = parse<TokenEntity & { data?: string; id?: string }>(
                store.get(key)
            );
            if (!t || typeof t !== "object" || (t.data && t.id)) {
                store.remove(key);
                return null;
            }
            if (!t.lastUpdatedAt) {
                t.lastUpdatedAt = String(Date.now());
                store.set(key, JSON.stringify(t));
            } else if (ttlExpired(t.lastUpdatedAt)) {
                store.remove(key);
                return null;
            }
            if (!t.credentialType || !t.clientId || !t.secret) {
                store.remove(key);
                return null;
            }
            if (
                (t.credentialType === "AccessToken" ||
                    t.credentialType === "RefreshToken") &&
                t.expiresOn &&
                nowSec + 300 > Number(t.expiresOn)
            ) {
                store.remove(key);
                return null;
            }
            return t;
        };

        // pass 1 — real's removeStaleAccounts: prune old-schema accounts
        // (and every token they own) past the retention TTL
        for (let s = 0; s < 3; s++) {
            const idx = readIdx<string[]>(acctIdxKey(s));
            if (!idx?.length) {
                continue;
            }
            const kept: string[] = [];
            for (const k of idx) {
                const a = parse<AccountEntity & { data?: string; id?: string }>(
                    store.get(k)
                );
                if (!a || typeof a !== "object") {
                    store.remove(k);
                } else if (!a.lastUpdatedAt) {
                    a.lastUpdatedAt = String(Date.now());
                    store.set(k, JSON.stringify(a));
                    kept.push(k);
                } else if (ttlExpired(a.lastUpdatedAt)) {
                    const tIdx = readIdx<OldTokenKeys>(tokIdxKey(s));
                    if (tIdx) {
                        for (const type of [
                            "idToken",
                            "accessToken",
                            "refreshToken",
                        ] as const) {
                            tIdx[type] = (tIdx[type] ?? []).filter((tk) => {
                                if (tk.includes(a.homeAccountId)) {
                                    store.remove(tk);
                                    return false;
                                }
                                return true;
                            });
                        }
                        writeIdx(tokIdxKey(s), tIdx);
                    }
                    store.remove(k);
                } else if (a.data && a.id) {
                    // encrypted by a previous session's key
                    store.remove(k);
                } else {
                    kept.push(k);
                }
            }
            writeIdx(acctIdxKey(s), kept);
        }

        const curTokKey = `${P}.token.keys.${clientId}`;
        const curTok = readIdx<OldTokenKeys>(curTokKey) ?? {
            idToken: [],
            accessToken: [],
            refreshToken: [],
        };
        const curAcct = readIdx<string[]>(`${P}.account.keys`) ?? [];
        const credKey = (t: TokenEntity, type: string) =>
            `${P}|${t.homeAccountId}|${t.environment}|${type}|` +
            `${(type === "refreshtoken" && (t as any).familyId) || clientId}|` +
            `${t.realm ?? ""}|${t.target?.toLowerCase() ?? ""}|`;

        // pass 2 — migrate id tokens first (they carry the KMSI claim),
        // building the account's tenantProfiles from the token claims
        for (let s = 0; s < 3; s++) {
            const idx = readIdx<OldTokenKeys>(tokIdxKey(s));
            if (!idx?.idToken?.length) {
                continue;
            }
            const kept: string[] = [];
            for (const k of idx.idToken) {
                const t = oldToken(k);
                if (!t) {
                    continue;
                }
                kept.push(k);
                const curKey = curAcct.find((ak) =>
                    ak.includes(t.homeAccountId)
                );
                const acct =
                    (curKey && parse<AccountEntity>(store.getUser(curKey))) ||
                    parse<AccountEntity>(
                        store.get(
                            (readIdx<string[]>(acctIdxKey(s)) ?? []).find(
                                (ak) => ak.includes(t.homeAccountId)
                            ) ?? ""
                        )
                    );
                if (!acct) {
                    // no account for this token: don't migrate it
                    continue;
                }
                const claims = decodeJwt(t.secret);
                const newKey = credKey(t, "idtoken");
                const cur = parse<TokenEntity>(store.getUser(newKey));
                if (
                    !cur ||
                    (t.lastUpdatedAt ?? "") > (cur.lastUpdatedAt ?? "")
                ) {
                    const profiles = ((acct.tenantProfiles ??= []) as {
                        tenantId: string;
                    }[]);
                    const tid: string =
                        claims.tid ?? claims.tfp ?? claims.acr ?? acct.realm;
                    if (tid && !profiles.some((p) => p.tenantId === tid)) {
                        profiles.push({
                            tenantId: tid,
                            localAccountId: claims.oid ?? claims.sub ?? "",
                            name: claims.name,
                            username:
                                claims.preferred_username ?? claims.upn ?? "",
                            isHomeTenant:
                                tid === acct.homeAccountId.split(".")[1],
                        } as { tenantId: string });
                    }
                    const kmsi = isKmsi(claims);
                    const newAcctKey = `${P}|${acct.homeAccountId}|${acct.environment}|${acct.realm}`;
                    await store.setUser(newAcctKey, JSON.stringify(acct), kmsi);
                    if (!curAcct.includes(newAcctKey)) {
                        curAcct.push(newAcctKey);
                    }
                    await store.setUser(newKey, JSON.stringify(t), kmsi);
                    if (!curTok.idToken.includes(newKey)) {
                        curTok.idToken.push(newKey);
                    }
                }
            }
            idx.idToken = kept;
            writeIdx(tokIdxKey(s), idx);
        }

        // homeAccountId -> KMSI, from the CURRENT schema's id tokens
        const kmsiMap: Record<string, boolean> = {};
        for (const k of curTok.idToken) {
            const t = parse<TokenEntity>(store.getUser(k));
            if (t) {
                kmsiMap[t.homeAccountId] = isKmsi(decodeJwt(t.secret));
            }
        }

        // pass 3 — migrate access + refresh tokens (skipped without an
        // id token / kmsi entry, like real)
        for (const [type, key] of [
            ["accessToken", "accesstoken"],
            ["refreshToken", "refreshtoken"],
        ] as const) {
            for (let s = 0; s < 3; s++) {
                const idx = readIdx<OldTokenKeys>(tokIdxKey(s));
                if (!idx?.[type]?.length) {
                    continue;
                }
                const kept: string[] = [];
                for (const k of idx[type]) {
                    const t = oldToken(k);
                    if (!t) {
                        continue;
                    }
                    kept.push(k);
                    if (!(t.homeAccountId in kmsiMap)) {
                        continue;
                    }
                    const newKey = credKey(t, key);
                    const cur = parse<TokenEntity>(store.getUser(newKey));
                    if (
                        !cur ||
                        (t.lastUpdatedAt ?? "") > (cur.lastUpdatedAt ?? "")
                    ) {
                        await store.setUser(
                            newKey,
                            JSON.stringify(t),
                            kmsiMap[t.homeAccountId]
                        );
                        if (!curTok[type].includes(newKey)) {
                            curTok[type].push(newKey);
                        }
                    }
                }
                idx[type] = kept;
                writeIdx(tokIdxKey(s), idx);
            }
        }
        if (
            curTok.idToken.length +
                curTok.accessToken.length +
                curTok.refreshToken.length >
            0
        ) {
            writeIdx(curTokKey, curTok);
        }
        if (curAcct.length > 0) {
            writeIdx(`${P}.account.keys`, curAcct);
        }
    });
}
