/**
 * Area 13 — cache entity semantics (C16): access-token scope dedupe on save,
 * multi-match clearing on lookup, guest-tenant profile merging into the base
 * account entity, msal.2 -> msal.3 schema migration at initialize, and KMSI
 * plaintext persistence in localStorage mode.
 */
import {
    stdConfig,
    gotoHarness,
    create,
    tryEval,
    tryResult,
    storageDump,
    loginViaPopup,
    digestIdpLog,
    idp,
    CLIENT_ID,
} from "../lib.mjs";

export const area = "cache";

/** Sign in via popup, then reset the IdP request log (keeps the session warm). */
async function seedLogin(ctx, config) {
    await gotoHarness(ctx);
    const login = await loginViaPopup(ctx, config);
    if (login.err) throw new Error(`seed login failed: ${JSON.stringify(login.err)}`);
    await idp.reset();
    await idp.session({ active: "1", user: "0" });
    return login;
}

const tokenGrants = (reqs) =>
    reqs.map((r) =>
        r.endpoint === "token"
            ? `token:${r.body?.grant_type}`
            : `${r.endpoint}${r.query?.prompt ? `:prompt=${r.query.prompt}` : ""}`
    );

/** Digest one AccountInfo incl. tenant-profile expansion + kmsi. */
const ACCOUNT_DIGEST = `(a) => a && {
    username: a.username,
    homeAccountId: a.homeAccountId,
    localAccountId: a.localAccountId,
    tenantId: a.tenantId,
    name: a.name ?? null,
    kmsi: a.kmsi === undefined ? "<absent>" : a.kmsi,
    hasIdTokenClaims: !!a.idTokenClaims && Object.keys(a.idTokenClaims).length > 0,
    profiles: a.tenantProfiles
        ? [...(a.tenantProfiles.values?.() ?? a.tenantProfiles)].map((p) => ({
              tenantId: p.tenantId,
              isHomeTenant: p.isHomeTenant,
          }))
        : null,
}`;

const allAccounts = (ctx) =>
    tryEval(
        ctx,
        (digestSrc) =>
            globalThis.__msal.getAllAccounts().map((0, eval)(digestSrc)),
        ACCOUNT_DIGEST
    );

/** Read the msal.3 token-key index from a storage area. */
const atIndex = (ctx, storage = "sessionStorage") =>
    ctx.page.evaluate((s) => {
        const store = window[s];
        const k = Object.keys(store).find((x) =>
            x.startsWith("msal.3.token.keys")
        );
        return k ? JSON.parse(store.getItem(k)).accessToken : [];
    }, storage);

export const scenarios = [
    {
        id: "cache.at-scope-dedupe",
        note: "acquiring a superset-scope AT deletes the intersecting-scope AT on save; silent then serves the NEW token",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            // second acquisition grants a superset of the cached AT's scopes
            await idp.config({
                scope: "openid profile User.Read Mail.Read",
                access_token: "mock-access-token-superset",
            });
            const superset = await tryResult(
                ctx,
                async () =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read", "Mail.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                        forceRefresh: true,
                    })
            );
            const atKeysAfter = await atIndex(ctx);
            // narrower request: the OLD token must be gone, the new one served
            const followUp = await tryResult(
                ctx,
                async () =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                    })
            );
            return {
                superset,
                atKeysAfter,
                followUp,
                networkCalls: tokenGrants(await idp.requests()),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "cache.at-multi-match-clear",
        note: ">1 cached ATs matching a silent request are ALL removed and the token is refreshed over the network",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            // seed a second AT whose scopes also cover the request (bypasses
            // save-time dedupe by writing the entity + index directly)
            await ctx.page.evaluate((cid) => {
                const idxKey = `msal.3.token.keys.${cid}`;
                const keys = JSON.parse(sessionStorage.getItem(idxKey));
                const srcKey = keys.accessToken[0];
                const entity = JSON.parse(sessionStorage.getItem(srcKey));
                entity.target = "User.Read Extra.Scope";
                entity.secret = "mock-access-token-clone";
                const parts = srcKey.split("|");
                parts[6] = "user.read extra.scope";
                const cloneKey = parts.join("|");
                sessionStorage.setItem(cloneKey, JSON.stringify(entity));
                keys.accessToken.push(cloneKey);
                sessionStorage.setItem(idxKey, JSON.stringify(keys));
            }, CLIENT_ID);
            const result = await tryResult(
                ctx,
                async () =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                    })
            );
            return {
                result,
                atKeysAfter: await atIndex(ctx),
                networkCalls: tokenGrants(await idp.requests()),
            };
        },
    },
    {
        id: "cache.tenant-profile-merge",
        note: "guest-tenant token merges a tenantProfile into ONE base account entity (guest isHomeTenant:false)",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            // guest tenant issues a different tid, same client_info (home id)
            await idp.config({ "claims.tid": "tenant-b" });
            const guest = await tryResult(
                ctx,
                async () =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                        forceRefresh: true,
                    })
            );
            const guestTenantId = await tryEval(ctx, async () => {
                const r = await globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: globalThis.__msal
                        .getAllAccounts()
                        .find((a) => a.tenantId === "tenant-b"),
                });
                return { tenantId: r.tenantId, fromCache: r.fromCache };
            });
            const accountIdx = await ctx.page.evaluate(() =>
                JSON.parse(sessionStorage.getItem("msal.3.account.keys"))
            );
            return {
                guest,
                guestTenantId,
                accounts: await allAccounts(ctx),
                accountIdx,
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "cache.schema-migration",
        note: "msal.2-schema entities are migrated to msal.3 at initialize; stale (>5d) and expired entries pruned",
        async run(ctx) {
            const config = stdConfig(ctx);
            await gotoHarness(ctx);
            await idp.reset();
            // seed a v4-era (schema 2) cache before the client exists
            await ctx.page.evaluate((cid) => {
                const b64url = (o) =>
                    btoa(JSON.stringify(o))
                        .replace(/\+/g, "-")
                        .replace(/\//g, "_")
                        .replace(/=+$/, "");
                const now = Date.now();
                const fresh = String(now);
                const stale = String(now - 6 * 864e5); // > 5-day retention
                const nowSec = Math.floor(now / 1000);
                const env = "localhost:4599";
                const jwt = (user) =>
                    `${b64url({ alg: "none" })}.${b64url({
                        aud: cid,
                        iss: "https://localhost:4599/tenant/v2.0",
                        ...user,
                        exp: nowSec + 3600,
                        iat: nowSec,
                    })}.sig`;
                const ada = {
                    sub: "sub-123",
                    oid: "oid-123",
                    tid: "tenant-123",
                    preferred_username: "ada@contoso.com",
                    name: "Ada Lovelace",
                };
                const grace = {
                    sub: "sub-456",
                    oid: "oid-456",
                    tid: "tenant-123",
                    preferred_username: "grace@contoso.com",
                    name: "Grace Hopper",
                };
                const home = (u) => `${u.oid}.${u.tid}`;
                const acctKey = (u) => `msal.2|${home(u)}|${env}|${u.tid}`;
                const set = (k, v) =>
                    sessionStorage.setItem(k, JSON.stringify(v));
                const account = (u, lastUpdatedAt) => ({
                    homeAccountId: home(u),
                    environment: env,
                    realm: u.tid,
                    localAccountId: u.oid,
                    username: u.preferred_username,
                    authorityType: "MSSTS",
                    name: u.name,
                    clientInfo: b64url({ uid: u.oid, utid: u.tid }),
                    lastUpdatedAt,
                });
                set(acctKey(ada), account(ada, fresh));
                set(acctKey(grace), account(grace, stale));
                const idKey = (u) =>
                    `msal.2|${home(u)}|${env}|idtoken|${cid}|${u.tid}||`;
                const idTok = (u, lastUpdatedAt) => ({
                    credentialType: "IdToken",
                    homeAccountId: home(u),
                    environment: env,
                    clientId: cid,
                    secret: jwt(u),
                    realm: u.tid,
                    lastUpdatedAt,
                });
                set(idKey(ada), idTok(ada, fresh));
                set(idKey(grace), idTok(grace, fresh));
                const rtKey = `msal.2|${home(ada)}|${env}|refreshtoken|${cid}|||`;
                set(rtKey, {
                    credentialType: "RefreshToken",
                    homeAccountId: home(ada),
                    environment: env,
                    clientId: cid,
                    secret: "legacy-rt-0",
                    lastUpdatedAt: fresh,
                });
                // expired AT: pruned during migration, never reaches msal.3
                const atKey = `msal.2|${home(ada)}|${env}|accesstoken|${cid}|tenant-123|openid profile user.read|`;
                set(atKey, {
                    credentialType: "AccessToken",
                    homeAccountId: home(ada),
                    environment: env,
                    clientId: cid,
                    secret: "legacy-at-0",
                    realm: "tenant-123",
                    target: "openid profile User.Read",
                    cachedAt: String(nowSec - 7200),
                    expiresOn: String(nowSec - 600),
                    extendedExpiresOn: String(nowSec - 600),
                    tokenType: "Bearer",
                    lastUpdatedAt: fresh,
                });
                set("msal.2.account.keys", [acctKey(ada), acctKey(grace)]);
                set(`msal.2.token.keys.${cid}`, {
                    idToken: [idKey(ada), idKey(grace)],
                    accessToken: [atKey],
                    refreshToken: [rtKey],
                });
            }, CLIENT_ID);
            await create(ctx, config);
            return {
                accounts: await allAccounts(ctx),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "cache.kmsi-plaintext-localstorage",
        note: "KMSI (signin_state) accounts are stored PLAINTEXT in localStorage mode and survive losing the encryption cookie",
        async run(ctx) {
            const config = stdConfig(ctx, {
                cache: { cacheLocation: "localStorage" },
            });
            await gotoHarness(ctx);
            await idp.reset();
            await idp.config({ "claims.signin_state": '["kmsi"]' });
            const login = await loginViaPopup(ctx, config);
            const accounts = await allAccounts(ctx);
            const storage = await storageDump(ctx);
            // simulate a browser restart: the session encryption cookie is
            // gone, plaintext KMSI entries must still be readable
            await ctx.page.evaluate(() => {
                document.cookie =
                    "msal.cache.encryption=;path=/;expires=Thu, 01 Jan 1970 00:00:00 GMT;SameSite=None;Secure";
            });
            await gotoHarness(ctx);
            await create(ctx, config);
            return {
                login,
                accounts,
                storage,
                accountsAfterRestart: await allAccounts(ctx),
            };
        },
    },
];
