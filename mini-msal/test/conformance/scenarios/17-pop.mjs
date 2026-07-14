/**
 * Area 17 — authenticationScheme "pop"/"ssh-cert" (C21): PoP token binding.
 * With authenticationScheme: "pop" real generates an RSA-2048 keypair, sends
 * token_type=pop + req_cnf (base64url {kid, xms_ksl:"sw"}) on the token
 * endpoint, caches the AT as AccessToken_With_AuthScheme with keyId (cache
 * key gains a "pop" scheme suffix), returns tokenType "pop" and the AT as a
 * SignedHttpRequest JWT — re-signed on every cache hit. request.popKid skips
 * keygen AND signing (raw AT returned). "ssh-cert" sends the request's
 * sshJwk as req_cnf and requires sshJwk/sshKid (config errors when absent).
 *
 * Non-determinism (fresh keypair per run: kid, SHR sig/ts/nonce, req_cnf)
 * is digested in-node into structural booleans/keys before snapshotting.
 */
import {
    stdConfig,
    gotoHarness,
    tryEval,
    tryResult,
    storageDump,
    loginViaPopup,
    digestIdpLog,
    idp,
} from "../lib.mjs";

export const area = "pop";

async function seedLogin(ctx, config) {
    await gotoHarness(ctx);
    const login = await loginViaPopup(ctx, config);
    if (login.err) {
        throw new Error(`seed login failed: ${JSON.stringify(login.err)}`);
    }
    await idp.reset();
    await idp.session({ active: "1", user: "0" });
    return login;
}

const b64json = (s) => JSON.parse(Buffer.from(s, "base64url").toString());

/** Structural digest of a SignedHttpRequest JWT (values are per-run). */
function shrDigest(shr) {
    const parts = String(shr ?? "").split(".");
    if (parts.length !== 3) {
        return { parts: parts.length, raw: String(shr).slice(0, 40) };
    }
    const header = b64json(parts[0]);
    const payload = b64json(parts[1]);
    return {
        headerKeys: Object.keys(header),
        typ: header.typ,
        alg: header.alg,
        // header.kid is base64url({ kid: <thumbprint> })
        headerKidShape: Object.keys(b64json(header.kid)),
        payloadKeys: Object.keys(payload),
        m: payload.m ?? null,
        u: payload.u ?? null,
        p: payload.p ?? null,
        q: payload.q ?? null,
        atParts: String(payload.at ?? "").split(".").length,
        nonceLen: String(payload.nonce ?? "").length,
        tsIsEpochSeconds:
            typeof payload.ts === "number" && payload.ts > 1.5e9 && payload.ts < 2.5e9,
        cnfJwkKeys: payload.cnf ? Object.keys(payload.cnf.jwk ?? {}).sort() : null,
        sigPresent: parts[2].length > 0,
        _payload: payload, // stripped before returning observations
        _headerKid: b64json(header.kid).kid,
    };
}

const strip = (d) => {
    const { _payload, _headerKid, ...rest } = d;
    return rest;
};

/** Digest the token-endpoint requests: body keys + decoded req_cnf shape. */
function tokenCalls(reqs) {
    return digestIdpLog(reqs)
        .filter((r) => r.endpoint === "token")
        .map((r) => ({
            grant: r.body.grant_type,
            token_type: r.body.token_type ?? null,
            bodyKeys: Object.keys(r.body).sort(),
            reqCnf: r.body.req_cnf
                ? (() => {
                      try {
                          // pop: base64url JSON; ssh: raw JWK JSON
                          const o = r.body.req_cnf.trim().startsWith("{")
                              ? JSON.parse(r.body.req_cnf)
                              : b64json(r.body.req_cnf);
                          return {
                              keys: Object.keys(o).sort(),
                              xms_ksl: o.xms_ksl ?? null,
                              kidLen: String(o.kid ?? "").length,
                          };
                      } catch {
                          return "unparseable";
                      }
                  })()
                : null,
        }));
}

/** Find cached AT entities (raw, pre-normalization) in the page's storage. */
async function atEntities(ctx) {
    const dump = await storageDump(ctx);
    const store = dump.sessionStorage;
    const keysKey = Object.keys(store).find((k) =>
        k.startsWith("msal.3.token.keys")
    );
    return (store[keysKey]?.accessToken ?? []).map((k) => ({
        key: k,
        entity: store[k],
    }));
}

const popSilent = (arg) =>
    globalThis.__msal.acquireTokenSilent({
        scopes: ["User.Read"],
        account: globalThis.__msal.getAllAccounts()[0],
        authenticationScheme: "pop",
        resourceRequestMethod: "get",
        resourceRequestUri: "https://api.contoso.test:8080/path/sub?foo=bar",
        ...arg,
    });

export const scenarios = [
    {
        id: "pop.silent-shr",
        note: "pop silent: bearer AT skipped (scheme-aware lookup) -> RT grant with token_type/req_cnf; SHR result; cache-hit re-signs; popKid skips keygen+signing",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);

            // 1) pop WITHOUT forceRefresh: the seed's bearer AT must NOT
            //    satisfy a pop request -> refresh_token grant on the wire
            const first = await tryResult(ctx, popSilent);
            const entitiesAfterFirst = await atEntities(ctx);
            const popEnt = entitiesAfterFirst.find(
                (e) => e.entity.tokenType === "pop"
            );
            const firstShr = first.ok ? shrDigest(first.ok.accessToken) : null;

            // 2) same request again: cache hit, SHR re-signed (fresh nonce,
            //    same embedded AT)
            const second = await tryResult(ctx, popSilent);
            const secondShr = second.ok
                ? shrDigest(second.ok.accessToken)
                : null;
            const callsBeforePopKid = (await idp.requests()).length;

            // 3) popKid: no keygen, req_cnf = {kid} only, raw (unsigned) AT
            const third = await tryResult(ctx, popSilent, {
                popKid: "my-pop-kid",
                forceRefresh: true,
            });
            const entitiesAfterThird = await atEntities(ctx);
            const popKidEnt = entitiesAfterThird.find(
                (e) => e.entity.keyId === "my-pop-kid"
            );

            const scrub = (r) =>
                r.ok ? { ...r, ok: { ...r.ok, accessToken: "<by-shape>" } } : r;
            return {
                first: scrub(first),
                firstShr: firstShr && strip(firstShr),
                firstShrChecks: firstShr &&
                    popEnt && {
                        atMatchesCachedSecret:
                            firstShr._payload.at === popEnt.entity.secret,
                        atCnfKidMatchesEntityKeyId:
                            b64json(firstShr._payload.at.split(".")[1]).cnf
                                ?.kid === popEnt.entity.keyId,
                        headerKidMatchesEntityKeyId:
                            firstShr._headerKid === popEnt.entity.keyId,
                    },
                popEntityShape: popEnt && {
                    keySuffix: popEnt.key.split("|").pop(),
                    keyHasSchemeCredType: popEnt.key.includes(
                        "accesstoken_with_authscheme"
                    ),
                    credentialType: popEnt.entity.credentialType,
                    tokenType: popEnt.entity.tokenType,
                    keyIdLen: String(popEnt.entity.keyId ?? "").length,
                },
                bearerSurvives: entitiesAfterFirst.some(
                    (e) => (e.entity.tokenType ?? "Bearer") === "Bearer"
                ),
                second: scrub(second),
                cacheHitChecks: firstShr &&
                    secondShr && {
                        fromCache: second.ok?.fromCache,
                        sameEmbeddedAt:
                            secondShr._payload?.at === firstShr._payload?.at,
                        freshNonce:
                            secondShr._payload?.nonce !==
                            firstShr._payload?.nonce,
                        sameKid: secondShr._headerKid === firstShr._headerKid,
                    },
                third: scrub(third),
                thirdUnsigned: third.ok
                    ? third.ok.accessToken === popKidEnt?.entity.secret
                    : null,
                popKidEntity: popKidEnt && {
                    credentialType: popKidEnt.entity.credentialType,
                    tokenType: popKidEnt.entity.tokenType,
                    keyId: popKidEnt.entity.keyId,
                },
                networkCallsBetweenFirstAndPopKid: callsBeforePopKid,
                tokenCalls: tokenCalls(await idp.requests()),
                storage: await storageDump(ctx),
            };
        },
    },
    {
        id: "pop.ssh-scheme-and-errors",
        note: "ssh-cert: missing sshJwk/sshKid -> config errors; with both, req_cnf = raw JWK, entity keyId from response key_id, raw AT (no signing) on network and cache hits",
        async run(ctx) {
            const config = stdConfig(ctx);
            await seedLogin(ctx, config);
            const sshJwk = JSON.stringify({
                kty: "RSA",
                n: "test-modulus",
                e: "AQAB",
                kid: "test-ssh-kid",
            });

            const missingJwk = await tryEval(ctx, () =>
                globalThis.__msal.acquireTokenSilent({
                    scopes: ["User.Read"],
                    account: globalThis.__msal.getAllAccounts()[0],
                    authenticationScheme: "ssh-cert",
                })
            );
            const missingKid = await tryEval(
                ctx,
                (jwk) =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                        authenticationScheme: "ssh-cert",
                        sshJwk: jwk,
                    }),
                sshJwk
            );
            const ok = await tryResult(
                ctx,
                (jwk) =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                        authenticationScheme: "ssh-cert",
                        sshJwk: jwk,
                        sshKid: "test-ssh-kid",
                    }),
                sshJwk
            );
            const hit = await tryResult(
                ctx,
                (jwk) =>
                    globalThis.__msal.acquireTokenSilent({
                        scopes: ["User.Read"],
                        account: globalThis.__msal.getAllAccounts()[0],
                        authenticationScheme: "ssh-cert",
                        sshJwk: jwk,
                        sshKid: "test-ssh-kid",
                    }),
                sshJwk
            );
            const sshEnt = (await atEntities(ctx)).find(
                (e) => e.entity.tokenType === "ssh-cert"
            );
            return {
                missingJwk,
                missingKid,
                ok,
                hit,
                sshEntity: sshEnt && {
                    keySuffix: sshEnt.key.split("|").pop(),
                    keyHasSchemeCredType: sshEnt.key.includes(
                        "accesstoken_with_authscheme"
                    ),
                    credentialType: sshEnt.entity.credentialType,
                    tokenType: sshEnt.entity.tokenType,
                    keyId: sshEnt.entity.keyId,
                },
                tokenCalls: tokenCalls(await idp.requests()),
                storage: await storageDump(ctx),
            };
        },
    },
];
