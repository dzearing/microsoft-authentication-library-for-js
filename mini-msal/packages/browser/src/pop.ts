/**
 * Proof-of-Possession feature: authenticationScheme "pop" (real MSAL's
 * PopTokenGenerator + CryptoOps). Generates an RSA-2048 (RS256) binding
 * keypair per request, sends its thumbprint as req_cnf, and wraps issued
 * access tokens as SignedHttpRequest JWTs — re-signed on every cache hit.
 * Keypairs persist in real's IndexedDB keystore (msal.db / msal.db.keys,
 * private key unextractable) so cached pop ATs stay signable across page
 * loads. Compose via createClient(config, [pop]).
 */
import { BrowserAuthError, b64url, enc } from "./index.js";
import type { ClientContext, TokenRequest } from "./index.js";

const RSA: RsaHashedKeyGenParams = {
    name: "RSASSA-PKCS1-v1_5",
    hash: "SHA-256",
    modulusLength: 2048,
    publicExponent: new Uint8Array([0x01, 0x00, 0x01]),
};

const b64urlStr = (s: string): string => b64url(enc.encode(s));

const sortedJson = (o: object): string =>
    JSON.stringify(o, Object.keys(o).sort());

interface BoundKeyPair {
    privateKey: CryptoKey;
    publicKey: CryptoKey;
    requestMethod?: string;
    requestUri?: string;
}

// real's DatabaseStorage: db msal.db v1, object store msal.db.keys
const DB = "msal.db";
const idb = (): Promise<IDBDatabase> =>
    new Promise((resolve, reject) => {
        const open = indexedDB.open(DB, 1);
        open.onupgradeneeded = () =>
            open.result.createObjectStore(`${DB}.keys`);
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
    });

const idbOp = async <T,>(
    mode: IDBTransactionMode,
    op: (store: IDBObjectStore) => IDBRequest<T>
): Promise<T> => {
    const db = await idb();
    try {
        return await new Promise<T>((resolve, reject) => {
            const req = op(db.transaction(`${DB}.keys`, mode).objectStore(`${DB}.keys`));
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    } finally {
        db.close();
    }
};

export function pop(ctx: ClientContext): void {
    // in-memory first (like real's AsyncMemoryStorage), IndexedDB fallback
    const mem = new Map<string, BoundKeyPair>();
    const getPair = async (kid: string): Promise<BoundKeyPair | undefined> =>
        mem.get(kid) ??
        (await idbOp("readonly", (s) => s.get(kid)).catch(
            () => undefined
        ));

    ctx.pop = {
        /** real's generateCnf: fresh keypair, kid = b64url(sha256(sorted
         * {e,kty,n})), req_cnf = b64url({kid, xms_ksl:"sw"}) */
        async cnf(req: TokenRequest) {
            const pair = await crypto.subtle.generateKey(RSA, true, [
                "sign",
                "verify",
            ]);
            const pub = await crypto.subtle.exportKey("jwk", pair.publicKey);
            const kid = b64url(
                await crypto.subtle.digest(
                    "SHA-256",
                    enc.encode(
                        sortedJson({ e: pub.e, kty: pub.kty, n: pub.n })
                    )
                )
            );
            // real re-imports the private key unextractable before storing
            const priv = await crypto.subtle.importKey(
                "jwk",
                await crypto.subtle.exportKey("jwk", pair.privateKey),
                RSA,
                false,
                ["sign"]
            );
            const entry: BoundKeyPair = {
                privateKey: priv,
                publicKey: pair.publicKey,
                requestMethod: req.resourceRequestMethod,
                requestUri: req.resourceRequestUri,
            };
            mem.set(kid, entry);
            await idbOp("readwrite", (s) => s.put(entry, kid)).catch(() => {
                /* IndexedDB unavailable: memory-only, like real */
            });
            return {
                kid,
                reqCnf: b64urlStr(JSON.stringify({ kid, xms_ksl: "sw" })),
            };
        },

        /** real's signPopToken: SHR JWT {typ:"pop", alg, kid} around the AT
         * with the resource binding claims + the public JWK as cnf.jwk */
        async sign(at: string, kid: string, req: TokenRequest) {
            const pair = await getPair(kid);
            if (!pair) {
                throw new BrowserAuthError("crypto_key_not_found");
            }
            const pub = await crypto.subtle.exportKey("jwk", pair.publicKey);
            const u = req.resourceRequestUri
                ? new URL(req.resourceRequestUri)
                : undefined;
            const payload = {
                at,
                ts: Math.floor(Date.now() / 1000),
                m: req.resourceRequestMethod?.toUpperCase(),
                u: u?.host,
                nonce: req.shrNonce ?? crypto.randomUUID(),
                p: u?.pathname,
                q: u?.search ? [[], u.search.slice(1)] : undefined,
                client_claims: req.shrClaims || undefined,
                cnf: { jwk: JSON.parse(sortedJson(pub)) },
            };
            const header = {
                typ: req.shrOptions?.header?.typ ?? "pop",
                alg: pub.alg,
                kid: b64urlStr(JSON.stringify({ kid })),
            };
            const token = `${b64urlStr(JSON.stringify(header))}.${b64urlStr(
                JSON.stringify(payload)
            )}`;
            const sig = await crypto.subtle.sign(
                RSA,
                pair.privateKey,
                enc.encode(token)
            );
            return `${token}.${b64url(sig)}`;
        },
    };
}
