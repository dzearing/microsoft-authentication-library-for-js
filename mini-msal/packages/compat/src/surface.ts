/**
 * The long tail of real msal-browser's export surface (D5): enums/constants,
 * utility namespaces, storage classes, EventHandler, perf helpers and
 * SignedHttpRequest. Everything here is a faithful-but-lean port of the real
 * 5.16.0 source — observable behavior (names, codes, shapes, values) matches;
 * internals are compressed. Pinned by init.exported-surface-full.
 */
import {
    AuthError,
    BrowserAuthError,
    ClientAuthError,
    ClientConfigurationError,
    Logger,
    EventType,
    InteractionType,
    InteractionStatus,
    type EventMessage,
    type TokenRequest,
} from "@mini-msal/browser";
import {
    makeBoundKeyPair,
    signPop,
    keystore,
    type BoundKeyPair,
} from "@mini-msal/browser/pop";
import {
    loadEncryptionCookie,
    encryptEntry,
    decryptEntry,
} from "@mini-msal/browser/local-storage";

/** real's browser-config error class (aka.ms message, like BrowserAuthError) */
export class BrowserConfigurationAuthError extends AuthError {
    name = "BrowserConfigurationAuthError";
    constructor(errorCode: string, subError?: string) {
        super(errorCode, undefined, subError);
    }
}

// ---- real's exported enums/constants ----

export const ApiId = {
    acquireTokenRedirect: 861,
    acquireTokenPopup: 862,
    ssoSilent: 863,
    acquireTokenSilent_authCode: 864,
    handleRedirectPromise: 865,
    acquireTokenByCode: 866,
    acquireTokenSilent_silentFlow: 61,
    logout: 961,
    logoutPopup: 962,
    hydrateCache: 963,
    loadExternalTokens: 964,
} as const;

export const AzureCloudInstance = {
    None: "none",
    AzurePublic: "https://login.microsoftonline.com",
    AzurePpe: "https://login.windows-ppe.net",
    AzureChina: "https://login.chinacloudapi.cn",
    AzureGermany: "https://login.microsoftonline.de",
    AzureUsGovernment: "https://login.microsoftonline.us",
} as const;

export const JsonWebTokenTypes = {
    Jwt: "JWT",
    Jwk: "JWK",
    Pop: "pop",
} as const;

export const ResponseMode = {
    QUERY: "query",
    FRAGMENT: "fragment",
    FORM_POST: "form_post",
} as const;

export const DEFAULT_IFRAME_TIMEOUT_MS = 10000;

export const BrowserRootPerformanceEvents = {
    AcquireTokenByCode: "acquireTokenByCode",
    AcquireTokenPopup: "acquireTokenPopup",
    AcquireTokenPreRedirect: "acquireTokenPreRedirect",
    AcquireTokenRedirect: "acquireTokenRedirect",
    AcquireTokenSilent: "acquireTokenSilent",
    InitializeClientApplication: "initializeClientApplication",
    LoadExternalTokens: "loadExternalTokens",
    LocalStorageUpdated: "localStorageUpdated",
    SsoCapable: "ssoCapable",
    SsoSilent: "ssoSilent",
    WaitForBridgeLateResponse: "waitForBridgeLateResponse",
} as const;

// ---- stubbedPublicClientApplication ----

const stubbed = () =>
    new BrowserConfigurationAuthError(
        "stubbed_public_client_application_called"
    );
const stubbedReject = () => Promise.reject(stubbed());

/** real's do-nothing default instance (msal-react's pre-provider value) */
export const stubbedPublicClientApplication = {
    initialize: stubbedReject,
    acquireTokenPopup: stubbedReject,
    acquireTokenRedirect: stubbedReject,
    acquireTokenSilent: stubbedReject,
    acquireTokenByCode: stubbedReject,
    getAllAccounts: () => [],
    getAccount: () => null,
    handleRedirectPromise: stubbedReject,
    loginPopup: stubbedReject,
    loginRedirect: stubbedReject,
    logoutRedirect: stubbedReject,
    logoutPopup: stubbedReject,
    ssoSilent: stubbedReject,
    addEventCallback: () => null,
    removeEventCallback: () => {},
    addPerformanceCallback: () => "",
    removePerformanceCallback: () => false,
    getLogger: () => {
        throw stubbed();
    },
    setLogger: () => {},
    setActiveAccount: () => {},
    getActiveAccount: () => null,
    initializeWrapperLibrary: () => {},
    setNavigationClient: () => {},
    getConfiguration: () => {
        throw stubbed();
    },
    hydrateCache: stubbedReject,
    clearCache: stubbedReject,
};

// ---- AuthenticationHeaderParser ----

/** parses SHR nonces out of Authentication-Info / WWW-Authenticate headers */
export class AuthenticationHeaderParser {
    constructor(private headers: Record<string, string>) {}
    getShrNonce(): string {
        const authenticationInfo = this.headers["Authentication-Info"];
        if (authenticationInfo) {
            const c = this.parseChallenges(authenticationInfo);
            if (c.nextnonce) {
                return c.nextnonce;
            }
            throw new ClientConfigurationError("invalid_authentication_header");
        }
        const wwwAuthenticate = this.headers["WWW-Authenticate"];
        if (wwwAuthenticate) {
            const c = this.parseChallenges(wwwAuthenticate);
            if (c.nonce) {
                return c.nonce;
            }
            throw new ClientConfigurationError("invalid_authentication_header");
        }
        throw new ClientConfigurationError(
            "missing_nonce_authentication_header"
        );
    }
    parseChallenges(header: string): Record<string, string> {
        const map: Record<string, string> = {};
        for (const challenge of header
            .substring(header.indexOf(" ") + 1)
            .split(",")) {
            const [key, value] = challenge.split("=");
            map[key] = unescape(value.replace(/['"]+/g, ""));
        }
        return map;
    }
}

// ---- EventMessageUtils / EventHandler ----

export class EventMessageUtils {
    /** derives the wrapper-library InteractionStatus from an event */
    static getInteractionStatusFromEvent(
        message: EventMessage,
        currentStatus?: string
    ): string | null {
        switch (message.eventType) {
            case EventType.ACQUIRE_TOKEN_START:
                if (
                    message.interactionType === InteractionType.Redirect ||
                    message.interactionType === InteractionType.Popup
                ) {
                    return InteractionStatus.AcquireToken;
                }
                break;
            case EventType.HANDLE_REDIRECT_START:
                return InteractionStatus.HandleRedirect;
            case EventType.LOGOUT_START:
                return InteractionStatus.Logout;
            case EventType.LOGOUT_END:
                if (currentStatus && currentStatus !== InteractionStatus.Logout) {
                    break;
                }
                return InteractionStatus.None;
            case EventType.HANDLE_REDIRECT_END:
                if (
                    currentStatus &&
                    currentStatus !== InteractionStatus.HandleRedirect
                ) {
                    break;
                }
                return InteractionStatus.None;
            case EventType.ACQUIRE_TOKEN_SUCCESS:
            case EventType.ACQUIRE_TOKEN_FAILURE:
            case EventType.RESTORE_FROM_BFCACHE:
                if (
                    message.interactionType === InteractionType.Redirect ||
                    message.interactionType === InteractionType.Popup
                ) {
                    if (
                        currentStatus &&
                        currentStatus !== InteractionStatus.AcquireToken
                    ) {
                        break;
                    }
                    return InteractionStatus.None;
                }
                break;
        }
        return null;
    }
}

type EventCallback = (message: EventMessage) => void;

/** real's standalone event dispatcher (per-callback event-type filters,
 * cross-tab broadcast of login/logout/active-account events) */
export class EventHandler {
    private eventCallbacks = new Map<string, [EventCallback, string[]]>();
    private logger: Logger;
    private broadcastChannel?: BroadcastChannel;
    constructor(logger?: Logger) {
        this.logger = logger || new Logger({});
        if (typeof BroadcastChannel !== "undefined") {
            this.broadcastChannel = new BroadcastChannel(
                "msal.broadcast.event"
            );
        }
        this.invokeCrossTabCallbacks = this.invokeCrossTabCallbacks.bind(this);
    }
    addEventCallback(
        callback: EventCallback,
        eventTypes?: string[],
        callbackId?: string
    ): string | null {
        if (typeof window !== "undefined") {
            const id = callbackId || crypto.randomUUID();
            if (this.eventCallbacks.has(id)) {
                this.logger.error(
                    `Event callback with id: ${id} is already registered`
                );
                return null;
            }
            this.eventCallbacks.set(id, [callback, eventTypes || []]);
            return id;
        }
        return null;
    }
    removeEventCallback(callbackId: string): void {
        this.eventCallbacks.delete(callbackId);
    }
    emitEvent(
        eventType: string,
        correlationId?: string,
        interactionType?: string,
        payload?: unknown,
        error?: unknown
    ): void {
        const message = {
            eventType,
            interactionType: interactionType || null,
            payload: payload || null,
            error: error || null,
            correlationId,
            timestamp: Date.now(),
        } as EventMessage;
        switch (eventType) {
            case EventType.LOGIN_SUCCESS:
            case EventType.LOGOUT_SUCCESS:
            case EventType.ACTIVE_ACCOUNT_CHANGED:
                // send to other open tabs / MSAL instances on same domain
                this.broadcastChannel?.postMessage(message);
        }
        this.invokeCallbacks(message);
    }
    invokeCallbacks(message: EventMessage): void {
        this.eventCallbacks.forEach(([callback, eventTypes]) => {
            if (
                eventTypes.length === 0 ||
                eventTypes.includes(message.eventType)
            ) {
                callback(message);
            }
        });
    }
    invokeCrossTabCallbacks(event: MessageEvent): void {
        this.invokeCallbacks(event.data);
    }
    subscribeCrossTab(): void {
        this.broadcastChannel?.addEventListener(
            "message",
            this.invokeCrossTabCallbacks
        );
    }
    unsubscribeCrossTab(): void {
        this.broadcastChannel?.removeEventListener(
            "message",
            this.invokeCrossTabCallbacks
        );
    }
}

// ---- storage classes (real's IWindowStorage implementations) ----

export class MemoryStorage {
    private cache = new Map<string, string>();
    async initialize(): Promise<void> {}
    getItem(key: string): string | null {
        return this.cache.get(key) || null;
    }
    getUserData(key: string): string | null {
        return this.getItem(key);
    }
    setItem(key: string, value: string): void {
        this.cache.set(key, value);
    }
    async setUserData(key: string, value: string): Promise<void> {
        this.setItem(key, value);
    }
    removeItem(key: string): void {
        this.cache.delete(key);
    }
    getKeys(): string[] {
        return [...this.cache.keys()];
    }
    containsKey(key: string): boolean {
        return this.cache.has(key);
    }
    clear(): void {
        this.cache.clear();
    }
    decryptData(): Promise<null> {
        return Promise.resolve(null);
    }
}

export class SessionStorage {
    constructor() {
        if (!window.sessionStorage) {
            throw new BrowserConfigurationAuthError("storage_not_supported");
        }
    }
    async initialize(): Promise<void> {}
    getItem(key: string): string | null {
        return window.sessionStorage.getItem(key);
    }
    getUserData(key: string): string | null {
        return this.getItem(key);
    }
    setItem(key: string, value: string): void {
        window.sessionStorage.setItem(key, value);
    }
    async setUserData(key: string, value: string): Promise<void> {
        this.setItem(key, value);
    }
    removeItem(key: string): void {
        window.sessionStorage.removeItem(key);
    }
    getKeys(): string[] {
        return Object.keys(window.sessionStorage);
    }
    containsKey(key: string): boolean {
        return Object.prototype.hasOwnProperty.call(
            window.sessionStorage,
            key
        );
    }
    decryptData(): Promise<null> {
        return Promise.resolve(null);
    }
}

interface EncryptedData {
    id: string;
    nonce: string;
    data: string;
    lastUpdatedAt?: string;
}

/** real's encrypted-localStorage backend: plaintext getItem/setItem,
 * encrypted user data behind an in-memory mirror synced across tabs —
 * the same at-rest scheme as ./local-storage (they interoperate) */
export class LocalStorage {
    private memoryStorage = new MemoryStorage();
    private broadcast: BroadcastChannel;
    private initialized = false;
    private keyId = "";
    private baseKey!: CryptoKey;
    constructor(
        private clientId: string,
        private logger?: Logger,
        private performanceClient?: unknown
    ) {
        if (!window.localStorage) {
            throw new BrowserConfigurationAuthError("storage_not_supported");
        }
        this.broadcast = new BroadcastChannel("msal.broadcast.cache");
    }
    async initialize(correlationId?: string): Promise<void> {
        ({ keyId: this.keyId, baseKey: this.baseKey } =
            await loadEncryptionCookie());
        await this.importExistingCache(correlationId);
        this.broadcast.addEventListener("message", (event) => {
            this.updateCache(event, correlationId);
        });
        this.initialized = true;
    }
    private guard(): void {
        if (!this.initialized) {
            throw new BrowserAuthError(
                "uninitialized_public_client_application"
            );
        }
    }
    getItem(key: string): string | null {
        return window.localStorage.getItem(key);
    }
    getUserData(key: string): string | null {
        this.guard();
        return this.memoryStorage.getItem(key);
    }
    async decryptData(
        key: string,
        data: EncryptedData,
        correlationId?: string
    ): Promise<object | null> {
        this.guard();
        if (data.id !== this.keyId) {
            // encrypted with a previous session's key
            return null;
        }
        const decrypted = await decryptEntry(
            this.baseKey,
            this.keyId,
            this.getContext(key),
            JSON.stringify(data)
        );
        if (!decrypted) {
            return null;
        }
        try {
            return { ...JSON.parse(decrypted), lastUpdatedAt: data.lastUpdatedAt };
        } catch {
            return null;
        }
    }
    setItem(key: string, value: string): void {
        window.localStorage.setItem(key, value);
    }
    async setUserData(
        key: string,
        value: string,
        correlationId?: string,
        timestamp?: string,
        kmsi?: boolean
    ): Promise<void> {
        this.guard();
        if (kmsi) {
            // KMSI entities persist plaintext (survive cookie loss)
            this.setItem(key, value);
        } else {
            this.setItem(
                key,
                await encryptEntry(
                    this.baseKey,
                    this.keyId,
                    this.getContext(key),
                    value,
                    timestamp ?? String(Date.now())
                )
            );
        }
        this.memoryStorage.setItem(key, value);
        this.broadcast.postMessage({
            key,
            value,
            context: this.getContext(key),
        });
    }
    removeItem(key: string): void {
        if (this.memoryStorage.containsKey(key)) {
            this.memoryStorage.removeItem(key);
            this.broadcast.postMessage({
                key,
                value: null,
                context: this.getContext(key),
            });
        }
        window.localStorage.removeItem(key);
    }
    getKeys(): string[] {
        return Object.keys(window.localStorage);
    }
    containsKey(key: string): boolean {
        return Object.prototype.hasOwnProperty.call(window.localStorage, key);
    }
    /** removes all known MSAL keys */
    clear(): void {
        this.memoryStorage.clear();
        this.getKeys().forEach((key) => {
            if (key.startsWith("msal") || key.includes(this.clientId)) {
                this.removeItem(key);
            }
        });
    }
    /** decrypt all known MSAL entities into the in-memory mirror, pruning
     * what this session's key can't read, and write the indexes back */
    async importExistingCache(correlationId?: string): Promise<void> {
        const acctIdx = "msal.3.account.keys";
        const accounts = await this.importArray(
            this.readIndex<string[]>(acctIdx) ?? [],
            correlationId
        );
        this.writeIndex(acctIdx, accounts.length > 0, accounts);
        const tokIdx = `msal.3.token.keys.${this.clientId}`;
        const t = this.readIndex<Record<string, string[]>>(tokIdx);
        const tokens = {
            idToken: await this.importArray(t?.idToken ?? [], correlationId),
            accessToken: await this.importArray(
                t?.accessToken ?? [],
                correlationId
            ),
            refreshToken: await this.importArray(
                t?.refreshToken ?? [],
                correlationId
            ),
        };
        this.writeIndex(
            tokIdx,
            tokens.idToken.length +
                tokens.accessToken.length +
                tokens.refreshToken.length >
                0,
            tokens
        );
    }
    private readIndex<T>(key: string): T | null {
        try {
            return JSON.parse(this.getItem(key) ?? "null");
        } catch {
            return null;
        }
    }
    private writeIndex(key: string, live: boolean, value: unknown): void {
        live
            ? this.setItem(key, JSON.stringify(value))
            : window.localStorage.removeItem(key);
    }
    /** decrypt one entity; unencrypted data comes back as-is, foreign/corrupt
     * entries as null */
    async getItemFromEncryptedCache(
        key: string,
        correlationId?: string
    ): Promise<string | null> {
        return decryptEntry(
            this.baseKey,
            this.keyId,
            this.getContext(key),
            this.getItem(key)
        );
    }
    /** decrypt an array of cache keys; returns the keys successfully
     * imported, removing the rest */
    async importArray(
        arr: string[],
        correlationId?: string
    ): Promise<string[]> {
        const imported: string[] = [];
        for (const key of arr) {
            const value = await this.getItemFromEncryptedCache(
                key,
                correlationId
            );
            if (value) {
                this.memoryStorage.setItem(key, value);
                imported.push(key);
            } else {
                this.removeItem(key);
            }
        }
        return imported;
    }
    /** encryption context: clientId for app-specific entries, "" for shared */
    getContext(key: string): string {
        return key.includes(this.clientId) ? this.clientId : "";
    }
    /** cross-tab mirror sync (real's broadcast listener) */
    updateCache(event: MessageEvent, correlationId?: string): void {
        const { key, value, context } = event.data ?? {};
        if (!key || (context && context !== this.clientId)) {
            return;
        }
        value
            ? this.memoryStorage.setItem(key, value)
            : this.memoryStorage.removeItem(key);
    }
}

// ---- perf helpers ----

/** real's performance.mark/measure wrapper for perf-timeline events */
export class BrowserPerformanceMeasurement {
    measureName: string;
    startMark: string;
    endMark: string;
    constructor(
        name: string,
        private correlationId: string
    ) {
        this.measureName = BrowserPerformanceMeasurement.makeMeasureName(
            name,
            correlationId
        );
        this.startMark = BrowserPerformanceMeasurement.makeStartMark(
            name,
            correlationId
        );
        this.endMark = BrowserPerformanceMeasurement.makeEndMark(
            name,
            correlationId
        );
    }
    static makeMeasureName(name: string, correlationId: string): string {
        return `msal.measure.${name}.${correlationId}`;
    }
    static makeStartMark(name: string, correlationId: string): string {
        return `msal.start.${name}.${correlationId}`;
    }
    static makeEndMark(name: string, correlationId: string): string {
        return `msal.end.${name}.${correlationId}`;
    }
    static supportsBrowserPerformance(): boolean {
        return (
            typeof window !== "undefined" &&
            typeof window.performance !== "undefined" &&
            typeof window.performance.mark === "function" &&
            typeof window.performance.measure === "function" &&
            typeof window.performance.clearMarks === "function" &&
            typeof window.performance.clearMeasures === "function" &&
            typeof window.performance.getEntriesByName === "function"
        );
    }
    static flushMeasurements(
        correlationId: string,
        measurements: { name: string }[]
    ): void {
        if (BrowserPerformanceMeasurement.supportsBrowserPerformance()) {
            try {
                measurements.forEach((m) => {
                    const measureName =
                        BrowserPerformanceMeasurement.makeMeasureName(
                            m.name,
                            correlationId
                        );
                    if (
                        performance.getEntriesByName(measureName, "measure")
                            .length > 0
                    ) {
                        performance.clearMeasures(measureName);
                        performance.clearMarks(
                            BrowserPerformanceMeasurement.makeStartMark(
                                measureName,
                                correlationId
                            )
                        );
                        performance.clearMarks(
                            BrowserPerformanceMeasurement.makeEndMark(
                                measureName,
                                correlationId
                            )
                        );
                    }
                });
            } catch {
                /* silently catch, like real */
            }
        }
    }
    startMeasurement(): void {
        if (BrowserPerformanceMeasurement.supportsBrowserPerformance()) {
            try {
                performance.mark(this.startMark);
            } catch {}
        }
    }
    endMeasurement(): void {
        if (BrowserPerformanceMeasurement.supportsBrowserPerformance()) {
            try {
                performance.mark(this.endMark);
                performance.measure(
                    this.measureName,
                    this.startMark,
                    this.endMark
                );
            } catch {}
        }
    }
    flushMeasurement(): number | null {
        if (BrowserPerformanceMeasurement.supportsBrowserPerformance()) {
            try {
                const entries = performance.getEntriesByName(
                    this.measureName,
                    "measure"
                );
                if (entries.length > 0) {
                    const durationMs = entries[0].duration;
                    performance.clearMeasures(this.measureName);
                    performance.clearMarks(this.startMark);
                    performance.clearMarks(this.endMark);
                    return durationMs;
                }
            } catch {}
        }
        return null;
    }
}

/** real's no-op IPerformanceClient */
export class StubPerformanceClient {
    generateId(): string {
        return "callback-id";
    }
    startMeasurement(measureName: string, correlationId?: string) {
        return {
            end: () => null,
            discard: () => {},
            add: () => {},
            increment: () => {},
            event: {
                eventId: this.generateId(),
                status: 1, // PerformanceEventStatus.InProgress
                authority: "",
                libraryName: "",
                libraryVersion: "",
                clientId: "",
                name: measureName,
                startTimeMs: Date.now(),
                correlationId: correlationId || "",
            },
        };
    }
    endMeasurement(): null {
        return null;
    }
    discardMeasurements(): void {}
    removePerformanceCallback(): boolean {
        return true;
    }
    addPerformanceCallback(): string {
        return "";
    }
    emitEvents(): void {}
    addFields(): void {}
    addGlobalFields(): void {}
    incrementFields(): void {}
    cacheEventByCorrelationId(): void {}
}

// ---- enforceResourceParameter (MCP resource validation) ----

const containsResourceParam = (params?: object): boolean =>
    !!params && Object.prototype.hasOwnProperty.call(params, "resource");

export function enforceResourceParameter(
    isMcp: boolean | undefined,
    request: {
        resource?: string;
        correlationId?: string;
        extraParameters?: object;
        extraQueryParameters?: object;
    }
): void {
    if (!isMcp) {
        return;
    }
    if (
        request.resource &&
        (containsResourceParam(request.extraParameters) ||
            containsResourceParam(request.extraQueryParameters))
    ) {
        throw new ClientAuthError("misplaced_resource_parameter");
    }
    if (!request.resource) {
        throw new ClientAuthError("resource_parameter_required");
    }
}

// ---- BrowserUtils namespace ----

type PerfLike = {
    startMeasurement(name: string, cid?: string): {
        end(o?: object, e?: unknown): unknown;
    };
    incrementFields?(fields: object, cid?: string): void;
};

const parseAuthResponseFromUrl = () => {
    const urlHash = window.location.hash;
    const urlQuery = window.location.search;
    let hasResponseInHash = false;
    let hasResponseInQuery = false;
    let payload = "";
    let params: URLSearchParams | undefined;
    if (urlHash && urlHash.length > 1) {
        const hashContent = urlHash.replace(/^#/, "");
        const hashParams = new URLSearchParams(hashContent);
        if (hashParams.has("state")) {
            hasResponseInHash = true;
            payload = hashContent;
            params = hashParams;
        }
    }
    if (urlQuery && urlQuery.length > 1) {
        const queryContent = urlQuery.replace(/^\?/, "");
        const queryParams = new URLSearchParams(queryContent);
        if (queryParams.has("state")) {
            hasResponseInQuery = true;
            payload = queryContent;
            params = queryParams;
        }
    }
    if (hasResponseInHash && hasResponseInQuery) {
        // hybrid response format: combine both
        payload = `${urlQuery.replace(/^\?/, "")}${urlHash.replace(/^#/, "")}`;
        params = new URLSearchParams(payload);
    }
    if (!payload || !params) {
        throw new BrowserAuthError("empty_response");
    }
    const state = params.get("state");
    if (!state) {
        throw new BrowserAuthError("no_state_in_hash");
    }
    // real's ProtocolUtils.parseRequestState: lib state is the b64 JSON
    // before the "|" user-state delimiter
    let libraryState: { id?: string; meta?: Record<string, string> };
    try {
        libraryState = JSON.parse(atob(decodeURIComponent(state).split("|")[0]));
    } catch {
        throw new ClientAuthError("invalid_state");
    }
    const { id, meta } = libraryState;
    if (!id || !meta) {
        throw new BrowserAuthError(
            "unable_to_parse_state",
            "missing_library_state"
        );
    }
    return {
        params,
        payload,
        urlHash,
        urlQuery,
        hasResponseInHash,
        hasResponseInQuery,
        libraryState: { id, meta },
    };
};

const isInIframe = () => window.parent !== window;

const isInPopup = () => {
    if (isInIframe()) {
        return false;
    }
    try {
        return (
            parseAuthResponseFromUrl().libraryState.meta["interactionType"] ===
            InteractionType.Popup
        );
    } catch {
        return false;
    }
};

const blockReloadInHiddenIframes = () => {
    // a response hash in a hidden iframe means a silent call reloaded us
    if (/[#&](code|error)=/.test(window.location.hash) && isInIframe()) {
        throw new BrowserAuthError("block_iframe_reload");
    }
};

const blockNonBrowserEnvironment = () => {
    if (typeof window === "undefined") {
        throw new BrowserAuthError("non_browser_environment");
    }
};

const blockAPICallsBeforeInitialize = (initialized: boolean) => {
    if (!initialized) {
        throw new BrowserAuthError("uninitialized_public_client_application");
    }
};

const preflightCheck = (initialized: boolean) => {
    blockNonBrowserEnvironment();
    blockReloadInHiddenIframes();
    if (isInPopup()) {
        throw new BrowserAuthError("block_nested_popups");
    }
    blockAPICallsBeforeInitialize(initialized);
};

// pending bridge monitor, cancellable when an interaction is overridden
let activeBridgeMonitor: {
    timeoutId: number;
    channel: BroadcastChannel;
    reject: (e: unknown) => void;
} | null = null;

/** real's exported utility namespace */
export const BrowserUtils = {
    blockAPICallsBeforeInitialize,
    blockAcquireTokenInPopups: () => {
        if (isInPopup()) {
            throw new BrowserAuthError("block_nested_popups");
        }
    },
    blockNonBrowserEnvironment,
    blockRedirectInIframe: (allowRedirectInIframe: boolean) => {
        if (isInIframe() && !allowRedirectInIframe) {
            throw new BrowserAuthError("redirect_in_iframe");
        }
    },
    blockReloadInHiddenIframes,
    buildMergedClaims: (
        claims: string | undefined,
        clientCapabilities?: string[],
        correlationId = ""
    ): string => {
        let c: Record<string, any> = {};
        if (claims) {
            try {
                const parsed = JSON.parse(claims);
                if (
                    typeof parsed !== "object" ||
                    parsed === null ||
                    Array.isArray(parsed)
                ) {
                    throw 0;
                }
                c = parsed;
            } catch {
                throw new ClientConfigurationError("invalid_claims");
            }
        }
        c.id_token ??= {};
        for (const [k, v] of Object.entries({
            signin_state: { essential: false },
            login_hint: { essential: false },
        })) {
            if (!(k in c.id_token)) {
                c.id_token[k] = v;
            }
        }
        if (clientCapabilities?.length) {
            (c.access_token ??= {}).xms_cc ??= { values: clientCapabilities };
        }
        return JSON.stringify(c);
    },
    cancelPendingBridgeResponse: (logger: Logger, correlationId?: string) => {
        if (activeBridgeMonitor) {
            clearTimeout(activeBridgeMonitor.timeoutId);
            activeBridgeMonitor.channel.close();
            activeBridgeMonitor.reject(
                new BrowserAuthError("interaction_in_progress_cancelled")
            );
            activeBridgeMonitor = null;
        }
    },
    clearAuthResponseFromUrl: (contentWindow: Window) => {
        if (typeof contentWindow.history?.replaceState === "function") {
            contentWindow.history.replaceState(
                null,
                "",
                `${contentWindow.location.origin}${contentWindow.location.pathname}`
            );
        }
    },
    clearHash: (contentWindow: Window) => {
        contentWindow.location.hash = "";
        // Office.js sets history.replaceState to null
        if (typeof contentWindow.history.replaceState === "function") {
            contentWindow.history.replaceState(
                null,
                "",
                `${contentWindow.location.origin}${contentWindow.location.pathname}${contentWindow.location.search}`
            );
        }
    },
    createGuid: () => crypto.randomUUID(),
    getCurrentUri: () =>
        typeof window !== "undefined" && window.location
            ? window.location.href.split("?")[0].split("#")[0]
            : "",
    getHomepage: (correlationId?: string) => {
        const u = new URL(window.location.href);
        return `${u.protocol}//${u.host}/`;
    },
    invoke:
        <A extends unknown[], R>(
            callback: (...args: A) => R,
            eventName: string,
            logger: Logger,
            telemetryClient: PerfLike,
            correlationId?: string
        ) =>
        (...args: A): R => {
            logger.trace(`Executing function '${eventName}'`, correlationId);
            const inProgressEvent = telemetryClient.startMeasurement(
                eventName,
                correlationId
            );
            if (correlationId) {
                telemetryClient.incrementFields?.(
                    { [`ext.${eventName}CallCount`]: 1 },
                    correlationId
                );
            }
            try {
                const result = callback(...args);
                inProgressEvent.end({ success: true });
                return result;
            } catch (e) {
                inProgressEvent.end({ success: false }, e);
                throw e;
            }
        },
    invokeAsync:
        <A extends unknown[], R>(
            callback: (...args: A) => Promise<R>,
            eventName: string,
            logger: Logger,
            telemetryClient: PerfLike,
            correlationId?: string
        ) =>
        async (...args: A): Promise<R> => {
            logger.trace(`Executing function '${eventName}'`, correlationId);
            const inProgressEvent = telemetryClient.startMeasurement(
                eventName,
                correlationId
            );
            if (correlationId) {
                telemetryClient.incrementFields?.(
                    { [`ext.${eventName}CallCount`]: 1 },
                    correlationId
                );
            }
            try {
                const result = await callback(...args);
                inProgressEvent.end({ success: true });
                return result;
            } catch (e) {
                inProgressEvent.end({ success: false }, e);
                throw e;
            }
        },
    isInIframe,
    isInPopup,
    parseAuthResponseFromUrl,
    preconnect: (authority: string) => {
        const link = document.createElement("link");
        link.rel = "preconnect";
        link.href = new URL(authority).origin;
        link.crossOrigin = "anonymous";
        document.head.appendChild(link);
        // the browser closes unused connections after a few seconds
        window.setTimeout(() => {
            try {
                document.head.removeChild(link);
            } catch {}
        }, 10000);
    },
    preflightCheck,
    redirectPreflightCheck: (
        initialized: boolean,
        config: {
            system?: { allowRedirectInIframe?: boolean };
            cache?: { cacheLocation?: string };
        }
    ) => {
        preflightCheck(initialized);
        BrowserUtils.blockRedirectInIframe(
            !!config.system?.allowRedirectInIframe
        );
        if (config.cache?.cacheLocation === "memoryStorage") {
            throw new BrowserConfigurationAuthError(
                "in_mem_redirect_unavailable"
            );
        }
    },
    replaceHash: (url: string) => {
        const urlParts = url.split("#");
        urlParts.shift();
        window.location.hash = urlParts.length > 0 ? urlParts.join("#") : "";
    },
    waitForBridgeResponse: (
        timeoutMs: number,
        logger: Logger,
        request: { state?: string; correlationId?: string },
        performanceClient?: unknown
    ): Promise<string> =>
        new Promise((resolve, reject) => {
            let id = "";
            try {
                id = JSON.parse(
                    atob((request.state ?? "").split("|")[0])
                ).id;
            } catch {}
            const channel = new BroadcastChannel(id);
            const timeoutId = window.setTimeout(() => {
                activeBridgeMonitor = null;
                channel.close();
                reject(
                    new BrowserAuthError("timed_out", "redirect_bridge_timeout")
                );
            }, timeoutMs);
            activeBridgeMonitor = { timeoutId, channel, reject };
            channel.onmessage = (event) => {
                activeBridgeMonitor = null;
                clearTimeout(timeoutId);
                channel.close();
                const responseString = event.data?.payload;
                responseString
                    ? resolve(responseString)
                    : reject(
                          new BrowserAuthError("redirect_bridge_empty_response")
                      );
            };
        }),
};

// ---- SignedHttpRequest ----

/** real's standalone SHR helper: generate a binding keypair, sign payloads
 * with it, drop it — reuses ./pop's keystore (same IndexedDB msal.db) */
export class SignedHttpRequest {
    private mem = new Map<string, BoundKeyPair>();
    constructor(
        private shrParameters: {
            resourceRequestMethod?: string;
            resourceRequestUri?: string;
            shrClaims?: string;
            shrNonce?: string;
            shrOptions?: { header?: { typ?: string } };
        },
        shrOptions?: { loggerOptions?: object }
    ) {}
    /** generates + caches a keypair; returns the public key digest (kid) */
    async generatePublicKeyThumbprint(): Promise<string> {
        const { kid, entry } = await makeBoundKeyPair(
            this.shrParameters as TokenRequest
        );
        await keystore.put(this.mem, kid, entry);
        return kid;
    }
    /** signs a payload (e.g. an access token) with the given key */
    async signRequest(
        payload: string,
        publicKeyThumbprint: string,
        claims?: object
    ): Promise<string> {
        const pair = await keystore.get(this.mem, publicKeyThumbprint);
        if (!pair) {
            throw new BrowserAuthError("crypto_key_not_found");
        }
        return signPop(
            pair,
            publicKeyThumbprint,
            payload,
            this.shrParameters as TokenRequest,
            claims
        );
    }
    /** removes the cached keypair for the given kid */
    async removeKeys(
        publicKeyThumbprint: string,
        correlationId?: string
    ): Promise<void> {
        await keystore.del(this.mem, publicKeyThumbprint);
    }
}
