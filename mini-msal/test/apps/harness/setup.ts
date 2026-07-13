/**
 * Shared conformance-harness glue. Exposes on globalThis:
 *   __lib     — the module namespace under test (real msal-browser or mini)
 *   __target  — "real" | "mini"
 *   __cap     — capture arrays: events / perf / logs (digested, stable)
 *   __create(config, opts) — construct + (optionally) initialize a client,
 *                            wire capture, expose it as __msal
 *   __serializeError(e)    — stable error digest for assertions
 *
 * The harness deliberately does NOTHING on page load (no auto
 * handleRedirectPromise) — scenarios drive everything via page.evaluate.
 */
export function setupHarness(lib: any, target: string): void {
    const g = globalThis as any;
    const cap = {
        events: [] as any[],
        perf: [] as any[],
        perfRaw: [] as any[],
        logs: [] as any[],
        seq: 0,
    };
    g.__lib = lib;
    g.__target = target;
    g.__cap = cap;

    g.__serializeError = (e: any) => ({
        name: e?.name ?? null,
        errorCode: e?.errorCode ?? null,
        subError: e?.subError || null,
        // message text is library-specific prose; keep a short prefix only
        message: String(e?.errorMessage ?? e?.message ?? e).slice(0, 120),
    });

    // navigation instrumentation (C13): URL digests keep snapshots stable
    const digestNavUrl = (url: string) => {
        const u = new URL(url, location.href);
        return {
            page: u.origin + u.pathname,
            queryKeys: [...u.searchParams.keys()].sort(),
        };
    };
    const navRecorder = {
        navigateInternal: (url: string, options: any) => {
            g.__navCalls.push({ call: "internal", url: digestNavUrl(url), options });
            return Promise.resolve(true);
        },
        navigateExternal: (url: string, options: any) => {
            g.__navCalls.push({ call: "external", url: digestNavUrl(url), options });
            return Promise.resolve(true);
        },
    };

    g.__create = async (config: any, opts: any = {}) => {
        g.__navCalls = [];
        if (opts.onRedirectNavigate === "record-cancel") {
            config.auth = {
                ...config.auth,
                onRedirectNavigate: (url: string) => {
                    g.__navCalls.push({
                        call: "onRedirectNavigate",
                        url: digestNavUrl(url),
                    });
                    return false;
                },
            };
        }
        if (opts.navigationClient === "config") {
            config.system = { ...config.system, navigationClient: navRecorder };
        }
        if (opts.captureLogs) {
            config.system = {
                ...config.system,
                loggerOptions: {
                    logLevel: 3, // Verbose
                    piiLoggingEnabled: false,
                    loggerCallback: (level: number, message: string) => {
                        cap.logs.push({ level, message });
                    },
                },
            };
        }
        if (opts.perfClient && lib.BrowserPerformanceClient) {
            // real perf events require opting in to a non-stub client
            config.telemetry = {
                ...config.telemetry,
                client: new lib.BrowserPerformanceClient(config),
            };
        }
        const pca = new lib.PublicClientApplication(config);
        pca.addEventCallback((m: any) => {
            cap.events.push({
                seq: cap.seq++,
                eventType: m.eventType,
                interactionType: m.interactionType ?? null,
                payloadKeys: m.payload
                    ? Object.keys(m.payload).sort()
                    : null,
                payloadAccount:
                    m.payload?.account?.username ??
                    m.payload?.username ??
                    null,
                error: m.error ? g.__serializeError(m.error) : null,
            });
        });
        try {
            pca.addPerformanceCallback?.((events: any[]) => {
                for (const e of events) {
                    // full-shape clone for perf-event-shape scenarios; keys
                    // with undefined values survive as null (JSON drops them)
                    const raw: any = {};
                    for (const k of Object.keys(e)) {
                        raw[k] = e[k] === undefined ? null : e[k];
                    }
                    cap.perfRaw.push(raw);
                    cap.perf.push({
                        name: e.name,
                        success: e.success,
                        hasDuration: typeof e.durationMs === "number",
                        correlationId: e.correlationId ?? null,
                        errorCode: e.errorCode ?? null,
                    });
                }
            });
        } catch {
            /* mini has no performance client */
        }
        if (opts.navigationClient === "setter") {
            pca.setNavigationClient(navRecorder);
        }
        if (!opts.noInit) {
            await pca.initialize();
        }
        g.__msal = pca;
        return target;
    };

    document.body.textContent = `conformance harness: ${target}`;
}
