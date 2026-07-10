/**
 * Telemetry feature: performance events, opt-in like real MSAL — events fire
 * only when the app passes `telemetry.client: new BrowserPerformanceClient()`
 * in config (real's default is a stub that swallows everything). Compose via
 * createClient(config, [popup, telemetry]) — telemetry LAST so it can wrap
 * the methods other features attach.
 */
import type { ClientContext, PerfClient } from "./index.js";

/** One completed top-level operation (the subset real's events observably carry). */
export interface PerformanceEvent {
    name: string;
    correlationId: string;
    durationMs: number;
    success: boolean;
    errorCode?: string;
}

type PerfCallback = (events: PerformanceEvent[]) => void;

/** The methods this feature attaches to the client. */
export interface TelemetryClient {
    addPerformanceCallback(cb: PerfCallback): string;
    removePerformanceCallback(id: string): boolean;
}

/** Opt-in perf client, config `telemetry.client` (like real's). */
export class BrowserPerformanceClient implements PerfClient {
    private cbs = new Map<string, PerfCallback>();
    constructor(_config?: unknown) {}
    addPerformanceCallback(cb: PerfCallback): string {
        const id = crypto.randomUUID();
        this.cbs.set(id, cb);
        return id;
    }
    removePerformanceCallback(id: string): boolean {
        return this.cbs.delete(id);
    }
    emitEvents(events: PerformanceEvent[]): void {
        this.cbs.forEach((cb) => cb(events));
    }
}

export function telemetry(ctx: ClientContext): void {
    const c = ctx.client;
    const pc = ctx.config.telemetry?.client;
    // registration works without an opt-in client, like real's stub — the
    // callback just never fires
    c.addPerformanceCallback = (cb: PerfCallback): string =>
        pc?.addPerformanceCallback(cb) ?? crypto.randomUUID();
    c.removePerformanceCallback = (id: string): boolean =>
        pc?.removePerformanceCallback(id) ?? false;
    if (!pc) return;
    // wrap each top-level API in a measurement, emitted on settle with real's
    // event name (cache hits included — real measures the whole call)
    const wrap = (method: string, name: string) => {
        const orig = c[method];
        if (!orig) return;
        c[method] = async (...args: any[]) => {
            const started = performance.now();
            const done = (success: boolean, cid?: string, errorCode?: string) =>
                pc.emitEvents([
                    {
                        name,
                        correlationId: cid ?? crypto.randomUUID(),
                        durationMs: Math.round(performance.now() - started),
                        success,
                        errorCode,
                    },
                ]);
            try {
                const r = await orig(...args);
                done(true, r?.correlationId);
                return r;
            } catch (e: any) {
                done(false, args[0]?.correlationId, e?.errorCode);
                throw e;
            }
        };
    };
    wrap("initialize", "initializeClientApplication");
    wrap("acquireTokenSilent", "acquireTokenSilent");
    wrap("ssoSilent", "ssoSilent");
    wrap("acquireTokenPopup", "acquireTokenPopup");
}
