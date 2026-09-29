/**
 * pdf.js's worker in a separate Node process, so what pdf.js does inside one stream cannot take the
 * host with it. A PDF of 100 KB can hold a content stream that inflates to a 100-million-character
 * string, which pdf.js reads into gigabytes of glyphs before any of it reaches the parser: in-process,
 * that ended the host with a fatal out-of-memory error no caller can catch (a worker thread's memory
 * limit did not hold either, for one allocation that large). In its own process under a heap limit
 * (`pdfParserConfig.processMemoryMb`), the same file ends that process, and the parse rejects with
 * PDF_PROCESS_FAILED. Node only; the browser runs pdf.js in a Web Worker of its own.
 *
 * The processes are a pool of at most one per CPU, each reading one document at a time: a process
 * per parse started 101 for 100 parses at once (7.5 GB), each slowed by the others. A parse finding
 * every process busy waits for one. Each process reports the CPU time it has spent (see
 * CLOCK_SOURCE), which is what the time budget charges a document.
 */

import type { ChildProcess } from 'child_process';
import type { Readable } from 'stream';
import { getAbortError } from '../../utils/errorUtils.js';

/**
 * Typed arrays arrive over IPC as views into the message's own buffer (at an offset); pdf.js reads a
 * view's whole buffer, so each is copied to one of its own. Used at both ends: here, and in the child
 * as OWN_ARRAYS_SOURCE (the same function as source: a compiled function's own text can carry a
 * bundler's helpers, which the child does not have).
 */
const ownArrays = (value: any): any => {
    const own = (item: any, depth: number): any => {
        if (ArrayBuffer.isView(item)) return item.byteOffset || item.byteLength !== item.buffer.byteLength ? (item as any).slice() : item;
        if (depth > 64 || item === null || typeof item !== 'object') return item;
        if (Array.isArray(item)) { for (let i = 0; i < item.length; i++) item[i] = own(item[i], depth + 1); return item; }
        // Only a changed value is written, as an own property: a key named `__proto__` in a message is
        // data, never the prototype.
        if (Object.getPrototypeOf(item) === Object.prototype) for (const key of Object.keys(item)) { const next = own(item[key], depth + 1); if (next !== item[key]) Object.defineProperty(item, key, { value: next, writable: true, enumerable: true, configurable: true }); }
        return item;
    };
    return own(value, 0);
};

const OWN_ARRAYS_SOURCE = `(value) => {
    const own = (item, depth) => {
        if (ArrayBuffer.isView(item)) return item.byteOffset || item.byteLength !== item.buffer.byteLength ? item.slice() : item;
        if (depth > 64 || item === null || typeof item !== 'object') return item;
        if (Array.isArray(item)) { for (let i = 0; i < item.length; i++) item[i] = own(item[i], depth + 1); return item; }
        if (Object.getPrototypeOf(item) === Object.prototype) for (const key of Object.keys(item)) { const next = own(item[key], depth + 1); if (next !== item[key]) Object.defineProperty(item, key, { value: next, writable: true, enumerable: true, configurable: true }); }
        return item;
    };
    return own(value, 0);
}`;

/**
 * The child's clock, in a thread of its own: every 20 ms it writes the process's CPU time, in ms, to
 * fd 4 when it has changed. pdf.js can hold the child's main thread for all of one request (a page
 * drawing one 1 MB string hundreds of times never returns to the event loop), so only another thread
 * can tell the host, mid-request, how much work the document has cost; and CPU time is what the host
 * charges, since no wait of the host's, nor another process taking the CPU, adds to it. A reading the
 * host can no longer receive means the host is gone: the child ends itself rather than finish a
 * document for no one (a busy child outlived its parent by seconds at full CPU).
 */
const CLOCK_SOURCE = `
const { writeSync } = require('fs');
let last = -1;
setInterval(() => {
    const usage = process.cpuUsage();
    const ms = Math.floor((usage.user + usage.system) / 1000);
    if (ms === last) return;
    last = ms;
    try { writeSync(4, ms + '\\n'); } catch (error) { if (error.code !== 'EAGAIN') process.kill(process.pid, 'SIGKILL'); }
}, 20);
`;

/** The child: pdf.js's worker, reading and writing its messages through the IPC channel. */
const CHILD_SOURCE = `
const ownArrays = ${OWN_ARRAYS_SOURCE};
// Arrays sent to the host hold at most 100,000 entries: a string drawn once can be a glyph array of
// fifty million in the operator list, and reading that message blocked the host for seconds. Nothing
// the host reads is that long (image data are typed arrays, which this leaves alone).
const capArrays = (value, depth) => {
    if (depth > 64 || value === null || typeof value !== 'object' || ArrayBuffer.isView(value)) return value;
    if (Array.isArray(value)) {
        if (value.length > 100000) value = value.slice(0, 100000);
        for (let i = 0; i < value.length; i++) value[i] = capArrays(value[i], depth + 1);
        return value;
    }
    if (Object.getPrototypeOf(value) === Object.prototype) for (const key of Object.keys(value)) { const next = capArrays(value[key], depth + 1); if (next !== value[key]) Object.defineProperty(value, key, { value: next, writable: true, enumerable: true, configurable: true }); }
    return value;
};
// The clock (fd 4): its first line says whether this process can measure its CPU time ('-' if not).
let clock = null;
try { clock = new (require('worker_threads').Worker)(${JSON.stringify(CLOCK_SOURCE)}, { eval: true }); clock.unref(); } catch { clock = null; }
try {
    const usage = process.cpuUsage();
    require('fs').writeSync(4, clock ? Math.floor((usage.user + usage.system) / 1000) + '\\n' : '-\\n');
} catch { process.exit(3); }
const listeners = new Set();
const port = {
    postMessage(message) { if (process.connected) process.send(capArrays(message, 0)); },
    addEventListener(type, listener, options) {
        if (type !== 'message') return;
        listeners.add(listener);
        if (options && options.signal) options.signal.addEventListener('abort', () => listeners.delete(listener));
    },
    removeEventListener(type, listener) { listeners.delete(listener); },
};
process.on('message', (data) => { data = ownArrays(data); for (const listener of listeners) listener({ data }); });
process.on('disconnect', () => process.exit(0));
import(process.env.OFFICEPARSER_PDF_WORKER).then(
    ({ WorkerMessageHandler }) => WorkerMessageHandler.initializeFromPort(port),
    () => process.exit(3),
);
`;

/** How long a new process may take to start pdf.js before the parse runs pdf.js in this one. */
const START_TIMEOUT_MS = 15_000;
/** Documents one process reads before it is replaced, and processes kept waiting for the next parse. */
const USES_PER_PROCESS = 50;
const MAX_IDLE = 2;

export interface PdfProcess {
    /** A pdf.js `PDFWorker` whose worker is the process, for `getDocument({ worker })`. */
    worker: any;
    /** Rejects (with the reason, as text) if the process ends while in use, other than by `release` or `end`. */
    died: Promise<never>;
    readonly alive: boolean;
    /**
     * The CPU time, in ms, the process has spent so far, as it last reported it: what reading a
     * document costs, whatever else the host or the machine is doing meanwhile. Where the process
     * cannot measure it (it could not start a thread), the time elapsed instead.
     */
    clock(): number;
    /** Ends the process now (a document pdf.js must stop reading). Nothing once the process was released. */
    end(): void;
    /** Hands the process back for the next parse, or ends it. Nothing once it was released or ended. */
    release(): void;
}

interface Handle {
    child: ChildProcess;
    worker: any;
    alive: boolean;
    ending: boolean;
    uses: number;
    memoryMb: number;
    workerUrl: string;
    /** The CPU time, in ms, the process last reported; NaN when it cannot measure it, undefined before its first report. */
    cpuMs: number | undefined;
    /** Whether the process holds one of the pool's places (from its start until it ends). */
    placed: boolean;
    rejectDied?: (reason: string) => void;
}

/** A parse waiting for a process: `grant` hands it an idle one, or `null` for a place to start one in. */
interface Waiter {
    memoryMb: number;
    workerUrl: string;
    grant(handle: Handle | null): void;
}

const idle: Handle[] = [];
/** Every process started and not yet ended, ended when this process exits. */
const live = new Set<Handle>();
const waiters: Waiter[] = [];
/** Places taken: processes running or starting. At most `maxProcesses`. */
let placesTaken = 0;
let maxProcesses = 0;
let pumping = false;
let exitHooked = false;

const setRef = (handle: Handle, ref: boolean): void => {
    const channel = (handle.child as any).channel;
    if (ref) { handle.child.ref(); channel?.ref?.(); } else { handle.child.unref(); channel?.unref?.(); }
};

/** Gives back the place a process held, once, and lets the next waiting parse have it. */
const vacate = (handle: Handle): void => {
    live.delete(handle);
    const index = idle.indexOf(handle);
    if (index >= 0) idle.splice(index, 1);
    if (!handle.placed) return;
    handle.placed = false;
    placesTaken--;
    pump();
};

const kill = (handle: Handle): void => {
    handle.ending = true;
    handle.alive = false;
    try { handle.worker?.destroy?.(); } catch { /* already gone */ }
    try { handle.child.kill(); } catch { /* already gone */ }
    vacate(handle);
};

/**
 * Hands idle processes and free places to the waiting parses, first come first served: an idle
 * process started for the same pdf.js and memory limit, else a place to start one, else (at the
 * limit) an idle process started for another configuration is ended to make room.
 */
const pump = (): void => {
    if (pumping) return;
    pumping = true;
    try {
        while (waiters.length) {
            const waiter = waiters[0];
            const index = idle.findIndex(h => h.alive && h.memoryMb === waiter.memoryMb && h.workerUrl === waiter.workerUrl);
            if (index >= 0) { waiters.shift(); waiter.grant(idle.splice(index, 1)[0]); continue; }
            if (placesTaken < maxProcesses) { waiters.shift(); placesTaken++; waiter.grant(null); continue; }
            const other = idle.shift();
            if (!other) break;
            kill(other);
        }
        while (idle.length > MAX_IDLE) kill(idle.shift()!);
    } finally {
        pumping = false;
    }
};

/** Starts a process running pdf.js's worker, resolving once pdf.js in it is ready; undefined if it cannot. */
const start = async (pdfjs: any, workerUrl: string, memoryMb: number): Promise<Handle | undefined> => {
    let spawn: typeof import('child_process').spawn;
    try { ({ spawn } = await import('child_process')); } catch { return undefined; }
    const env: Record<string, string | undefined> = { ...process.env, OFFICEPARSER_PDF_WORKER: workerUrl, ELECTRON_RUN_AS_NODE: '1' };
    // The host's options (a loader, instrumentation, its own heap size) are not the child's.
    delete env.NODE_OPTIONS;
    let child: ChildProcess;
    try {
        child = spawn(process.execPath, [`--max-old-space-size=${memoryMb}`, '--input-type=commonjs', '-e', CHILD_SOURCE], {
            stdio: ['ignore', 'ignore', 'ignore', 'ipc', 'pipe'],
            serialization: 'advanced',
            env,
            windowsHide: true,
        });
    } catch {
        return undefined;
    }
    const handle: Handle = { child, worker: undefined, alive: true, ending: false, uses: 0, memoryMb, workerUrl, cpuMs: undefined, placed: false };
    live.add(handle);
    if (!exitHooked) {
        exitHooked = true;
        // A process reading a document when this one exits is ended with it, not left to finish.
        process.on('exit', () => { for (const h of live) { try { h.child.kill(); } catch { /* already gone */ } } });
    }
    let clockRead!: () => void;
    const clockKnown = new Promise<void>(resolve => { clockRead = resolve; });
    const clockStream = child.stdio[4] as Readable | null | undefined;
    if (clockStream) {
        // Only the last whole reading counts; what trails it is kept for the next chunk (a reading is a few digits).
        let partial = '';
        clockStream.setEncoding('latin1');
        clockStream.on('data', (chunk: string) => {
            const text = partial + chunk;
            const end = text.lastIndexOf('\n');
            partial = text.slice(end + 1).slice(-32);
            if (end < 0) return;
            const line = text.slice(text.lastIndexOf('\n', end - 1) + 1, end);
            const ms = line === '-' ? NaN : Number(line);
            if (line === '-' || (line !== '' && Number.isFinite(ms))) { handle.cpuMs = ms; clockRead(); }
        });
        clockStream.on('error', () => { /* the process ended */ });
        // The clock never keeps this process alive; the IPC channel does, while a parse uses the process.
        (clockStream as any).unref?.();
    }
    // The first message is pdf.js's own "ready", which pdf.js's end of the port does not wait for.
    const ready = new Promise<boolean>(resolve => {
        const timer = setTimeout(() => resolve(false), START_TIMEOUT_MS);
        const pdfReady = new Promise<void>(resolveReady => child.once('message', () => resolveReady()));
        Promise.all([pdfReady, clockKnown]).then(() => { clearTimeout(timer); resolve(true); });
        child.once('error', () => { clearTimeout(timer); resolve(false); });
        child.once('exit', () => { clearTimeout(timer); resolve(false); });
    });
    child.on('error', () => { handle.alive = false; });
    child.on('exit', (code, signal) => {
        handle.alive = false;
        vacate(handle);
        if (!handle.ending) handle.rejectDied?.(`the pdf.js process ended (${signal ?? `exit code ${code}`})`);
    });
    if (!(await ready) || !handle.alive) { kill(handle); return undefined; }
    const listeners = new Set<(event: { data: unknown }) => void>();
    child.on('message', (data: unknown) => {
        const event = { data: ownArrays(data) };
        for (const listener of listeners) listener(event);
    });
    const port = {
        postMessage(message: unknown) { if (child.connected) child.send(message as any); },
        addEventListener(type: string, listener: (event: { data: unknown }) => void, options?: { signal?: AbortSignal }) {
            if (type !== 'message') return;
            listeners.add(listener);
            options?.signal?.addEventListener('abort', () => listeners.delete(listener));
        },
        removeEventListener(_type: string, listener: (event: { data: unknown }) => void) { listeners.delete(listener); },
    };
    try { handle.worker = new pdfjs.PDFWorker({ port }); } catch { kill(handle); return undefined; }
    return handle;
};

/**
 * A process running pdf.js's worker for one parse: an idle one started for the same pdf.js and memory
 * limit, or a new one, waiting (until `signal` aborts) while every place in the pool is busy. Undefined
 * when none can start (no `child_process`, a runtime that cannot run this one's executable as Node);
 * pdf.js then runs in this process.
 *
 * What is returned is a lease: once released (or ended), its `end` and `release` do nothing, so a
 * holder acting late (an abort listener firing after its parse finished) cannot end a process that
 * another parse is now using.
 */
export const acquirePdfProcess = async (pdfjs: any, workerUrl: string, memoryMb: number, signal?: AbortSignal | null): Promise<PdfProcess | undefined> => {
    if (!maxProcesses) {
        try {
            const os = await import('os');
            maxProcesses = Math.max(1, typeof os.availableParallelism === 'function' ? os.availableParallelism() : os.cpus().length);
        } catch {
            maxProcesses = 1;
        }
    }
    const granted = await new Promise<Handle | null>((resolve, reject) => {
        if (signal?.aborted) { reject(getAbortError()); return; }
        const onAbort = () => {
            const index = waiters.indexOf(waiter);
            if (index < 0) return;
            waiters.splice(index, 1);
            reject(getAbortError());
        };
        const waiter: Waiter = { memoryMb, workerUrl, grant: (handle) => { signal?.removeEventListener('abort', onAbort); resolve(handle); } };
        signal?.addEventListener('abort', onAbort, { once: true });
        waiters.push(waiter);
        pump();
    });
    let handle = granted;
    if (!handle) {
        handle = (await start(pdfjs, workerUrl, memoryMb)) ?? null;
        if (!handle) { placesTaken--; pump(); return undefined; }
        handle.placed = true;
    }
    const current = handle;
    current.uses++;
    setRef(current, true);
    let held = true;
    const died = new Promise<never>((_, reject) => { current.rejectDied = reject; });
    // Settled or not, never an unhandled rejection: a caller races it.
    died.catch(() => { /* raced by the caller */ });
    return {
        worker: current.worker,
        died,
        get alive() { return current.alive; },
        clock: () => (current.cpuMs !== undefined && !Number.isNaN(current.cpuMs) ? current.cpuMs : performance.now()),
        end: () => {
            if (!held) return;
            held = false;
            current.rejectDied = undefined;
            kill(current);
        },
        release: () => {
            if (!held) return;
            held = false;
            current.rejectDied = undefined;
            if (current.alive && current.uses < USES_PER_PROCESS) {
                setRef(current, false);
                idle.push(current);
                pump();
            } else {
                kill(current);
            }
        },
    };
};
