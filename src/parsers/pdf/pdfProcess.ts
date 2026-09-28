/**
 * pdf.js's worker in a separate Node process, so what pdf.js does inside one stream cannot take the
 * host with it. A PDF of 100 KB can hold a content stream that inflates to a 100-million-character
 * string, which pdf.js reads into gigabytes of glyphs before any of it reaches the parser: in-process,
 * that ended the host with a fatal out-of-memory error no caller can catch (a worker thread's memory
 * limit did not hold either, for one allocation that large). In its own process under a heap limit
 * (`pdfParserConfig.processMemoryMb`), the same file ends that process, and the parse rejects with
 * PDF_PROCESS_FAILED. Node only; the browser runs pdf.js in a Web Worker of its own.
 */

import type { ChildProcess } from 'child_process';

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
        if (Object.getPrototypeOf(item) === Object.prototype) for (const key of Object.keys(item)) item[key] = own(item[key], depth + 1);
        return item;
    };
    return own(value, 0);
};

const OWN_ARRAYS_SOURCE = `(value) => {
    const own = (item, depth) => {
        if (ArrayBuffer.isView(item)) return item.byteOffset || item.byteLength !== item.buffer.byteLength ? item.slice() : item;
        if (depth > 64 || item === null || typeof item !== 'object') return item;
        if (Array.isArray(item)) { for (let i = 0; i < item.length; i++) item[i] = own(item[i], depth + 1); return item; }
        if (Object.getPrototypeOf(item) === Object.prototype) for (const key of Object.keys(item)) item[key] = own(item[key], depth + 1);
        return item;
    };
    return own(value, 0);
}`;

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
    if (Object.getPrototypeOf(value) === Object.prototype) for (const key of Object.keys(value)) value[key] = capArrays(value[key], depth + 1);
    return value;
};
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
    /** Ends the process now (a document pdf.js must stop reading). */
    end(): void;
    /** Hands the process back for the next parse, or ends it. */
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
    rejectDied?: (reason: string) => void;
}

const idle: Handle[] = [];

const setRef = (handle: Handle, ref: boolean): void => {
    const channel = (handle.child as any).channel;
    if (ref) { handle.child.ref(); channel?.ref?.(); } else { handle.child.unref(); channel?.unref?.(); }
};

const kill = (handle: Handle): void => {
    handle.ending = true;
    handle.alive = false;
    try { handle.worker?.destroy?.(); } catch { /* already gone */ }
    try { handle.child.kill(); } catch { /* already gone */ }
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
            stdio: ['ignore', 'ignore', 'ignore', 'ipc'],
            serialization: 'advanced',
            env,
            windowsHide: true,
        });
    } catch {
        return undefined;
    }
    const handle: Handle = { child, worker: undefined, alive: true, ending: false, uses: 0, memoryMb, workerUrl };
    const listeners = new Set<(event: { data: unknown }) => void>();
    const ready = new Promise<boolean>(resolve => {
        const timer = setTimeout(() => resolve(false), START_TIMEOUT_MS);
        child.once('message', () => { clearTimeout(timer); resolve(true); });
        child.once('error', () => { clearTimeout(timer); resolve(false); });
        child.once('exit', () => { clearTimeout(timer); resolve(false); });
    });
    child.on('error', () => { handle.alive = false; });
    child.on('exit', (code, signal) => {
        handle.alive = false;
        const index = idle.indexOf(handle);
        if (index >= 0) idle.splice(index, 1);
        if (!handle.ending) handle.rejectDied?.(`the pdf.js process ended (${signal ?? `exit code ${code}`})`);
    });
    // The first message is pdf.js's own "ready", which pdf.js's end of the port does not wait for.
    if (!(await ready) || !handle.alive) { kill(handle); return undefined; }
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
 * limit, or a new one. Undefined when none can start (no `child_process`, a runtime that cannot run
 * this one's executable as Node); pdf.js then runs in this process.
 */
export const acquirePdfProcess = async (pdfjs: any, workerUrl: string, memoryMb: number): Promise<PdfProcess | undefined> => {
    let handle: Handle | undefined;
    for (let i = idle.length - 1; i >= 0; i--) {
        const candidate = idle[i];
        if (candidate.alive && candidate.memoryMb === memoryMb && candidate.workerUrl === workerUrl) { idle.splice(i, 1); handle = candidate; break; }
    }
    handle ??= await start(pdfjs, workerUrl, memoryMb);
    if (!handle) return undefined;
    const current = handle;
    current.uses++;
    setRef(current, true);
    const died = new Promise<never>((_, reject) => { current.rejectDied = reject; });
    // Settled or not, never an unhandled rejection: a caller races it.
    died.catch(() => { /* raced by the caller */ });
    return {
        worker: current.worker,
        died,
        get alive() { return current.alive; },
        end: () => kill(current),
        release: () => {
            current.rejectDied = undefined;
            if (current.alive && current.uses < USES_PER_PROCESS && idle.length < MAX_IDLE) {
                setRef(current, false);
                idle.push(current);
            } else {
                kill(current);
            }
        },
    };
};
