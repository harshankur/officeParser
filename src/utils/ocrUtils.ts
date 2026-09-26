/**
 * OCR (Optical Character Recognition) Utilities
 * 
 * This module provides functions for extracting text from images using Tesseract.js.
 * Used when `config.ocr` is enabled to extract text from embedded images in documents.
 * 
 * Includes a worker pool via OcrSchedulerManager to improve performance when 
 * processing multiple images.
 * 
 * @module ocrUtils
 */

import { FullOfficeParserConfig, OcrConfig, OfficeErrorType, OfficeWarningType } from '../types.js';
import { isBrowser } from './envUtils.js';
import { buildOfficeError, getAbortError, logWarning } from './errorUtils.js';
import { median } from './numberUtils.js';

/**
 * Internal interface for tracking jobs in the scheduler queue.
 */
interface OcrJob {
    image: any;
    config: OcrConfig;
    resolve: (text: string) => void;
    reject: (err: any) => void;
    startTime: number;
    timeoutMs: number;
    isFinished?: boolean;
    /** Rejects when the job is cancelled or the pool terminated, so an abandoned recognition is not awaited forever. */
    cancelled: Promise<never>;
    /** Rejects `cancelled` with the given error. */
    cancel: (err: any) => void;
}

/**
 * Internal interface for tracking workers within the pool.
 */
interface ManagedWorker {
    worker: any;
    language: string;
    lastUsed: number;
    isBusy: boolean;
    activeJob?: OcrJob;
    /** Set once `recognize` has been called for the active job (see `runWorker`). */
    recognizing?: boolean;
    /** A re-initialization in flight, which the worker must not be terminated during (see `terminate`). */
    reinitializing?: Promise<unknown>;
}

/**
 * Test-only observation points (not part of the package's API): `afterRecognizeCall` runs right
 * after a worker's `recognize` is called, before Tesseract has sent it the job, which is the moment
 * a cancellation must not terminate the worker; `createWorker` stands in for Tesseract's, so a test
 * can supply a worker whose steps it controls (one whose re-initialization never finishes).
 */
export const ocrTestHooks: { afterRecognizeCall?: () => void; createWorker?: (language: string) => Promise<any> } = {};

/**
 * The image as bytes in memory. Tesseract loads its input before sending the job to its worker;
 * given bytes, that step does no I/O, so the job is sent within the microtasks that follow the
 * `recognize` call. Anything else is passed on for Tesseract to load.
 */
async function imageBytes(image: any): Promise<any> {
    if (image instanceof Uint8Array) return image;
    if (typeof Blob !== 'undefined' && image instanceof Blob) return new Uint8Array(await image.arrayBuffer());
    return image;
}

/**
 * Reconstructs the two-dimensional page layout of a Tesseract result from its per-word bounding
 * boxes, so recognized text reads spatially (columns line up, right-hand text stays right) instead
 * of as one flat reading-order string. Each visual line's words are placed at a character column
 * derived from their x position (using the page's median glyph width), the block is left-normalized
 * so there is no large leading indent, lines keep their top-to-bottom order, and a wide vertical gap
 * becomes a blank line. Falls back to the flat `page.text` when no word geometry is available.
 *
 * Exported for unit tests (it is not re-exported from the package entry point).
 */
/** Widest character column the layout reconstruction will pad to (a very wide page is ~200). */
const MAX_OCR_COLUMNS = 1000;

export function layoutOcrText(page: any): string {
    const blocks = page?.blocks;
    const flat: string = typeof page?.text === 'string' ? page.text : '';
    if (!Array.isArray(blocks) || !blocks.length) return flat;

    interface OcrWord { text: string; x0: number; x1: number; }
    interface OcrLine { words: OcrWord[]; y0: number; }
    const lines: OcrLine[] = [];
    for (const block of blocks) {
        for (const para of block?.paragraphs || []) {
            for (const line of para?.lines || []) {
                const words: OcrWord[] = [];
                for (const w of line?.words || []) {
                    const text = (w?.text || '').trim();
                    if (text) words.push({ text, x0: w.bbox?.x0 ?? 0, x1: w.bbox?.x1 ?? 0 });
                }
                if (words.length) lines.push({ words, y0: line.bbox?.y0 ?? line?.words?.[0]?.bbox?.y0 ?? 0 });
            }
        }
    }
    if (!lines.length) return flat;

    // Character width unit: median per-glyph width across all words.
    const glyphWidths: number[] = [];
    for (const l of lines) for (const w of l.words) if (w.text.length) glyphWidths.push((w.x1 - w.x0) / w.text.length);
    const charWidth = median(glyphWidths.filter(v => v > 0)) || 8;

    // Left-normalize so the leftmost word sits at column 0 (no giant leading indent).
    let minX = Infinity;
    for (const l of lines) for (const w of l.words) minX = Math.min(minX, w.x0);
    if (!Number.isFinite(minX)) minX = 0;

    lines.sort((a, b) => a.y0 - b.y0);
    const pitches: number[] = [];
    for (let i = 1; i < lines.length; i++) pitches.push(lines[i].y0 - lines[i - 1].y0);
    // Line pitch is the typical row advance: exclude near-zero gaps (two lines Tesseract split across
    // columns at the same y), which would otherwise drag the median down on a multi-column scan. If an
    // unusually wide face or very tight leading leaves the filtered set empty, fall back to the
    // unfiltered positive median before the last-resort 2x charWidth, so rowBand does not overshoot the
    // real pitch and merge consecutive single-column lines into one row.
    const linePitch = median(pitches.filter(v => v > charWidth)) || median(pitches.filter(v => v > 0)) || charWidth * 2;

    // Merge lines that sit at (nearly) the same y into one visual row. A multi-column or table scan
    // puts each column's line in its own Tesseract block at the same vertical position; placing each on
    // its own output line staircases the columns down the page. The band is a tight fraction of the row
    // pitch, so single-column lines (a full pitch apart) are never merged. Words are then placed by x.
    interface OcrRow { words: OcrWord[]; y0: number; }
    const rowBand = 0.4 * linePitch;
    const rows: OcrRow[] = [];
    for (const l of lines) {
        const last = rows[rows.length - 1];
        if (last && l.y0 - last.y0 <= rowBand) last.words.push(...l.words);
        else rows.push({ words: [...l.words], y0: l.y0 });
    }

    const out: string[] = [];
    let prevY: number | null = null;
    for (const r of rows) {
        if (prevY !== null && r.y0 - prevY > 1.8 * linePitch) out.push(''); // blank line for a big vertical gap
        prevY = r.y0;
        let s = '';
        let col = 0;
        // Left-to-right across every column merged into this row.
        for (const w of [...r.words].sort((a, b) => a.x0 - b.x0)) {
            // The column comes from Tesseract's own box coordinates, so it is clamped: a bogus x0
            // (or a degenerate glyph width) must not turn into a line of hundreds of thousands of
            // spaces, which would be a memory spike driven straight by the input image.
            const target = Math.min(MAX_OCR_COLUMNS, Math.max(col, Math.round((w.x0 - minX) / charWidth)));
            if (target > col) { s += ' '.repeat(target - col); col = target; }
            else if (s.length && !s.endsWith(' ')) { s += ' '; col += 1; } // always keep words apart
            s += w.text;
            col += w.text.length;
        }
        out.push(s.trimEnd());
    }
    return out.join('\n');
}

/**
 * Wraps a promise in a timeout.
 *
 * @param promise - The promise to wrap
 * @param ms - Timeout duration in milliseconds
 * @param errMsg - Error message to throw if timeout occurs
 * @returns The wrapped promise
 */
function withTimeout<T>(promise: Promise<T>, ms: number, errMsg: string): Promise<T> {
    let id: any;
    const timeout = new Promise<never>((_, reject) => {
        id = setTimeout(() => {
            reject(new Error(errMsg));
        }, ms);
    });
    return Promise.race([promise, timeout]).then(
        (res) => { clearTimeout(id); return res; },
        (err) => { clearTimeout(id); throw err; }
    );
}

/**
 * Manages a pool of Tesseract workers with "Smart Affinity".
 * 
 * Instead of a simple scheduler, this manager allows workers to persist with 
 * a specific language affinity. If a new language is requested and the pool 
 * is at capacity, it re-initializes the Least Recently Used (LRU) idle worker 
 * rather than resetting the entire pool.
 * 
 * Implements lazy loading of tesseract.js to ensure no background processes
 * are spawned unless OCR is explicitly used.
 */
class OcrSchedulerManager {
    private static instance: OcrSchedulerManager;
    private pool: ManagedWorker[] = [];
    private queue: OcrJob[] = [];
    private readonly MAX_WORKERS: number = 4;
    private idleTimeout: number = 10000; // 10s default
    private timeoutId: NodeJS.Timeout | null = null;
    /** Jobs not yet resolved or rejected: queued, waiting for a worker, or being recognized. */
    private readonly unfinished = new Set<OcrJob>();
    private isProcessing: boolean = false;

    private constructor() { }

    /**
     * Returns the singleton instance of the manager.
     */
    public static getInstance(): OcrSchedulerManager {
        if (!OcrSchedulerManager.instance) {
            OcrSchedulerManager.instance = new OcrSchedulerManager();
        }
        return OcrSchedulerManager.instance;
    }

    /**
     * Checks if the singleton instance has been initialized.
     */
    public static hasInstance(): boolean {
        return !!OcrSchedulerManager.instance;
    }

    /**
     * Resets the inactivity timer. If the timer reaches its duration, 
     * all workers are terminated automatically.
     */
    private resetIdleTimer(): void {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
        }

        if (this.idleTimeout > 0) {
            this.timeoutId = setTimeout(() => {
                this.timeoutId = null;
                // A recognition (or a worker's first start) can outlast the idle period: the pool is
                // idle only when no job is waiting or running, and until then the timer starts over.
                if (this.unfinished.size > 0) this.resetIdleTimer();
                else void this.terminate();
            }, this.idleTimeout);
        }
    }

    /**
     * Performs OCR on an image using the smart worker pool.
     * 
     * @param image - Image data (Buffer, string path, or Blob)
     * @param config - OCR configuration (language, custom paths, timeouts, signal)
     * @returns Recognized text
     */
    public async recognize(image: any, config?: OcrConfig): Promise<string> {
        const signal = config?.abortSignal;
        if (signal?.aborted) {
            return Promise.reject(getAbortError());
        }

        return new Promise((resolve, reject) => {
            // Update idle timeout if provided.
            const effectiveAutoTerminate = config?.timeout?.autoTerminate;
            if (effectiveAutoTerminate !== undefined) {
                this.idleTimeout = effectiveAutoTerminate;
            }

            // Reset the inactivity timer every time a new job is requested
            this.resetIdleTimer();

            let abortListener: (() => void) | null = null;
            let finished = false;
            let job: OcrJob;
            let cancel: (err: any) => void = () => { };
            const cancelled = new Promise<never>((_, rejectCancelled) => { cancel = rejectCancelled; });
            cancelled.catch(() => { });

            const cleanResolve = (val: string) => {
                if (finished) return;
                finished = true;
                if (job) { job.isFinished = true; this.unfinished.delete(job); }
                if (abortListener && signal) {
                    signal.removeEventListener('abort', abortListener);
                }
                resolve(val);
            };

            const cleanReject = (err: any) => {
                if (finished) return;
                finished = true;
                if (job) { job.isFinished = true; this.unfinished.delete(job); }
                if (abortListener && signal) {
                    signal.removeEventListener('abort', abortListener);
                }
                reject(err);
            };

            // Priority: timeout.recognition (new) > 30 s default.
            const recogTimeout = config?.timeout?.recognition ?? 30000;

            // Create job
            job = {
                image,
                config: config || {},
                resolve: cleanResolve,
                reject: cleanReject,
                startTime: Date.now(),
                timeoutMs: recogTimeout,
                cancelled,
                cancel
            };
            this.unfinished.add(job);

            if (signal) {
                abortListener = () => {
                    if (finished) return;
                    
                    const err = getAbortError();
                    cleanReject(err);
                    cancel(err);

                    // 1. Remove job from queue if it hasn't run yet
                    const idx = this.queue.indexOf(job);
                    if (idx !== -1) {
                        this.queue.splice(idx, 1);
                    }

                    // 2. A worker recognizing this job is removed from the pool and terminated, to stop
                    // the work. Not at once: Tesseract sends the job to its worker in the microtasks
                    // after `recognize` is called, and a worker terminated before that send makes
                    // Tesseract's send() reject with nothing to catch it, which ends a Node process. So
                    // the worker is terminated on the next macrotask, when the job has been sent. A
                    // worker still starting or re-initializing for this job is left to the code
                    // awaiting that step, which finds the job finished and cleans up.
                    const workerIndex = this.pool.findIndex(mw => mw.activeJob === job && mw.recognizing);
                    if (workerIndex !== -1) {
                        const [managedWorker] = this.pool.splice(workerIndex, 1);
                        setTimeout(() => {
                            Promise.resolve().then(() => managedWorker.worker.terminate()).catch(() => { });
                        }, 0);
                        // Trigger queue processing for subsequent tasks
                        this.processQueue();
                    }
                };
                signal.addEventListener('abort', abortListener);
            }

            // Add job to queue and trigger processing
            this.queue.push(job);
            this.processQueue();
        });
    }

    /**
     * Attempts to process the next job in the queue using an available worker.
     * Designed to be race-free and support concurrent/parallel job execution.
     */
    private async processQueue(): Promise<void> {
        if (this.isProcessing) return;
        this.isProcessing = true;

        try {
            while (this.queue.length > 0) {
                const nextJob = this.queue[0];
                if (nextJob.isFinished) {
                    this.queue.shift();
                    continue;
                }
                const requestedLanguage = nextJob.config.language || 'eng';

                // 1. Find an idle worker with the EXACT language affinity
                let managed = this.pool.find(mw => !mw.isBusy && mw.language === requestedLanguage);

                // 2. If not found and we have room, create a new worker
                if (!managed && this.pool.length < this.MAX_WORKERS) {
                    const job = this.queue.shift();
                    if (!job) continue;

                    this.createAndRunWorker(job, requestedLanguage);
                    continue;
                }

                // 3. If still not found and we are at capacity, find the LRU idle worker and re-initialize it
                if (!managed) {
                    const idleWorkers = this.pool.filter(mw => !mw.isBusy);
                    if (idleWorkers.length > 0) {
                        const job = this.queue.shift();
                        if (!job) continue;

                        managed = idleWorkers.reduce((prev, curr) => (prev.lastUsed < curr.lastUsed ? prev : curr));
                        this.reinitializeAndRunWorker(managed, job, requestedLanguage);
                        continue;
                    }
                }

                // 4. If we have a worker ready, execute the job
                if (managed) {
                    const job = this.queue.shift();
                    if (!job) continue;

                    this.runWorker(managed, job);
                    continue;
                }

                // No workers can be allocated right now (all busy and pool at capacity). Break work loop.
                break;
            }
        } finally {
            this.isProcessing = false;
        }
    }

    /**
     * Helper to dynamically instantiate a Tesseract worker, register it to the pool, and run the job.
     */
    private async createAndRunWorker(job: OcrJob, requestedLanguage: string): Promise<void> {
        // Priority: timeout.workerLoad (new) > 60 s default.
        const loadTimeout = job.config.timeout?.workerLoad ?? 60000;
        let managed: ManagedWorker | null = null;

        try {
            const { createWorker } = await import('tesseract.js');
            const options: any = { logger: () => { } };
            if (job.config.workerPath) options.workerPath = job.config.workerPath;
            if (job.config.corePath) options.corePath = job.config.corePath;
            if (job.config.langPath) options.langPath = job.config.langPath;

            const workerPromise = ocrTestHooks.createWorker ? ocrTestHooks.createWorker(requestedLanguage) : createWorker(requestedLanguage, 1, options);

            // To prevent dangling worker threads on timeout or abort, we register a post-resolution hook
            // that terminates the worker if the promise finishes after the timeout has fired or the job is finished.
            let hasTimedOutOrAborted = false;
            workerPromise.then(
                async (worker) => {
                    if (hasTimedOutOrAborted || job.isFinished) {
                        try {
                            await worker.terminate();
                        } catch (e) {}
                    }
                },
                () => {}
            );

            const worker = loadTimeout > 0
                ? await withTimeout(workerPromise, loadTimeout, `OCR worker initialization timed out after ${loadTimeout}ms`).catch(err => {
                    hasTimedOutOrAborted = true;
                    throw err;
                })
                : await workerPromise;

            // If the job finished/aborted while loading, clean up the worker and skip execution.
            if (job.isFinished) {
                hasTimedOutOrAborted = true;
                try {
                    await worker.terminate();
                } catch (e) {}
                this.processQueue();
                return;
            }

            managed = {
                worker,
                language: requestedLanguage,
                lastUsed: Date.now(),
                isBusy: false
            };
            this.pool.push(managed);

            await this.runWorker(managed, job);
        } catch (err) {
            job.reject(err);
            this.processQueue();
        }
    }

    /**
     * Helper to reinitialize an existing idle worker with a different language affinity and run the job.
     */
    private async reinitializeAndRunWorker(managed: ManagedWorker, job: OcrJob, requestedLanguage: string): Promise<void> {
        // Priority: timeout.workerLoad (new) > 60 s default.
        const loadTimeout = job.config.timeout?.workerLoad ?? 60000;
        managed.isBusy = true;
        managed.lastUsed = Date.now();
        managed.activeJob = job;

        try {
            const reinitPromise = managed.worker.reinitialize(requestedLanguage);
            const reinitialized = loadTimeout > 0
                ? withTimeout(reinitPromise, loadTimeout, `OCR worker re-initialization timed out after ${loadTimeout}ms`)
                : reinitPromise;
            // What terminate() waits for before ending the worker: bounded by the load timeout, so a
            // language download that never finishes cannot hold terminateOcr() forever.
            managed.reinitializing = reinitialized.then(() => { }, () => { });
            await reinitialized;
            managed.reinitializing = undefined;
            managed.language = requestedLanguage;

            // If the job finished/aborted while reinitializing, clean up and skip execution.
            if (job.isFinished) {
                const index = this.pool.indexOf(managed);
                if (index !== -1) {
                    this.pool.splice(index, 1);
                }
                try {
                    await managed.worker.terminate();
                } catch (e) {}
                this.processQueue();
                return;
            }

            await this.runWorker(managed, job);
        } catch (err: any) {
            // Re-initialization failed/timed out, remove worker from pool and terminate
            const index = this.pool.indexOf(managed);
            if (index !== -1) {
                this.pool.splice(index, 1);
            }
            try {
                await managed.worker.terminate();
            } catch (e) {}

            job.reject(err);
            this.processQueue();
        }
    }

    /**
     * Helper to execute OCR text recognition on the worker and return the results.
     */
    private async runWorker(managed: ManagedWorker, job: OcrJob): Promise<void> {
        // Priority: timeout.recognition (new) > 30 s default.
        const recogTimeout = job.config.timeout?.recognition ?? 30000;
        managed.isBusy = true;
        managed.lastUsed = Date.now();
        managed.activeJob = job;

        try {
            const image = await imageBytes(job.image);
            // Cancelled while the image loaded: nothing was sent to the worker, which stays in the pool.
            if (job.isFinished) return;
            // When preserving layout (default), ask Tesseract for the block/word tree so we can
            // rebuild the 2-D page layout from per-word boxes; otherwise the flat text is enough.
            const wantLayout = job.config.preserveLayout !== false;
            const recognizePromise = wantLayout
                ? managed.worker.recognize(image, {}, { text: true, blocks: true })
                : managed.worker.recognize(image);
            managed.recognizing = true;
            ocrTestHooks.afterRecognizeCall?.();
            // A cancelled job's worker is terminated, so its recognition never settles: stop waiting.
            // Raced inside the timeout, so cancelling also clears the timeout's timer, which would
            // otherwise keep the process alive for the rest of it.
            const recognition = Promise.race([recognizePromise, job.cancelled]);
            const { data } = recogTimeout > 0
                ? await withTimeout(recognition, recogTimeout, `OCR recognition timed out after ${recogTimeout}ms`)
                : await recognition;

            job.resolve(wantLayout ? layoutOcrText(data) : (data.text || ''));
        } catch (err: any) {
            // If it timed out, terminate and remove worker to avoid reusing a stuck process
            if (err.message?.includes('timed out')) {
                const index = this.pool.indexOf(managed);
                if (index !== -1) {
                    this.pool.splice(index, 1);
                }
                try {
                    await managed.worker.terminate();
                } catch (e) {}
            }
            job.reject(err);
        } finally {
            managed.recognizing = false;
            if (this.pool.includes(managed)) {
                managed.isBusy = false;
                managed.activeJob = undefined;
                managed.lastUsed = Date.now();
            }
            this.processQueue();
        }
    }

    /**
     * Terminates all workers in the pool and resets the state. Jobs still waiting or running fail at
     * once (their parses report OCR_FAILED) rather than hang on workers that are going away. A worker
     * is never terminated while Tesseract may still be sending it a job, which would end a Node
     * process (see the abort listener in `recognize`): a recognizing worker is terminated on the next
     * macrotask, and a re-initializing one once that settles.
     */
    public async terminate(): Promise<void> {
        if (this.timeoutId) {
            clearTimeout(this.timeoutId);
            this.timeoutId = null;
        }

        // Not reported here: each job's parse reports its own as OCR_FAILED, to its own config.
        const err = buildOfficeError(OfficeErrorType.OCR_TERMINATED);
        this.queue = [];
        for (const job of [...this.unfinished]) {
            job.cancel(err);
            job.reject(err);
        }

        const workers = this.pool;
        this.pool = [];
        await Promise.all(workers.map(async managed => {
            if (managed.recognizing) await new Promise(resolve => setTimeout(resolve, 0));
            if (managed.reinitializing) await managed.reinitializing;
            try { await managed.worker.terminate(); } catch { /* already gone */ }
        }));
    }
}

/**
 * Reads an image buffer's media type from its magic bytes. The browser path wraps the bytes in a
 * Blob, and that Blob's type is what the decoder trusts, so it has to match the actual bytes: PDF
 * page images arrive as PNG, embedded pictures can be anything the document carried.
 *
 * Exported for unit tests (it is not re-exported from the package entry point).
 *
 * @param buf - The image bytes
 * @returns The detected media type, defaulting to `image/png`
 */
export function sniffImageMime(buf: Buffer): string {
    if (buf.length >= 4) {
        if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'image/png';
        if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'image/jpeg';
        if (buf[0] === 0x42 && buf[1] === 0x4d) return 'image/bmp';
        if (buf[0] === 0x47 && buf[1] === 0x49 && buf[2] === 0x46) return 'image/gif';
        if (buf.length >= 12 && buf.toString('ascii', 0, 4) === 'RIFF' && buf.toString('ascii', 8, 12) === 'WEBP') return 'image/webp';
        if ((buf[0] === 0x49 && buf[1] === 0x49 && buf[2] === 0x2a) || (buf[0] === 0x4d && buf[1] === 0x4d && buf[2] === 0x00)) return 'image/tiff';
    }
    return 'image/png';
}

/**
 * Performs Optical Character Recognition (OCR) on an image to extract text.
 * 
 * Uses Tesseract.js to recognize text in the provided image buffer.
 * This is useful for extracting text from screenshots, scanned documents,
 * charts with labels, or any image containing text.
 * 
 * This function uses a shared worker pool to minimize initialization overhead.
 * 
 * @param image - The image data as a Buffer, file path, or Blob
 * @param config - Optional configuration for language and custom worker paths
 * @param mimeType - Media type of `image` when it is a Buffer; sniffed from its signature otherwise
 * @returns A promise that resolves to the recognized text as a string
 * @throws {Error} If the image cannot be processed or Tesseract initialization fails
 * 
 * @example
 * ```typescript
 * // Extract text from an English image
 * const text = await performOcr(imageBuffer, { language: 'eng' });
 * ```
 * 
 * @see https://github.com/naptha/tesseract.js for supported languages and options
 */
export const performOcr = async (image: Buffer | string, config?: OcrConfig, mimeType?: string): Promise<string> => {
    // Prepare image data
    let inputImage: any = image;

    // In browser environment, convert Buffer to Blob for better compatibility
    // @ts-ignore
    if (isBrowser && typeof Blob !== 'undefined' && Buffer.isBuffer(image)) {
        inputImage = new Blob([image as any], { type: mimeType || sniffImageMime(image) });
    }

    return await OcrSchedulerManager.getInstance().recognize(inputImage, config);
};

/** Whether a parse's caller has cancelled it, through `abortSignal` or the OCR-only `ocrConfig.abortSignal`. */
const parseCancelled = (config: FullOfficeParserConfig): boolean =>
    !!(config.abortSignal?.aborted || config.ocrConfig?.abortSignal?.aborted);

/**
 * OCR of one image while parsing: the recognized text, or `undefined` when recognition failed (an
 * `OCR_FAILED` warning, and the parse goes on without that text).
 *
 * A cancelled parse is not a failed recognition. Once the caller's signal has fired (checked before
 * recognizing, after it, and whenever recognition rejects), this throws the AbortError, so the
 * parse rejects instead of resolving without the text. An OCR timeout stays a failed recognition.
 *
 * @param image - The image bytes (or a path)
 * @param config - The parse's resolved configuration (its `ocrConfig` and signals)
 * @param name - The attachment's name, for the warning
 * @param mimeType - Media type of `image`; sniffed from its signature otherwise
 */
export const ocrDuringParse = async (image: Buffer | string, config: FullOfficeParserConfig, name: string, mimeType?: string): Promise<string | undefined> => {
    if (parseCancelled(config)) throw getAbortError();
    try {
        const text = (await performOcr(image, { ...config.ocrConfig }, mimeType)).trim();
        if (parseCancelled(config)) throw getAbortError();
        return text;
    } catch (error: any) {
        if (error?.name === 'AbortError') throw error;
        if (parseCancelled(config)) throw getAbortError();
        logWarning(OfficeWarningType.OCR_FAILED, config, name, error);
        return undefined;
    }
};

/**
 * Terminates all OCR workers and cleans up resources.
 * 
 * Should be called when the application is shutting down or OCR is no longer needed
 * to prevent memory leaks and dangling worker processes.
 */
export const terminateOcr = async (): Promise<void> => {
    if (OcrSchedulerManager.hasInstance()) {
        await OcrSchedulerManager.getInstance().terminate();
    }
};
