/**
 * The files pdf.js loads while it reads a PDF, and what pdf.js says about content it leaves out.
 *
 * pdf.js reads the text of a font the PDF does not embed through a CMap it does not hold when the font
 * names one of the predefined CJK encodings (UniJIS-UTF16-H, 90ms-RKSJ-H...), and decodes JBIG2, CCITT
 * fax and JPEG 2000 pictures with WebAssembly decoders it loads when a picture needs one. Given no
 * folder for them, every glyph in such a font was lost (and the document called scanned) and every
 * such picture decoded to nothing. They are read from pdfjs-dist's own `cmaps/` and `wasm/` folders in
 * Node, and from `pdfCMapUrl` and `pdfDecoderUrl` (by default, the CDN copy of the same pdf.js) in the
 * browser, through a loader of this module's (pdf.js's `BinaryDataFactory`) that asks for nothing else:
 * only a CMap or decoder file by name, inside its folder, bounded in size and time, and not at all from
 * the built-in CDN address when `allowRemoteAssets` is false. A file that cannot be had is reported once
 * per document (PDF_ASSET_UNAVAILABLE), with the page that first needed it.
 *
 * pdf.js also leaves some pictures out without drawing anything in their place (one past its
 * `maxImageSize`, a small inline picture it cannot decode, one without dimensions, an XObject it
 * cannot read), and says so only in a console warning from its worker. Its worker runs at warning
 * verbosity, and those warnings are taken where they can be (the separate process's messages, this
 * process's console, the messages of a browser worker this module starts) and reported with the page
 * being read.
 *
 * @module pdfAssets
 */

import { DEFAULT_PDF_CMAP_URL, DEFAULT_PDF_DECODER_URL } from '../../defaults.js';
import { FullOfficeParserConfig, OfficeWarningType } from '../../types.js';
import { isBrowser } from '../../utils/envUtils.js';
import { getAbortError, logWarning } from '../../utils/errorUtils.js';

/** The largest CMap or decoder file read (the largest pdf.js ships is about 250 KB). */
export const MAX_ASSET_BYTES = 4 * 1024 * 1024;
/**
 * How long a document may wait on the files of one folder, in all: reading one file may take no longer,
 * and once the files read from a folder have kept the document waiting this long (counting the time any
 * of them was being read), nothing more is asked of that folder.
 */
export const ASSET_TIMEOUT_MS = 15_000;
/** How long a browser worker of this module's may take to start pdf.js before pdf.js starts its own. */
const WORKER_START_MS = 15_000;
/** The pdf.js warnings of each kind (see PDFJS_WARNINGS) one document's worker may send, and how much of each is kept. */
export const MAX_WORKER_WARNINGS = 1000;
export const MAX_WARNING_LENGTH = 300;
/** A packed CMap's file name, as pdf.js names one it has (`UniJIS-UTF16-H.bcmap`). */
const CMAP_FILE = /^[A-Za-z0-9-]{1,40}\.bcmap$/;
/** The decoders pdf.js loads for pictures: JBIG2 and CCITT fax, and JPEG 2000. */
const DECODER_FILES = new Set(['jbig2.wasm', 'openjpeg.wasm']);

/**
 * The warnings pdf.js's worker gives for content it leaves out, as it words them (pdf.js 6). A test
 * checks that each is still in the installed worker, so an update of pdf.js that rewords one is seen.
 */
export const PDFJS_WARNINGS = {
    imageTooLarge: 'Image exceeded maximum allowed size and was removed.',
    inlineImageUndecodable: 'Unable to decode inline image: ',
    imageWithoutSize: 'Image dimensions are missing, or not numbers.',
    xObjectNotDrawn: 'getOperatorList - ignoring XObject: ',
    xObjectTextNotRead: 'getTextContent - ignoring XObject: ',
    imageUndecodable: 'Unable to decode image ',
} as const;

/**
 * Which of PDFJS_WARNINGS a line from pdf.js's worker ("Warning: ...") is, or undefined for any other
 * line (pdf.js's own notes, which the parser does not act on). The same test, as source, runs where
 * the lines are written (see PDFJS_WARNING_KIND_SOURCE).
 */
export function pdfjsWarningKind(line: string): keyof typeof PDFJS_WARNINGS | undefined {
    if (!line.startsWith('Warning: ')) return undefined;
    for (const [kind, words] of Object.entries(PDFJS_WARNINGS) as [keyof typeof PDFJS_WARNINGS, string][]) {
        if (line.startsWith(words, 'Warning: '.length)) return kind;
    }
    return undefined;
}

/**
 * The test of pdfjsWarningKind as source, for the separate process and the browser Worker: a function
 * of a line giving its kind or ''. Each sends only the lines it names, at most MAX_WORKER_WARNINGS of
 * each kind a document, so pdf.js's notes (one for every picture pdf.js is told to leave out when
 * pictures are not extracted) cannot use up what may be sent before a line the parser reports.
 */
export const PDFJS_WARNING_KIND_SOURCE = `((kinds) => (line) => {
    if (typeof line !== 'string' || !line.startsWith('Warning: ')) return '';
    for (const [kind, words] of kinds) if (line.startsWith(words, 9)) return kind;
    return '';
})(${JSON.stringify(Object.entries(PDFJS_WARNINGS))})`;

export type PdfAssetKind = 'cMapUrl' | 'wasmUrl';
/** Why a file could not be had: it is not there (or no folder is set), reading it failed, or `allowRemoteAssets` forbids its address. */
export type PdfAssetReason = 'missing' | 'failed' | 'blocked';

/** Where one kind of file is read from. */
export interface PdfAssetFolder {
    /** An absolute address ending in '/' (`file:` in Node, `http(s):` in the browser); '' for none. */
    url: string;
    /** The option that sets it, which a warning names. */
    option: 'pdfCMapUrl' | 'pdfDecoderUrl';
    /** The folder is the built-in CDN address and `allowRemoteAssets` is false: nothing is read from it. */
    blocked: boolean;
}

export interface PdfAssetFolders {
    cMapUrl: PdfAssetFolder;
    wasmUrl: PdfAssetFolder;
}

/** What a document's read tells the files' loader: the page being read, and whether the read has stopped. */
export interface PdfReadPosition {
    readonly pageBeingRead: number | undefined;
    /** Nothing more is asked of pdf.js (a budget ran out, or the parse ended). */
    readonly spent: boolean;
    /** The parse ended while the document was read: nothing more is reported. */
    readonly cancelled: boolean;
}

/** What pdf.js is told when the loader will not read a file it asks for: it is not one of pdf.js's CMaps or decoders. */
class AssetRefused extends Error { }

/** A file that could not be had, and why. */
class AssetUnavailable extends Error {
    constructor(readonly reason: Exclude<PdfAssetReason, 'blocked'>, message: string) {
        super(message);
    }
}

/** pdfjs-dist's own `cmaps/` and `wasm/` folders, as `file:` addresses; undefined where they cannot be found. */
async function installedFolders(): Promise<Record<PdfAssetKind, string> | undefined> {
    try {
        // @ts-ignore - 'require' is available in Node.js/CommonJS environment
        const workerPath: string = require.resolve('pdfjs-dist/legacy/build/pdf.worker.mjs');
        const path = await import('path');
        const { pathToFileURL } = await import('url');
        const root = path.join(path.dirname(workerPath), '..', '..');
        return {
            cMapUrl: `${pathToFileURL(path.join(root, 'cmaps')).href}/`,
            wasmUrl: `${pathToFileURL(path.join(root, 'wasm')).href}/`,
        };
    } catch {
        return undefined;
    }
}

/**
 * An address as an absolute one ending in '/': a path or `file:` address in Node, an address relative
 * to the page in the browser. '' stays ''.
 */
async function folderAddress(value: string): Promise<string> {
    if (!value) return '';
    const slashed = value.endsWith('/') ? value : `${value}/`;
    if (/^[a-z][a-z0-9+.-]*:\/\//i.test(slashed) && !/^[a-z]:[\\/]/i.test(value)) return new URL(slashed).href;
    if (isBrowser) {
        const base = typeof document !== 'undefined' ? document.baseURI : (globalThis as any).location?.href;
        return new URL(slashed, base).href;
    }
    const path = await import('path');
    const { pathToFileURL } = await import('url');
    return `${pathToFileURL(path.resolve(value)).href}/`;
}

/**
 * Where a parse reads CMaps and decoders from. In Node, pdfjs-dist's own folders unless `pdfCMapUrl`
 * or `pdfDecoderUrl` is set (to something other than its default or ''); where they cannot be found,
 * or in the browser, the option's value. The built-in CDN address is blocked when `allowRemoteAssets`
 * is false. A value that is not text is reported (INVALID_CONFIG_VALUE) and the default used.
 */
export async function resolvePdfAssetFolders(config: FullOfficeParserConfig): Promise<PdfAssetFolders> {
    const local = isBrowser ? undefined : await installedFolders();
    const folder = async (kind: PdfAssetKind, option: PdfAssetFolder['option'], fallback: string): Promise<PdfAssetFolder> => {
        let value: unknown = config[option];
        if (value === undefined || value === null) value = fallback;
        if (typeof value !== 'string') {
            logWarning(OfficeWarningType.INVALID_CONFIG_VALUE, config, { option, value, expected: "a folder's address ending in '/'", fallback });
            value = fallback;
        }
        const given = value as string;
        const own = given !== fallback && given !== '';
        if (!own && local) return { url: local[kind], option, blocked: false };
        const blocked = config.allowRemoteAssets === false && given !== '' && given === fallback;
        let url = '';
        try { url = await folderAddress(given); } catch { url = ''; }
        return { url, option, blocked };
    };
    return {
        cMapUrl: await folder('cMapUrl', 'pdfCMapUrl', DEFAULT_PDF_CMAP_URL),
        wasmUrl: await folder('wasmUrl', 'pdfDecoderUrl', DEFAULT_PDF_DECODER_URL),
    };
}

/** `promise`, or a rejection once `signal` aborts (with its reason). */
function untilAborted<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
    if (signal.aborted) return Promise.reject(signal.reason);
    return new Promise<T>((resolve, reject) => {
        const onAbort = () => reject(signal.reason);
        signal.addEventListener('abort', onAbort, { once: true });
        promise.then(
            value => { signal.removeEventListener('abort', onAbort); resolve(value); },
            error => { signal.removeEventListener('abort', onAbort); reject(error); },
        );
    });
}

/** Reads `name` from the `file:` folder `base`, which it must stay inside of. */
async function readLocalFile(base: string, name: string, signal: AbortSignal): Promise<Uint8Array> {
    const path = await import('path');
    const fs = await import('fs');
    const { fileURLToPath } = await import('url');
    const folder = path.resolve(fileURLToPath(base));
    const file = path.resolve(folder, name);
    if (path.dirname(file) !== folder) throw new AssetUnavailable('failed', `${name} is not inside ${base}`);
    let handle: import('fs/promises').FileHandle;
    try {
        handle = await fs.promises.open(file, 'r');
    } catch (e: any) {
        if (e?.code === 'ENOENT' || e?.code === 'ENOTDIR') throw new AssetUnavailable('missing', `${name} is not in ${base}`);
        throw e;
    }
    try {
        const stat = await handle.stat();
        if (!stat.isFile()) throw new AssetUnavailable('missing', `${name} in ${base} is not a file`);
        if (stat.size > MAX_ASSET_BYTES) throw new AssetUnavailable('failed', `${name} is larger than ${MAX_ASSET_BYTES} bytes`);
        const bytes = await handle.readFile({ signal });
        if (bytes.length > MAX_ASSET_BYTES) throw new AssetUnavailable('failed', `${name} is larger than ${MAX_ASSET_BYTES} bytes`);
        return new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    } finally {
        await handle.close().catch(() => { /* already closed */ });
    }
}

/** Fetches `url`, at most MAX_ASSET_BYTES of it. */
async function fetchRemoteFile(url: URL, signal: AbortSignal): Promise<Uint8Array> {
    if (typeof fetch !== 'function') throw new AssetUnavailable('failed', 'this runtime has no fetch()');
    const response = await fetch(url.href, { signal });
    if (response.status === 404 || response.status === 410) throw new AssetUnavailable('missing', `${url.href} answered ${response.status}`);
    if (!response.ok) throw new AssetUnavailable('failed', `${url.href} answered ${response.status}`);
    const tooLarge = () => new AssetUnavailable('failed', `${url.href} is larger than ${MAX_ASSET_BYTES} bytes`);
    const declared = Number(response.headers.get('content-length'));
    if (declared > MAX_ASSET_BYTES) { response.body?.cancel().catch(() => { /* not read */ }); throw tooLarge(); }
    const reader = response.body?.getReader();
    if (!reader) {
        const whole = new Uint8Array(await response.arrayBuffer());
        if (whole.length > MAX_ASSET_BYTES) throw tooLarge();
        return whole;
    }
    const chunks: Uint8Array[] = [];
    let total = 0;
    while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        total += value.length;
        if (total > MAX_ASSET_BYTES) { reader.cancel().catch(() => { /* not read */ }); throw tooLarge(); }
        chunks.push(value);
    }
    const bytes = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length; }
    return bytes;
}

/** A file that could not be had, as reported: the first page that needed it, and why. */
interface Unavailable {
    kind: PdfAssetKind;
    reason: PdfAssetReason;
}

/**
 * How long the files of one folder have kept a document waiting: the time any of them was being read,
 * so files read side by side count their time once.
 */
class FolderWait {
    private reading = 0;
    private since = 0;
    private before = 0;
    waited(): number {
        return this.before + (this.reading ? Date.now() - this.since : 0);
    }
    start(): void {
        if (!this.reading++) this.since = Date.now();
    }
    end(): void {
        if (!--this.reading) this.before += Date.now() - this.since;
    }
}

/**
 * One parse's CMap and decoder files, and what its pdf.js worker says about content it leaves out
 * (see the module's comment). It reports through the parse's config, naming the page being read
 * (see `follow`), and nothing once the read has ended or the parse has (`close`).
 */
export class PdfAssetSession {
    /** Aborted when the parse ends, ending every file still being read. */
    private readonly ended = new AbortController();
    private position: PdfReadPosition | undefined;
    private closed = false;
    private readonly unavailable = new Map<string, Unavailable>();
    /** The files being read, by name: pdf.js asking for one again before it has arrived is given the same read. */
    private readonly reads = new Map<string, Promise<Uint8Array>>();
    /** How long each folder has kept the document waiting (see ASSET_TIMEOUT_MS). */
    private readonly waits: Record<PdfAssetKind, FolderWait> = { cMapUrl: new FolderWait(), wasmUrl: new FolderWait() };
    /** Why pdf.js could not decode a picture, by the id it gave it (its warning comes before the empty picture). */
    private readonly undecodable = new Map<string, string>();
    /** The worker's lines acted on, by kind (see workerWarning). */
    private readonly warningsTaken = new Map<string, number>();

    constructor(private readonly folders: PdfAssetFolders, private readonly config: FullOfficeParserConfig, private readonly maxImagePixels: number) { }

    /** Names the read whose page warnings are about: each document read of the parse in turn. */
    follow(position: PdfReadPosition): void {
        this.position = position;
    }

    /** Ends the files still being read and every report: the parse has ended. */
    close(): void {
        this.closed = true;
        this.ended.abort(getAbortError());
    }

    /** Whether a warning may be reported now: a document is being read and the parse has not ended. */
    private get reporting(): boolean {
        return !this.closed && !!this.position && !this.position.cancelled;
    }

    /**
     * What `getDocument` is given so pdf.js reads CMaps and decoders through this session: every file
     * is asked of the main thread (`useWorkerFetch: false`), where the loader below reads it. pdf.js's
     * worker imports a decoder's JavaScript fallback (`wasm/*_nowasm_fallback.js`, for a runtime that
     * cannot run WebAssembly) from `wasmUrl` itself, so that is given only for a folder that may be read.
     * `standardFontDataUrl` and `iccUrl` are not given: the standard fonts' data changes nothing the
     * parser extracts, and pdf.js's ICC colour needs a synchronous request Node does not have.
     */
    documentOptions(): Record<string, unknown> {
        const usable = (folder: PdfAssetFolder) => (folder.url && !folder.blocked ? folder.url : undefined);
        const session = this;
        return {
            cMapUrl: usable(this.folders.cMapUrl),
            cMapPacked: true,
            wasmUrl: usable(this.folders.wasmUrl),
            useWorkerFetch: false,
            // pdf.js makes one with `new` and asks it for each file with `fetch({ kind, filename })`.
            BinaryDataFactory: class OfficeParserAssetLoader {
                constructor(_options: unknown) { /* the session holds the folders */ }
                fetch(request: { kind?: unknown; filename?: unknown }): Promise<Uint8Array> {
                    return session.fetchAsset(request);
                }
            },
        };
    }

    /**
     * One file pdf.js asks for: a CMap or a decoder, by its own name, from its folder. Anything else is
     * refused: the standard fonts' data (pdf.js then uses its own metrics, as before) and any name
     * that is not one of pdf.js's files. A file that cannot be had is reported once (see report), and
     * refused at once when pdf.js asks for it again: pdf.js keeps only the CMaps it could load, so each
     * font naming one that could not be had asked for it anew, and a server that never answered cost
     * the parse 15 seconds a font. A file asked for again while it is being read is given that read.
     */
    async fetchAsset(request: { kind?: unknown; filename?: unknown }): Promise<Uint8Array> {
        const kind = request?.kind;
        if (kind !== 'cMapUrl' && kind !== 'wasmUrl') throw new AssetRefused(`officeParser does not load pdf.js's ${String(kind).slice(0, 40)} files`);
        const name = typeof request.filename === 'string' ? request.filename : '';
        if (kind === 'cMapUrl' ? !CMAP_FILE.test(name) : !DECODER_FILES.has(name)) {
            throw new AssetRefused(`${JSON.stringify(name.slice(0, 60))} is not one of pdf.js's ${kind === 'cMapUrl' ? 'CMap' : 'decoder'} files`);
        }
        const known = this.unavailable.get(name);
        if (known) throw this.report(kind, name, known.reason);
        const pending = this.reads.get(name);
        // Each asker has bytes of its own: pdf.js may hand the ones it is given on to its worker.
        if (pending) return (await pending).slice();
        const read = this.read(kind, name);
        this.reads.set(name, read);
        const forget = () => { this.reads.delete(name); };
        read.then(forget, forget);
        return read;
    }

    /** Reads a file pdf.js asked for (see fetchAsset). */
    private async read(kind: PdfAssetKind, name: string): Promise<Uint8Array> {
        const folder = this.folders[kind];
        if (folder.blocked) throw this.report(kind, name, 'blocked');
        if (!folder.url) throw this.report(kind, name, 'missing', 'no folder is set for it');
        const wait = this.waits[kind];
        const left = ASSET_TIMEOUT_MS - wait.waited();
        if (left <= 0) throw this.report(kind, name, 'failed', `files from ${folder.url} have kept this document waiting ${ASSET_TIMEOUT_MS / 1000} s in all, so nothing more is read from there`);
        // Bounded by the time the folder has left, the parse's abortSignal and the parse's end.
        const controller = new AbortController();
        const ranOut = left < ASSET_TIMEOUT_MS
            ? `reading it took longer than the ${Math.ceil(left / 1000)} s left of the ${ASSET_TIMEOUT_MS / 1000} s a document waits on files from one folder`
            : `reading it took longer than ${ASSET_TIMEOUT_MS / 1000} s`;
        const timer = setTimeout(() => controller.abort(new AssetUnavailable('failed', ranOut)), left);
        const sources = [this.ended.signal, this.config.abortSignal].filter((s): s is AbortSignal => !!s);
        const stop = (source: AbortSignal) => () => controller.abort(source.reason);
        const stops = sources.map(source => { const onAbort = stop(source); if (source.aborted) onAbort(); else source.addEventListener('abort', onAbort, { once: true }); return { source, onAbort }; });
        wait.start();
        try {
            const url = new URL(name, folder.url);
            const read = url.protocol === 'file:' && !isBrowser
                ? readLocalFile(folder.url, name, controller.signal)
                : url.protocol === 'http:' || url.protocol === 'https:' || url.protocol === 'file:'
                    ? fetchRemoteFile(url, controller.signal)
                    : Promise.reject(new AssetUnavailable('failed', `${url.protocol} addresses are not read`));
            return await untilAborted(read, controller.signal);
        } catch (e: any) {
            // The parse has ended (or was aborted): nothing to report.
            if (sources.some(source => source.aborted)) throw e;
            const reason = e instanceof AssetUnavailable ? e.reason : 'failed';
            const detail = e instanceof AssetUnavailable ? e.message : `${folder.url}${name}: ${String(e?.message ?? e).slice(0, 200)}`;
            throw this.report(kind, name, reason, detail);
        } finally {
            clearTimeout(timer);
            wait.end();
            for (const { source, onAbort } of stops) source.removeEventListener('abort', onAbort);
        }
    }

    /** Reports a file that could not be had, once per file, and returns the error pdf.js is given. */
    private report(kind: PdfAssetKind, name: string, reason: PdfAssetReason, detail?: string): Error {
        const folder = this.folders[kind];
        if (!this.unavailable.has(name)) {
            this.unavailable.set(name, { kind, reason });
            if (this.reporting || (!this.closed && !this.position)) {
                const page = this.position?.pageBeingRead;
                logWarning(OfficeWarningType.PDF_ASSET_UNAVAILABLE, this.config, { asset: name, kind: kind === 'cMapUrl' ? 'cmap' : 'decoder', reason, option: folder.option, folder: folder.url, detail, page }, undefined, { page });
            }
        }
        return new AssetUnavailable(reason === 'blocked' ? 'missing' : reason, `officeParser could not load ${name} (${reason})`);
    }

    /** The CMaps that could not be had: text in the fonts using them is missing. */
    get unavailableCMaps(): string[] {
        return [...this.unavailable].filter(([, entry]) => entry.kind === 'cMapUrl').map(([name]) => name.replace(/\.bcmap$/, ''));
    }

    /**
     * What a picture pdf.js could not decode is reported with: pdf.js's reason where its worker gave
     * one, and the decoder that could not be had where the picture needed it (or where no reason is
     * known and one could not be).
     */
    undecodablePicture(objId: string): string {
        const reason = this.undecodable.get(objId);
        let text = `pdf.js could not decode the image ${objId}${reason ? ` (${reason})` : ''}`;
        const decoders = [...this.unavailable].filter(([name, entry]) => entry.kind === 'wasmUrl'
            && (!reason || (name === 'jbig2.wasm' ? /jbig2|ccitt/i : /openjpeg|jpx|jpeg ?2000/i).test(reason)));
        if (decoders.length) {
            text += `: the decoder ${decoders.map(([name, entry]) => `${name} (${entry.reason})`).join(', ')} could not be loaded; set pdfDecoderUrl to a folder holding pdf.js's wasm/ files`;
        }
        return text;
    }

    /**
     * A warning line from pdf.js's worker ("Warning: ..."). The ones about content pdf.js left out are
     * reported with the page being read: pictures only when pictures are extracted (`extractAttachments`;
     * without it pdf.js is told to leave every picture out, and says so for each), text always. Others
     * are pdf.js's own notes, which it printed nowhere before either. At most MAX_WORKER_WARNINGS lines
     * of each kind a document are acted on; a line that is not acted on is not counted, so it cannot
     * use up the count before one that is reported.
     */
    workerWarning(line: unknown): void {
        if (typeof line !== 'string') return;
        const kind = pdfjsWarningKind(line);
        if (!kind || !this.reporting || this.position!.spent) return;
        const pictures = !!this.config.extractAttachments;
        if (!pictures && kind !== 'imageUndecodable' && kind !== 'xObjectTextNotRead') return;
        const taken = this.warningsTaken.get(kind) ?? 0;
        if (taken >= MAX_WORKER_WARNINGS) return;
        this.warningsTaken.set(kind, taken + 1);
        const text = line.slice('Warning: '.length, MAX_WARNING_LENGTH);
        const page = this.position!.pageBeingRead;
        const where = page !== undefined ? `from page ${page}` : 'from the document';
        const quoted = (prefix: string) => text.slice(prefix.length).replace(/^"/, '').replace(/"\.?$/, '');
        if (kind === 'imageUndecodable') {
            const match = /^Unable to decode image "([^"]{1,100})": "?(.*?)"?\.?$/s.exec(text);
            if (match) this.undecodable.set(match[1], match[2]);
            return;
        }
        if (kind === 'xObjectTextNotRead') {
            logWarning(OfficeWarningType.CONTENT_PART_NOT_READ, this.config, { part: `an XObject ${page !== undefined ? `on page ${page}` : 'of the document'}`, reason: `pdf.js could not read it (${quoted(PDFJS_WARNINGS.xObjectTextNotRead)})` }, undefined, { page });
            return;
        }
        let lost: string;
        if (kind === 'imageTooLarge') {
            lost = `pdf.js left out a picture larger than pdfParserConfig.maxImagePixels (${this.maxImagePixels} pixels); raise the option to extract it`;
        } else if (kind === 'inlineImageUndecodable') {
            lost = `pdf.js could not decode an inline picture (${quoted(PDFJS_WARNINGS.inlineImageUndecodable)})`;
        } else if (kind === 'imageWithoutSize') {
            lost = 'pdf.js left out a picture whose width or height is missing';
        } else {
            lost = `pdf.js left out an XObject it could not read, with any picture in it (${quoted(PDFJS_WARNINGS.xObjectNotDrawn)})`;
        }
        logWarning(OfficeWarningType.IMAGE_EXTRACTION_FAILED, this.config, where, lost, { page });
    }
}

/** The file a stack frame's code is in, as the stack names it (V8's `at f (file:1:2)`, or `f@file:1:2`). */
function frameFile(frame: string): string | undefined {
    const match = /(?:^\s*at (?:.*? \()?|@)(.+?):\d+:\d+\)?\s*$/.exec(frame);
    return match?.[1];
}

/** The files of the code that called the function this is called in, nearest first (at most `depth`). */
function callerFiles(depth: number): string[] {
    const errorClass = Error as { stackTraceLimit?: number };
    const limit = errorClass.stackTraceLimit;
    try {
        if (typeof limit === 'number') errorClass.stackTraceLimit = depth + 2;
        const stack = new Error().stack;
        if (typeof stack !== 'string') return [];
        const files = stack.split('\n').map(frameFile).filter((file): file is string => !!file);
        // This function's frame, and the caller's own.
        return files.slice(2, depth + 2);
    } catch {
        return [];
    } finally {
        if (typeof limit === 'number') errorClass.stackTraceLimit = limit;
    }
}

/**
 * The file pdf.js's worker runs from in this process, as a stack names it: the module pdf.js's fake
 * worker is given (the caller's preloaded `globalThis.pdfjsWorker`, or the one imported from
 * `workerSrc`), found by asking its `WorkerMessageHandler.setup` to register its listeners with a
 * handler of this function's, which notes where it was called from. Undefined where it cannot be
 * found: the module does not load, `setup` no longer calls the handler, or the runtime gives no stack.
 */
export async function pdfjsWorkerFile(pdfjs: any): Promise<string | undefined> {
    try {
        const handler = await pdfjs.PDFWorker?._setupFakeWorkerGlobal;
        if (typeof handler?.setup !== 'function') return undefined;
        let file: string | undefined;
        handler.setup({ on: () => { file ??= callerFiles(1)[0]; }, send: () => { /* nothing is sent */ } }, null);
        return file;
    } catch {
        return undefined;
    }
}

/**
 * Sends the warnings pdf.js's worker writes in this process to `sink` until the returned function is
 * called. pdf.js's worker module, run in this process, writes them with `console.warn` ("Warning: ...")
 * from its own `warn`: while any sink is open, a line that starts so and was written by code in
 * `workerFile` (see pdfjsWorkerFile) goes to the newest sink and not the console (at verbosity 0
 * pdf.js printed nothing either); every other call reaches the console as before, the caller's own
 * "Warning: " lines with it. A line is pdf.js's when one of the few frames nearest the call is in that
 * file: pdf.js's worker in this process answers the main thread's messages a microtask later, so no
 * code of the caller's runs inside its calls but the `onWarning` a reported line reaches, whose own
 * lines pass on. What this cannot tell apart is left as it is: a line the caller writes from the very
 * file pdf.js's worker is in, or from code called from it (an application bundling pdf.js's worker
 * into its own code), is taken while a PDF is open, and a line pdf.js writes for a document of the
 * caller's own read by the same worker module at that moment is taken as the parse's.
 */
/* eslint-disable no-console -- this takes pdf.js's console warnings; it writes nothing to the console itself */
const consoleSinks: { sink: (line: string) => void; workerFile: string }[] = [];
let consoleWrapper: ((...args: unknown[]) => void) | undefined;
let consoleOriginal: ((...args: unknown[]) => void) | undefined;
/** The frames nearest a `console.warn` call looked at for pdf.js's worker (wrappers others put round the console come first). */
const WARN_FRAMES = 8;
export function captureConsoleWarnings(sink: (line: string) => void, workerFile: string): () => void {
    const entry = { sink, workerFile };
    consoleSinks.push(entry);
    if (!consoleWrapper || console.warn !== consoleWrapper) {
        const passOn = console.warn;
        consoleOriginal = passOn;
        // A line written while one is being reported (an `onWarning` of the caller's writing its own,
        // with pdf.js's `warn` still below it on the stack) is the caller's.
        let reporting = false;
        const wrapper = function (this: unknown, ...args: unknown[]): void {
            const current = consoleSinks[consoleSinks.length - 1];
            if (current && !reporting && typeof args[0] === 'string' && args[0].startsWith('Warning: ') && callerFiles(WARN_FRAMES).includes(current.workerFile)) {
                reporting = true;
                try { current.sink(args[0].slice(0, MAX_WARNING_LENGTH)); } catch { /* a report is never the console's concern */ } finally { reporting = false; }
                return;
            }
            passOn.apply(console, args as []);
        };
        consoleWrapper = wrapper;
        console.warn = wrapper;
    }
    let open = true;
    return () => {
        if (!open) return;
        open = false;
        const index = consoleSinks.lastIndexOf(entry);
        if (index >= 0) consoleSinks.splice(index, 1);
        // Put back only where nothing replaced the wrapper since; a wrapper left in another's chain
        // passes every call on once no sink is open.
        if (!consoleSinks.length && console.warn === consoleWrapper && consoleOriginal) {
            console.warn = consoleOriginal;
            consoleWrapper = undefined;
            consoleOriginal = undefined;
        }
    };
}
/* eslint-enable no-console */

/** A pdf.js worker this module started in the browser, and how to end it. */
export interface WarningWorker {
    /** A pdf.js `PDFWorker` on the worker, for `getDocument({ worker })`. */
    worker: any;
    /** Ends the worker. */
    close(): void;
}

/**
 * pdf.js's worker in a browser Worker that sends pdf.js's warnings back (as `{ officeparserWarning }`
 * messages, which pdf.js's own listeners never see) to `sink`. The Worker runs a small module from a
 * Blob that sends each line of PDFJS_WARNINGS `console.warn` is given (at most MAX_WORKER_WARNINGS of
 * each kind) and passes other calls on, and then imports `workerSrc`;
 * once pdf.js in it says it is ready, a `PDFWorker` is made on it. Undefined where it cannot be made:
 * no Worker or Blob, a Content Security Policy that forbids a `blob:` worker, a worker port the caller
 * set (`GlobalWorkerOptions.workerPort`), or pdf.js not starting in WORKER_START_MS. pdf.js then
 * starts its own worker, as it always did, and the pictures it leaves out go unreported.
 */
export async function openWarningWorker(pdfjs: any, workerSrc: string, sink: (line: string) => void, signal?: AbortSignal | null): Promise<WarningWorker | undefined> {
    if (!workerSrc || pdfjs.GlobalWorkerOptions?.workerPort) return undefined;
    if (typeof Worker === 'undefined' || typeof Blob === 'undefined' || typeof URL?.createObjectURL !== 'function') return undefined;
    let href: string;
    try { href = new URL(workerSrc, typeof document !== 'undefined' ? document.baseURI : undefined).href; } catch { return undefined; }
    const source = `const send = (line) => { try { self.postMessage({ officeparserWarning: line }); } catch { /* not reported */ } };
const kindOf = ${PDFJS_WARNING_KIND_SOURCE};
const sent = new Map();
const passOn = console.warn.bind(console);
console.warn = (...args) => {
    const line = typeof args[0] === 'string' ? args[0] : '';
    if (!line.startsWith('Warning: ')) { passOn(...args); return; }
    const kind = kindOf(line);
    const count = sent.get(kind) || 0;
    if (kind && count < ${MAX_WORKER_WARNINGS}) { sent.set(kind, count + 1); send(line.slice(0, ${MAX_WARNING_LENGTH})); }
};
try { await import(${JSON.stringify(href)}); } catch { self.postMessage({ officeparserWorkerFailed: true }); }`;
    let blobUrl: string | undefined;
    let worker: Worker;
    try {
        blobUrl = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
        worker = new Worker(blobUrl, { type: 'module' });
    } catch {
        if (blobUrl) URL.revokeObjectURL(blobUrl);
        return undefined;
    }
    const listeners = new Set<(event: MessageEvent) => void>();
    let ready!: (started: boolean) => void;
    const started = new Promise<boolean>(resolve => { ready = resolve; });
    worker.addEventListener('message', (event: MessageEvent) => {
        const data = event.data;
        if (data && typeof data === 'object' && typeof data.officeparserWarning === 'string') { sink(data.officeparserWarning); return; }
        if (data && typeof data === 'object' && data.officeparserWorkerFailed === true) { ready(false); return; }
        // pdf.js's first message: its worker is listening.
        if (data && typeof data === 'object' && data.action === 'ready' && !listeners.size) { ready(true); return; }
        for (const listener of [...listeners]) listener(event);
    });
    worker.addEventListener('error', () => ready(false));
    const timer = setTimeout(() => ready(false), WORKER_START_MS);
    const onAbort = () => ready(false);
    signal?.addEventListener('abort', onAbort, { once: true });
    const ok = await started;
    clearTimeout(timer);
    signal?.removeEventListener('abort', onAbort);
    URL.revokeObjectURL(blobUrl);
    if (!ok) { worker.terminate(); return undefined; }
    const port = {
        postMessage: (message: unknown, transfer?: Transferable[]) => worker.postMessage(message, transfer ?? []),
        addEventListener(type: string, listener: (event: MessageEvent) => void, options?: { signal?: AbortSignal }) {
            if (type !== 'message') return;
            listeners.add(listener);
            options?.signal?.addEventListener('abort', () => listeners.delete(listener));
        },
        removeEventListener(_type: string, listener: (event: MessageEvent) => void) { listeners.delete(listener); },
    };
    let pdfWorker: any;
    try { pdfWorker = new pdfjs.PDFWorker({ port, verbosity: 1 }); } catch { worker.terminate(); return undefined; }
    return {
        worker: pdfWorker,
        close: () => {
            try { pdfWorker.destroy(); } catch { /* already ended */ }
            worker.terminate();
        },
    };
}
