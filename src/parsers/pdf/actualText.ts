/**
 * The `/ActualText` of a PDF's marked content: the text a span of glyphs stands for, as a producer
 * writes it where the glyphs alone do not say it (a ligature, a conjunct of an Indic script, a word in
 * glyphs that map to nothing). Chromium and LibreOffice wrap every such cluster in
 * `/Span << /ActualText (...) >> BDC ... EMC`, and map its glyphs to nothing, or all to the first: read
 * from the glyphs, "हिन्दी" came out as "ह ी". pdf.js keeps only a marked-content item's tag and MCID,
 * so the page's content streams are read here, for their marked-content operators alone, in the order
 * pdf.js meets them (forms drawn with `Do` in place): the n-th begin here is the n-th begin pdf.js gives.
 *
 * A small reader of the file's structure: its cross-reference tables and streams (and their `/Prev`
 * chain), object streams, Flate data with PNG predictors, the page tree with its inherited resources.
 * An encrypted file is not read (its streams cannot be), nor a stream in another filter. Bounded: the
 * bytes inflated, the objects and pages looked at, the depth of forms; past a bound, what is left is
 * not read and its glyphs' text stands.
 *
 * @module parsers/pdf/actualText
 */

import { Unzlib } from 'fflate';

/** One marked-content event: a begin (with its ActualText, where it has one) or an end. */
export type MarkedEvent = { begin: true; actualText?: string } | { begin: false };

/** At most this many bytes inflated for the whole document, objects read, and forms deep. */
const MAX_INFLATED = 64 * 1024 * 1024;
const MAX_OBJECTS = 200_000;
const MAX_FORM_DEPTH = 8;
/** At most this many events a page, and for the whole document. */
const MAX_EVENTS = 500_000;
const MAX_DOCUMENT_EVENTS = 10_000_000;
/** Compressed data is inflated this much at a time. */
const INFLATE_SLICE = 64 * 1024;

type PdfValue = number | string | boolean | null | PdfName | PdfRef | PdfValue[] | PdfDict | PdfStream | PdfString;
class PdfName { constructor(readonly name: string) {} }
class PdfRef { constructor(readonly num: number, readonly gen: number) {} }
class PdfString { constructor(readonly bytes: Uint8Array) {} }
type PdfDict = Map<string, PdfValue>;
interface PdfStream { dict: PdfDict; start: number; length: number; src: Uint8Array }

/** The longest word (an operator, a number, a keyword) a content stream holds: longer is none. */
const MAX_WORD = 64;

/** Whether a word is a PDF number: a sign, then digits with at most one point among or before them. */
function isNumberWord(word: string): boolean {
    let i = word[0] === '+' || word[0] === '-' ? 1 : 0, digits = 0, points = 0;
    for (; i < word.length; i++) {
        const c = word.charCodeAt(i);
        if (c >= 0x30 && c <= 0x39) digits++;
        else if (c === 0x2e && points === 0) points++;
        else return false;
    }
    return digits > 0;
}
const isWhite = (c: number) => c === 0x20 || c === 0x0a || c === 0x0d || c === 0x09 || c === 0x0c || c === 0x00;
const isDelim = (c: number) => c === 0x28 || c === 0x29 || c === 0x3c || c === 0x3e || c === 0x5b || c === 0x5d || c === 0x7b || c === 0x7d || c === 0x2f || c === 0x25;

/** A lexer over PDF syntax: objects in the file, and operands and operators in a content stream. */
class Lexer {
    pos: number;
    constructor(readonly src: Uint8Array, start = 0, readonly end = src.length) { this.pos = start; }

    skip(): void {
        const s = this.src;
        while (this.pos < this.end) {
            const c = s[this.pos];
            if (isWhite(c)) this.pos++;
            else if (c === 0x25) { while (this.pos < this.end && s[this.pos] !== 0x0a && s[this.pos] !== 0x0d) this.pos++; }
            else break;
        }
    }

    /** The next token: a value, or an operator (a bare word) as `{ op }`; undefined at the end. */
    next(depth = 0): PdfValue | { op: string } | undefined {
        this.skip();
        const s = this.src;
        if (this.pos >= this.end) return undefined;
        const c = s[this.pos];
        if (c === 0x2f) { // /Name
            let i = ++this.pos;
            while (i < this.end && !isWhite(s[i]) && !isDelim(s[i])) i++;
            let name = '';
            for (let k = this.pos; k < i; k++) name += String.fromCharCode(s[k]);
            this.pos = i;
            return new PdfName(name.replace(/#([0-9a-fA-F]{2})/g, (_, h) => String.fromCharCode(parseInt(h, 16))));
        }
        if (c === 0x28) return this.literal();
        if (c === 0x3c) {
            if (s[this.pos + 1] === 0x3c) { this.pos += 2; return this.dict(depth); }
            return this.hex();
        }
        if (c === 0x5b) {
            this.pos++;
            const out: PdfValue[] = [];
            while (out.length < 1 << 20) {
                this.skip();
                if (this.pos >= this.end) break;
                if (s[this.pos] === 0x5d) { this.pos++; break; }
                const v = depth < 32 ? this.next(depth + 1) : undefined;
                if (v === undefined) break;
                if (v && typeof v === 'object' && 'op' in v) { if (v.op === 'R' && out.length >= 2) { const gen = out.pop() as number, num = out.pop() as number; out.push(new PdfRef(num, gen)); } continue; }
                out.push(v as PdfValue);
            }
            return out;
        }
        if (c === 0x5d || c === 0x3e || c === 0x29 || c === 0x7b || c === 0x7d) { this.pos++; return { op: String.fromCharCode(c) }; }
        // A number, or a word (an operator, true, false, null). A word longer than any a content stream
        // names is none of them (a run of digits a hundred thousand long cost time in its square to test).
        let i = this.pos;
        while (i < this.end && !isWhite(s[i]) && !isDelim(s[i])) i++;
        const start = this.pos;
        this.pos = i;
        if (i - start > MAX_WORD) return { op: '' };
        let word = '';
        for (let k = start; k < i; k++) word += String.fromCharCode(s[k]);
        if (isNumberWord(word)) return Number(word);
        if (word === 'true') return true;
        if (word === 'false') return false;
        if (word === 'null') return null;
        return { op: word };
    }

    private literal(): PdfString {
        const s = this.src, out: number[] = [];
        let depth = 0;
        this.pos++;
        while (this.pos < this.end) {
            const c = s[this.pos++];
            if (c === 0x5c) {
                const e = s[this.pos++];
                const map: Record<number, number> = { 0x6e: 10, 0x72: 13, 0x74: 9, 0x62: 8, 0x66: 12 };
                if (e in map) out.push(map[e]);
                else if (e >= 0x30 && e <= 0x37) {
                    let v = e - 0x30;
                    for (let k = 0; k < 2 && s[this.pos] >= 0x30 && s[this.pos] <= 0x37; k++) v = v * 8 + s[this.pos++] - 0x30;
                    out.push(v & 0xff);
                } else if (e === 0x0d) { if (s[this.pos] === 0x0a) this.pos++; }
                else if (e !== 0x0a && e !== undefined) out.push(e);
            } else if (c === 0x28) { depth++; out.push(c); }
            else if (c === 0x29) { if (depth === 0) break; depth--; out.push(c); }
            else out.push(c);
        }
        return new PdfString(Uint8Array.from(out));
    }

    private hex(): PdfString {
        const s = this.src, out: number[] = [];
        this.pos++;
        let half = -1;
        while (this.pos < this.end) {
            const c = s[this.pos++];
            if (c === 0x3e) break;
            const v = c >= 0x30 && c <= 0x39 ? c - 0x30 : c >= 0x41 && c <= 0x46 ? c - 55 : c >= 0x61 && c <= 0x66 ? c - 87 : -1;
            if (v < 0) continue;
            if (half < 0) half = v; else { out.push(half * 16 + v); half = -1; }
        }
        if (half >= 0) out.push(half * 16);
        return new PdfString(Uint8Array.from(out));
    }

    private dict(depth: number): PdfDict {
        const out: PdfDict = new Map();
        const s = this.src;
        while (out.size < 4096) {
            this.skip();
            if (this.pos >= this.end) break;
            if (s[this.pos] === 0x3e && s[this.pos + 1] === 0x3e) { this.pos += 2; break; }
            const key = this.next(depth + 1);
            if (!(key instanceof PdfName)) { if (key === undefined) break; continue; }
            const at = this.pos;
            let value = depth < 32 ? this.next(depth + 1) : undefined;
            // `n g R`: a reference.
            if (typeof value === 'number') {
                const save = this.pos;
                const gen = this.next(depth + 1), r = this.next(depth + 1);
                if (typeof gen === 'number' && r && typeof r === 'object' && 'op' in r && r.op === 'R') value = new PdfRef(value, gen);
                else this.pos = save;
            }
            if (value === undefined || (value && typeof value === 'object' && 'op' in value)) { if (this.pos === at) this.pos++; continue; }
            out.set(key.name, value as PdfValue);
        }
        return out;
    }
}

/** A string's text: UTF-16BE where it starts with its byte order mark (or UTF-8's), else PDFDocEncoding (as Latin-1). */
function textOf(s: PdfString): string {
    const b = s.bytes;
    if (b.length >= 2 && b[0] === 0xfe && b[1] === 0xff) {
        let out = '';
        for (let i = 2; i + 1 < b.length; i += 2) out += String.fromCharCode((b[i] << 8) | b[i + 1]);
        return out;
    }
    if (b.length >= 3 && b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder().decode(b.subarray(3));
    let out = '';
    for (const c of b) out += String.fromCharCode(c);
    return out;
}

/** Reads a PDF's structure, lazily, within its bounds. */
class PdfReader {
    private offsets = new Map<number, { at: number } | { stream: number; index: number }>();
    private cache = new Map<number, PdfValue | undefined>();
    private reading = new Set<number>();
    private objStreams = new Map<number, { nums: number[]; offs: number[]; data: Uint8Array; first: number } | null>();
    private inflated = 0;
    trailer: PdfDict = new Map();

    constructor(readonly src: Uint8Array) {}

    /** Reads the cross-reference sections from the last; false where none can be read, or the file is encrypted. */
    open(): boolean {
        const tail = new TextDecoder('latin1').decode(this.src.subarray(Math.max(0, this.src.length - 2048)));
        const m = /startxref\s+(\d+)/g;
        let start = -1;
        for (let r = m.exec(tail); r; r = m.exec(tail)) start = Number(r[1]);
        const seen = new Set<number>();
        const queue = [start];
        while (queue.length && seen.size < 64) {
            const at = queue.shift()!;
            if (at < 0 || at >= this.src.length || seen.has(at)) continue;
            seen.add(at);
            for (const next of this.section(at)) queue.push(next);
        }
        return this.offsets.size > 0 && !this.trailer.has('Encrypt') && this.trailer.has('Root');
    }

    /** Merges a section's trailer keys the newer sections did not set. */
    private mergeTrailer(dict: PdfDict): void {
        for (const [k, v] of dict) if (!this.trailer.has(k)) this.trailer.set(k, v);
    }

    /** One cross-reference section (a table or a stream); the offsets of the sections it refers to. */
    private section(at: number): number[] {
        const lx = new Lexer(this.src, at);
        lx.skip();
        const word = lx.next();
        const refers = (dict: PdfDict) => [dict.get('XRefStm'), dict.get('Prev')].filter((v): v is number => typeof v === 'number');
        if (word && typeof word === 'object' && 'op' in word && word.op === 'xref') {
            while (this.offsets.size < MAX_OBJECTS) {
                const save = lx.pos;
                const first = lx.next();
                if (typeof first !== 'number') { lx.pos = save; break; }
                const count = lx.next();
                if (typeof count !== 'number') return [];
                for (let i = 0; i < count && this.offsets.size < MAX_OBJECTS; i++) {
                    const off = lx.next(); lx.next(); const kind = lx.next();
                    if (typeof off !== 'number' || !kind || typeof kind !== 'object' || !('op' in kind)) return [];
                    if (kind.op === 'n' && !this.offsets.has(first + i)) this.offsets.set(first + i, { at: off });
                }
            }
            const t = lx.next();
            if (!(t && typeof t === 'object' && 'op' in t && t.op === 'trailer')) return [];
            const dict = lx.next();
            if (!(dict instanceof Map)) return [];
            this.mergeTrailer(dict);
            return refers(dict);
        }
        // A cross-reference stream: `n g obj << /Type /XRef ... >> stream`.
        const obj = this.objectAt(at);
        if (!obj || !isStream(obj) || (obj.dict.get('Type') as PdfName | undefined)?.name !== 'XRef') return [];
        const data = this.decode(obj);
        if (!data) return [];
        const w = (obj.dict.get('W') as number[] | undefined) ?? [];
        if (w.length < 3 || w.some(v => typeof v !== 'number' || v < 0 || v > 8)) return [];
        const index = (obj.dict.get('Index') as number[] | undefined) ?? [0, Number(obj.dict.get('Size')) || 0];
        const field = (pos: number, width: number, fallback: number) => { if (!width) return fallback; let v = 0; for (let k = 0; k < width; k++) v = v * 256 + data[pos + k]; return v; };
        const row = w[0] + w[1] + w[2];
        let pos = 0;
        for (let k = 0; k + 1 < index.length && this.offsets.size < MAX_OBJECTS; k += 2) {
            for (let i = 0; i < index[k + 1] && pos + row <= data.length && this.offsets.size < MAX_OBJECTS; i++, pos += row) {
                const type = field(pos, w[0], 1), a = field(pos + w[0], w[1], 0), b = field(pos + w[0] + w[1], w[2], 0);
                const num = index[k] + i;
                if (this.offsets.has(num)) continue;
                if (type === 1) this.offsets.set(num, { at: a });
                else if (type === 2) this.offsets.set(num, { stream: a, index: b });
            }
        }
        this.mergeTrailer(obj.dict);
        return refers(obj.dict);
    }

    /** The object written at a byte offset (`n g obj ... endobj`), a stream with its data's place. */
    private objectAt(at: number): PdfValue | undefined {
        const lx = new Lexer(this.src, at);
        const num = lx.next(), gen = lx.next(), kw = lx.next();
        if (typeof num !== 'number' || typeof gen !== 'number' || !(kw && typeof kw === 'object' && 'op' in kw && kw.op === 'obj')) return undefined;
        const value = lx.next();
        if (!(value instanceof Map)) return value && typeof value === 'object' && 'op' in value ? undefined : value as PdfValue;
        const save = lx.pos;
        const after = lx.next();
        if (!(after && typeof after === 'object' && 'op' in after && after.op === 'stream')) { lx.pos = save; return value; }
        let start = lx.pos;
        if (this.src[start] === 0x0d) start++;
        if (this.src[start] === 0x0a) start++;
        let length = this.resolve(value.get('Length'));
        if (typeof length !== 'number' || length < 0 || start + length > this.src.length) {
            // A length that cannot be read: the data up to `endstream`.
            const end = indexOf(this.src, 'endstream', start);
            length = end < 0 ? 0 : end - start;
        }
        return { dict: value, start, length, src: this.src };
    }

    /** An object by its number, read once. */
    get(num: number): PdfValue | undefined {
        if (this.cache.has(num)) return this.cache.get(num);
        if (this.reading.has(num) || this.cache.size > MAX_OBJECTS) return undefined;
        this.reading.add(num);
        let value: PdfValue | undefined;
        const at = this.offsets.get(num);
        if (at && 'at' in at) value = this.objectAt(at.at);
        else if (at) {
            let os = this.objStreams.get(at.stream);
            if (os === undefined) {
                const stream = this.get(at.stream);
                const data = stream && isStream(stream) ? this.decode(stream) : undefined;
                os = null;
                if (data && isStream(stream!)) {
                    const n = Number(stream.dict.get('N')) || 0, first = Number(stream.dict.get('First')) || 0;
                    const lx = new Lexer(data, 0, Math.min(first, data.length));
                    const nums: number[] = [], offs: number[] = [];
                    for (let i = 0; i < n && i < MAX_OBJECTS; i++) {
                        const a = lx.next(), b = lx.next();
                        if (typeof a !== 'number' || typeof b !== 'number') break;
                        nums.push(a); offs.push(b);
                    }
                    os = { nums, offs, data, first };
                }
                this.objStreams.set(at.stream, os);
            }
            if (os && at.index < os.offs.length) {
                const lx = new Lexer(os.data, os.first + os.offs[at.index]);
                const v = lx.next();
                if (!(v && typeof v === 'object' && 'op' in v)) value = v as PdfValue;
            }
        }
        this.reading.delete(num);
        this.cache.set(num, value);
        return value;
    }

    /** A value, a reference followed. */
    resolve(v: PdfValue | undefined): PdfValue | undefined {
        for (let hops = 0; v instanceof PdfRef && hops < 8; hops++) v = this.get(v.num);
        return v instanceof PdfRef ? undefined : v;
    }

    /** A stream's data, decoded: as written, or Flate (with a PNG predictor); undefined for another filter or past the bound. */
    decode(stream: PdfStream): Uint8Array | undefined {
        const raw = stream.src.subarray(stream.start, stream.start + stream.length);
        const filter = this.resolve(stream.dict.get('Filter'));
        const filters = (Array.isArray(filter) ? filter : filter ? [filter] : []).map(f => this.resolve(f));
        if (!filters.length) return raw;
        if (filters.length > 1 || !(filters[0] instanceof PdfName) || (filters[0].name !== 'FlateDecode' && filters[0].name !== 'Fl')) return undefined;
        const left = MAX_INFLATED - this.inflated;
        if (left <= 0) return undefined;
        const chunks: Uint8Array[] = [];
        let size = 0, over = false;
        const inflater = new Unzlib((chunk: Uint8Array) => {
            if (over) return;
            size += chunk.length;
            if (size > left) { over = true; return; }
            chunks.push(chunk);
        });
        // Fed a slice at a time, and no more once past the bound: pushed whole, a megabyte that inflates
        // to a gigabyte was inflated (a gigabyte held, seconds spent) before its size was looked at.
        // A stream cut short, or with a bad check, gives what it inflated.
        try {
            for (let at = 0; at < raw.length && !over; at += INFLATE_SLICE) inflater.push(raw.subarray(at, at + INFLATE_SLICE), at + INFLATE_SLICE >= raw.length);
        } catch { /* what was inflated stands */ }
        this.inflated += Math.min(size, left);
        if (over) { this.inflated = MAX_INFLATED; return undefined; }
        const data = new Uint8Array(size);
        let at = 0;
        for (const c of chunks) { data.set(c, at); at += c.length; }
        const parms = this.resolve(stream.dict.get('DecodeParms'));
        const predictor = parms instanceof Map ? Number(this.resolve(parms.get('Predictor'))) || 1 : 1;
        if (predictor < 10) return data;
        const columns = Number(this.resolve((parms as PdfDict).get('Columns'))) || 1;
        const colors = Number(this.resolve((parms as PdfDict).get('Colors'))) || 1, bits = Number(this.resolve((parms as PdfDict).get('BitsPerComponent'))) || 8;
        const bpp = Math.max(1, Math.ceil(colors * bits / 8)), rowLen = Math.ceil(columns * colors * bits / 8);
        if (rowLen <= 0 || rowLen > 1 << 24) return undefined;
        const rows = Math.floor(data.length / (rowLen + 1));
        const out = new Uint8Array(rows * rowLen);
        for (let r = 0; r < rows; r++) {
            const type = data[r * (rowLen + 1)], src = r * (rowLen + 1) + 1, dst = r * rowLen;
            for (let i = 0; i < rowLen; i++) {
                const x = data[src + i], a = i >= bpp ? out[dst + i - bpp] : 0, b = r ? out[dst - rowLen + i] : 0, c = r && i >= bpp ? out[dst - rowLen + i - bpp] : 0;
                let v = x;
                if (type === 1) v = x + a;
                else if (type === 2) v = x + b;
                else if (type === 3) v = x + ((a + b) >> 1);
                else if (type === 4) { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); v = x + (pa <= pb && pa <= pc ? a : pb <= pc ? b : c); }
                out[dst + i] = v & 0xff;
            }
        }
        return out;
    }

    /** The document's pages, in order, each with the resources it has or inherits. */
    pages(max: number): { dict: PdfDict; resources: PdfDict | undefined }[] {
        const out: { dict: PdfDict; resources: PdfDict | undefined }[] = [];
        const root = this.resolve(this.trailer.get('Root'));
        const top = root instanceof Map ? this.resolve(root.get('Pages')) : undefined;
        const stack: { node: PdfValue | undefined; resources: PdfDict | undefined; depth: number }[] = [{ node: top, resources: undefined, depth: 0 }];
        const seen = new Set<PdfDict>();
        while (stack.length && out.length < max) {
            const { node, resources, depth } = stack.pop()!;
            if (!(node instanceof Map) || seen.has(node) || depth > 64) continue;
            seen.add(node);
            const own = this.resolve(node.get('Resources'));
            const res = own instanceof Map ? own : resources;
            const kids = this.resolve(node.get('Kids'));
            if (Array.isArray(kids)) { for (let i = kids.length - 1; i >= 0; i--) stack.push({ node: this.resolve(kids[i]), resources: res, depth: depth + 1 }); }
            else out.push({ dict: node, resources: res });
        }
        return out;
    }
}

function isStream(v: PdfValue): v is PdfStream & PdfValue {
    return !!v && typeof v === 'object' && !(v instanceof Map) && 'dict' in (v as object) && 'start' in (v as object);
}

/** The first place `word` is written in `src` from `from`, or -1. */
function indexOf(src: Uint8Array, word: string, from: number): number {
    const first = word.charCodeAt(0);
    outer: for (let i = src.indexOf(first, from); i >= 0 && i + word.length <= src.length; i = src.indexOf(first, i + 1)) {
        for (let k = 1; k < word.length; k++) if (src[i + k] !== word.charCodeAt(k)) continue outer;
        return i;
    }
    return -1;
}

/** At most this many tokens are read in one stream, and in the whole document (each stream read once). */
const MAX_TOKENS = 5_000_000;
const MAX_DOCUMENT_TOKENS = 50_000_000;

/**
 * A content stream's marked-content operators, in order, as written: what a page's events are made of
 * with the resources it is drawn with (a named property list, a form drawn with `Do`). Read once for each
 * stream, whatever draws it: a stream shared by a thousand pages was read a thousand times.
 */
type MarkedOp = { op: 'BMC' } | { op: 'BDC'; props: PdfValue | undefined } | { op: 'EMC' } | { op: 'Do'; name: string };

/**
 * The marked-content events of a document's pages, read from the file when a page asks for them (see
 * eventsOf): the structure, each stream's operators and each form's events read once, every read within
 * the document's bounds (bytes inflated, tokens, events), past which a page has none and its glyphs'
 * text stands.
 */
export class MarkedContent {
    private reader: PdfReader | undefined;
    private pages: { dict: PdfDict; resources: PdfDict | undefined }[] | undefined;
    private opened = false;
    private readonly programs = new Map<PdfValue, MarkedOp[]>();
    /** A form's events, with the resources it was last drawn with (pages of resources of their own each keep none). */
    private readonly forms = new Map<PdfValue, { resources: PdfDict | undefined; events: MarkedEvent[] }>();
    private tokensLeft = MAX_DOCUMENT_TOKENS;
    private eventsLeft = MAX_DOCUMENT_EVENTS;

    constructor(private readonly bytes: Uint8Array, private readonly pageCount: number) {}

    /**
     * A page's events (index 0 the first page's), in the order pdf.js meets them; undefined for a file
     * whose structure cannot be read here (encrypted, damaged, in a filter not read), or once the
     * document's bounds are spent. A page whose content cannot be read has none.
     */
    eventsOf(index: number): MarkedEvent[] | undefined {
        try {
            if (!this.opened) {
                this.opened = true;
                const reader = new PdfReader(this.bytes);
                if (reader.open()) {
                    const pages = reader.pages(this.pageCount);
                    if (pages.length === this.pageCount) { this.reader = reader; this.pages = pages; }
                }
            }
            const page = this.pages?.[index];
            if (!this.reader || !page || this.eventsLeft <= 0 || this.tokensLeft <= 0) return undefined;
            const events = this.pageEvents(this.reader, page);
            this.eventsLeft -= events.length;
            return events;
        } catch {
            return [];
        }
    }

    /** A stream's operators (see MarkedOp), read once. */
    private program(reader: PdfReader, stream: PdfStream & PdfValue): MarkedOp[] | undefined {
        const known = this.programs.get(stream);
        if (known) return known;
        const data = reader.decode(stream);
        if (!data) return undefined;
        const ops: MarkedOp[] = [];
        const lx = new Lexer(data);
        const operands: (PdfValue | { op: string })[] = [];
        let tokens = Math.min(MAX_TOKENS, this.tokensLeft);
        const budget = tokens;
        while (ops.length < MAX_EVENTS && tokens-- > 0) {
            const t = lx.next();
            if (t === undefined) break;
            if (!(t && typeof t === 'object' && 'op' in t)) { if (operands.length < 64) operands.push(t); continue; }
            const op = t.op;
            if (op === 'BI') {
                // An inline image: its data, to the `EI` after whitespace, is no operator.
                const id = indexOf(data, 'ID', lx.pos);
                let end = id < 0 ? -1 : indexOf(data, 'EI', id + 3);
                while (end > 0 && !(isWhite(data[end - 1]) && (end + 2 >= data.length || isWhite(data[end + 2]) || isDelim(data[end + 2])))) end = indexOf(data, 'EI', end + 2);
                lx.pos = end < 0 ? data.length : end + 2;
            } else if (op === 'BMC' || op === 'EMC') ops.push({ op });
            else if (op === 'BDC') { const props = operands[operands.length - 1]; ops.push({ op, props: props && typeof props === 'object' && 'op' in props ? undefined : props as PdfValue }); }
            else if (op === 'Do') { const name = operands[operands.length - 1]; if (name instanceof PdfName) ops.push({ op, name: name.name }); }
            operands.length = 0;
        }
        this.tokensLeft -= budget - Math.max(0, tokens);
        this.programs.set(stream, ops);
        return ops;
    }

    /** A page's events: its content streams' operators with its resources, forms drawn in place. */
    private pageEvents(reader: PdfReader, page: { dict: PdfDict; resources: PdfDict | undefined }): MarkedEvent[] {
        /** The events of operators drawn with `resources` (EMC included even at no depth, see below). */
        const play = (ops: MarkedOp[], resources: PdfDict | undefined, depth: number, active: Set<PdfValue>, events: MarkedEvent[]) => {
            for (const o of ops) {
                if (events.length >= MAX_EVENTS) return;
                if (o.op === 'BMC') events.push({ begin: true });
                else if (o.op === 'EMC') events.push({ begin: false });
                else if (o.op === 'BDC') {
                    let props = o.props;
                    if (props instanceof PdfName) {
                        const named = resources ? reader.resolve(resources.get('Properties')) : undefined;
                        props = named instanceof Map ? reader.resolve(named.get(props.name)) : undefined;
                    }
                    const actual = props instanceof Map ? reader.resolve(props.get('ActualText')) : undefined;
                    events.push(actual instanceof PdfString ? { begin: true, actualText: textOf(actual) } : { begin: true });
                } else if (depth < MAX_FORM_DEPTH) {
                    const xobjects = resources ? reader.resolve(resources.get('XObject')) : undefined;
                    const form = xobjects instanceof Map ? reader.resolve(xobjects.get(o.name)) : undefined;
                    if (!form || !isStream(form) || (reader.resolve(form.dict.get('Subtype')) as PdfName | undefined)?.name !== 'Form' || active.has(form)) continue;
                    const own = reader.resolve(form.dict.get('Resources'));
                    const res = own instanceof Map ? own : resources;
                    // A form's events are read once for the resources it is drawn with, and repeated where it is
                    // drawn again: a form drawn ten times in each of eight forms nested would be read a hundred
                    // million times.
                    const cached = this.forms.get(form);
                    let drawn = cached && cached.resources === res ? cached.events : undefined;
                    if (!drawn) {
                        drawn = [];
                        const body = this.program(reader, form);
                        active.add(form);
                        if (body) play(body, res, depth + 1, active, drawn);
                        active.delete(form);
                        this.forms.set(form, { resources: res, events: drawn });
                    }
                    for (let k = 0; k < drawn.length && events.length < MAX_EVENTS; k++) events.push(drawn[k]);
                }
            }
        };
        const contents = reader.resolve(page.dict.get('Contents'));
        const streams = (Array.isArray(contents) ? contents.map(c => reader.resolve(c)) : [contents]).filter((st): st is PdfStream & PdfValue => !!st && isStream(st));
        const raw: MarkedEvent[] = [];
        // The streams are one content stream, as a page's are.
        for (const st of streams) {
            const ops = this.program(reader, st);
            if (!ops) return [];
            play(ops, page.resources, 0, new Set(), raw);
        }
        // An end with nothing open is dropped, as pdf.js drops it.
        const events: MarkedEvent[] = [];
        let level = 0;
        for (const e of raw) {
            if (e.begin) { level++; events.push(e); }
            else if (level > 0) { level--; events.push(e); }
        }
        return events;
    }
}

/** A pdf.js text content item, as far as reading it needs. */
interface TextContentItem { type?: string; str?: string; transform?: number[]; width?: number }

/**
 * A page's text items as they read: each span with an ActualText (see readMarkedContent) as one item of
 * that text over the span's width (its glyphs' own text, in it, left out), and the control characters a
 * font's lost mapping gives (U+0002, U+0005) taken out of every item. The events are used only where
 * they are as many as pdf.js's marked-content begins (else the page is as pdf.js reads it).
 */
export function readableItems<T extends TextContentItem>(items: T[], events: MarkedEvent[] | undefined): T[] {
    return withoutCoveredSpaces(advanced(spans(items, events)));
}

/**
 * Items of whitespace alone drawn inside a glyph of their line (a space pdf.js put where a glyph's width
 * was lost, over the glyphs after it) left out: no gap is there. A space lies inside a glyph when most of
 * it is in the glyph's box, or it starts before a glyph before it ends; a space whose end runs a little
 * into the next word (as one a producer sets between words of right-to-left text, or at a font's change)
 * is a word's space and stays. Looked for among the few items drawn around each.
 */
function withoutCoveredSpaces<T extends TextContentItem>(items: T[]): T[] {
    const level = (it: T) => it.transform && it.transform.length >= 6 && Math.abs(it.transform[1]) < 1e-6 && Math.abs(it.transform[2]) < 1e-6
        && it.transform.every(Number.isFinite);
    return items.filter((item, i) => {
        if (typeof item.str !== 'string' || item.str.trim() || !item.str || !item.width || !Number.isFinite(item.width) || !level(item)) return true;
        const x0 = item.transform![4], x1 = x0 + item.width, y = item.transform![5];
        for (let k = Math.max(0, i - 4); k < Math.min(items.length, i + 5); k++) {
            const o = items[k];
            if (k === i || typeof o.str !== 'string' || !o.str.trim() || !level(o) || !(Math.abs(o.transform![5] - y) <= 0.5)) continue;
            const ox0 = o.transform![4], ox1 = ox0 + (o.width || 0);
            const overlap = Math.min(x1, ox1) - Math.max(x0, ox0);
            // (One reaching a fifth of an em past the glyph is that glyph's lost width and a word space after it.)
            const em = Math.abs(item.transform![0]) || Math.abs(o.transform![0]) || 1;
            if ((overlap > 0.5 * (x1 - x0) || (ox0 < x0 - 0.5 && ox1 > x0 + 0.5)) && !(x1 - ox1 >= 0.2 * em)) return false;
        }
        return true;
    });
}

/** The user-perceived characters in a text: its grapheme clusters, or its letters where they cannot be told. */
function clusters(text: string): number {
    const Segmenter = (Intl as unknown as { Segmenter?: new (locale?: string, options?: object) => { segment(t: string): Iterable<unknown> } }).Segmenter;
    if (Segmenter) { let n = 0; for (const _ of new Segmenter(undefined, { granularity: 'grapheme' }).segment(text)) { void _; n++; } return Math.max(1, n); }
    return Math.max(1, (text.match(/\p{L}/gu) ?? []).length);
}

/**
 * Items pdf.js gave no width (a glyph whose text holds a virama or another mark it takes for a
 * diacritic, as an Indic conjunct's does: its own width lost, and the text after it on its line drawn
 * that much too far left, so read out of order) given one (nine tenths of an em a cluster), and the items after
 * them on their line moved by it, in the order they are drawn. A mark alone keeps none. The space
 * pdf.js puts before the text after such an item (the gap its lost width left) is no space: "क्ष त्रिय"
 * for "क्षत्रिय".
 */
function advanced<T extends TextContentItem>(items: T[]): T[] {
    // (The line's baseline: none before its first item, so the shift starts afresh on every line, the first too.)
    // (`lost`: the item without its width last given one, until the next item on its line shows whether pdf.js
    // already left the room: one drawn at least a third of an em past it was, and is not moved.)
    let shift = 0, lineY: number | undefined, afterLost = false, lost: { x: number; width: number; em: number } | undefined;
    return items.map(item => {
        const t = item.transform;
        if (typeof item.str !== 'string' || !t || t.length < 6 || !t.every(Number.isFinite) || Math.abs(t[1]) > 1e-6 || Math.abs(t[2]) > 1e-6) return item;
        if (lineY === undefined || !(Math.abs(t[5] - lineY) <= 0.5)) { shift = 0; lineY = t[5]; afterLost = false; lost = undefined; }
        else if (lost && item.str.trim()) {
            // (A blank decides nothing: pdf.js puts one over the lost glyph, as wide as it and the word space after it.)
            // (Room left: the space before this item, if any, is the word's, not the lost width's gap.)
            if (t[4] >= lost.x + lost.em / 3) { shift -= lost.width; afterLost = false; }
            lost = undefined;
        }
        let moved = shift ? { ...item, transform: [t[0], t[1], t[2], t[3], t[4] + shift, t[5]] } : item;
        if (afterLost && /^\s/.test(item.str) && item.str.trim()) moved = { ...moved, str: item.str.replace(/^\s+/, '') };
        if (item.str.trim()) afterLost = false;
        if (!item.width && /\p{L}/u.test(item.str) && /[\p{Mn}\p{Mc}]/u.test(item.str)) {
            afterLost = true;
            // Nine tenths of an em a cluster (more than less: text drawn over its neighbour still reads in
            // order, where a gap left would read as a space, or a column's edge).
            const width = 0.9 * Math.abs(t[0]) * clusters(item.str.trim());
            shift += width;
            lost = { x: t[4], width, em: Math.abs(t[0]) || 1 };
            return { ...moved, width };
        }
        return moved;
    });
}

function spans<T extends TextContentItem>(items: T[], events: MarkedEvent[] | undefined): T[] {
    const controls = /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g;
    const clean = (item: T): T => typeof item.str === 'string' && controls.test(item.str) ? { ...item, str: item.str.replace(controls, '') } : item;
    const isBegin = (item: T) => item.type === 'beginMarkedContent' || item.type === 'beginMarkedContentProps';
    const begins = events?.filter(e => e.begin) as ({ begin: true; actualText?: string })[] | undefined;
    if (!begins?.some(e => e.actualText !== undefined) || items.filter(isBegin).length !== begins.length) return items.map(clean);
    const out: T[] = [];
    let at = 0, depth = 0, text = '';
    let target: T | undefined, minX = Infinity, maxX = -Infinity, level = true;
    const finish = () => {
        if (target && level && Number.isFinite(minX)) {
            target.transform = [...target.transform!];
            target.transform[4] = minX;
            target.width = maxX - minX;
        }
        target = undefined;
    };
    for (const item of items) {
        if (isBegin(item)) {
            const event = begins[at++];
            out.push(item);
            if (depth) depth++;
            else if (event?.actualText !== undefined) { depth = 1; text = event.actualText; target = undefined; minX = Infinity; maxX = -Infinity; level = true; }
            continue;
        }
        if (item.type === 'endMarkedContent') {
            if (depth && --depth === 0) finish();
            out.push(item);
            continue;
        }
        if (!depth) { out.push(clean(item)); continue; }
        // In a span with an ActualText: the first glyph's item holds it, over every glyph's width.
        const t = item.transform;
        if (t && t.length >= 6) {
            if (Math.abs(t[1]) > 1e-6 || Math.abs(t[2]) > 1e-6) level = false;
            minX = Math.min(minX, t[4]);
            maxX = Math.max(maxX, t[4] + (item.width || 0));
        }
        if (!target) {
            // A space pdf.js put before the glyphs (a ligature's own width lost) stays before the text.
            target = { ...item, str: (/^\s/.test(item.str ?? '') ? ' ' : '') + text.replace(controls, '') };
            out.push(target);
        }
    }
    finish();
    return out;
}
