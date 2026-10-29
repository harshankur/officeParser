/**
 * A small reader of a PDF file's structure, for what pdf.js reads and does not give: its cross-reference
 * tables and streams (and their `/Prev` chain), object streams, Flate data with PNG predictors, the page
 * tree with its inherited resources, and the operators of a page's content streams that the readers built
 * on it ask for (marked content, see actualText; the fonts set, see glyphNames). An encrypted file is not
 * read (its streams cannot be), nor a stream in another filter. Bounded: the bytes inflated, the objects,
 * pages and tokens read; past a bound, what is left is not read and pdf.js's reading stands.
 *
 * @module parsers/pdf/pdfFile
 */

import { Unzlib } from 'fflate';

/** At most this many bytes inflated for the whole document, objects read, and forms deep. */
const MAX_INFLATED = 64 * 1024 * 1024;
const MAX_OBJECTS = 200_000;
/** What the search for `endstream` of streams with a length that lies may read beyond twice the file's size. */
const MAX_SEARCH_SLACK = 1024 * 1024;
export const MAX_FORM_DEPTH = 8;
/**
 * Compressed data is inflated this much at a time. Deflate inflates to at most about a thousand times
 * its size, and the inflater grows a buffer to hold all a push gives before handing it on, so a push
 * of 8 KB holds at most some 16 MB: a 64 KB push of a run of spaces held 128 MB to give 64 MB.
 */
const INFLATE_SLICE = 8 * 1024;

export type PdfValue = number | string | boolean | null | PdfName | PdfRef | PdfValue[] | PdfDict | PdfStream | PdfString;
export class PdfName { constructor(readonly name: string) {} }
export class PdfRef { constructor(readonly num: number, readonly gen: number) {} }
export class PdfString { constructor(readonly bytes: Uint8Array) {} }
export type PdfDict = Map<string, PdfValue>;
export interface PdfStream { dict: PdfDict; start: number; length: number; src: Uint8Array }

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
export class Lexer {
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
export function textOf(s: PdfString): string {
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
export class PdfReader {
    private offsets = new Map<number, { at: number } | { stream: number; index: number }>();
    private cache = new Map<number, PdfValue | undefined>();
    private reading = new Set<number>();
    private objStreams = new Map<number, { nums: number[]; offs: number[]; data: Uint8Array; first: number } | null>();
    private inflated = 0;
    /** The bytes searched for an `endstream` so far (see objectAt). */
    private searched = 0;
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
            // A length that cannot be read: the data up to `endstream`. The search is paid for from a budget
            // of the file's own size and a little more, so a file of many streams with a length that lies
            // and no `endstream` after them is not searched to its end once for each.
            if (this.searched > 2 * this.src.length + MAX_SEARCH_SLACK) return { dict: value, start, length: 0, src: this.src };
            const end = indexOf(this.src, 'endstream', start);
            this.searched += (end < 0 ? this.src.length : end) - start;
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

export function isStream(v: PdfValue): v is PdfStream & PdfValue {
    return !!v && typeof v === 'object' && !(v instanceof Map) && 'dict' in (v as object) && 'start' in (v as object);
}

/** The first place `word` is written in `src` from `from`, or -1. */
export function indexOf(src: Uint8Array, word: string, from: number): number {
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

/** A page's content streams, in order (its `/Contents`, one stream or an array of them). */
export function contentStreams(reader: PdfReader, page: PdfDict): (PdfStream & PdfValue)[] {
    const contents = reader.resolve(page.get('Contents'));
    return (Array.isArray(contents) ? contents.map(c => reader.resolve(c)) : [contents]).filter((st): st is PdfStream & PdfValue => !!st && isStream(st));
}

/**
 * The form XObject `Do` draws by its name in `resources`, with the resources it is drawn with (its own,
 * else those it is drawn in); undefined where the name is not a form's.
 */
export function formOf(reader: PdfReader, resources: PdfDict | undefined, name: string): { form: PdfStream & PdfValue; resources: PdfDict | undefined } | undefined {
    const xobjects = resources ? reader.resolve(resources.get('XObject')) : undefined;
    const form = xobjects instanceof Map ? reader.resolve(xobjects.get(name)) : undefined;
    if (!form || !isStream(form) || (reader.resolve(form.dict.get('Subtype')) as PdfName | undefined)?.name !== 'Form') return undefined;
    const own = reader.resolve(form.dict.get('Resources'));
    return { form, resources: own instanceof Map ? own : resources };
}

/** At most this many operators a stream's program holds (see program). */
const MAX_PROGRAM = 500_000;

/**
 * The operators of a content stream the readers built on this one use, in order, as written: a marked
 * content's begin (with its properties) or end, a form drawn with `Do`, a font set with `Tf` (by its name
 * in the resources, none where its operand is no name). A `Tf` with fewer than its two operands is not one,
 * as pdf.js skips it.
 */
export type ContentOp = { op: 'BMC' } | { op: 'BDC'; props: PdfValue | undefined } | { op: 'EMC' } | { op: 'Do'; name: string } | { op: 'Tf'; name: string | undefined };

/**
 * A PDF file as the readers built on this one read it, opened when a page is first asked for: its
 * structure, its pages with their resources, and each content stream's operators (see ContentOp) read
 * once, whatever draws it and whichever reader asks (a stream shared by a thousand pages was read a
 * thousand times), within the document's bound of tokens.
 */
export class PdfFile {
    private reader: PdfReader | undefined;
    private pageList: { dict: PdfDict; resources: PdfDict | undefined }[] | undefined;
    private opened = false;
    private readonly programs = new Map<PdfValue, ContentOp[]>();
    /** One `Tf` operator a name, shared by every program (a stream setting a font a million times holds one). */
    private readonly fontOps = new Map<string | undefined, ContentOp>();
    private tokensLeft = MAX_DOCUMENT_TOKENS;

    constructor(private readonly bytes: Uint8Array, private readonly pageCount: number) {}

    /**
     * A page (index 0 the first) with the file's reader; undefined for a file whose structure cannot be
     * read here (encrypted, damaged, its pages not as many as pdf.js counts), or once the document's tokens
     * are spent.
     */
    page(index: number): { reader: PdfReader; page: { dict: PdfDict; resources: PdfDict | undefined } } | undefined {
        this.open();
        const page = this.pageList?.[index];
        return this.reader && page && this.tokensLeft > 0 ? { reader: this.reader, page } : undefined;
    }

    /**
     * The file's reader, for a reader that wants an object by its number (a picture's stream, see
     * pdfJpeg) and not a page's content; undefined for a file whose structure cannot be read here
     * (encrypted, damaged, its pages not as many as pdf.js counts).
     */
    objects(): PdfReader | undefined {
        this.open();
        return this.reader;
    }

    /** Opens the file once, on the first ask. */
    private open(): void {
        if (this.opened) return;
        this.opened = true;
        try {
            const reader = new PdfReader(this.bytes);
            if (reader.open()) {
                const pages = reader.pages(this.pageCount);
                if (pages.length === this.pageCount) { this.reader = reader; this.pageList = pages; }
            }
        } catch { /* not read here */ }
    }

    /** Takes up to `wanted` of the document's tokens, for a reader reading a stream of its own (a CMap); how many it may read. */
    takeTokens(wanted: number): number {
        const taken = Math.max(0, Math.min(wanted, this.tokensLeft));
        this.tokensLeft -= taken;
        return taken;
    }

    /** Gives back the tokens taken and not read (see takeTokens). */
    returnTokens(unread: number): void {
        this.tokensLeft += Math.max(0, unread);
    }

    /** A content stream's operators (see ContentOp), read once; undefined where its data cannot be read. */
    program(reader: PdfReader, stream: PdfStream & PdfValue): ContentOp[] | undefined {
        const known = this.programs.get(stream);
        if (known) return known;
        const data = reader.decode(stream);
        if (!data) return undefined;
        const ops: ContentOp[] = [];
        const lx = new Lexer(data);
        const operands: (PdfValue | { op: string })[] = [];
        let tokens = Math.min(MAX_TOKENS, this.tokensLeft);
        const budget = tokens;
        while (ops.length < MAX_PROGRAM && tokens-- > 0) {
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
            else if (op === 'Tf' && operands.length >= 2) {
                const operand = operands[operands.length - 2];
                const name = operand instanceof PdfName ? operand.name : undefined;
                let set = this.fontOps.get(name);
                if (!set) { set = { op: 'Tf', name }; this.fontOps.set(name, set); }
                ops.push(set);
            }
            operands.length = 0;
        }
        this.tokensLeft -= budget - Math.max(0, tokens);
        this.programs.set(stream, ops);
        return ops;
    }
}
