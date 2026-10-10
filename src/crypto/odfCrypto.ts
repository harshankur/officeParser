/**
 * Decrypts password-protected OpenDocument files (.odt/.ods/.odp/.odg) per the ODF 1.2 packaging
 * spec (§3.4 "Encryption").
 *
 * Unlike OOXML, an encrypted ODF is still an ordinary zip: the entry *names* are visible, but each
 * listed entry's bytes are deflate-compressed and then AES-encrypted, with the per-entry parameters
 * recorded in `META-INF/manifest.xml`. We read the manifest, decrypt and inflate every encrypted
 * entry, then rebuild a plain zip the normal ODF parser can read. Modern AES-256/192/128-CBC is
 * supported; the legacy Blowfish scheme (OpenOffice.org < 3.4) is not, and raises a clear error.
 *
 * Because the input is untrusted, both the zip inflation (via the parser's streaming `extractFiles`
 * guard) and the second, per-entry deflate stream (`inflateRawCapped`) are bounded by
 * `decompressionLimits`, and the attacker-controlled PBKDF2 `iteration-count` is validated first.
 *
 * @module crypto/odfCrypto
 */

import { createDecipheriv, createHash, pbkdf2Sync } from 'crypto';
import { Inflate, zipSync, Zippable } from 'fflate';
import { DecompressionLimits, OfficeParserConfig } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { extractFiles } from '../utils/zipUtils.js';
import { WRONG_PASSWORD, DecryptionError } from './wrongPassword.js';

/**
 * Real ODF files use 1024-100000 PBKDF2 iterations per entry; reject anything above 1,000,000 (10x
 * headroom) as a CPU DoS. Key derivation runs through native WebCrypto (`crypto.subtle`) on the
 * primary path, with `pbkdf2Sync` only as a fallback, so this bounds even the slow fallback to well
 * under a second per entry.
 */
const MAX_ITERATIONS = 1_000_000;
/**
 * Document-wide PBKDF2 budget. An encrypted ODF derives a fresh key for every encrypted entry, so a
 * hostile file with thousands of entries, each individually plausible at the per-entry cap, could
 * still grind for a long time once the correct password is supplied. Real documents write ~100000
 * iterations per entry; even a media-heavy presentation (200+ encrypted image entries, i.e. 20M+
 * iterations) stays under this. Spending the whole budget is not cheap: it costs roughly ten seconds
 * of CPU on the native path, so the budget is carried across the caller's password retries (see
 * {@link decryptOdf}'s `budget`) and unverified work is capped far lower by
 * {@link MAX_UNVERIFIED_ITERATIONS}. Only a file that has already proved its password gets to spend
 * the full amount.
 */
const MAX_TOTAL_ITERATIONS = 50_000_000;
/**
 * PBKDF2 budget before any entry has proved the password. A wrong password (or a hostile file whose
 * entries can never verify) must cost one entry's key derivation, not the whole document budget, and
 * the caller re-asks up to `MAX_PASSWORD_ATTEMPTS` times. Real documents verify on their first
 * encrypted entry, so this leaves ~20 entries of slack at the usual 100000 iterations.
 */
const MAX_UNVERIFIED_ITERATIONS = 2_000_000;
const DEFAULT_MAX_BYTES = 512 * 1024 * 1024;

/**
 * Upper bound on `META-INF/manifest.xml` before it is parsed. The manifest is one short XML element
 * per zip entry (~200 bytes), so even a presentation with thousands of entries stays well under a
 * megabyte; the file is untrusted, so refuse an absurd manifest instead of parsing it.
 */
const MAX_MANIFEST_BYTES = 4 * 1024 * 1024;

/** Detection reads only the tiny manifest; never let a sniff inflate more than this. */
const MAX_SNIFF_BYTES = 4 * 1024 * 1024;
/** A sniff is an inconclusive guess, not something the caller did wrong, so it never reports. */
const DETECTION_SILENT: OfficeParserConfig = { onWarning: () => { } };

interface EncryptionData {
    fullPath: string;
    checksumType: string;
    checksum: Buffer;
    algoName: string;
    iv: Buffer;
    keyDerivName: string;
    keySize: number;
    iterationCount: number;
    salt: Buffer;
    startKeyName: string;
}

/** True when the buffer even looks like a zip (PK\x03\x04); avoids streaming a PDF/RTF through Unzip. */
function isZip(buf: Uint8Array): boolean {
    return buf.length >= 4 && buf[0] === 0x50 && buf[1] === 0x4b && buf[2] === 0x03 && buf[3] === 0x04;
}

/** ODF package extensions (document + template), used as a "this might be ODF" hint from the caller. */
const ODF_EXTS = new Set(['odt', 'ods', 'odp', 'odg', 'ott', 'ots', 'otp', 'otg', 'odf', 'odm']);
function isOdfExt(ext?: string): boolean {
    return !!ext && ODF_EXTS.has(ext.toLowerCase().replace(/^\./, ''));
}

/**
 * Reads the first local file header when it is a `mimetype` entry, so a docx/xlsx/pptx (which start
 * with `[Content_Types].xml`/`_rels`) never pays a full zip walk. Returns `{ stored, content }`, where
 * `stored` is true for an uncompressed entry (method 0) and `content` is the media-type prefix (only
 * read when stored). Returns null when the archive does not start with a `mimetype` entry at all.
 */
function firstMimetypeEntry(buf: Uint8Array): { stored: boolean; content: string } | null {
    // Local file header: sig(4) ver(2) flags(2) method(2)@8 time(2) date(2) crc(4) csize(4) usize(4)
    // namelen(2)@26 extralen(2)@28 name@30.
    if (!isZip(buf) || buf.length < 30) return null;
    const method = buf[8] | (buf[9] << 8);
    const nameLen = buf[26] | (buf[27] << 8);
    const extraLen = buf[28] | (buf[29] << 8);
    if (nameLen !== 8) return null; // "mimetype".length
    if (Buffer.from(buf.subarray(30, 38)).toString('latin1') !== 'mimetype') return null;
    const dataStart = 30 + nameLen + extraLen;
    // Only the media-type prefix is needed, and only a stored entry can be read without inflating;
    // subarray clamps to the buffer so a bogus offset is safe.
    const content = method === 0 ? Buffer.from(buf.subarray(dataStart, dataStart + 64)).toString('latin1') : '';
    return { stored: method === 0, content };
}

/**
 * True when a buffer is worth the (bounded) ODF encryption-manifest sniff. A conformant ODF package
 * (encrypted or not) stores an `application/vnd.oasis.opendocument.*` `mimetype` FIRST, uncompressed
 * (ODF 1.2 §3.3), so that is the fast yes. Two fallbacks keep a NON-conformant encrypted ODF
 * decryptable rather than failing later as unreadable ciphertext: a *deflated* `mimetype` first entry
 * (an EPUB's is always stored per OCF, so a compressed one is not an EPUB), and a caller
 * extension/fileType hint that names an ODF type. OOXML and EPUB take neither branch, so both are still
 * ruled out without walking the archive.
 */
function looksLikeOdfPackage(buf: Uint8Array, extHint?: string): boolean {
    const mt = firstMimetypeEntry(buf);
    if (mt) {
        if (mt.stored) return mt.content.startsWith('application/vnd.oasis.opendocument');
        return true; // deflated mimetype: not an EPUB, treat as a possible non-conformant ODF
    }
    return isOdfExt(extHint);
}

function maxBytesOf(limits?: DecompressionLimits): number {
    const v = limits?.maxUncompressedBytes;
    return v !== undefined && Number.isFinite(v) && v >= 0 ? v : DEFAULT_MAX_BYTES;
}

/** Streaming, limit-enforcing unzip (reuses the parser's zip-bomb guard); returns a name -> bytes map. */
async function unzipCapped(buf: Uint8Array, filter: (name: string) => boolean, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<Record<string, Uint8Array>> {
    const src = Buffer.isBuffer(buf) ? buf : Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength);
    const files = await extractFiles(src, filter, limits ?? {}, config);
    const map: Record<string, Uint8Array> = {};
    for (const f of files) map[f.path] = f.content;
    return map;
}

async function readManifest(buf: Uint8Array, limits?: DecompressionLimits, config?: OfficeParserConfig): Promise<string | null> {
    if (!isZip(buf)) return null;
    const map = await unzipCapped(buf, name => name === 'META-INF/manifest.xml', limits, config);
    const bytes = map['META-INF/manifest.xml'];
    return bytes ? Buffer.from(bytes).toString('utf8') : null;
}

/**
 * True when the buffer is a zip whose `META-INF/manifest.xml` marks at least one entry encrypted.
 * A plain ODF (or any other zip, e.g. a docx) has no `<manifest:encryption-data>` and returns false.
 */
export async function isEncryptedOdf(buf: Uint8Array, limits?: DecompressionLimits, extHint?: string, abortSignal?: AbortSignal | null): Promise<boolean> {
    checkAbortSignal(abortSignal);
    // Fast reject: every zip-backed format (docx/xlsx/pptx/epub) reaches this on every parse, but only
    // an ODF package can be an encrypted ODF. Decide from the first entry's `mimetype` (plus a deflated-
    // mimetype / ODF-extension fallback) instead of streaming the whole archive for a manifest only ODF
    // has, so a non-conformant encrypted ODF still decrypts rather than failing later as ciphertext.
    if (!looksLikeOdfPackage(buf, extHint)) return false;
    try {
        // Sniffing is capped (the manifest is tiny in any real file) and silent (a failed guess is not
        // the caller's fault), exactly like detectOfficeTypeFromZip. The real decrypt below re-reads the
        // manifest under the caller's full budget and reports genuine problems.
        const sniffLimits: DecompressionLimits = {
            ...limits,
            maxUncompressedBytes: Math.min(limits?.maxUncompressedBytes ?? MAX_SNIFF_BYTES, MAX_SNIFF_BYTES),
        };
        const manifest = await readManifest(buf, sniffLimits, { ...DETECTION_SILENT, abortSignal });
        return !!manifest && manifest.includes('encryption-data');
    } catch (error) {
        if (typeof error === 'object' && error !== null && 'name' in error && error.name === 'AbortError') throw error;
        return false;
    }
}

function attr(scope: string, name: string): string {
    const m = new RegExp(`\\b${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}="([^"]*)"`).exec(scope);
    return m ? m[1] : '';
}

/** Parses every `<manifest:file-entry>` that carries encryption metadata. */
function parseEncryptionData(manifest: string): Map<string, EncryptionData> {
    const map = new Map<string, EncryptionData>();
    // An entry is either self-closing or child-bearing: a naive `[\s\S]*?(?:\/>|</...>)` would stop at
    // the first inner self-closing child (e.g. `<manifest:algorithm/>`), truncating the entry before
    // `key-derivation`, and the corrected `[\s\S]*?</manifest:file-entry>` restarts its scan at every
    // start token, so a hostile manifest of unclosed tokens costs O(n^2). Split on the start tag
    // instead and bound each entry to its own piece: one linear pass, identical result on real files
    // (file-entry elements never nest).
    const START = '<manifest:file-entry';
    const CLOSE = '</manifest:file-entry>';
    const pieces = manifest.split(START);
    for (let i = 1; i < pieces.length; i++) {
        const rest = pieces[i];
        if (/^\w/.test(rest)) continue;              // `\b` after the tag name: `<manifest:file-entryX` is not one
        const tagEnd = rest.indexOf('>');
        if (tagEnd < 0) continue;                    // unterminated start tag
        const selfClosing = rest.charCodeAt(tagEnd - 1) === 0x2f; // '/'
        const close = selfClosing ? -1 : rest.indexOf(CLOSE);
        if (!selfClosing && close < 0) continue;     // no closing tag before the next entry
        const block = START + (selfClosing ? rest.slice(0, tagEnd + 1) : rest.slice(0, close + CLOSE.length));
        if (!block.includes('encryption-data')) continue;
        const algo = /<manifest:algorithm\b[^>]*>/.exec(block)?.[0] || '';
        const kd = /<manifest:key-derivation\b[^>]*>/.exec(block)?.[0] || '';
        const skg = /<manifest:start-key-generation\b[^>]*>/.exec(block)?.[0] || '';
        const enc = /<manifest:encryption-data\b[^>]*>/.exec(block)?.[0] || '';
        map.set(attr(block, 'manifest:full-path'), {
            fullPath: attr(block, 'manifest:full-path'),
            checksumType: attr(enc, 'manifest:checksum-type'),
            checksum: Buffer.from(attr(enc, 'manifest:checksum'), 'base64'),
            algoName: attr(algo, 'manifest:algorithm-name'),
            iv: Buffer.from(attr(algo, 'manifest:initialisation-vector'), 'base64'),
            keyDerivName: attr(kd, 'manifest:key-derivation-name'),
            keySize: parseInt(attr(kd, 'manifest:key-size'), 10) || 16,
            iterationCount: parseInt(attr(kd, 'manifest:iteration-count'), 10) || 1024,
            salt: Buffer.from(attr(kd, 'manifest:salt'), 'base64'),
            startKeyName: skg ? attr(skg, 'manifest:start-key-generation-name') : '',
        });
    }
    return map;
}

/** AES key size in bits from the algorithm URI, defaulting to 256 (the ODF 1.2 default). */
function aesBitsFromAlgo(algoName: string): number {
    if (/aes256/i.test(algoName)) return 256;
    if (/aes192/i.test(algoName)) return 192;
    if (/aes128/i.test(algoName)) return 128;
    if (/blowfish/i.test(algoName)) return -1; // signals unsupported
    return 256;
}

/**
 * Checks a decrypted entry against its `manifest:checksum`: `true` when it matches (the password is
 * proved), `false` when it definitely does not (the password is wrong), `null` when the entry carries
 * no checksum we can reproduce.
 *
 * ODF checksums cover the first 1024 bytes of the compressed (pre-encryption) data, but only the "1K"
 * checksum types (`...#sha256-1k`, `SHA1/1K`) hash exactly that prefix. The exact compressed length is
 * no longer known either: CBC left the tail zero-padded to the 16-byte block. Past 1024 + one block the
 * padding cannot reach into the hashed prefix, so a single hash is authoritative; below that we try
 * each of the 16 possible unpadded lengths (16 hashes of at most 1 KiB, negligible) and report a miss
 * as "unknown" rather than "wrong password", since a writer that padded differently must not turn a
 * correct password into a rejection. The raw inflate in the caller settles those.
 */
function checksumVerdict(enc: EncryptionData, compressed: Buffer): boolean | null {
    if (!enc.checksum.length || !/1k/i.test(enc.checksumType)) return null;
    const algo = /sha256/i.test(enc.checksumType) ? 'sha256' : 'sha1';
    const padded = compressed.length <= 1024 + 16;
    for (let pad = 0; pad < (padded ? 16 : 1); pad++) {
        const n = Math.min(compressed.length - pad, 1024);
        if (n <= 0) break;
        const actual = createHash(algo).update(compressed.subarray(0, n)).digest();
        if (actual.subarray(0, enc.checksum.length).equals(enc.checksum)) return true;
    }
    return padded ? null : false;
}

/** Mutable PBKDF2 budget, shared across one document's password attempts. */
export interface KeyDerivationBudget {
    /** Iterations spent so far, against {@link MAX_TOTAL_ITERATIONS}. */
    spent: number;
}

/**
 * Raw-inflates `data`, aborting once the output would exceed `maxBytes`. The input is pushed in
 * bounded chunks so a deflate bomb inflates at most ~one chunk past the cap before we stop, rather
 * than allocating gigabytes. Malformed input (a wrong AES key produces garbage) makes fflate throw,
 * which the caller treats as a wrong password.
 */
function inflateRawCapped(data: Uint8Array, maxBytes: number): Uint8Array {
    const parts: Uint8Array[] = [];
    let total = 0;
    let over = false;
    const inf = new Inflate(chunk => {
        if (over) return;
        total += chunk.length;
        if (total > maxBytes) { over = true; return; }
        parts.push(chunk);
    });
    const CH = 1 << 16;
    for (let o = 0; o < data.length && !over; o += CH) {
        inf.push(data.subarray(o, Math.min(o + CH, data.length)), o + CH >= data.length);
    }
    if (over) throw new DecryptionError('an encrypted ODF entry decompresses beyond the allowed size');
    const out = new Uint8Array(total);
    let off = 0;
    for (const p of parts) { out.set(p, off); off += p.length; }
    return out;
}

/**
 * PBKDF2-HMAC-SHA1, preferring the platform's native WebCrypto. Node's `pbkdf2Sync` is native and
 * fast, but the browser bundle polyfills it in pure JS, where 100000+ iterations per entry freeze the
 * main thread for tens of seconds on a real document. `crypto.subtle.deriveBits` is native in both
 * Node (>= 20) and the browser, and this decrypt path is already async, so we use it when present and
 * fall back to `pbkdf2Sync` only on the rare runtime without WebCrypto.
 */
async function pbkdf2Sha1(startKey: Uint8Array, salt: Uint8Array, iterations: number, bits: number): Promise<Buffer> {
    const subtle = globalThis.crypto?.subtle;
    if (subtle) {
        // Copy into fresh, plain ArrayBuffer-backed views: WebCrypto's BufferSource type rejects the
        // SharedArrayBuffer-capable type of a Node Buffer, and these inputs are tiny (<= 32 bytes).
        const material = await subtle.importKey('raw', new Uint8Array(startKey), 'PBKDF2', false, ['deriveBits']);
        const derived = await subtle.deriveBits({ name: 'PBKDF2', salt: new Uint8Array(salt), iterations, hash: 'SHA-1' }, material, bits);
        return Buffer.from(derived);
    }
    return pbkdf2Sync(Buffer.from(startKey), Buffer.from(salt), iterations, bits / 8, 'sha1');
}

/**
 * Decrypts an encrypted ODF buffer with the given password, returning a plain (re-zipped) ODF.
 * Throws {@link WRONG_PASSWORD} on a bad password, or a {@link DecryptionError} for an unsupported
 * cipher, an implausible key-derivation parameter, a decompression-limit breach, or a broken package.
 *
 * `budget` lets a caller that retries with several passwords (see `decryptIfNeeded`) share one
 * key-derivation budget across those attempts, so a hostile file cannot re-spend the full allowance on
 * every retry. Omitted, each call gets its own.
 */
export async function decryptOdf(buf: Uint8Array, password: string, limits?: DecompressionLimits, config?: OfficeParserConfig, budget?: KeyDerivationBudget): Promise<Uint8Array> {
    checkAbortSignal(config?.abortSignal);
    const manifest = await readManifest(buf, limits, config);
    if (!manifest) throw new DecryptionError('encrypted ODF: missing META-INF/manifest.xml');
    if (manifest.length > MAX_MANIFEST_BYTES) throw new DecryptionError(`encrypted ODF: implausible META-INF/manifest.xml size ${manifest.length} bytes`);
    const encMap = parseEncryptionData(manifest);
    const all = await unzipCapped(buf, () => true, limits, config); // encrypted entries are STORED -> raw bytes
    const maxBytes = maxBytesOf(limits);

    const out: Zippable = {};
    const pwBytes = Buffer.from(password, 'utf8');
    // Both decompression and key-derivation caps are enforced across the whole document, not per entry:
    // a single plausible entry says nothing about a file that repeats it ten thousand times. The
    // key-derivation budget additionally spans the caller's password retries when one is passed in.
    const work = budget ?? { spent: 0 };
    let inflatedTotal = 0;
    // Set once an entry has actually proved the password (a matching checksum, or a clean inflate).
    // Until then only MAX_UNVERIFIED_ITERATIONS may be spent, so a wrong password fails in about the
    // cost of one entry instead of the whole document budget.
    let verified = false;

    for (const [name, bytes] of Object.entries(all)) {
        checkAbortSignal(config?.abortSignal);
        const enc = encMap.get(name);
        if (!enc) { out[name] = bytes; continue; } // mimetype, manifest.xml, directories: copy as-is

        const bits = aesBitsFromAlgo(enc.algoName);
        if (bits < 0) throw new DecryptionError(`encrypted ODF: '${enc.algoName}' (Blowfish) is not supported; only AES-CBC is`);
        if (!/cbc/i.test(enc.algoName)) throw new DecryptionError(`encrypted ODF: unsupported cipher '${enc.algoName}' (only AES-CBC; LibreOffice 24.8+ AES-256-GCM is not supported)`);
        // LibreOffice 24.8+ defaults to Argon2id key derivation, which we do not implement. Reject it
        // loudly instead of silently running PBKDF2 and misreporting the result as a wrong password.
        if (enc.keyDerivName && !/pbkdf2/i.test(enc.keyDerivName)) {
            throw new DecryptionError(`encrypted ODF: unsupported key derivation '${enc.keyDerivName}' (only PBKDF2; LibreOffice 24.8+ Argon2 encryption is not supported)`);
        }
        // iteration-count is attacker-controlled; bound it per entry and across the whole document.
        if (!Number.isFinite(enc.iterationCount) || enc.iterationCount < 1 || enc.iterationCount > MAX_ITERATIONS) {
            throw new DecryptionError(`encrypted ODF: implausible iteration-count ${enc.iterationCount}`);
        }
        work.spent += enc.iterationCount;
        if (work.spent > MAX_TOTAL_ITERATIONS) {
            throw new DecryptionError('encrypted ODF: total key-derivation work exceeds the allowed budget');
        }
        if (!verified && work.spent > MAX_UNVERIFIED_ITERATIONS) {
            throw new DecryptionError('encrypted ODF: no entry verified before the key-derivation budget for an unverified password ran out');
        }

        // Start key: hash of the password; then PBKDF2 (HMAC-SHA1) stretches it to the AES key. The
        // key length comes from the cipher (bits/8), which is inherently one of 16/24/32.
        const startAlgo = /sha256/i.test(enc.startKeyName) ? 'sha256' : 'sha1';
        const startKey = createHash(startAlgo).update(pwBytes).digest();
        const key = await pbkdf2Sha1(startKey, enc.salt, enc.iterationCount, bits);
        // Native key derivation cannot be interrupted; do not start decryption
        // if the caller cancelled while it was running.
        checkAbortSignal(config?.abortSignal);

        // The encrypted length is a whole number of AES blocks; trim any trailing partial block.
        const aligned = bytes.subarray(0, bytes.length - (bytes.length % 16));
        const decipher = createDecipheriv(`aes-${bits}-cbc`, key, enc.iv);
        decipher.setAutoPadding(false);
        const compressed = Buffer.concat([decipher.update(aligned), decipher.final()]);

        // The entry's own checksum is the cheapest password check there is, so consult it on every
        // entry and give up on the first definite mismatch (see checksumVerdict for what "definite"
        // means). A whole-entry checksum type, or a short entry a writer padded unexpectedly, leaves the
        // verdict open; the raw inflate below settles those (it effectively never succeeds on a bad key).
        const verdict = checksumVerdict(enc, compressed);
        if (verdict === false) throw WRONG_PASSWORD;
        if (verdict === true) verified = true;

        let plain: Uint8Array;
        try { plain = inflateRawCapped(compressed, maxBytes - inflatedTotal); }
        catch (e) {
            if (e instanceof DecryptionError) throw e; // a real limit breach, not a wrong password
            throw WRONG_PASSWORD;                      // an undamaged package only fails to inflate on a bad key
        }
        if (plain.length) verified = true; // real deflate output: the key, and so the password, was right
        inflatedTotal += plain.length;
        out[name] = plain;
    }

    checkAbortSignal(config?.abortSignal);
    // The rebuilt zip is fed straight back into the ODF parser, so there is no point recompressing it;
    // storing (level 0) skips the CPU. mimetype must still be first and stored for ODF recognition.
    if (out['mimetype']) {
        const mimetype = out['mimetype'];
        delete out['mimetype'];
        return zipSync({ mimetype: [mimetype as Uint8Array, { level: 0 }], ...out }, { level: 0 });
    }
    return zipSync(out, { level: 0 });
}
