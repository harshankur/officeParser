import * as assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { getEventListeners } from 'node:events';
import { readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { zipSync, strToU8 } from 'fflate';
import { decryptIfNeeded } from '../../src/crypto/decryptContainer';
import { decryptOdf, isEncryptedOdf } from '../../src/crypto/odfCrypto';
import { OfficeParser } from '../../src/OfficeParser';
import { detectOfficeTypeFromZip, extractFiles } from '../../src/utils/zipUtils';

const quiet = { onWarning: () => { } };

// Stored entries make the compressed input span multiple feed chunks without
// relying on random data, expensive compression, or a machine-speed deadline. With a signal,
// extraction yields to the event loop once a megabyte of input (16 chunks of 64 KiB), so a first
// entry of 2 MB makes it span turns.
const FIRST_ENTRY_BYTES = 2 * 1024 * 1024;
const archive = Buffer.from(zipSync({
    'first.bin': new Uint8Array(FIRST_ENTRY_BYTES),
    mimetype: strToU8('application/vnd.oasis.opendocument.text'),
    'content.xml': strToU8('<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Example text</text:p></office:text></office:body></office:document-content>'),
}, { level: 0 }));

export async function archiveCancellationTests(): Promise<void> {
    console.log('- archive cancellation...');
    await odfEncryptionCancellationTests();
    const preAborted = new AbortController();
    preAborted.abort();
    let inspected = 0;
    await assert.rejects(extractFiles(archive, () => { inspected++; return true; }, {}, {
        ...quiet, abortSignal: preAborted.signal,
    }), { name: 'AbortError' });
    assert.equal(inspected, 0, 'pre-aborted extraction must not inspect entries');

    await assertMidExtractionAbort(archive);
    // Exercise the inflater too, using repeatable high-entropy bytes so the
    // compressed archive still crosses the input-chunk boundary.
    let seed = 1;
    const payload = Uint8Array.from({ length: FIRST_ENTRY_BYTES }, () => {
        seed ^= seed << 13;
        seed ^= seed >>> 17;
        seed ^= seed << 5;
        return seed & 255;
    });
    const deflated = Buffer.from(zipSync({
        'first.bin': payload,
        'later.txt': strToU8('Later entry'),
    }));
    assert.ok(deflated.length > 1024 * 1024);
    await assertMidExtractionAbort(deflated);
    assert.deepEqual(
        await extractFiles(deflated, () => true, {}, { ...quiet, abortSignal: new AbortController().signal }),
        await extractFiles(deflated, () => true, {}, quiet),
    );

    // An abort queued as the next turn's work runs before extraction's own next turn (immediates run
    // in order), so it lands mid-extraction on any machine.
    const detection = new AbortController();
    const detectionTurn = setImmediate(() => detection.abort());
    try {
        await assert.rejects(detectOfficeTypeFromZip(archive, {}, detection.signal), { name: 'AbortError' });
    } finally {
        clearImmediate(detectionTurn);
    }

    // An abort from a timer lands between turns too: Node's timers run on every turn of the loop.
    const entries: Record<string, Uint8Array> = {};
    for (let i = 0; i < 16; i++) entries[`entry${String(i).padStart(2, '0')}.bin`] = new Uint8Array(1024 * 1024);
    const sixteen = Buffer.from(zipSync(entries, { level: 0 }));
    const timed = new AbortController();
    const timedVisits: string[] = [];
    const timedAbort = setTimeout(() => timed.abort(), 0);
    try {
        await assert.rejects(extractFiles(sixteen, name => { timedVisits.push(name); return true; }, {}, { ...quiet, abortSignal: timed.signal }), { name: 'AbortError' });
        assert.ok(!timedVisits.includes('entry15.bin'), `a timer's abort stops extraction before its last entry; visited ${timedVisits.length}`);
    } finally {
        clearTimeout(timedAbort);
    }

    // Extraction with a signal is about as fast as without: a turn of the event loop once a megabyte,
    // never a timer (clamped to a millisecond or more) after every 64 KiB, which made a 40 MB archive
    // take 850 ms against 10.
    const large = Buffer.from(zipSync(Object.fromEntries(Array.from({ length: 40 }, (_, i) => [`part${i}.bin`, new Uint8Array(1024 * 1024)])), { level: 0 }));
    const timeOf = async (signal?: AbortSignal) => {
        const started = performance.now();
        await extractFiles(large, () => true, {}, { ...quiet, abortSignal: signal });
        return performance.now() - started;
    };
    let without = await timeOf(), withSignal = await timeOf(new AbortController().signal);
    if (withSignal > 2 * without + 50) {
        without = Math.min(without, await timeOf());
        withSignal = Math.min(withSignal, await timeOf(new AbortController().signal));
    }
    assert.ok(withSignal <= 2 * without + 50, `a 40 MB archive takes ${Math.round(withSignal)} ms with a signal, ${Math.round(without)} ms without`);

    // This also proves the entry point checks cancellation before file I/O.
    await assert.rejects(OfficeParser.parseOffice('missing-cancelled-file.odt', {
        ...quiet, abortSignal: preAborted.signal,
    }), { name: 'AbortError' });

    const live = new AbortController();
    const config = { ...quiet, abortSignal: live.signal };
    const actual = await extractFiles(archive, () => true, {}, config);
    assert.deepEqual(actual, await extractFiles(archive, () => true, {}, quiet));
    assert.equal(getEventListeners(live.signal, 'abort').length, 0, 'success must remove its abort listener');
    assert.equal(await detectOfficeTypeFromZip(archive, {}, live.signal), 'odt');
    const parsed = await OfficeParser.parseOffice(archive, { ...config, fileType: 'odt' });
    assert.match(String((await parsed.to('text')).value), /Example text/);
    assert.deepEqual(await extractFiles(archive, () => false, {}, config), []);
    await assert.rejects(extractFiles(Buffer.from('invalid'), () => true, {}, config));
    await assert.rejects(extractFiles(archive.subarray(0, archive.length - 22), () => true, {}, config));
    await assert.rejects(extractFiles(archive, () => true, { maxUncompressedBytes: 1 }, config));
    await assert.rejects(extractFiles(archive, () => true, { maxZipEntries: 1 }, config));
    assert.equal(getEventListeners(live.signal, 'abort').length, 0, 'failures must remove their abort listeners');
    // A late abort after success must not affect another extraction.
    live.abort();
    assert.equal((await extractFiles(archive, () => true, {}, quiet)).length, 3);

    hostShapeTests();
}

/**
 * The shim officeParser's browser bundles once carried in their banner, verbatim: a
 * `setImmediate` without `clearImmediate`, which any page or polyfill may still define.
 */
const BUNDLE_BANNER_SHIM = `// officeparser browser bundle
// Shim for setImmediate (not available in all browsers)
if (typeof setImmediate === 'undefined') {
  window.setImmediate = function(callback) { return setTimeout(callback, 0); };
}`;

/**
 * Hosts extraction may run on, each made in Node by taking globals away (and, for the banner, running
 * its shim), with the turn `eventLoopTurns` takes there: `immediate`, `channel`, `timer` or `microtask`.
 */
const HOST_SHAPES: Array<{ name: string; turn: 'immediate' | 'channel' | 'timer' | 'microtask'; setup: string[] }> = [
    { name: 'banner', turn: 'channel', setup: ['delete g.setImmediate; delete g.clearImmediate;', 'g.window = g;', BUNDLE_BANNER_SHIM] },
    { name: 'none', turn: 'channel', setup: ['delete g.setImmediate; delete g.clearImmediate;'] },
    { name: 'banner without MessageChannel', turn: 'timer', setup: ['delete g.setImmediate; delete g.clearImmediate;', 'g.window = g;', BUNDLE_BANNER_SHIM, 'delete g.MessageChannel;'] },
    { name: 'none without MessageChannel', turn: 'timer', setup: ['delete g.setImmediate; delete g.clearImmediate; delete g.MessageChannel;'] },
    { name: 'bare', turn: 'microtask', setup: ['delete g.setImmediate; delete g.clearImmediate; delete g.MessageChannel; delete g.setTimeout; delete g.clearTimeout;'] },
    { name: 'clearImmediate throws', turn: 'immediate', setup: [`g.clearImmediate = () => { throw new Error('x'); };`] },
];

/**
 * Extraction with a signal settles on every host: where `setImmediate` comes without `clearImmediate`
 * (the old bundle banner's shim), where there is no `setImmediate`, no `MessageChannel`, no timer at all,
 * and where the host's canceller throws. A bundle-banner host left every parse of an archive over a
 * megabyte pending, with an uncaught "clearImmediate is not a function".
 *
 * Each shape runs in a child process, since it takes Node's globals away. The child does not call
 * `process.exit`: its exiting on its own (status 0) within the timeout proves extraction left no port or
 * timer holding the process.
 */
function hostShapeTests(): void {
    for (const shape of HOST_SHAPES) {
        const script = join(tmpdir(), `officeparser-host-shape-${process.pid}-${Math.random().toString(36).slice(2)}.ts`);
        writeFileSync(script, hostShapeScript(shape));
        try {
            const child = spawnSync(process.execPath, [...process.execArgv, script], { encoding: 'utf8', timeout: 60_000 });
            const label = `host shape "${shape.name}" (${shape.turn} turns)`;
            assert.equal(child.status, 0, `${label}: the child exits on its own (status ${child.status}, signal ${child.signal}); ${(child.stderr ?? '').slice(-400)}`);
            assert.equal(child.stderr ?? '', '', `${label}: nothing printed to stderr`);
            const report = JSON.parse((child.stdout ?? '').trim().split('\n').pop() ?? '');
            assert.deepEqual(report.failures, [], `${label}: ${report.failures.join('; ')}`);
            assert.equal(report.uncaught, 0, `${label}: no uncaught exceptions`);
        } finally {
            rmSync(script, { force: true });
        }
    }
}

/** The child's script for one host shape: applies it, extracts under it, and prints one JSON line. */
function hostShapeScript(shape: (typeof HOST_SHAPES)[number]): string {
    const src = join(__dirname, '..', '..', 'src');
    return [
        `const g = globalThis;`,
        // (Its own deadlines, kept from before the shape takes them away.)
        `const realSetTimeout = g.setTimeout, realClearTimeout = g.clearTimeout;`,
        ...shape.setup,
        `let uncaught = 0;`,
        `process.on('uncaughtException', () => { uncaught++; });`,
        `const { zipSync, strToU8 } = require(${JSON.stringify(require.resolve('fflate'))});`,
        `const { extractFiles } = require(${JSON.stringify(join(src, 'utils', 'zipUtils'))});`,
        `const turn = ${JSON.stringify(shape.turn)};`,
        `const quiet = { onWarning: () => {} };`,
        `const failures = [];`,
        `const expect = (ok, message) => { if (!ok) failures.push(message); };`,
        `const settle = (promise, ms) => new Promise(resolve => {`,
        `    const deadline = realSetTimeout(() => resolve({ settled: false }), ms);`,
        `    promise.then(value => { realClearTimeout(deadline); resolve({ settled: true, value }); },`,
        `        error => { realClearTimeout(deadline); resolve({ settled: true, error: error && error.name }); });`,
        `});`,
        `const describe = r => !r.settled ? 'not settled' : r.error ? 'rejected ' + r.error : 'resolved ' + r.value.length;`,
        `(async () => {`,
        `    const large = Buffer.from(zipSync(Object.fromEntries(Array.from({ length: 40 }, (_, i) => ['part' + i + '.bin', new Uint8Array(1024 * 1024)])), { level: 0 }));`,
        // Success: the same entries as without a signal, about as fast (a clamped timer a megabyte, 4 ms in a
        // browser, where the turn is a timer).
        `    const timeOf = async signal => { const started = performance.now(); const r = await settle(extractFiles(large, () => true, {}, { ...quiet, abortSignal: signal }), 20_000); return { ms: performance.now() - started, r }; };`,
        `    const plain = await timeOf();`,
        `    let signalled = await timeOf(new AbortController().signal);`,
        `    expect(plain.r.value && plain.r.value.length === 40, 'without a signal: ' + describe(plain.r));`,
        `    expect(signalled.r.value && signalled.r.value.length === 40 && signalled.r.value.every((e, i) => e.path === plain.r.value[i].path && e.content.equals(plain.r.value[i].content)),`,
        `        'with a signal, the entries read without one: ' + describe(signalled.r));`,
        `    const bound = without => 2 * without + 50 + (turn === 'timer' ? 4 * 40 : 0);`,
        `    let without = plain.ms;`,
        `    if (signalled.ms > bound(without)) { without = Math.min(without, (await timeOf()).ms); signalled = { ms: Math.min(signalled.ms, (await timeOf(new AbortController().signal)).ms) }; }`,
        `    expect(signalled.ms <= bound(without), 'a 40 MB archive takes ' + Math.round(signalled.ms) + ' ms with a signal, ' + Math.round(without) + ' ms without');`,
        // An abort queued as a microtask runs at the end of the turn's synchronous work, before the next turn
        // on every path: it stops the archive before its later entry.
        `    const two = Buffer.from(zipSync({ 'first.bin': new Uint8Array(2 * 1024 * 1024), 'later.txt': strToU8('Later entry') }, { level: 0 }));`,
        `    const mid = new AbortController(); const visited = [];`,
        `    const midResult = await settle(extractFiles(two, name => { visited.push(name); if (name === 'first.bin') queueMicrotask(() => mid.abort()); return true; }, {}, { ...quiet, abortSignal: mid.signal }), 3_000);`,
        `    expect(midResult.error === 'AbortError' && visited.join() === 'first.bin', 'an abort mid-extraction: ' + describe(midResult) + ', visited ' + visited.join());`,
        // An abort from a timer: between turns where they are immediates or timers. Where they are messages,
        // Node's MessagePort drains up to 1,000 queued messages in one go, so timers do not run between them
        // here, as they do in a browser (and Node itself always has setImmediate with clearImmediate); with
        // microtasks no timer runs before the end. There, it settles.
        `    const timed = new AbortController(); const timedVisits = [];`,
        `    realSetTimeout(() => timed.abort(), 0);`,
        `    const timedResult = await settle(extractFiles(large, name => { timedVisits.push(name); return true; }, {}, { ...quiet, abortSignal: timed.signal }), 3_000);`,
        `    if (turn === 'immediate' || turn === 'timer') expect(timedResult.error === 'AbortError' && !timedVisits.includes('part39.bin'), 'an abort from a timer: ' + describe(timedResult) + ', visited ' + timedVisits.length);`,
        `    else expect(timedResult.settled, 'an abort from a timer: ' + describe(timedResult));`,
        `    const pre = new AbortController(); pre.abort(); let inspected = 0;`,
        `    const preResult = await settle(extractFiles(large, () => { inspected++; return true; }, {}, { ...quiet, abortSignal: pre.signal }), 3_000);`,
        `    expect(preResult.error === 'AbortError' && inspected === 0, 'pre-aborted: ' + describe(preResult) + ', inspected ' + inspected);`,
        ...(shape.name === 'banner' ? [
            // A whole parse: a DOCX holding a 2 MB stored picture, with the signal a caller would give.
            `    const { OfficeParser } = require(${JSON.stringify(join(src, 'OfficeParser'))});`,
            `    const docx = Buffer.from(zipSync({`,
            `        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),`,
            `        '_rels/.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'),`,
            `        'word/document.xml': strToU8('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hello from the banner host</w:t></w:r></w:p></w:body></w:document>'),`,
            `        'word/media/image1.bin': [new Uint8Array(2 * 1024 * 1024), { level: 0 }],`,
            `    }));`,
            `    const parsed = await settle(OfficeParser.parseOffice(docx, { ...quiet, abortSignal: AbortSignal.timeout(30_000) }), 10_000);`,
            `    const text = parsed.value ? String((await parsed.value.to('text')).value) : '';`,
            `    expect(/hello from the banner host/.test(text), 'parseOffice with AbortSignal.timeout: ' + (parsed.settled ? parsed.error ? 'rejected ' + parsed.error : JSON.stringify(text.slice(0, 80)) : 'not settled'));`,
        ] : []),
        `})().catch(error => failures.push('threw: ' + (error && error.stack || error))).then(() => console.log(JSON.stringify({ failures, uncaught })));`,
        '',
    ].join('\n');
}

async function assertMidExtractionAbort(input: Buffer): Promise<void> {
    const controller = new AbortController();
    const visited: string[] = [];
    let abortTurn: ReturnType<typeof setImmediate> | undefined;
    try {
        // (Queued before extraction's next turn, so it runs first: see archiveCancellationTests.)
        await assert.rejects(extractFiles(input, name => {
            visited.push(name);
            if (name === 'first.bin') abortTurn = setImmediate(() => controller.abort());
            return true;
        }, {}, { ...quiet, abortSignal: controller.signal }), { name: 'AbortError' });
        assert.deepEqual(visited, ['first.bin'], 'cancellation must stop input before later entries');
        assert.equal(getEventListeners(controller.signal, 'abort').length, 0, 'abort listener must be removed');
        await new Promise(resolve => setTimeout(resolve, 10));
        assert.deepEqual(visited, ['first.bin'], 'no scheduled work may resume after rejection');
    } finally {
        if (abortTurn) clearImmediate(abortTurn);
    }

}

async function odfEncryptionCancellationTests(): Promise<void> {
    // An extension-selected package must sniff its manifest even when mimetype is
    // not the first entry. Its contents only need to mark encryption: cancellation
    // must prevent password handling and actual decryption from starting.
    const encrypted = Buffer.from(zipSync({
        'first.bin': new Uint8Array(FIRST_ENTRY_BYTES),
        'META-INF/manifest.xml': strToU8('<manifest:manifest><manifest:encryption-data/></manifest:manifest>'),
    }, { level: 0 }));
    assert.equal(await isEncryptedOdf(encrypted, {}, 'odt'), true);
    const aborted = new AbortController();
    aborted.abort();
    await assert.rejects(isEncryptedOdf(encrypted, {}, 'odt', aborted.signal), { name: 'AbortError' });

    const sniff = new AbortController();
    const sniffTurn = setImmediate(() => sniff.abort());
    try {
        await assert.rejects(isEncryptedOdf(encrypted, {}, 'odt', sniff.signal), { name: 'AbortError' });
        assert.equal(getEventListeners(sniff.signal, 'abort').length, 0);
    } finally {
        clearImmediate(sniffTurn);
    }

    const parsing = new AbortController();
    let passwordRequests = 0;
    const parseTurn = setImmediate(() => parsing.abort());
    try {
        await assert.rejects(OfficeParser.parseOffice(encrypted, {
            ...quiet, fileType: 'odt', abortSignal: parsing.signal,
            onPassword: async () => { passwordRequests++; return undefined; },
        }), { name: 'AbortError' });
        assert.equal(passwordRequests, 0, 'cancelled sniff must not request a password');
    } finally {
        clearImmediate(parseTurn);
    }

    const decrypting = new AbortController();
    let decryptTurn: ReturnType<typeof setImmediate> | undefined;
    try {
        await assert.rejects(decryptIfNeeded(encrypted, {
            ...quiet, abortSignal: decrypting.signal,
            onPassword: async () => {
                decryptTurn = setImmediate(() => decrypting.abort());
                return 'example password';
            },
        }, 'odt'), { name: 'AbortError' });
        assert.equal(getEventListeners(decrypting.signal, 'abort').length, 0);
    } finally {
        if (decryptTurn) clearImmediate(decryptTurn);
    }

    const passwordPrompt = new AbortController();
    await assert.rejects(decryptIfNeeded(encrypted, {
        ...quiet, abortSignal: passwordPrompt.signal,
        onPassword: async () => { passwordPrompt.abort(); return undefined; },
    }, 'odt'), { name: 'AbortError' });
    await assert.rejects(decryptOdf(encrypted, 'example password', {}, {
        ...quiet, abortSignal: aborted.signal,
    }), { name: 'AbortError' });
    const realEncrypted = readFileSync(join(__dirname, '../files/encrypted/encrypted.odt'));
    const decryptedDocument = await OfficeParser.parseOffice(realEncrypted, {
        ...quiet, fileType: 'odt', password: 'test123', abortSignal: new AbortController().signal,
    });
    assert.match(String((await decryptedDocument.to('text')).value), /SECRET CONTENT 42/);

    const retry = new AbortController();
    let retries = 0;
    await assert.rejects(decryptIfNeeded(realEncrypted, {
        ...quiet, password: 'incorrect password', abortSignal: retry.signal,
        onPassword: async (reason) => {
            assert.equal(reason, 'incorrect');
            retries++;
            retry.abort();
            return undefined;
        },
    }, 'odt'), { name: 'AbortError' });
    assert.equal(retries, 1, 'cancellation must stop password retries');
    assert.equal(await isEncryptedOdf(archive, {}, 'odt', new AbortController().signal), false);
    assert.equal(await isEncryptedOdf(encrypted.subarray(0, 80), {}, 'odt'), false);
}
