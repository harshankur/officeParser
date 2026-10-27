/**
 * Generates the deterministic PDF fixtures used by the parser test suite.
 *
 * These are written as raw PDF syntax (no external tooling) so the bytes are stable across machines
 * and CI. Run once and commit the output:
 *
 *     node scripts/generate-pdf-fixtures.js
 *
 * Outputs (under test/files/pdf/):
 *   - columns_untagged.pdf : two text columns, no structure tree, for reading-order tests.
 *   - rotated.pdf          : the same content on a /Rotate 90 page.
 *   - encrypted.pdf        : single page, RC4 40-bit standard security, user password "test123".
 *   - tagged_headings.pdf  : tagged headings beside headings guessed from font size, for headingDetection.
 *   - page_sizes.pdf       : pages of several sizes (landscape, a /Rotate 90 page, A4), for sections.
 *   - invisible_text.pdf   : text in rendering modes 3 and 7 (a Type 3 font's and Helvetica's) beside and
 *                            among visible text, as a scan's recognised text layer is drawn.
 *   - mixed_size_line.pdf  : one line set in three sizes on one baseline, for the runs' baselines.
 *
 * The encryption is the PDF standard security handler, revision 2 (RC4, 40-bit). RC4 is implemented
 * inline because modern OpenSSL/Node no longer expose it, and qpdf is not assumed to be installed.
 */

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const OUT_DIR = path.join(__dirname, '..', 'test', 'files', 'pdf');

// ── raw PDF assembly ─────────────────────────────────────────────────────────

/** Escapes a string for a PDF literal `( )` string. */
function esc(s) {
    return s.replace(/\\/g, '\\\\').replace(/\(/g, '\\(').replace(/\)/g, '\\)');
}

/** Builds a content stream that draws each { x, y, size, text } line with Helvetica. */
function contentStream(lines) {
    let s = 'BT\n';
    for (const l of lines) {
        s += `/F1 ${l.size} Tf\n1 0 0 1 ${l.x} ${l.y} Tm\n(${esc(l.text)}) Tj\n`;
    }
    s += 'ET\n';
    return Buffer.from(s, 'latin1');
}

/** Object number the first `opts.extra` object gets; 6 stays reserved for /Encrypt. */
const FIRST_EXTRA_OBJ = 7;

/** Wraps a dictionary string and a payload into a stream object. */
function streamObject(dict, payload) {
    return Buffer.concat([
        Buffer.from(`<< ${dict ? `${dict} ` : ''}/Length ${payload.length} >>\nstream\n`, 'latin1'),
        payload,
        Buffer.from('\nendstream', 'latin1'),
    ]);
}

/** Serializes `objects` (index i is object i+1) with an xref table and trailer. */
function serialize(objects, trailerExtra, idHex) {
    const header = Buffer.from('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n', 'latin1');
    let body = header;
    const offsets = [];
    for (let i = 0; i < objects.length; i++) {
        offsets[i] = body.length;
        body = Buffer.concat([body, Buffer.from(`${i + 1} 0 obj\n`, 'latin1'), objects[i], Buffer.from('\nendobj\n', 'latin1')]);
    }
    const xrefOffset = body.length;
    let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (let i = 0; i < objects.length; i++) {
        xref += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`;
    }
    const trailer = `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R${trailerExtra} /ID [<${idHex}> <${idHex}>] >>\nstartxref\n${xrefOffset}\n%%EOF`;
    return Buffer.concat([body, Buffer.from(xref, 'latin1'), Buffer.from(trailer, 'latin1')]);
}

/**
 * Assembles a one-page PDF from a content-stream buffer.
 *
 * `extra` objects are appended from object {@link FIRST_EXTRA_OBJ} on, so a fixture that needs an
 * image XObject or an outline can reference them by a number it knows up front.
 *
 * @param {Buffer} content
 * @param {{ rotate?: number, encrypt?: { userPassword: string }, extra?: Buffer[], resources?: string, fonts?: string, catalog?: string }} opts
 */
function buildPdf(content, opts = {}) {
    const rotate = opts.rotate || 0;
    const objects = []; // objects[i] = Buffer for object (i+1)

    objects[0] = Buffer.from(`<< /Type /Catalog /Pages 2 0 R${opts.catalog || ''} >>`, 'latin1');
    objects[1] = Buffer.from('<< /Type /Pages /Kids [3 0 R] /Count 1 >>', 'latin1');
    const rotateEntry = rotate ? ` /Rotate ${rotate}` : '';
    objects[2] = Buffer.from(
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792]${rotateEntry} ` +
        `/Resources << /Font << /F1 5 0 R${opts.fonts || ''} >>${opts.resources || ''} >> /Contents 4 0 R >>`, 'latin1');
    // object 4 (content) filled below, possibly encrypted
    objects[4] = Buffer.from('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>', 'latin1');

    let encryptRef = '';
    let idHex = '30313233343536373839414243444546'; // fixed 16-byte /ID
    if (opts.encrypt) {
        const enc = buildStandardEncryption(opts.encrypt.userPassword, idHex);
        objects[5] = Buffer.from(
            `<< /Filter /Standard /V 1 /R 2 /O <${enc.O}> /U <${enc.U}> /P -1 >>`, 'latin1');
        // encrypt the content stream with the object key for object 4, gen 0
        content = enc.encryptObject(content, 4, 0);
        encryptRef = ` /Encrypt 6 0 R`;
    }

    objects[3] = streamObject('', content);

    for (let i = 0; i < (opts.extra || []).length; i++) {
        // Object 6 is /Encrypt when the fixture is encrypted; keep the slot either way so the extra
        // objects always land on the same numbers.
        if (!objects[5]) objects[5] = Buffer.from('null', 'latin1');
        objects[FIRST_EXTRA_OBJ - 1 + i] = opts.extra[i];
    }

    return serialize(objects, encryptRef, idHex);
}

// ── RC4 + PDF standard security handler (revision 2, 40-bit) ─────────────────

function rc4(key, data) {
    const s = new Uint8Array(256);
    for (let i = 0; i < 256; i++) s[i] = i;
    let j = 0;
    for (let i = 0; i < 256; i++) {
        j = (j + s[i] + key[i % key.length]) & 0xff;
        [s[i], s[j]] = [s[j], s[i]];
    }
    const out = Buffer.alloc(data.length);
    let a = 0, b = 0;
    for (let k = 0; k < data.length; k++) {
        a = (a + 1) & 0xff;
        b = (b + s[a]) & 0xff;
        [s[a], s[b]] = [s[b], s[a]];
        out[k] = data[k] ^ s[(s[a] + s[b]) & 0xff];
    }
    return out;
}

const PAD = Buffer.from([
    0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
    0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A,
]);

function padPassword(pw) {
    const bytes = Buffer.from(pw, 'latin1');
    return Buffer.concat([bytes, PAD]).subarray(0, 32);
}

function md5(...buffers) {
    const h = crypto.createHash('md5');
    for (const b of buffers) h.update(b);
    return h.digest();
}

/** Builds /O, /U and a per-object encryptor for revision 2 (40-bit RC4). */
function buildStandardEncryption(userPassword, idHex) {
    const userPad = padPassword(userPassword);
    const ownerPad = padPassword(userPassword); // owner == user for the fixture
    const ownerKey = md5(ownerPad).subarray(0, 5);
    const O = rc4(ownerKey, userPad);

    const pLe = Buffer.alloc(4);
    pLe.writeInt32LE(-1, 0);
    const idBytes = Buffer.from(idHex, 'hex');
    const fileKey = md5(userPad, O, pLe, idBytes).subarray(0, 5);
    const U = rc4(fileKey, PAD);

    function encryptObject(data, objNum, gen) {
        const ext = Buffer.from([objNum & 0xff, (objNum >> 8) & 0xff, (objNum >> 16) & 0xff, gen & 0xff, (gen >> 8) & 0xff]);
        const objKey = md5(fileKey, ext).subarray(0, Math.min(fileKey.length + 5, 16));
        return rc4(objKey, data);
    }

    return { O: O.toString('hex'), U: U.toString('hex'), encryptObject };
}

// ── fixture definitions ──────────────────────────────────────────────────────

function twoColumnLines() {
    const lines = [];
    const leftWords = ['ALPHA', 'BETA', 'GAMMA', 'DELTA', 'EPSILON', 'OMEGA'];
    const rightWords = ['ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'NINE'];
    let y = 720;
    for (let i = 0; i < leftWords.length; i++) {
        lines.push({ x: 72, y, size: 12, text: `LEFT COLUMN ${leftWords[i]}` });
        lines.push({ x: 330, y, size: 12, text: `RIGHT COLUMN ${rightWords[i]}` });
        y -= 24;
    }
    return lines;
}

/** A 2x2 RGB image XObject, uncompressed, referenced as /Im0 by the fixtures that place an image. */
function imageXObject() {
    const pixels = Buffer.from([
        0xFF, 0x00, 0x00, 0x00, 0xFF, 0x00,
        0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x00,
    ]);
    return streamObject('/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8', pixels);
}

/**
 * A page that opens with the y-flipping `cm` every Chrome/Skia PDF starts with, then paints an image
 * and three coloured strings (one inside a `q ... Q` bracket, so the restore is exercised too).
 *
 * The flip is the point: text and image positions only line up with what pdf.js reports if the page
 * `cm` and the `q`/`Q` nesting are tracked. Without them the colour marks land mirrored (no colour is
 * ever found) and the image box lands at the bottom of the page instead of the top.
 */
function cmTransformContent() {
    return Buffer.from([
        '1 0 0 -1 0 792 cm',                    // Skia-style flip: the rest of the page is top-down
        'q 200 0 0 -100 50 100 cm /Im0 Do Q',   // image across the top, inside its own q/Q bracket
        '1 0 0 rg',
        'BT /F1 14 Tf 1 0 0 -1 72 140 Tm (RED HEADING TEXT) Tj ET',
        'q 0 0 1 rg BT /F1 12 Tf 1 0 0 -1 72 180 Tm (BLUE BRACKETED TEXT) Tj ET Q',
        'BT /F1 12 Tf 1 0 0 -1 72 220 Tm (RED AGAIN AFTER RESTORE) Tj ET',
        '',
    ].join('\n'), 'latin1');
}

/**
 * A title, a 4x3 grid table, an image off to the right and a closing paragraph - the mix that makes
 * reading-order splicing observable. Generated twice, once unrotated and once on a /Rotate 90 page,
 * so a test can assert the rotation does not reshuffle the flow.
 */
function tableImageContent() {
    const lines = [{ x: 72, y: 740, size: 14, text: 'ALPHA TOP TITLE' }];
    const rows = [['AA1', 'BB1', 'CC1'], ['AA2', 'BB2', 'CC2'], ['AA3', 'BB3', 'CC3'], ['AA4', 'BB4', 'CC4']];
    let y = 700;
    for (const row of rows) {
        lines.push({ x: 72, y, size: 11, text: row[0] });
        lines.push({ x: 220, y, size: 11, text: row[1] });
        lines.push({ x: 380, y, size: 11, text: row[2] });
        y -= 24;
    }
    lines.push({ x: 72, y: 120, size: 12, text: 'OMEGA BOTTOM PARAGRAPH' });
    // The image sits low on the page but far to the right: on a /Rotate 90 page its rendered y is its
    // authored x, so splicing on rendered geometry would drag it to the end of the flow.
    const image = Buffer.from('q 180 0 0 100 380 400 cm /Im0 Do Q\n', 'latin1');
    return Buffer.concat([image, contentStream(lines)]);
}

/**
 * Artifact-marked text in three places: the running-header band, the running-footer band, and the
 * middle of the page (a watermark). Positions are given in authored space; the caller decides the
 * /Rotate, and for the rotated variant the bands are the authored left/right edges, because it is
 * the RENDERED position that decides whether a run is a running header or footer.
 */
function artifactsContent(rotate) {
    const mark = (tag, l) => `/Artifact ${tag} BDC\nBT /F1 ${l.size} Tf 1 0 0 1 ${l.x} ${l.y} Tm (${esc(l.text)}) Tj ET\nEMC`;
    const header = { size: 9, text: 'RUNNING HEADER' };
    const footer = { size: 9, text: 'RUNNING FOOTER' };
    const watermark = { size: 24, text: 'WATERMARK DRAFT' };
    if (rotate === 90) {
        // Rendered y is authored x here: header at the authored left edge, footer at the right one.
        // Both must still end inside the MediaBox - pdf.js drops glyphs painted past the page edge.
        Object.assign(header, { x: 20, y: 400 });
        Object.assign(footer, { x: 530, y: 400 });
        Object.assign(watermark, { x: 200, y: 400 });
    } else {
        Object.assign(header, { x: 72, y: 762 });
        Object.assign(footer, { x: 72, y: 30 });
        Object.assign(watermark, { x: 150, y: 400 });
    }
    const body = contentStream([
        { x: 72, y: 600, size: 12, text: 'FIRST BODY PARAGRAPH' },
        { x: 72, y: 200, size: 12, text: 'SECOND BODY PARAGRAPH' },
    ]);
    return Buffer.concat([
        Buffer.from(`${mark('<< /Type /Pagination /Subtype /Header >>', header)}\n`, 'latin1'),
        Buffer.from(`${mark('<< /Type /Pagination /Subtype /Footer >>', footer)}\n`, 'latin1'),
        Buffer.from(`${mark('<< /Type /Watermark >>', watermark)}\n`, 'latin1'),
        body,
    ]);
}

/**
 * An outline nested `depth` levels deep, one bookmark per level. pdf.js returns such a chain intact,
 * so the parser has to bound its own walk: a recursive conversion of a deep-enough chain overflows
 * the stack and takes the whole parse with it.
 */
function deepOutlineObjects(depth) {
    const rootObj = Buffer.from(`<< /Type /Outlines /First ${FIRST_EXTRA_OBJ + 1} 0 R /Last ${FIRST_EXTRA_OBJ + 1} 0 R /Count ${depth} >>`, 'latin1');
    const items = [];
    for (let i = 0; i < depth; i++) {
        const self = FIRST_EXTRA_OBJ + 1 + i;
        const parent = i === 0 ? FIRST_EXTRA_OBJ : self - 1;
        const child = i === depth - 1 ? '' : ` /First ${self + 1} 0 R /Last ${self + 1} 0 R /Count 1`;
        items.push(Buffer.from(`<< /Title (Level ${i + 1}) /Parent ${parent} 0 R /Dest [3 0 R /Fit]${child} >>`, 'latin1'));
    }
    return [rootObj, ...items];
}

/**
 * A two-page tagged PDF: a heading, a two-item list whose first item carries a footnote, a closing
 * paragraph, and a second list on page two. The footnote text is covered by a Link annotation with
 * an internal destination.
 *
 * It exercises three things no other fixture can: notes that hang off a list item (their runs must
 * be claimed by the tag walk, not left over as stray paragraphs), the internal-link placeholder
 * inside a note body (which is reachable only through `notes`, never through `children`), and list
 * ids that must stay unique across pages because numbering is keyed on them.
 */
function buildTaggedListsPdf() {
    const page1 = Buffer.from([
        '/H1 << /MCID 0 >> BDC BT /F1 18 Tf 1 0 0 1 72 720 Tm (TAGGED FIXTURE HEADING) Tj ET EMC',
        '/Lbl << /MCID 1 >> BDC BT /F1 12 Tf 1 0 0 1 72 680 Tm (1.) Tj ET EMC',
        '/LBody << /MCID 2 >> BDC BT /F1 12 Tf 1 0 0 1 96 680 Tm (First item with a note reference) Tj ET EMC',
        '/Lbl << /MCID 4 >> BDC BT /F1 12 Tf 1 0 0 1 72 650 Tm (2.) Tj ET EMC',
        '/LBody << /MCID 5 >> BDC BT /F1 12 Tf 1 0 0 1 96 650 Tm (Second item without a note) Tj ET EMC',
        '/P << /MCID 6 >> BDC BT /F1 12 Tf 1 0 0 1 72 560 Tm (Closing paragraph after the list) Tj ET EMC',
        '/Note << /MCID 3 >> BDC BT /F1 9 Tf 1 0 0 1 72 600 Tm (1 See the linked note text here) Tj ET EMC',
        '/Note << /MCID 7 >> BDC BT /F1 9 Tf 1 0 0 1 72 520 Tm (2 A second linked note under the paragraph) Tj ET EMC',
        '',
    ].join('\n'), 'latin1');
    const page2 = Buffer.from([
        '/Lbl << /MCID 0 >> BDC BT /F1 12 Tf 1 0 0 1 72 700 Tm (1.) Tj ET EMC',
        '/LBody << /MCID 1 >> BDC BT /F1 12 Tf 1 0 0 1 96 700 Tm (Page two first item) Tj ET EMC',
        '/Lbl << /MCID 2 >> BDC BT /F1 12 Tf 1 0 0 1 72 670 Tm (2.) Tj ET EMC',
        '/LBody << /MCID 3 >> BDC BT /F1 12 Tf 1 0 0 1 96 670 Tm (Page two second item) Tj ET EMC',
        '',
    ].join('\n'), 'latin1');

    const o = [];
    const put = (n, body) => { o[n - 1] = Buffer.from(body, 'latin1'); };
    put(1, '<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 9 0 R /MarkInfo << /Marked true >> >>');
    put(2, '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>');
    put(3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> ' +
        '/Contents 5 0 R /StructParents 0 /Annots [8 0 R 30 0 R] >>');
    put(4, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> ' +
        '/Contents 6 0 R /StructParents 1 >>');
    o[4] = streamObject('', page1);
    o[5] = streamObject('', page2);
    put(7, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    // Link over the footnote text, pointing back at the heading's position on page one.
    put(8, '<< /Type /Annot /Subtype /Link /Rect [72 594 300 612] /Border [0 0 0] /Dest [3 0 R /XYZ 72 730 0] >>');
    put(9, '<< /Type /StructTreeRoot /K [10 0 R] /ParentTree 11 0 R /ParentTreeNextKey 2 >>');
    put(10, '<< /Type /StructElem /S /Document /P 9 0 R /K [12 0 R 13 0 R 21 0 R 22 0 R] >>');
    // A page's /StructParents key indexes into the parent tree; within that array the position is the
    // MCID, so each entry names the element that owns that marked-content sequence.
    put(11, '<< /Nums [0 [12 0 R 15 0 R 16 0 R 17 0 R 19 0 R 20 0 R 21 0 R 29 0 R] 1 [24 0 R 25 0 R 27 0 R 28 0 R]] >>');
    put(12, '<< /Type /StructElem /S /H1 /P 10 0 R /Pg 3 0 R /K 0 >>');
    put(13, '<< /Type /StructElem /S /L /P 10 0 R /K [14 0 R 18 0 R] >>');
    put(14, '<< /Type /StructElem /S /LI /P 13 0 R /K [15 0 R 16 0 R] >>');
    put(15, '<< /Type /StructElem /S /Lbl /P 14 0 R /Pg 3 0 R /K 1 >>');
    put(16, '<< /Type /StructElem /S /LBody /P 14 0 R /Pg 3 0 R /K [2 17 0 R] >>');
    put(17, '<< /Type /StructElem /S /Note /P 16 0 R /Pg 3 0 R /K 3 >>');
    put(18, '<< /Type /StructElem /S /LI /P 13 0 R /K [19 0 R 20 0 R] >>');
    put(19, '<< /Type /StructElem /S /Lbl /P 18 0 R /Pg 3 0 R /K 4 >>');
    put(20, '<< /Type /StructElem /S /LBody /P 18 0 R /Pg 3 0 R /K 5 >>');
    put(21, '<< /Type /StructElem /S /P /P 10 0 R /Pg 3 0 R /K [6 29 0 R] >>');
    put(22, '<< /Type /StructElem /S /L /P 10 0 R /K [23 0 R 26 0 R] >>');
    put(23, '<< /Type /StructElem /S /LI /P 22 0 R /K [24 0 R 25 0 R] >>');
    put(24, '<< /Type /StructElem /S /Lbl /P 23 0 R /Pg 4 0 R /K 0 >>');
    put(25, '<< /Type /StructElem /S /LBody /P 23 0 R /Pg 4 0 R /K 1 >>');
    put(26, '<< /Type /StructElem /S /LI /P 22 0 R /K [27 0 R 28 0 R] >>');
    put(27, '<< /Type /StructElem /S /Lbl /P 26 0 R /Pg 4 0 R /K 2 >>');
    put(28, '<< /Type /StructElem /S /LBody /P 26 0 R /Pg 4 0 R /K 3 >>');
    // A second note, this one under the closing paragraph rather than a list item, with its own link.
    // A footnote reached through `notes` is invisible to a rewrite that only walks `children`, so its
    // internal link is where a leaked `#__pdfsec_k` placeholder shows up.
    put(29, '<< /Type /StructElem /S /Note /P 21 0 R /Pg 3 0 R /K 7 >>');
    put(30, '<< /Type /Annot /Subtype /Link /Rect [72 514 320 532] /Border [0 0 0] /Dest [4 0 R /XYZ 72 730 0] >>');
    return serialize(o, '', '30313233343536373839414243444546');
}

/**
 * A tagged, two-page document where tagged headings and guessed ones meet, for `headingDetection`.
 *
 * Page one is read from its tags: an `H1` and an `H2` set in body-size type (only a tag makes them
 * headings), three `P`s, and one short 30pt line outside the tags, which is spliced in from geometry
 * and is a heading by its size alone. Page two's tree names one of its four marked-content sequences,
 * too little to trust, so the whole page is read from geometry and its 24pt line is a guessed heading.
 */
function buildTaggedHeadingsPdf() {
    const page1 = Buffer.from([
        '/H1 << /MCID 0 >> BDC BT /F1 12 Tf 1 0 0 1 72 700 Tm (Purpose) Tj ET EMC',
        '/P << /MCID 1 >> BDC BT /F1 12 Tf 1 0 0 1 72 670 Tm (This paragraph is tagged body text under the first heading.) Tj ET EMC',
        '/H2 << /MCID 2 >> BDC BT /F1 12 Tf 1 0 0 1 72 630 Tm (Scope) Tj ET EMC',
        '/P << /MCID 3 >> BDC BT /F1 12 Tf 1 0 0 1 72 600 Tm (A second tagged paragraph follows the second heading.) Tj ET EMC',
        '/P << /MCID 4 >> BDC BT /F1 12 Tf 1 0 0 1 72 570 Tm (A third tagged paragraph keeps twelve points the body size.) Tj ET EMC',
        'BT /F1 30 Tf 1 0 0 1 72 760 Tm (2024) Tj ET',
        '',
    ].join('\n'), 'latin1');
    const page2 = Buffer.from([
        '/P << /MCID 0 >> BDC BT /F1 24 Tf 1 0 0 1 72 700 Tm (Annual Report) Tj ET EMC',
        '/P << /MCID 1 >> BDC BT /F1 12 Tf 1 0 0 1 72 660 Tm (The tree of this page names one of its four sequences.) Tj ET EMC',
        '/P << /MCID 2 >> BDC BT /F1 12 Tf 1 0 0 1 72 630 Tm (So the page is read from its geometry instead.) Tj ET EMC',
        '/P << /MCID 3 >> BDC BT /F1 12 Tf 1 0 0 1 72 600 Tm (Its large first line is a heading by size alone.) Tj ET EMC',
        '',
    ].join('\n'), 'latin1');

    const o = [];
    const put = (n, body) => { o[n - 1] = Buffer.from(body, 'latin1'); };
    put(1, '<< /Type /Catalog /Pages 2 0 R /StructTreeRoot 8 0 R /MarkInfo << /Marked true >> >>');
    put(2, '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>');
    put(3, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> ' +
        '/Contents 5 0 R /StructParents 0 >>');
    put(4, '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 7 0 R >> >> ' +
        '/Contents 6 0 R /StructParents 1 >>');
    o[4] = streamObject('', page1);
    o[5] = streamObject('', page2);
    put(7, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    put(8, '<< /Type /StructTreeRoot /K [9 0 R] /ParentTree 10 0 R /ParentTreeNextKey 2 >>');
    put(9, '<< /Type /StructElem /S /Document /P 8 0 R /K [11 0 R 12 0 R 13 0 R 14 0 R 15 0 R 16 0 R] >>');
    put(10, '<< /Nums [0 [11 0 R 12 0 R 13 0 R 14 0 R 15 0 R] 1 [null 16 0 R]] >>');
    put(11, '<< /Type /StructElem /S /H1 /P 9 0 R /Pg 3 0 R /K 0 >>');
    put(12, '<< /Type /StructElem /S /P /P 9 0 R /Pg 3 0 R /K 1 >>');
    put(13, '<< /Type /StructElem /S /H2 /P 9 0 R /Pg 3 0 R /K 2 >>');
    put(14, '<< /Type /StructElem /S /P /P 9 0 R /Pg 3 0 R /K 3 >>');
    put(15, '<< /Type /StructElem /S /P /P 9 0 R /Pg 3 0 R /K 4 >>');
    put(16, '<< /Type /StructElem /S /P /P 9 0 R /Pg 4 0 R /K 1 >>');
    return serialize(o, '', '30313233343536373839414243444547');
}

/**
 * A document of pages of several sizes, for sections read from a PDF's pages: two US Letter pages, a
 * landscape one, an upright one shown landscape (`/Rotate 90`, the same page as the one before it as it
 * is seen), an upright Letter page again and an A4 page. Each page says which it is.
 */
function buildPageSizesPdf() {
    const pages = [
        { box: [612, 792], text: 'Letter page one' },
        { box: [612, 792], text: 'Letter page two' },
        { box: [792, 612], text: 'Landscape page three' },
        { box: [612, 792], rotate: 90, text: 'Turned page four' },
        { box: [612, 792], text: 'Letter page five' },
        { box: [595, 842], text: 'A4 page six' },
    ];
    const o = [];
    const put = (n, body) => { o[n - 1] = Buffer.from(body, 'latin1'); };
    // Objects: 1 catalog, 2 pages, 3 font, then a page and its content for each page.
    put(1, '<< /Type /Catalog /Pages 2 0 R >>');
    put(2, `<< /Type /Pages /Kids [${pages.map((_, i) => `${4 + 2 * i} 0 R`).join(' ')}] /Count ${pages.length} >>`);
    put(3, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>');
    pages.forEach((p, i) => {
        put(4 + 2 * i, `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${p.box[0]} ${p.box[1]}]${p.rotate ? ` /Rotate ${p.rotate}` : ''} ` +
            `/Resources << /Font << /F1 3 0 R >> >> /Contents ${5 + 2 * i} 0 R >>`);
        o[4 + 2 * i] = streamObject('', contentStream([
            { x: 72, y: p.box[1] - 100, size: 18, text: p.text },
            { x: 72, y: p.box[1] - 140, size: 12, text: `This page is ${p.box[0]} by ${p.box[1]} points${p.rotate ? `, shown turned ${p.rotate} degrees` : ''}.` },
        ]));
    });
    return serialize(o, '', '30313233343536373839414243444548');
}

/**
 * A Type 3 font with no BaseFont and no FontName, its glyphs drawing nothing (as a scan's invisible text
 * layer is often set), mapped to the letters of `text` by its ToUnicode: the font, its one empty glyph
 * procedure and its CMap are three objects from `first` on. Returns the objects and the hex string that
 * shows `text` in it.
 */
function type3Font(first, text) {
    const letters = [...new Set(text)];
    const code = new Map(letters.map((c, i) => [c, i + 1]));
    const hex2 = n => n.toString(16).padStart(2, '0').toUpperCase();
    const cmap = [
        '/CIDInit /ProcSet findresource begin 12 dict begin begincmap',
        '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
        '/CMapName /Adobe-Identity-UCS def /CMapType 2 def',
        '1 begincodespacerange <00> <FF> endcodespacerange',
        `${letters.length} beginbfchar`,
        ...letters.map(c => `<${hex2(code.get(c))}> <${c.charCodeAt(0).toString(16).padStart(4, '0').toUpperCase()}>`),
        'endbfchar endcmap CMapName currentdict /CMap defineresource pop end end',
    ].join('\n');
    const names = letters.map(c => `/g${code.get(c)}`).join(' ');
    const procs = letters.map(c => `/g${code.get(c)} ${first + 1} 0 R`).join(' ');
    const font = `<< /Type /Font /Subtype /Type3 /FontBBox [0 0 600 700] /FontMatrix [0.001 0 0 0.001 0 0] ` +
        `/CharProcs << ${procs} >> /Encoding << /Type /Encoding /Differences [1 ${names}] >> ` +
        `/FirstChar 1 /LastChar ${letters.length} /Widths [${letters.map(() => 600).join(' ')}] /Resources << >> /ToUnicode ${first + 2} 0 R >>`;
    return {
        objects: [Buffer.from(font, 'latin1'), streamObject('', Buffer.from('600 0 d0', 'latin1')), streamObject('', Buffer.from(cmap, 'latin1'))],
        hex: [...text].map(c => hex2(code.get(c))).join(''),
    };
}

/**
 * Text drawn invisibly (rendering mode 3, and 7, which adds it to the clip) among visible text: a line in a
 * Type 3 font inside a `q`/`Q` (the mode is restored with the graphics state, so the line after it is
 * seen), a Helvetica line in mode 7, a line whose second half is invisible in the same text object, a
 * large invisible line (no heading by its size), and visible body lines.
 */
function invisibleTextPdf() {
    const scanned = type3Font(FIRST_EXTRA_OBJ, 'Scanned words of a page');
    const content = Buffer.from([
        'BT /F1 12 Tf 1 0 0 1 72 720 Tm (Visible first line of the page.) Tj ET',
        `q BT 3 Tr /F2 12 Tf 1 0 0 1 72 700 Tm <${scanned.hex}> Tj ET Q`,
        'BT /F1 12 Tf 1 0 0 1 72 680 Tm (Visible again after the state is restored.) Tj ET',
        'q BT 7 Tr /F1 12 Tf 1 0 0 1 72 660 Tm (Clipped and unseen words.) Tj ET Q',
        'BT /F1 12 Tf 1 0 0 1 72 640 Tm (Seen words then ) Tj 3 Tr (unseen words.) Tj 0 Tr ET',
        'BT 3 Tr /F1 24 Tf 1 0 0 1 72 600 Tm (Large unseen line) Tj 0 Tr ET',
        'BT /F1 12 Tf 1 0 0 1 72 570 Tm (A visible body line of ordinary text.) Tj ET',
        'BT /F1 12 Tf 1 0 0 1 72 550 Tm (Another visible body line of ordinary text.) Tj ET',
        '',
    ].join('\n'), 'latin1');
    return buildPdf(content, { extra: scanned.objects, fonts: ` /F2 ${FIRST_EXTRA_OBJ} 0 R` });
}

/** One line in three sizes on one baseline (24, 10 and 16 points), and a line under it. */
function mixedSizeLineContent() {
    return Buffer.from([
        'BT /F1 24 Tf 1 0 0 1 72 700 Tm (Large) Tj /F1 10 Tf 1 0 0 1 150 700 Tm (small) Tj /F1 16 Tf 1 0 0 1 190 700 Tm (middle) Tj ET',
        'BT /F1 12 Tf 1 0 0 1 72 670 Tm (A plain line under it.) Tj ET',
        '',
    ].join('\n'), 'latin1');
}

function main() {
    fs.mkdirSync(OUT_DIR, { recursive: true });

    const columns = contentStream(twoColumnLines());
    fs.writeFileSync(path.join(OUT_DIR, 'columns_untagged.pdf'), buildPdf(columns));
    fs.writeFileSync(path.join(OUT_DIR, 'rotated.pdf'), buildPdf(columns, { rotate: 90 }));

    const secret = contentStream([{ x: 72, y: 700, size: 24, text: 'SECRET CONTENT' }]);
    fs.writeFileSync(path.join(OUT_DIR, 'encrypted.pdf'), buildPdf(secret, { encrypt: { userPassword: 'test123' } }));

    const withImage = { extra: [imageXObject()], resources: ` /XObject << /Im0 ${FIRST_EXTRA_OBJ} 0 R >>` };
    fs.writeFileSync(path.join(OUT_DIR, 'cm_transform.pdf'), buildPdf(cmTransformContent(), withImage));

    const tableImage = tableImageContent();
    fs.writeFileSync(path.join(OUT_DIR, 'table_image.pdf'), buildPdf(tableImage, withImage));
    fs.writeFileSync(path.join(OUT_DIR, 'rotated_table_image.pdf'), buildPdf(tableImage, { ...withImage, rotate: 90 }));

    fs.writeFileSync(path.join(OUT_DIR, 'artifacts.pdf'), buildPdf(artifactsContent(0)));
    fs.writeFileSync(path.join(OUT_DIR, 'rotated_artifacts.pdf'), buildPdf(artifactsContent(90), { rotate: 90 }));

    fs.writeFileSync(path.join(OUT_DIR, 'deep_outline.pdf'), buildPdf(
        contentStream([{ x: 72, y: 700, size: 12, text: 'OUTLINE TARGET PAGE' }]),
        { extra: deepOutlineObjects(200), catalog: ` /Outlines ${FIRST_EXTRA_OBJ} 0 R` }));

    fs.writeFileSync(path.join(OUT_DIR, 'tagged_lists.pdf'), buildTaggedListsPdf());
    fs.writeFileSync(path.join(OUT_DIR, 'tagged_headings.pdf'), buildTaggedHeadingsPdf());
    fs.writeFileSync(path.join(OUT_DIR, 'page_sizes.pdf'), buildPageSizesPdf());
    fs.writeFileSync(path.join(OUT_DIR, 'invisible_text.pdf'), invisibleTextPdf());
    fs.writeFileSync(path.join(OUT_DIR, 'mixed_size_line.pdf'), buildPdf(mixedSizeLineContent()));

    console.log('Wrote fixtures to', OUT_DIR);
}

main();
