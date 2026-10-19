#!/usr/bin/env node
/**
 * OCR language data for the officeParser site, served from the owner's own Cloudflare R2 storage
 * (https://tessdata.harshankur.com) so the site's OCR fetches nothing from a CDN. Every language
 * published in the @tesseract.js-data npm scope is there (more than tesseract.js's own list names), as
 * the site's language field takes any code, and tesseract.js fetched any of them from jsDelivr.
 *
 *   node scripts/upload-tessdata.mjs download [--langs eng,deu] [--out .tessdata]
 *       Downloads each language's LSTM data (@tesseract.js-data/<lang>, Apache-2.0, from the npm
 *       registry) into <out>/4.0.0_best_int/<lang>.traineddata.gz.
 *   node scripts/upload-tessdata.mjs cors --bucket <name>
 *       Lets the site (and a local preview) read the bucket: GET and HEAD from its origins.
 *   node scripts/upload-tessdata.mjs upload --bucket <name> [--out .tessdata]
 *       Puts every downloaded file in the bucket under 4.0.0_best_int/, with wrangler (log in first
 *       with `npx wrangler login`).
 *
 * The bucket is made, and given its domain, once:
 *   npx wrangler r2 bucket create <name>
 *   npx wrangler r2 bucket domain add <name> --domain tessdata.harshankur.com --zone-id <harshankur.com's zone id>
 *
 * The files keep their gzip compression as data: tesseract.js unpacks them itself, so they are not
 * sent with a Content-Encoding a browser would undo. Their path names the data's version, so they are
 * cached for a year.
 */
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { gunzipSync } from 'node:zlib';

const require = createRequire(import.meta.url);
const DATA_FOLDER = '4.0.0_best_int';
const SITE_ORIGINS = ['https://officeparser.harshankur.com', 'http://localhost:3477'];
const PUBLIC_URL = 'https://tessdata.harshankur.com';

const [command, ...rest] = process.argv.slice(2);
const option = (name) => {
    const at = rest.indexOf(`--${name}`);
    return at >= 0 ? rest[at + 1] : undefined;
};
const out = path.resolve(option('out') ?? '.tessdata');
const bucket = option('bucket');

/**
 * Every language with data on npm: the @tesseract.js-data scope (searched page by page until a page
 * holds none of it), with the installed tesseract.js's own list in case the search misses one.
 */
async function languages() {
    const asked = option('langs');
    if (asked) return asked.split(',').map(l => l.trim()).filter(Boolean);
    const codes = new Set(Object.values(require('tesseract.js/src/constants/languages.js')));
    for (let from = 0; ; from += 250) {
        const page = await (await fetch(`https://registry.npmjs.org/-/v1/search?text=%40tesseract.js-data&size=250&from=${from}`)).json();
        const scoped = (page.objects ?? []).map(o => o.package.name).filter(n => n.startsWith('@tesseract.js-data/'));
        for (const name of scoped) codes.add(name.slice('@tesseract.js-data/'.length));
        if (!scoped.length) break;
    }
    return [...codes].sort();
}

/** The bytes of one file in a tar archive, by its path, or null. One forward pass over the headers. */
function tarMember(tar, wanted) {
    for (let at = 0; at + 512 <= tar.length;) {
        const header = tar.subarray(at, at + 512);
        if (header.every(b => b === 0)) break;
        const field = (from, to) => header.subarray(from, to).toString('latin1').replace(/\0.*$/s, '');
        const name = field(0, 100);
        const prefix = field(345, 500);
        const size = parseInt(field(124, 136).trim() || '0', 8);
        const full = prefix ? `${prefix}/${name}` : name;
        const start = at + 512;
        if (full === wanted) return tar.subarray(start, start + size);
        at = start + Math.ceil(size / 512) * 512;
    }
    return null;
}

async function download() {
    const folder = path.join(out, DATA_FOLDER);
    fs.mkdirSync(folder, { recursive: true });
    const failed = [];
    const legacyOnly = [];
    for (const lang of await languages()) {
        const target = path.join(folder, `${lang}.traineddata.gz`);
        if (fs.existsSync(target)) { console.log(`  - ${lang}: already downloaded`); continue; }
        try {
            const meta = await (await fetch(`https://registry.npmjs.org/@tesseract.js-data%2f${encodeURIComponent(lang)}`)).json();
            const version = meta['dist-tags']?.latest;
            const tarball = version && meta.versions?.[version]?.dist?.tarball;
            if (!tarball) throw new Error('no package on npm');
            const tgz = Buffer.from(await (await fetch(tarball)).arrayBuffer());
            const tar = gunzipSync(tgz);
            const data = tarMember(tar, `package/${DATA_FOLDER}/${lang}.traineddata.gz`);
            if (!data && tarMember(tar, `package/4.0.0/${lang}.traineddata.gz`)) {
                // Data for Tesseract's legacy engine only (kur, tgl): officeParser runs the LSTM engine, which
                // cannot read it from any host, so there is nothing to serve.
                legacyOnly.push(lang);
                console.log(`  - ${lang}: legacy-engine data only, none for the LSTM engine`);
                continue;
            }
            if (!data || data[0] !== 0x1f || data[1] !== 0x8b) throw new Error(`no gzip ${DATA_FOLDER}/${lang}.traineddata.gz in ${tarball}`);
            fs.writeFileSync(target, data);
            console.log(`  ✓ ${lang}: ${(data.length / 1048576).toFixed(1)} MB (@tesseract.js-data/${lang}@${version})`);
        } catch (e) {
            failed.push(lang);
            console.error(`  ✗ ${lang}: ${e.message}`);
        }
    }
    // The data's licence, served beside it: Tesseract's trained data is Apache-2.0 (the npm packages ship
    // no licence file), the text tesseract.js-core carries.
    fs.copyFileSync(require.resolve('tesseract.js-core/LICENSE'), path.join(folder, 'LICENSE'));
    console.log(`\nIn ${folder}.${legacyOnly.length ? ` No LSTM data: ${legacyOnly.join(', ')}.` : ''}${failed.length ? ` Not downloaded: ${failed.join(', ')}.` : ''}`);
    if (failed.length) process.exitCode = 1;
}

function wrangler(args) {
    const run = spawnSync('npx', ['wrangler', ...args], { stdio: 'inherit' });
    if (run.status !== 0) throw new Error(`wrangler ${args.slice(0, 3).join(' ')} failed`);
}

function needBucket() {
    if (!bucket) { console.error('Name the bucket: --bucket <name>'); process.exit(1); }
}

function cors() {
    needBucket();
    const file = path.join(out, 'cors.json');
    fs.mkdirSync(out, { recursive: true });
    fs.writeFileSync(file, JSON.stringify({ rules: [{ allowed: { origins: SITE_ORIGINS, methods: ['GET', 'HEAD'] } }] }, null, 2));
    wrangler(['r2', 'bucket', 'cors', 'set', bucket, '--file', file, '--force']);
}

async function upload() {
    needBucket();
    const folder = path.join(out, DATA_FOLDER);
    const files = fs.existsSync(folder) ? fs.readdirSync(folder).filter(f => f.endsWith('.traineddata.gz') || f === 'LICENSE') : [];
    if (!files.length) { console.error(`Nothing in ${folder}: run download first.`); process.exit(1); }
    let put = 0;
    for (const file of files) {
        // A file the bucket already serves at this size is skipped, so an interrupted upload resumes. Asked
        // under a query of its own: a 404 for the file's own address would stay in Cloudflare's cache for
        // minutes after the upload.
        const local = fs.statSync(path.join(folder, file)).size;
        const served = await fetch(`${PUBLIC_URL}/${DATA_FOLDER}/${file}?check=${Date.now()}`, { method: 'HEAD' }).catch(() => null);
        if (served?.ok && Number(served.headers.get('content-length')) === local) { console.log(`  - ${file}: already there`); continue; }
        const args = ['r2', 'object', 'put', `${bucket}/${DATA_FOLDER}/${file}`, '--file', path.join(folder, file),
            '--content-type', file === 'LICENSE' ? 'text/plain; charset=utf-8' : 'application/gzip',
            '--cache-control', 'public, max-age=31536000, immutable', '--remote'];
        try { wrangler(args); } catch { wrangler(args); } // once more, for a dropped connection
        put++;
    }
    console.log(`\nUploaded ${put} files to ${bucket}/${DATA_FOLDER}/; ${files.length - put} were there already.`);
}

const commands = { download, cors, upload };
if (!commands[command]) {
    console.error('Usage: node scripts/upload-tessdata.mjs download|cors|upload [--bucket <name>] [--langs eng,deu] [--out <dir>]');
    process.exit(1);
}
await commands[command]();
