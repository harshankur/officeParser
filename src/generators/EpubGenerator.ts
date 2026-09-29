import { Zippable, zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';
import { escapeXml } from '../utils/sanitize.js';
import { decodeBase64, documentLanguage, MIME_EXT, resolveZipInstant } from '../utils/officeGenUtils.js';

const VOID_TAGS = ['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'];

/** Block-level tags whose HTML5 content model does not permit them inside a <p> (a <p> among them). */
const BLOCK_TAGS_INVALID_IN_P = new Set(['address', 'article', 'aside', 'blockquote', 'details', 'dialog', 'div', 'dl', 'fieldset',
    'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'main', 'menu', 'nav', 'ol',
    'p', 'pre', 'section', 'table', 'ul']);

/**
 * HTML5's content model forbids block elements inside <p> (a <p> can only hold "phrasing
 * content"), but browsers silently fix this via the HTML5 parsing algorithm's
 * auto-closing rule: seeing a block start tag implicitly closes the open <p> first.
 * XML parsers have no such rule - they just build the tree exactly as written, and many EPUB
 * rendering engines refuse to lay out a block box found inside a paragraph and simply drop it.
 * HtmlGenerator writes a paragraph holding a block as its parts, so this is for markup it is
 * handed whole (an `onNode` replacement, a heading's content).
 *
 * Fixes this by promoting any `<p ...>` that contains a nested block tag to a `<div ...>`
 * instead, matching what a browser's auto-correction effectively produces. The tags are read
 * once, in order, each `</p>` matched to the `<p>` it closes: paired with the first `</p>` after
 * it, a paragraph holding a table with a <p> in a cell was closed at the cell's `</p>` and the
 * document was no longer well-formed. Comments are passed over (their text is not markup).
 */
const promoteParagraphsWithBlockContent = (html: string): string => {
    const open: { at: number; hasBlock: boolean }[] = [];
    const edits: { at: number; length: number; text: string }[] = [];
    for (let i = html.indexOf('<'); i !== -1; i = html.indexOf('<', i + 1)) {
        if (html.startsWith('<!--', i)) {
            const end = html.indexOf('-->', i + 4);
            if (end === -1) break;
            i = end + 2;
            continue;
        }
        const closing = html[i + 1] === '/';
        let nameEnd = i + (closing ? 2 : 1);
        while (nameEnd < html.length && /[A-Za-z0-9]/.test(html[nameEnd])) nameEnd++;
        const name = html.slice(i + (closing ? 2 : 1), nameEnd).toLowerCase();
        if (closing) {
            if (name !== 'p') continue;
            const paragraph = open.pop();
            if (paragraph?.hasBlock) edits.push({ at: paragraph.at, length: 2, text: '<div' }, { at: i, length: 3, text: '</div' });
            continue;
        }
        // The innermost open paragraph holds this block; a paragraph holding it holds a block too.
        if (BLOCK_TAGS_INVALID_IN_P.has(name) && open.length) open[open.length - 1].hasBlock = true;
        if (name === 'p') open.push({ at: i, hasBlock: false });
    }
    if (!edits.length) return html;
    edits.sort((a, b) => a.at - b.at);
    const parts: string[] = [];
    let cursor = 0;
    for (const edit of edits) {
        parts.push(html.slice(cursor, edit.at), edit.text);
        cursor = edit.at + edit.length;
    }
    parts.push(html.slice(cursor));
    return parts.join('');
};

/**
 * Converts HtmlGenerator's HTML output into well-formed XHTML, which EPUB reading
 * systems parse as strict XML (unlike browsers, which tolerate HTML's looseness).
 *
 * This is more than cosmetic: a single raw `&` or unclosed tag makes the whole content
 * document fail to open. The conversion:
 *  - strips `<script>` blocks. EpubGenerator renders through HtmlGenerator with
 *    `standalone: false`, which already omits the envelope-level stylesheet and Chart.js/
 *    spreadsheet scripts entirely, but a chart *node* still emits its own inline
 *    `<script>` (chart-init JS) regardless of that flag, since it's content, not envelope.
 *    A reading system can't execute it anyway, and its JS operators can contain raw `&`/`<`
 *    that are illegal as XML character data, so it's stripped here;
 *  - promotes `<p>` tags that contain nested block content (see
 *    promoteParagraphsWithBlockContent above) to `<div>`, since XML readers don't apply
 *    HTML5's auto-closing correction that hides this in a browser;
 *  - normalises HTML named entities (`&nbsp;`) to numeric references, since XML predefines
 *    only `&amp;`/`&lt;`/`&gt;`/`&quot;`/`&apos;`;
 *  - escapes stray ampersands (e.g. in `href` query strings) not already part of a valid
 *    reference;
 *  - gives HTML boolean attributes an explicit value (`checked` -> `checked="checked"`);
 *  - self-closes void elements (`<br>` -> `<br/>`).
 */
const toXhtml = (html: string): string => {
    let out = html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');

    out = promoteParagraphsWithBlockContent(out);

    // Named -> numeric entities (nbsp is the only named entity HtmlGenerator emits).
    out = out.replace(/&nbsp;/g, '&#160;');

    // Escape ampersands that don't already open a valid XML entity reference.
    out = out.replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g, '&amp;');

    // Give bare boolean attributes an explicit value. Scoped to the specific tags that
    // emit them (checkbox task-list items, media iframes) so body text like "the selected
    // option" is never rewritten.
    out = out.replace(/<input\b([^>]*?)\schecked(\s*\/?>)/gi, '<input$1 checked="checked"$2');
    out = out.replace(/<(iframe|video|audio)\b([^>]*?)\s(allowfullscreen|autoplay|controls|loop|muted)(\s*\/?>|\s)/gi, '<$1$2 $3="$3"$4');

    // Self-close void elements (non-greedy attr capture so an already-present trailing `/`
    // isn't duplicated, e.g. `<meta .../>` must not become `<meta ...//>`).
    const voidTagPattern = new RegExp(`<(${VOID_TAGS.join('|')})((?:\\s[^>]*?)?)\\s*/?>`, 'gi');
    out = out.replace(voidTagPattern, (_m, tag, attrs) => `<${tag}${attrs}/>`);

    return out;
};

/**
 * Minimal, book-friendly CSS injected into every EPUB content document. Kept static and
 * free of `&`/`<` so it is XML-safe inline; the reading system supplies typography, so
 * this only covers structural essentials the stripped page-chrome would otherwise lose.
 */
const EPUB_STYLESHEET = `img { max-width: 100%; height: auto; }
table { border-collapse: collapse; margin: 1em 0; }
td, th { border: 1px solid #ccc; padding: 4px 8px; }`;


/**
 * Generates a minimal, valid EPUB 3 file from an AST.
 *
 * Every AST node is rendered as a single XHTML content document (reusing `HtmlGenerator`
 * for the actual markup, since EPUB content documents are XHTML) and packaged with the
 * required `mimetype`, `META-INF/container.xml`, OPF manifest, and navigation document.
 *
 * `HtmlGenerator` embeds images as base64 `data:` URIs, but EPUB reading systems do not
 * render `data:` URIs - images must be packaged as separate resources referenced by a
 * relative path. So each data-URI image is extracted into `OEBPS/images/`, declared in
 * the manifest, and its `<img src>` rewritten to point at the packaged file.
 */
export class EpubGenerator extends BaseGenerator<'epub'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'epub'>) {
        super('epub', ast, config);
    }

    /**
     * Resolves the modification instant used for both the EPUB 3 `dcterms:modified` property
     * (which the specification requires) and every zip entry's mtime.
     *
     * Takes the value from `effectiveMetadata`, i.e. `metadataOverrides.modified` if the caller
     * set one, otherwise the source document's own `metadata.modified`, and only falls back to
     * the current time when neither exists. That last fallback is the sole non-reproducible
     * option, so it is the last resort rather than the default.
     *
     * **Both outputs matter for reproducibility.** `dcterms:modified` is the visible one, but
     * `zipSync` defaults each entry's mtime to `Date.now()`, so pinning only the OPF still
     * yields archives that differ byte-for-byte on every run. That second source is easy to
     * miss because DOS zip timestamps have two-second granularity - back-to-back generation
     * looks stable and only a gap longer than that reveals it.
     *
     * `iso` is `YYYY-MM-DDThh:mm:ssZ` (UTC, whole seconds) as EPUB requires; `toISOString()`
     * emits milliseconds, so they are stripped.
     */
    private resolveModified(): { iso: string; mtime: Date } {
        return resolveZipInstant(this.effectiveMetadata.modified);
    }

    async generate(): Promise<ConversionResult<'epub'>> {
        const htmlGenerator = new HtmlGenerator(this.ast, {
            ...this.config,
            // EPUB packages every image into a separate file, rewriting its data: URI to a relative src
            // (readers don't render data: URIs). That rewrite only fires on a data: URI, so the
            // maxInlineImageBytes cap must not apply here: a large image would otherwise degrade to a bare
            // `<img src="name">` the packager can't rewrite, giving a broken image + no manifest entry.
            maxInlineImageBytes: Infinity,
            // Force sourceAttributes off: those data-* attributes are wire-format plumbing for
            // structured consumers and change the mermaid shape's rendered appearance, neither of
            // which belongs in a packaged EPUB.
            htmlConfig: { ...this.config.htmlConfig, standalone: false, sourceAttributes: false },
        } as GeneratorConfig<'html'>);
        // Each attachment's picture is a file in the package, which every <img> showing it points at: its
        // data was inlined at every <img> and extracted again, so a small book showing one large
        // picture many times was built as an HTML string past what memory holds.
        const imageResources: Record<string, Uint8Array> = {};
        const imageManifestItems: string[] = [];
        const dataUriToHref = new Map<string, string>();
        const attachmentHref = new Map<string, string>();
        let imageCounter = 0;
        htmlGenerator.imageSourceFor = (attachment) => {
            const mime = (attachment.mimeType || 'image/png').toLowerCase();
            if (!/^image\/[a-z0-9.+-]+$/.test(mime) || !attachment.name) return undefined;
            let href = attachmentHref.get(attachment.name);
            if (!href) {
                let data: Uint8Array;
                try { data = decodeBase64(attachment.data); } catch { return undefined; }
                imageCounter++;
                href = `images/image${imageCounter}.${MIME_EXT[mime] || 'img'}`;
                imageResources[`OEBPS/${href}`] = data;
                imageManifestItems.push(`<item id="img${imageCounter}" href="${href}" media-type="${mime}"/>`);
                attachmentHref.set(attachment.name, href);
            }
            return href;
        };
        const htmlResult = await htmlGenerator.generate();
        let bodyHtml = typeof htmlResult.value === 'string' ? htmlResult.value : '';

        // Extract any other base64 data-URI images (a picture given by a data: URL) into packaged
        // files (EPUB readers don't render `data:` URIs). Each distinct image becomes one
        // OEBPS/images/imageN.ext resource, a manifest <item>, and a rewritten relative `src`.
        // Deduped so a repeated image is packaged once.
        bodyHtml = bodyHtml.replace(/(<img\b[^>]*\bsrc=")(data:(image\/[a-zA-Z0-9.+-]+);base64,([^"]+))(")/gi,
            (_full, pre, dataUri, mime, b64, post) => {
                let href = dataUriToHref.get(dataUri);
                if (!href) {
                    imageCounter++;
                    const ext = MIME_EXT[mime.toLowerCase()] || 'img';
                    href = `images/image${imageCounter}.${ext}`;
                    dataUriToHref.set(dataUri, href);
                    try {
                        imageResources[`OEBPS/${href}`] = decodeBase64(b64);
                        imageManifestItems.push(`<item id="img${imageCounter}" href="${href}" media-type="${mime}"/>`);
                    } catch {
                        // Undecodable data - leave the original src untouched rather than
                        // emit a manifest entry for a resource we couldn't write.
                        dataUriToHref.delete(dataUri);
                        return `${pre}${dataUri}${post}`;
                    }
                }
                return `${pre}${href}${post}`;
            });

        const xhtmlBody = toXhtml(bodyHtml);

        // Via effectiveMetadata so `metadataOverrides` applies here as it does everywhere else.
        const meta = this.effectiveMetadata;
        const title = meta.title || 'Untitled';
        const author = meta.author;
        const description = meta.description;
        // dc:subject is the OPF's only vocabulary slot for either of these; both are repeatable.
        const subject = meta.subject;
        const keywords = meta.keywords;
        const nativeProps = (meta.nativeProperties || {}) as Record<string, any>;
        const language = documentLanguage(meta);
        // OPF metadata is a closed Dublin Core vocabulary: a caller-defined key has no valid
        // element to live in, and inventing one risks failing EPUB validation outright.
        this.warnUnrepresentableCustomMetadata('EPUB');
        const identifier = nativeProps.identifier || `urn:x-officeparser:${this.slugify(title)}-${xhtmlBody.length}`;
        const { iso: modified, mtime } = this.resolveModified();

        const chapterXhtml = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="${escapeXml(language)}">
<head>
<meta charset="utf-8"/>
<title>${escapeXml(title)}</title>
<style type="text/css">
${EPUB_STYLESHEET}
</style>
</head>
<body>
${xhtmlBody}
</body>
</html>`;

        // Built as a list rather than inline ternaries so an absent field contributes nothing at
        // all. Inline `${x ? ... : ''}` leaves the surrounding indentation and newline behind, so
        // every optional field the document lacks used to emit a stray blank line into the OPF.
        const optionalDcElements = [
            author ? `<dc:creator>${escapeXml(author)}</dc:creator>` : '',
            description ? `<dc:description>${escapeXml(description)}</dc:description>` : '',
            // dc:subject is repeatable and is the OPF's only slot for either of these.
            subject ? `<dc:subject>${escapeXml(subject)}</dc:subject>` : '',
            keywords ? `<dc:subject>${escapeXml(keywords)}</dc:subject>` : '',
        ].filter(Boolean).map(el => `    ${el}`).join('\n');

        const opf = `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="pub-id">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="pub-id">${escapeXml(identifier)}</dc:identifier>
    <dc:title>${escapeXml(title)}</dc:title>
${optionalDcElements}
    <dc:language>${escapeXml(language)}</dc:language>
    <meta property="dcterms:modified">${modified}</meta>
  </metadata>
  <manifest>
    <item id="chapter1" href="chapter1.xhtml" media-type="application/xhtml+xml"/>
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>${imageManifestItems.length ? '\n    ' + imageManifestItems.join('\n    ') : ''}
  </manifest>
  <spine>
    <itemref idref="chapter1"/>
  </spine>
</package>`;

        const navXhtml = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><meta charset="utf-8"/><title>Navigation</title></head>
<body>
<nav epub:type="toc" id="toc">
<h1>${escapeXml(title)}</h1>
<ol>
<li><a href="chapter1.xhtml">${escapeXml(title)}</a></li>
</ol>
</nav>
</body>
</html>`;

        const containerXml = `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`;

        const encoder = new TextEncoder();
        // EPUB requires the mimetype entry to be the first file in the archive, stored
        // uncompressed (level 0) - readers use it to sniff the format before parsing any XML.
        const zipFiles: Zippable = {
            mimetype: [encoder.encode('application/epub+zip'), { level: 0 }],
            'META-INF/container.xml': encoder.encode(containerXml),
            'OEBPS/content.opf': encoder.encode(opf),
            'OEBPS/nav.xhtml': encoder.encode(navXhtml),
            'OEBPS/chapter1.xhtml': encoder.encode(chapterXhtml),
            ...imageResources,
        };

        return {
            // An explicit mtime is what makes the archive reproducible; fflate would otherwise
            // stamp every entry with Date.now(). See resolveModified().
            value: zipSync(zipFiles, { mtime }),
            messages: this.messages,
        };
    }
}
