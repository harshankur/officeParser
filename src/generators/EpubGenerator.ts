import { Zippable, zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';

const VOID_TAGS = ['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'];

/** Block-level tags whose HTML5 content model does not permit them inside a <p>. */
const BLOCK_TAGS_INVALID_IN_P = ['div', 'table', 'ul', 'ol', 'dl', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'section', 'figure'];

/**
 * HTML5's content model forbids block elements inside <p> (a <p> can only hold "phrasing
 * content"), but browsers silently fix this via the HTML5 parsing algorithm's
 * auto-closing rule: seeing a block start tag implicitly closes the open <p> first.
 * XML parsers have no such rule - they just build the tree exactly as written. Since
 * HtmlGenerator always wraps images in `<div class="image-container">`, a paragraph
 * whose only child is an image becomes `<p><div>...</div></p>`: well-formed XML, but
 * many EPUB rendering engines refuse to lay out a block box found inside a paragraph and
 * simply drop it - silently, with no parse error, which is why the image vanishes.
 *
 * Fixes this by promoting any `<p ...>` that contains a nested block tag to a `<div ...>`
 * instead, matching what a browser's auto-correction effectively produces. Paragraphs
 * don't nest, so the first `</p>` after each `<p>` is always its match.
 */
const promoteParagraphsWithBlockContent = (html: string): string => {
    const blockTagPattern = new RegExp(`<(?:${BLOCK_TAGS_INVALID_IN_P.join('|')})\\b`, 'i');
    let result = '';
    let cursor = 0;
    const pOpenRegex = /<p(\s[^>]*)?>/gi;
    let match: RegExpExecArray | null;

    while ((match = pOpenRegex.exec(html)) !== null) {
        if (match.index < cursor) continue; // inside content already emitted by a prior promotion

        result += html.slice(cursor, match.index);
        const contentStart = match.index + match[0].length;
        const closeMatch = /<\/p>/i.exec(html.slice(contentStart));
        if (!closeMatch) {
            // No closing tag found (shouldn't happen with well-formed generator output) -
            // leave as-is rather than risk corrupting the rest of the document.
            result += match[0];
            cursor = contentStart;
            pOpenRegex.lastIndex = cursor;
            continue;
        }

        const inner = html.slice(contentStart, contentStart + closeMatch.index);
        const attrs = match[1] || '';
        result += blockTagPattern.test(inner) ? `<div${attrs}>${inner}</div>` : `<p${attrs}>${inner}</p>`;

        cursor = contentStart + closeMatch.index + closeMatch[0].length;
        pOpenRegex.lastIndex = cursor;
    }
    result += html.slice(cursor);
    return result;
};

/**
 * Converts HtmlGenerator's HTML output into well-formed XHTML, which EPUB reading
 * systems parse as strict XML (unlike browsers, which tolerate HTML's looseness).
 *
 * This is more than cosmetic: a single raw `&` or unclosed tag makes the whole content
 * document fail to open. The conversion:
 *  - strips `<script>` blocks — EpubGenerator renders through HtmlGenerator with
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

const escapeXml = (text: string): string => text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

/** Maps an image MIME type to a file extension for the packaged resource. */
const MIME_EXT: Record<string, string> = {
    'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/png': 'png', 'image/gif': 'gif',
    'image/svg+xml': 'svg', 'image/webp': 'webp', 'image/bmp': 'bmp', 'image/tiff': 'tiff'
};

/** Decodes a base64 string to raw bytes, cross-env (atob exists in Node 16+ and browsers). */
const decodeBase64 = (b64: string): Uint8Array => {
    const bin = atob(b64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
};

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

    async generate(): Promise<ConversionResult<'epub'>> {
        const htmlGenerator = new HtmlGenerator(this.ast, {
            ...this.config,
            htmlConfig: { ...this.config.htmlConfig, standalone: false },
        } as GeneratorConfig<'html'>);
        const htmlResult = await htmlGenerator.generate();
        let bodyHtml = typeof htmlResult.value === 'string' ? htmlResult.value : '';

        // Extract base64 data-URI images into packaged files (EPUB readers don't render
        // `data:` URIs). Each distinct image becomes one OEBPS/images/imageN.ext resource,
        // a manifest <item>, and a rewritten relative `src`. Deduped so a repeated image
        // is packaged once.
        const imageResources: Record<string, Uint8Array> = {};
        const imageManifestItems: string[] = [];
        const dataUriToHref = new Map<string, string>();
        let imageCounter = 0;
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

        const title = this.ast.metadata?.title || 'Untitled';
        const author = this.ast.metadata?.author;
        const description = this.ast.metadata?.description;
        const nativeProps = (this.ast.metadata?.nativeProperties || {}) as Record<string, any>;
        const language = nativeProps.language || 'en';
        const identifier = nativeProps.identifier || `urn:x-officeparser:${this.slugify(title)}-${xhtmlBody.length}`;
        const modified = new Date().toISOString().replace(/\.\d+Z$/, 'Z');

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

        const opf = `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="pub-id">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:identifier id="pub-id">${escapeXml(identifier)}</dc:identifier>
    <dc:title>${escapeXml(title)}</dc:title>
    ${author ? `<dc:creator>${escapeXml(author)}</dc:creator>` : ''}
    ${description ? `<dc:description>${escapeXml(description)}</dc:description>` : ''}
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
            value: zipSync(zipFiles),
            messages: this.messages,
        };
    }
}
