import { Zippable, zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';

const VOID_TAGS = ['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'];

/**
 * Converts HtmlGenerator's HTML output into well-formed XHTML, which EPUB reading
 * systems parse as strict XML (unlike browsers, which tolerate HTML's looseness).
 *
 * This is more than cosmetic: a single raw `&` or unclosed tag makes the whole content
 * document fail to open. The conversion:
 *  - strips `<style>`/`<script>` blocks — browser page-chrome CSS and chart-init JS that
 *    a reading system ignores anyway, and whose CSS comments / JS operators contain raw
 *    `&` and `<` that are illegal as XML character data;
 *  - normalises HTML named entities (`&nbsp;`) to numeric references, since XML predefines
 *    only `&amp;`/`&lt;`/`&gt;`/`&quot;`/`&apos;`;
 *  - escapes stray ampersands (e.g. in `href` query strings) not already part of a valid
 *    reference;
 *  - gives HTML boolean attributes an explicit value (`checked` -> `checked="checked"`);
 *  - self-closes void elements (`<br>` -> `<br/>`).
 */
const toXhtml = (html: string): string => {
    let out = html
        .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
        .replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');

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

/**
 * Generates a minimal, valid EPUB 3 file from an AST.
 *
 * Every AST node is rendered as a single XHTML content document (reusing `HtmlGenerator`
 * for the actual markup, since EPUB content documents are XHTML) and packaged with the
 * required `mimetype`, `META-INF/container.xml`, OPF manifest, and navigation document.
 * Images are already embedded as base64 data URIs by `HtmlGenerator`, so the content
 * document is fully self-contained without a separate image manifest.
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
        const bodyHtml = typeof htmlResult.value === 'string' ? htmlResult.value : '';
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
    <item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>
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
        };

        return {
            value: zipSync(zipFiles),
            messages: this.messages,
        };
    }
}
