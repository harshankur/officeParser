import { Zippable, zipSync } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeParserAST } from '../types.js';
import { BaseGenerator } from './BaseGenerator.js';
import { HtmlGenerator } from './HtmlGenerator.js';

const VOID_TAGS = ['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr'];

/**
 * Self-closes HTML5 void elements (`<br>` -> `<br/>`) so HtmlGenerator's output becomes
 * well-formed XHTML, which EPUB 3 content documents require.
 */
const toXhtml = (html: string): string => {
    const voidTagPattern = new RegExp(`<(${VOID_TAGS.join('|')})((?:\\s+[^>]*)?)\\s*\\/?>`, 'gi');
    return html.replace(voidTagPattern, (_m, tag, attrs) => `<${tag}${attrs}/>`);
};

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
