/**
 * Security regression tests for output sanitization.
 *
 * Every string in the AST is treated as attacker-controlled (it comes from an
 * untrusted document). These tests lock in that document-supplied values can't
 * break out of their destination context (HTML attribute/tag, inline script,
 * CSS, a CSV formula, a Markdown link, an RTF group) in the generated output.
 */
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { strFromU8, unzipSync, zipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import { OfficeParserAST, OfficeWarningType } from '../../src/types';
import { resolveGeneratorConfig, resolveParserConfig } from '../../src/utils/configUtils';
import {
    escapeHtml, escapeXml, sanitizeCssValue, sanitizeUrl, sanitizeImageUrl,
    serializeForInlineScript, csvSafeCell, escapeRtf, markdownEscapeText, sanitizeMarkdownUrl, sanitizeRtfUrl
} from '../../src/utils/sanitize';
import { extractFiles } from '../../src/utils/zipUtils';
import { getOfficeError, getWrappedError } from '../../src/utils/errorUtils';
import { OfficeErrorType } from '../../src/types';

let passed = 0;
let failed = 0;
const check = (name: string, cond: boolean, detail = '') => {
    if (cond) { passed++; }
    else { failed++; console.error(`  ✗ FAIL: ${name}${detail ? ` — ${detail}` : ''}`); }
};

function astWith(content: any[]): OfficeParserAST {
    return {
        type: 'docx',
        metadata: { title: 'Security Test' },
        attachments: [],
        content,
        toText: () => '',
        getImages: () => []
    } as any;
}

function unitTests() {
    console.log('- Sanitize module (unit)...');

    // escapeHtml / escapeXml include the single quote.
    check('escapeHtml quotes', escapeHtml(`a<b>&"'`) === 'a&lt;b&gt;&amp;&quot;&#39;');
    check('escapeXml apos', escapeXml(`'`) === '&apos;');

    // CSS value sanitizer: drop breakout / resource-fetching constructs, keep colors.
    check('css tag breakout dropped', sanitizeCssValue('red"><script>') === '');
    check('css url() dropped', sanitizeCssValue('url(javascript:alert(1))') === '');
    check('css expression dropped', sanitizeCssValue('expression(alert(1))') === '');
    check('css semicolon stripped', !sanitizeCssValue('red;background:blue').includes(';'));
    check('css rgb preserved', sanitizeCssValue('rgb(255,0,0)') === 'rgb(255,0,0)');
    // Obfuscated url() must not reassemble once control chars / comments are stripped.
    check('css newline-obfuscated url dropped', !/url\s*\(/i.test(sanitizeCssValue('u\nrl(http://evil)')));
    check('css comment-obfuscated url dropped', !/url\s*\(/i.test(sanitizeCssValue('url/*x*/(http://evil)')));
    // CSS backslash escapes are resolved away by the browser, so `u\rl(` IS `url(` to a
    // renderer. These were the gap: the strip ran downstream of the denylist test, so the
    // sanitizer returned a live url() it had just declared safe.
    check('css escape-obfuscated url dropped', !/url\s*\(/i.test(sanitizeCssValue('u\\rl(http://evil/x)')));
    check('css escape-obfuscated expression dropped', !/expression\s*\(/i.test(sanitizeCssValue('expr\\ession(alert(1))')));
    check('css escape-obfuscated image-set dropped', !/image-set\s*\(/i.test(sanitizeCssValue('image\\-set(x)')));
    // Contract-level, not payload-level: every denylisted construct must stay dropped under an
    // escaped spelling. This is what catches the next variant rather than the last one.
    for (const construct of ['url', 'expression', 'image-set', 'element', '-moz-binding']) {
        const escaped = construct[0] + '\\' + construct.slice(1) + '(http://evil/x)';
        check(`css escaped "${construct}(" dropped`, sanitizeCssValue(escaped) === '',
            `sanitizeCssValue(${JSON.stringify(escaped)}) = ${JSON.stringify(sanitizeCssValue(escaped))}`);
    }
    // A legitimate value that merely contains a backslash still survives (minus the backslash).
    check('css plain value survives escape strip', sanitizeCssValue('12\\px') === '12px');

    // Formula guard must not be bypassable by leading whitespace.
    check('csv leading-space formula guarded', csvSafeCell(' =1+1', ',').includes(`'`));
    check('csv leading-space at guarded', csvSafeCell('  @SUM(1)', ',').includes(`'`));

    // URL sanitizer: block script schemes (incl. control-char obfuscation), keep http/relative.
    check('url javascript blocked', sanitizeUrl('javascript:alert(1)') === '');
    check('url obfuscated blocked', sanitizeUrl('java\tscript:alert(1)') === '');
    check('url vbscript blocked', sanitizeUrl('vbscript:msgbox(1)') === '');
    check('url data blocked (link)', sanitizeUrl('data:text/html,<script>') === '');
    check('url https allowed', sanitizeUrl('https://example.com/a?b=1') === 'https://example.com/a?b=1');
    check('url fragment allowed', sanitizeUrl('#section') === '#section');

    // Image URL sanitizer additionally allows data:image, still blocks scripts.
    check('img data:image allowed', sanitizeImageUrl('data:image/png;base64,AAAA') === 'data:image/png;base64,AAAA');
    check('img data:text/html blocked', sanitizeImageUrl('data:text/html,<script>') === '');
    check('img javascript blocked', sanitizeImageUrl('javascript:alert(1)') === '');

    // Inline-script serializer escapes the </script> sequence.
    check('inline script escapes <', !serializeForInlineScript({ x: '</script>' }).includes('</script>'));
    check('inline script has \\u003C', serializeForInlineScript({ x: '</script>' }).includes('\\u003C'));

    // CSV formula/DDE guard.
    check('csv = guarded', csvSafeCell('=1+1', ',').startsWith(`"'=`) || csvSafeCell('=1+1', ',') === `'=1+1`);
    check('csv @ guarded', csvSafeCell('@SUM(1)', ',').startsWith(`'@`));
    check('csv + formula guarded', csvSafeCell('+1+1', ',').startsWith(`'+`));
    check('csv signed number preserved', csvSafeCell('+7', ',') === '+7');
    check('csv negative number preserved', csvSafeCell('-5.3', ',') === '-5.3');
    check('csv plain preserved', csvSafeCell('hello', ',') === 'hello');
    check('csv delimiter quoted', csvSafeCell('a,b', ',') === '"a,b"');

    // RTF control-char / quote escaping.
    check('rtf braces escaped', escapeRtf('a{b}\\c') === 'a\\{b\\}\\\\c');
    check('rtf quote escaped', escapeRtf('"') === "\\'22");
    check('rtf url javascript blocked', sanitizeRtfUrl('javascript:alert(1)') === '');
    check('rtf url file blocked', sanitizeRtfUrl('file:///etc/passwd') === '');
    check('rtf url UNC blocked', sanitizeRtfUrl('\\\\host\\share') === '');
    check('rtf url https allowed', sanitizeRtfUrl('https://example.com/a') === 'https://example.com/a');
    check('rtf url relative allowed', sanitizeRtfUrl('a/b.html') === 'a/b.html');

    // Markdown: only tag-opening "<" is encoded; bare "<" is preserved for round-trip.
    check('md tag < encoded', markdownEscapeText('<img onerror=x>') === '&lt;img onerror=x>');
    check('md bare < preserved', markdownEscapeText('a < b') === 'a < b');
    check('md url javascript blocked', sanitizeMarkdownUrl('javascript:alert(1)') === '');
    check('md url paren encoded', sanitizeMarkdownUrl('http://x/a(b)').includes('%28'));
    check('md img data allowed', sanitizeMarkdownUrl('data:image/png;base64,AA', { allowDataImage: true }) === 'data:image/png;base64,AA');
}

async function htmlTests() {
    console.log('- HtmlGenerator (integration)...');
    const XSS = 'red"><script>alert(1)</script>';

    const styleAst = astWith([
        { type: 'paragraph', children: [
            { type: 'text', text: 'hi', formatting: { color: XSS } }
        ] }
    ]);
    const html = (await OfficeGenerator.generate(styleAst, 'html', { includeFormatting: true })).value as string;
    check('html: color XSS not raw', !html.includes('<script>alert(1)'), 'style breakout survived');

    const anchorAst = astWith([
        { type: 'paragraph', metadata: { anchorIds: ['x"><script>alert(2)</script>'] }, children: [
            { type: 'text', text: 'hi' }
        ] } as any
    ]);
    const html2 = (await OfficeGenerator.generate(anchorAst, 'html')).value as string;
    check('html: anchorId XSS not raw', !html2.includes('<script>alert(2)'), 'id/name breakout survived');

    // Image width flows into a style="" attribute — it must be CSS-sanitized so it can't
    // break out with a quote (event-handler injection) or smuggle a url() resource fetch.
    // `url` is the real ImageMetadata field; `src` is not one, so an AST using it renders
    // src="" and exercises far less of the path than it appears to.
    const imgAst = astWith([
        { type: 'image', text: 'alt', metadata: { width: '1px" onerror="alert(4)', url: 'data:image/png;base64,AAAA' } } as any
    ]);
    const imgHtml = (await OfficeGenerator.generate(imgAst, 'html', { includeFormatting: true })).value as string;
    // The escaped data-width attribute legitimately echoes the text; the breakout signature
    // is a REAL `onerror="` attribute (quote closed the style early), which must be absent.
    check('html: image width no attr breakout', !/onerror\s*=\s*"/.test(imgHtml), `width broke out: ${imgHtml}`);
    // A width the sanitizer fully rejects emits NO style attribute at all, so asserting
    // "no url() in the style" against it matches nothing and passes vacuously - that is exactly
    // how this test sat green while the escape-obfuscation bypass went unnoticed. Use a value
    // with a legitimate leading length so a style attribute is genuinely produced, assert it
    // rendered, and only then assert the payload did not survive inside it.
    const imgUrlAst = astWith([
        { type: 'image', text: 'alt', metadata: { width: '50px', url: 'data:image/png;base64,AAAA' } } as any
    ]);
    const imgUrlHtml = (await OfficeGenerator.generate(imgUrlAst, 'html', { includeFormatting: true })).value as string;
    const imgStyle = imgUrlHtml.match(/\sstyle="([^"]*)"/)?.[1] || '';
    check('html: image style attribute is actually emitted', imgStyle.length > 0,
        `no style attribute, so the url() check below would be vacuous: ${imgUrlHtml}`);
    check('html: image width no url() fetch', !/url\(/i.test(imgStyle), `width injected url() into style: ${imgStyle}`);
    // And the hostile widths - both plain and escape-obfuscated - must yield no style at all.
    for (const hostile of ['1px;background:url(http://evil/x)', '1px;background:u\\rl(http://evil/x)']) {
        const ast = astWith([{ type: 'image', text: 'alt', metadata: { width: hostile, url: 'data:image/png;base64,AAAA' } } as any]);
        const out = (await OfficeGenerator.generate(ast, 'html', { includeFormatting: true })).value as string;
        const style = out.match(/\sstyle="([^"]*)"/)?.[1] || '';
        check(`html: hostile width ${JSON.stringify(hostile)} emits no url()`, !/url\(/i.test(style),
            `style="${style}"`);
    }

    const linkAst = astWith([
        { type: 'paragraph', children: [
            { type: 'text', text: 'click', metadata: { link: 'javascript:alert(3)', linkType: 'external' } }
        ] } as any
    ]);
    const html3 = (await OfficeGenerator.generate(linkAst, 'html')).value as string;
    check('html: javascript link neutralized', !html3.includes('href="javascript:'), 'javascript href survived');
}

async function markdownTests() {
    console.log('- MarkdownGenerator (integration)...');

    const scriptAst = astWith([
        { type: 'paragraph', children: [
            { type: 'text', text: '<script>alert(1)</script>' }
        ] }
    ]);
    const md = (await OfficeGenerator.generate(scriptAst, 'md')).value as string;
    check('md: raw script tag encoded', !md.includes('<script>'), 'raw <script> survived to markdown');

    const linkAst = astWith([
        { type: 'paragraph', children: [
            { type: 'text', text: 'x', metadata: { link: 'javascript:alert(1)', linkType: 'external' } }
        ] } as any
    ]);
    const md2 = (await OfficeGenerator.generate(linkAst, 'md')).value as string;
    check('md: javascript link dropped', !md2.includes('javascript:'), 'javascript: survived markdown link');

    // --- Sinks that emitted document text without escaping ---------------------------------
    // Text nodes were escaped, but seven other constructs interpolated their content directly.
    // Each has its own delimiter, so each needs its own treatment - which is why these are
    // asserted individually rather than through one shared helper.
    //
    // The payload uses `/` as the attribute separator on purpose: a whitespace-stripping guard
    // (which is what the attribute-list sink had) stops `<img src=x onerror=…>` but not this.
    const PAYLOAD = '<img/src=x/onerror=alert(1)>';

    // Document-reachable sinks: driven through the real parser from real Markdown source, not a
    // hand-built AST, so the test proves the whole parse -> generate path and not just the
    // generator half.
    const viaDocument = async (source: string): Promise<string> => {
        const tmp = path.join(os.tmpdir(), `op-sec-${Date.now()}-${Math.random().toString(36).slice(2)}.md`);
        fs.writeFileSync(tmp, source);
        try {
            const ast = await OfficeParser.parseOffice(tmp, {} as any);
            return String((await ast.to('md')).value);
        } finally { fs.unlinkSync(tmp); }
    };

    const docSinks: Array<[string, string, RegExp]> = [
        // [name, source, a pattern proving the construct actually rendered]
        ['inline math', `Text $${PAYLOAD}$ end.`, /\$/],
        ['block math', `$$\n${PAYLOAD}\n$$`, /\$\$/],
        ['wikilink', `[[Page${PAYLOAD}]]`, /\[\[/],
        ['wikilink alias', `[[Page|Alias${PAYLOAD}]]`, /\[\[/],
        ['abbreviation', `*[HTML]: Hyper ${PAYLOAD} Lang\n\nThe HTML spec.`, /\*\[HTML\]:/],
        ['footnote key', `text[^${PAYLOAD}]\n\n[^${PAYLOAD}]: def`, /\[\^/],
    ];
    for (const [name, source, renderedPattern] of docSinks) {
        const out = await viaDocument(source);
        check(`md: ${name} actually rendered`, renderedPattern.test(out),
            `construct absent from output, so the escape check below would be vacuous: ${JSON.stringify(out.slice(0, 120))}`);
        check(`md: ${name} cannot carry a raw tag`, !out.includes(PAYLOAD),
            `payload survived verbatim: ${JSON.stringify(out.slice(0, 160))}`);
    }

    // The attribute list is the one document-reachable sink whose correct behaviour is to DROP
    // the value rather than encode it (it lands in metadata, which is not entity-decoded), so
    // it gets a positive control instead of a "still rendered" guard: a legitimate width must
    // survive, a hostile one must vanish entirely.
    const attrHostile = await viaDocument(`![a](x.png){width=50%${PAYLOAD}}`);
    check('md: attribute list cannot carry a raw tag', !attrHostile.includes(PAYLOAD),
        `payload survived: ${JSON.stringify(attrHostile.slice(0, 160))}`);
    const attrBenign = await viaDocument('![a](x.png){width=50%}');
    check('md: attribute list still emits a legitimate width', attrBenign.includes('{width=50%}'),
        `legitimate attribute list was dropped too: ${JSON.stringify(attrBenign.slice(0, 160))}`);

    // Sinks reachable only from a programmatic AST (both parsers allowlist admonitionType, and
    // no parser ever sets an admonition title or a non-conforming citation key). The generator
    // has to stand alone against these - it is a public API.
    for (const dialect of ['extended', 'gitlab', 'pandoc', 'commonmark']) {
        const admAst = astWith([{ type: 'admonition',
            metadata: { admonitionType: `note${PAYLOAD}`, title: `T${PAYLOAD}` },
            children: [{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }] } as any]);
        const out = (await OfficeGenerator.generate(admAst, 'md', { mdConfig: { dialect } } as any)).value as string;
        check(`md: admonition (${dialect}) actually rendered`, out.includes('body'),
            'admonition body absent, so the escape check below would be vacuous');
        check(`md: admonition (${dialect}) cannot carry a raw tag`, !out.includes(PAYLOAD),
            `payload survived: ${JSON.stringify(out.slice(0, 160))}`);
    }

    const citAst = astWith([{ type: 'paragraph', children: [
        { type: 'text', text: 'c', metadata: { citationKey: `k${PAYLOAD}` } }] } as any]);
    const citOut = (await OfficeGenerator.generate(citAst, 'md')).value as string;
    check('md: citation actually rendered', /\[@/.test(citOut),
        'no citation emitted, so the escape check below would be vacuous');
    check('md: citation key cannot carry a raw tag', !citOut.includes(PAYLOAD), citOut);

    // Under the commonmark preset math has NO delimiter at all - the text lands straight in the
    // document body, which makes it the worst case rather than an edge case.
    const mathAst = astWith([{ type: 'code', text: PAYLOAD, metadata: { math: 'inline' } } as any]);
    const mathOut = (await OfficeGenerator.generate(mathAst, 'md', { mdConfig: { dialect: 'commonmark' } } as any)).value as string;
    check('md: undelimited math cannot carry a raw tag', !mathOut.includes(PAYLOAD), mathOut);

    // Fidelity half: the escaping must not destroy legitimate content. `$a < b$` is the case
    // that rules out "just drop every <".
    const latex = await viaDocument('Given $a < b$ and $E = mc^2$ here.');
    check('md: legitimate LaTeX comparison survives', latex.includes('$a < b$'),
        `real math was corrupted: ${JSON.stringify(latex.slice(0, 160))}`);
}

async function csvTests() {
    console.log('- CsvGenerator (integration)...');
    const sheetAst = astWith([
        { type: 'sheet', metadata: { sheetName: 'S1' }, children: [
            { type: 'row', children: [
                { type: 'cell', children: [{ type: 'text', text: '=HYPERLINK("http://evil")' }] }
            ] }
        ] } as any
    ]);
    const csv = (await OfficeGenerator.generate(sheetAst, 'csv')).value as string;
    check('csv: formula cell guarded', !/(^|,|\n)=HYPERLINK/.test(csv), `formula not guarded: ${JSON.stringify(csv)}`);

    // A `#` comment line (sheet name / metadata) must not split into a formula cell:
    // the delimiter inside the value has to be neutralized.
    const commentAst = astWith([
        { type: 'sheet', metadata: { sheetName: 'good,=1+1' }, children: [
            { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'a' }] }] }
        ] },
        { type: 'sheet', metadata: { sheetName: 'S2' }, children: [
            { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'b' }] }] }
        ] },
    ] as any);
    (commentAst as any).metadata = { title: 'pwn,=cmd()' };
    const csv2 = (await OfficeGenerator.generate(commentAst, 'csv', { renderMetadata: true } as any)).value as string;
    const cellStartsFormula = csv2.split('\n').some(line => line.split(',').slice(1).some(c => /^[=+\-@]/.test(c)));
    check('csv: comment line no formula split', !cellStartsFormula, `comment split into formula: ${JSON.stringify(csv2)}`);
}

/**
 * `BaseContentNode.htmlAttributes` replays source attributes into generated HTML, so it is an
 * injection surface by construction. These build the AST directly rather than parsing, because
 * that is the path that bypasses the parser's own filtering - the generator has to stand alone.
 */
async function htmlAttributeBagTests() {
    console.log('- HtmlGenerator attribute bag (integration)...');
    const gen = async (htmlAttributes: Record<string, string>) =>
        (await OfficeGenerator.generate(
            astWith([{ type: 'paragraph', htmlAttributes, children: [{ type: 'text', text: 'x' }] }] as any),
            'html', { htmlConfig: { standalone: false } } as any
        )).value as string;

    const onclick = await gen({ onclick: 'alert(1)', onerror: 'alert(2)' });
    check('bag: event handlers dropped', !/onclick|onerror/i.test(onclick), onclick);

    const jsHref = await gen({ href: 'javascript:alert(1)' });
    check('bag: javascript: URL dropped', !/javascript:/i.test(jsHref), jsHref);

    const dataHtml = await gen({ src: 'data:text/html,<script>alert(1)</script>' });
    check('bag: data:text/html src dropped', !/data:text\/html/i.test(dataHtml), dataHtml);

    const srcdoc = await gen({ srcdoc: '<script>alert(1)</script>' });
    check('bag: srcdoc dropped', !/srcdoc/i.test(srcdoc), srcdoc);

    // A key carrying its own quote/`=` is the shape of an attribute-injection payload.
    const breakout = await gen({ 'x" onclick="alert(1)': 'y' });
    check('bag: attribute-injecting key dropped', !/onclick/i.test(breakout), breakout);

    const styleExpr = await gen({ style: 'width:expression(alert(1))' });
    check('bag: style never carried', !/expression\(/i.test(styleExpr), styleExpr);

    // Values are escaped, so a quote in a value cannot terminate the attribute early.
    const quoted = await gen({ 'data-note': 'he said "hi" <b>' });
    check('bag: value escaped', !/data-note="he said "/.test(quoted) && /&quot;|&#/.test(quoted), quoted);

    // Duplicate attributes are merely invalid in HTML but FATAL in the XHTML EpubGenerator emits -
    // an unopenable EPUB. Nothing else in the gate parses generated output as XML.
    const dupe = await gen({ class: 'from-source', 'data-k': 'v' });
    for (const tag of dupe.match(/<[a-zA-Z][^>]*>/g) || []) {
        const names = [...tag.matchAll(/\s([a-zA-Z_:][\w:.-]*)\s*=/g)].map(m => m[1].toLowerCase());
        check('bag: no duplicate attribute names', new Set(names).size === names.length, tag);
    }
}

/**
 * `metadataOverrides` is the first path where a caller supplies metadata *keys*, not just values.
 * Every prior metadata key came from a fixed vocabulary in our own code, so the key side was never
 * an injection surface; `custom` makes it one. Both halves need escaping in every destination.
 */
async function metadataOverrideTests() {
    console.log('- metadataOverrides (keys and values)...');

    const ast = astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'Body' }] }]);
    const hostileKey = 'x"><script>alert(1)</script><meta name="y';
    const hostileValue = '"><script>alert(2)</script>';

    // HTML: both key and value land inside a double-quoted attribute.
    const { value: html } = await OfficeGenerator.generate(ast, 'html', {
        metadataOverrides: { title: hostileValue, custom: { [hostileKey]: hostileValue } },
    } as any);
    check('html: injected key cannot open a tag', !/<script>alert\(1\)/.test(html as string),
        'custom metadata key escaped out of the meta attribute');
    check('html: injected value cannot open a tag', !/<script>alert\(2\)/.test(html as string),
        'metadata value escaped out of the meta attribute');

    // EPUB renders through the same HTML path and then into XML, where an unescaped value is
    // not merely an injection but makes the whole package fail to parse.
    const epub = (await OfficeGenerator.generate(ast, 'epub', {
        metadataOverrides: { title: hostileValue },
    } as any)).value as Uint8Array;
    const opf = strFromU8(unzipSync(epub)['OEBPS/content.opf']);
    check('epub: hostile title is escaped in the OPF', !opf.includes('<script>'),
        'raw markup reached the OPF package document');
    check('epub: OPF remains well-formed XML', !/<dc:title>[^<]*[<>][^<]*<\/dc:title>/.test(
        opf.replace(/<dc:title>|<\/dc:title>/g, m => m)) || opf.includes('&lt;'),
        'unescaped angle bracket inside dc:title');

    // Markdown frontmatter: a value containing a newline could otherwise close the `---` block
    // early and inject document content, or forge additional frontmatter keys.
    const { value: md } = await OfficeGenerator.generate(ast, 'md', {
        metadataOverrides: { title: 'a\n---\ninjected: true' },
    } as any);
    const frontmatter = String(md).split('---')[1] ?? '';
    check('md: newline in a metadata value cannot forge frontmatter keys',
        !/^injected:/m.test(frontmatter), 'value broke out of the frontmatter block');

    // CSV renders metadata as comments; a delimiter or newline must not fabricate rows/columns.
    // Needs a sheet-bearing AST: a paragraph-only document produces no CSV at all, so asserting
    // against it would pass without ever exercising the metadata path.
    const sheetAst = astWith([
        { type: 'sheet', metadata: { sheetName: 'S1' }, children: [
            { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'a' }] }] }
        ] } as any
    ]);
    const { value: csv } = await OfficeGenerator.generate(sheetAst, 'csv', {
        renderMetadata: true,
        metadataOverrides: { title: 'a,b\n=cmd|calc', custom: { 'k\n=HYPERLINK(1)': 'v' } },
    } as any);
    const csvText = typeof csv === 'string' ? csv : '';
    check('csv: metadata override is actually rendered', csvText.includes('# Title:'),
        `metadata comments absent, so the checks below would be vacuous: ${JSON.stringify(csvText.slice(0, 80))}`);
    check('csv: metadata comment cannot spawn a new line',
        !csvText.split('\n').some(l => l.trim().startsWith('=')),
        'a formula escaped onto its own line from a metadata comment');
    check('csv: every metadata line stays a comment',
        csvText.split('\n').filter(l => l.trim() !== '').slice(0, 3).every(l => l.startsWith('#') || l === 'a'),
        'a newline in a metadata value broke out of the comment prefix');

    // Plain text renders metadata as a structured `Key: value` block closed by a rule. A line
    // break in a value would forge fields the document never had - no code execution, but a lie
    // about the document's provenance, which consumers parsing that block would believe.
    const { value: textOut } = await OfficeGenerator.generate(ast, 'text', {
        renderMetadata: true,
        metadataOverrides: { title: 'Real\nAuthor: Attacker\n-------------------' },
    } as any);
    const headerLines = String(textOut).split('\n');
    check('text: metadata header is rendered', headerLines[0].startsWith('Title: '),
        'renderMetadata produced no header, so the check below would be vacuous');
    check('text: newline in a metadata value cannot forge a field',
        !headerLines.some(l => l.startsWith('Author: ')),
        `forged an Author line the document never had: ${JSON.stringify(headerLines.slice(0, 4))}`);

    // A malformed date must not render literal "Invalid Date" as if it were real provenance.
    const { value: badDate } = await OfficeGenerator.generate(ast, 'text', {
        renderMetadata: true, metadataOverrides: { created: 'not-a-date' },
    } as any);
    check('text: malformed date is omitted, not printed as "Invalid Date"',
        !String(badDate).includes('Invalid Date'), 'literal Invalid Date reached the header');

    // The EPUB timestamp is interpolated into the OPF without escaping, which is only safe
    // because it is normalised through toISOString(). Asserting it directly so that a future
    // change reintroducing a verbatim passthrough fails here rather than silently allowing
    // markup into the package document.
    const hostileDate = (await OfficeGenerator.generate(ast, 'epub', {
        metadataOverrides: { modified: '2024-01-01T00:00:00Z"/><script>alert(3)</script><meta x="' as any },
    } as any)).value as Uint8Array;
    const hostileOpf = strFromU8(unzipSync(hostileDate)['OEBPS/content.opf']);
    check('epub: dcterms:modified cannot carry markup', !hostileOpf.includes('<script>'),
        'an unnormalised timestamp injected markup into the OPF');
    check('epub: dcterms:modified is a well-formed instant',
        /<meta property="dcterms:modified">\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z<\/meta>/.test(hostileOpf),
        `timestamp not normalised: ${hostileOpf.match(/dcterms:modified">[^<]*/)?.[0]}`);

    // RTF: a brace or backslash in a value would otherwise close the \info group early.
    const { value: rtf } = await OfficeGenerator.generate(ast, 'rtf', {
        renderMetadata: true,
        metadataOverrides: { title: '}\\b evil{' },
    } as any);
    const info = String(rtf).slice(String(rtf).indexOf('{\\info'));
    check('rtf: braces in a metadata value are escaped',
        info.includes('\\}') || info.includes('\\{'), 'unescaped brace inside the \\info group');
}

/**
 * Config resolution is an attack surface distinct from document content: a host application that
 * accepts a JSON config blob hands us an object whose keys the caller did not choose. A config
 * parsed from JSON can carry `__proto__` as a genuine own enumerable key (an object literal
 * cannot), so a recursive merge writes it straight onto `Object.prototype` and corrupts every
 * object in the process - not just our output.
 */
function configPollutionTests() {
    console.log('- Config resolution (prototype pollution)...');

    const clean = () => {
        for (const k of ['polluted', 'pollutedNested', 'pollutedParser', 'pollutedCtor']) {
            delete (Object.prototype as any)[k];
        }
    };
    clean(); // start from a known state so an earlier failure can't cascade into a false pass

    // Every sub-config goes through the same merge, so every one is a route in. The probe value
    // must be one the config's own validation accepts - a rejected value falls back to the
    // default, which would make the "merge applied" guard fail even though the merge ran.
    const subConfigs: Array<[string, string, string]> = [
        ['htmlConfig', 'containerWidth', '640px'], ['mdConfig', 'dialect', 'github'],
        ['pdfConfig', 'format', 'Letter'], ['csvConfig', 'columnDelimiter', ';'],
        ['textConfig', 'newlineDelimiter', '\\r\\n'], ['chunksConfig', 'strategy', 'fixed-size'],
    ];
    for (const [sub, probeKey, probeValue] of subConfigs) {
        const raw = JSON.parse(`{"${sub}":{"${probeKey}":"${probeValue}","__proto__":{"polluted":"YES"}}}`);
        const cfg: any = resolveGeneratorConfig('html' as any, undefined as any, raw);
        const expected = JSON.parse(`"${probeValue}"`);
        // Guard first: if the merge silently did nothing, the pollution assertion below would
        // pass for the wrong reason. This is the failure mode that let an earlier vacuous test
        // in this very file sit green while the defect it named went unnoticed.
        check(`config: ${sub} merge actually applied`, cfg[sub]?.[probeKey] === expected,
            `nothing merged into ${sub}, so the pollution check below would be vacuous`);
        check(`config: __proto__ in ${sub} cannot reach Object.prototype`,
            ({} as any).polluted === undefined, `Object.prototype.polluted = ${({} as any).polluted}`);
        check(`config: ${sub} merge returns a clean prototype`,
            Object.getPrototypeOf(cfg) === Object.prototype, 'returned config inherits an attacker-chosen prototype');
        clean();
    }

    // Nested depth: the recursion must carry the guard down, not just check the top level.
    const nested = JSON.parse('{"htmlConfig":{"injections":{"headEnd":"__PROBE__","__proto__":{"pollutedNested":"YES"}}}}');
    const nestedCfg: any = resolveGeneratorConfig('html' as any, undefined as any, nested);
    check('config: nested merge actually applied', nestedCfg.htmlConfig?.injections?.headEnd === '__PROBE__',
        'nothing merged, so the nested pollution check would be vacuous');
    check('config: nested __proto__ cannot reach Object.prototype', ({} as any).pollutedNested === undefined);
    clean();

    // `constructor` is the other name that reaches a prototype through an ordinary write.
    const ctor = JSON.parse('{"htmlConfig":{"containerWidth":"720px","constructor":{"prototype":{"pollutedCtor":"YES"}}}}');
    const ctorCfg: any = resolveGeneratorConfig('html' as any, undefined as any, ctor);
    check('config: constructor-route merge actually applied', ctorCfg.htmlConfig?.containerWidth === '720px',
        'nothing merged, so the constructor check would be vacuous');
    check('config: constructor route cannot reach Object.prototype', ({} as any).pollutedCtor === undefined);
    check('config: constructor is not shadowed on the sub-config',
        !Object.prototype.hasOwnProperty.call(ctorCfg.htmlConfig, 'constructor'),
        'attacker-supplied constructor landed as an own property');
    clean();

    // Parser config takes a different path (Object.assign, not the recursive merge). Object.assign
    // does not pollute Object.prototype - it writes via [[Set]], so `__proto__` hits the inherited
    // setter - but that setter REPLACES the target's prototype, so the returned config silently
    // inherits attacker properties. Assert the returned object's prototype directly.
    const parserRaw = JSON.parse('{"newlineDelimiter":"__PROBE__","__proto__":{"pollutedParser":"YES"}}');
    const parserCfg: any = resolveParserConfig(parserRaw);
    check('config: parser merge actually applied', parserCfg.newlineDelimiter === '__PROBE__',
        'nothing merged, so the parser checks below would be vacuous');
    check('config: parser __proto__ cannot reach Object.prototype', ({} as any).pollutedParser === undefined);
    check('config: parser config keeps a clean prototype',
        Object.getPrototypeOf(parserCfg) === Object.prototype,
        'Object.assign invoked the __proto__ setter and replaced the config prototype');
    check('config: parser config did not inherit attacker properties',
        parserCfg.pollutedParser === undefined, `inherited pollutedParser = ${parserCfg.pollutedParser}`);
    clean();
}

/**
 * `styleMap` is caller config, not document content, but it is public API and a host app may
 * build one from user-influenced values. Two of its emission paths bypassed the escaping every
 * other node type gets: the spreadsheet row and sheet rebuild the class attribute from the raw
 * mapping array instead of reusing the escaped `className`, and both styleMap attribute loops
 * escaped the value while interpolating the NAME unchecked.
 */
async function styleMapTests() {
    console.log('- HtmlGenerator styleMap (integration)...');

    const xlsx = path.join(__dirname, '..', 'files', 'test.xlsx');
    if (!fs.existsSync(xlsx)) { check('styleMap: xlsx fixture present', false, 'missing test.xlsx'); return; }
    const sheetAst = await OfficeParser.parseOffice(xlsx, {} as any);

    // Spreadsheet row: hostile class AND hostile attribute name.
    const rowOut = String((await sheetAst.to('html', { styleMap: [{ selector: { nodeType: 'row' },
        output: { tag: 'tr', classes: ['r" onmouseover="alert(1)'], attributes: { 'q" onfocus="alert(2)" w': 'v' } } }] } as any)).value);
    const tr = rowOut.match(/<tr[^>]*excel-row[^>]*>/)?.[0] ?? '';
    check('styleMap: spreadsheet row actually rendered', tr.length > 0,
        'no excel-row <tr> emitted, so the checks below would be vacuous');
    check('styleMap: row class cannot break out', !/onmouseover\s*=\s*"/.test(tr), tr);
    check('styleMap: row attribute name cannot break out', !/onfocus\s*=\s*"/.test(tr), tr);

    // Sheet container.
    const sheetOut = String((await sheetAst.to('html', { styleMap: [{ selector: { nodeType: 'sheet' },
        output: { tag: 'div', classes: ['s" onmouseover="alert(3)'] } }] } as any)).value);
    const div = sheetOut.match(/<div[^>]*spreadsheet-sheet[^>]*>/)?.[0] ?? '';
    check('styleMap: sheet container actually rendered', div.length > 0,
        'no spreadsheet-sheet <div> emitted, so the check below would be vacuous');
    check('styleMap: sheet class cannot break out', !/onmouseover\s*=\s*"/.test(div), div);

    // Paragraph path: attribute name only (its class path was already escaped).
    const pAst = astWith([{ type: 'paragraph', metadata: { style: 'Custom' },
        children: [{ type: 'text', text: 'hi' }] } as any]);
    const sm = (output: any) => ({ styleMap: [{ selector: { nodeType: 'paragraph', attributes: { style: 'Custom' } }, output }] } as any);
    const pOut = String((await OfficeGenerator.generate(pAst, 'html', sm({ tag: 'p', attributes: { 'x" onmouseover="alert(4)" z': 'y' } }))).value);
    check('styleMap: paragraph actually rendered', /<p[^>]*>hi/.test(pOut),
        'no paragraph emitted, so the check below would be vacuous');
    check('styleMap: paragraph attribute name cannot break out', !/onmouseover\s*=\s*"/.test(pOut), pOut);

    // Positive control: rejecting hostile names must not also drop legitimate ones, or the
    // "fix" would be silently breaking styleMap for every real user.
    const benign = String((await OfficeGenerator.generate(pAst, 'html', sm({ tag: 'p', classes: ['lead'], attributes: { 'data-role': 'intro' } }))).value);
    check('styleMap: legitimate class still emitted', /class="[^"]*lead/.test(benign), benign);
    check('styleMap: legitimate data- attribute still emitted', /data-role="intro"/.test(benign), benign);

    // Duplicate attribute names are merely invalid in HTML but FATAL in EpubGenerator's XHTML,
    // so scan every emitted tag - the sheet <div> is the one that reaches the EPUB path.
    for (const tag of sheetOut.match(/<[a-zA-Z][^>]*>/g) || []) {
        const names = (tag.match(/\s([a-zA-Z-]+)=/g) || []).map(a => a.trim().slice(0, -1).toLowerCase());
        const dupes = names.filter((n, i) => names.indexOf(n) !== i);
        if (dupes.length > 0) { check('styleMap: no duplicate attribute names', false, `${dupes.join(',')} in ${tag}`); return; }
    }
    check('styleMap: no duplicate attribute names', true);

    // --- output.tag ---------------------------------------------------------------------
    // HtmlGenerator now honours styleMap output.tag (it previously wrote the value and then
    // shadowed it in every switch branch, so it was silently ignored). The shadowing was the
    // only thing stopping a hostile tag from injecting, so honouring it REQUIRES the allowlist:
    // a tag name is interpolated into both `<TAG>` and `</TAG>`, where no escaping applies.
    const fragment = { htmlConfig: { standalone: false } };
    const tagOut = async (tag: string, warns?: any[]) => String((await OfficeGenerator.generate(pAst, 'html', {
        ...fragment, ...(warns ? { onWarning: (w: any) => warns.push(w) } : {}),
        styleMap: [{ selector: { nodeType: 'paragraph', attributes: { style: 'Custom' } }, output: { tag } }],
    } as any)).value);

    // Honoured for the semantic elements a style mapping exists to express.
    for (const tag of ['h2', 'blockquote', 'section', 'em']) {
        const out = await tagOut(tag);
        check(`styleMap: output.tag "${tag}" is honoured`, out.includes(`<${tag}>`) && out.includes(`</${tag}>`),
            `mapping ignored: ${JSON.stringify(out.slice(0, 120))}`);
    }
    // Rejected, with a fallback to the default tag and a warning - never emitted.
    for (const tag of ['script', 'iframe', 'style', 'object', 'p><script>alert(1)</script><p']) {
        const warns: any[] = [];
        const out = await tagOut(tag, warns);
        check(`styleMap: output.tag ${JSON.stringify(tag.slice(0, 24))} is rejected`, !out.includes(`<${tag}`),
            `hostile tag reached output: ${JSON.stringify(out.slice(0, 160))}`);
        check(`styleMap: rejected tag falls back to the default`, /<p[\s>]/.test(out),
            `no fallback element emitted: ${JSON.stringify(out.slice(0, 160))}`);
        check(`styleMap: rejected tag warns`, warns.some(w => w.code === OfficeWarningType.INVALID_STYLE_MAP_TAG),
            'silently ignoring a caller-supplied tag gives them no way to find out');
    }
    check('styleMap: no script element from a hostile tag',
        !(await tagOut('p><script>alert(1)</script><p')).includes('<script'), 'script element emitted');
}

/**
 * RTF was the only generator with no URL scheme allowlist - `escapeRtf` neutralizes the field
 * metacharacters but says nothing about where the link points. A `file://` or UNC HYPERLINK in
 * Word is a phishing / NTLM-credential-leak vector, not just a rendering quirk.
 */
async function rtfUrlTests() {
    console.log('- RtfGenerator (URL schemes)...');

    const linkAst = (url: string) => astWith([{ type: 'paragraph', children: [
        { type: 'text', text: 'clickme', metadata: { link: url, linkType: 'external' } }] } as any]);
    const rtfFor = async (url: string) => String((await OfficeGenerator.generate(linkAst(url), 'rtf')).value);

    for (const url of ['javascript:alert(1)', 'vbscript:msgbox(1)', 'data:text/html,<script>',
                       'file:///C:/Windows/System32/calc.exe', '\\\\evil.com\\share\\x', '//evil.com/share']) {
        const rtf = await rtfFor(url);
        check(`rtf: ${JSON.stringify(url).slice(0, 40)} emits no HYPERLINK field`,
            !/HYPERLINK/.test(rtf), rtf.match(/HYPERLINK "[^"]*"/)?.[0] ?? rtf.slice(0, 120));
        // Degrade, don't delete: the link text is document content and must survive.
        check(`rtf: rejected link keeps its text`, rtf.includes('clickme'),
            `link text was dropped along with the URL: ${rtf.slice(0, 120)}`);
    }

    // Positive control. Without these the allowlist could be "reject everything" and still pass.
    for (const url of ['https://example.com/a?b=1', 'http://x.test/p', 'mailto:a@b.test', 'tel:+123', '#anchor', 'relative/path.html']) {
        const rtf = await rtfFor(url);
        check(`rtf: ${url} still emits a HYPERLINK field`, /HYPERLINK "/.test(rtf),
            `legitimate URL was dropped: ${rtf.slice(0, 140)}`);
    }

    // UNC is rejected for RTF specifically and must NOT become a global policy: in a browser
    // `//host/share` is an ordinary protocol-relative URL and blocking it there would break
    // legitimate links from older HTML sources.
    check('rtf: UNC rejection is RTF-only, HTML still allows protocol-relative',
        sanitizeUrl('//example.com/x') === '//example.com/x',
        `sanitizeUrl unexpectedly rejected a protocol-relative URL: ${JSON.stringify(sanitizeUrl('//example.com/x'))}`);

    // Field metacharacters must still be neutralized on an otherwise-allowed URL.
    const quoted = await rtfFor('https://example.com/a"}{\\b evil');
    check('rtf: quotes/braces in an allowed URL are escaped',
        !/HYPERLINK "[^"]*"\}\{\\b/.test(quoted), quoted.match(/HYPERLINK "[^"]*"/)?.[0] ?? '');
}

/**
 * ODF encodes runs of identical cells/rows with `table:number-columns-repeated` /
 * `table:number-rows-repeated` instead of repeating markup, so a few hundred bytes of XML can ask
 * the parser to materialize an arbitrary number of nodes, and the two multiply. The ZIP limits do
 * not help: the XML is tiny before decompression and the expansion happens afterwards.
 *
 * These assert the bound holds without breaking real documents, which legitimately carry very
 * large repeat counts (LibreOffice writes `number-rows-repeated="1048566"` for trailing empties).
 */
async function odfRepeatExpansionTests() {
    console.log('- OpenOfficeParser (repeated-cell expansion)...');

    const enc = (t: string) => new TextEncoder().encode(t);
    const doc = (inner: string) => `<?xml version="1.0" encoding="UTF-8"?><office:document-content ` +
        `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
        `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
        `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">` +
        `<office:body>${inner}</office:body></office:document-content>`;
    const build = (mime: string, inner: string) =>
        Buffer.from(zipSync({ mimetype: enc(mime), 'content.xml': enc(doc(inner)) }));
    const ODS = 'application/vnd.oasis.opendocument.spreadsheet';
    const ODT = 'application/vnd.oasis.opendocument.text';

    const countCells = (ast: any) => {
        let n = 0; const walk = (x: any) => { if (x.type === 'cell') n++; (x.children || []).forEach(walk); };
        ast.content.forEach(walk); return n;
    };
    const parse = async (buf: Buffer, fileType: string, limit?: number) => {
        const warns: any[] = [];
        const cfg: any = { fileType, onWarning: (w: any) => warns.push(w) };
        if (limit !== undefined) cfg.decompressionLimits = { maxTableCells: limit };
        const ast = await OfficeParser.parseOffice(buf, cfg);
        return { cells: countCells(ast), warned: warns.some(w => w.code === OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED) };
    };

    const LIMIT = 5000;

    // Single axis: a huge column repeat on a non-empty cell.
    const cols = build(ODS, `<office:spreadsheet><table:table table:name="S"><table:table-row>` +
        `<table:table-cell table:number-columns-repeated="5000000"><text:p>X</text:p></table:table-cell>` +
        `</table:table-row></table:table></office:spreadsheet>`);
    const rc = await parse(cols, 'ods', LIMIT);
    check('odf: column repeat is bounded', rc.cells <= LIMIT, `materialized ${rc.cells} cells against a limit of ${LIMIT}`);
    check('odf: column clamp warns', rc.warned, 'truncation must not be silent');

    // Both axes: this is the combination that exhausted memory, since each row repetition
    // deep-copies the whole cell array.
    const both = build(ODS, `<office:spreadsheet><table:table table:name="S">` +
        `<table:table-row table:number-rows-repeated="10000">` +
        `<table:table-cell table:number-columns-repeated="10000"><text:p>X</text:p></table:table-cell>` +
        `</table:table-row></table:table></office:spreadsheet>`);
    const rb = await parse(both, 'ods', LIMIT);
    check('odf: rows x cols product is bounded', rb.cells <= LIMIT, `materialized ${rb.cells} cells against a limit of ${LIMIT}`);

    // ODT/ODP keep empty cells on purpose (the grid is structural), so they have no empty-cell
    // skip and the budget is the only thing bounding them.
    const odt = build(ODT, `<office:text><table:table table:name="T"><table:table-row>` +
        `<table:table-cell table:number-columns-repeated="5000000"/>` +
        `</table:table-row></table:table></office:text>`);
    const ro = await parse(odt, 'odt', LIMIT);
    check('odf: ODT empty-cell repeat is bounded', ro.cells <= LIMIT, `materialized ${ro.cells} cells`);

    // MANY tables, each with a huge repeat. The budget is per document, so splitting the
    // expansion across tables must not multiply past the cap - the earlier single-table tests
    // would pass even with a per-table budget, which is exactly the hole this covers.
    const manyTables = build(ODT, '<office:text>' +
        (`<table:table table:name="T"><table:table-row><table:table-cell ` +
         `table:number-columns-repeated="1000000"><text:p>X</text:p></table:table-cell>` +
         `</table:table-row></table:table>`).repeat(20) + '</office:text>');
    const rm = await parse(manyTables, 'odt', LIMIT);
    check('odf: budget is per-document, not per-table', rm.cells <= LIMIT,
        `20 tables materialized ${rm.cells} cells against a per-document limit of ${LIMIT}`);

    // A garbage (non-numeric) repeat must render the cell once, not drain the whole budget and
    // silently drop every legitimate cell that follows it.
    const garbage = build(ODS, `<office:spreadsheet>` +
        `<table:table table:name="A"><table:table-row><table:table-cell ` +
        `table:number-columns-repeated="abc"><text:p>GARBAGE</text:p></table:table-cell></table:table-row></table:table>` +
        `<table:table table:name="B"><table:table-row><table:table-cell><text:p>LEGIT</text:p></table:table-cell></table:table-row></table:table>` +
        `</office:spreadsheet>`);
    const gWarns: any[] = [];
    const gAst = await OfficeParser.parseOffice(garbage, { fileType: 'ods', onWarning: (w: any) => gWarns.push(w), decompressionLimits: { maxTableCells: LIMIT } } as any);
    const gText = gAst.toText();
    check('odf: a garbage repeat does not drain the budget', gText.includes('LEGIT'),
        'a non-numeric repeat count consumed the budget and dropped a later legitimate cell');
    check('odf: a garbage repeat does not spuriously warn',
        !gWarns.some(w => w.code === OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED),
        'a non-numeric repeat count tripped the limit warning on a tiny document');

    // A huge repeat on an EMPTY spreadsheet cell (the normal ODF way to mark a trailing empty
    // run) must be skipped in O(1), not spun once per column. 2e8 would take ~1.4s as a loop.
    const emptyRun = build(ODS, `<office:spreadsheet><table:table table:name="S"><table:table-row>` +
        `<table:table-cell table:number-columns-repeated="200000000"/></table:table-row></table:table></office:spreadsheet>`);
    const t0 = Date.now();
    const eAst = await OfficeParser.parseOffice(emptyRun, { fileType: 'ods' } as any);
    const eMs = Date.now() - t0;
    check('odf: an empty repeated run is skipped, not spun', eMs < 200 && countCells(eAst) === 0,
        `empty run of 2e8 columns took ${eMs}ms and produced ${countCells(eAst)} cells`);

    // The bound must not fire on ordinary documents. A real .ods carries repeat counts in the
    // millions on empty runs; those cost nothing because empty spreadsheet cells are skipped.
    const realOds = path.join(__dirname, '..', 'files', 'test.ods');
    if (fs.existsSync(realOds)) {
        const warns: any[] = [];
        const ast = await OfficeParser.parseOffice(realOds, { onWarning: (w: any) => warns.push(w) } as any);
        const n = countCells(ast);
        check('odf: a real spreadsheet still parses fully', n > 0 && n < 10000, `${n} cells`);
        check('odf: a real spreadsheet is not clamped',
            !warns.some(w => w.code === OfficeWarningType.TABLE_CELL_LIMIT_EXCEEDED),
            'the bound fired on a legitimate document');
    }
}

/**
 * `abortSignal` is one of the escape hatches a consumer relies on for adversarial input, so it
 * has to actually interrupt work rather than only decline to start it. It previously did the
 * latter: parsers read it once before parsing and never again, and every generator except
 * ChunkingGenerator ignored it entirely.
 *
 * The generator cases matter individually because three generators *override*
 * `processNodeRecursive`, so a check in the base class alone leaves them inert - which is
 * precisely how HtmlGenerator and MarkdownGenerator were missed on the first pass.
 */
async function abortSignalTests() {
    console.log('- abortSignal (parser + generators)...');

    const enc = (t: string) => new TextEncoder().encode(t);
    const ods = Buffer.from(zipSync({
        mimetype: enc('application/vnd.oasis.opendocument.spreadsheet'),
        'content.xml': enc(`<?xml version="1.0" encoding="UTF-8"?><office:document-content ` +
            `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
            `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ` +
            `xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body>` +
            `<office:spreadsheet><table:table table:name="S"><table:table-row>` +
            `<table:table-cell table:number-columns-repeated="1000000"><text:p>X</text:p></table:table-cell>` +
            `</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`),
    }));

    // Parser: an aborted signal makes parseOffice reject rather than returning a parsed AST.
    //
    // This asserts only the honest guarantee, and deliberately does NOT claim to prove the
    // in-loop parser checks specifically. In a single thread that claim would be a lie: the
    // parser reads the signal at the top of parseOpenOffice, before the synchronous
    // parseContentXml loop, and nothing can flip the signal DURING that synchronous loop - no
    // callback, timer or microtask runs until the loop yields, and it does not. So any signal
    // this test could set is caught either by the top check (if set before it) or not at all (if
    // it could only be set mid-loop). The in-loop checks earn their place in two cases this test
    // cannot pin deterministically: an abort landing during the async archive extraction, and a
    // signal driven from another thread/worker. They are cheap, so they stay; the point here is
    // simply that an aborted parse does not silently succeed.
    const preAborted = new AbortController();
    preAborted.abort();
    let parserAborted = false;
    try { await OfficeParser.parseOffice(ods, { fileType: 'ods', abortSignal: preAborted.signal } as any); }
    catch (e: any) { parserAborted = /abort/i.test(String(e?.message)); }
    check('abort: an aborted parse rejects rather than succeeding', parserAborted,
        'parseOffice returned a parsed AST despite an aborted signal');

    // Generators: every text-based one, since three of them override the shared traversal.
    const src = path.join(__dirname, '..', 'files', 'test.docx');
    if (!fs.existsSync(src)) { check('abort: docx fixture present', false, 'missing test.docx'); return; }
    const ast = await OfficeParser.parseOffice(src, { extractAttachments: true } as any);

    for (const fmt of ['html', 'md', 'text', 'rtf', 'csv']) {
        const aborted = new AbortController();
        aborted.abort();
        let threw = false;
        try { await OfficeGenerator.generate(ast as any, fmt as any, { abortSignal: aborted.signal } as any); }
        catch (e: any) { threw = /abort/i.test(String(e?.message)); }
        check(`abort: ${fmt} generator honours the signal`, threw,
            'generation completed despite an already-aborted signal');
    }

    // Positive control: an un-aborted signal must not interfere with normal generation.
    const live = new AbortController();
    const { value } = await OfficeGenerator.generate(ast as any, 'md', { abortSignal: live.signal } as any);
    check('abort: an un-aborted signal does not block generation', String(value).length > 0,
        'a live signal suppressed output');
}

/** Silences the console for guards that report through it, without disabling the guard itself. */
const QUIET = { outputErrorToConsole: false } as any;

/** Encodes test fixture XML for zipSync. */
const zipEnc = (text: string) => new TextEncoder().encode(text);

/** Deterministic incompressible filler, so a truncation lands inside real deflate output. */
const zipFiller = (size: number) => {
    const bytes = new Uint8Array(size);
    let seed = 12345;
    for (let i = 0; i < size; i++) { seed = (seed * 1103515245 + 12345) & 0x7fffffff; bytes[i] = seed & 0xff; }
    return bytes;
};

/** Resolves to the rejection message, or '' if the promise resolved instead. */
async function rejectionMessage(run: () => Promise<unknown>): Promise<string> {
    try { await run(); return ''; }
    catch (e: any) { return String(e?.message); }
}

async function corruptArchiveTests() {
    console.log('- Corrupt / non-ZIP archive input (issue #107)...');

    const EXPECTED_MESSAGE = '[OfficeParser]: No readable entries found in ZIP data. The input is corrupt, truncated, or not a ZIP archive: every ZIP-based document format requires at least one entry.';

    // extractFiles used to surface fflate's "invalid zip data" for non-ZIP input; the 7.3.0
    // streaming rewrite silently resolved with [] instead, making a corrupt file
    // indistinguishable from a genuinely empty document (issue #107). Zero entries can never
    // be a valid office document, so this must reject, with the exact message from the table.
    for (const [name, buf] of [
        ['plain text', Buffer.from('not a real docx')],
        ['empty buffer', Buffer.alloc(0)],
        ['stray PK magic inside text', Buffer.from('hello PK\x03\x04 world, still not a zip')],
        // An archive that is well-formed but holds nothing is equally impossible as a document.
        ['valid but empty archive', Buffer.from(zipSync({}))],
        // Garbage carrying the End Of Central Directory signature must still be caught here,
        // rather than passing the truncation check and reporting the wrong reason.
        ['garbage containing an EOCD signature',
            Buffer.concat([Buffer.from('junk'), Buffer.from([0x50, 0x4b, 0x05, 0x06]), Buffer.alloc(30)])],
    ] as const) {
        const message = await rejectionMessage(() => extractFiles(buf, () => true, {}, QUIET));
        check(`corrupt zip: ${name} rejects with the exact typed message`, message === EXPECTED_MESSAGE,
            `got ${JSON.stringify(message)}`);
    }

    // Control: a real archive whose entries are ALL filtered out must still resolve empty.
    // The entry count is taken before the filter runs, so this stays a success, not a reject.
    const zipped = Buffer.from(zipSync({ 'unrelated.txt': zipEnc('x') }));
    let filteredOk = false;
    try { filteredOk = (await extractFiles(zipped, () => false, {}, QUIET)).length === 0; }
    catch { /* a rejection here would itself be the regression */ }
    check('corrupt zip: fully-filtered valid archive still resolves empty', filteredOk,
        'a valid archive with no matching entries must not be treated as corrupt');

    // End to end: the public parse API rejects instead of returning an empty AST. Exact
    // equality, not a substring: the single-report guard makes the prefix deterministic.
    const e2eMessage = await rejectionMessage(() =>
        OfficeParser.parseOffice(Buffer.from('not a real docx'), { fileType: 'docx', ...QUIET } as any));
    check('corrupt zip: parseOffice rejects for a corrupt docx buffer', e2eMessage === EXPECTED_MESSAGE,
        e2eMessage ? `unexpected message: ${JSON.stringify(e2eMessage)}` : 'parseOffice resolved instead of rejecting');

    // Errors raised inside extractFiles must honour the caller's reporting config like every
    // other issue, rather than always writing to the console.
    const routed: any[] = [];
    await rejectionMessage(() => extractFiles(Buffer.from('nope'), () => true, {},
        { onWarning: (issue: any) => routed.push(issue) } as any));
    check('corrupt zip: extraction errors route through onWarning',
        routed.length === 1 && routed[0].code === 'ZIP_NO_ENTRIES_FOUND',
        `got ${JSON.stringify(routed.map(i => i.code))}`);
}

async function truncatedArchiveTests() {
    console.log('- Truncated archive input...');

    const EXPECTED_MESSAGE = '[OfficeParser]: Malformed ZIP data: no End of Central Directory record was found at the end of the input. Either the file was cut off during download or transfer, or extra data follows the archive; in both cases the entries recovered from it cannot be trusted to be the whole document.';

    // The streaming reader rebuilds entries from local file headers alone, so a cut archive
    // still yields whatever preceded the cut and used to resolve as if nothing were wrong.
    // Requiring the trailer that ends every complete archive is what catches this.
    const archive = Buffer.from(zipSync({
        'word/document.xml': zipFiller(40000),
        'word/styles.xml': zipFiller(40000),
    }));

    // The trailer must sit at the end of the input, not merely somewhere in it. A ZIP comment
    // is length-limited to 16 bits, so a conformant archive ends within 64 KiB of its own
    // trailer; anything further is either a cut-off file or a payload appended after the
    // archive. Readers that locate the trailer from the end reject both, which is also what
    // this library did before the streaming rewrite, so the last case here is deliberate
    // rather than incidental: it keeps a smuggled payload from riding along inside a document.
    const APPENDED_BYTE = 0x41;
    for (const [name, buf] of [
        ['trailer sliced off', archive.subarray(0, archive.length - 10)],
        ['cut inside the central directory', archive.subarray(0, Math.floor(archive.length * 0.99))],
        ['data appended past the comment limit', Buffer.concat([archive, Buffer.alloc(70 * 1024, APPENDED_BYTE)])],
    ] as const) {
        const message = await rejectionMessage(() => extractFiles(buf as Buffer, () => true, {}, QUIET));
        check(`truncated zip: ${name} rejects with the exact typed message`, message === EXPECTED_MESSAGE,
            `got ${JSON.stringify(message)}`);
    }

    // Control for the boundary above: a trailer still reachable within the comment window is
    // a readable archive, so trailing bytes alone must not be treated as corruption.
    let withinWindow = -1;
    try {
        withinWindow = (await extractFiles(
            Buffer.concat([archive, Buffer.alloc(60 * 1024, APPENDED_BYTE)]), () => true, {}, QUIET)).length;
    } catch { /* a rejection here would be the over-strict failure */ }
    check('truncated zip: trailing bytes within the comment window still extract', withinWindow === 2,
        `expected 2 entries, got ${withinWindow}`);

    // A cut landing inside an entry's compressed data leaves that entry's completion callback
    // pending forever, so the promise could only settle if something else settles it. Raced
    // against a timer because the failure mode under test is "never settles", not "wrong value".
    const midEntry = archive.subarray(0, archive.length >> 1) as Buffer;
    const HUNG = Symbol('hung');
    const outcome = await Promise.race([
        extractFiles(midEntry, () => true, {}, QUIET).then(() => 'resolved').catch(() => 'rejected'),
        new Promise(resolve => setTimeout(() => resolve(HUNG), 5000)),
    ]);
    check('truncated zip: a cut inside entry data settles rather than hanging', outcome === 'rejected',
        outcome === HUNG ? 'extractFiles never settled' : `unexpectedly ${String(outcome)}`);

    // Control: the same archive intact must extract both entries.
    let intactCount = -1;
    try { intactCount = (await extractFiles(archive, () => true, {}, QUIET)).length; }
    catch { /* a rejection here would itself be the regression */ }
    check('truncated zip: the intact archive still extracts normally', intactCount === 2,
        `expected 2 entries, got ${intactCount}`);

    // End to end, with the exact single-prefixed message.
    const e2eMessage = await rejectionMessage(() =>
        OfficeParser.parseOffice(archive.subarray(0, archive.length - 10) as Buffer, { fileType: 'docx', ...QUIET } as any));
    check('truncated zip: parseOffice rejects a truncated docx', e2eMessage === EXPECTED_MESSAGE,
        e2eMessage ? `unexpected message: ${JSON.stringify(e2eMessage)}` : 'parseOffice resolved instead of rejecting');
}

async function missingMainPartTests() {
    console.log('- Readable archives missing their required part...');

    const requiredPartMessage = (fileType: string, part: string) =>
        `[OfficeParser]: Your ${fileType} file is a readable ZIP archive but is missing its required '${part}' part, so it cannot be a valid ${fileType} document. The file is corrupt, incomplete, or mislabeled. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce the error.`;

    const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';
    const ODT_MIME = 'application/vnd.oasis.opendocument.text';

    // A ZIP that extracts perfectly can still be a photo bundle, a partial upload or a
    // mislabeled file. Each of these is a valid archive with the format's main part removed,
    // which before this check parsed into an empty AST that no caller could distinguish from
    // a genuinely empty document.
    for (const [name, fileType, part, entries] of [
        ['docx without word/document.xml', 'docx', 'word/document.xml',
            { '[Content_Types].xml': zipEnc('<Types/>'), 'word/styles.xml': zipEnc('<styles/>') }],
        ['xlsx without xl/workbook.xml', 'xlsx', 'xl/workbook.xml',
            { 'xl/worksheets/sheet1.xml': zipEnc('<worksheet/>') }],
        ['pptx without ppt/presentation.xml', 'pptx', 'ppt/presentation.xml',
            { 'ppt/slides/slide1.xml': zipEnc('<p:sld/>') }],
        ['odt without content.xml', 'odt', 'content.xml',
            { mimetype: zipEnc(ODT_MIME), 'styles.xml': zipEnc('<styles/>') }],
        ['epub without an OPF', 'epub', 'OPF package document (.opf)',
            { 'META-INF/container.xml': zipEnc('<container/>'), 'ch1.xhtml': zipEnc('<html/>') }],
        // The reported symptom of #107 reproduced with a perfectly valid archive: a zip of
        // photos handed over as a docx. The entry-count guard cannot see this one.
        ['a photo archive handed over as docx', 'docx', 'word/document.xml',
            { 'photos/a.jpg': zipEnc('x'.repeat(500)), 'notes.txt': zipEnc('hi') }],
        // Anchoring regression: an ODF file can carry Object N/content.xml for an embedded
        // chart. That must never stand in for the document body when the real one is absent.
        ['ods whose only content.xml is an embedded object', 'ods', 'content.xml',
            { mimetype: zipEnc(ODS_MIME), 'Object 1/content.xml': zipEnc('<chart/>') }],
    ] as const) {
        const buf = Buffer.from(zipSync(entries as any));
        const message = await rejectionMessage(() =>
            OfficeParser.parseOffice(buf, { fileType, ...QUIET } as any));
        check(`missing part: ${name} rejects with the exact typed message`,
            message === requiredPartMessage(fileType, part), `got ${JSON.stringify(message)}`);
    }

    // Positive controls: minimal but complete archives must still parse, with their text
    // intact. Without these the checks above could be satisfied by rejecting everything.
    const docx = Buffer.from(zipSync({
        'word/document.xml': zipEnc('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello docx</w:t></w:r></w:p></w:body></w:document>'),
    }));
    const docxAst = await OfficeParser.parseOffice(docx, { fileType: 'docx', ...QUIET } as any);
    check('missing part: a minimal complete docx still parses', docxAst.toText().includes('Hello docx'),
        `got ${JSON.stringify(docxAst.toText())}`);

    const pptx = Buffer.from(zipSync({
        'ppt/presentation.xml': zipEnc('<p:presentation/>'),
        'ppt/slides/slide1.xml': zipEnc('<?xml version="1.0"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hello slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>'),
    }));
    const pptxAst = await OfficeParser.parseOffice(pptx, { fileType: 'pptx', ...QUIET } as any);
    check('missing part: a minimal complete pptx still parses', pptxAst.toText().includes('Hello slide'),
        `got ${JSON.stringify(pptxAst.toText())}`);
    // ppt/presentation.xml is extracted for the check above, and the slide loop treats every
    // unrecognized file as a slide, so it must be skipped explicitly or it becomes an extra
    // empty slide in the deck.
    check('missing part: the presentation part does not become a phantom slide',
        pptxAst.content.length === 1, `expected 1 slide node, got ${pptxAst.content.length}`);

    const odt = Buffer.from(zipSync({
        mimetype: zipEnc(ODT_MIME),
        'content.xml': zipEnc('<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Hello odt</text:p></office:text></office:body></office:document-content>'),
    }));
    const odtAst = await OfficeParser.parseOffice(odt, { fileType: 'odt', ...QUIET } as any);
    check('missing part: a minimal complete odt still parses', odtAst.toText().includes('Hello odt'),
        `got ${JSON.stringify(odtAst.toText())}`);
}

async function incompleteArchiveWarningTests() {
    console.log('- Legitimately empty archives warn rather than fail...');

    const collect = () => { const issues: any[] = []; return { issues, config: { ...QUIET, onWarning: (i: any) => issues.push(i) } as any }; };

    // A workbook holding only chartsheets has no worksheets and no cell text. That is valid,
    // so it warns; the missing-part check above is what covers a workbook that is not one.
    const chartsOnly = collect();
    const xlsx = Buffer.from(zipSync({
        'xl/workbook.xml': zipEnc('<workbook/>'),
        'xl/_rels/workbook.xml.rels': zipEnc('<Relationships/>'),
    }));
    const xlsxAst = await OfficeParser.parseOffice(xlsx, { fileType: 'xlsx', ...chartsOnly.config });
    const sheetWarnings = chartsOnly.issues.filter(i => i.code === 'NO_WORKSHEETS_FOUND');
    check('empty archive: a chartsheet-only workbook resolves', xlsxAst.type === 'xlsx');
    check('empty archive: it warns exactly once about missing worksheets', sheetWarnings.length === 1,
        `got ${JSON.stringify(chartsOnly.issues.map(i => i.code))}`);
    check('empty archive: the worksheet warning text is exact',
        sheetWarnings[0]?.message === 'Workbook contains no worksheet parts (xl/worksheets/). If the workbook holds only chartsheets this is expected and there is simply no cell text to extract; otherwise the file may be incomplete.',
        `got ${JSON.stringify(sheetWarnings[0]?.message)}`);

    // PowerPoint can save a deck with no slides at all, so this warns rather than failing.
    const noSlides = collect();
    const pptx = Buffer.from(zipSync({ 'ppt/presentation.xml': zipEnc('<p:presentation/>') }));
    const pptxAst = await OfficeParser.parseOffice(pptx, { fileType: 'pptx', ...noSlides.config });
    const slideWarnings = noSlides.issues.filter(i => i.code === 'NO_SLIDES_FOUND');
    check('empty archive: a slide-less presentation resolves', pptxAst.type === 'pptx');
    check('empty archive: it warns exactly once about missing slides', slideWarnings.length === 1,
        `got ${JSON.stringify(noSlides.issues.map(i => i.code))}`);
    check('empty archive: the slide warning text is exact',
        slideWarnings[0]?.message === 'Presentation contains no slides (ppt/slides/). A legitimately empty presentation produces this too, but if you expected content the file may be incomplete.',
        `got ${JSON.stringify(slideWarnings[0]?.message)}`);
}

async function odfTypeResolutionTests() {
    console.log('- ODF type resolution (mimetype vs caller hint)...');

    const ODS_MIME = 'application/vnd.oasis.opendocument.spreadsheet';
    const spreadsheetBody = zipEnc('<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:spreadsheet><table:table table:name="S1"><table:table-row><table:table-cell><text:p>CellValue</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>');

    // The three ODF types share a parser that branches on the body shape it expects. With no
    // mimetype entry it used to assume text regardless of what the caller said, so a valid
    // spreadsheet was walked as a text document and came back empty.
    const noMimetype = Buffer.from(zipSync({ 'content.xml': spreadsheetBody }));
    const hinted = await OfficeParser.parseOffice(noMimetype, { fileType: 'ods', ...QUIET } as any);
    check('odf type: a caller hint resolves a spreadsheet with no mimetype entry', hinted.type === 'ods',
        `got type ${hinted.type}`);
    check('odf type: that spreadsheet\'s cells are actually parsed', hinted.toText().includes('CellValue'),
        `got ${JSON.stringify(hinted.toText())}`);

    // When the archive declares its own type, that stays authoritative over the hint.
    const withMimetype = Buffer.from(zipSync({ mimetype: zipEnc(ODS_MIME), 'content.xml': spreadsheetBody }));
    const declared = await OfficeParser.parseOffice(withMimetype, { fileType: 'odt', ...QUIET } as any);
    check('odf type: the archive mimetype beats a conflicting hint', declared.type === 'ods',
        `got type ${declared.type}`);

    // From a file path there is no fileType in config at all; the dispatcher supplies the
    // extension, which is the only thing this path can go on.
    const tmp = path.join(os.tmpdir(), `officeparser-no-mimetype-${process.pid}.ods`);
    fs.writeFileSync(tmp, noMimetype);
    try {
        const fromPath = await OfficeParser.parseOffice(tmp, QUIET);
        check('odf type: the file extension resolves a spreadsheet with no mimetype entry',
            fromPath.type === 'ods' && fromPath.toText().includes('CellValue'),
            `got type ${fromPath.type}, text ${JSON.stringify(fromPath.toText())}`);
    } finally { fs.unlinkSync(tmp); }

    // Supplying that type must not write it back into the caller's config, or every later parse
    // reusing the object would be pinned to the wrong format: parse an .odt, then a .docx with
    // the same config, and the second one is routed to the ODF parser. Config ownership in
    // general is covered by configOwnershipTests below; this pins the dispatch half of it.
    const reused: any = {
        ...QUIET, extractAttachments: true, ocr: false, fileType: null,
        decompressionLimits: { maxUncompressedBytes: 512 * 1024 * 1024, maxZipEntries: 10000, maxTableCells: 1000000 },
    };
    const odt = Buffer.from(zipSync({
        mimetype: zipEnc('application/vnd.oasis.opendocument.text'),
        'content.xml': zipEnc('<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text><text:p>Text doc</text:p></office:text></office:body></office:document-content>'),
    }));
    const docx = Buffer.from(zipSync({
        'word/document.xml': zipEnc('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Word doc</w:t></w:r></w:p></w:body></w:document>'),
    }));
    await OfficeParser.parseOffice(odt, { ...reused, fileType: 'odt' });
    check('odf type: parsing does not pin the caller\'s config to that type',
        reused.fileType === null, `caller config fileType became ${JSON.stringify(reused.fileType)}`);
    const afterOdt = await OfficeParser.parseOffice(docx, { ...reused, fileType: 'docx' });
    check('odf type: a reused config still routes a later docx to the Word parser',
        afterOdt.type === 'docx' && afterOdt.toText().includes('Word doc'),
        `got type ${afterOdt.type}, text ${JSON.stringify(afterOdt.toText())}`);
}

async function configOwnershipTests() {
    console.log('- Config ownership across parses...');

    // A parse installs per-call state on the config it is given: the collector that gathers one
    // document's warnings for its `ast.warnings`. If config resolution hands back the caller's
    // own object, that state attaches to something the caller may reuse, and every later parse
    // appends its warnings to earlier, already-returned ASTs while retaining their arrays.
    //
    // This only happens for a config complete enough to skip the defaults merge, so these use
    // one that genuinely qualifies: `ocrConfig` carrying `language` and `workerPath`.
    const fullConfig = (): any => ({
        ocrConfig: { language: 'eng', workerPath: '', abortSignal: null },
        outputErrorToConsole: false,
        fileType: 'pptx',
    });

    // A presentation with no slides warns, which gives each parse exactly one warning to trace.
    const slideless = Buffer.from(zipSync({ 'ppt/presentation.xml': zipEnc('<p:presentation/>') }));

    const shared = fullConfig();
    const PARSE_COUNT = 4;
    const asts: any[] = [];
    for (let i = 0; i < PARSE_COUNT; i++) asts.push(await OfficeParser.parseOffice(slideless, shared));

    const counts = asts.map(ast => ast.warnings.length);
    check('config ownership: each parse keeps only its own warnings',
        counts.every(count => count === 1), `warnings per AST: ${JSON.stringify(counts)}`);
    check('config ownership: a returned AST is not modified by a later parse',
        asts[0].warnings.length === 1,
        `the first AST accumulated ${asts[0].warnings.length} warnings over ${PARSE_COUNT} parses`);

    // The caller's object must come back exactly as it went in.
    check('config ownership: parsing does not add fields to the caller\'s config',
        !('decompressionLimits' in shared), 'decompressionLimits was written onto the caller\'s object');
    check('config ownership: parsing does not replace the caller\'s onWarning',
        shared.onWarning === undefined, 'the warning collector was written onto the caller\'s object');

    // The caller's own handler still fires, once per warning, for every parse.
    const observed: string[] = [];
    const withHandler = { ...fullConfig(), onWarning: (issue: any) => observed.push(issue.code) };
    await OfficeParser.parseOffice(slideless, withHandler);
    await OfficeParser.parseOffice(slideless, withHandler);
    check('config ownership: the caller\'s handler fires once per warning per parse',
        observed.length === 2 && observed.every(code => code === 'NO_SLIDES_FOUND'),
        `got ${JSON.stringify(observed)}`);

    // Copying must not extend to values whose identity carries meaning. A cloned AbortSignal is
    // no longer tied to its controller, so cancellation would silently stop working.
    const aborted = new AbortController();
    aborted.abort();
    const resolved: any = resolveParserConfig({ ...fullConfig(), abortSignal: aborted.signal } as any);
    check('config ownership: abortSignal keeps its identity through resolution',
        resolved.abortSignal === aborted.signal, 'the signal was copied instead of referenced');
    let abortName = '';
    try { await OfficeParser.parseOffice(slideless, { ...fullConfig(), abortSignal: aborted.signal }); }
    catch (e: any) { abortName = e?.name; }
    check('config ownership: an aborted signal still cancels a parse', abortName === 'AbortError',
        `expected AbortError, got ${JSON.stringify(abortName)}`);

    // Containers, by contrast, must be fresh so per-parse writes cannot reach the caller.
    const source = fullConfig();
    const copy: any = resolveParserConfig(source);
    check('config ownership: nested config containers are copied',
        copy.ocrConfig !== source.ocrConfig, 'ocrConfig was shared with the caller');

    // Generation has the same contract. Its per-run normalization rewrites an unusable
    // containerWidth to 'auto'; done to the caller's object that both edits a value they still
    // hold and hides the problem from every later run, so the same config would report it once
    // and then look clean. An AST built by hand carries no config of its own, which is what
    // lets a complete generator config skip the merge and reach this path.
    const generatorConfig: any = resolveGeneratorConfig('html', undefined, { outputErrorToConsole: false } as any);
    generatorConfig.htmlConfig.containerWidth = 'not-a-width';
    const generatorWarnings: string[] = [];
    generatorConfig.onWarning = (issue: any) => generatorWarnings.push(issue.code);

    const ast = astWith([{ type: 'paragraph', text: 'Hi', children: [{ type: 'text', text: 'Hi', formatting: {} }] }]);
    delete (ast as any).config;
    await OfficeGenerator.generate(ast as any, 'html' as any, generatorConfig);
    const widthAfterFirstRun = generatorConfig.htmlConfig.containerWidth;
    await OfficeGenerator.generate(ast as any, 'html' as any, generatorConfig);

    check('config ownership: generating does not rewrite the caller\'s containerWidth',
        widthAfterFirstRun === 'not-a-width',
        `caller's width became ${JSON.stringify(widthAfterFirstRun)} after one generate`);
    check('config ownership: an invalid width warns on every run, not just the first',
        generatorWarnings.filter(code => code === 'INVALID_CONTAINER_WIDTH').length === 2,
        `got ${JSON.stringify(generatorWarnings)}`);

    // The same identity rule applies on the generator side.
    const generatorSource: any = resolveGeneratorConfig('html', undefined, { outputErrorToConsole: false } as any);
    const generatorCopy: any = resolveGeneratorConfig('html', undefined, generatorSource);
    check('config ownership: generator containers are copied',
        generatorCopy.htmlConfig !== generatorSource.htmlConfig, 'htmlConfig was shared with the caller');
    check('config ownership: generator callbacks keep their identity',
        generatorCopy.onNode === generatorSource.onNode, 'onNode was replaced during resolution');
}

function errorReportingTests() {
    console.log('- Error reporting (single report, single prefix)...');

    // Parser errors pass through getWrappedError on their way out. It exists to give raw
    // third-party failures OfficeParser context, but a typed error has already been reported
    // and already carries the header, so re-wrapping it reported the same issue twice, added
    // a second '[OfficeParser]: ' prefix, and flattened its code to FILE_CORRUPTED.
    const reported: any[] = [];
    const config = { onWarning: (issue: any) => reported.push(issue) } as any;
    const typed = getOfficeError(OfficeErrorType.REQUIRED_PART_MISSING, config,
        { fileType: 'docx', part: 'word/document.xml' });
    const wrapped = getWrappedError(typed, config);

    check('error reporting: a typed error passes through the wrapper untouched', wrapped === typed,
        'getWrappedError rebuilt an error that was already an OfficeParser error');
    check('error reporting: it is reported exactly once', reported.length === 1,
        `got ${reported.length} reports: ${JSON.stringify(reported.map(i => i.code))}`);
    check('error reporting: the reported code is preserved, not flattened',
        reported[0]?.code === OfficeErrorType.REQUIRED_PART_MISSING, `got ${reported[0]?.code}`);
    check('error reporting: the message carries exactly one header',
        (String(wrapped.message).match(/\[OfficeParser\]: /g) || []).length === 1,
        `got ${JSON.stringify(wrapped.message)}`);
    check('error reporting: the structured issue is exposed on the error',
        (typed as any).officeIssue?.code === OfficeErrorType.REQUIRED_PART_MISSING,
        'officeIssue missing from the returned error');

    // A raw third-party error still gets wrapped, which is the behavior being preserved.
    const rawReported: any[] = [];
    const raw = new Error('invalid zip data');
    const rawWrapped = getWrappedError(raw, { onWarning: (i: any) => rawReported.push(i) } as any);
    check('error reporting: an untyped error is still wrapped', rawWrapped !== raw
        && String(rawWrapped.message) === '[OfficeParser]: invalid zip data',
        `got ${JSON.stringify(rawWrapped.message)}`);
    check('error reporting: an untyped error is reported once as corruption',
        rawReported.length === 1 && rawReported[0].code === OfficeErrorType.FILE_CORRUPTED,
        `got ${JSON.stringify(rawReported.map(i => i.code))}`);
}

async function main() {
    console.log('Running sanitization security tests...\n');
    unitTests();
    configPollutionTests();
    await htmlTests();
    await htmlAttributeBagTests();
    await markdownTests();
    await csvTests();
    await metadataOverrideTests();
    await styleMapTests();
    await rtfUrlTests();
    await odfRepeatExpansionTests();
    await abortSignalTests();
    await corruptArchiveTests();
    await truncatedArchiveTests();
    await missingMainPartTests();
    await incompleteArchiveWarningTests();
    await odfTypeResolutionTests();
    await configOwnershipTests();
    errorReportingTests();

    console.log(`\n${failed === 0 ? '✓' : '✗'} Sanitization tests: ${passed} passed, ${failed} failed`);
    if (failed > 0) process.exit(1);
}

main().catch(err => { console.error(err); process.exit(1); });
