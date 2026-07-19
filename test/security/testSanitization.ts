/**
 * Security regression tests for output sanitization.
 *
 * Every string in the AST is treated as attacker-controlled (it comes from an
 * untrusted document). These tests lock in that document-supplied values can't
 * break out of their destination context (HTML attribute/tag, inline script,
 * CSS, a CSV formula, a Markdown link, an RTF group) in the generated output.
 */
import * as assert from 'assert';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { strFromU8, unzipSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import { OfficeParserAST } from '../../src/types';
import { resolveGeneratorConfig, resolveParserConfig } from '../../src/utils/configUtils';
import {
    escapeHtml, escapeXml, sanitizeCssValue, sanitizeUrl, sanitizeImageUrl,
    serializeForInlineScript, csvSafeCell, escapeRtf, markdownEscapeText, sanitizeMarkdownUrl
} from '../../src/utils/sanitize';

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

async function main() {
    console.log('Running sanitization security tests...\n');
    unitTests();
    configPollutionTests();
    await htmlTests();
    await htmlAttributeBagTests();
    await markdownTests();
    await csvTests();
    await metadataOverrideTests();

    console.log(`\n${failed === 0 ? '✓' : '✗'} Sanitization tests: ${passed} passed, ${failed} failed`);
    if (failed > 0) process.exit(1);
}

main().catch(err => { console.error(err); process.exit(1); });
