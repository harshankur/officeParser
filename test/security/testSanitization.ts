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
import { strFromU8, unzipSync, zipSync, zlibSync } from 'fflate';
import { OfficeGenerator } from '../../src/OfficeGenerator';
import { OfficeParser } from '../../src/OfficeParser';
import { OfficeTemplate } from '../../src/OfficeTemplate';
import { OfficeParserAST, OfficeWarningType } from '../../src/types';
import { resolveGeneratorConfig, resolveParserConfig } from '../../src/utils/configUtils';
import {
    escapeHtml, escapeXml, sanitizeCssValue, sanitizeUrl, sanitizeImageUrl,
    serializeForInlineScript, csvSafeCell, escapeRtf, markdownEscapeText, sanitizeMarkdownUrl, sanitizeRtfUrl,
    escapeLatex, latexComment, sanitizeLatexUrl, sanitizeLatexMath, sanitizeLatexImagePath,
    sanitizeCommentText, latexSourceComment
} from '../../src/utils/sanitize';
import { extractFiles } from '../../src/utils/zipUtils';
import { parseXmlString } from '../../src/utils/xmlUtils';
import { getOfficeError, getWrappedError } from '../../src/utils/errorUtils';
import { terminateOcr } from '../../src/utils/ocrUtils';
import { OfficeErrorType } from '../../src/types';
import { buildTaggedNodes } from '../../src/parsers/pdf/structTree';
import { computeDocContext } from '../../src/parsers/pdf/textLayout';

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
        getImages: () => []
    } as any;
}

function unitTests() {
    console.log('- Sanitize module (unit)...');
    // sanitizeCommentText: only comment-closing sequences change; ordinary text is untouched.
    check('comment text: plain text unchanged', sanitizeCommentText(' a note - with -- dashes ') === ' a note - with -- dashes ');
    check('comment text: --> neutralized', sanitizeCommentText('a --> b') === 'a --&gt; b');
    check('comment text: --!> neutralized', sanitizeCommentText('a --!> b') === 'a --!&gt; b');
    check('comment text: leading > neutralized', sanitizeCommentText('> b') === '&gt; b');
    check('comment text: leading -> neutralized', sanitizeCommentText('-> b') === '-&gt; b');
    check('comment text: non-string is empty', sanitizeCommentText(undefined as any) === '');

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
    // A quote would end the HYPERLINK argument once a reader decodes `\'22`: it is percent-encoded.
    const quoted = sanitizeRtfUrl('https://ok.example/" "file://evil.example/share/i');
    check('rtf url quote cannot end the field argument', !quoted.includes('"') && !quoted.includes("\\'22") && quoted.includes('%22'), quoted);
    // Full-width formula triggers, and a cell holding a delimiter of another locale, are guarded.
    check('csv full-width = guarded', csvSafeCell('\uFF1D1+1', ',').includes(`'`));
    check('csv ; quoted in a comma file', csvSafeCell('a;=1+1', ',') === '"a;=1+1"');
    check('css: a character reference cannot add a declaration', !sanitizeCssValue('red&#59position:fixed').includes('&') && !sanitizeCssValue('&#117rl(http://x)').includes('&'));
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

async function htmlSourceAttributesTests() {
    console.log('- HtmlGenerator sourceAttributes emission...');
    // Every sourceAttributes sink writes an attacker-influenced value into a data-* attribute. It
    // must be entity-escaped so it can neither break out of the attribute (") nor open a tag (< >).
    // Each payload attempts both. standalone:false keeps the output to the body so the document
    // shell can't produce a false positive.
    const ATTR = 'x" onerror="alert(1)';
    const TAG = '"><script>alert(1)</script>';
    const cfg = { htmlConfig: { sourceAttributes: true, standalone: false } };

    const cases: Array<[string, any[]]> = [
        ['wikilink data-target', [{ type: 'paragraph', children: [
            { type: 'text', text: 'link', metadata: { wikilink: true, link: ATTR, linkType: 'internal' } }] }]],
        ['wikilink data-alias', [{ type: 'paragraph', children: [
            { type: 'text', text: TAG, metadata: { wikilink: true, link: 'Page', linkType: 'internal' } }] }]],
        ['citation data-key', [{ type: 'paragraph', children: [
            { type: 'text', text: 'c', metadata: { citationKey: ATTR } }] }]],
        ['inline math data-math', [{ type: 'code', text: ATTR, metadata: { math: 'inline' } }]],
        ['block math data-math', [{ type: 'code', text: TAG, metadata: { math: 'block' } }]],
        ['mermaid data-mermaid', [{ type: 'code', text: TAG, metadata: { language: 'mermaid' } }]],
        ['block source comment data-html-comment', [{ type: 'comment', text: TAG, metadata: { sourceSyntax: 'html' } }]],
        ['inline source comment data-html-comment', [{ type: 'paragraph', children: [
            { type: 'comment', text: ATTR, metadata: { sourceSyntax: 'html' } }] }]],
    ];
    for (const [name, content] of cases) {
        const out = (await OfficeGenerator.generate(astWith(content), 'html', cfg as any)).value as string;
        check(`html: ${name} no attribute breakout`, !/onerror\s*=\s*"/.test(out), `broke out: ${JSON.stringify(out.slice(0, 200))}`);
        check(`html: ${name} no tag breakout`, !/<script/i.test(out), `tag injected: ${JSON.stringify(out.slice(0, 200))}`);
    }

    // Positive control: the emission must actually happen, or the checks above are vacuous.
    const okMath = (await OfficeGenerator.generate(astWith([{ type: 'code', text: 'E=mc^2', metadata: { math: 'inline' } }]), 'html', cfg as any)).value as string;
    check('html: sourceAttributes actually emits data-math', /data-math="E=mc\^2"/.test(okMath), okMath.slice(0, 200));
    const okComment = (await OfficeGenerator.generate(astWith([{ type: 'comment', text: ' note ', metadata: { sourceSyntax: 'html' } }]), 'html', cfg as any)).value as string;
    check('html: sourceAttributes actually emits data-html-comment', okComment.includes('<span data-html-comment=" note "></span>'), okComment.slice(0, 200));
}

/**
 * Remove HTML comments the way a browser tokenizes them: `<!-->` and `<!--->` close at once; otherwise a
 * comment runs to the first `-->` or `--!>` (or the end). What is left is what would render/execute.
 */
function stripCommentsLikeABrowser(html: string): string {
    let out = '';
    let i = 0;
    while (i < html.length) {
        const open = html.indexOf('<!--', i);
        if (open === -1) { out += html.slice(i); break; }
        out += html.slice(i, open);
        const body = open + 4;
        if (html.startsWith('>', body)) { i = body + 1; continue; }
        if (html.startsWith('->', body)) { i = body + 2; continue; }
        const close = html.slice(body).search(/--!?>/);
        if (close === -1) break;
        i = body + close + html.slice(body + close).match(/^--!?>/)![0].length;
    }
    return out;
}

async function sourceCommentBreakoutTests() {
    console.log('- Source comment (<!-- -->) breakout in Markdown and HTML output...');
    // A comment closes at the first `-->` or `--!>`, and a leading `>`/`->` closes an empty one, so
    // AST text carrying those could end the comment early and turn the rest into live markup. Whatever
    // the payload, nothing outside a comment may survive once comments are removed as a browser would.
    const payloads = [
        ' x --> <script>alert(1)</script> ',
        ' x --!> <script>alert(1)</script> ',
        '> <script>alert(1)</script> ',
        '-> <script>alert(1)</script> ',
    ];
    for (const text of payloads) {
        const cases: Array<[string, any[]]> = [
            ['block', [{ type: 'comment', text, metadata: { sourceSyntax: 'html' } }]],
            ['inline', [{ type: 'paragraph', children: [{ type: 'text', text: 'a' }, { type: 'comment', text, metadata: { sourceSyntax: 'html' } }] }]],
        ];
        for (const [where, content] of cases) {
            for (const format of ['md', 'html'] as const) {
                const out = String((await OfficeGenerator.generate(astWith(content), format, { htmlConfig: { standalone: false } } as any)).value);
                check(`${format} ${where} comment ${JSON.stringify(text.slice(0, 7))}: no live markup escapes the comment`,
                    !/<script/i.test(stripCommentsLikeABrowser(out)), JSON.stringify(out.slice(0, 160)));
            }
        }
    }
    // Positive control: an ordinary comment is emitted unchanged, so the check above isn't vacuous.
    const ok = String((await OfficeGenerator.generate(astWith([{ type: 'comment', text: ' plain ', metadata: { sourceSyntax: 'html' } }]), 'md')).value);
    // (astWith carries a title, so the output starts with frontmatter; the comment is the body.)
    check('md: an ordinary source comment is emitted verbatim', ok.endsWith('\n\n<!-- plain -->'), JSON.stringify(ok));
}

async function iframePreservationTests() {
    console.log('- HtmlParser iframe preservation (opt-in)...');
    const parseHtml = (html: string, extra: any = {}) =>
        OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html', ...extra } as any);
    const toHtml = async (ast: any) => String((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: false } } as any)).value);

    // Default: a non-YouTube iframe is dropped entirely (the standing security posture).
    const offHtml = await toHtml(await parseHtml('<iframe src="https://example.com/x"></iframe>'));
    check('iframe: dropped by default', !/<iframe/i.test(offHtml), offHtml.slice(0, 160));

    // Opted in: a legitimate https iframe survives to HTML and Markdown.
    const onAst = await parseHtml('<iframe src="https://player.example.com/v/1" width="640" height="360"></iframe>', { htmlParserConfig: { preserveIframes: true } });
    check('iframe: preserved when opted in (html)', /<iframe src="https:\/\/player\.example\.com\/v\/1"/.test(await toHtml(onAst)), 'not preserved');
    check('iframe: preserved when opted in (md)', /<iframe src="https:\/\/player\.example\.com/.test(String((await OfficeGenerator.generate(onAst, 'md')).value)), 'not in md');

    // A query-string src must survive with its & escaped exactly once, not compounding every cycle.
    const qsHtml = await toHtml(await parseHtml('<iframe src="https://player.example.com/v?a=1&amp;b=2"></iframe>', { htmlParserConfig: { preserveIframes: true } }));
    check('iframe: query-string src escaped once, not double-escaped', qsHtml.includes('a=1&amp;b=2') && !qsHtml.includes('&amp;amp;'), qsHtml.slice(0, 200));

    // Hostile schemes: even with preservation on, the src must not survive generation.
    for (const badSrc of ['javascript:alert(1)', 'data:text/html,alert(1)']) {
        const badAst = await parseHtml(`<iframe src="${badSrc}"></iframe>`, { htmlParserConfig: { preserveIframes: true } });
        const badHtml = await toHtml(badAst);
        check(`iframe: hostile src ${JSON.stringify(badSrc.slice(0, 16))} yields no live iframe (html)`, !/<iframe/i.test(badHtml) && !/javascript:|data:text\/html/i.test(badHtml), badHtml.slice(0, 200));
        const badMd = String((await OfficeGenerator.generate(badAst, 'md')).value);
        check(`iframe: hostile src ${JSON.stringify(badSrc.slice(0, 16))} yields no live iframe (md)`, !/javascript:|data:text\/html/i.test(badMd), badMd.slice(0, 200));
    }

    // Allowlist: only listed hosts survive.
    const allowAst = await parseHtml('<iframe src="https://player.vimeo.com/video/1"></iframe><iframe src="https://evil.example/x"></iframe>', { htmlParserConfig: { preserveIframes: ['vimeo.com'] } });
    const allowHtml = await toHtml(allowAst);
    check('iframe allowlist: listed host kept', allowHtml.includes('player.vimeo.com'), allowHtml.slice(0, 200));
    check('iframe allowlist: unlisted host dropped', !allowHtml.includes('evil.example'), allowHtml.slice(0, 200));

    // The allowlist must not be fooled by lookalike hosts: a suffix that isn't a dot-boundary,
    // a host that merely contains the entry, or the entry smuggled into the userinfo.
    for (const badHost of ['vimeo.com.evil.com', 'evilvimeo.com', 'notvimeo.com', 'vimeo.com@evil.com']) {
        const bypassHtml = await toHtml(await parseHtml(`<iframe src="https://${badHost}/x"></iframe>`, { htmlParserConfig: { preserveIframes: ['vimeo.com'] } }));
        check(`iframe allowlist: lookalike host ${JSON.stringify(badHost)} rejected`, !/<iframe/i.test(bypassHtml), bypassHtml.slice(0, 200));
    }

    // The preserved-iframe src must not compound its entity-escaping across Markdown save/reload
    // cycles: a `&amp;` in a query string must stay `&amp;`, not grow into `&amp;amp;` each time.
    const mdCfg: any = { fileType: 'md', htmlParserConfig: { preserveIframes: true } };
    let cyc = String((await OfficeGenerator.generate(await parseHtml('<iframe src="https://x.co/v?a=1&amp;b=2"></iframe>', { htmlParserConfig: { preserveIframes: true } }), 'md')).value);
    for (let i = 0; i < 2; i++) cyc = String((await OfficeGenerator.generate(await OfficeParser.parseOffice(Buffer.from(cyc), mdCfg), 'md')).value);
    check('iframe: src does not compound &amp; over md save/reload cycles', cyc.includes('a=1&amp;b=2') && !cyc.includes('&amp;amp;'), cyc.slice(0, 200));

    // HTML with its optional end tags omitted (`<p>`, `<li>`, `<tr>`, `<td>` left open) nests
    // nothing, as a browser reads it: thousands of them parse, quickly. Nesting that is really that
    // deep is refused with the typed error before anything can run out of stack on it.
    for (const [label, html, count] of [
        ['paragraphs', '<p>x'.repeat(5000), (ast: any) => ast.content.filter((n: any) => n.type === 'paragraph').length === 5000],
        ['list items', `<ul>${'<li>x'.repeat(5000)}</ul>`, (ast: any) => ast.content.filter((n: any) => n.type === 'list').length === 5000],
        ['table rows and cells', `<table>${'<tr><td>x<td>y'.repeat(5000)}</table>`, (ast: any) => ast.content[0]?.children?.length === 5000],
    ] as const) {
        const started = Date.now();
        let error = '';
        const ast = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html', ...QUIET } as any).catch((e: any) => { error = e.message; return null; });
        check(`html: 5000 ${label} with omitted end tags parse flat`, !!ast && count(ast) && Date.now() - started < 5000, `${Date.now() - started}ms ${error}`);
    }
    const deep = await rejectionMessage(() => OfficeParser.parseOffice(Buffer.from('<div>'.repeat(30000)), { fileType: 'html', ...QUIET } as any));
    check('html: 30000 nested elements reject with the nesting-depth error', /nesting depth/i.test(deep), deep);
}

async function markdownTests() {
    // `<!--` scanning is linear: an unclosed opener must not re-scan the rest of the document. Each of
    // these took seconds (80k lines: about 29 s) when a lazy regex body searched ahead per opener.
    for (const [label, md] of [
        ['80k unclosed <!-- lines', '<!-- x\n'.repeat(80000)],
        ['20k unclosed <!-- in one line', '<!-- x '.repeat(20000)],
        ['20k unclosed <!-- in one paragraph', 'a <!-- x\n'.repeat(20000)],
        // Display math inside a paragraph: a `$$` closing on a later line rejoins those lines, and an
        // unclosed one is literal. Neither may scan the rest of the paragraph once per opener.
        ['20k unclosed $$ in one paragraph', 'a $$x\n'.repeat(20000) + 'b'],
        ['20k $$ pairs across lines of one paragraph', 'a $$x\ny$$ b\n'.repeat(20000)],
        ['20k unclosed $$ in one line', 'a $$ x '.repeat(20000) + '$'],
        ['1MB display formula', `a $$${'x\\$'.repeat(250000)}$$ b`],
    ] as const) {
        const started = Date.now();
        await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' } as any);
        const ms = Date.now() - started;
        check(`markdown: ${label} parses in linear time`, ms < 3000, `${ms}ms`);
    }
    // Literal text that looks like an internal placeholder is text, never a lifted block (it used to crash).
    const lookalike = await OfficeParser.parseOffice(Buffer.from('__CODE_BLOCK_0__\n\n__HTML_COMMENT_0__\n\n__MATH_BLOCK_9__'), { fileType: 'md' } as any);
    check('markdown: placeholder look-alikes stay text', lookalike.content.every(n => n.type === 'paragraph'), JSON.stringify(lookalike.content.map(n => n.type)));

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

    // An image that is a link takes the same URL policy as a linked run in every format, and its link
    // title cannot leave its attribute or quotes.
    {
        const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
        const linked = { ...astWith([{ type: 'paragraph', children: [
            { type: 'image', metadata: { attachmentName: 'p.png', altText: 'a', link: 'javascript:alert(1)', linkType: 'external', linkTitle: `" onmouseover="alert(1)` } },
            { type: 'image', metadata: { url: 'https://example.com/i.png', altText: 'b', link: 'javascript:alert(2)', linkType: 'external' } },
        ] } as any]), attachments: [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: png }] } as any;
        for (const format of ['html', 'md', 'docx', 'odt', 'rtf', 'tex'] as const) {
            const out = (await OfficeGenerator.generate(linked, format, { onWarning: () => { } } as any)).value;
            const text = typeof out === 'string' ? out : Object.values(unzipSync(out as Uint8Array)).map(b => strFromU8(b)).join('\n');
            check(`image link: a javascript: target is refused (${format})`, !/javascript:/i.test(text), text.slice(0, 300));
            if (format === 'html') check('image link: its title stays in its attribute (html)', !/"\s*onmouseover=/.test(text), text.slice(0, 300));
        }
    }

    // An embed's label is link text or a directive label, from whatever source: escaped in every
    // embed form, so it cannot write a tag.
    for (const [embeds, embed] of [
        ['link', { embedType: 'youtube', videoId: 'abc', label: PAYLOAD }], ['thumbnail', { embedType: 'youtube', videoId: 'abc', label: PAYLOAD }],
        ['directive', { embedType: 'youtube', videoId: 'abc', label: PAYLOAD }], ['link', { embedType: 'iframe', url: 'https://example.com/e', label: PAYLOAD }],
    ] as const) {
        const out = (await OfficeGenerator.generate(astWith([{ type: 'embed', metadata: embed } as any]), 'md', { mdConfig: { dialect: { extends: 'extended', embeds } } } as any)).value as string;
        check(`md: an embed label cannot carry a raw tag (${embeds}, ${embed.embedType})`, !out.includes(PAYLOAD), out);
    }

    // Inline syntax that never closes costs time in proportion to its length: a line of backticks or
    // of `[x](` took minutes, rescanned from every opening character. So did an opener right after an
    // escaped backtick, a run of comments, and underscores that cannot close.
    for (const unit of ['`', 'a`', '``a', '[x](', '[x](y', '[', '![', '[^', '[[', '[a][', '<span style="x">', '<u>', '<sub>', '<sup>', '**a', '*a',
        '`\\`', '\\``x', '<!-- a -->', ' _x', '_a_b ', '***a', '[a](b "(', '[a \\]', '[[[[a', '\\_']) {
        const text = unit.repeat(Math.ceil(200_000 / unit.length));
        const started = Date.now();
        let error = '';
        await OfficeParser.parseOffice(Buffer.from('p ' + text + ' q'), { fileType: 'md' } as any).catch((e: any) => { error = e.message; });
        const ms = Date.now() - started;
        check(`md: 200 KB of ${JSON.stringify(unit)} parses in linear time`, !error && ms < 5000, `${ms}ms ${error}`);
    }
    // Blocks that never close (fences, in a list item or not, MDX components, `:::` admonitions, an
    // HTML table) are not rescanned for each opener, nor is a heading holding many unclosed `{#`.
    const decreasing = (indent: string) => { let s = '', k = 600; while (s.length < 200_000 && k >= 3) s += `${indent}${'`'.repeat(k--)}\n`; return s; };
    for (const [label, text] of [
        ['top-level fences', decreasing('')], ['fences in a list item', `- item\n\n${decreasing('    ')}`], ['table rows with escaped pipes', `| a | b |\n| - | - |\n${'| x\\|y | `p\\|q` |\n'.repeat(8000)}`],
        ['unclosed MDX components', '<A>x'.repeat(50_000)], ['MDX tags without a >', '<A '.repeat(66_667)], ['unclosed ::: admonitions', ':::note\n'.repeat(25_000)],
        ['an unclosed HTML table', `<table>${'<tr><td>x'.repeat(22_000)}`], ['a heading of unclosed {#', `# a${' {#x'.repeat(50_000)}`],
    ] as const) {
        const started = Date.now();
        await OfficeParser.parseOffice(Buffer.from(text), { fileType: 'md' } as any);
        check(`md: 200 KB of ${label} parses in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
    // A long run of one character inside a construct (a line of spaces in a link target, a table, an
    // aligned div, a definition, a heading) is scanned once. Patterns anchored to the end of a line or
    // holding runs that can match the same characters used to retry every position of such a run:
    // quadratic, and for a table's delimiter row cubic (5,000 spaces took 81 seconds).
    const run = ' '.repeat(200_000);
    for (const [label, text] of [
        ['a paragraph line', `a${run}x\nb`], ['a link target', `[a](u${run}x)`], ['a link title', `[a](u "${run}x)`],
        ['a table and a line of spaces', `a|b\n${run}x`], ['a table attribute line', `| a |\n| - |\n{${run}x`],
        ['an aligned div', `<div align="center">a${run}x</div>`], ['an iframe line', `<iframe${run}x`], ['a definition', `T\n: a${run}x`],
        ['a heading', `# a${run}x {#i}`], ['a line of backslashes in a table row', `| a |\n| - |\n| ${'\\'.repeat(200_000)}x`],
    ] as const) {
        const started = Date.now();
        await OfficeParser.parseOffice(Buffer.from(text), { fileType: 'md' } as any);
        check(`md: 200 KB run in ${label} parses in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
    // The same for the generators: text holding a long run, in a paragraph, emphasis, a list item, a
    // cell and a code block, written as Markdown, text and LaTeX.
    for (const [label, content] of [
        ['a paragraph', [{ type: 'paragraph', children: [{ type: 'text', text: `a${run}x ` }] }]],
        ['bold text', [{ type: 'paragraph', children: [{ type: 'text', text: `a${run}x`, formatting: { bold: true } }] }]],
        ['a list item', [{ type: 'list', metadata: { listType: 'unordered', listId: 'l', indentation: 0, itemIndex: 0 }, children: [{ type: 'text', text: `a${run}x\n${'\\'.repeat(1000)}` }] }]],
        ['a table cell', [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: `a${run}x` }] }] }] }]],
        ['a code block of blank lines', [{ type: 'code', text: `a${'\n'.repeat(200_000)}x` }, { type: 'paragraph', children: [{ type: 'text', text: 'n', notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [{ type: 'text', text: 'f' }] }] }] }]],
    ] as const) {
        for (const format of ['md', 'text', 'tex'] as const) {
            const started = Date.now();
            await OfficeGenerator.generate(astWith(content as any), format, { onWarning: () => { } } as any);
            check(`${format}: 200 KB run in ${label} is written in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
        }
    }
    // A line of hundreds of thousands of inline nodes is appended, not spread into one call.
    let manyNodes = '';
    await OfficeParser.parseOffice(Buffer.from('*a'.repeat(500_000)), { fileType: 'md' } as any).catch((e: any) => { manyNodes = e.message; });
    check('md: a line of 500,000 emphasis runs parses', !manyNodes, manyNodes);
    // Lines that each open something no later line closes (a reference, footnote or abbreviation
    // label, an anchor or a component tag) are each scanned a bounded distance. A label or tag pattern
    // that can run across lines scanned from every such line to the end of the document.
    for (const line of ['[a', '[^a', '*[a', '<a a', '<Aa<A', '<a'] as const) {
        const started = Date.now();
        await OfficeParser.parseOffice(Buffer.from(`${line}\n`.repeat(40_000)), { fileType: 'md' } as any);
        check(`md: 40,000 lines of ${JSON.stringify(line)} parse in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
    // The patterns read since: link targets holding parenthesis pairs (a pattern that let a group match
    // nothing backtracked exponentially on one Wikipedia URL), emphasis, strikethrough and highlight
    // openers that never close, code spans across lines, thematic breaks, setext underlines, empty
    // items, nested quotes, list continuations, definitions and footnotes continued across blank lines.
    for (const [label, text] of [
        ['nested link parentheses', `[a](${'(b'.repeat(20_000)})`], ['unclosed link parentheses', `[a](x${'(b'.repeat(40_000)}`],
        ['closed link parentheses', `[a](b${'(c)'.repeat(40_000)}`], ['unclosed emphasis', '*a '.repeat(60_000)], ['unclosed bold', '**a '.repeat(50_000)],
        ['escaped closers', '*a\\*'.repeat(40_000)], ['unclosed strikethrough', '~~a '.repeat(50_000)], ['unclosed highlights', '==a '.repeat(50_000)],
        ['code spans across lines', 'x `a\n'.repeat(40_000)], ['thematic breaks', '* * *\n'.repeat(40_000)], ['setext underlines', 'text\n-\n'.repeat(40_000)],
        ['empty list items', '-\n'.repeat(80_000)], ['nested quote markers', `> ${'>'.repeat(200_000)} x`], ['list continuation lines', `- a\n${'  b\n'.repeat(50_000)}`],
        ['a definition list', 't\n: d\n'.repeat(40_000)], ['a footnote continued across blank lines', `a[^1]\n\n[^1]: a${'\n\n    b'.repeat(20_000)}`],
        ['footnotes across blank lines', 'x[^1]\n\n' + Array.from({ length: 10_000 }, (_, i) => `[^${i}]: a\n\n    b\n`).join('\n')],
        ['a quote of blocks', `> - a\n${'> - b\n'.repeat(40_000)}`],
        ['empty anchors in a line', 'a <a id="x"></a>'.repeat(60_000)], ['unclosed anchors in a line', '<a id="x" '.repeat(60_000)],
        ['an HTML table blank lines run through', `<table>\n${'<tr><td>a</td></tr>\n\n'.repeat(20_000)}</table>`],
        ['HTML tables opened and never closed, between blank lines', '<table>\n\n'.repeat(20_000)],
        ['HTML tables with Markdown in their cells', '<table><tr><td>\n\n*x* y\n\n</td></tr></table>\n\n'.repeat(10_000)],
        ['blocks under a list of many indents', `${Array.from({ length: 5_000 }, (_, i) => `${' '.repeat(i % 400)}- item`).join('\n')}${'\n\n    more'.repeat(20_000)}`],
        ['indented code across blank lines', '    a\n\n'.repeat(60_000)], ['lazy underlines in a quote', '> q\n===\n'.repeat(40_000)],
    ] as const) {
        const started = Date.now();
        await OfficeParser.parseOffice(Buffer.from(text), { fileType: 'md' } as any);
        check(`md: ${label} parse in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
    // A sheet's cell coordinates cannot make a grid no memory holds: one cell at Excel's last position,
    // cells along a diagonal (each row and column its own), or a span of a billion columns. Every
    // writer fills the grid between cells, and HTML draws all of it.
    const T0 = (text: string) => ({ type: 'text', text });
    const sheetOf = (rows: any[]) => ({ type: 'xlsx', metadata: {}, attachments: [], content: [{ type: 'sheet', metadata: { sheetName: 'S' }, children: rows }] } as any);
    for (const [label, ast] of [
        ['a cell at XFD1048576', sheetOf([{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [T0('a')] }] }, { type: 'row', children: [{ type: 'cell', metadata: { row: 1_048_575, col: 16_383 }, children: [T0('far')] }] }])],
        ['5,000 cells along a diagonal', sheetOf(Array.from({ length: 5_000 }, (_, i) => ({ type: 'row', children: [{ type: 'cell', metadata: { row: i, col: i }, children: [T0('d' + i)] }] })))],
        ['a span of a billion columns', sheetOf([{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0, colSpan: 1e9, rowSpan: 1e9 }, children: [T0('wide')] }] }])],
    ] as const) {
        for (const format of ['html', 'csv', 'md', 'tex', 'docx', 'odt', 'rtf', 'text', 'epub'] as const) {
            const started = Date.now();
            const out: any = (await OfficeGenerator.generate(ast, format, { onWarning: () => {} } as any)).value;
            const size = typeof out === 'string' ? out.length : out.length ?? out.byteLength;
            check(`${format}: ${label} is written bounded`, Date.now() - started < 5000 && size < 5_000_000, `${Date.now() - started}ms, ${size} bytes`);
        }
    }
    // Spans on cells without coordinates count too (an 80-byte HTML table spanning 16 million columns
    // ran HTML and PDF out of memory), the budget is the document's (ten sheets each just within one
    // budget of their own made a 4 KB workbook's HTML too large to write), and HTML spans are clamped
    // as a browser clamps them.
    const spanned = await OfficeParser.parseOffice(Buffer.from('<table><tr><td colspan="16000000" rowspan="2">x</td></tr><tr><td>y</td></tr></table>'), { fileType: 'html' } as any);
    check('html: colspan and rowspan are clamped as a browser clamps them', (spanned.content[0].children![0].children![0].metadata as any).colSpan === 1000);
    const handSpan = { type: 'docx', metadata: {}, attachments: [], content: [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', metadata: { colSpan: 1e8, rowSpan: 1e8 }, children: [T0('x')] }] }, { type: 'row', children: [{ type: 'cell', children: [T0('y')] }] }] }] } as any;
    const manySheets = { type: 'xlsx', metadata: {}, attachments: [], content: Array.from({ length: 20 }, () => sheetOf([{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [T0('a')] }] }, { type: 'row', children: [{ type: 'cell', metadata: { row: 1999, col: 999 }, children: [T0('b')] }] }]).content[0]) } as any;
    for (const [label, ast] of [['a span of 100 million on a cell without coordinates', handSpan], ['20 sheets each holding 2 million positions', manySheets]] as const) {
        for (const format of ['html', 'md', 'csv', 'tex', 'docx', 'odt', 'epub', 'pdf'] as const) {
            const started = Date.now();
            const out: any = (await OfficeGenerator.generate(ast, format, { onWarning: () => {}, pdfConfig: { engine: 'native' } } as any)).value;
            const size = typeof out === 'string' ? out.length : out.length ?? out.byteLength;
            check(`${format}: ${label} is written bounded`, Date.now() - started < 5000 && size < 20_000_000, `${Date.now() - started}ms, ${size} bytes`);
        }
    }
    // A value of another type than the AST defines (an array for a string, text for a number) cannot
    // get past a writer's escaping: every one is coerced before the writers run.
    const P0 = '1" onx="PWN\' <PWN>';
    const confused = { type: 'docx', metadata: { title: [P0] }, attachments: [], content: [
        { type: 'heading', metadata: { level: P0, anchorIds: [[P0]] }, children: [T0('h')] },
        { type: 'list', metadata: { listType: 'ordered', listId: 'l', indentation: P0, itemIndex: P0 }, children: [T0('li')] },
        { type: 'paragraph', children: [T0('i'), { type: 'image', metadata: { url: 'https://x/a.png', altText: [P0], title: [P0], width: [P0] } }] },
        { type: 'code', text: 'x', metadata: { math: [P0], language: [P0] } },
        { type: 'embed', metadata: { embedType: 'youtube', videoId: [P0], label: [P0] } },
        { type: 'admonition', metadata: { admonitionType: [P0] }, children: [{ type: 'paragraph', children: [T0('a')] }] },
        { type: 'paragraph', children: [{ type: 'text', text: 'c', formatting: { color: 'red&#59position:fixed' } }] },
    ] } as any;
    for (const format of ['html', 'epub', 'md', 'rtf'] as const) {
        const out: any = (await OfficeGenerator.generate(confused, format, { onWarning: () => {}, includeFormatting: true, htmlConfig: { standalone: format === 'epub' } } as any)).value;
        const text = typeof out === 'string' ? out : Object.values(unzipSync(new Uint8Array(out))).map((u: any) => new TextDecoder().decode(u)).join('\n');
        // Live means: a control word taking the value in RTF; elsewhere the value as a tag or attribute of
        // a tag, or a second CSS declaration. (Markdown syntax, such as alt text in `![...]` or a fence's
        // info string, is escaped by the renderer, so a fenced block is left out of the Markdown checked.)
        const html = format === 'md' ? text.replace(/^(`{3,})[^\n]*\n[\s\S]*?\n\1$/gm, '').replace(/^`{3,}[^\n]*$/gm, '') : text;
        const live = format === 'rtf' ? /(^|[^\\])\\s1" onx|\\ilvl1"/.test(text) : /<[a-zA-Z][^<>]*onx="PWN|<PWN>|;\s*position:fixed|&#59/.test(html);
        check(`${format}: values of the wrong type are escaped`, !live, text.slice(0, 300));
    }
    // Named anchors are ids of the node they stand at, each node's added once: many of them in a line,
    // or before a block, cost time in proportion to their number, read and written.
    for (const [label, text, fileType] of [
        ['md: 60,000 anchors in a line', 'a <a id="x1"></a>'.repeat(60_000), 'md'],
        ['html: 60,000 anchors in a paragraph', `<p>${'a <a id="x"></a>'.repeat(60_000)}</p>`, 'html'],
        ['html: 60,000 anchors before a heading', `${'<a id="x"></a>'.repeat(60_000)}<h2>h</h2>`, 'html'],
    ] as const) {
        const started = Date.now();
        const ast = await OfficeParser.parseOffice(Buffer.from(text), { fileType } as any);
        await ast.to('md');
        await ast.to('html');
        check(`${label} are read and written in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
    // Footnotes that refer to themselves, to each other, or along a long chain: a note referring to
    // itself recursed until the stack ran out, and a chain made an AST too deep to serialize. A note in
    // a note refers to none; other references stay text.
    for (const [label, text] of [
        ['a footnote referring to itself', 'a[^1]\n\n[^1]: see [^1]'], ['two footnotes referring to each other', 'a[^1]\n\n[^1]: one [^2]\n[^2]: two [^1]'],
        ['a chain of 20,000 footnotes', `x[^1]\n\n${Array.from({ length: 20_000 }, (_, i) => `[^${i + 1}]: n [^${i + 2}]`).join('\n')}`],
    ] as const) {
        const started = Date.now();
        const outcome = await OfficeParser.parseOffice(Buffer.from(text), { fileType: 'md' } as any).then(ast => { JSON.stringify(ast); return 'ok'; }, (e: any) => e.message);
        check(`md: ${label} parses, serializes, in linear time`, outcome === 'ok' && Date.now() - started < 5000, `${outcome} ${Date.now() - started}ms`);
    }
    // A table cell's alignment wrappers are found once each, not searched for a closing tag from each.
    const alignDivs = `| a |\n| - |\n| ${'<div style="text-align: left">x'.repeat(80_000)} |`;
    const alignStarted = Date.now();
    await OfficeParser.parseOffice(Buffer.from(alignDivs), { fileType: 'md' } as any);
    check('md: a cell of 80,000 unclosed alignment wrappers parses in linear time', Date.now() - alignStarted < 5000, `${Date.now() - alignStarted}ms`);
    // Writing many blocks, or a paragraph of many runs, never reads back all it has written: reading
    // the end of a string grown by appending makes V8 copy the whole of it, which made the Markdown
    // writer quadratic in the number of blocks (48,000 paragraphs took 3 seconds).
    const many = 100_000;
    for (const [label, content] of [
        ['paragraphs', Array.from({ length: many }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `p${i}` }] }))],
        ['list items', Array.from({ length: many }, (_, i) => ({ type: 'list', metadata: { listType: 'unordered', listId: 'l', indentation: 0, itemIndex: i }, children: [{ type: 'text', text: `i${i}` }] }))],
        ['runs in one paragraph', [{ type: 'paragraph', children: Array.from({ length: many }, (_, i) => ({ type: 'text', text: `r${i} `, formatting: i % 2 ? { bold: true } : { italic: true } })) }]],
        ['paragraphs on one page', [{ type: 'page', children: Array.from({ length: many }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `q${i}` }] })) }]],
    ] as const) {
        for (const format of ['md', 'text', 'html', 'tex'] as const) {
            const started = Date.now();
            await OfficeGenerator.generate(astWith(content as any), format, { onWarning: () => { } } as any);
            check(`${format}: 100,000 ${label} are written in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
        }
    }
    // A colour table found each run's colour by scanning it.
    const colours = astWith([{ type: 'paragraph', children: Array.from({ length: many }, (_, i) => ({ type: 'text', text: `r${i} `, formatting: { color: `#${i.toString(16).padStart(6, '0')}` } })) }] as any);
    const coloursStarted = Date.now();
    await OfficeGenerator.generate(colours, 'rtf', { onWarning: () => { } } as any);
    check('rtf: 100,000 runs of distinct colours are written in linear time', Date.now() - coloursStarted < 5000, `${Date.now() - coloursStarted}ms`);
    // The same for text built while reading or writing: an ODF table cell of many paragraphs, and LaTeX
    // output of a paragraph of many inline comments or a math block of many display environments,
    // each tested the end of all the text so far at every step.
    const cellOdt = await OfficeGenerator.generate(astWith([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: Array.from({ length: many }, (_, i) => ({ type: 'paragraph', children: [{ type: 'text', text: `c${i}` }] })) }] }] }] as any), 'odt');
    const cellStarted = Date.now();
    await OfficeParser.parseOffice(Buffer.from(cellOdt.value as any), { fileType: 'odt' } as any);
    check('odt: a table cell of 100,000 paragraphs parses in linear time', Date.now() - cellStarted < 5000, `${Date.now() - cellStarted}ms`);
    // ODT: a cell's inline runs are gathered into a paragraph once each, pictures without text too.
    const runsStarted = Date.now();
    const runsOdt = await OfficeGenerator.generate(astWith([{ type: 'table', children: [{ type: 'row', children: [
        { type: 'cell', children: Array.from({ length: many }, (_, i) => ({ type: 'text', text: `r${i} ` })) },
        { type: 'cell', children: Array.from({ length: many }, () => ({ type: 'break' })) },
    ] }] }] as any), 'odt', { onWarning: () => { } } as any);
    const runsXml = strFromU8(unzipSync(runsOdt.value as Uint8Array)['content.xml']);
    check('odt: a cell of 100,000 runs, and one of 100,000 breaks, are written in linear time, the runs one paragraph', Date.now() - runsStarted < 5000 && runsXml.includes('r0 r1 r2'), `${Date.now() - runsStarted}ms`);
    // A formula's commands are read once each: rewriting KaTeX macros (`\R`), and planning packages and
    // definitions for 100,000 distinct unknown commands.
    const unknownCommands = Array.from({ length: many }, (_, i) => `\\x${i.toString(26).replace(/\d/g, d => 'abcdefghij'[+d])}`).join(' ');
    for (const [label, markdown] of [['100,000 inline comments', 'a<!--x-->'.repeat(many)], ['a math block of 100,000 display environments', `$$\nx ${'\\begin{equation}y\\end{equation}'.repeat(many)}\n$$`],
        ['a formula of 100,000 KaTeX macros', `$${'\\R \\lang '.repeat(many)}$`], ['a formula of 100,000 unknown commands', `$${unknownCommands}$`]] as const) {
        const parsed = await OfficeParser.parseOffice(Buffer.from(markdown), { fileType: 'md' } as any);
        const started = Date.now();
        await OfficeGenerator.generate(parsed, 'tex', { onWarning: () => { } } as any);
        check(`tex: ${label} are written in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    }
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

    // Parsed attribute values are decoded, so the checks above see what a browser would: an encoded
    // scheme or quote is caught, not smuggled through as inert-looking text.
    const parsed = async (html: string, fmt: 'html' | 'md') => (await (await OfficeParser.parseOffice(Buffer.from(html),
        { fileType: 'html', htmlParserConfig: { preserveAttributes: true } } as any)).to(fmt, { htmlConfig: { standalone: false } } as any)).value as string;
    // Attribute names of every tag, with quoted values consumed whole (text inside a value is not a name).
    const attributeNames = (html: string) => [...html.matchAll(/<[a-zA-Z][^\s>/]*((?:\s+[^\s"'>/=]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?)*)\s*\/?>/g)]
        .flatMap(tag => [...tag[1].matchAll(/([^\s"'>/=]+)(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+))?/g)].map(a => a[1].toLowerCase()));
    // A Markdown link target as an HTML5-aware renderer reads it: character references decoded
    // (`&colon;` included), then percent-decoded.
    const markdownTargets = (md: string) => [...md.matchAll(/\]\(([^)\s]*)/g)].map(m => {
        const decoded = m[1].replace(/&colon;/g, ':').replace(/&#(\d+);/g, (_x, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_x, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');
        try { return decodeURIComponent(decoded); } catch { return decoded; }
    });
    const encodedScheme = '<p><a href="java&#115;cript:alert(1)">x</a> <a href="jav&#x61;script&colon;alert(2)">y</a></p>';
    const encodedQuote = '<p data-note="a&quot; onclick=&quot;alert(1)"><a href="http://x.com/&quot; onmouseover=&quot;alert(1)">z</a> <img src="i.png" alt="&quot; onerror=&quot;alert(1)"></p>';
    for (const source of [encodedScheme, encodedQuote]) {
        const html = await parsed(source, 'html');
        check('decoded attributes: no event handler and no javascript: link in HTML', !attributeNames(html).some(n => /^on/.test(n)) && ![...html.matchAll(/href="([^"]*)"/g)].some(m => /^\s*javascript:/i.test(m[1].replace(/&colon;/g, ':').replace(/&amp;/g, '&'))), html);
        const md = await parsed(source, 'md');
        check('decoded attributes: no Markdown link target reads as javascript:', !markdownTargets(md).some(t => /^\s*javascript:/i.test(t)), md);
    }

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

    // The newer config paths: texParserConfig, the value validation and the unknown-key check. A
    // prototype key is neither merged nor reported as an unknown option, and a huge invalid value
    // cannot make the warning message huge.
    const texRaw = JSON.parse('{"texParserConfig":{"today":"x","__proto__":{"pollutedParser":"YES"}}}');
    const texParser: any = resolveParserConfig(texRaw);
    check('config: texParserConfig merge applied', texParser.texParserConfig.today === 'x');
    check('config: __proto__ in texParserConfig cannot reach Object.prototype', ({} as any).pollutedParser === undefined);
    const issues: any[] = [];
    const genRaw = JSON.parse('{"__proto__":{"polluted":"YES"},"texConfig":{"documentClass":"article","__proto__":{"polluted":"YES"},"constructor":{"prototype":{"polluted":"YES"}}}}');
    resolveGeneratorConfig('tex' as any, undefined as any, { ...genRaw, onWarning: (i: any) => issues.push(i) } as any);
    check('config: prototype keys in a generator config reach nothing', ({} as any).polluted === undefined);
    check('config: prototype keys are not reported as unknown options', !issues.some(i => /__proto__|constructor/.test(i.message)), issues.map(i => i.message).join(' | '));
    const huge: any[] = [];
    resolveGeneratorConfig('tex' as any, undefined as any, { texConfig: { documentClass: 'x'.repeat(1_000_000) }, onWarning: (i: any) => huge.push(i) } as any);
    check('config: an invalid value is shown truncated in its warning', huge.length === 1 && huge[0].message.length < 400, `${huge[0]?.message.length}`);
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

    // A comment's author and initials, and a header's text, are escaped in their destinations: braces
    // and backslashes in them cannot close the destination or write body text.
    const annotated = String((await OfficeGenerator.generate({ ...astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'body', comments: [{ type: 'comment', text: 'c}\\par INJECTED{', metadata: { author: 'A}\\par INJECTED{', initials: 'I}{\\b', date: '}{\\par' } }] }] } as any]), auxiliary: { headers: [{ type: 'paragraph', children: [{ type: 'text', text: 'h}\\par INJECTED{' }] }] } } as any, 'rtf')).value);
    const annotatedBack = await OfficeParser.parseOffice(Buffer.from(annotated, 'latin1'), { fileType: 'rtf' } as any);
    const annotatedComment = (annotatedBack.content[0]?.children ?? []).flatMap((n: any) => n.comments ?? [])[0];
    check('rtf: a comment\'s author, initials and text and a header\'s text stay in their destinations', annotatedBack.content.length === 1 && annotatedBack.content[0].text === 'body' && annotatedComment?.metadata?.author === 'A}\\par INJECTED{' && annotatedComment?.metadata?.initials === 'I}{\\b' && annotatedComment?.text === 'c}\\par INJECTED{' && annotatedBack.auxiliary?.headers?.[0]?.text === 'h}\\par INJECTED{', annotated);

    // Positive control. Without these the allowlist could be "reject everything" and still pass.
    for (const url of ['https://example.com/a?b=1', 'http://x.test/p', 'mailto:a@b.test', 'tel:+123', '#anchor', 'relative/path.html']) {
        const rtf = await rtfFor(url);
        check(`rtf: ${url} still emits a HYPERLINK field`, /HYPERLINK "/.test(rtf),
            `legitimate URL was dropped: ${rtf.slice(0, 140)}`);
    }

    // A protocol-relative URL keeps working on the web as https (older HTML sources use it), and cannot
    // resolve to file://host from a page opened from disk; RTF refuses it outright.
    check('html: a protocol-relative URL is written as https',
        sanitizeUrl('//example.com/x') === 'https://example.com/x',
        `sanitizeUrl gave ${JSON.stringify(sanitizeUrl('//example.com/x'))}`);

    // Field metacharacters must still be neutralized on an otherwise-allowed URL.
    const quoted = await rtfFor('https://example.com/a"}{\\b evil');
    check('rtf: quotes/braces in an allowed URL are escaped',
        !/HYPERLINK "[^"]*"\}\{\\b/.test(quoted), quoted.match(/HYPERLINK "[^"]*"/)?.[0] ?? '');
}

/**
 * The DOCX generator hand-writes WordprocessingML into a ZIP package, so every value that reaches
 * the XML is an injection surface: text, hyperlink targets, internal-link anchors, bookmark names
 * and colors. These assert the package stays well-formed and script-free under hostile input while
 * legitimate content survives (degrade, don't delete).
 */
async function docxSanitizationTests() {
    console.log('- DocxGenerator (WML injection surfaces)...');

    const docFor = async (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['word/document.xml']);
    const relsFor = async (bytes: Uint8Array) => {
        const f = unzipSync(bytes)['word/_rels/document.xml.rels'];
        return f ? strFromU8(f) : '';
    };
    const gen = async (content: any[]) => (await OfficeGenerator.generate(astWith(content), 'docx')).value as Uint8Array;

    // ── Hyperlink target scheme rejection ─────────────────────────────────────
    const linkContent = (url: string) => [{ type: 'paragraph', children: [
        { type: 'text', text: 'clickme', metadata: { link: url, linkType: 'external' } }] }];
    for (const url of ['javascript:alert(1)', 'vbscript:msgbox(1)', 'data:text/html,<script>',
                       'file:///C:/Windows/System32/calc.exe', '\\\\evil.com\\share\\x', '//evil.com/share']) {
        const bytes = await gen(linkContent(url));
        const doc = await docFor(bytes);
        const rels = await relsFor(bytes);
        check(`docx: ${JSON.stringify(url).slice(0, 32)} emits no hyperlink relationship`,
            !/<w:hyperlink r:id=/.test(doc) && !/TargetMode="External"/.test(rels),
            doc.match(/<w:hyperlink[^>]*>/)?.[0] ?? rels.slice(0, 120));
        check('docx: rejected link keeps its text', doc.includes('clickme'),
            `link text was dropped along with the URL: ${doc.slice(0, 160)}`);
    }
    // Positive controls: without these the allowlist could be "reject everything" and still pass.
    for (const url of ['https://example.com/a?b=1', 'http://x.test/p', 'mailto:a@b.test', 'tel:+123']) {
        const bytes = await gen(linkContent(url));
        const doc = await docFor(bytes);
        const rels = await relsFor(bytes);
        check(`docx: ${url} still emits a hyperlink relationship`,
            /<w:hyperlink r:id=/.test(doc) && /TargetMode="External"/.test(rels),
            `legitimate URL was dropped: ${doc.slice(0, 160)}`);
    }

    // ── Text / attribute XML injection ────────────────────────────────────────
    const payload = `</w:t></w:r></w:p><script>alert(1)</script>&<>"'`;
    {
        const bytes = await gen([{ type: 'paragraph', children: [{ type: 'text', text: payload }] }]);
        const doc = await docFor(bytes);
        check('docx: hostile text does not break out of <w:t>', !/<script>/.test(doc),
            doc.match(/<script>[^<]*/)?.[0] ?? '');
        // Every emitted XML part must remain well-formed (strict @xmldom throws on malformed input).
        for (const [name, data] of Object.entries(unzipSync(bytes))) {
            if (!name.endsWith('.xml') && !name.endsWith('.rels')) continue;
            let ok = true;
            try { parseXmlString(strFromU8(data as Uint8Array)); } catch { ok = false; }
            check(`docx: ${name} well-formed under hostile text`, ok);
        }
        // The literal characters survive a re-parse (escaped, not executed).
        const back = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType: 'docx' });
        const allText = JSON.stringify(back.content);
        check('docx: payload text preserved through round-trip', allText.includes('alert(1)'),
            `payload text was lost: ${allText.slice(0, 200)}`);
    }

    // ── Control-character stripping ───────────────────────────────────────────
    {
        const bytes = await gen([{ type: 'paragraph', children: [{ type: 'text', text: 'a\x00\x01\x1F\x08b' }] }]);
        const doc = await docFor(bytes);
        check('docx: invalid XML control chars stripped', !/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(doc),
            'control characters leaked into document.xml');
        let ok = true; try { parseXmlString(doc); } catch { ok = false; }
        check('docx: document.xml well-formed after control-char strip', ok);
    }

    // ── Internal-link anchor / bookmark name sanitization ─────────────────────
    {
        const bytes = await gen([
            { type: 'heading', metadata: { level: 1, id: 'evil" name<>&/\\' }, children: [{ type: 'text', text: 'Title' }] },
            { type: 'paragraph', children: [{ type: 'text', text: 'jump', metadata: { link: '#evil" name<>&/\\', linkType: 'internal' } }] },
        ]);
        const doc = await docFor(bytes);
        const anchors = [...doc.matchAll(/w:(?:anchor|name)="([^"]*)"/g)].map(m => m[1]);
        check('docx: bookmark/anchor names contain only safe chars',
            anchors.length > 0 && anchors.every(a => /^[A-Za-z_][A-Za-z0-9_]*$/.test(a)),
            `unsafe anchor/name emitted: ${JSON.stringify(anchors)}`);
        let ok = true; try { parseXmlString(doc); } catch { ok = false; }
        check('docx: document.xml well-formed with hostile anchor', ok);
    }

    // ── Color validation ──────────────────────────────────────────────────────
    {
        const bytes = await gen([{ type: 'paragraph', children: [
            { type: 'text', text: 'colored', formatting: { color: 'red"/><w:color w:val="injected' } as any }] }]);
        const doc = await docFor(bytes);
        const colors = [...doc.matchAll(/<w:color w:val="([^"]*)"/g)].map(m => m[1]);
        check('docx: color values are valid hex or absent',
            colors.every(c => /^[0-9A-Fa-f]{6}$/.test(c) || c === 'auto'),
            `invalid color leaked: ${JSON.stringify(colors)}`);
    }

    // ── Attachment name cannot escape the media directory ─────────────────────
    {
        const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
        const ast = astWith([{ type: 'image', metadata: { attachmentName: '../../evil.png' } }]);
        (ast as any).attachments = [{ name: '../../evil.png', mimeType: 'image/png', data: png }];
        const bytes = (await OfficeGenerator.generate(ast, 'docx')).value as Uint8Array;
        const entries = Object.keys(unzipSync(bytes));
        check('docx: no zip entry escapes its directory', entries.every(e => !e.includes('..')),
            `traversal entry present: ${entries.filter(e => e.includes('..')).join(', ')}`);
        const media = entries.filter(e => e.startsWith('word/media/'));
        check('docx: media file names are synthesized, not attacker-controlled',
            media.every(e => /^word\/media\/image\d+\.[a-z]+$/.test(e)), `unexpected media entry: ${media.join(', ')}`);
    }

    // ── Absurd colSpan/rowSpan is clamped (no unbounded cell explosion) ────────
    {
        const ast = astWith([{ type: 'table', children: [
            { type: 'row', children: [{ type: 'cell', metadata: { colSpan: 1e9, rowSpan: 1e9 }, children: [{ type: 'text', text: 'x' }] }] } ] }]);
        const bytes = (await OfficeGenerator.generate(ast, 'docx')).value as Uint8Array;
        check('docx: absurd span does not blow up output size', bytes.length < 200_000, `output was ${bytes.length} bytes`);
        const doc = await docFor(bytes);
        let ok = true; try { parseXmlString(doc); } catch { ok = false; }
        check('docx: document.xml well-formed with clamped spans', ok);
        const gridSpan = doc.match(/<w:gridSpan w:val="(\d+)"/);
        check('docx: gridSpan clamped to a sane value', !gridSpan || Number(gridSpan[1]) <= 1000,
            `gridSpan was ${gridSpan?.[1]}`);
    }

    // ── styleMap output.tag cannot inject into w:pStyle ───────────────────────
    {
        const pAst = astWith([{ type: 'paragraph', metadata: { style: 'Custom' }, children: [{ type: 'text', text: 'hi' }] } as any]);
        const sm = (tag: string) => ({ styleMap: [{ selector: { nodeType: 'paragraph', attributes: { style: 'Custom' } }, output: { tag } }] } as any);
        const hostile = (await OfficeGenerator.generate(pAst, 'docx', sm('h1"/><w:pStyle w:val="Injected'))).value as Uint8Array;
        const hdoc = strFromU8(unzipSync(hostile)['word/document.xml']);
        check('docx: hostile styleMap tag is not mapped to a style', !/Injected/.test(hdoc), hdoc.match(/<w:pStyle[^>]*>/)?.[0] ?? '');
        let ok = true; try { parseXmlString(hdoc); } catch { ok = false; }
        check('docx: document.xml well-formed with hostile styleMap tag', ok);
        // Positive control: a legitimate tag still maps to its Word style.
        const benign = (await OfficeGenerator.generate(pAst, 'docx', sm('h1'))).value as Uint8Array;
        const bdoc = strFromU8(unzipSync(benign)['word/document.xml']);
        check('docx: legitimate styleMap tag h1 maps to Heading1', /<w:pStyle w:val="Heading1"\/>/.test(bdoc), bdoc.match(/<w:pStyle[^>]*>/)?.[0] ?? '');
    }
}

/**
 * The ODT generator hand-writes OpenDocument XML into a zip, so text, hyperlink hrefs, bookmark
 * names, colors, attachment names and styleMap tags are all injection surfaces. These assert the
 * package stays well-formed and script-free under hostile input while legitimate content survives.
 */
async function odtSanitizationTests() {
    console.log('- OdtGenerator (ODF injection surfaces)...');

    const contentFor = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)['content.xml']);
    const gen = async (content: any[]) => (await OfficeGenerator.generate(astWith(content), 'odt')).value as Uint8Array;
    const wellFormed = (xml: string) => { try { parseXmlString(xml); return true; } catch { return false; } };

    // ── Hyperlink scheme rejection ────────────────────────────────────────────
    const linkContent = (url: string) => [{ type: 'paragraph', children: [{ type: 'text', text: 'clickme', metadata: { link: url, linkType: 'external' } }] }];
    for (const url of ['javascript:alert(1)', 'vbscript:msgbox(1)', 'data:text/html,<script>', 'file:///etc/passwd', '\\\\evil.com\\share\\x', '//evil.com/share']) {
        const doc = contentFor(await gen(linkContent(url)));
        check(`odt: ${JSON.stringify(url).slice(0, 32)} emits no text:a`, !/<text:a /.test(doc), doc.match(/<text:a[^>]*>/)?.[0] ?? '');
        check('odt: rejected link keeps its text', doc.includes('clickme'), doc.slice(0, 160));
    }
    for (const url of ['https://example.com/a?b=1', 'http://x.test/p', 'mailto:a@b.test', 'tel:+123']) {
        const doc = contentFor(await gen(linkContent(url)));
        check(`odt: ${url} still emits a text:a`, /<text:a [^>]*xlink:href="/.test(doc), doc.slice(0, 160));
    }

    // ── Text / attribute XML injection ────────────────────────────────────────
    {
        const payload = `</text:p></office:text><script>alert(1)</script>&<>"'`;
        const bytes = await gen([{ type: 'paragraph', children: [{ type: 'text', text: payload }] }]);
        const doc = contentFor(bytes);
        check('odt: hostile text does not break out', !/<script>/.test(doc), doc.match(/<script>[^<]*/)?.[0] ?? '');
        for (const [name, data] of Object.entries(unzipSync(bytes))) {
            if (!name.endsWith('.xml')) continue;
            check(`odt: ${name} well-formed under hostile text`, wellFormed(strFromU8(data as Uint8Array)));
        }
        const back = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType: 'odt' });
        check('odt: payload text preserved through round-trip', JSON.stringify(back.content).includes('alert(1)'), '');
    }

    // ── Control-character stripping ───────────────────────────────────────────
    {
        const doc = contentFor(await gen([{ type: 'paragraph', children: [{ type: 'text', text: 'a\x00\x01\x1F\x08b' }] }]));
        check('odt: invalid XML control chars stripped', !/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(doc), 'control chars leaked');
        check('odt: content.xml well-formed after control-char strip', wellFormed(doc));
    }

    // ── Bookmark / anchor name sanitization + uniqueness ──────────────────────
    {
        const doc = contentFor(await gen([
            { type: 'heading', text: 'Dup', metadata: { level: 1, id: 'evil" name<>&/\\' }, children: [{ type: 'text', text: 'Dup' }] },
            { type: 'heading', text: 'Dup', metadata: { level: 1 }, children: [{ type: 'text', text: 'Dup' }] },
            { type: 'paragraph', children: [{ type: 'text', text: 'jump', metadata: { link: '#evil" name<>&/\\', linkType: 'internal' } }] },
        ]));
        const names = [...doc.matchAll(/text:(?:bookmark|name)="([^"]*)"/g)].map(m => m[1]);
        check('odt: bookmark names contain only safe chars', names.length > 0 && names.every(a => /^[A-Za-z_][A-Za-z0-9_]*$/.test(a)), JSON.stringify(names));
        const bmNames = [...doc.matchAll(/<text:bookmark text:name="([^"]+)"/g)].map(m => m[1]);
        check('odt: duplicate raw names produce unique bookmark names', new Set(bmNames).size === bmNames.length, JSON.stringify(bmNames));
        check('odt: content.xml well-formed with hostile anchor', wellFormed(doc));
    }

    // ── Color validation ──────────────────────────────────────────────────────
    {
        const doc = contentFor(await gen([{ type: 'paragraph', children: [{ type: 'text', text: 'c', formatting: { color: 'red"/><style:x' } as any }] }]));
        const colors = [...doc.matchAll(/fo:color="([^"]*)"/g)].map(m => m[1]);
        check('odt: color values are valid #hex or absent', colors.every(c => /^#[0-9A-Fa-f]{6}$/.test(c)), JSON.stringify(colors));
        check('odt: content.xml well-formed with hostile color', wellFormed(doc));
    }

    // ── Attachment name cannot escape Pictures/ ───────────────────────────────
    {
        const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
        const ast = astWith([{ type: 'image', metadata: { attachmentName: '../../evil.png' } }]);
        (ast as any).attachments = [{ name: '../../evil.png', mimeType: 'image/png', data: png }];
        const entries = Object.keys(unzipSync((await OfficeGenerator.generate(ast, 'odt')).value as Uint8Array));
        check('odt: no zip entry escapes its directory', entries.every(e => !e.includes('..')), entries.filter(e => e.includes('..')).join(', '));
        const pics = entries.filter(e => e.startsWith('Pictures/'));
        check('odt: media file names are synthesized', pics.every(e => /^Pictures\/image\d+\.[a-z]+$/.test(e)), pics.join(', '));
    }

    // ── Absurd colSpan is clamped ─────────────────────────────────────────────
    {
        const ast = astWith([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', metadata: { colSpan: 1e9, rowSpan: 1e9 }, children: [{ type: 'text', text: 'x' }] }] }] }]);
        const bytes = (await OfficeGenerator.generate(ast, 'odt')).value as Uint8Array;
        check('odt: absurd span does not blow up output', bytes.length < 200_000, `${bytes.length} bytes`);
        const doc = contentFor(bytes);
        check('odt: content.xml well-formed with clamped spans', wellFormed(doc));
        const span = doc.match(/table:number-columns-spanned="(\d+)"/);
        check('odt: span clamped to a sane value', !span || Number(span[1]) <= 1000, span?.[1] ?? '');
    }

    // ── styleMap output.tag cannot inject a style ─────────────────────────────
    {
        const pAst = astWith([{ type: 'paragraph', metadata: { style: 'Custom' }, children: [{ type: 'text', text: 'hi' }] } as any]);
        const sm = (tag: string) => ({ styleMap: [{ selector: { nodeType: 'paragraph', attributes: { style: 'Custom' } }, output: { tag } }] } as any);
        const hostile = contentFor((await OfficeGenerator.generate(pAst, 'odt', sm('h1"/><style:x style:name="Injected'))).value as Uint8Array);
        check('odt: hostile styleMap tag is not mapped to a style', !/Injected/.test(hostile), hostile.match(/text:style-name="[^"]*"/)?.[0] ?? '');
        check('odt: content.xml well-formed with hostile styleMap tag', wellFormed(hostile));
        const benign = contentFor((await OfficeGenerator.generate(pAst, 'odt', sm('h1'))).value as Uint8Array);
        check('odt: legitimate styleMap tag h1 maps to Heading_20_1', /text:style-name="Heading_20_1"/.test(benign) || /Heading_20_1/.test(benign), benign.match(/<text:p[^>]*>/)?.[0] ?? '');
    }

    // ── Whitespace-encoding metachars survive as literal text ─────────────────
    {
        const doc = contentFor(await gen([{ type: 'paragraph', children: [{ type: 'text', text: '<text:s text:c="99"/>' }] }]));
        check('odt: literal text:s in source is escaped, not emitted as markup', /&lt;text:s/.test(doc) && !/<text:s text:c="99"/.test(doc), doc.slice(0, 160));
    }
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
    // The limit grows by one cell per byte of the document (a large workbook is not cut short for its
    // size), so a file of a few hundred bytes gets a few hundred more than LIMIT.
    const limitFor = (buf: Buffer) => LIMIT + buf.length;

    // Single axis: a huge column repeat on a non-empty cell.
    const cols = build(ODS, `<office:spreadsheet><table:table table:name="S"><table:table-row>` +
        `<table:table-cell table:number-columns-repeated="5000000"><text:p>X</text:p></table:table-cell>` +
        `</table:table-row></table:table></office:spreadsheet>`);
    const rc = await parse(cols, 'ods', LIMIT);
    check('odf: column repeat is bounded', rc.cells <= limitFor(cols), `materialized ${rc.cells} cells against a limit of ${limitFor(cols)}`);
    check('odf: column clamp warns', rc.warned, 'truncation must not be silent');

    // Both axes: this is the combination that exhausted memory, since each row repetition
    // deep-copies the whole cell array.
    const both = build(ODS, `<office:spreadsheet><table:table table:name="S">` +
        `<table:table-row table:number-rows-repeated="10000">` +
        `<table:table-cell table:number-columns-repeated="10000"><text:p>X</text:p></table:table-cell>` +
        `</table:table-row></table:table></office:spreadsheet>`);
    const rb = await parse(both, 'ods', LIMIT);
    check('odf: rows x cols product is bounded', rb.cells <= limitFor(both), `materialized ${rb.cells} cells against a limit of ${limitFor(both)}`);

    // ODT/ODP keep empty cells on purpose (the grid is structural), so they have no empty-cell
    // skip and the budget is the only thing bounding them.
    const odt = build(ODT, `<office:text><table:table table:name="T"><table:table-row>` +
        `<table:table-cell table:number-columns-repeated="5000000"/>` +
        `</table:table-row></table:table></office:text>`);
    const ro = await parse(odt, 'odt', LIMIT);
    check('odf: ODT empty-cell repeat is bounded', ro.cells <= limitFor(odt), `materialized ${ro.cells} cells`);

    // MANY tables, each with a huge repeat. The budget is per document, so splitting the
    // expansion across tables must not multiply past the cap - the earlier single-table tests
    // would pass even with a per-table budget, which is exactly the hole this covers.
    const manyTables = build(ODT, '<office:text>' +
        (`<table:table table:name="T"><table:table-row><table:table-cell ` +
         `table:number-columns-repeated="1000000"><text:p>X</text:p></table:table-cell>` +
         `</table:table-row></table:table>`).repeat(20) + '</office:text>');
    const rm = await parse(manyTables, 'odt', LIMIT);
    check('odf: budget is per-document, not per-table', rm.cells <= limitFor(manyTables),
        `20 tables materialized ${rm.cells} cells against a per-document limit of ${limitFor(manyTables)}`);

    // A garbage (non-numeric) repeat must render the cell once, not drain the whole budget and
    // silently drop every legitimate cell that follows it.
    const garbage = build(ODS, `<office:spreadsheet>` +
        `<table:table table:name="A"><table:table-row><table:table-cell ` +
        `table:number-columns-repeated="abc"><text:p>GARBAGE</text:p></table:table-cell></table:table-row></table:table>` +
        `<table:table table:name="B"><table:table-row><table:table-cell><text:p>LEGIT</text:p></table:table-cell></table:table-row></table:table>` +
        `</office:spreadsheet>`);
    const gWarns: any[] = [];
    const gAst = await OfficeParser.parseOffice(garbage, { fileType: 'ods', onWarning: (w: any) => gWarns.push(w), decompressionLimits: { maxTableCells: LIMIT } } as any);
    const gText = (await gAst.to('text')).value;
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

    for (const fmt of ['html', 'md', 'text', 'rtf', 'csv', 'tex']) {
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
const QUIET = { onWarning: () => { } } as any;

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
    check('missing part: a minimal complete docx still parses', (await docxAst.to('text')).value.includes('Hello docx'),
        `got ${JSON.stringify((await docxAst.to('text')).value)}`);

    const pptx = Buffer.from(zipSync({
        'ppt/presentation.xml': zipEnc('<p:presentation/>'),
        'ppt/slides/slide1.xml': zipEnc('<?xml version="1.0"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Hello slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>'),
    }));
    const pptxAst = await OfficeParser.parseOffice(pptx, { fileType: 'pptx', ...QUIET } as any);
    check('missing part: a minimal complete pptx still parses', (await pptxAst.to('text')).value.includes('Hello slide'),
        `got ${JSON.stringify((await pptxAst.to('text')).value)}`);
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
    check('missing part: a minimal complete odt still parses', (await odtAst.to('text')).value.includes('Hello odt'),
        `got ${JSON.stringify((await odtAst.to('text')).value)}`);
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
    check('odf type: that spreadsheet\'s cells are actually parsed', (await hinted.to('text')).value.includes('CellValue'),
        `got ${JSON.stringify((await hinted.to('text')).value)}`);

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
            fromPath.type === 'ods' && (await fromPath.to('text')).value.includes('CellValue'),
            `got type ${fromPath.type}, text ${JSON.stringify((await fromPath.to('text')).value)}`);
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
        afterOdt.type === 'docx' && (await afterOdt.to('text')).value.includes('Word doc'),
        `got type ${afterOdt.type}, text ${JSON.stringify((await afterOdt.to('text')).value)}`);
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
    const generatorConfig: any = resolveGeneratorConfig('html', undefined, { onWarning: () => { } } as any);
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
    const generatorSource: any = resolveGeneratorConfig('html', undefined, { onWarning: () => { } } as any);
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

/**
 * An error reaches the caller's onWarning, and nothing is printed, wherever it is raised: a parser's
 * depth limit, a generator asked for an unknown format, an invalid style map. The OCR pool's
 * termination is reported by each parse whose image it interrupted, not printed on its own.
 */
async function errorRoutingTests() {
    console.log('- Error routing (onWarning, never the console)...');
    const printed: string[] = [];
    const { error: consoleError, warn: consoleWarn } = console;
    console.error = (...args: unknown[]) => { printed.push(String(args[0])); };
    console.warn = (...args: unknown[]) => { printed.push(String(args[0])); };
    try {
        const cases: [string, string, (onWarning: (i: any) => void) => Promise<unknown>][] = [
            ['html: markup nested past the limit', 'MAX_NESTING_DEPTH_EXCEEDED', onWarning => OfficeParser.parseOffice(Buffer.from('<b>'.repeat(300)), { fileType: 'html', onWarning } as any)],
            ['rtf: groups nested past the limit (no stack overflow)', 'MAX_NESTING_DEPTH_EXCEEDED', onWarning => OfficeParser.parseOffice(Buffer.from(`{\\rtf1 ${'{\\b a'.repeat(5000)}${'}'.repeat(5000)}}`), { fileType: 'rtf', onWarning } as any)],
            ['generate: an unknown format', 'FORMAT_UNSUPPORTED', onWarning => OfficeGenerator.generate(astWith([]), 'nope' as any, { onWarning } as any)],
            ['generate: an invalid style map', 'INVALID_SELECTOR', onWarning => OfficeGenerator.generate(astWith([]), 'md', { onWarning, styleMap: ['p[style-name= => '] } as any)],
        ];
        for (const [label, code, run] of cases) {
            const codes: string[] = [];
            const thrown = await run(issue => codes.push(issue.code)).then(() => 'none', (e: any) => e?.officeIssue?.code ?? e?.message);
            check(`error routing: ${label} throws ${code} and reports it to onWarning`, thrown === code && codes.includes(code), `threw ${thrown}, reported ${codes.join(', ')}`);
        }
        const rtfDepth = await OfficeParser.parseOffice(Buffer.from(`{\\rtf1 ${'{\\b a'.repeat(200)}${'}'.repeat(200)}}`), { fileType: 'rtf' } as any).then(ast => ast.content.length, (e: any) => e.message);
        check('rtf: 200 nested groups still parse', rtfDepth === 1, String(rtfDepth));
        await terminateOcr();
    } finally {
        console.error = consoleError;
        console.warn = consoleWarn;
    }
    check('error routing: nothing was printed', printed.length === 0, printed.join(' | '));
}

async function mdInlineFormattingTests() {
    console.log('- MarkdownGenerator inline formatting (opt-in <span style>)...');
    const cfg = { mdConfig: { fallbackToHtml: { inlineFormatting: true } } };
    // color/backgroundColor/size land in a style attribute; each value must be CSS-sanitized so it
    // can't break out of the attribute or inject a url()/expression().
    const payloads = ['red;}body{display:none', 'expression(alert(1))', 'url(javascript:alert(1))', 'x"><script>alert(1)</script>', 'red;background:url(//evil)'];
    for (const payload of payloads) {
        const ast = astWith([{ type: 'paragraph', children: [
            { type: 'text', text: 'x', formatting: { color: payload, backgroundColor: payload, size: payload } }] } as any]);
        const md = (await OfficeGenerator.generate(ast, 'md', cfg as any)).value as string;
        check(`md inline: payload ${JSON.stringify(payload.slice(0, 18))} no CSS/tag breakout`,
            !/expression\s*\(|url\s*\(|<script|javascript:/i.test(md) && !/style="[^"]*"[^>]*"/.test(md), md.slice(0, 200));
    }
    // Positive control: a legitimate color still round-trips (the checks above aren't vacuous).
    const ok = (await OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [
        { type: 'text', text: 'x', formatting: { color: '#c00' } }] } as any]), 'md', cfg as any)).value as string;
    check('md inline: legitimate color emitted', /<span style="color: ?#c00/.test(ok), ok.slice(0, 200));
}

/**
 * Control words TeX will execute in a piece of LaTeX, in order: tokenized the way TeX reads the
 * source (a backslash followed by letters is a control word; a backslash followed by any other
 * character is a control symbol, so `\\input` is a line break then the letters "input"), skipping
 * verbatim/lstlisting bodies and `%` comments, which TeX does not interpret.
 */
function liveLatexSource(tex: string): string {
    return tex
        .replace(/\\begin\{(verbatim|lstlisting)\}(\[[^\]\n]*\])?\n[\s\S]*?\n\\end\{\1\}/g, '')
        .replace(/(^|[^\\])((?:\\\\)*)%.*$/gm, '$1$2');
}

function liveControlWords(tex: string): string[] {
    const live = liveLatexSource(tex);
    const words: string[] = [];
    for (let i = 0; i < live.length; i++) {
        if (live[i] !== '\\') continue;
        let j = i + 1;
        while (j < live.length && /[A-Za-z@]/.test(live[j])) j++;
        if (j > i + 1) { words.push(live.slice(i + 1, j)); i = j - 1; } else { i++; }
    }
    return words;
}

/**
 * Every command the LaTeX generator writes on its own account. Document content never adds to this
 * set: text is escaped, URLs are encoded, and math that uses anything outside the formula is
 * written as literal text. So a hostile document's output may use only these.
 */
const LATEX_GENERATOR_COMMANDS = new Set((
    'documentclass usepackage ifPDFTeX else fi hypersetup setcounter maxdimen lstset ttfamily small makeatletter makeatother renewcommand paragraph ' +
    'subparagraph startsection z plus minus normalfont normalsize bfseries definecolor newunicodechar DeclareUnicodeCharacter ensuremath ding title author date ' +
    'begin end maketitle section subsection subsubsection chapter label phantomsection hyperref href footnote endnote footnotemark footnotetext addtocounter ' +
    'stepcounter theendnotes textbf textit texttt uline sout textsuperscript textsubscript textcolor colorbox setlength fboxsep strut fontsize selectfont protect ' +
    'item centering raggedleft raggedright arraybackslash dimexpr linewidth tabcolsep arrayrulewidth relax hline cline endhead multicolumn multirow cellcolor ' +
    'includegraphics ifdim textheight fbox hfil break newpage clearpage rule quad hspace phantom leftskip rightskip hangindent hangafter par cite ' +
    'textbackslash textasciitilde textasciicircum textless textgreater textbar textasciigrave pagestyle fancyhf fancyhead fancyfoot headrulewidth headheight ' +
    'fancypagestyle thepage frame frametitle note setbeamertemplate titlepage scriptsize footnotesize checkmark square boxtimes frac ' +
    'RequirePackage officeparserdriver ifpdf ifXeTeX iftutex IfFontExistsTF IfFileExists setmainfont setsansfont setmonofont xeCJKsetup ' +
    'xeCJKDeclareCharClass setCJKmainfont setCJKsansfont setCJKmonofont setCJKfallbackfamilyfont CJKrmdefault CJKsfdefault CJKttdefault ' +
    'ltjsetparameter setmainjfont setsansjfont setmonojfont ' +
    // Greek letters, which pdfLaTeX draws as math symbols.
    'alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi pi rho varsigma sigma tau upsilon phi chi psi omega ' +
    'Gamma Delta Theta Lambda Xi Pi Sigma Upsilon Phi Psi Omega'
).split(/\s+/));

/**
 * The generator's own lines that use commands no document content may ever produce (`\def`,
 * `\directlua`): the DVI driver choice and luatexja's font fallbacks. They are removed, by their
 * exact form, before the output is checked, so those commands anywhere else still fail it.
 */
const withoutGeneratorPrivileged = (tex: string) => tex
    .replace(/^\\def\\officeparserdriver\{\}\\ifpdf\\else\\ifXeTeX\\else\\def\\officeparserdriver\{dvipdfmx,\}\\fi\\fi$/m, '')
    .replace(/^ *\\directlua\{luaotfload\.add_fallback\("officeparser(rm|sf)", \{("[A-Za-z-]+\.(otf|ttf):mode=node;"(, )?)*\}\)\}%$/gm, '');

async function latexSanitizationTests() {
    console.log('- LaTeX generator (escaping, URLs, comments, math, bundle paths)...');

    // escapeLatex: every special character is neutralized, ligatures are broken, and no line
    // break survives raw (a blank line inside a macro argument would abort the run).
    check('latex: specials escaped', escapeLatex('\\{}$&%#_~^<>|`[]') ===
        '\\textbackslash{}\\{\\}\\$\\&\\%\\#\\_\\textasciitilde{}\\textasciicircum{}\\textless{}\\textgreater{}\\textbar{}\\textasciigrave{}{[}{]}');
    check('latex: ligatures broken', escapeLatex("--- '' ,,") === "-{}-{}- '{}' ,{},");
    check('latex: every line-break form becomes the newline argument',
        escapeLatex('a\nb\r\nc\rd\u000Be\u2028f\u2029g', '|') === 'a|b|c|d|e|f|g', escapeLatex('a\nb\r\nc\rd\u000Be\u2028f\u2029g', '|'));
    check('latex: controls, BOM and lone surrogates dropped', escapeLatex('a\u0000b\u0007c\uFEFFd\uD800e') === 'abcde');
    check('latex: decomposed accents composed (pdfLaTeX has no combining marks)', escapeLatex('e\u0301') === '\u00E9');
    check('latex: injection payload is inert', !liveControlWords(escapeLatex('\\input{/etc/passwd}\\write18{rm -rf /}')).some(w => w !== 'textbackslash'));

    // latexComment: whatever line breaks the text holds, nothing after the % escapes onto a live line.
    const comment = latexComment('x\n\\input{a}\r\\input{b}\u2028\\input{c}');
    check('latex comment: every line commented', comment.split('\n').filter(Boolean).every(l => l.startsWith('%')) && comment.endsWith('\n'), JSON.stringify(comment));

    // sanitizeLatexUrl: scheme policy, then nothing TeX-special left raw.
    for (const bad of ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'data:text/html,x', 'vbscript:x', 'file:///etc/passwd', '\\\\host\\share', '//host/share', 'java\u0000script:alert(1)']) {
        check(`latex url: ${JSON.stringify(bad)} rejected`, sanitizeLatexUrl(bad) === '');
    }
    const url = sanitizeLatexUrl('https://x.com/a b/{c}\\d^e|f~g$h?i=1&j=_#k%zz%41é');
    check('latex url: special characters encoded or escaped',
        url === 'https://x.com/a\\%20b/\\%7Bc\\%7D\\%5Cd\\%5Ee\\%7Cf\\%7Eg\\%24h?i=1\\&j=\\_\\#k\\%25zz\\%41\\%C3\\%A9', url);
    check('latex url: no raw brace or command survives', !/[{}]/.test(url) && liveControlWords(url).length === 0);
    // Its own encodings too: a bare `%` is a comment when the link sits in an argument (a heading, a footnote).
    check('latex url: every % it writes is escaped', !/(^|[^\\])%/.test(sanitizeLatexUrl('https://example.edu/~alice/Zürich?q=a b')), sanitizeLatexUrl('https://example.edu/~alice/Zürich?q=a b'));
    // escapeLatex copies the text between special characters in one pass: built a character at a time,
    // 20 MB took seconds and hundreds of megabytes, and 140 MB exhausted the heap.
    const escapeTimed = (n: number) => {
        const text = 'word & more_text 50% {x} -- '.repeat(Math.ceil(n / 28)).slice(0, n);
        const started = Date.now();
        const out = escapeLatex(text);
        return { ms: Date.now() - started, ok: out.startsWith('word \\& more\\_text 50\\% \\{x\\} -{}- ') };
    };
    const escapeSmall = escapeTimed(1_000_000), escapeLarge = escapeTimed(4_000_000);
    check('latex: escapeLatex is linear in its input', escapeSmall.ok && escapeLarge.ok && escapeLarge.ms < Math.max(200, 7 * escapeSmall.ms), `${escapeSmall.ms}ms vs ${escapeLarge.ms}ms`);
    const escapeHeap = process.memoryUsage().heapUsed;
    const escapedRun = escapeLatex('plain words here '.repeat(1_200_000));
    check('latex: escaping a 20 MB run holds about its own size', process.memoryUsage().heapUsed - escapeHeap < 150_000_000 && escapedRun.length === 20_400_000, `heap +${Math.round((process.memoryUsage().heapUsed - escapeHeap) / 1e6)}MB`);

    // latexSourceComment: a hidden note stays inside `%` lines whatever line breaks it holds, and a
    // `-->` inside it cannot end the comment early for the LaTeX parser.
    const hiddenNote = latexSourceComment(' a\r\\input{/etc/passwd}\u2028\\write18{id}\n--> \\def\\x{} ');
    check('latex source comment: every line commented', hiddenNote.split('\n').filter(Boolean).every(l => l.startsWith('%')) && hiddenNote.endsWith('\n'), JSON.stringify(hiddenNote));
    check('latex source comment: an inner --> is broken', (hiddenNote.match(/-->/g) || []).length === 1 && hiddenNote.trimEnd().endsWith('-->'), JSON.stringify(hiddenNote));

    // sanitizeLatexImagePath: TeX reads the file when compiling, so only a plain relative path inside
    // the document's folder passes, and it needs no escaping.
    for (const good of ['a.png', 'figures/diagram', 'img/fig_1.v2.pdf', '_x/y-z.jpg', 'pic with space.png', 'my figures/a b.png']) {
        check(`latex image path: ${JSON.stringify(good)} kept`, sanitizeLatexImagePath(good) === good);
    }
    // HTML/Markdown paths are URL references: percent-escapes are decoded, then the result is checked.
    check('latex image path: %20 decoded to a space', sanitizeLatexImagePath('pic%20with%20space.png') === 'pic with space.png');
    for (const bad of ['%2E%2E/secret.png', '%2Fetc%2Fpasswd', 'a%7Bb%7D.png', 'a%zz.png', 'a%25b.png']) {
        check(`latex image path: encoded ${JSON.stringify(bad)} refused`, sanitizeLatexImagePath(bad) === null);
    }
    for (const bad of ['../secret.png', 'a/../../b.png', './a.png', '/etc/passwd', '~/x.png', 'C:\\x.png', 'C:/x.png', '\\\\host\\share\\x.png',
        'https://x.com/a.png', 'file:///etc/passwd', '.hidden.png', 'a//b.png', 'a/', '-flag.png', 'a  b.png', 'a /b.png', 'a/ b.png', 'a\tb.png', 'a{b}.png', 'a\\input{x}.png',
        'a$b.png', '|kpsewhich x', "a'b.png", 'é.png', '', 'x'.repeat(256)]) {
        check(`latex image path: ${JSON.stringify(bad).slice(0, 40)} refused`, sanitizeLatexImagePath(bad) === null);
    }

    // sanitizeLatexMath: refused commands, catcode tricks and structure breakers never pass.
    const refused = ['\\input{/etc/passwd}', '\\include{x}', '\\write18{id}', '\\immediate\\write18{id}', '\\openout1=x', '\\openin1=/etc/passwd',
        '\\read1 to\\x', '\\directlua{os.execute("id")}', '\\catcode`\\@=11', '\\def\\x{y}', '\\let\\a\\b', '\\newcommand{\\x}{y}', '\\renewcommand{\\x}{y}',
        '\\NewDocumentCommand\\x{}{}', '\\csname input\\endcsname{x}', '\\scantokens{x}', '\\verb|x|', '\\includegraphics{/etc/passwd}', '\\pdffiledump{x}',
        '\\filedump{x}', '\\XeTeXinputencoding{x}', '\\everypar{x}', '\\usepackage{x}', '\\makeatletter', '\\href{javascript:x}{y}', '\\url{x}',
        '\\setlength\\textwidth{0pt}', '\\endinput', '\\stop', '\\show\\x', '\\lstinputlisting{/etc/passwd}', '\\ExplSyntaxOn', '\\special{x}', '\\font\\x=cmr10',
        // Commands that spell a refused command from text: each read /etc/os-release in a real compile.
        '\\UseName{input}{/etc/passwd}', '\\ExpandArgs{c}\\relax{input}{/etc/passwd}', '\\tokenized{\\string\\in put}', '\\csuse{input}{x}', '\\begincsname input\\endcsname',
        '\\csappto{maketitle}{x}', '\\enddocument'];
    for (const m of refused) {
        const r = sanitizeLatexMath(m, 'inline');
        check(`latex math: ${JSON.stringify(m)} refused`, !r.ok && r.commands.length > 0, JSON.stringify(r));
    }
    const caret = sanitizeLatexMath('^^5cinput{x}', 'inline');
    check('latex math: ^^ notation refused', !caret.ok && caret.commands.includes('^^'));
    for (const broken of ['a}', '{a', '\\begin{matrix}a', 'a\\end{matrix}', '\\begin{matrix}a\\end{pmatrix}', '{\\begin{matrix}}a\\end{matrix}',
        '\\begin{verbatim}x\\end{verbatim}', '\\begin{filecontents}{x}y\\end{filecontents}', '\\begin{document}', 'a\\)', '\\[x', 'x\\', '\\begin{align}x\\end{align} y']) {
        const r = sanitizeLatexMath(broken, broken.includes('align') ? 'block' : 'inline');
        check(`latex math: structural ${JSON.stringify(broken)} refused`, !r.ok && r.commands.length === 0, JSON.stringify(r));
    }
    const neutral = sanitizeLatexMath('a%b$c#d&e\\\\f', 'inline');
    check('latex math: %, $, #, & and \\\\ neutralized inline', neutral.ok && neutral.latex === 'a\\%b\\$c\\#d\\&e f', JSON.stringify(neutral));
    const blank = sanitizeLatexMath('a\n\n\nb', 'block');
    check('latex math: blank lines collapsed', blank.ok && blank.latex === 'a\nb', JSON.stringify(blank));
    const fine = sanitizeLatexMath('\\frac{a}{b} + \\begin{pmatrix}1&2\\\\3&4\\end{pmatrix} + \\text{ok} + \\mathbb{R}', 'inline');
    check('latex math: ordinary math untouched', fine.ok && fine.latex === '\\frac{a}{b} + \\begin{pmatrix}1&2\\\\3&4\\end{pmatrix} + \\text{ok} + \\mathbb{R}', JSON.stringify(fine));

    // Whole-document property: a payload in every sink leaves only generator-written commands.
    const P = '}\\input{/etc/passwd}\\write18{id}%\n\\end{document}\\begin{x}]$&#_^~';
    const hostile: any = {
        type: 'docx',
        metadata: { title: P, author: P, subject: P, keywords: P, description: P, lastModifiedBy: P, language: P, customProperties: { [P]: P } },
        attachments: [{ name: `../${P}.png`, mimeType: 'image/png', data: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==' }],
        auxiliary: { headers: [{ type: 'header', children: [{ type: 'paragraph', children: [{ type: 'text', text: P }] }] }] },
        content: [
            { type: 'heading', metadata: { level: 1, anchorIds: [P] }, children: [{ type: 'text', text: P, notes: [{ type: 'note', metadata: { noteType: 'footnote' }, children: [{ type: 'text', text: P }] }] }] },
            { type: 'paragraph', comments: [{ type: 'comment', metadata: { author: P, date: P }, children: [{ type: 'text', text: P }] }], children: [
                { type: 'text', text: P, formatting: { bold: true, color: P, backgroundColor: P, size: P, font: P } },
                { type: 'text', text: P, metadata: { link: `https://x.com/${P}`, linkType: 'external' } },
                { type: 'text', text: P, metadata: { link: `#${P}`, linkType: 'internal' } },
                { type: 'text', text: P, metadata: { citationKey: P } },
                { type: 'text', text: P, notes: [{ type: 'note', metadata: { noteType: 'endnote' }, children: [{ type: 'text', text: P }] }] },
                { type: 'code', metadata: { math: 'inline' }, text: P },
                { type: 'code', metadata: { math: 'inline' }, text: '\\frac{a}{b}' },
            ] },
            { type: 'list', metadata: { listId: P, listType: 'ordered', indentation: 1e9, itemIndex: 1e12 }, children: [{ type: 'text', text: P }] },
            { type: 'definitionList', children: [{ type: 'definitionTerm', children: [{ type: 'text', text: P }] }, { type: 'definitionDescription', children: [{ type: 'text', text: P }] }] },
            { type: 'table', children: [{ type: 'row', children: [
                { type: 'cell', metadata: { colSpan: 1e9, rowSpan: 1e9, backgroundColor: P }, children: [{ type: 'paragraph', children: [{ type: 'text', text: P, notes: [{ type: 'note', metadata: { noteType: 'footnote' }, children: [{ type: 'text', text: P }] }] }] }] },
            ] }, { type: 'row', children: [{ type: 'cell', children: [{ type: 'code', text: `${P}\\end{verbatim}`, metadata: { language: P } }] }] }] },
            { type: 'code', text: `x\n\\end{verbatim}\n${P}`, metadata: { language: 'python' } },
            { type: 'code', text: `x\n\\end{lstlisting}\n${P}`, metadata: { language: 'python' } },
            // beamer reads a fragile frame raw up to a line `\end{frame}`, so one in a code block ended the frame.
            { type: 'code', text: `x\n\\end{frame}\n${P}` },
            { type: 'code', text: `x\n  \\end{frame}\n${P}`, metadata: { language: 'python' } },
            { type: 'code', metadata: { math: 'block' }, text: P },
            { type: 'image', metadata: { attachmentName: `../${P}.png`, altText: P, width: `${P}%`, align: P } },
            { type: 'image', metadata: { url: `javascript:${P}`, altText: P } },
            { type: 'admonition', metadata: { admonitionType: P, title: P }, children: [{ type: 'paragraph', children: [{ type: 'text', text: P }] }] },
            { type: 'embed', metadata: { embedType: 'youtube', videoId: P, label: P } },
            { type: 'sheet', metadata: { sheetName: P }, children: [{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [{ type: 'text', text: P }] }] }] },
            { type: 'paragraph', metadata: { alignment: P, paragraphIndentation: { left: 1e15, hanging: 1e15 } }, children: [{ type: 'text', text: P }] },
        ],
        getImages: () => [],
    };
    // The same document in Greek, Cyrillic, Chinese, Japanese and Korean, so the per-script font setup is checked too.
    const scripts = { ...hostile, content: [...hostile.content, { type: 'paragraph', children: [{ type: 'text', text: `Ωμέγα Привет 中文 日本語のテキスト 한국어 ${P}` }] }] };
    for (const [label, doc, config] of [['article', hostile, {}], ['beamer', hostile, { texConfig: { documentClass: 'beamer' } }], ['fragment', hostile, { texConfig: { standalone: false } }],
        ['article with CJK, Greek and Cyrillic', scripts, {}]] as const) {
        const out = (await OfficeGenerator.generate(doc as any, 'tex', { renderMetadata: true, onWarning: () => { }, ...config } as any)).value as string;
        const foreign = [...new Set(liveControlWords(withoutGeneratorPrivileged(out)).filter(w => !LATEX_GENERATOR_COMMANDS.has(w)))];
        check(`latex ${label}: hostile content adds no command of its own`, foreign.length === 0, foreign.join(', '));
        const docEnds = (liveLatexSource(out).match(/(^|[^\\])\\end\{document\}/g) || []).length;
        check(`latex ${label}: exactly one live document end`, label === 'fragment' ? docEnds === 0 : docEnds === 1, `found ${docEnds}`);
        if (doc === scripts) check(`latex ${label}: the script fonts were set up`, out.includes('\\setCJKmainfont') && out.includes('\\setmainjfont') && out.includes('cmunrm.otf') && out.includes('luaotfload.add_fallback'));
        check(`latex ${label}: hostile spans and indices stay bounded`, out.length < 250000 && !/\d{8,}pt|\{\d{8,}\}/.test(out), `length ${out.length}`);
        // liveLatexSource treats verbatim bodies as inert; in a beamer frame one holding `\end{frame}` is not.
        const frameEnds = [...out.matchAll(/\\begin\{(verbatim|lstlisting)\}[^\n]*\n([\s\S]*?)\n\\end\{\1\}/g)].filter(m => /\\end\s*\{\s*frame\s*\}/.test(m[2]));
        check(`latex ${label}: no verbatim body holds \\end{frame}`, label !== 'beamer' || frameEnds.length === 0, `${frameEnds.length} found`);
    }
    const bundle = (await OfficeGenerator.generate(hostile, 'tex', { texConfig: { bundle: true }, onWarning: () => { } } as any)).value as Uint8Array;
    const names = Object.keys(unzipSync(bundle));
    check('latex bundle: entry names cannot traverse or escape', names.every(n => n === 'main.tex' || /^images\/[A-Za-z0-9-]+\.[a-z]+$/.test(n)), names.join(', '));
}

async function latexParserTests() {
    console.log('- LaTeX parser (expansion bombs, include cycles, traversal, nesting)...');
    const parse = async (src: string | Buffer) => {
        const codes: string[] = [];
        const started = Date.now();
        const ast = await OfficeParser.parseOffice(typeof src === 'string' ? Buffer.from(src) : src,
            { fileType: 'tex', extractAttachments: true, onWarning: (w: any) => codes.push(w.code) } as any);
        return { ast, codes, ms: Date.now() - started, json: JSON.stringify(ast.content) };
    };
    const zip = (files: Record<string, string>) => Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, new TextEncoder().encode(v)]))));
    const TIME_BUDGET_MS = 5000;
    const SIZE_BUDGET = 5_000_000;

    const tenfold = (a: string, b: string) => `\\def\\${b}{${`\\${a}`.repeat(10)}}`;
    const bomb = `\\def\\a{xx}${tenfold('a', 'b')}${tenfold('b', 'c')}${tenfold('c', 'd')}${tenfold('d', 'e')}${tenfold('e', 'f')}${tenfold('f', 'g')}${tenfold('g', 'h')}\\h`;
    for (const [label, src] of [
        ['exponential \\def', bomb],
        ['self-recursive macro', '\\newcommand{\\loop}{a\\loop}\\loop'],
        ['argument-doubling macro', '\\newcommand{\\x}[1]{\\x{#1#1}}\\x{a}'],
        ['recursive macro in math', '\\newcommand{\\m}{\\m\\m}$\\m$'],
        ['self-opening environment', '\\newenvironment{r}{\\begin{r}}{}\\begin{r}x\\end{r}'],
        ['argument-doubling xparse command', '\\NewDocumentCommand{\\x}{m}{\\x{#1#1}}\\x{a}'],
        ['self-opening xparse environment', '\\NewDocumentEnvironment{r}{O{x}}{\\begin{r}}{}\\begin{r}x\\end{r}'],
        ['recursive \\IfBooleanTF', '\\NewDocumentCommand{\\l}{s}{\\IfBooleanTF{#1}{\\l*\\l*}{\\l*\\l*}}\\l'],
    ] as const) {
        const r = await parse(src);
        check(`latex parser: ${label} is bounded`, r.ms < TIME_BUDGET_MS && r.json.length < SIZE_BUDGET && r.codes.includes('LATEX_EXPANSION_LIMIT_REACHED'),
            `${r.ms}ms, ${r.json.length} bytes, ${r.codes.join(',')}`);
    }
    for (const [label, src] of [
        ['50k nested groups', '{'.repeat(50000) + 'deep' + '}'.repeat(50000)],
        ['3k nested environments', '\\begin{quote}'.repeat(3000) + 'deep' + '\\end{quote}'.repeat(3000)],
        ['2k nested lists', '\\begin{itemize}\\item deep '.repeat(2000) + '\\end{itemize}'.repeat(2000)],
    ] as const) {
        const r = await parse(src);
        check(`latex parser: ${label} parse without recursion blow-up and keep the text`, r.ms < TIME_BUDGET_MS && r.json.includes('deep'), `${r.ms}ms`);
    }
    // Argument paths (formatting commands, notes, captions, links, nested tables) count against the same
    // depth bound as groups, so a tall stack of them ends with a warning, not a stack overflow; and a
    // `\maketitle` inside `\title` cannot re-enter the title block.
    for (const [label, src] of [
        ['3k nested \\footnote', '\\footnote{'.repeat(3000) + 'deep' + '}'.repeat(3000)],
        ['5k nested \\mbox', '\\mbox{'.repeat(5000) + 'deep' + '}'.repeat(5000)],
        ['3k nested \\caption', '\\caption{'.repeat(3000) + 'deep' + '}'.repeat(3000)],
        ['2k nested \\textbf', '\\textbf{'.repeat(2000) + 'deep' + '}'.repeat(2000)],
        ['1.5k nested tabulars', '\\begin{tabular}{c}'.repeat(1500) + 'deep' + '\\end{tabular}'.repeat(1500)],
        ['3k nested theorems', '\\begin{theorem}'.repeat(3000) + 'deep' + '\\end{theorem}'.repeat(3000)],
        ['3k nested proofs', '\\begin{proof}'.repeat(3000) + 'deep' + '\\end{proof}'.repeat(3000)],
        ['3k nested language environments', '\\begin{french}'.repeat(3000) + 'deep' + '\\end{french}'.repeat(3000)],
        ['5k nested \\foreignlanguage', '\\foreignlanguage{german}{'.repeat(5000) + 'deep' + '}'.repeat(5000)],
    ] as const) {
        let error = '';
        const r = await parse(src).catch((e: any) => { error = e.message; return null; });
        check(`latex parser: ${label} stops at the depth bound`, !!r && r.ms < TIME_BUDGET_MS && r.json.includes('deep') && r.codes.includes('LATEX_EXPANSION_LIMIT_REACHED'), error || `${r?.ms}ms ${r?.codes.join(',')}`);
    }
    for (const src of ['\\documentclass{article}\\title{a\\maketitle}\\begin{document}\\maketitle x\\end{document}',
        '\\documentclass{beamer}\\title{T}\\author{\\maketitle}\\begin{document}\\begin{frame}\\titlepage\\end{frame}\\end{document}']) {
        let error = '';
        await parse(src).catch((e: any) => { error = e.message; });
        check(`latex parser: \\maketitle inside the title block does not recurse (${src.slice(15, 22)})`, !error, error);
    }

    // Linear time: work that grows with the document, never with its square. Each case took tens of
    // seconds before (the pending text or raw body was re-read in full per character or space).
    for (const [label, src] of [
        ['1MB of spaced prose', 'word '.repeat(200000)],
        ['1MB \\verb body', '\\verb|' + 'x'.repeat(1_000_000) + '|'],
        ['1MB \\iffalse body', '\\iffalse ' + 'x'.repeat(1_000_000) + '\\fi'],
        ['1MB of } in verbatim', '\\begin{verbatim}' + '}'.repeat(1_000_000) + '\\end{verbatim}'],
        ['20k unclosed % <!-- lines', 'text\n' + '% <!-- x\n'.repeat(20000) + 'end'],
        ['100k conditionals in an unclosed \\iffalse', '\\iffalse ' + '\\ifnum1<2 x'.repeat(100000)],
        ['100k \\ifXeTeX\\else pairs never closed', '\\ifXeTeX a\\else b'.repeat(100000)],
        ['100k stray \\else and \\fi', '\\else x\\fi y'.repeat(100000)],
        ['100k unclosed \\ifcsname', '\\ifcsname x '.repeat(100000)],
        ['100k \\newif switches', Array.from({ length: 100000 }, (_, k) => `\\newif\\ifs${k}`).join('')],
        ['1MB \\ExplSyntaxOn body', '\\ExplSyntaxOn ' + 'x_y:n '.repeat(160000)],
        ['1MB unclosed xparse d() argument', '\\NewDocumentCommand\\d{d()}{#1}\\d(' + 'x'.repeat(1_000_000)],
        ['10k theorems numbered within sections', '\\newtheorem{t}{T}[section]' + '\\section{s}\\begin{t}x\\end{t}'.repeat(10000)],
        ['20k unclosed ConTeXt \\startitemize', '\\starttext ' + '\\startitemize \\item x '.repeat(20000)],
        ['40k xparse \\ends with no \\begin', '\\NewDocumentEnvironment{A}{}{}{}\\NewDocumentEnvironment{B}{}{}{}' + '\\begin{A}'.repeat(40000) + '\\end{B}'.repeat(40000)],
        ['100k unknown \\if words in a skipped branch', '\\iffalse ' + '\\ifunknown x '.repeat(100000)],
        // Scans that read comments, URLs and verbatim text as TeX does: each looks for a construct's end once.
        ['100k commented-out \\begin lines in a tabular', '\\begin{tabular}{l}\n' + '% \\begin{tabular}{l} x & y\n'.repeat(100000) + '\\end{tabular}'],
        ['100k near misses of \\begin and \\end in a tabular', '\\begin{tabular}{l}' + '\\bf \\em \\beg \\en \\% '.repeat(60000) + '\\end{tabular}'],
        ['100k \\verb on one line', 'x \\verb|a%b| '.repeat(100000)],
        ['100k unclosed \\verb on one line', '\\verb!a '.repeat(100000)],
        ['100k \\verb in one table cell row', '\\begin{tabular}{l}' + '\\verb|a&b| '.repeat(100000) + '\\end{tabular}'],
        ['20k unclosed \\url{ with %', '\\url{a%\n'.repeat(20000)],
        ['100k % lines in an equation', '\\begin{equation}' + 'x % c\n'.repeat(100000) + '\\end{equation}'],
        ['100k unclosed verbatim environments', '\\begin{verbatim}x'.repeat(100000)],
        ['a 1MB \\input line', '\\input ' + 'a'.repeat(1_000_000)],
    ] as const) {
        const r = await parse(src);
        check(`latex parser: ${label} parses in linear time`, r.ms < TIME_BUDGET_MS, `${r.ms}ms`);
    }
    // A .bib database is read in one pass, however its entries are broken, joined or long.
    for (const [label, bib] of [
        ['100k unclosed entries', '@book{k,title={'.repeat(100000)],
        ['300k bare @', '@'.repeat(300000)],
        ['20k entries under \\nocite{*}', Array.from({ length: 20000 }, (_, k) => `@misc{k${k}, author = {A and B and C}, title = {T${k}}}`).join('\n')],
        ['an author list of 100k names', `@misc{x, author = {${'A B and '.repeat(100000)}Z}}`],
        ['100k # joins', `@misc{x, title = ${'a # '.repeat(100000)}b}`],
    ] as const) {
        const r = await parse(zip({ 'main.tex': '\\documentclass{article}\\begin{document}\\nocite{*}\\cite{x}\\bibliography{refs}\\bibliography{refs}\\end{document}', 'refs.bib': bib }));
        check(`latex parser: a .bib of ${label} is read in linear time`, r.ms < TIME_BUDGET_MS && r.json.length < 20 * SIZE_BUDGET, `${r.ms}ms ${r.json.length}`);
    }
    // Picking a project's main file reads every file once, whatever the files include.
    const manyFiles = Object.fromEntries(Array.from({ length: 3000 }, (_, k) => [`ch/${k}.tex`, `\\input{ch/${k + 1}}\\include{${k + 2}}\\subfile{../x}\n% \\documentclass{article}\n`]));
    const mainPick = await parse(zip({ ...manyFiles, 'main.tex': '\\input{preamble}\\begin{document}Body\\end{document}', 'preamble.tex': '\\documentclass{article}' }));
    check('latex parser: the main file of 3,000 files including each other is found in linear time', mainPick.ms < TIME_BUDGET_MS && mainPick.json.includes('Body'), `${mainPick.ms}ms`);

    // Memory: an expansion is sized before it is built, and a column spec's repetition is bounded by length.
    const heapBefore = process.memoryUsage().heapUsed;
    const amplify = await parse('\\newcommand{\\x}[1]{' + '#1'.repeat(20000) + '}\\x{' + 'y'.repeat(20000) + '}');
    const spec = await parse('\\begin{tabular}{*{1000}{' + 'c'.repeat(100000) + '}}a\\end{tabular}');
    const heapGrowth = process.memoryUsage().heapUsed - heapBefore;
    check('latex parser: a macro repeating a long argument is refused before it is built', amplify.codes.includes('LATEX_EXPANSION_LIMIT_REACHED') && amplify.ms < TIME_BUDGET_MS, amplify.codes.join(','));
    check('latex parser: a long repeated column spec is bounded', spec.ms < TIME_BUDGET_MS && heapGrowth < 200_000_000, `${spec.ms}ms, heap +${Math.round(heapGrowth / 1e6)}MB`);

    // Included text counts against the expansion budget: including a large file many times over is bounded.
    const repeated = await parse(zip({ 'main.tex': '\\documentclass{article}\\begin{document}' + '\\input{big}'.repeat(400) + '\\end{document}', 'big.tex': 'z '.repeat(500000) }));
    check('latex parser: repeated \\input of a large file is bounded', repeated.ms < 3 * TIME_BUDGET_MS && repeated.codes.includes('LATEX_EXPANSION_LIMIT_REACHED'), `${repeated.ms}ms ${repeated.codes.join(',')}`);

    // Table and switch names from Object.prototype are unknown names, never inherited table entries.
    const proto = await parse('\\usepackage[constructor]{babel}\\usepackage[__proto__]{inputenc}\\babeltags{__proto__ = french}\\newtheorem{constructor}{C}'
        + '\\setdefaultlanguage{toString}\\begin{document}\\begin{constructor}a\\end{constructor}\\begin{toString}b\\end{toString}\\textvalueOf{c}'
        + '\\foreignlanguage{__proto__}{d}\\newif\\ifconstructor\\constructortrue\\ifconstructor e\\fi\\begin{hasOwnProperty}f\\end{hasOwnProperty}'
        + '\\\'{\\constructor}g\\"{\\toString}h\\end{document}');
    check('latex parser: prototype names in the language, theorem and encoding tables are plain names',
        ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h'].every(t => proto.json.includes(t)) && !proto.json.includes('native code') && !proto.json.includes('function') && ({} as any).polluted === undefined, proto.json.slice(0, 200));

    const unclosed = await parse('\\begin{itemize}\\item a \\textbf{b \\begin{tabular}{ll} x & y');
    check('latex parser: unclosed groups and environments do not throw', unclosed.json.includes('a'));

    // No file system access, ever: a lone .tex cannot reach outside itself, and a project zip only its own files.
    const single = await parse('\\input{/etc/passwd}\\input{../../../etc/passwd}\\includegraphics{/etc/passwd}');
    check('latex parser: \\input of a system path reads nothing', !single.json.includes('root:') && single.codes.includes('LATEX_FILE_NOT_FOUND') && single.ast.attachments.length === 0);
    const traversal = await parse(zip({ 'main.tex': '\\documentclass{article}\\begin{document}\\input{../outside}\\input{/abs}\\includegraphics{../x.png}\\end{document}', 'x.png': 'png' }));
    check('latex parser: project paths cannot leave the project', traversal.ast.attachments.length === 0 && traversal.codes.includes('LATEX_FILE_NOT_FOUND'));
    const cycle = await parse(zip({ 'main.tex': '\\documentclass{article}\\begin{document}A\\input{b}\\end{document}', 'b.tex': 'B\\input{main}' }));
    check('latex parser: include cycle stops with a warning', cycle.ms < TIME_BUDGET_MS && cycle.codes.includes('LATEX_EXPANSION_LIMIT_REACHED') && cycle.json.includes('B'));

    // Parsed text is data: a document spelling out LaTeX commands in \verb is not re-executed on regeneration.
    const verb = await parse('\\verb|\\input{/etc/passwd}|');
    const regenerated = (await verb.ast.to('tex')).value as string;
    check('latex parser: verbatim command text regenerates escaped', regenerated.includes('\\textbackslash{}input') && !liveControlWords(regenerated).includes('input'));

    // Images carried as PDFs (a filecontents block, or a PDF in a project) are read under fixed bounds:
    // anything suspect stays the PDF it is, unread.
    const pdfWith = (image: string) => `%PDF-1.5\n1 0 obj\n<< /Type /Page /Resources << /XObject << /Im0 2 0 R >> >> /Contents 3 0 R >>\nendobj\n2 0 obj\n${image}\nendobj\n3 0 obj\n<< /Length 26 >>\nstream\nq 1 0 0 1 0 0 cm /Im0 Do Q\nendstream\nendobj\n%%EOF\n`;
    const carried = (pdf: string, name = 'x.pdf') => `\\documentclass{article}\n\\begin{filecontents*}{${name}}\n${pdf}\\end{filecontents*}\n\\begin{document}\\includegraphics{x.pdf}\\end{document}`;
    const tinyZlib = Buffer.from(zlibSync(new Uint8Array(4096))).toString('latin1');
    for (const [label, image] of [
        ['a claimed 10-gigapixel image', '<< /Type /XObject /Subtype /Image /Width 100000 /Height 100000 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length 8 >>\nstream\nxxxxxxxx\nendstream'],
        ['a 36-megapixel image from a few bytes', `<< /Type /XObject /Subtype /Image /Width 6000 /Height 6000 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length ${tinyZlib.length} >>\nstream\n${tinyZlib}\nendstream`],
        ['deeply nested dictionaries', '<< /A '.repeat(20000) + '>>'.repeat(20000)],
        ['broken ASCII85', '<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter [/ASCII85Decode /FlateDecode] /Length 12 >>\nstream\n{{{{vvvv~>xx\nendstream'],
        ['a Length past the end', '<< /Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8 /Length 99999999 >>\nstream\nabcd'],
    ] as const) {
        const started = process.memoryUsage().heapUsed;
        const r = await parse(carried(pdfWith(image)));
        const grew = process.memoryUsage().heapUsed - started;
        check(`latex parser: ${label} in a carried PDF stays a PDF, within bounds`, r.ms < TIME_BUDGET_MS && grew < 100_000_000 && r.ast.attachments.length === 1 && r.ast.attachments[0].mimeType === 'application/pdf',
            `${r.ms}ms, heap +${Math.round(grew / 1e6)}MB, ${r.ast.attachments.map(a => a.mimeType).join(',')}`);
    }
    // The object scan is linear: text an unclosed string or dictionary swallowed is read once, not
    // again from every `obj` inside it (300 KB took minutes).
    for (const [label, unit] of [['unclosed strings', '1 0 obj\n('], ['unclosed dictionaries', '1 0 obj\n<</K('], ['unclosed arrays', '1 0 obj\n[']] as const) {
        const r = await parse(carried('%PDF-1.5\n' + unit.repeat(Math.ceil(300_000 / unit.length)) + '\n'));
        check(`latex parser: a carried PDF of ${label} is scanned in linear time`, r.ms < TIME_BUDGET_MS && r.ast.attachments[0]?.mimeType === 'application/pdf', `${r.ms}ms`);
    }
    // Decoding is bounded by what it produces, typed arrays included: an image past the per-image
    // budget is refused before a buffer is allocated, and one at it decodes in bounded time and memory.
    // Binary image data travels in a project zip (a .tex is text, so filecontents cannot carry it).
    const decodeCase = async (width: number, bpc: number, data: Uint8Array) => {
        const stream = Buffer.from(data).toString('latin1');
        const pdf = pdfWith(`<< /Type /XObject /Subtype /Image /Width ${width} /Height ${width} /ColorSpace /DeviceRGB /BitsPerComponent ${bpc} /Filter /FlateDecode /DecodeParms << /Predictor 12 /Colors 3 /BitsPerComponent ${bpc} /Columns ${width} >> /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
        const project = Buffer.from(zipSync({
            'main.tex': new TextEncoder().encode('\\documentclass{article}\\begin{document}\\includegraphics{x.pdf}\\end{document}'),
            'x.pdf': new Uint8Array(Buffer.from(pdf, 'latin1')),
        }));
        const before = process.memoryUsage().arrayBuffers;
        const r = await parse(project);
        return { r, grew: process.memoryUsage().arrayBuffers - before };
    };
    const refused = await decodeCase(4000, 16, zlibSync(new Uint8Array(4096)));
    check('latex parser: a carried image past the decoding budget stays a PDF, nothing allocated', refused.r.ms < TIME_BUDGET_MS && refused.grew < 50_000_000 && refused.r.ast.attachments[0]?.mimeType === 'application/pdf', `${refused.r.ms}ms, +${Math.round(refused.grew / 1e6)}MB`);
    const atCap = await decodeCase(3999, 8, zlibSync(new Uint8Array((3999 * 3 + 1) * 3999), { level: 9 }));
    check('latex parser: a carried image at the decoding budget decodes in bounded time and memory', atCap.r.ms < TIME_BUDGET_MS && atCap.grew < 300_000_000 && atCap.r.ast.attachments[0]?.mimeType === 'image/png', `${atCap.r.ms}ms, +${Math.round(atCap.grew / 1e6)}MB`);
    const escaping = await parse('\\begin{filecontents*}{../../outside.tex}\nEscaped\n\\end{filecontents*}\\begin{document}\\input{../../outside}\\end{document}');
    check('latex parser: a filecontents file cannot be written outside the project', !escaping.json.includes('Escaped') && escaping.codes.includes('LATEX_FILE_NOT_FOUND'));
}

/**
 * Hardening a second parser review asked for: keys a document names cannot reach a shared prototype,
 * no construct costs time in the square of its size (or doubles per nesting level), spans are
 * bounded, the XML library writes nothing to the console, and an EPUB's pictures and chapters are
 * read once however often they are referred to.
 */
async function parserHardeningTests() {
    console.log('- Parser hardening (prototype keys, linear time, spans, console, EPUB)...');
    const files = path.join(__dirname, '..', 'files');
    const enc = (s: string) => new TextEncoder().encode(s);
    const repack = (file: string, edit: (z: Record<string, Uint8Array>) => void) => {
        const zip = unzipSync(new Uint8Array(fs.readFileSync(path.join(files, file))));
        edit(zip);
        return Buffer.from(zipSync(zip));
    };
    const builtIns = new Set(Object.getOwnPropertyNames(Object.prototype));
    const polluted = () => Object.getOwnPropertyNames(Object.prototype).filter(k => !builtIns.has(k));
    const parseQuiet = async (buffer: Buffer, fileType: string, extra: object = {}) => {
        try { return { ast: await OfficeParser.parseOffice(buffer, { fileType, onWarning: () => {}, ...extra } as any), error: '' }; }
        catch (e: any) { return { ast: undefined, error: String(e?.message ?? e) }; }
    };
    const hasFunctionValue = (value: any, seen = new Set<any>()): boolean => {
        if (typeof value === 'function') return true;
        if (!value || typeof value !== 'object' || seen.has(value)) return false;
        seen.add(value);
        return Object.keys(value).some(k => k !== 'to' && hasFunctionValue(value[k], seen));
    };

    // Prototype keys: a drawing with no relationships whose picture embeds `__proto__` (it wrote the
    // picture's description onto every object), numbering overrides at levels `__proto__` and
    // `constructor` (they wrote a start number onto every object) and an abstract numbering id of
    // `toString` (the parse failed on the prototype's function).
    const drawing = '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="n" descr="POLLUTED"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="__proto__"/></xdr:blipFill></xdr:pic></xdr:twoCellAnchor></xdr:wsDr>';
    const xlsxDrawing = await parseQuiet(repack('test.xlsx', z => { z['xl/drawings/drawing77.xml'] = enc(drawing); }), 'xlsx', { extractAttachments: true });
    check('xlsx: a picture embedding __proto__ writes nothing onto Object.prototype', !xlsxDrawing.error && polluted().length === 0 && ({} as any).altText === undefined, `${xlsxDrawing.error} ${polluted()}`);
    const numbering = '<?xml version="1.0"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="__proto__"><w:startOverride w:val="1337"/></w:lvlOverride><w:lvlOverride w:ilvl="constructor"><w:startOverride w:val="42"/></w:lvlOverride></w:num><w:num w:numId="2"><w:abstractNumId w:val="toString"/></w:num></w:numbering>';
    const docxNumbering = await parseQuiet(repack('test.docx', z => { z['word/numbering.xml'] = enc(numbering); }), 'docx');
    check('docx: numbering overrides at __proto__ and constructor write nothing onto a prototype', !docxNumbering.error && polluted().length === 0 && ({} as any).start === undefined && (Object as any).start === undefined, `${docxNumbering.error} ${polluted()}`);
    for (const key of polluted()) delete (Object.prototype as any)[key];
    // A value looked up in a table of names (a highlight colour, an alignment, a package's type) is
    // never what a plain object inherits.
    for (const key of ['constructor', 'toString', '__proto__']) {
        const docx = await parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/(<w:rPr>)/, `$1<w:highlight w:val="${key}"/>`)); }), 'docx');
        const odt = await parseQuiet(repack('test.odt', z => { z['content.xml'] = enc(new TextDecoder().decode(z['content.xml']).replace(/fo:text-align="[^"]*"/g, `fo:text-align="${key}"`)); }), 'odt');
        const pptx = await parseQuiet(repack('test.pptx', z => { for (const k of Object.keys(z)) if (/^ppt\/slides\/slide\d+\.xml$/.test(k)) z[k] = enc(new TextDecoder().decode(z[k]).replace(/\balgn="[^"]*"/g, `algn="${key}"`)); }), 'pptx');
        check(`docx, odt, pptx: a highlight or alignment named ${key} reads as no function or prototype`, [docx, odt, pptx].every(r => !r.error && !hasFunctionValue(r.ast?.content) && !hasFunctionValue(r.ast?.metadata)), [docx, odt, pptx].map(r => r.error).join(' '));
        const odtMime = await parseQuiet(repack('test.odt', z => { z['mimetype'] = enc(key); }), undefined as any);
        check(`zip: a mimetype of ${key} is an unsupported type, not a failure inside`, /supports/.test(odtMime.error) || !!odtMime.ast, odtMime.error);
    }
    // Names a document gives its own properties stay its own, whatever they are.
    const frontMatter = await parseQuiet(Buffer.from('---\n__proto__: [a, b]\nconstructor: x\ntitle: T\n---\n\nText\n'), 'md');
    const custom = frontMatter.ast?.metadata.customProperties as any;
    check('md: front matter named __proto__ is a property, not a prototype', !!custom && Object.getPrototypeOf(custom) === Object.prototype && Object.prototype.hasOwnProperty.call(custom, '__proto__') && custom.constructor === 'x', JSON.stringify(custom));

    // Linear time.
    const timed = async (label: string, run: () => Promise<unknown>) => {
        const started = Date.now();
        await run();
        check(`${label} in linear time`, Date.now() - started < 5000, `${Date.now() - started}ms`);
    };
    const tex = (body: string) => `\\documentclass{article}\\begin{document}${body}\\end{document}`;
    for (const [label, src] of [
        ['keywords nested 120 deep', tex('\\keywords{'.repeat(120) + 'x' + '}'.repeat(120))],
        ['pdfinfo nested 120 deep', tex('\\hypersetup{pdfinfo={A={'.repeat(120) + 'x' + '}}}'.repeat(120))],
        ['keywords environments nested 120 deep', tex('\\begin{keywords}'.repeat(120) + 'x' + '\\end{keywords}'.repeat(120))],
        ['a 200 KB % !TEX encoding line', '% !TEX encoding = a' + ' '.repeat(200000) + '!\n' + tex('x')],
        ['40k unclosed \\usepackage[', '\\usepackage['.repeat(40000) + tex('x')],
        ['40k unclosed \\inputencoding{', '\\inputencoding{'.repeat(40000)],
        ['40k unclosed \\documentclass[', '\\documentclass['.repeat(40000) + '\\begin{document}x\\end{document}'],
        ['40k unclosed \\documentclass{', '\\documentclass{'.repeat(40000)],
        ['80k unclosed filecontents', '\\begin{filecontents}'.repeat(80000) + tex('x')],
        ['a \\cmidrule before 200 KB of spaces', tex('\\begin{tabular}{l}\\cmidrule' + ' '.repeat(200000) + 'x\\end{tabular}')],
        ['a \\rowcolor before 200 KB of spaces', tex('\\begin{tabular}{l}\\rowcolor' + ' '.repeat(200000) + 'x\\end{tabular}')],
        ['a longtable \\caption before 200 KB of spaces', tex('\\begin{longtable}{l}\\caption' + ' '.repeat(200000) + 'x\\end{longtable}')],
        ['a \\multirow before 200 KB of spaces', tex('\\begin{tabular}{l}\\multirow' + ' '.repeat(200000) + 'x\\end{tabular}')],
        ['a \\cellcolor before 200 KB of spaces', tex('\\begin{tabular}{l}\\cellcolor' + ' '.repeat(200000) + 'x\\end{tabular}')],
        ['40k \\begin{ past the nesting limit', tex('{'.repeat(130) + '\\begin{'.repeat(40000))],
        ['a width of 200 KB of spaces', tex('\\includegraphics[width=1' + ' '.repeat(200000) + 'x]{a}')],
        ['keywords with 200 KB of spaces before \\and', tex('\\keywords{a' + ' '.repeat(200000) + 'b \\and c}')],
    ] as const) await timed(`latex parser: ${label} parses`, () => parseQuiet(Buffer.from(src), 'tex'));
    await timed('latex parser: a \\subfile of 40k \\begin{document} parses', () => parseQuiet(Buffer.from(zipSync({ 'main.tex': enc(tex('\\subfile{a}')), 'a.tex': enc('\\begin{document}'.repeat(40000)) })), 'zip' as any));
    const sheetXml = new TextDecoder().decode(unzipSync(new Uint8Array(fs.readFileSync(path.join(files, 'test.xlsx'))))['xl/worksheets/sheet1.xml']);
    const withSheetData = (body: string) => repack('test.xlsx', z => { z['xl/worksheets/sheet1.xml'] = enc(sheetXml.replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/, `<sheetData>${body}</sheetData>`)); });
    await timed('xlsx: 160k unclosed rows parse', () => parseQuiet(withSheetData('<row r="1"><c r="A1"><v>1</v></c>'.repeat(160000)), 'xlsx'));
    await timed('xlsx: a row of 160k unclosed cells parses', () => parseQuiet(withSheetData('<row r="1">' + '<c r="A1"><v>1</v>'.repeat(160000) + '</row>'), 'xlsx'));
    await timed('xlsx: a cell of 160k unclosed values parses', () => parseQuiet(withSheetData('<row r="1"><c r="A1" t="inlineStr"><is>' + '<t>1'.repeat(160000) + '</is></c></row>'), 'xlsx'));
    const unclosedRows = await parseQuiet(withSheetData('<row r="1"><c r="A1"><v>1</v></c><row r="2"><c r="A2"><v>2</v></c>'), 'xlsx');
    check('xlsx: unclosed rows keep their cells', JSON.stringify(unclosedRows.ast?.content).includes('"2"'), unclosedRows.error);
    await timed('md: a block whose second line is 300 KB, then 150k lines, parses', () => parseQuiet(Buffer.from('a\n|' + ' '.repeat(300000) + '-x\n' + 'a\n'.repeat(150000)), 'md'));
    await timed('md: 600 nested items holding unclosed fences, then 1M blank lines, parse', () => {
        let src = '';
        for (let j = 0; j < 600; j++) src += ' '.repeat(2 * j) + '- a\n' + ' '.repeat(2 * j + 2) + '```\n';
        return parseQuiet(Buffer.from(src + '\n'.repeat(1_000_000)), 'md');
    });
    await timed('md: a line of 400 KB of spaces before a block parses', () => parseQuiet(Buffer.from('- a\n\n' + ' '.repeat(400000) + 'x\n'), 'md'));
    await timed('md: indented code of 200k blank lines before its last line parses', () => parseQuiet(Buffer.from('    a' + '\n'.repeat(200000) + '    b\n'), 'md'));
    await timed('html: 100k unclosed <style>s parse', () => parseQuiet(Buffer.from('<p><style>'.repeat(100000)), 'html'));
    await timed('html: 80k cells of unclosed <script>s parse', () => parseQuiet(Buffer.from('<table><tr>' + '<td><script>'.repeat(80000)), 'html'));
    await timed('csv: a 1 MB run of digits is checked as a formula', async () => csvSafeCell('1'.repeat(1_000_000) + 'x'));
    await timed('xlsx: 1 MB of & in an inline string decodes', () => parseQuiet(withSheetData('<row r="1"><c r="A1" t="inlineStr"><is><t>' + '&amp;'.repeat(200000) + '</t></is></c></row>'), 'xlsx'));
    await timed('docx: an image width of 200 KB of spaces is written', () => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [{ type: 'image', metadata: { url: 'a.png', width: '1' + ' '.repeat(200000) + 'x' } }] }]), 'docx' as any, { onWarning: () => {} } as any));
    await timed('md: block math of 200 KB of spaces is written inline', () => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'a' }, { type: 'code', text: 'x' + ' '.repeat(200000) + 'y', metadata: { math: 'block' } }] }]), 'md' as any, { onWarning: () => {} } as any));

    // Raw HTML in Markdown, read as links and pictures, passes the writers' URL checks like any other.
    const rawHtml = await parseQuiet(Buffer.from('<a href="javascript:alert(1)">x</a> <a href="java&#115;cript:alert(2)">y</a> <img src="javascript:alert(3)" alt="z" onerror="alert(4)">\n\n<p><a href="javascript:alert(5)">w</a><img src="x" onerror="alert(6)"></p>\n'), 'md');
    for (const format of ['md', 'html'] as const) {
        const out = (await OfficeGenerator.generate(rawHtml.ast!, format, { onWarning: () => {} } as any)).value as string;
        check(`md: raw HTML links and pictures reach ${format} output without a script URL or handler`, !rawHtml.error && !/javascript:|onerror|alert\(/i.test(out), out.slice(0, 300));
    }
    await timed('md: 160k unclosed inline tags parse', () => parseQuiet(Buffer.from('x <b><kbd class="a"><a href="u">'.repeat(40000)), 'md'));
    await timed('md: 80k HTML block paragraphs parse', () => parseQuiet(Buffer.from('<p>x</p>\n\n'.repeat(80000)), 'md'));

    // One note, comment or string referred to many times is built once and written once: built per
    // reference, a few kilobytes of DOCX, EPUB or XLSX filled the heap, and writers copied the note
    // into their output at every reference.
    const bigText = 'word '.repeat(40000); // 200 KB
    const heapBudget = async (label: string, run: () => Promise<unknown>) => {
        global.gc?.();
        const before = process.memoryUsage().heapUsed;
        const result = await run();
        const grown = process.memoryUsage().heapUsed - before;
        check(`${label} without copying it per reference`, grown < 300_000_000, `${Math.round(grown / 1e6)}MB`);
        return result;
    };
    const docxRefs = (tag: 'footnote' | 'endnote' | 'comment') => repack('test.docx', z => {
        const doc = new TextDecoder().decode(z['word/document.xml']);
        const ref = tag === 'comment' ? '<w:r><w:commentReference w:id="7"/></w:r>' : `<w:r><w:${tag}Reference w:id="7"/></w:r>`;
        z['word/document.xml'] = enc(doc.replace(/<w:body>/, `<w:body><w:p>${ref.repeat(2000)}</w:p>`));
        const part = tag === 'comment' ? 'comments' : `${tag}s`;
        const el = tag === 'comment' ? 'w:comment' : `w:${tag}`;
        z[`word/${part}.xml`] = enc(`<?xml version="1.0"?><w:${part} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><${el} w:id="7"><w:p><w:r><w:t>${bigText}</w:t></w:r></w:p></${el}></w:${part}>`);
    });
    for (const tag of ['footnote', 'endnote', 'comment'] as const) {
        const r: any = await heapBudget(`docx: a ${tag} referred to 2000 times parses`, () => parseQuiet(docxRefs(tag), 'docx'));
        const shared: any[] = [];
        const collect = (nodes: any[] | undefined) => nodes?.forEach((n: any) => { shared.push(...(n.notes ?? []), ...(n.comments ?? [])); collect(n.children); });
        collect(r.ast?.content);
        const counts = new Map<any, number>();
        for (const n of shared) counts.set(n, (counts.get(n) ?? 0) + 1);
        const [top, topCount] = [...counts].sort((x, y) => y[1] - x[1])[0] ?? [undefined, 0];
        check(`docx: a ${tag} referred to 2000 times keeps its text, in one shared node`, !r.error && topCount === 2000 && top.text.length > 190_000, `${r.error} ${topCount}`);
    }
    const htmlNote = `<p>x${'<sup data-footnote-ref="a"></sup>'.repeat(2000)}</p><section data-footnotes><div data-footnote-id="a">${bigText}<b>${bigText}</b></div></section>`;
    await heapBudget('html: a footnote referred to 2000 times parses', () => parseQuiet(Buffer.from(htmlNote), 'html'));
    await heapBudget('md: an HTML block referring to one footnote 2000 times parses', () => parseQuiet(Buffer.from('# T\n\n' + htmlNote + '\n'), 'md'));
    await heapBudget('epub: a chapter referring to one footnote 2000 times parses', () => parseQuiet(Buffer.from(zipSync({
        'mimetype': enc('application/epub+zip'),
        'META-INF/container.xml': enc('<?xml version="1.0"?><container><rootfiles><rootfile full-path="content.opf"/></rootfiles></container>'),
        'content.opf': enc('<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata><manifest><item id="c" href="c.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="c"/></spine></package>'),
        'c.xhtml': enc(`<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><body>${htmlNote}</body></html>`),
    })), 'epub'));
    await heapBudget('xlsx: a rich shared string shown in 2000 cells parses', () => parseQuiet(repack('test.xlsx', z => {
        z['xl/sharedStrings.xml'] = enc(`<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>${bigText}</t></r><r><rPr><b/></rPr><t>${bigText}</t></r></si></sst>`);
        z['xl/worksheets/sheet1.xml'] = enc(sheetXml.replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/, `<sheetData>${Array.from({ length: 2000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join('')}</sheetData>`));
    }), 'xlsx'));
    const sharedNote = await parseQuiet(Buffer.from('[^a] '.repeat(2000) + '\n\n[^a]: ' + 'word '.repeat(2000) + '\n'), 'md');
    for (const format of ['text', 'rtf', 'tex', 'odt', 'chunks', 'html', 'md', 'docx', 'epub'] as const) {
        const out = (await OfficeGenerator.generate(sharedNote.ast!, format, { onWarning: () => {} } as any)).value as any;
        const size = typeof out === 'string' ? out.length : Array.isArray(out) ? JSON.stringify(out).length : Object.values(unzipSync(new Uint8Array(out))).reduce((n: number, f: any) => n + f.length, 0);
        // DOCX refers to a note again with a NOTEREF field (about 150 bytes a reference); the note copied
        // per reference would be 20 MB.
        check(`${format}: a note referred to 2000 times is written once`, size < (format === 'docx' ? 600_000 : 400_000), `${size} chars`);
    }

    // A network path cannot reach a page opened from disk as file://host (SMB, with the user's
    // credentials on Windows): `//host` is written as https://host, a backslash form is refused.
    const unc = astWith([{ type: 'paragraph', children: [
        { type: 'text', text: 'l1', metadata: { link: '\\\\evil\\share\\doc', linkType: 'external' } },
        { type: 'text', text: 'l2', metadata: { link: '//evil/share', linkType: 'external' } },
        { type: 'image', metadata: { url: '\\\\evil\\share\\x.png', altText: 'a' } },
        { type: 'image', metadata: { url: '//evil/y.png', altText: 'b' } },
        { type: 'image', metadata: { url: '\\/evil/z.png', altText: 'c' } },
    ] }]);
    for (const format of ['html', 'md', 'epub'] as const) {
        const value = (await OfficeGenerator.generate(unc, format, { onWarning: () => {} } as any)).value as any;
        const out = typeof value === 'string' ? value : Object.values(unzipSync(new Uint8Array(value))).map((u: any) => new TextDecoder().decode(u)).join('\n');
        check(`${format}: no network path is written, and //host is https`, !/\\\\evil|["(]\/\/evil|\\\/evil/.test(out) && out.includes('https://evil/share'), out.match(/[^\n]*evil[^\n]*/g)?.slice(0, 3).join(' | '));
    }

    // Notes and comments referring to each other (each holds two references to the one before): every
    // pass over the AST and every writer takes each shared node once, so depth 40 is instant (it
    // doubled per level), and a comment is written once.
    for (const kind of ['footnote', 'comment'] as const) {
        const el = kind === 'comment' ? 'w:comment' : 'w:footnote';
        const part = kind === 'comment' ? 'comments' : 'footnotes';
        const ref = (id: number) => kind === 'comment' ? `<w:r><w:commentReference w:id="${id}"/></w:r>` : `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
        let items = '';
        for (let k = 1; k <= 40; k++) items += `<${el} w:id="${k}"><w:p><w:r><w:t>item${k} </w:t></w:r>${k > 1 ? ref(k - 1) + ref(k - 1) : ''}</w:p></${el}>`;
        const chain = await parseQuiet(repack('test.docx', z => {
            z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/<w:body>/, `<w:body><w:p><w:r><w:t>x</w:t></w:r>${ref(40)}</w:p>`));
            z[`word/${part}.xml`] = enc(`<?xml version="1.0"?><w:${part} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${items}</w:${part}>`);
        }), 'docx');
        for (const format of ['text', 'md', 'html', 'rtf', 'tex', 'docx', 'odt', 'epub', 'chunks'] as const) {
            const started = Date.now();
            const value = (await OfficeGenerator.generate(chain.ast!, format, { onWarning: () => {} } as any)).value as any;
            const size = typeof value === 'string' ? value.length : Array.isArray(value) ? JSON.stringify(value).length : value.length;
            check(`${format}: ${kind}s referring to each other 40 deep are written in linear time and size`, Date.now() - started < 3000 && size < 2_000_000, `${Date.now() - started}ms ${size}`);
        }
    }
    // A source comment in a shared note keeps it shared (pruned per path, it was copied per reference).
    const commentedNote = await parseQuiet(Buffer.from('[^a] '.repeat(2000) + '\n\n[^a]: <!-- c --> ' + 'word '.repeat(2000) + '\n'), 'md');
    for (const format of ['text', 'rtf', 'docx', 'odt', 'chunks', 'epub'] as const) {
        const value = (await OfficeGenerator.generate(commentedNote.ast!, format, { onWarning: () => {} } as any)).value as any;
        const size = typeof value === 'string' ? value.length : Array.isArray(value) ? JSON.stringify(value).length : Object.values(unzipSync(new Uint8Array(value))).reduce((n: number, f: any) => n + f.length, 0);
        check(`${format}: a shared note holding a source comment is written once`, size < (format === 'docx' ? 600_000 : 400_000), `${size}`);
    }
    // One picture shown many times: HTML inlines up to a document budget, EPUB packages it once, RTF
    // encodes it without running out of memory.
    const bigPicture = { type: 'image', mimeType: 'image/png', name: 'big.png', extension: 'png', data: Buffer.concat([Buffer.from('89504e470d0a1a0a', 'hex'), Buffer.alloc(30_000_000)]).toString('base64') };
    const manyPictures: any = { ...astWith(Array.from({ length: 12 }, () => ({ type: 'paragraph', children: [{ type: 'image', metadata: { attachmentName: 'big.png' } }] }))), attachments: [bigPicture] };
    for (const format of ['html', 'epub', 'rtf', 'md'] as const) {
        const warnings: any[] = [];
        let outcome = 'ok';
        try { await OfficeGenerator.generate(manyPictures, format, { onWarning: (w: any) => warnings.push(w) } as any); } catch (e: any) { outcome = String(e?.message).slice(0, 80); }
        check(`${format}: one 30 MB picture shown 12 times is written within the inline budget`, outcome === 'ok', outcome);
    }
    // Markdown: an inline source comment cannot end its line's block; srcset and ping are checked per URL;
    // a reference definition repeats within a budget.
    const inlineComment = (await OfficeGenerator.generate(astWith([{ type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: 'h' }, { type: 'comment', text: 'x\n\n<img src=x onerror=alert(1)>\n\n', metadata: { sourceSyntax: 'html' } }] }]), 'md', { onWarning: () => {} } as any)).value as string;
    const reread = await parseQuiet(Buffer.from(inlineComment), 'md');
    check('md: an inline source comment with a blank line stays in its line', !JSON.stringify(reread.ast?.content).includes('"image"') && !/\n\n<img/.test(inlineComment), inlineComment);
    const srcsetHtml = (await OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [{ type: 'image', metadata: { url: 'https://ok/a.png' }, htmlAttributes: { srcset: 'https://ok/a.png 1x, //evil/b.png 2x, \\\\evil\\c.png 3x, data:image/png;base64,iVBOR= 4x', ping: 'https://ok/p //evil/q javascript:x' } }] }]), 'html', { onWarning: () => {} } as any)).value as string;
    check('html: srcset and ping are checked per URL', !/"\/\/evil|\\\\evil|javascript/.test(srcsetHtml) && srcsetHtml.includes('https://evil/b.png 2x') && srcsetHtml.includes('data:image/png;base64,iVBOR= 4x'), srcsetHtml.match(/srcset="[^"]*"|ping="[^"]*"/g)?.join(' '));
    const references = await parseQuiet(Buffer.from('[a][r] '.repeat(2000) + '\n\n[r]: https://e.com/' + 'x'.repeat(100_000) + '\n'), 'md');
    const referencesHtml = (await OfficeGenerator.generate(references.ast!, 'html', { onWarning: () => {} } as any)).value as string;
    check('md: a long reference target used 2000 times repeats within a budget', referencesHtml.length < 40_000_000, `${referencesHtml.length}`);

    // PDF: layout, marked content and annotations take time linear in a page (each was quadratic in
    // it, and pages sharing one content stream multiplied that at no size cost); RTF: a picture's
    // groups and a paragraph's notes are serialized once.
    const pdfOf = (content: string, pages: number, extraPage = '', extraObjects: string[] = []) => {
        const { deflateSync } = require('zlib') as typeof import('zlib');
        const data = deflateSync(Buffer.from(content));
        const objs: (string | Buffer)[] = [
            '<< /Type /Catalog /Pages 2 0 R >>',
            `<< /Type /Pages /Kids [${Array.from({ length: pages }, (_, i) => `${5 + extraObjects.length + i} 0 R`).join(' ')}] /Count ${pages} >>`,
            '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
            Buffer.concat([Buffer.from(`<< /Length ${data.length} /Filter /FlateDecode >>\nstream\n`), data, Buffer.from('\nendstream')]),
            ...extraObjects,
        ];
        for (let i = 0; i < pages; i++) objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 3200000 2000] /Resources << /Font << /F1 3 0 R >> >> /Contents 4 0 R ${extraPage} >>`);
        const parts: Buffer[] = [Buffer.from('%PDF-1.7\n')];
        const offsets: number[] = [];
        let at = parts[0].length;
        objs.forEach((o, i) => { const b = Buffer.concat([Buffer.from(`${i + 1} 0 obj\n`), typeof o === 'string' ? Buffer.from(o) : o, Buffer.from('\nendobj\n')]); offsets.push(at); at += b.length; parts.push(b); });
        parts.push(Buffer.from(`xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objs.length + 1} /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`));
        return Buffer.concat(parts);
    };
    await timed('pdf: 4 pages sharing one line of 16,000 spaced words lay out', () => parseQuiet(pdfOf('BT /F1 10 Tf 0 50 Td\n' + '(ab) Tj 100 0 Td\n'.repeat(16000) + 'ET', 4), 'pdf'));
    await timed('pdf: 20,000 text items inside content nested 20,000 deep are read', () => parseQuiet(pdfOf('/Span <</MCID 1>> BDC\n'.repeat(20000) + 'BT /F1 10 Tf 0 1900 Td\n' + '(a) Tj 0 -0.09 Td\n'.repeat(20000) + 'ET', 1), 'pdf'));
    const quads = Array.from({ length: 20000 }, (_, i) => `${i} 1000 ${i + 1} 1000 ${i} 990 ${i + 1} 990`).join(' ');
    await timed('pdf: a highlight of 20,000 quads and a link listed 20,000 times over 20,000 words are read', () => parseQuiet(pdfOf('BT /F1 10 Tf 0 995 Td\n' + '(ab) Tj 3 0 Td\n'.repeat(20000) + 'ET', 1,
        `/Annots [6 0 R ${'5 0 R '.repeat(20000)}]`,
        ['<< /Type /Annot /Subtype /Link /Rect [0 990 100 1000] /A << /S /URI /URI (https://example.com/) >> >>', `<< /Type /Annot /Subtype /Highlight /Rect [0 990 20000 1000] /QuadPoints [${quads}] >>`]), 'pdf'));
    const pictRtf = await heapBudget('rtf: a picture of 8,000 groups is read', () => parseQuiet(Buffer.from('{\\rtf1\\ansi {\\pict\\pngblip ' + '{}'.repeat(8000) + '}}'), 'rtf', { extractAttachments: true, includeRawContent: true }));
    check('rtf: a picture of 8,000 groups is read without error', !(pictRtf as any).error, (pictRtf as any).error);
    await timed('rtf: a paragraph of 80,000 footnotes is read', () => parseQuiet(Buffer.from('{\\rtf1\\ansi ' + 'word {\\super\\chftn}{\\footnote\\pard {\\super\\chftn} note}'.repeat(80000) + '\\par}'), 'rtf'));

    // ODF, PPTX and charts: a repeated chart cell, labels per series, an object or comments part named
    // many times, list-style lookups, chart ranges and runs of spaces are all bounded (each took
    // minutes or gigabytes, or crashed the process, from a few KB).
    const odfNs = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"';
    const odfOf = (kind: 'text' | 'spreadsheet', body: string, extra: Record<string, string> = {}, automatic = '') => {
        const mime = `application/vnd.oasis.opendocument.${kind}`;
        const files: Record<string, Uint8Array> = {
            mimetype: enc(mime),
            'META-INF/manifest.xml': enc(`<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mime}"/></manifest:manifest>`),
            'content.xml': enc(`<?xml version="1.0" encoding="UTF-8"?><office:document-content ${odfNs}><office:automatic-styles>${automatic}</office:automatic-styles><office:body><office:${kind}>${body}</office:${kind}></office:body></office:document-content>`),
        };
        for (const [k, v] of Object.entries(extra)) files[k] = enc(v);
        return Buffer.from(zipSync(files));
    };
    const chartDoc = (inner: string) => `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${odfNs}><office:body><office:chart><chart:chart>${inner}</chart:chart></office:chart></office:body></office:document-content>`;
    await heapBudget('odt: a chart cell repeated 100,000,000 times is read', () => parseQuiet(odfOf('text', '<text:p>hi</text:p><draw:frame><draw:object xlink:href="./Object 1"/></draw:frame>',
        { 'Object 1/content.xml': chartDoc('<table:table><table:table-header-rows><table:table-row><table:table-cell/><table:table-cell table:number-columns-repeated="100000000"><text:p>S</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-row><table:table-cell><text:p>L</text:p></table:table-cell><table:table-cell table:number-columns-repeated="100000000" office:value="1"/></table:table-row></table:table>') }), 'odt'));
    const rows = '<table:table-row><table:table-cell><text:p>L</text:p></table:table-cell><table:table-cell office:value="1"/></table:table-row>'.repeat(2000);
    await timed('odt: one chart object named by 20,000 frames is read', () => parseQuiet(odfOf('text', `<text:p>${'<draw:frame><draw:object xlink:href="./Object 1"/></draw:frame>'.repeat(20000)}</text:p>`,
        { 'Object 1/content.xml': chartDoc(`<table:table><table:table-header-rows><table:table-row><table:table-cell/><table:table-cell><text:p>S</text:p></table:table-cell></table:table-row></table:table-header-rows>${rows}</table:table>`) }), 'odt'));
    await timed('odt: 20,000 lists among 20,000 list styles are read', () => parseQuiet(odfOf('text', '<text:list text:style-name="X"/>'.repeat(20000), {}, '<text:list-style style:name="s"/>'.repeat(20000)), 'odt'));
    await heapBudget('ods: 2,000 chart series over a range of 50,000 cells are resolved', () => parseQuiet(odfOf('spreadsheet', '<table:table table:name="Sheet1"><table:table-row><table:table-cell table:number-columns-repeated="50000" office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row></table:table>',
        { 'Object 1/content.xml': chartDoc(`<chart:plot-area>${'<chart:series chart:values-cell-range-address="Sheet1.A1:.ZZZZ9"/>'.repeat(2000)}<chart:categories table:cell-range-address="Sheet1.A1:.ZZZZ9"/></chart:plot-area>`) }), 'ods', { extractAttachments: true }));
    const spaced = await parseQuiet(odfOf('text', `<text:p>${'<text:s text:c="10000"/>'.repeat(50000)}</text:p>`), 'odt');
    const spacedText = (await OfficeGenerator.generate(spaced.ast!, 'text', { onWarning: () => {} } as any)).value as string;
    check('odt: runs of spaces are bounded in all', spacedText.length < 20_000_000, `${spacedText.length}`);
    const pns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
    const pptxOf = (extra: Record<string, string>) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'),
        '_rels/.rels': enc('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>'),
        'ppt/presentation.xml': enc(`<?xml version="1.0"?><p:presentation ${pns}><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst></p:presentation>`),
        'ppt/slides/slide1.xml': enc(`<?xml version="1.0"?><p:sld ${pns}><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>hi</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
        ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, enc(v)])),
    }));
    await heapBudget('pptx: a chart of 20,000 series and 30,000 labels is read', () => parseQuiet(pptxOf({ 'ppt/charts/chart1.xml': `<?xml version="1.0"?><c:chartSpace ${pns}><c:chart><c:plotArea><c:barChart><c:ser><c:cat>${'<c:v>1</c:v>'.repeat(30000)}</c:cat></c:ser>${'<c:ser/>'.repeat(20000)}</c:barChart></c:plotArea></c:chart></c:chartSpace>` }), 'pptx', { extractAttachments: true }));
    const sharedComments = await heapBudget('pptx: one comments part named 2,000 times is read', () => parseQuiet(pptxOf({
        'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${Array.from({ length: 2000 }, (_, i) => `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>`).join('')}</Relationships>`,
        'ppt/comments/comment1.xml': `<?xml version="1.0"?><p:cmLst ${pns}>${'<p:cm authorId="0"><p:text>c</p:text></p:cm>'.repeat(20000)}</p:cmLst>`,
    }), 'pptx')) as any;
    check('pptx: a comments part named 2,000 times is attached once', sharedComments.ast?.content[0]?.comments?.length === 20000, `${sharedComments.ast?.content[0]?.comments?.length} ${sharedComments.error}`);

    // XML readers: content nested in its own kind is read once per level, not again at each (notes in
    // notes doubled per level; tables in tables, text boxes in text boxes, shared-string runs and PPTX
    // paragraphs grew with the square of the depth), and a document's XML holds a bounded number of
    // elements (32 KB of empty elements filled a 4 GB heap).
    let notes = '<text:p>x</text:p>';
    for (let i = 0; i < 24; i++) notes = `<text:p>a<text:note text:id="n${i}"><text:note-body>${notes}</text:note-body></text:note></text:p>`;
    await timed('odt: footnotes nested 24 deep are read', () => parseQuiet(odfOf('text', notes), 'odt'));
    let annotations = '<text:p>x</text:p>';
    for (let i = 0; i < 24; i++) annotations = `<text:p>a<office:annotation><dc:creator xmlns:dc="http://purl.org/dc/elements/1.1/">A</dc:creator>${annotations}</office:annotation></text:p>`;
    await timed('odt: comments nested 24 deep are read', () => parseQuiet(odfOf('text', annotations), 'odt'));
    let cells = '<text:p>x</text:p>';
    for (let i = 0; i < 300; i++) cells = `<table:table><table:table-row><table:table-cell>${cells}</table:table-cell></table:table-row></table:table>`;
    await timed('ods: tables nested 300 deep are read', () => parseQuiet(odfOf('spreadsheet', cells.replace('<table:table>', '<table:table table:name="S">')), 'ods'));
    let boxes = '<w:p><w:r><w:t>x</w:t></w:r></w:p>';
    for (let i = 0; i < 40; i++) boxes = `<w:p><w:r><w:pict><v:shape xmlns:v="urn:schemas-microsoft-com:vml"><v:textbox><w:txbxContent>${boxes}</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>`;
    await timed('docx: text boxes nested 40 deep are read', () => parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/<w:body>/, `<w:body>${boxes}`)); }), 'docx'));
    let runs = '<t>x</t>';
    for (let i = 0; i < 20000; i++) runs = `<r>${runs}</r>`;
    await timed('xlsx: shared-string runs nested 20,000 deep are read', () => parseQuiet(repack('test.xlsx', z => { z['xl/sharedStrings.xml'] = enc(`<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si>${runs}</si></sst>`); }), 'xlsx'));
    let tableRows = '<a:tc><a:txBody><a:p><a:r><a:t>x</a:t></a:r></a:p></a:txBody></a:tc>';
    for (let i = 0; i < 600; i++) tableRows = `<a:tr><a:tc><a:txBody><a:p><a:r><a:t>c</a:t></a:r></a:p></a:txBody>${tableRows}</a:tc></a:tr>`;
    await timed('pptx: table rows and cells nested 600 deep are read', () => parseQuiet(pptxOf({ 'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${pns}><p:cSld><p:spTree><p:graphicFrame><a:graphic><a:graphicData><a:tbl>${tableRows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>` }), 'pptx'));
    const elementBomb = await parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/<w:body>/, `<w:body>${'<z/>'.repeat(3_000_000)}`)); }), 'docx');
    check('docx: 3,000,000 empty elements fail with XML_ELEMENT_LIMIT_EXCEEDED', /XML element limit exceeded/.test(elementBomb.error), elementBomb.error.slice(0, 120));
    const raised = await parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/<w:body>/, `<w:body>${'<z/>'.repeat(300_000)}`)); }), 'docx', { decompressionLimits: { maxXmlElements: 400_000 } });
    const lowered = await parseQuiet(repack('test.docx', z => z), 'docx', { decompressionLimits: { maxXmlElements: 100 } });
    check('docx: maxXmlElements is the limit a parse is held to', !raised.error && /XML element limit exceeded/.test(lowered.error), `${raised.error} | ${lowered.error}`);

    // Spans are held to what a browser allows, and never below 1.
    const gridSpan = await parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace(/<w:body>/, '<w:body><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2147483647"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:gridSpan w:val="-5"/></w:tcPr><w:p><w:r><w:t>neg</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')); }), 'docx');
    const spans = (ast: any) => { const out: any[] = []; const walk = (ns: any[]) => ns?.forEach((n: any) => { if (n.type === 'cell') out.push([n.metadata?.colSpan, n.metadata?.rowSpan, n.metadata?.col]); walk(n.children); }); walk(ast?.content); return out; };
    check('docx: a gridSpan of 2^31 is held to 1000, a negative one to 1', JSON.stringify(spans(gridSpan.ast).slice(0, 2)) === JSON.stringify([[1000, null, 0], [null, null, 1000]].map(r => r.map(v => v ?? undefined))), JSON.stringify(spans(gridSpan.ast).slice(0, 2)));
    const htmlSpans = await parseQuiet(Buffer.from('<table><tr><td rowspan="-3" colspan="0">a</td><td rowspan="99999999" colspan="5000">b</td></tr></table>'), 'html');
    check('html: spans below 1 are 1, above the limits are the limits', JSON.stringify(spans(htmlSpans.ast)) === JSON.stringify([[undefined, undefined, undefined], [1000, 65534, undefined]]), JSON.stringify(spans(htmlSpans.ast)));
    const odtSpans = await parseQuiet(repack('test.odt', z => { z['content.xml'] = enc(new TextDecoder().decode(z['content.xml']).replace(/<office:text\b[^>]*>/, m => m + '<table:table table:name="T"><table:table-row><table:table-cell table:number-columns-spanned="2147483647" table:number-rows-spanned="99999999"><text:p>x</text:p></table:table-cell></table:table-row></table:table>')); }), 'odt');
    check('odt: spans of billions are held to the limits', spans(odtSpans.ast).some(([c, r]) => c === 1000 && r === 65534), JSON.stringify(spans(odtSpans.ast)));

    // The XML library reports nothing to the console: an entity it cannot resolve is kept as written.
    const errorLog = console.error;
    let logged = 0;
    console.error = () => { logged++; };
    let entities: Awaited<ReturnType<typeof parseQuiet>>;
    try {
        entities = await parseQuiet(repack('test.docx', z => { z['word/document.xml'] = enc(new TextDecoder().decode(z['word/document.xml']).replace('<w:t>', '<w:t>' + '&nbsp;&foo;'.repeat(50))); }), 'docx');
    } finally { console.error = errorLog; }
    check('xml: unresolved entities write nothing to the console and are kept', logged === 0 && JSON.stringify(entities.ast?.content).includes('&foo;'), `${logged} ${entities.error}`);

    // EPUB: a picture shown many times is one attachment, and a chapter the spine lists many times is read once.
    const png = Buffer.concat([Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489', 'hex'), Buffer.alloc(1_000_000)]);
    const epub = (images: number, spine: number) => Buffer.from(zipSync({
        'mimetype': enc('application/epub+zip'),
        'META-INF/container.xml': enc('<?xml version="1.0"?><container><rootfiles><rootfile full-path="content.opf"/></rootfiles></container>'),
        'content.opf': enc(`<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata><manifest><item id="c" href="c.xhtml" media-type="application/xhtml+xml"/><item id="i" href="a.png" media-type="image/png"/></manifest><spine>${'<itemref idref="c"/>'.repeat(spine)}</spine></package>`),
        'c.xhtml': enc(`<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><body><p>${'x'.repeat(100000)}</p>${'<img src="a.png" alt="a"/>'.repeat(images)}</body></html>`),
        'a.png': new Uint8Array(png),
    }));
    const heapBefore = process.memoryUsage().heapUsed;
    const shownOften = await parseQuiet(epub(3000, 1), 'epub', { extractAttachments: true });
    const heapGrowth = process.memoryUsage().heapUsed - heapBefore;
    const pictures: any[] = [];
    const collect = (ns: any[] | undefined) => ns?.forEach((n: any) => { if (n.type === 'image') pictures.push(n.metadata?.attachmentName); collect(n.children); });
    collect(shownOften.ast?.content);
    check('epub: a picture shown 3000 times is one attachment every picture names', shownOften.ast?.attachments.length === 1 && pictures.length === 3000 && pictures.every(n => n === shownOften.ast!.attachments[0].name) && heapGrowth < 500_000_000, `${shownOften.ast?.attachments.length} ${pictures.length} ${heapGrowth}`);
    const listedOften = await parseQuiet(epub(0, 5000), 'epub');
    check('epub: a chapter the spine lists 5000 times is read once', listedOften.ast?.content.length === 1, `${listedOften.ast?.content.length} ${listedOften.error}`);

    const warned = async (buffer: Buffer, fileType: string, extra: object = {}) => {
        const codes: string[] = [];
        const result = await parseQuiet(buffer, fileType, { ...extra, onWarning: (issue: any) => codes.push(issue.code) });
        return { ...result, codes };
    };
    const epubPages = await parseQuiet(epub(3000, 1), 'epub', { decompressionLimits: { maxXmlElements: 1000 } });
    check('epub: a chapter counts against maxXmlElements', /XML element limit exceeded/.test(epubPages.error), epubPages.error.slice(0, 120));

    // Parts holding their own kind (notes in notes, comments in comments, fonts, styles, drawings and
    // chart series in themselves, math in math) are read once per element, not once per level above it,
    // and a numbering definition is shared by the lists naming it rather than copied into each.
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';
    const docxOf = (body: string, extra: Record<string, string | Uint8Array> = {}) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        'word/document.xml': enc(`<?xml version="1.0"?><w:document ${W}><w:body>${body}</w:body></w:document>`),
        ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, typeof v === 'string' ? enc(v) : v])),
    }));
    const xlsxOf = (extra: Record<string, string>) => Buffer.from(zipSync({
        '[Content_Types].xml': enc('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'),
        'xl/workbook.xml': enc('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>'),
        'xl/_rels/workbook.xml.rels': enc('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
        'xl/worksheets/sheet1.xml': enc('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>v</t></is></c></row></sheetData></worksheet>'),
        ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, enc(v)])),
    }));
    const nest = (inner: string, depth: number, wrap: (s: string, i: number) => string) => { for (let i = 0; i < depth; i++) inner = wrap(inner, i); return inner; };
    const empties = (n: number) => '<x/>'.repeat(n);
    const bodyText = '<w:p><w:r><w:t>body</w:t></w:r></w:p>';
    await timed('docx: footnotes nested 2,000 deep around 20,000 elements are read', () => parseQuiet(docxOf(bodyText, { 'word/footnotes.xml': `<?xml version="1.0"?><w:footnotes ${W}>${nest('<w:p><w:r><w:t>x</w:t></w:r></w:p>' + empties(20000), 2000, (s, i) => `<w:footnote w:id="${i + 5}">${s}</w:footnote>`)}</w:footnotes>` }), 'docx'));
    await timed('docx: comments nested 2,000 deep around 20,000 elements are read', () => parseQuiet(docxOf(bodyText, { 'word/comments.xml': `<?xml version="1.0"?><w:comments ${W}>${nest('<w:p><w:r><w:t>x</w:t></w:r></w:p>' + empties(20000), 2000, (s, i) => `<w:comment w:id="${i + 5}">${s}</w:comment>`)}</w:comments>` }), 'docx'));
    await timed('docx: drawings nested 2,000 deep in a run are read', () => parseQuiet(docxOf(`<w:p><w:r><w:t>x</w:t>${nest(empties(20000), 2000, s => `<w:drawing>${s}</w:drawing>`)}</w:r></w:p>`), 'docx', { extractAttachments: true }));
    await timed('docx: 20,000 lists naming one numbering of 4,000 levels are read', () => parseQuiet(docxOf('<w:p><w:r><w:t>x</w:t></w:r></w:p>', { 'word/numbering.xml': `<?xml version="1.0"?><w:numbering ${W}><w:abstractNum w:abstractNumId="1">${'<w:lvl w:ilvl="0"/>'.repeat(4000)}</w:abstractNum>${Array.from({ length: 20000 }, (_, i) => `<w:num w:numId="${i}"><w:abstractNumId w:val="1"/></w:num>`).join('')}</w:numbering>` }), 'docx'));
    const commentRels = (target: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="c1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="${target}"/></Relationships>`;
    await heapBudget('xlsx: comments nested 2,000 deep around 200 KB of text are read', () => parseQuiet(xlsxOf({ 'xl/worksheets/_rels/sheet1.xml.rels': commentRels('../comments1.xml'), 'xl/comments1.xml': `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><commentList>${nest(`<text><t>${'A'.repeat(200000)}</t></text>`, 2000, s => `<comment ref="A1"><text><t>x</t></text>${s}</comment>`)}</commentList></comments>` }), 'xlsx'));
    await timed('xlsx: fonts nested 2,000 deep around 20,000 elements are read', () => parseQuiet(xlsxOf({ 'xl/styles.xml': `<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts>${nest(empties(20000), 2000, s => `<font>${s}</font>`)}</fonts></styleSheet>` }), 'xlsx'));
    await heapBudget('pptx: comments nested 2,000 deep around 200 KB of text are read', () => parseQuiet(pptxOf({ 'ppt/slides/_rels/slide1.xml.rels': commentRels('../comments/comment1.xml'), 'ppt/comments/comment1.xml': `<?xml version="1.0"?><p:cmLst ${pns}>${nest(`<p:text>${'A'.repeat(200000)}</p:text>`, 2000, s => `<p:cm authorId="0"><p:text>x${s}</p:text></p:cm>`)}</p:cmLst>` }), 'pptx'));
    const mns = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
    await timed('pptx: 200 paragraphs of math nested 1,000 deep are read', () => parseQuiet(pptxOf({ 'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${pns} ${mns}><p:cSld><p:spTree><p:sp><p:txBody>${`<a:p><x>${nest(`<m:r><m:t>${'A'.repeat(1000)}</m:t></m:r>`, 1000, s => `<m:oMath>${s}</m:oMath>`)}</x></a:p>`.repeat(200)}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>` }), 'pptx'));
    await timed('pptx: chart series nested 2,000 deep around 20,000 values are read', () => parseQuiet(pptxOf({ 'ppt/charts/chart1.xml': `<?xml version="1.0"?><c:chartSpace ${pns}><c:chart>${nest('<c:v>1</c:v>'.repeat(20000), 2000, s => `<c:ser><c:val>${s}</c:val></c:ser>`)}</c:chart></c:chartSpace>` }), 'pptx', { extractAttachments: true }));
    await timed('odt: styles nested 2,000 deep around 60,000 elements are read', () => parseQuiet(odfOf('text', '<text:p>hi</text:p>', {}, nest(empties(60000), 2000, (s, i) => `<style:style style:name="s${i}">${s}</style:style>`)), 'odt'));

    // rawContent: a node's markup holds everything nested in it, so nested tables repeat it at every
    // level and repeated cells once per copy; it is serialized once per cell and bounded in all.
    const nestedTables = odfOf('text', nest('<text:p>x</text:p>' + '<text:span/>'.repeat(20000), 300, s => `<table:table><table:table-row><table:table-cell>${s}</table:table-cell></table:table-row></table:table>`));
    await heapBudget('odt: rawContent of tables nested 300 deep around 20,000 elements is read', () => timed('odt: rawContent of tables nested 300 deep is read', () => parseQuiet(nestedTables, 'odt', { includeRawContent: true })));
    const textBoxes = docxOf(nest('<w:p><w:r><w:t>x</w:t></w:r></w:p>' + empties(150000), 500, s => `<w:p><w:pict><w:txbxContent>${s}</w:txbxContent></w:pict></w:p>`));
    const boundedRaw = await warned(textBoxes, 'docx', { includeRawContent: true, decompressionLimits: { maxRawContentLength: 400_000 } });
    let rawTotal = 0;
    const sumRaw = (ns: any[] | undefined, seen = new Set<any>()) => ns?.forEach((n: any) => { if (seen.has(n)) return; seen.add(n); rawTotal += n.rawContent?.length ?? 0; sumRaw(n.children, seen); });
    sumRaw(boundedRaw.ast?.content);
    check('docx: rawContent stops at maxRawContentLength with RAW_CONTENT_LIMIT_EXCEEDED', !boundedRaw.error && rawTotal > 0 && rawTotal <= 400_000 && boundedRaw.codes.includes('RAW_CONTENT_LIMIT_EXCEEDED'), `${rawTotal} ${boundedRaw.codes} ${boundedRaw.error}`);
    const repeatedRaw = await warned(odfOf('text', `<table:table><table:table-row><table:table-cell table:number-columns-repeated="1000"><text:p>${'<text:span>y</text:span>'.repeat(2000)}</text:p></table:table-cell></table:table-row></table:table>`), 'odt', { includeRawContent: true, decompressionLimits: { maxRawContentLength: 1_000_000 } });
    rawTotal = 0;
    sumRaw(repeatedRaw.ast?.content);
    check('odt: a repeated cell\'s rawContent counts once per copy', rawTotal <= 1_000_000 && repeatedRaw.codes.includes('RAW_CONTENT_LIMIT_EXCEEDED'), `${rawTotal} ${repeatedRaw.codes}`);

    // Repeated ODF cells copy their content within a budget: 700 bytes asked for 100,000 copies of a
    // 2,000-span cell, 200 MB of output that ran generators out of memory, and walking each copy's
    // shared content made the parse itself take repeats x content.
    const repeatedCell = (repeats: number, rowRepeats = 1) => odfOf('spreadsheet', `<table:table table:name="S"><table:table-row table:number-rows-repeated="${rowRepeats}"><table:table-cell table:number-columns-repeated="${repeats}"><text:p>${'<text:span>y</text:span>'.repeat(2000)}</text:p></table:table-cell><table:table-cell><text:p>after</text:p></table:table-cell></table:table-row></table:table>`);
    for (const [label, doc] of [['a cell repeated 100,000 times', repeatedCell(100000)], ['a row of it repeated 1,000,000 times', repeatedCell(1000, 1000000)]] as const) {
        let html = '';
        const started = Date.now();
        const repeated = await warned(doc, 'ods');
        if (repeated.ast) html = (await OfficeGenerator.generate(repeated.ast, 'html', { onWarning: () => {} } as any)).value as string;
        check(`ods: ${label} is copied within maxRepeatedContent, with REPEATED_CONTENT_LIMIT_EXCEEDED`, !repeated.error && html.length < 20_000_000 && Date.now() - started < 5000 && repeated.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED') && html.includes('after'), `${html.length} ${Date.now() - started}ms ${repeated.codes} ${repeated.error}`);
    }
    const cellsOf = (ast: any) => { const out: any[] = []; const walk = (ns: any[]) => ns?.forEach((n: any) => { if (n.type === 'cell') out.push(n); else walk(n.children); }); walk(ast?.content); return out; };
    const small = await warned(repeatedCell(100), 'ods', { decompressionLimits: { maxRepeatedContent: 100_000 } });
    const large = await warned(repeatedCell(100), 'ods', { decompressionLimits: { maxRepeatedContent: 100_000_000 } });
    const smallCells = cellsOf(small.ast), largeCells = cellsOf(large.ast);
    check('ods: maxRepeatedContent is the limit, and later cells keep their columns', smallCells.length < 10 && largeCells.length === 101 && smallCells[smallCells.length - 1]?.metadata?.col === 100 && !large.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${smallCells.length} ${largeCells.length} ${smallCells[smallCells.length - 1]?.metadata?.col}`);
    const xlsxCellsDoc = xlsxOf({ 'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${'<row><c t="inlineStr"><is><t>x</t></is></c><c><v>1</v></c></row>'.repeat(2500)}</sheetData></worksheet>` });
    const xlsxCells = await warned(xlsxCellsDoc, 'xlsx', { decompressionLimits: { maxTableCells: 1000 } });
    // maxTableCells, plus one cell per byte of the document (see budgetUtils).
    check('xlsx: a sheet holds maxTableCells cells (plus one a byte), with TABLE_CELL_LIMIT_EXCEEDED', xlsxCellsDoc.length < 4000 && cellsOf(xlsxCells.ast).length === 1000 + xlsxCellsDoc.length && xlsxCells.codes.includes('TABLE_CELL_LIMIT_EXCEEDED'), `${cellsOf(xlsxCells.ast).length} ${xlsxCellsDoc.length} ${xlsxCells.codes}`);
    // A workbook whose cells do not compress to less than a byte each (every real one) is read whole
    // under any limit: 4 MB of 1.2 million cells lost its last 200,000 to the default million.
    const distinctCellsDoc = xlsxOf({ 'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: 3000 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${(i * 7919) % 100003}.${i % 97}</v></c></row>`).join('')}</sheetData></worksheet>` });
    const distinctCells = await warned(distinctCellsDoc, 'xlsx', { decompressionLimits: { maxTableCells: 1000 } });
    check('xlsx: a workbook of more cells than maxTableCells, but fewer than its bytes, is read whole', distinctCellsDoc.length > 2000 && cellsOf(distinctCells.ast).length === 3000 && !distinctCells.codes.includes('TABLE_CELL_LIMIT_EXCEEDED'), `${cellsOf(distinctCells.ast).length} ${distinctCellsDoc.length} ${distinctCells.codes}`);

    // Shared strings are read as the sheets are, without building XML elements: a million distinct
    // strings (5.5 MB) was refused by the element budget. Plain strings cost an element per eight; a
    // rich string's runs one each.
    const sharedStringCells = (n: number) => `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: n }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>${i}</v></c></row>`).join('')}</sheetData></worksheet>`;
    const sst = (items: string) => `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${items}</sst>`;
    const plainStrings = await warned(xlsxOf({ 'xl/sharedStrings.xml': sst(Array.from({ length: 3000 }, (_, i) => `<si><t>item ${i}</t></si>`).join('')), 'xl/worksheets/sheet1.xml': sharedStringCells(3000) }), 'xlsx', { decompressionLimits: { maxXmlElements: 1000 } });
    const plainTexts = cellsOf(plainStrings.ast).map((c: any) => c.children?.[0]?.text ?? c.children?.[0]?.children?.[0]?.text);
    check('xlsx: 3,000 plain shared strings are read within 1,000 elements', !plainStrings.error && plainTexts.length === 3000 && plainTexts[2999] === 'item 2999', `${plainStrings.error} ${plainTexts.length} ${plainTexts[2999]}`);
    const richStrings = await warned(xlsxOf({ 'xl/sharedStrings.xml': sst('<si><r><t>a</t></r><r><t>b</t></r></si>'.repeat(1000)), 'xl/worksheets/sheet1.xml': sharedStringCells(1) }), 'xlsx', { decompressionLimits: { maxXmlElements: 1000 } });
    check('xlsx: 2,000 rich runs count against maxXmlElements', /XML element limit exceeded/.test(richStrings.error ?? ''), String(richStrings.error).slice(0, 100));
    // What an XML reader gives: references decoded, CDATA as written, phonetic readings (`rPh`) not the
    // text, and a toggle a run turns off (`<b val="0"/>`) off.
    const readStrings = await warned(xlsxOf({
        'xl/sharedStrings.xml': sst('<si><t>A &amp; B &#x263A;</t><rPh sb="0" eb="1"><t>PHONETIC</t></rPh></si><si><r><rPr><b val="0"/><i/><sz val="9"/><color rgb="FF112233"/><rFont val="Arial &amp; Co"/></rPr><t xml:space="preserve">plain </t></r><r><rPr><b/><vertAlign val="superscript"/></rPr><t><![CDATA[<bold>]]></t></r></si>'),
        'xl/worksheets/sheet1.xml': sharedStringCells(2),
    }), 'xlsx');
    const readCells = cellsOf(readStrings.ast);
    const runsOf = (cell: any): any[] => { const out: any[] = []; const walk = (ns: any[]) => ns?.forEach((n: any) => { if (n.type === 'text') out.push(n); else walk(n.children); }); walk(cell?.children); return out; };
    const [plainRun] = runsOf(readCells[0]);
    const [first, second] = runsOf(readCells[1]);
    check('xlsx: a shared string is its text, without its phonetic reading', plainRun?.text === 'A & B \u263A', JSON.stringify(plainRun?.text));
    check('xlsx: a rich shared string keeps each run and its formatting', first?.text === 'plain ' && !first.formatting?.bold && first.formatting?.italic && first.formatting?.size === '9pt' && first.formatting?.color === '#112233' && first.formatting?.font === 'Arial & Co' && second?.text === '<bold>' && second.formatting?.bold && second.formatting?.superscript, JSON.stringify([first, second]));
    // The repeat budget grows with the document: maxRepeatedContent plus 16 characters a byte.
    const repeatedStringDoc = xlsxOf({ 'xl/sharedStrings.xml': sst(`<si><t>${'n'.repeat(500)}</t></si>`), 'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: 400 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join('')}</sheetData></worksheet>` });
    const repeatedIssues: any[] = [];
    await parseQuiet(repeatedStringDoc, 'xlsx', { decompressionLimits: { maxRepeatedContent: 1000 }, onWarning: (issue: any) => repeatedIssues.push(issue) });
    const repeatedLimit = repeatedIssues.find(issue => issue.code === 'REPEATED_CONTENT_LIMIT_EXCEEDED');
    check('xlsx: maxRepeatedContent grows by 16 characters a byte of the document', repeatedLimit?.message.includes(`(${1000 + 16 * repeatedStringDoc.length} characters`), repeatedLimit?.message.slice(0, 120));

    // Plain text lays out a table from its cells once: rendered first and then again, each nested
    // table doubled the work (26 levels took minutes).
    let deepTable: any = { type: 'paragraph', children: [{ type: 'text', text: 'x' }] };
    for (let d = 0; d < 26; d++) deepTable = { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [deepTable] }] }] };
    await timed('text: tables nested 26 deep are written', () => OfficeGenerator.generate(astWith([deepTable]), 'text' as any, { onWarning: () => {} } as any));

    let deepSheet: any = { type: 'paragraph', children: [{ type: 'text', text: 'x' }] };
    for (let d = 0; d < 3000; d++) deepSheet = { type: 'sheet', metadata: { sheetName: 's' }, children: [{ type: 'row', children: [{ type: 'cell', children: [deepSheet] }] }] };
    let deepError: any;
    try { await OfficeGenerator.generate(astWith([deepSheet]), 'html' as any, { onWarning: () => {} } as any); } catch (e) { deepError = e; }
    check('generators: an AST nested past the stack is MAX_NESTING_DEPTH_EXCEEDED', deepError?.officeIssue?.code === 'MAX_NESTING_DEPTH_EXCEEDED', String(deepError));

    // Content repeated by reference is bounded in all (maxRepeatedContent): one XLSX string shown in
    // many cells, one style's font or one relationship's link given to many runs, a style's font on a
    // repeated ODF cell. Each made hundreds of MB of output, or a heap abort, from a few KB.
    const sharedStringBook = (si: string, cells: number) => xlsxOf({
        'xl/sharedStrings.xml': `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si>${si}</si></sst>`,
        'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: cells / 100 }, (_, r) => `<row r="${r + 1}">${'<c t="s"><v>0</v></c>'.repeat(100)}</row>`).join('')}</sheetData></worksheet>`,
    });
    for (const [label, si, cells] of [['a 1 MB string shown in 400 cells', `<t>${'x'.repeat(1_000_000)}</t>`, 400], ['a string of 20,000 runs shown in 50,000 cells', '<r><rPr><b/></rPr><t>x</t></r>'.repeat(20000), 50000]] as const) {
        const started = Date.now();
        const shown = await heapBudget(`xlsx: ${label} is read`, () => warned(sharedStringBook(si, cells), 'xlsx')) as any;
        const csv = shown.ast ? (await OfficeGenerator.generate(shown.ast, 'csv', { onWarning: () => {} } as any)).value as string : '';
        check(`xlsx: ${label} repeats within maxRepeatedContent, with REPEATED_CONTENT_LIMIT_EXCEEDED`, !shown.error && csv.length < 40_000_000 && Date.now() - started < 5000 && shown.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${csv.length} ${Date.now() - started}ms ${shown.codes} ${shown.error}`);
    }
    const styledRuns = await parseQuiet(xlsxOf({
        'xl/sharedStrings.xml': '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>rich</t></r></si></sst>',
        'xl/styles.xml': '<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="3"><fill/><fill/><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></fills><cellXfs count="2"><xf fontId="0" fillId="0"/><xf fontId="0" fillId="2"/></cellXfs></styleSheet>',
        'xl/worksheets/sheet1.xml': '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s" s="1"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>',
    }), 'xlsx');
    const runBackgrounds = cellsOf(styledRuns.ast).map(c => c.children?.[0]?.formatting?.backgroundColor ?? '');
    check('xlsx: cells sharing a rich string keep their own cell style', JSON.stringify(runBackgrounds) === JSON.stringify(['#00FF00', '']), JSON.stringify(runBackgrounds));
    const longFont = 'F'.repeat(65536);
    const fontRuns = await warned(docxOf(`<w:p>${'<w:r><w:rPr><w:rStyle w:val="A"/></w:rPr><w:t>x</w:t></w:r>'.repeat(7500)}</w:p>`, { 'word/styles.xml': `<?xml version="1.0"?><w:styles ${W}><w:style w:type="character" w:styleId="A"><w:rPr><w:rFonts w:ascii="${longFont}"/></w:rPr></w:style></w:styles>` }), 'docx');
    const linkRuns = await warned(docxOf(`<w:p>${'<w:hyperlink r:id="rId5"><w:r><w:t>x</w:t></w:r></w:hyperlink> '.repeat(2000)}</w:p>`, { 'word/_rels/document.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://e.com/${'a'.repeat(65536)}" TargetMode="External"/></Relationships>` }), 'docx');
    const fontCell = await warned(odfOf('spreadsheet', '<table:table table:name="S"><table:table-row><table:table-cell table:style-name="ce1" table:number-columns-repeated="5000"><text:p><text:span text:style-name="T1">x</text:span></text:p></table:table-cell></table:table-row></table:table>', {},
        `<style:style style:name="T1" style:family="text"><style:text-properties xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" fo:font-family="${longFont}"/></style:style><style:style style:name="ce1" style:family="table-cell"><style:text-properties xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" fo:font-family="${longFont}"/></style:style>`), 'ods');
    for (const [label, result] of [['docx: a 64 KB style font on 7,500 runs', fontRuns], ['docx: a 64 KB link target on 2,000 links', linkRuns], ['ods: a 64 KB cell font repeated 5,000 times', fontCell]] as const) {
        let longest = 0;
        for (const format of ['html', 'rtf'] as const) {
            const out = result.ast ? (await OfficeGenerator.generate(result.ast, format, { onWarning: () => {} } as any)).value as string : '';
            longest = Math.max(longest, out.length);
        }
        check(`${label} repeats within maxRepeatedContent, with REPEATED_CONTENT_LIMIT_EXCEEDED`, !result.error && longest > 0 && longest < 40_000_000 && result.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${longest} ${result.codes} ${result.error}`);
    }

    // PDF: a form XObject drawing another many times is read within a budget of text items and
    // operators (seven forms deep, 2.5 KB, ended the process out of memory).
    const pdfForms = (depth: number, leaf: string) => {
        const ids = Array.from({ length: depth }, (_, i) => 6 + i);
        const xobjects = (upto: number) => '<< ' + ids.slice(0, upto).map((id, i) => `/X${i} ${id} 0 R`).join(' ') + ' >>';
        const stream = (body: string, extra = '') => `<< /Length ${body.length}${extra} >>\nstream\n${body}\nendstream`;
        const objects = ['<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
            `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject ${xobjects(depth)} >> /Contents 5 0 R >>`,
            '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>', stream(`BT /F1 12 Tf 10 700 Td (page) Tj ET /X${depth - 1} Do`),
            ...ids.map((_, i) => stream(i === 0 ? leaf : `/X${i - 1} Do `.repeat(10), ` /Type /XObject /Subtype /Form /BBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> /XObject ${xobjects(i)} >>`))];
        let pdf = '%PDF-1.4\n';
        const offsets: number[] = [];
        objects.forEach((o, i) => { offsets.push(pdf.length); pdf += `${i + 1} 0 obj\n${o}\nendobj\n`; });
        const xref = pdf.length;
        pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
        return Buffer.from(pdf, 'latin1');
    };
    for (const [label, leaf] of [['text', 'BT /F1 12 Tf 10 10 Td (xy) Tj ET'], ['shapes', '0 0 1 1 re f']] as const) {
        const started = Date.now();
        const forms = await heapBudget(`pdf: forms of ${label} nested 8 deep, each drawn 10 times, are read`, () => warned(pdfForms(8, leaf), 'pdf')) as any;
        check(`pdf: forms of ${label} nested 8 deep stop at the content limit with PDF_CONTENT_LIMIT_EXCEEDED`, !forms.error && Date.now() - started < 5000 && forms.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED') && JSON.stringify(forms.ast?.content).includes('page'), `${Date.now() - started}ms ${forms.codes} ${forms.error}`);
    }
    const strictPdf = await warned(pdfForms(5, 'BT /F1 12 Tf 10 10 Td (xy) Tj ET'), 'pdf', { pdfParserConfig: { maxTextItems: 1 } });
    const roomyPdf = await warned(pdfForms(5, 'BT /F1 12 Tf 10 10 Td (xy) Tj ET'), 'pdf', { pdfParserConfig: { maxTextItems: 100_000, maxOperators: 1_000_000 } });
    check('pdf: maxTextItems is the base of the limit', strictPdf.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED') && !roomyPdf.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${strictPdf.codes} | ${roomyPdf.codes}`);
    // A table of contents' dot leaders (tagged TOCI) collapse in time linear in the entry, however its
    // dots and spaces run: an entry four times as long takes about four times as long.
    const tocMs = (text: string) => {
        const run: any = { text, x: 72, yTop: 90, yBaseline: 100, width: 300, height: 12, fontSize: 10, dir: 'ltr', angle: 0, mcid: 'm0', inArtifact: false, formatting: {} };
        const layout: any = { useTags: true, detectColumns: true, mergeHyphenatedWords: true, lineToleranceFactor: 0.35, spaceToleranceFactor: 0.25, headingDetection: 'auto', normalizeText: true, extractTextColor: false, includeBounds: true };
        const begun = performance.now();
        buildTaggedNodes({ role: 'Root', children: [{ role: 'TOCI', children: [{ type: 'content', id: 'm0' }] }] }, new Map([['m0', [run]]]), { pageNumber: 1, authoredW: 612, authoredH: 792, rotation: 0 }, computeDocContext([run], layout, '\n'), { ignoreNotes: false, listCounter: { n: 0 } });
        return performance.now() - begun;
    };
    for (const [label, lead, unit] of [['one leader of many runs of dots', 'Title', ' ....'], ['a run of spaces before three dots', 'x', ' '], ['runs of three dots', 'x', '... '], ['spaced single dots', 'x', '. ']] as const) {
        const small = tocMs(lead + unit.repeat(50_000) + '...12');
        const large = tocMs(lead + unit.repeat(200_000) + '...12');
        check(`pdf: TOC dot leaders in ${label} collapse in linear time`, large < 5000 && large < Math.max(100, small * 8), `${Math.round(small)}ms, four times as long: ${Math.round(large)}ms`);
    }

    // Templates, raw source and chunks.
    await timed('template: a part of 160,000 unclosed <w:p is rendered', () => OfficeTemplate.render(Buffer.from(zipSync({ '[Content_Types].xml': enc('<Types/>'), 'word/document.xml': enc('<w:p '.repeat(160000)) })), { data: { a: 1 } }).catch(() => undefined));
    await timed('docx: raw source of 32,000 paragraphs is read', () => parseQuiet(docxOf('<w:p><w:r><w:t>x</w:t></w:r></w:p>\n'.repeat(32000)), 'docx', { includeRawContent: true, serializeRawContent: false }));
    const nestedSource = await parseQuiet(docxOf('<w:tbl><w:tr><w:tc><w:tbl><w:tr><w:tc><w:p><w:r><w:t>in</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>out</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:br/><w:t>b</w:t></w:r></w:p>'), 'docx', { includeRawContent: true, serializeRawContent: false });
    const sources: string[] = [];
    const collectSources = (ns: any[] | undefined) => ns?.forEach((n: any) => { if (n.type === 'paragraph' && n.rawContent) sources.push(n.rawContent); collectSources(n.children); });
    collectSources(nestedSource.ast?.content);
    check('docx: raw source of a paragraph is the paragraph, nested or after a break', JSON.stringify(sources) === JSON.stringify(['<w:p><w:r><w:t>in</w:t></w:r></w:p>', '<w:p><w:r><w:t>out</w:t></w:r></w:p>', '<w:p><w:r><w:br/><w:t>b</w:t></w:r></w:p>']), JSON.stringify(sources));
    const longHeading = await parseQuiet(docxOf(`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${'H'.repeat(1_000_000)}</w:t></w:r></w:p>${'<w:p><w:r><w:t>body text</w:t></w:r></w:p>'.repeat(1000)}`), 'docx');
    const chunked = longHeading.ast ? (await OfficeGenerator.generate(longHeading.ast, 'chunks', { onWarning: () => {} } as any)).value as any : [];
    const chunkList: any[] = Array.isArray(chunked) ? chunked : chunked.chunks;
    const chunkMetadata = chunkList.reduce((n: number, c: any) => n + JSON.stringify(c.metadata).length, 0);
    check('chunks: a 1 MB heading is not repeated in each chunk\'s metadata', chunkList.length > 0 && chunkMetadata < 5_000_000, `${chunkList.length} ${chunkMetadata}`);

    // What a part, picture or definition gives each place using it is bounded too: a chart value its
    // table repeats, a picture's name and a chart's text at every frame showing them, and a LaTeX
    // heading, theorem title or beamer title at every reference or use.
    const outputOf = async (ast: any, format: string) => ast ? ((await OfficeGenerator.generate(ast, format as any, { onWarning: () => {} } as any)).value as string) : '';
    const chartValue = 'x'.repeat(100_000);
    const repeatedValues = await warned(odfOf('text', '<text:p><draw:frame><draw:object xlink:href="./Object 1"/></draw:frame></text:p>', { 'Object 1/content.xml': chartDoc(`<table:table><table:table-header-rows><table:table-row><table:table-cell/><table:table-cell table:number-columns-repeated="2500"><text:p>S</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-row><table:table-cell><text:p>L</text:p></table:table-cell><table:table-cell office:value="${chartValue}" table:number-columns-repeated="2500"/></table:table-row></table:table>`) }), 'odt', { extractAttachments: true });
    const repeatedValuesText = await outputOf(repeatedValues.ast, 'text');
    check('odt: a 100 KB chart value repeated 2,500 times repeats within maxRepeatedContent', !repeatedValues.error && repeatedValuesText.length < 20_000_000 && repeatedValues.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${repeatedValuesText.length} ${repeatedValues.codes} ${repeatedValues.error}`);
    const longName = 'a'.repeat(60_000);
    const namedPictures = await parseQuiet(docxOf('<w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId9"/></w:drawing></w:r></w:p>'.repeat(2000), { 'word/_rels/document.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${longName}.png"/></Relationships>`, [`word/media/${longName}.png`]: 'png' }), 'docx', { extractAttachments: true });
    const namedHtml = await outputOf(namedPictures.ast, 'html');
    const pictureNames = new Set<string>();
    const collectNames = (ns: any[] | undefined) => ns?.forEach((n: any) => { if (n.metadata?.attachmentName) pictureNames.add(n.metadata.attachmentName); collectNames(n.children); });
    collectNames(namedPictures.ast?.content);
    const [shownName] = [...pictureNames];
    check('docx: a 60 KB picture name shown 2,000 times is shortened alike on the picture and its attachment', !namedPictures.error && namedHtml.length < 20_000_000 && pictureNames.size === 1 && shownName.length <= 128 && shownName.endsWith('.png') && namedPictures.ast!.attachments.some(a => a.name === shownName), `${namedHtml.length} ${pictureNames.size} ${shownName?.length} ${namedPictures.ast?.attachments.map(a => a.name.length)}`);
    const framedChart = await warned(pptxOf({
        'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${pns}><p:cSld><p:spTree>${'<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId2"/></a:graphicData></a:graphic></p:graphicFrame>'.repeat(2000)}</p:spTree></p:cSld></p:sld>`,
        'ppt/slides/_rels/slide1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
        'ppt/charts/chart1.xml': `<?xml version="1.0"?><c:chartSpace ${pns}><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numCache><c:pt idx="0"><c:v>${'9'.repeat(100_000)}</c:v></c:pt></c:numCache></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`,
    }), 'pptx', { extractAttachments: true });
    const framedText = await outputOf(framedChart.ast, 'text');
    const framedHtml = await outputOf(framedChart.ast, 'html');
    const chartJson = /<script type="application\/json" id="([^"]+)">([\s\S]*?)<\/script>/.exec(framedHtml);
    let chartJsonOk = false;
    try { chartJsonOk = !!chartJson && Array.isArray(JSON.parse(chartJson[2]).dataSets) && framedHtml.includes(`getElementById('${chartJson[1]}')`); } catch { chartJsonOk = false; }
    check('pptx: one 100 KB chart framed 2,000 times is written within bounds, its HTML data once', !framedChart.error && framedText.length < 20_000_000 && framedHtml.length < 20_000_000 && chartJsonOk && framedHtml.split('type="application/json"').length === 2, `${framedText.length} ${framedHtml.length} ${chartJsonOk} ${framedChart.error}`);
    const longTitle = 'w'.repeat(100_000);
    for (const [label, src] of [
        ['a \\nameref to a 100 KB heading, 2,000 times', `\\documentclass{article}\\begin{document}\\section{${longTitle}}\\label{k}\n${'\\nameref{k}\n\n'.repeat(2000)}\\end{document}`],
        ['a 100 KB theorem title on 2,000 theorems', `\\documentclass{article}\\newtheorem{t}{${longTitle}}\\begin{document}\n${'\\begin{t}x\\end{t}\n'.repeat(2000)}\\end{document}`],
        ['a 100 KB beamer title on 2,000 title frames', `\\documentclass{beamer}\\title{${longTitle}}\\begin{document}\n${'\\begin{frame}\\maketitle\\end{frame}\n'.repeat(2000)}\\end{document}`],
    ] as const) {
        const started = Date.now();
        const repeatedTex = await warned(Buffer.from(src), 'tex');
        const texText = await outputOf(repeatedTex.ast, 'text');
        check(`latex: ${label} repeats within maxRepeatedContent`, !repeatedTex.error && texText.length < 20_000_000 && Date.now() - started < 5000 && repeatedTex.codes.includes('REPEATED_CONTENT_LIMIT_EXCEEDED'), `${texText.length} ${Date.now() - started}ms ${repeatedTex.codes} ${repeatedTex.error}`);
    }

    // PDF: pdf.js's own work is bounded. It runs in a separate process under a memory limit (a stream
    // inflating to a string of tens of millions of characters ended the host out of memory), within a
    // time budget (a form drawing a 1 MB string off the page hundreds of times gave nothing to count),
    // and annotations a page tree shares are charged per page (58 KB ran the process out of memory).
    const pdfFrom = (objects: (string | Buffer)[]) => {
        const parts: Buffer[] = [Buffer.from('%PDF-1.7\n')];
        const offsets: number[] = [];
        let at = parts[0].length;
        objects.forEach((o, i) => { const b = Buffer.concat([Buffer.from(`${i + 1} 0 obj\n`), typeof o === 'string' ? Buffer.from(o, 'latin1') : o, Buffer.from('\nendobj\n')]); offsets.push(at); at += b.length; parts.push(b); });
        parts.push(Buffer.from(`xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map(o => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${at}\n%%EOF\n`));
        return Buffer.concat(parts);
    };
    const { deflateSync } = require('zlib') as typeof import('zlib');
    const flate = (content: Buffer, dict = '') => { const data = deflateSync(content, { level: 9 }); return Buffer.concat([Buffer.from(`<< /Length ${data.length} /Filter /FlateDecode ${dict} >>\nstream\n`), data, Buffer.from('\nendstream')]); };
    const pagesSharingAnnots = (pages: number, annots: number) => pdfFrom([
        '<< /Type /Catalog /Pages 2 0 R >>', `<< /Type /Pages /Kids [${Array.from({ length: pages }, (_, i) => `${5 + i} 0 R`).join(' ')}] /Count ${pages} >>`,
        '<< /Type /Annot /Subtype /Link /Rect [0 0 10 10] /A << /S /URI /URI (https://e.com/) >> >>', `[${'3 0 R '.repeat(annots)}]`,
        ...Array.from({ length: pages }, () => '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Annots 4 0 R >>'),
    ]);
    let started = Date.now();
    const sharedAnnots = await warned(pagesSharingAnnots(300, 3000), 'pdf');
    check('pdf: 300 pages sharing 3,000 annotations stop at maxAnnotations and maxTimeMs', !sharedAnnots.error && Date.now() - started < 15_000 && sharedAnnots.codes.includes('ANNOTATION_EXTRACTION_FAILED'), `${Date.now() - started}ms ${sharedAnnots.codes} ${sharedAnnots.error}`);
    const roomyAnnots = await warned(pagesSharingAnnots(10, 100), 'pdf', { pdfParserConfig: { maxAnnotations: 100_000 } });
    check('pdf: maxAnnotations is the base of the limit', !roomyAnnots.codes.includes('ANNOTATION_EXTRACTION_FAILED'), `${roomyAnnots.codes}`);
    const bigString = (length: number) => Buffer.concat([Buffer.from('BT /F1 10 Tf 0 50 Td ('), Buffer.alloc(length, 'a'), Buffer.from(') Tj ET\n')]);
    const drawnOften = pdfFrom([
        '<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> /XObject << /X 6 0 R >> >> /Contents 4 0 R >>',
        flate(Buffer.from('q /X Do Q\n'.repeat(300))), '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
        flate(bigString(1_000_000), '/Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >>'),
    ]);
    started = Date.now();
    const drawn = await warned(drawnOften, 'pdf');
    check('pdf: a form drawing a 1 MB string 300 times stops at maxTimeMs', !drawn.error && Date.now() - started < 15_000 && drawn.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${Date.now() - started}ms ${drawn.codes} ${drawn.error}`);
    started = Date.now();
    const quickly = await warned(drawnOften, 'pdf', { pdfParserConfig: { maxTimeMs: 200 } });
    check('pdf: maxTimeMs is the base of the limit', !quickly.error && Date.now() - started < 4000 && quickly.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${Date.now() - started}ms ${quickly.codes}`);
    started = Date.now();
    const abortedPdf = await parseQuiet(drawnOften, 'pdf', { abortSignal: AbortSignal.timeout(500), pdfParserConfig: { maxTimeMs: 60_000 } });
    check('pdf: abortSignal stops pdf.js in the middle of a stream', /abort/i.test(abortedPdf.error) && Date.now() - started < 4000, `${Date.now() - started}ms ${abortedPdf.error}`);
    const inflating = pdfFrom([
        '<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
        flate(bigString(30_000_000)), '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    ]);
    let processError: any;
    try { await OfficeParser.parseOffice(inflating, { fileType: 'pdf', onWarning: () => {}, pdfParserConfig: { processMemoryMb: 128 } } as any); } catch (e) { processError = e; }
    check('pdf: a stream pdf.js cannot hold in its memory limit fails the parse, not the process', processError?.officeIssue?.code === 'PDF_PROCESS_FAILED', String(processError));
    const afterwards = await parseQuiet(fs.readFileSync(path.join(files, 'test.pdf')), 'pdf');
    check('pdf: the next parse reads normally', !afterwards.error && (afterwards.ast?.content.length ?? 0) > 0, afterwards.error);

    // Destinations are resolved within the time budget: bookmarks each naming a page deep in a long
    // page tree walked the tree per bookmark outside it (500 KB, two minutes).
    const pdfObjectStream = (objects: string[]) => {
        let header = '', body = '', offset = 0;
        objects.forEach((o, i) => { header += `${i + 1} ${offset} `; body += o + '\n'; offset += Buffer.byteLength(o + '\n', 'latin1'); });
        const n = objects.length, streamNum = n + 1, xrefNum = n + 2;
        const packed = deflateSync(Buffer.from(header + body, 'latin1'), { level: 9 });
        const start = Buffer.from('%PDF-1.7\n', 'latin1');
        const streamObj = Buffer.concat([Buffer.from(`${streamNum} 0 obj\n<</Type/ObjStm/N ${n}/First ${Buffer.byteLength(header, 'latin1')}/Length ${packed.length}/Filter/FlateDecode>>\nstream\n`, 'latin1'), packed, Buffer.from('\nendstream\nendobj\n')]);
        const xrefAt = start.length + streamObj.length;
        const rows = Buffer.alloc((n + 3) * 7);
        const put = (i: number, t: number, a: number, b: number) => { rows[i * 7] = t; rows.writeUInt32BE(a, i * 7 + 1); rows.writeUInt16BE(b, i * 7 + 5); };
        put(0, 0, 0, 65535);
        for (let i = 1; i <= n; i++) put(i, 2, streamNum, i - 1);
        put(streamNum, 1, start.length, 0); put(xrefNum, 1, xrefAt, 0);
        const xrefData = deflateSync(rows, { level: 9 });
        return Buffer.concat([start, streamObj, Buffer.from(`${xrefNum} 0 obj\n<</Type/XRef/Size ${n + 3}/W[1 4 2]/Root 1 0 R/Length ${xrefData.length}/Filter/FlateDecode>>\nstream\n`, 'latin1'), xrefData, Buffer.from(`\nendstream\nendobj\nstartxref\n${xrefAt}\n%%EOF\n`, 'latin1')]);
    };
    const depth = 10000, bookmarks = 3000;
    const chainAt = (k: number) => 6 + k, targetAt = (i: number) => 6 + depth + i, itemAt = (i: number) => 6 + depth + bookmarks + i;
    const deepOutline = pdfObjectStream([
        '<</Type/Catalog/Pages 2 0 R/Outlines 5 0 R>>', `<</Type/Pages/Kids[3 0 R ${chainAt(0)} 0 R]/Count 1>>`, '<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]>>', '<</Dummy 0>>',
        `<</Type/Outlines/First ${itemAt(0)} 0 R/Last ${itemAt(bookmarks - 1)} 0 R/Count ${bookmarks}>>`,
        ...Array.from({ length: depth }, (_, k) => `<</Type/Pages/Parent ${k === 0 ? 2 : chainAt(k - 1)} 0 R/Kids[${k === depth - 1 ? Array.from({ length: bookmarks }, (_, i) => `${targetAt(i)} 0 R`).join(' ') : `${chainAt(k + 1)} 0 R`}]/Count 0>>`),
        ...Array.from({ length: bookmarks }, () => `<</Type/Page/Parent ${chainAt(depth - 1)} 0 R>>`),
        ...Array.from({ length: bookmarks }, (_, i) => `<</Title(b${i})/Parent 5 0 R${i > 0 ? `/Prev ${itemAt(i - 1)} 0 R` : ''}${i < bookmarks - 1 ? `/Next ${itemAt(i + 1)} 0 R` : ''}/Dest[${targetAt(i)} 0 R/XYZ 0 0 0]>>`),
    ]);
    started = Date.now();
    const outlined = await warned(deepOutline, 'pdf', { pdfParserConfig: { maxTimeMs: 200 } });
    check('pdf: bookmarks naming pages deep in a long page tree are resolved within maxTimeMs', !outlined.error && Date.now() - started < 12_000 && outlined.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${Date.now() - started}ms ${deepOutline.length} bytes ${outlined.codes} ${outlined.error}`);
    await pdfProcessTests({ pdfFrom, flate, pdfForms, drawnOften, parseQuiet, files });

    // Tables write within bounds however ragged or spanned: rows padded to the widest (a 4 KB DOCX,
    // 100 million positions), a staircase of rows each spanning a thousand below (a 1.5 KB EPUB, 500 MB
    // of DOCX), a text layout padding every cell to the widest (1.6 KB, 200 MB), and a chunk header
    // repeated before each row (213 MB); and an EPUB names its pictures in linear time.
    const tableCell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const ragged = await parseQuiet(docxOf(`<w:tbl><w:tr>${tableCell('h').repeat(10000)}</w:tr>${`<w:tr>${tableCell('x')}</w:tr>`.repeat(10000)}</w:tbl>`), 'docx');
    const wideCell = await parseQuiet(docxOf(`<w:tbl><w:tr>${tableCell('A'.repeat(100_000))}</w:tr>${`<w:tr>${tableCell('x')}</w:tr>`.repeat(2000)}</w:tbl>`), 'docx');
    const staircase = await parseQuiet(Buffer.from(`<html><body><table>${'<tr><td rowspan="1000">a</td></tr>'.repeat(9000)}</table></body></html>`), 'html');
    for (const [label, result] of [['a row of 10,000 cells over 10,000 rows of one', ragged], ['one 100 KB cell over 2,000 rows', wideCell], ['9,000 rows each spanning 1,000', staircase]] as const) {
        let longest = 0, slowest = 0;
        for (const format of ['text', 'md', 'csv', 'chunks', 'tex', 'docx', 'odt', 'html', 'rtf'] as const) {
            const t0 = Date.now();
            const out = result.ast ? (await OfficeGenerator.generate(result.ast, format, { onWarning: () => {} } as any)).value as any : '';
            slowest = Math.max(slowest, Date.now() - t0);
            longest = Math.max(longest, typeof out === 'string' ? out.length : Array.isArray(out) ? JSON.stringify(out).length : out?.byteLength ?? 0);
        }
        check(`tables: ${label} is written within bounds in every format`, !result.error && longest < 25_000_000 && slowest < 5000, `${longest} ${slowest}ms ${result.error}`);
    }
    const texWide = await warned(docxOf(`<w:tbl><w:tr>${tableCell('h').repeat(1200)}</w:tr></w:tbl>`), 'docx');
    const texWarnings: string[] = [];
    if (texWide.ast) await OfficeGenerator.generate(texWide.ast, 'tex', { onWarning: (w: any) => texWarnings.push(w.code) } as any);
    check('tex: cells past its widest table are reported, not dropped silently', texWarnings.includes('CONTENT_NOT_REPRESENTABLE'), `${texWarnings}`);
    const pictureBook = Buffer.from(zipSync({
        'mimetype': enc('application/epub+zip'),
        'META-INF/container.xml': enc('<?xml version="1.0"?><container><rootfiles><rootfile full-path="o.opf"/></rootfiles></container>'),
        'o.opf': enc(`<package><metadata/><manifest>${Array.from({ length: 2000 }, (_, i) => `<item id="i${i}" href="c${i}.xhtml" media-type="application/xhtml+xml"/>`).join('')}</manifest><spine>${Array.from({ length: 2000 }, (_, i) => `<itemref idref="i${i}"/>`).join('')}</spine></package>`),
        ...Object.fromEntries(Array.from({ length: 2000 }, (_, i) => [`c${i}.xhtml`, enc(`<html><body>${'<img src="data:image/png;base64,AA==">'.repeat(50)}</body></html>`)])),
    }));
    await timed('epub: 2,000 chapters of 50 pictures each are named', () => parseQuiet(pictureBook, 'epub', { extractAttachments: true }));

    // Headers, footers and the like go through every pass the content does (the grid budget above
    // all), and a sparse sheet's rows are filled back in within it: an empty row naming row 4,094, or a
    // sparse sheet in a footer, made 270 MB of DOCX from a few hundred bytes of AST.
    const unzippedSize = (value: any) => Object.values(unzipSync(new Uint8Array(value))).reduce((n, entry) => n + entry.length, 0);
    const sparseSheet = (rowsOf: any[]) => ({ type: 'sheet', metadata: { sheetName: 'S' }, children: rowsOf });
    const farCell = (row: number, col: number) => ({ type: 'cell', metadata: { row, col }, children: [{ type: 'text', text: 'a' }] });
    const namedRow = astWith([sparseSheet([{ type: 'row', children: [farCell(0, 999)] }, { type: 'row', metadata: { row: 4094 }, children: [] }])]);
    const footerSheet = { ...astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }]), auxiliary: { footers: [sparseSheet([{ type: 'row', children: [farCell(0, 999)] }, { type: 'row', children: [farCell(4094, 0)] }])] } } as any;
    const staircaseTable = () => ({ type: 'table', children: Array.from({ length: 2000 }, () => ({ type: 'row', children: [{ type: 'cell', metadata: { rowSpan: 1000 }, children: [{ type: 'text', text: 'z' }] }] })) });
    const staircaseHeaders = { ...astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }]), auxiliary: { headers: [staircaseTable(), staircaseTable()], footers: [staircaseTable(), staircaseTable()] } } as any;
    for (const [label, ast] of [['an empty sheet row naming row 4,094', namedRow], ['a sparse sheet in a footer', footerSheet], ['tables of 2,000 rows each spanning 1,000 in headers and footers', staircaseHeaders]] as const) {
        const sizes: number[] = [];
        for (const format of ['docx', 'odt'] as const) sizes.push(unzippedSize((await OfficeGenerator.generate(ast, format, { onWarning: () => {} } as any)).value));
        check(`docx, odt: ${label} is written within the grid budget`, sizes.every(n => n < 20_000_000), sizes.join(' '));
    }
    const commentedHeader = { ...astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }]), auxiliary: { headers: [{ type: 'paragraph', children: [{ type: 'text', text: 'visible' }, { type: 'comment', text: 'HIDDEN-NOTE', metadata: { sourceSyntax: 'html' } }] }] } } as any;
    const headerDocx = unzipSync(new Uint8Array((await OfficeGenerator.generate(commentedHeader, 'docx', { onWarning: () => {} } as any)).value as any));
    const headerXml = Object.entries(headerDocx).filter(([name]) => /header\d*\.xml$/.test(name)).map(([, data]) => strFromU8(data)).join('');
    check('docx: a source comment in a header is left out, as in the body', headerXml.includes('visible') && !headerXml.includes('HIDDEN-NOTE'), headerXml.slice(0, 300));

    // A chart's table is built within the grid budget (a 9.6 KB PPTX, 3,000 labels by 3,000 series, ran
    // DOCX, ODT and LaTeX out of memory); names made unique resume their numbering (60,000 bookmarks named
    // `a` took two minutes); an AST sharing nodes along many paths is refused; lists share numbering
    // definitions; and a delimiter an AST's own config names is taken only when short.
    const bigChart = { ...astWith([{ type: 'chart', metadata: { attachmentName: 'c' } }]), attachments: [{ type: 'chart', name: 'c', mimeType: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', data: '', chartData: { title: 't', labels: Array.from({ length: 3000 }, () => 'a'), dataSets: Array.from({ length: 3000 }, () => ({ name: 's', values: [] })) } }] } as any;
    for (const format of ['docx', 'odt', 'tex'] as const) {
        const chartWarnings: string[] = [];
        const t0 = Date.now();
        let failure = '';
        try { await OfficeGenerator.generate(bigChart, format, { onWarning: (w: any) => chartWarnings.push(w.code) } as any); } catch (e) { failure = String(e); }
        check(`${format}: a chart of 3,000 labels by 3,000 series is written within the grid budget`, !failure && Date.now() - t0 < 5000 && chartWarnings.includes('CONTENT_NOT_REPRESENTABLE'), `${Date.now() - t0}ms ${failure} ${chartWarnings}`);
    }
    const manyBookmarks = await parseQuiet(docxOf(`<w:p>${'<w:bookmarkStart w:id="0" w:name="a"/>'.repeat(60000)}<w:r><w:t>x</w:t></w:r></w:p>`), 'docx');
    await timed('docx, odt: a paragraph naming one bookmark 60,000 times is written', async () => { for (const format of ['docx', 'odt'] as const) await OfficeGenerator.generate(manyBookmarks.ast!, format, { onWarning: () => {} } as any); });
    const sameHeadings = await parseQuiet(Buffer.from('# a\n\n'.repeat(20000)), 'md');
    await timed('docx, odt: 20,000 headings of one text are written', async () => { for (const format of ['docx', 'odt'] as const) await OfficeGenerator.generate(sameHeadings.ast!, format, { onWarning: () => {} } as any); });
    let doubled: any = { type: 'text', text: 'ab' };
    for (let i = 0; i < 40; i++) doubled = { type: 'text', children: [doubled, doubled] };
    let sharingError: any;
    const t1 = Date.now();
    try { await OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [doubled] }]), 'html', { onWarning: () => {} } as any); } catch (e) { sharingError = e; }
    check('generators: an AST whose nodes each hold the next twice, 40 deep, is refused at once', sharingError?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - t1 < 2000, `${Date.now() - t1}ms ${sharingError}`);
    const manyLists = await parseQuiet(Buffer.from('1. a\n\nx\n\n'.repeat(20000)), 'md');
    const listDocx = unzipSync(new Uint8Array((await OfficeGenerator.generate(manyLists.ast!, 'docx', { onWarning: () => {} } as any)).value as any));
    const numberingSize = listDocx['word/numbering.xml']?.length ?? 0;
    check('docx: 20,000 lists share their numbering definition', numberingSize > 0 && numberingSize < 5_000_000, `${numberingSize}`);
    const twoLists = await parseQuiet(Buffer.from('1. one\n2. two\n\ntext\n\n1. again\n'), 'md');
    const listsReread = await parseQuiet(Buffer.from((await OfficeGenerator.generate(twoLists.ast!, 'docx', { onWarning: () => {} } as any)).value as any), 'docx');
    const markdownAgain = (await OfficeGenerator.generate(listsReread.ast!, 'md', { onWarning: () => {} } as any)).value as string;
    check('docx: a list sharing a numbering definition still starts at its own number', /1\. again/.test(markdownAgain), markdownAgain);
    const longDelimiter = { ...astWith(Array.from({ length: 1000 }, () => ({ type: 'paragraph', children: [{ type: 'text', text: 'x' }] }))), config: { newlineDelimiter: 'n'.repeat(100_000) } } as any;
    const delimited = (await OfficeGenerator.generate(longDelimiter, 'text', { onWarning: () => {} } as any)).value as string;
    check('text: a long newline delimiter in an AST\'s own config is not taken', delimited.length < 1_000_000, `${delimited.length}`);

    // A list level far outside Word's nine (negative, or past a billion) is a level Word has.
    const levelNumbering = `<?xml version="1.0"?><w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    for (const level of ['-2000000000', '2000000000', 'x']) {
        await timed(`docx: a list level of ${level} is read and written`, async () => {
            const levels = await parseQuiet(docxOf(`<w:p><w:pPr><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>`.repeat(20), { 'word/numbering.xml': levelNumbering }), 'docx');
            for (const format of ['md', 'html', 'docx', 'odt', 'tex'] as const) await OfficeGenerator.generate(levels.ast!, format, { onWarning: () => {} } as any);
        });
    }
    // A sheet naming one drawing many times places it once.
    const drawingRels = Array.from({ length: 300 }, (_, i) => `<Relationship Id="i${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>`).join('');
    const drawingSheet = xlsxOf({
        'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>v</t></is></c></row></sheetData>${'<drawing r:id="rId1"/>'.repeat(2000)}</worksheet>`,
        'xl/worksheets/_rels/sheet1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>',
        'xl/drawings/drawing1.xml': '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
        'xl/drawings/_rels/drawing1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRels}</Relationships>`,
        'xl/media/image1.png': 'png',
    });
    const drawingParse = await parseQuiet(drawingSheet, 'xlsx', { extractAttachments: true });
    const drawingImages = JSON.stringify(drawingParse.ast?.content ?? []).split('"type":"image"').length - 1;
    check('xlsx: a sheet naming one drawing 2,000 times places its pictures once', !drawingParse.error && drawingImages <= 300, `${drawingParse.error} ${drawingImages}`);
    // Many sheets and references to them.
    await timed('html: 40,000 sheets are written', () => OfficeGenerator.generate(astWith(Array.from({ length: 40000 }, () => ({ type: 'sheet', children: [], metadata: { sheetName: 's' } }))), 'html', { onWarning: () => {} } as any));
    let chartRows = '<table:table-row><table:table-cell/><table:table-cell><text:p>s</text:p></table:table-cell></table:table-row>';
    for (let i = 0; i < 40000; i++) chartRows += `<table:table-row><table:table-cell><text:p>[Q${i}.A1]</text:p></table:table-cell></table:table-row>`;
    await timed('ods: a chart of 40,000 references beside 40,000 sheets parses', () => parseQuiet(odfOf('spreadsheet', '<table:table table:name="a"/>'.repeat(40000), { 'Object 1/content.xml': chartDoc(`<table:table>${chartRows}</table:table>`) }), 'ods', { extractAttachments: true }));
    let objectFrames = '';
    for (let i = 0; i < 80000; i++) objectFrames += `<draw:frame><draw:object xlink:href="./O${i}"/></draw:frame>`;
    await timed('odt: 80,000 embedded objects beside 8,000 parts parse', () => parseQuiet(odfOf('text', `<text:p>${objectFrames}</text:p>`, Object.fromEntries(Array.from({ length: 8000 }, (_, i) => [`z${i}/xcontent.xml`, '']))), 'odt'));
    // Hyperlinks nested in hyperlinks: the outermost gives the runs their link, once.
    const linkStarted = Date.now();
    const nestedLinks = await parseQuiet(docxOf(`<w:p>${'<w:hyperlink w:anchor="a">'.repeat(400)}${'<w:r><w:t>x</w:t></w:r>'.repeat(2000)}${'</w:hyperlink>'.repeat(400)}</w:p>`), 'docx');
    check('docx: 2,000 runs in hyperlinks nested 400 deep parse in linear time', Date.now() - linkStarted < 5000, `${Date.now() - linkStarted}ms`);
    let deepLinksError: any;
    try { await OfficeParser.parseOffice(docxOf(`<w:p>${'<w:hyperlink w:anchor="a">'.repeat(5000)}<w:r><w:t>x</w:t></w:r>${'</w:hyperlink>'.repeat(5000)}</w:p>`), { fileType: 'docx', onWarning: () => {} } as any); } catch (e) { deepLinksError = e; }
    check('docx: hyperlinks nested past what the stack holds are a nesting error', deepLinksError?.officeIssue?.code === 'MAX_NESTING_DEPTH_EXCEEDED', String(deepLinksError));
    const linkedRuns = JSON.stringify(nestedLinks.ast?.content ?? []).split('"link":"#a"').length - 1;
    check('docx: runs in nested hyperlinks keep the outermost link', !nestedLinks.error && linkedRuns === 2000, `${nestedLinks.error} ${linkedRuns}`);

    // Long lists are appended item by item: a spread of more than about 120,000 arguments throws.
    const longParagraph = astWith([{ type: 'paragraph', children: Array.from({ length: 200000 }, () => ({ type: 'text', text: 'x' })) }]);
    let longError: any;
    let longMarkdown = '';
    try { longMarkdown = (await OfficeGenerator.generate(longParagraph, 'md', { onWarning: () => {} } as any)).value as string; } catch (e) { longError = e; }
    check('md: a paragraph of 200,000 runs is written', !longError && longMarkdown.length >= 200000, `${longError?.officeIssue?.code ?? longError}`);
    const longTextBox = await parseQuiet(docxOf(`<w:p><w:r><w:drawing><w:txbxContent><w:p>${'<w:r><w:t>x</w:t></w:r>'.repeat(130000)}</w:p></w:txbxContent></w:drawing></w:r></w:p>`), 'docx');
    check('docx: a text box of 130,000 runs is read, after the paragraph drawing it', !longTextBox.error && (longTextBox.ast?.content[1] as any)?.text?.length === 130000, longTextBox.error);

    // Content Word wraps (content controls, custom XML, text boxes in runs) is read, not dropped.
    const wrapped = await parseQuiet(docxOf(
        '<w:sdt><w:sdtPr><w:alias w:val="toc"/></w:sdtPr><w:sdtContent><w:sdt><w:sdtContent><w:p><w:r><w:t>BODYCONTROL</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>'
        + '<w:customXml w:element="x"><w:p><w:r><w:t>CUSTOMXML</w:t></w:r></w:p></w:customXml>'
        + '<w:tbl><w:sdt><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>ROWCONTROL</w:t></w:r></w:p></w:tc><w:customXml w:element="c"><w:tc><w:p><w:r><w:t>CELLCUSTOM</w:t></w:r></w:p></w:tc></w:customXml></w:tr></w:sdtContent></w:sdt></w:tbl>'
        + '<w:p><w:r><w:t>before </w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><w:txbxContent><w:p><w:r><w:t>RUNTEXTBOX</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>BOXTABLE</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:txbxContent></w:drawing></mc:Choice><mc:Fallback><w:pict><w:txbxContent><w:p><w:r><w:t>RUNTEXTBOX</w:t></w:r></w:p></w:txbxContent></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>',
        { 'word/header1.xml': `<?xml version="1.0"?><w:hdr ${W}><w:sdt><w:sdtContent><w:sdt><w:sdtContent><w:p><w:r><w:t>HEADERCONTROL</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt></w:hdr>` },
    ), 'docx');
    const wrappedJson = JSON.stringify(wrapped.ast ?? {});
    const missing = ['BODYCONTROL', 'CUSTOMXML', 'ROWCONTROL', 'CELLCUSTOM', 'RUNTEXTBOX', 'BOXTABLE', 'HEADERCONTROL'].filter(marker => !wrappedJson.includes(marker));
    check('docx: content in content controls, custom XML and text boxes in runs is read', !wrapped.error && missing.length === 0, `${wrapped.error} missing ${missing}`);
    const textBoxOnce = (wrapped.ast?.content ?? []).filter((n: any) => n.text === 'RUNTEXTBOX').length;
    check('docx: a text box in alternate content is read once', textBoxOnce === 1, `${textBoxOnce}`);
    const boxedPicture = await parseQuiet(docxOf(
        '<w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rId9"/></w:drawing></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p>',
        { 'word/_rels/document.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/boxed.png"/></Relationships>', 'word/media/boxed.png': 'png' },
    ), 'docx', { extractAttachments: true });
    const boxedCount = JSON.stringify(boxedPicture.ast?.content ?? []).split('"attachmentName":"boxed.png"').length - 1;
    check('docx: a picture in a text box is shown once', !boxedPicture.error && boxedCount === 1, `${boxedPicture.error} ${boxedCount}`);
    // ODF: fields, list headers, grouped rows and cell text around spans are read; deleted text is not.
    const odtWrapped = await parseQuiet(odfOf('text',
        '<text:list><text:list-header><text:p>LISTHEADER</text:p></text:list-header><text:list-item><text:p>item</text:p></text:list-item></text:list>'
        + '<table:table><table:table-row-group><table:table-row><table:table-cell><text:p>ROWGROUP</text:p></table:table-cell></table:table-row></table:table-row-group><table:table-rows><table:table-row><table:table-cell><text:p>TABLEROWS</text:p></table:table-cell></table:table-row></table:table-rows></table:table>'
        + '<text:p>Figure <text:sequence text:name="Figure">SEQUENCE</text:sequence>, <text:date>DATEFIELD</text:date>, <text:bookmark-ref text:ref-name="b">CROSSREF</text:bookmark-ref>, <text:meta>METATEXT</text:meta>, <text:ruby><text:ruby-base>RUBYBASE</text:ruby-base><text:ruby-text>RUBYTEXT</text:ruby-text></text:ruby></text:p>'
        + '<text:tracked-changes><text:changed-region text:id="c1"><text:deletion><text:p>DELETEDTEXT</text:p></text:deletion></text:changed-region></text:tracked-changes>'), 'odt');
    const odtJson = JSON.stringify(odtWrapped.ast?.content ?? []);
    const odtMissing = ['LISTHEADER', 'ROWGROUP', 'TABLEROWS', 'SEQUENCE', 'DATEFIELD', 'CROSSREF', 'METATEXT', 'RUBYBASE', 'RUBYTEXT'].filter(marker => !odtJson.includes(marker));
    check('odt: fields, list headers and grouped rows are read, deleted text is not', !odtWrapped.error && odtMissing.length === 0 && !odtJson.includes('DELETEDTEXT'), `${odtWrapped.error} missing ${odtMissing}`);
    const odsSpans = await parseQuiet(odfOf('spreadsheet', '<table:table table:name="S"><table:table-row><table:table-cell><text:p>Total: <text:span>FIVE</text:span> units</text:p></table:table-cell></table:table-row><table:table-row-group><table:table-row><table:table-cell><text:p>GROUPED</text:p></table:table-cell></table:table-row></table:table-row-group></table:table>'), 'ods');
    const odsCells = cellsOf(odsSpans.ast).map((c: any) => c.text);
    check('ods: a cell keeps the text around its spans', !odsSpans.error && odsCells.includes('Total: FIVE units') && odsCells.includes('GROUPED'), `${odsSpans.error} ${JSON.stringify(odsCells)}`);
    const alternateShape = await parseQuiet(pptxOf({ 'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${pns} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:cSld><p:spTree><mc:AlternateContent><mc:Choice Requires="a14"><p:sp><p:txBody><a:p><a:r><a:t>ALTSHAPE</a:t></a:r></a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:txBody><a:p><a:r><a:t>ALTSHAPE</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>` }), 'pptx');
    const alternateCount = (JSON.stringify(alternateShape.ast?.content ?? []).match(/"text":"ALTSHAPE"/g) ?? []).length;
    check('pptx: a shape in alternate content is read once', !alternateShape.error && alternateCount === 2, `${alternateShape.error} ${alternateCount}`);

    // An element's own source is found in one pass, however many elements of its name nest in it.
    await timed('docx: raw source of a run of 160,000 nested w:t is found', () => parseQuiet(docxOf(`<w:p><w:r><w:t>${'<w:t>'.repeat(160000)}x${'</w:t>'.repeat(160000)}</w:t></w:r></w:p>`), 'docx', { includeRawContent: true, serializeRawContent: false }));
    const rawLimited = await parseQuiet(docxOf(`<w:p><w:r><w:t>${'x'.repeat(5000)}</w:t></w:r></w:p>`), 'docx', { includeRawContent: true, serializeRawContent: false, decompressionLimits: { maxRawContentLength: 1000 } });
    let rawLimitedTotal = 0;
    const sumLimited = (ns: any[] | undefined) => ns?.forEach((n: any) => { rawLimitedTotal += n.rawContent?.length ?? 0; sumLimited(n.children); });
    sumLimited(rawLimited.ast?.content);
    check('docx: raw source longer than the budget is left out', !rawLimited.error && rawLimitedTotal <= 1000, `${rawLimited.error} ${rawLimitedTotal}`);
    // An XLSX comments part named many times is read and attached once.
    const commentRel = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>';
    const manyCommentRels = await heapBudget('xlsx: a comments part of 1 MB named 2,000 times is read', () => parseQuiet(xlsxOf({
        'xl/worksheets/_rels/sheet1.xml.rels': `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${commentRel.repeat(2000)}</Relationships>`,
        'xl/comments1.xml': `<?xml version="1.0"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>a</author></authors><commentList><comment ref="A1" authorId="0"><text><t>${'x'.repeat(1_000_000)}</t></text></comment></commentList></comments>`,
    }), 'xlsx')) as any;
    const attachedComments = cellsOf(manyCommentRels.ast).reduce((sum: number, c: any) => sum + (c.comments?.length ?? 0), 0);
    check('xlsx: a comments part named 2,000 times is attached once', !manyCommentRels.error && attachedComments === 1, `${manyCommentRels.error} ${attachedComments}`);
    // Pictures' alt text is found through a map, not a scan of every drawing relationship per picture.
    let mediaRels = '<?xml version="1.0"?><Relationships>';
    for (let i = 0; i < 100000; i++) mediaRels += `<Relationship Id="r${i}" Target="../media/z"/>`;
    await timed('xlsx: 1,000 pictures beside 100,000 drawing relationships are read', () => parseQuiet(xlsxOf({
        'xl/drawings/_rels/drawing1.xml.rels': mediaRels + '</Relationships>',
        ...Object.fromEntries(Array.from({ length: 1000 }, (_, i) => [`xl/media/${i}`, ''])),
    }), 'xlsx', { extractAttachments: true }));
    // Text boxes nested in text boxes are read once each, after the paragraph drawing them.
    const boxRuns = '<w:t>x</w:t>'.repeat(400);
    await timed('docx: text boxes nested 1,000 deep, each of 400 runs, parse', () => parseQuiet(docxOf(`${`<w:p><w:r>${boxRuns}</w:r><w:r><w:pict><w:txbxContent>`.repeat(1000)}${'</w:txbxContent></w:pict></w:r></w:p>'.repeat(1000)}`), 'docx'));
    const boxedNote = await parseQuiet(docxOf('<w:p><w:r><w:t>outer</w:t></w:r><w:r><w:pict><w:txbxContent><w:p><w:bookmarkStart w:id="1" w:name="inbox"/><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></w:pict></w:r></w:p>'), 'docx');
    check('docx: a bookmark in a text box stays on the text box\'s paragraph', !boxedNote.error && JSON.stringify(boxedNote.ast?.content[1]?.metadata ?? {}).includes('inbox') && !JSON.stringify(boxedNote.ast?.content[0]?.metadata ?? {}).includes('inbox'), JSON.stringify(boxedNote.ast?.content.map(n => n.metadata)));
    // XLSX: an inline string keeps every rich run (not its phonetic reading), and a boolean shows as TRUE or FALSE.
    const inlineRuns = await parseQuiet(xlsxOf({ 'xl/worksheets/sheet1.xml': '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>ONE </t></r><r><t>TWO</t></r><rPh sb="0" eb="1"><t>READING</t></rPh></is></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="b"><v>0</v></c></row></sheetData></worksheet>' }), 'xlsx');
    const inlineCells = cellsOf(inlineRuns.ast).map((c: any) => c.text);
    check('xlsx: an inline string keeps every run, and booleans read TRUE and FALSE', !inlineRuns.error && JSON.stringify(inlineCells) === JSON.stringify(['ONE TWO', 'TRUE', 'FALSE']), `${inlineRuns.error} ${JSON.stringify(inlineCells)}`);
    // DOCX alternative-format chunks: read once each, within the document's budgets.
    const chunkRel = (id: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" Target="${target}"/>`;
    const chunkRels = (rels: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
    const mhtOf = (parts: string[]) => `MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary="B"\r\n\r\n${parts.map(p => `--B\r\n${p}\r\n`).join('')}--B--\r\n`;
    const altChunked = await parseQuiet(docxOf('<w:altChunk r:id="c1"/><w:altChunk r:id="c2"/><w:altChunk r:id="c3"/><w:tbl><w:tr><w:tc><w:altChunk r:id="c4"/></w:tc></w:tr></w:tbl>', {
        'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', 'afchunk.mht') + chunkRel('c2', 'chunk.htm') + chunkRel('c3', 'chunk.rtf') + chunkRel('c4', 'cell.txt')),
        'word/afchunk.mht': mhtOf(['Content-Type: text/html; charset="utf-8"\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<p>MHT=20CHUNK caf=C3=A9 <a href=3D"javascript:alert(1)">x</a></p>']),
        'word/chunk.htm': '<p>HTML CHUNK<script>alert(2)</script></p>',
        'word/chunk.rtf': '{\\rtf1\\ansi RTF CHUNK\\par}',
        'word/cell.txt': 'CELL CHUNK',
    }), 'docx');
    const chunkJson = JSON.stringify(altChunked.ast?.content ?? []);
    const chunkMissing = ['MHT CHUNK café', 'HTML CHUNK', 'RTF CHUNK', 'CELL CHUNK'].filter(marker => !chunkJson.includes(marker));
    check('docx: alternative-format chunks (MHT, HTML, RTF, text) are read where they stand', !altChunked.error && chunkMissing.length === 0, `${altChunked.error} missing ${chunkMissing}`);
    const chunkHtml = altChunked.ast ? (await OfficeGenerator.generate(altChunked.ast, 'html', { onWarning: () => {} } as any)).value as string : '';
    const chunkBody = chunkHtml.match(/<body[\s\S]*<\/body>/)?.[0] ?? '';
    check('docx: a chunk\'s scripts and script links do not reach HTML output', !!chunkBody && !/javascript:|<script|alert\(/i.test(chunkBody), chunkBody.slice(0, 300));
    const sameChunk = await parseQuiet(docxOf('<w:altChunk r:id="c1"/>'.repeat(100000), { 'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', 'chunk.htm')), 'word/chunk.htm': `<p>${'x'.repeat(100000)}</p>` }), 'docx');
    await timed('docx: one chunk part named 100,000 times is read', async () => sameChunk);
    check('docx: one chunk part named 100,000 times is read once', !sameChunk.error && JSON.stringify(sameChunk.ast?.content ?? []).length < 300000, sameChunk.error);
    const missingChunks = await warned(docxOf('<w:altChunk r:id="none"/>'.repeat(50000)), 'docx');
    check('docx: 50,000 chunks naming nothing warn once', !missingChunks.error && missingChunks.codes.filter((c: string) => c === 'ALT_CHUNK_NOT_READ').length === 1, `${missingChunks.error} ${missingChunks.codes.length}`);
    await timed('docx: an MHT chunk of 100,000 parts and 1 MB of = is read', () => parseQuiet(docxOf('<w:altChunk r:id="c1"/>', {
        'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', 'afchunk.mht')),
        'word/afchunk.mht': mhtOf(['Content-Type: text/html\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<p>' + '='.repeat(1_000_000) + '</p>', ...Array.from({ length: 100000 }, (_, i) => `Content-Type: image/png\r\nContent-Location: ${i}.png\r\nContent-Transfer-Encoding: base64\r\n\r\nAAAA`)]),
    }), 'docx', { extractAttachments: true }));
    // A DOCX chunk inflates within what the document around it left of maxUncompressedBytes.
    const innerDocx = Buffer.from(zipSync({ '[Content_Types].xml': enc('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'), 'word/document.xml': enc(`<?xml version="1.0"?><w:document ${W}><w:body><w:p><w:r><w:t>${'x'.repeat(3_000_000)}</w:t></w:r></w:p></w:body></w:document>`) }));
    const nestedBomb = await parseQuiet(docxOf('<w:altChunk r:id="c1"/>', { 'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', 'inner.docx')), 'word/inner.docx': new Uint8Array(innerDocx) }), 'docx', { decompressionLimits: { maxUncompressedBytes: 2_000_000 } });
    check('docx: a DOCX chunk inflating past what the document left is refused', /size|limit/i.test(nestedBomb.error), nestedBomb.error || 'parsed');
    // A chunk that cannot be read is skipped with a warning, never as a way past a budget: an unreadable
    // DOCX chunk's reported error is not reported, and a DOCX chunk past the size limit still fails.
    const unreadable = await warned(docxOf('<w:altChunk r:id="c1"/><w:p><w:r><w:t>after</w:t></w:r></w:p>', { 'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', 'inner.docx')), 'word/inner.docx': 'not a zip' }), 'docx');
    check('docx: an unreadable chunk is one warning, and the document is read', !unreadable.error && JSON.stringify(unreadable.codes) === '["ALT_CHUNK_NOT_READ"]' && unreadable.ast?.content[0]?.text === 'after', `${unreadable.error} ${unreadable.codes}`);
    // Only the parts an aFChunk relationship names are inflated as chunks, not every part of a chunk's kind.
    const ignoredEmbedding = await parseQuiet(docxOf('<w:p><w:r><w:t>main</w:t></w:r></w:p>', { 'word/embeddings/object.docx': new Uint8Array(innerDocx) }), 'docx', { decompressionLimits: { maxUncompressedBytes: 2_000_000 } });
    check('docx: an embedded DOCX no chunk names is not inflated', !ignoredEmbedding.error, ignoredEmbedding.error);
    // Fields, bookmarks and note references a paragraph holds by the hundred thousand are read in one pass.
    const noteParts = { 'word/footnotes.xml': `<?xml version="1.0"?><w:footnotes ${W}><w:footnote w:id="1"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:footnote></w:footnotes>` };
    await timed('docx: 100,000 fields begun and never ended before 100,000 runs are read', () => parseQuiet(docxOf(`<w:p>${'<w:r><w:fldChar w:fldCharType="begin"/><w:instrText> NOTEREF a \\f </w:instrText></w:r>'.repeat(100000)}${'<w:r><w:t>x</w:t></w:r>'.repeat(100000)}</w:p>`, noteParts), 'docx'));
    await timed('docx: 100,000 bookmarks left open before 50,000 note references are read', () => parseQuiet(docxOf(`<w:p>${Array.from({ length: 100000 }, (_, i) => `<w:bookmarkStart w:id="${i}" w:name="b${i}"/>`).join('')}${'<w:r><w:footnoteReference w:id="1"/></w:r>'.repeat(50000)}</w:p>`, noteParts), 'docx'));
    await timed('docx: 40,000 NOTEREF fields to one note are read', () => parseQuiet(docxOf(`<w:p><w:bookmarkStart w:id="0" w:name="a"/><w:r><w:footnoteReference w:id="1"/></w:r><w:bookmarkEnd w:id="0"/>${'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> NOTEREF a \\f \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'.repeat(40000)}</w:p>`, noteParts), 'docx'));
    // Alternate content nested deep, in a run and between blocks, and a Choice requiring 100,000 namespaces.
    await timed('docx: alternate content nested 20,000 deep in a run is read', () => parseQuiet(docxOf(`<w:p><w:r>${'<mc:AlternateContent><mc:Choice Requires="w">'.repeat(20000)}<w:t>x</w:t>${'</mc:Choice></mc:AlternateContent>'.repeat(20000)}</w:r></w:p>`), 'docx'));
    await timed('docx: alternate content nested 20,000 deep between blocks is read', () => parseQuiet(docxOf(`${'<mc:AlternateContent><mc:Choice Requires="w">'.repeat(20000)}<w:p><w:r><w:t>x</w:t></w:r></w:p>${'</mc:Choice></mc:AlternateContent>'.repeat(20000)}`), 'docx'));
    await timed('docx: 5,000 choices each requiring 1,000 namespaces are read', () => parseQuiet(docxOf(`<w:p><w:r>${`<mc:AlternateContent><mc:Choice Requires="${Array.from({ length: 1000 }, (_, i) => `p${i}`).join(' ')}"><w:t>c</w:t></mc:Choice><mc:Fallback><w:t>f</w:t></mc:Fallback></mc:AlternateContent>`.repeat(5000)}</w:r></w:p>`), 'docx'));
    // XLSX: a rich shared string's run formatting counts toward the repeated-content budget per cell.
    const sharedFontRuns = await parseQuiet(xlsxOf({
        'xl/sharedStrings.xml': `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><rFont val="${'F'.repeat(65536)}"/></rPr><t>a</t></r></si></sst>`,
        'xl/worksheets/sheet1.xml': `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${Array.from({ length: 7500 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="s"><v>0</v></c></row>`).join('')}</sheetData></worksheet>`,
    }), 'xlsx');
    const fontHtml = sharedFontRuns.ast ? ((await OfficeGenerator.generate(sharedFontRuns.ast, 'html', { onWarning: () => {} } as any)).value as string).length : -1;
    check('xlsx: one run\'s 64 KB font shown in 7,500 cells stays within the repeated-content budget', !sharedFontRuns.error && fontHtml > 0 && fontHtml < 40_000_000, `${sharedFontRuns.error} ${fontHtml}`);
    // ODS: a text box in a cell's paragraph keeps its text.
    const cellBox = await parseQuiet(odfOf('spreadsheet', '<table:table table:name="S"><table:table-row><table:table-cell><text:p>A<draw:frame><draw:text-box><text:p>INNER</text:p></draw:text-box></draw:frame>B</text:p></table:table-cell></table:table-row></table:table>'), 'ods');
    check('ods: a text box in a cell keeps its text', cellsOf(cellBox.ast)[0]?.text === 'AINNERB', JSON.stringify(cellsOf(cellBox.ast)[0]?.text));
    // Markdown: a code block's language cannot open a tag.
    const infoString = (await OfficeGenerator.generate(astWith([{ type: 'code', text: 'x', metadata: { language: 'js"><img src=x onerror=alert(1)>' } }]), 'md', { onWarning: () => {} } as any)).value as string;
    check('md: a code block\'s language holds no < or >', !/[<>]/.test(infoString.split('\n').find(line => line.startsWith('```')) ?? ''), infoString);
    // RTF: headers and footers go to the auxiliary, annotations are comments, shape text is read, and a
    // Unicode character's fallback is skipped.
    const rtfParts = await parseQuiet(Buffer.from(String.raw`{\rtf1\ansi {\header HEADERTEXT\par}{\headerl LEFTHEADER\par}{\footerf FIRSTFOOTER\par}Body{\*\atnauthor ANN}\chatn{\*\annotation COMMENTTEXT}\par{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt SHAPETEXT\par}}}\par A\u8364\'80B \u1055\'3f\u1088\'3f\par}`), 'rtf');
    const rtfBody = JSON.stringify(rtfParts.ast?.content ?? []);
    const rtfAux = JSON.stringify(rtfParts.ast?.auxiliary ?? {});
    check('rtf: headers and footers are read into the auxiliary, not the body', !rtfParts.error && ['HEADERTEXT', 'LEFTHEADER', 'FIRSTFOOTER'].every(m => rtfAux.includes(m) && !rtfBody.includes(m)), `${rtfParts.error} ${rtfAux}`);
    check('rtf: an annotation is a comment with its author, and shape text is read', /"type":"comment"[^]*COMMENTTEXT/.test(rtfBody) && rtfBody.includes('"author":"ANN"') && rtfBody.includes('SHAPETEXT') && !rtfBody.includes('shapeType'), rtfBody.slice(0, 400));
    check('rtf: a Unicode character\'s fallback is skipped', rtfBody.includes('A€B') && rtfBody.includes('Пр') && !rtfBody.includes('П?'), rtfBody.slice(-300));
    // Each destination sets the flow around it aside and back in constant time: 20,000 comments, notes
    // and text boxes in one table's cells, then in one list item, are read in linear time and leave the
    // table whole and the item an item.
    const destinationsStarted = Date.now();
    const destinationCell = String.raw`\pard\intbl A{\*\atnid J}\chatn{\*\annotation\pard\plain R\par}{\footnote\pard\plain N\par}{\shp{\*\shpinst{\shptxt\pard\plain T\par}}}\cell`;
    const destinationItem = String.raw`{\*\atnid J}\chatn{\*\annotation\pard\plain R\par}{\footnote\pard\plain N\par}{\shp{\*\shpinst{\shptxt\pard\plain T\par}}}`;
    const manyDestinations = await parseQuiet(Buffer.from(String.raw`{\rtf1\ansi\trowd\cellx3000` + destinationCell.repeat(20000) + String.raw`\row\pard\ls1\ilvl0 Item` + destinationItem.repeat(20000) + String.raw`\par}`), 'rtf');
    // A stated table depth is held to the reader's nesting limit: 20,000 paragraphs alternating between
    // \itap of four billion and \itap1, and 20,000 merged-down cells, are read in linear time and bounded size.
    const depthStarted = Date.now();
    const deepTables = await parseQuiet(Buffer.from(String.raw`{\rtf1\ansi` + String.raw`\pard\intbl\itap4294967295 x\par\pard\intbl y\par`.repeat(20000) + String.raw`\trowd\clvmgf\cellx100\pard\intbl a\cell\row` + String.raw`\trowd\clvmrg\cellx100\pard\intbl b\cell\row`.repeat(20000) + '}'), 'rtf');
    const deepSize = JSON.stringify(deepTables.ast?.content ?? []).length;
    check('rtf: 20,000 paragraphs at a stated depth of billions and 20,000 merged cells are read in linear time and size', !deepTables.error && Date.now() - depthStarted < 5000 && deepSize < 100_000_000, `${Date.now() - depthStarted}ms ${deepTables.error} ${deepSize}`);
    const destinationBlocks = manyDestinations.ast?.content ?? [];
    check('rtf: 20,000 comments, notes and text boxes in table cells and a list item are read in linear time', !manyDestinations.error && Date.now() - destinationsStarted < 5000 && destinationBlocks[0]?.type === 'table' && destinationBlocks[0]?.children?.[0]?.children?.length === 20000 && destinationBlocks[1]?.type === 'list', `${Date.now() - destinationsStarted}ms ${manyDestinations.error} ${destinationBlocks.slice(0, 2).map((n: any) => n.type)}`);
    // DOCX tracked changes read as the document now stands: insertions in, deletions and moved-away text out.
    const tracked = await parseQuiet(docxOf('<w:p><w:r><w:t xml:space="preserve">keep </w:t></w:r><w:moveFrom w:id="1" w:author="a"><w:r><w:t>MOVED</w:t></w:r></w:moveFrom><w:del w:id="2" w:author="a"><w:r><w:delText>DELETED</w:delText></w:r></w:del><w:ins w:id="3" w:author="a"><w:r><w:t>INSERTED</w:t></w:r></w:ins></w:p><w:p><w:moveTo w:id="4" w:author="a"><w:r><w:t>MOVED</w:t></w:r></w:moveTo></w:p>'), 'docx');
    const trackedTexts = (tracked.ast?.content ?? []).map((n: any) => n.text);
    check('docx: tracked moves read once, where the text now stands', !tracked.error && JSON.stringify(trackedTexts) === JSON.stringify(['keep INSERTED', 'MOVED']), `${tracked.error} ${JSON.stringify(trackedTexts)}`);
    // The element budget counts comments and processing instructions, and a DOCX chunk's lines, control
    // words and parts, before they are read.
    const budgetError = async (label: string, buffer: Buffer, fileType: string) => {
        const started = Date.now();
        const result = await parseQuiet(buffer, fileType);
        check(`${label} is refused within the element budget`, /XML element limit/.test(result.error) && Date.now() - started < 10000, `${Date.now() - started}ms ${result.error || 'parsed'}`);
    };
    await budgetError('docx: 3 million processing instructions', docxOf(`<w:p><w:r><w:t>a${'<?x?>'.repeat(3_000_000)}</w:t></w:r></w:p>`), 'docx');
    await budgetError('docx: 3 million comments', docxOf(`<w:p><w:r><w:t>a${'<!---->'.repeat(3_000_000)}</w:t></w:r></w:p>`), 'docx');
    const chunkOf = (target: string, content: string) => docxOf('<w:altChunk r:id="c1"/>', { 'word/_rels/document.xml.rels': chunkRels(chunkRel('c1', target)), [`word/${target}`]: content });
    await budgetError('docx: an RTF chunk of 3 million paragraphs', chunkOf('chunk.rtf', '{\\rtf1 ' + '\\par a'.repeat(3_000_000) + '}'), 'docx');
    await budgetError('docx: a text chunk of 3 million lines', chunkOf('chunk.txt', 'a\n'.repeat(3_000_000)), 'docx');
    await budgetError('docx: an MHT chunk of 3 million parts', chunkOf('chunk.mht', 'Content-Type: multipart/related; boundary="a"\r\n\r\n' + '--a\n'.repeat(3_000_000)), 'docx');
    await budgetError('docx: an MHT chunk of 3 million header lines', chunkOf('chunk.mht', 'Content-Type: text/html\r\n' + ' x\r\n'.repeat(3_000_000) + '\r\n<p>x</p>'), 'docx');
    // HTML text split by markup read as nothing (stray end tags, comments) is one text node.
    const strayEnds = await parseQuiet(Buffer.from(zipSync({
        'mimetype': enc('application/epub+zip'),
        'META-INF/container.xml': enc('<?xml version="1.0"?><container><rootfiles><rootfile full-path="content.opf"/></rootfiles></container>'),
        'content.opf': enc('<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata><manifest><item id="c" href="c.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="c"/></spine></package>'),
        'c.xhtml': enc(`<html><body><p>${'a</b>'.repeat(1_000_000)}</p></body></html>`),
    })), 'epub');
    let strayNodes = 0;
    const countNodes = (ns: any[] | undefined) => ns?.forEach((n: any) => { strayNodes++; countNodes(n.children); });
    countNodes(strayEnds.ast?.content);
    check('epub: text split by a million stray end tags is one text node', !strayEnds.error && strayNodes < 10, `${strayEnds.error} ${strayNodes}`);
    // The template renderer rebuilds a paragraph from slices, not a character at a time.
    await heapBudget('template: one paragraph of 20 million characters around a placeholder renders', () => timed('template: one paragraph of 20 million characters renders', () => OfficeTemplate.render(docxOf(`<w:p><w:r><w:t>{{a}}${'x'.repeat(20_000_000)}</w:t></w:r></w:p>`), { data: { a: 'y' } })));
    // SmartArt: a diagram's text is read, once however many frames show it, whatever its connections.
    const dgmNs = 'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const dgmPoint = (id: string, text: string, type = '') => `<dgm:pt modelId="${id}"${type ? ` type="${type}"` : ''}>${text ? `<dgm:t><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dgm:t>` : ''}</dgm:pt>`;
    const dgmData = (points: string, connections: string) => `<?xml version="1.0"?><dgm:dataModel ${dgmNs}><dgm:ptLst>${points}</dgm:ptLst><dgm:cxnLst>${connections}</dgm:cxnLst></dgm:dataModel>`;
    const dgmFrame = '<p:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId5"/></a:graphicData></a:graphic></p:graphicFrame>';
    const smartArtPptx = (frames: number, data: string) => pptxOf({
        'ppt/slides/slide1.xml': `<?xml version="1.0"?><p:sld ${pns}><p:cSld><p:spTree>${dgmFrame.repeat(frames)}</p:spTree></p:cSld></p:sld>`,
        'ppt/slides/_rels/slide1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/></Relationships>',
        'ppt/diagrams/data1.xml': data,
    });
    const smartArt = await parseQuiet(smartArtPptx(1, dgmData(dgmPoint('0', '', 'doc') + dgmPoint('1', 'FIRST') + dgmPoint('2', 'SECOND') + dgmPoint('3', 'CHILD') + dgmPoint('p', 'PRESENTATION', 'pres'), '<dgm:cxn modelId="a" srcId="0" destId="2" srcOrd="1"/><dgm:cxn modelId="b" srcId="0" destId="1" srcOrd="0"/><dgm:cxn modelId="c" srcId="2" destId="3" srcOrd="0"/>')), 'pptx');
    const smartItems = (smartArt.ast?.content[0]?.children ?? []).map((n: any) => `${n.metadata?.indentation}:${n.text}`);
    check('pptx: SmartArt text is read in order and nested, without its presentation points', !smartArt.error && JSON.stringify(smartItems) === JSON.stringify(['0:FIRST', '0:SECOND', '1:CHILD']), `${smartArt.error} ${JSON.stringify(smartItems)}`);
    const manyFrames = await parseQuiet(smartArtPptx(100000, dgmData(dgmPoint('0', '', 'doc') + dgmPoint('1', 'x'.repeat(100000)), '<dgm:cxn modelId="a" srcId="0" destId="1"/>')), 'pptx');
    check('pptx: one SmartArt shown by 100,000 frames is read once', !manyFrames.error && (manyFrames.ast?.content[0]?.children?.length ?? 0) === 1, `${manyFrames.error} ${manyFrames.ast?.content[0]?.children?.length}`);
    const chainPoints = Array.from({ length: 100000 }, (_, i) => dgmPoint(String(i + 1), 'n')).join('');
    const chainLinks = Array.from({ length: 100000 }, (_, i) => `<dgm:cxn modelId="c${i}" srcId="${i}" destId="${i + 1}"/>`).join('') + '<dgm:cxn modelId="loop" srcId="100000" destId="1"/>';
    await timed('pptx: SmartArt of a 100,000-point chain looping back is read', () => parseQuiet(smartArtPptx(1, dgmData(dgmPoint('0', '', 'doc') + chainPoints, chainLinks)), 'pptx'));
    const smartArtDocx = await parseQuiet(docxOf('<w:p><w:r><w:t>Intro</w:t></w:r><w:r><w:drawing><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId9"/></a:graphicData></a:graphic></w:drawing></w:r></w:p>', {
        'word/_rels/document.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="x" Target="diagrams/data1.xml"/></Relationships>',
        'word/diagrams/data1.xml': dgmData(dgmPoint('0', '', 'doc') + dgmPoint('1', 'ALPHA'), '<dgm:cxn modelId="a" srcId="0" destId="1"/>'),
    }), 'docx');
    check('docx: SmartArt text is a list after its paragraph', !smartArtDocx.error && JSON.stringify(smartArtDocx.ast?.content.map(n => [n.type, n.text])) === JSON.stringify([['paragraph', 'Intro'], ['list', 'ALPHA']]), JSON.stringify(smartArtDocx.ast?.content.map(n => [n.type, n.text])));
    // PPTX: modern comments (PowerPoint 365) and their replies are read with their authors.
    const p188Ns = 'xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
    const modernComments = await parseQuiet(pptxOf({
        'ppt/slides/_rels/slide1.xml.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/modernComment_100_ABC.xml"/></Relationships>',
        'ppt/comments/modernComment_100_ABC.xml': `<?xml version="1.0"?><p188:cmLst ${p188Ns}><p188:cm id="{1}" authorId="{A1}"><p188:txBody><a:p><a:r><a:t>MODERN</a:t></a:r></a:p></p188:txBody><p188:replyLst><p188:reply id="{2}" authorId="{A2}"><p188:txBody><a:p><a:r><a:t>REPLY</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst></p188:cm></p188:cmLst>`,
        'ppt/authors.xml': `<?xml version="1.0"?><p188:authorLst ${p188Ns}><p188:author id="{A1}" name="Ann" initials="A"/><p188:author id="{A2}" name="Rae" initials="R"/></p188:authorLst>`,
    }), 'pptx');
    const modernList = JSON.stringify((modernComments.ast?.content[0]?.comments ?? []).map((c: any) => [c.text, c.metadata?.author]));
    check('pptx: modern comments and replies are read with their authors', !modernComments.error && modernList === JSON.stringify([['MODERN', 'Ann'], ['REPLY', 'Rae']]) && modernComments.ast?.content.length === 1, `${modernComments.error} ${modernList}`);
    // RTF: a \uc fallback is kept as one slice and held to 16 characters; nested fields are searched once.
    await timed('rtf: a fallback of \\uc2000000000 before 20 MB of text is read', () => parseQuiet(Buffer.from('{\\rtf1\\uc2000000000\\u65 ' + 'a'.repeat(20_000_000) + '}'), 'rtf', { includeRawContent: true }));
    const ucCapped = await parseQuiet(Buffer.from('{\\rtf1\\uc2000000000\\u65 ' + 'a'.repeat(40) + '}'), 'rtf');
    check('rtf: a \\uc past 16 skips 16 characters', (ucCapped.ast?.content[0] as any)?.text === 'A' + 'a'.repeat(24), JSON.stringify((ucCapped.ast?.content[0] as any)?.text));
    const nestedFields = '{\\rtf1 ' + '{\\field '.repeat(250) + '{' + 'a'.repeat(4_000_000) + '}' + '}'.repeat(251);
    await timed('rtf: 4 MB of text in fields nested 250 deep is read', () => parseQuiet(Buffer.from(nestedFields), 'rtf'));
    // HTML: an end tag closing nothing costs nothing, however deep the element it stands in.
    await timed('html: 4 million stray end tags 250 elements deep parse', () => parseQuiet(Buffer.from('<html><body>' + '<b>'.repeat(250) + '</x>'.repeat(4_000_000) + 'z'), 'html'));
    // Templates: a value repeated at many placeholders is charged past its first use.
    const repeatedTemplate = docxOf(`<w:p><w:r><w:t>${'{{a}}'.repeat(200000)}</w:t></w:r></w:p>`);
    let repeatError: any;
    const repeatStarted = Date.now();
    try { await OfficeTemplate.render(repeatedTemplate, { data: { a: 'x'.repeat(100) } }); } catch (e) { repeatError = e; }
    check('template: a 100-character value at 200,000 placeholders is refused', repeatError?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - repeatStarted < 5000, `${Date.now() - repeatStarted}ms ${repeatError}`);
    const fewRepeats = await OfficeTemplate.render(docxOf(`<w:p><w:r><w:t>${'{{a}} '.repeat(10)}</w:t></w:r></w:p>`), { data: { a: 'x'.repeat(1000) } });
    check('template: a long value at ten placeholders renders', fewRepeats instanceof Uint8Array && fewRepeats.length > 0, String(fewRepeats));
    // HTML: a tag whose '>' stands between quotes is scanned for to the end at most three times.
    await timed('html: 120,000 tags each closed only inside quotes parse', () => parseQuiet(Buffer.from('<p>' + '<br ">"'.repeat(120000)), 'html'));
    await timed('html: 120,000 end tags each closed only inside quotes parse', () => parseQuiet(Buffer.from('<p>' + '</b ">"'.repeat(120000)), 'html'));
    // HTML: content under many wrappers is read once, not again at every level.
    for (const [label, html] of [
        ['paragraphs under 250 divs', '<div>'.repeat(250) + '<p>x</p>'.repeat(300000)],
        ['paragraphs under 250 divs each holding a paragraph', '<div><p>a</p>'.repeat(250) + '<p>x</p>'.repeat(300000)],
        ['items under 250 lists', '<ul>'.repeat(250) + '<li>x</li>'.repeat(300000)],
        ['paragraphs under 250 quotes', '<blockquote>'.repeat(250) + '<p>x</p>'.repeat(300000)],
        ['text under 250 links', '<p>' + '<a href="#a">'.repeat(250) + 'x <br>'.repeat(300000)],
    ] as const) await timed(`html: ${label} parse`, () => parseQuiet(Buffer.from(html), 'html'));
    const nestedMarks = await parseQuiet(Buffer.from('<div id="outer"><div id="inner"><p>x</p><p>y</p></div></div><p><abbr title="outer"><abbr title="inner">z</abbr></abbr></p>'), 'html');
    const firstIds = (nestedMarks.ast?.content[0]?.metadata as any)?.anchorIds;
    check('html: nested wrappers\' ids mark their first block, outermost first', JSON.stringify(firstIds) === JSON.stringify(['outer', 'inner']), JSON.stringify(firstIds));
    const abbrText = JSON.stringify(nestedMarks.ast?.content ?? []);
    check('html: nested abbreviations give their text the innermost title', abbrText.includes('"abbreviationTitle":"inner"') && !abbrText.includes('"abbreviationTitle":"outer"'), abbrText.slice(-300));
    const htmlNestedLinks = await parseQuiet(Buffer.from('<p><a href="https://one.example/">one <a href="https://two.example/">two</a></a></p>'), 'html');
    const linkJson = JSON.stringify(htmlNestedLinks.ast?.content ?? []);
    check('html: a link started inside a link ends the first, as a browser reads it', /"text":"one ".*"link":"https:\/\/one\.example\/"/.test(linkJson) && /"text":"two".*"link":"https:\/\/two\.example\/"/.test(linkJson), linkJson);
    // ODF: runs inside spans, links and fields nested in each other are read into the paragraph once.
    for (const [open, close] of [['<text:span>', '</text:span>'], ['<text:a xlink:href="#x">', '</text:a>'], ['<text:meta>', '</text:meta>']] as const) {
        await timed(`odt: 600,000 runs inside ${open} nested 1,000 deep parse`, () => parseQuiet(odfOf('text', `<text:p>${open.repeat(1000)}${'x<text:s/>'.repeat(300000)}${close.repeat(1000)}</text:p>`), 'odt'));
    }
    // A note every reference shares is resolved once (HTML, EPUB, Markdown).
    await timed('html: 16,000 references to a note of 16,000 paragraphs parse', () => parseQuiet(Buffer.from('<p>a<sup data-footnote-ref="k"></sup></p>'.repeat(16000) + '<section data-footnotes><div data-footnote-id="k">' + '<p>x</p>'.repeat(16000) + '</div></section>'), 'html'));
    await timed('md: 16,000 references to a footnote of 16,000 paragraphs parse', () => parseQuiet(Buffer.from('a[^k]\n\n'.repeat(16000) + '[^k]: x\n' + '\n    y\n'.repeat(16000)), 'md'));
    // Generators: an AST sharing nodes of an unknown type along many paths is refused before any pass expands it.
    let sharedUnknown: any = { type: 'text', text: 'a' };
    for (let i = 0; i < 40; i++) sharedUnknown = { type: 'wrap', children: [sharedUnknown, sharedUnknown] };
    for (const limit of [undefined, Infinity, 'x']) {
        let unknownError: any;
        const unknownStarted = Date.now();
        const withConfig = { ...astWith([{ type: 'paragraph', children: [sharedUnknown] }]), ...(limit !== undefined && { config: { decompressionLimits: { maxRepeatedContent: limit } } }) } as any;
        try { await OfficeGenerator.generate(withConfig, 'text', { onWarning: () => {} } as any); } catch (e) { unknownError = e; }
        check(`generators: unknown nodes shared along 2^40 paths are refused (repeat limit ${String(limit)})`, unknownError?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - unknownStarted < 2000, `${Date.now() - unknownStarted}ms ${unknownError}`);
    }
    // Generators: unknown wrappers nested 1,000 deep are written as their content once.
    let unknownWrapped: any = Array.from({ length: 100000 }, () => ({ type: 'text', text: 'a' }));
    for (let i = 0; i < 1000; i++) unknownWrapped = [{ type: 'wrap', formatting: { bold: true }, children: unknownWrapped }];
    const wrappedStarted = Date.now();
    const wrappedText = (await OfficeGenerator.generate(astWith([{ type: 'paragraph', children: unknownWrapped }]), 'text', { onWarning: () => {} } as any)).value as string;
    check('generators: 100,000 runs under 1,000 unknown wrappers are written in linear time', wrappedText.length >= 100000 && Date.now() - wrappedStarted < 5000, `${Date.now() - wrappedStarted}ms ${wrappedText.length}`);
    // HTML: the innermost link is the one its text follows.
    const innerLink = await parseQuiet(Buffer.from('<table><tr><td><a href="https://outer.example/"><table><tr><td><a href="https://inner.example/">in</a></td></tr></table></a></td></tr></table>'), 'html');
    const innerJson = JSON.stringify(innerLink.ast?.content ?? []);
    check('html: text in a link inside a link follows the inner one', innerJson.includes('inner.example') && !/"text":"in"[^}]*outer\.example/.test(innerJson), innerJson.slice(0, 400));
    // Generators: unknown nodes shared through notes are written once, notes carried up nested wrappers
    // are not copied again at every level, and shared arrays are read once by the sharing check.
    let noteShared: any = { type: 'wrap', text: 'leaf' };
    for (let i = 0; i < 60; i++) noteShared = { type: 'wrap', text: 'x', notes: [noteShared, noteShared] };
    await timed('generators: unknown nodes each holding the next twice in their notes, 60 deep, are written', () => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [noteShared] }]), 'text', { onWarning: () => {} } as any));
    const sharedNoteList = Array(16000).fill({ type: 'note', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'n' }] }] });
    let carried: any = { type: 'text', text: 'leaf' };
    for (let i = 0; i < 1000; i++) carried = { type: 'wrap', notes: sharedNoteList, children: [carried] };
    let carriedError: any;
    const carriedStarted = Date.now();
    try { await OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [carried] }]), 'text', { onWarning: () => {} } as any); } catch (e) { carriedError = e; }
    check('generators: 1,000 wrappers holding one list of 16,000 notes are refused at once', carriedError?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - carriedStarted < 2000, `${Date.now() - carriedStarted}ms ${carriedError}`);
    let fewCarried: any = { type: 'text', text: 'leaf' };
    for (let i = 0; i < 1000; i++) fewCarried = { type: 'wrap', notes: [{ type: 'note', children: [{ type: 'text', text: 'n' + i }] }], children: [fewCarried] };
    await timed('generators: 1,000 nested unknown wrappers each carrying a note are written', () => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [fewCarried] }]), 'md', { onWarning: () => {} } as any));
    const sharedKids = Array.from({ length: 100000 }, () => ({ type: 'text', text: 'x' }));
    let arrayError: any;
    const arrayStarted = Date.now();
    try { await OfficeGenerator.generate(astWith(Array.from({ length: 1000000 }, () => ({ type: 'paragraph', children: sharedKids, notes: sharedNoteList }))), 'text', { onWarning: () => {} } as any); } catch (e) { arrayError = e; }
    check('generators: a million nodes sharing one list of 100,000 children are refused at once', arrayError?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - arrayStarted < 5000, `${Date.now() - arrayStarted}ms ${arrayError}`);
    // Generators: values shared across nodes (records, id lists, note lists, slides' notes, nested and
    // sparse arrays) count at every path and node holding them, and each is read once.
    const refusedQuickly = async (label: string, content: any[], format = 'text', extra: object = {}) => {
        let refusal: any;
        const started = Date.now();
        try { await OfficeGenerator.generate({ ...astWith(content), ...extra } as any, format as any, { onWarning: () => {} } as any); } catch (e) { refusal = e; }
        check(`generators: ${label} is refused at once`, refusal?.officeIssue?.code === 'OUTPUT_TOO_LARGE' && Date.now() - started < 3000, `${Date.now() - started}ms ${refusal}`);
    };
    const bigRecord = Object.fromEntries(Array.from({ length: 20000 }, (_, i) => [`k${i}`, 1]));
    await refusedQuickly('one record of 20,000 fields shared by 20,000 runs', [{ type: 'paragraph', children: Array.from({ length: 20000 }, () => ({ type: 'text', text: 'x', formatting: bigRecord })) }], 'csv');
    const sharedIds = Array.from({ length: 3000 }, (_, i) => `id${i}`);
    await refusedQuickly('one list of 3,000 ids shared by 3,000 paragraphs', Array.from({ length: 3000 }, () => ({ type: 'paragraph', metadata: { anchorIds: sharedIds }, children: [{ type: 'text', text: 'x' }] })), 'docx');
    const smallNote = { type: 'note', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'n' }] }] };
    const referencing = { type: 'text', text: 'x', notes: Array(3000).fill(smallNote) };
    await refusedQuickly('a run holding 3,000 note references, shared 3,000 times', [{ type: 'paragraph', children: Array(3000).fill(referencing) }], 'odt');
    const sharedSlide = { type: 'slide', metadata: { slideNumber: 1 }, children: [], notes: Array(10000).fill(smallNote) };
    await refusedQuickly('a slide of 10,000 notes shared 10,000 times', Array(10000).fill(sharedSlide), 'rtf');
    let nested: any = ['v'];
    for (let i = 0; i < 8; i++) nested = Array(10).fill(nested);
    await refusedQuickly('a style that is ten lists nested eight deep', [{ type: 'paragraph', children: [{ type: 'text', text: 'x', metadata: { style: nested } }] }], 'csv');
    const sparse: any[] = [];
    sparse.length = 4294967295;
    sparse[0] = { type: 'text', text: 'x' };
    await refusedQuickly('a sparse list of four billion slots', [{ type: 'paragraph', children: sparse }], 'csv');
    const longText = 'x'.repeat(10_000_000);
    await refusedQuickly('one 10 MB text held by 1,000 runs', [{ type: 'paragraph', children: Array.from({ length: 1000 }, () => ({ type: 'text', text: longText })) }], 'docx');
    // Values that JSON can hold are read in time linear in the JSON: a wrapper's thousands of unknown
    // formatting fields, a paragraph's thousands of metadata fields split around blocks, and a chart
    // whose labels are lists.
    const manyFields = JSON.parse(JSON.stringify(astWith([{ type: 'paragraph', children: [{ type: 'wrap', formatting: Object.fromEntries(Array.from({ length: 16000 }, (_, i) => [`k${i}`, 1])), children: Array.from({ length: 16000 }, () => ({ type: 'text', text: 'x' })) }] }])));
    await timed('generators: a wrapper of 16,000 formatting fields around 16,000 runs is written', () => OfficeGenerator.generate(manyFields, 'csv', { onWarning: () => {} } as any));
    const splitFields = JSON.parse(JSON.stringify(astWith([{ type: 'paragraph', metadata: Object.fromEntries(Array.from({ length: 8000 }, (_, i) => [`m${i}`, 'v'])), children: Array.from({ length: 8000 }, () => [{ type: 'text', text: 'x' }, { type: 'code', text: 'y', metadata: { math: 'block' } }]).flat() }])));
    await timed('html: a paragraph of 8,000 metadata fields split around 8,000 blocks is written', () => OfficeGenerator.generate(splitFields, 'html', { onWarning: () => {} } as any));
    const listLabels = Array(10).fill(Array(10).fill(Array(10).fill(Array(10).fill(Array(10).fill('l')))));
    const chartAst = { ...astWith([{ type: 'chart', metadata: { attachmentName: 'c' } }]), attachments: [{ type: 'chart', name: 'c', mimeType: 'application/json', data: '', extension: 'json', chartData: { labels: Array(2000).fill(listLabels), dataSets: [{ name: listLabels, values: Array(2000).fill(listLabels), pointLabels: [] }], rawTexts: [] } }] } as any;
    await timed('tex, docx: a chart whose labels are nested lists is written', async () => { for (const format of ['tex', 'docx'] as const) await OfficeGenerator.generate(chartAst, format, { onWarning: () => {} } as any); });
    // CSV: the AST's own delimiter is one character that neither ends a row nor starts a formula.
    const astDelimited = (await OfficeGenerator.generate({ ...astWith([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [{ type: 'text', text: 'a' }] }] }, { type: 'cell', children: [{ type: 'paragraph', children: [{ type: 'text', text: "cmd|' /C calc'!A0" }] }] }] }] }]), config: { csvDelimiter: '\n=' } } as any, 'csv', { onWarning: () => {} } as any)).value as string;
    check('csv: an AST\'s delimiter cannot start a row or a formula', !/(^|\n)=/.test(astDelimited), JSON.stringify(astDelimited));
    // ODF: long part paths are indexed by their last folders only.
    const longPaths: Record<string, string> = { 'Obj/content.xml': '' };
    for (let i = 0; i < 20; i++) longPaths[`${i}/${'a/'.repeat(30000)}content.xml`] = '';
    await timed('odt: 20 parts of 30,000 folders each are indexed', () => parseQuiet(odfOf('text', '<text:p><draw:frame><draw:object xlink:href="./Obj"/></draw:frame></text:p>', longPaths), 'odt'));
    // Text layout: a page's line is set at most 1,000 columns wide, and its padding draws on a budget.
    const farApart = Array.from({ length: 2000 }, (_, i) => [{ type: 'text', text: 'abcdef', bounds: { x: 0, y: i * 20, width: 12, height: 10 } }, { type: 'text', text: 'ghijkl', bounds: { x: 1_000_000, y: i * 20, width: 12, height: 10 } }]).flat();
    const laidOut = (await OfficeGenerator.generate({ ...astWith([{ type: 'page', metadata: { pageNumber: 1, pageWidth: 2_000_000, pageHeight: 50_000 }, children: farApart }]), type: 'pdf' } as any, 'text', { onWarning: () => {} } as any)).value as string;
    check('text: 2,000 lines of two runs a million points apart are laid out narrow', laidOut.length < 4_000_000 && laidOut.includes('ghijkl'), `${laidOut.length}`);
    // One picture shown many times is decoded and embedded once per document.
    const { deflateSync: deflatePng, crc32 } = require('zlib') as typeof import('zlib');
    const pngChunk = (kind: string, data: Buffer) => { const length = Buffer.alloc(4); length.writeUInt32BE(data.length); const body = Buffer.concat([Buffer.from(kind), data]); const sum = Buffer.alloc(4); sum.writeUInt32BE(crc32(body) >>> 0); return Buffer.concat([length, body, sum]); };
    const noisePng = (side: number) => {
        const raw = Buffer.alloc((side * 3 + 1) * side);
        for (let i = 0; i < raw.length; i++) raw[i] = i % (side * 3 + 1) === 0 ? 0 : (i * 2654435761) >>> 24;
        const header = Buffer.alloc(13); header.writeUInt32BE(side, 0); header.writeUInt32BE(side, 4); header[8] = 8; header[9] = 2;
        return Buffer.concat([Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]), pngChunk('IHDR', header), pngChunk('IDAT', deflatePng(raw, { level: 1 })), pngChunk('IEND', Buffer.alloc(0))]);
    };
    const picture = noisePng(400).toString('base64');
    const onePictureShown = (times: number) => ({ ...astWith(Array.from({ length: times }, () => ({ type: 'image', metadata: { attachmentName: 'a.png' } }))), attachments: [{ type: 'image', name: 'a.png', mimeType: 'image/png', extension: 'png', data: picture }] }) as any;
    const nativePdf = (await OfficeGenerator.generate(onePictureShown(100), 'pdf', { onWarning: () => {}, pdfConfig: { engine: 'native' } } as any)).value as Uint8Array;
    check('pdf (native): a picture shown 100 times is embedded once', nativePdf.byteLength < 3 * picture.length, `${nativePdf.byteLength} vs ${picture.length}`);
    await timed('docx, odt: a picture shown 5,000 times is decoded once', async () => { for (const format of ['docx', 'odt'] as const) await OfficeGenerator.generate(onePictureShown(5000), format, { onWarning: () => {} } as any); });
    // The document's custom properties and a chart's data are weighed per key and per series.
    const sharedValues = Array.from({ length: 10000 }, (_, i) => `v${i}`);
    await refusedQuickly('10,000 custom properties sharing one list of 10,000 values', [], 'html', { metadata: { customProperties: Object.fromEntries(Array.from({ length: 10000 }, (_, i) => [`p${i}`, sharedValues])) } });
    const sharedSeries = { name: 's', values: Array.from({ length: 100000 }, (_, i) => i), pointLabels: [] };
    await refusedQuickly('1,000 chart series sharing one list of 100,000 values', [{ type: 'chart', metadata: { attachmentName: 'c' } }], 'html', { attachments: [{ type: 'chart', name: 'c', mimeType: 'application/json', data: '', extension: 'json', chartData: { labels: [], dataSets: Array(1000).fill(sharedSeries), rawTexts: [] } }] });
    // Records are weighed by what writers write at each holder (keys, short values), a list item by its
    // indentation, and a table by the empty positions its grid fills, at every path to it.
    const indented = { type: 'list', metadata: { listType: 'unordered', indentation: 64 }, children: [] };
    await refusedQuickly('a list item indented 64 levels, shared 2,000,000 times', Array(2_000_000).fill(indented), 'md');
    const longKey = { type: 'paragraph', htmlAttributes: { [`data-${'a'.repeat(65536)}`]: 'v' }, children: [{ type: 'text', text: 'x' }] };
    await refusedQuickly('an attribute of a 64 KB name, shared by 7,000 paragraphs', Array(7000).fill(longKey), 'html');
    const shortAttributes = { type: 'paragraph', htmlAttributes: Object.fromEntries(Array.from({ length: 32 }, (_, i) => [`data-k${i}`, 'v'.repeat(64)])), children: [] };
    await refusedQuickly('32 attributes of 64 characters, shared by 200,000 paragraphs', Array(200_000).fill(shortAttributes), 'html');
    const diagonal = { type: 'table', children: Array.from({ length: 1000 }, (_, i) => ({ type: 'row', children: [{ type: 'cell', metadata: { row: i, col: i }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'x' }] }] }] })) };
    await refusedQuickly('a table of 1,000 cells along a diagonal, shared 40 times', Array(40).fill(diagonal), 'csv');
    const oneDiagonal = (await OfficeGenerator.generate(astWith([diagonal]) as any, 'csv', { onWarning: () => {} } as any)).value as string;
    check('csv: a table of 1,000 cells along a diagonal is written once', oneDiagonal.length > 1_000_000, `${oneDiagonal.length}`);
    // A style record every cell of a workbook shares (as a parser gives it) is written, not refused: the
    // AST holds it once, and at each later cell it weighs only what it weighs past what a cell's own
    // record may. Counted in full at every cell, 700,000 styled cells were refused.
    const cellStyle = { bold: true, italic: true, underline: true, strikethrough: true, color: '#1F3864', backgroundColor: '#D9E1F2', size: '10.5pt', font: 'Segoe UI Semibold Condensed', alignment: 'center' };
    const styledSheet = { type: 'sheet', children: Array.from({ length: 100_000 }, (_, r) => ({ type: 'row', children: Array.from({ length: 7 }, (_, c) => ({ type: 'cell', metadata: { row: r, col: c }, children: [{ type: 'text', text: String(r * 7 + c), formatting: cellStyle }] })) })) };
    let styledCsv = '', styledError: any;
    try { styledCsv = (await OfficeGenerator.generate(astWith([styledSheet]) as any, 'csv', { onWarning: () => {} } as any)).value as string; } catch (e) { styledError = e; }
    check('csv: 700,000 cells sharing one style record are written', !styledError && styledCsv.split('\n').filter(line => /^\d/.test(line)).length === 100_000 && styledCsv.includes('699999'), `${styledError} ${styledCsv.length}`);
    // A grid too large for the budget is laid out closer, and that is reported, never silent; a parsed
    // sparse sheet gets the budget its size allows (16 positions a byte) and keeps its columns.
    const cellsFarApart = { type: 'sheet', children: [{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [{ type: 'text', text: 'a' }] }] }, { type: 'row', children: [{ type: 'cell', metadata: { row: 5000, col: 5000 }, children: [{ type: 'text', text: 'b' }] }] }] };
    const gridIssues: string[] = [];
    const farCsv = (await OfficeGenerator.generate(astWith([cellsFarApart]) as any, 'csv', { onWarning: (issue: any) => gridIssues.push(issue.code) } as any)).value as string;
    check('csv: cells too far apart for the grid budget are laid out closer, with TABLE_GRID_LIMIT_EXCEEDED', farCsv.length < 100 && gridIssues.includes('TABLE_GRID_LIMIT_EXCEEDED'), `${farCsv.length} ${gridIssues}`);
    const sparseRows = Array.from({ length: 3000 }, (_, r) => `<table:table-row><table:table-cell office:value-type="string"><text:p>id${r}</text:p></table:table-cell><table:table-cell table:number-columns-repeated="398"/><table:table-cell office:value-type="string"><text:p>v${r}</text:p></table:table-cell></table:table-row>`).join('');
    const wideSparseSheet = await warned(odfOf('spreadsheet', `<table:table table:name="Data">${sparseRows}</table:table>`), 'ods');
    const sparseIssues: string[] = [];
    const sparseCsv = wideSparseSheet.ast ? (await wideSparseSheet.ast.to('csv', { onWarning: (issue: any) => sparseIssues.push(issue.code) } as any)).value as string : '';
    check('ods: a sparse sheet of 3,000 rows and 400 columns keeps its columns', sparseCsv.split('\n')[0].split(',').length === 400 && !sparseIssues.includes('TABLE_GRID_LIMIT_EXCEEDED'), `${sparseCsv.split('\n')[0].split(',').length} ${sparseIssues}`);
    // Short rows are padded to their table's width within the budget, and past it that is reported.
    const tallNarrow = { type: 'table', children: [{ type: 'row', children: Array.from({ length: 13 }, (_, c) => ({ type: 'cell', children: [{ type: 'text', text: 'H' + c }] })) }, ...Array.from({ length: 100_000 }, (_, r) => ({ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: String(r) }] }] }))] };
    const padIssues: string[] = [];
    await OfficeGenerator.generate(astWith([tallNarrow]) as any, 'csv', { onWarning: (issue: any) => padIssues.push(issue.code) } as any);
    check('csv: rows left short by the padding budget are reported (TABLE_GRID_LIMIT_EXCEEDED)', padIssues.includes('TABLE_GRID_LIMIT_EXCEEDED'), `${padIssues}`);
    // Plain text lines a table's columns up within a budget of spaces (16 million, plus 16 a byte of the
    // document); past it cells are set off by one space, and that is reported.
    const wideColumn = { type: 'table', children: Array.from({ length: 70_000 }, (_, r) => ({ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: r === 0 ? 'y'.repeat(250) : 'x' }] }, { type: 'cell', children: [{ type: 'text', text: 'z' }] }] })) };
    const alignIssues: string[] = [];
    const aligned = (await OfficeGenerator.generate(astWith([wideColumn]) as any, 'text', { textConfig: { preserveLayout: true }, onWarning: (issue: any) => alignIssues.push(issue.code) } as any)).value as string;
    check('text: table alignment past its budget of spaces is reported (TABLE_GRID_LIMIT_EXCEEDED)', aligned.length < 20_000_000 && alignIssues.includes('TABLE_GRID_LIMIT_EXCEEDED'), `${aligned.length} ${alignIssues}`);
    const parsedNarrow = await warned(Buffer.from(['h0,h1,h2,h3,h4,h5,h6,h7,h8,h9,h10,h11,h12', ...Array.from({ length: 100_000 }, (_, r) => String(r))].join('\n')), 'csv');
    const narrowCsv = parsedNarrow.ast ? (await parsedNarrow.ast.to('csv')).value as string : '';
    check('csv: a parsed CSV of 100,000 short rows under a 13-column header is padded whole', narrowCsv.split('\n').filter(Boolean).every(line => line.split(',').length === 13), narrowCsv.slice(-80));
    // A link's `wikilink` is a boolean: false is not a wikilink (coerced to the string "false", it was).
    const notWiki = astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'Example', metadata: { link: 'https://example.com', linkType: 'external', wikilink: false } }] }]);
    const notWikiMd = (await OfficeGenerator.generate(notWiki as any, 'md', { onWarning: () => {} } as any)).value as string;
    check('md: a link with wikilink false is a Markdown link', notWikiMd.includes('[Example](https://example.com)') && !notWikiMd.includes('[['), notWikiMd);
    // A paragraph's text is the text of its runs, not written beside them: a paragraph of 96 MB in one run is written.
    const hugeText = 'x'.repeat(96_000_000);
    const hugeParagraph = (await OfficeGenerator.generate(astWith([{ type: 'paragraph', text: hugeText, children: [{ type: 'text', text: hugeText }] }]) as any, 'text', { onWarning: () => {} } as any)).value as string;
    check('text: a paragraph of 96 MB in one run is written, not refused as shared', hugeParagraph.length >= 96_000_000, `${hugeParagraph.length}`);
    // Text layout finds each list's first line from its children's, not by scanning every line under it.
    let deepList: any = { type: 'list', metadata: { listType: 'unordered' }, children: Array(200_000).fill({ type: 'text', text: 'abcdef', bounds: { x: 10, y: 10, width: 36, height: 10 } }) };
    for (let i = 1; i < 1000; i++) deepList = { type: 'list', metadata: { listType: 'unordered' }, children: [deepList] };
    await timed('text: lists nested 1,000 deep over 200,000 placed runs are laid out', () => OfficeGenerator.generate({ ...astWith([{ type: 'page', metadata: { pageWidth: 600, pageHeight: 800 }, children: [deepList] }]), type: 'pdf' } as any, 'text', { onWarning: () => {} } as any));
}

/**
 * The separate pdf.js processes and a document's budgets: a parse's abort signal reaches only that
 * parse, a parse that ended stops reading, the time budget charges the process's CPU time (a busy
 * host cuts nothing short), processes are pooled and end with the host, a parse in this process
 * leaves the next in a separate one, budget options are checked, and a drawing's operators never cost
 * the document its text.
 */
async function pdfProcessTests(h: {
    pdfFrom: (objects: (string | Buffer)[]) => Buffer;
    flate: (content: Buffer, dict?: string) => Buffer;
    pdfForms: (depth: number, leaf: string) => Buffer;
    drawnOften: Buffer;
    parseQuiet: (buffer: Buffer, fileType: string, extra?: object) => Promise<{ ast: OfficeParserAST | undefined; error: string }>;
    files: string;
}) {
    const { pdfFrom, flate, pdfForms, drawnOften, parseQuiet, files } = h;
    const read = async (buffer: Buffer, extra: object = {}) => {
        const warnings: { code: string; message: string }[] = [];
        const result = await parseQuiet(buffer, 'pdf', { ...extra, onWarning: (issue: any) => warnings.push({ code: issue.code, message: issue.message }) });
        const text = (result.ast?.content ?? []).map((p: any) => p.text).join('\n');
        return { ...result, warnings, codes: warnings.map(w => w.code), pages: result.ast?.content.length ?? 0, text };
    };
    const named = (warnings: { code: string; message: string }[], limit: string) => warnings.some(w => w.code === 'PDF_CONTENT_LIMIT_EXCEEDED' && w.message.includes(`pdfParserConfig.${limit}`));
    // Pages of notes, 50 lines each; with `drawing`, a first page stamping one symbol of 60 strokes 3,000 times (a plan).
    const notes = (pages: number, drawing: boolean) => {
        const objs: (string | Buffer)[] = ['<< /Type /Catalog /Pages 2 0 R >>', '', '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
        const kids: number[] = [];
        if (drawing) {
            let symbol = '';
            for (let i = 0; i < 60; i++) symbol += `${(0.2 + (i % 5) * 0.1).toFixed(1)} w ${(i % 3) / 2} ${(i % 4) / 3} ${(i % 7) / 6} RG ${i % 10} ${Math.floor(i / 10)} m ${(i * 7) % 10} ${(i * 3) % 10} l S\n`;
            let plan = 'BT /F1 14 Tf 40 800 Td (Floor plan) Tj ET\n';
            for (let k = 0; k < 3000; k++) plan += `q 1 0 0 1 ${20 + (k % 60) * 9} ${40 + Math.floor(k / 60) * 14} cm /S Do Q\n`;
            objs.push(flate(Buffer.from(symbol), '/Type /XObject /Subtype /Form /BBox [0 0 10 10]'), flate(Buffer.from(plan)));
            objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> /XObject << /S ${objs.length - 1} 0 R >> >> /Contents ${objs.length} 0 R >>`);
            kids.push(objs.length);
        }
        for (let p = 1; p <= pages; p++) {
            let body = 'BT /F1 10 Tf 12 TL 40 800 Td\n';
            for (let l = 1; l <= 50; l++) body += `(Note ${p}.${l}: fixture schedule item for room ${p * 50 + l}.) Tj T*\n`;
            objs.push(flate(Buffer.from(body + 'ET\n')));
            objs.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents ${objs.length} 0 R >>`);
            kids.push(objs.length);
        }
        objs[1] = `<< /Type /Pages /Kids [${kids.map(k => `${k} 0 R`).join(' ')}] /Count ${kids.length} >>`;
        return pdfFrom(objs);
    };
    const testPdf = fs.readFileSync(path.join(files, 'test.pdf'));

    // Operators: only those kept count, and running out of them stops operators, never text. A 21 KB
    // plan stamping one symbol 3,000 times gave its first page, 19 characters, and "no text" warnings.
    // (Separate processes first: pdf.js in this process comes last, at the end.)
    const plan = notes(20, true);
    const planned = await read(plan);
    check('pdf: a plan of 3,000 symbols is read whole, its paths costing none of the operator budget', !planned.error && planned.pages === 21 && planned.text.includes('Note 20.50') && !planned.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED') && !planned.codes.includes('PDF_NO_TEXT_EXTRACTED'), `${planned.pages} pages ${planned.codes} ${planned.error}`);
    const operatorChecks = async (separateProcess: boolean) => {
        const where = separateProcess ? 'separate process' : 'in process';
        const capped = await read(plan, { pdfParserConfig: { maxOperators: 0, separateProcess } });
        check(`pdf: past maxOperators every later page's text is still read (${where})`, !capped.error && capped.pages === 21 && capped.text.includes('Note 20.50') && named(capped.warnings, 'maxOperators') && !capped.codes.includes('PDF_NO_TEXT_EXTRACTED'), `${capped.pages} pages ${capped.codes} ${capped.error}`);
        // Read only for its fonts, a page's operator list stops once pdf.js has sent them: nested forms of
        // shapes cost nothing more, in this process too, where pdf.js has no time limit.
        let begun = Date.now();
        const fontsOnly = await read(pdfForms(8, '0 0 1 1 re f'), { pdfParserConfig: { extractTextColor: false, separateProcess } });
        check(`pdf: forms of shapes nested 8 deep, read for fonts only, stop once the fonts are known (${where})`, !fontsOnly.error && Date.now() - begun < 5000 && fontsOnly.text.includes('page'), `${Date.now() - begun}ms ${fontsOnly.codes} ${fontsOnly.error}`);
        begun = Date.now();
        const shapes = await read(pdfForms(8, '0 0 1 1 re f'), { pdfParserConfig: { separateProcess } });
        check(`pdf: forms of shapes nested 8 deep stop at maxOperators, keeping the page's text (${where})`, !shapes.error && Date.now() - begun < 5000 && shapes.text.includes('page') && named(shapes.warnings, 'maxOperators'), `${Date.now() - begun}ms ${shapes.codes} ${shapes.error}`);
    };
    await operatorChecks(true);
    const byText = await read(pdfForms(5, 'BT /F1 12 Tf 10 10 Td (xy) Tj ET'), { pdfParserConfig: { maxTextItems: 1 } });
    const byTime = await read(drawnOften, { pdfParserConfig: { maxTimeMs: 200 } });
    check('pdf: PDF_CONTENT_LIMIT_EXCEEDED names the limit that stopped the read', named(byText.warnings, 'maxTextItems') && !named(byText.warnings, 'maxTimeMs') && named(byTime.warnings, 'maxTimeMs') && !named(byTime.warnings, 'maxTextItems'), `${byText.warnings.map(w => w.message.slice(0, 80))} | ${byTime.warnings.map(w => w.message.slice(0, 80))}`);

    // Budget options: Infinity (or past what a timer holds) is no limit, a number may be given as text,
    // and anything else is reported and replaced by the default. Infinity made a timer of 1 ms, and
    // every page after the third came back empty with no warning.
    const nodeWarnings: string[] = [];
    const onNodeWarning = (warning: Error) => nodeWarnings.push(warning.name);
    process.on('warning', onNodeWarning);
    try {
        for (const [label, options] of [['maxTimeMs: Infinity', { maxTimeMs: Infinity }], ['maxTimeMs: 2^31', { maxTimeMs: 2 ** 31 }], ['maxTimeMs: 2^40', { maxTimeMs: 2 ** 40 }],
            ['maxTextItems: "20000" (text)', { maxTextItems: '20000' }], ['maxOperators and maxAnnotations: Infinity', { maxOperators: Infinity, maxAnnotations: Infinity }]] as const) {
            const whole = await read(testPdf, { pdfParserConfig: options });
            check(`pdf: ${label} reads the whole document, with no warning`, !whole.error && whole.pages === 8 && whole.codes.length === 0, `${whole.pages} pages ${whole.codes} ${whole.error}`);
        }
        for (const [key, value] of [['maxTimeMs', NaN], ['maxTimeMs', -1], ['maxTextItems', 'many'], ['maxOperators', null], ['maxAnnotations', NaN], ['processMemoryMb', 16]] as const) {
            const replaced = await read(testPdf, { pdfParserConfig: { [key]: value } });
            check(`pdf: ${key} of ${String(value)} is reported, and the default used`, !replaced.error && replaced.pages === 8 && replaced.codes.join() === 'INVALID_CONFIG_VALUE' && replaced.warnings[0].message.includes(`pdfParserConfig.${key}`), `${replaced.pages} pages ${replaced.codes} ${replaced.error}`);
        }
    } finally {
        process.off('warning', onNodeWarning);
    }
    check('pdf: no budget makes a timer Node cannot hold', !nodeWarnings.includes('TimeoutOverflowWarning') && !nodeWarnings.includes('TimeoutNaNWarning'), nodeWarnings.join());

    // An abort signal is the parse's alone. Left listening after the parse, one fired later and ended the
    // process by then reading another parse's document, which came back with 12 of 400 pages; a signal
    // shared by many parses gathered a listener each. (A memory limit of its own gives the two parses one
    // process, the one the first hands to the second.)
    const book = notes(150, false);
    const controller = new AbortController();
    const first = await parseQuiet(testPdf, 'pdf', { abortSignal: controller.signal, pdfParserConfig: { processMemoryMb: 700 } });
    const { getEventListeners } = require('events') as typeof import('events');
    const listening = getEventListeners(controller.signal, 'abort').length;
    const second = read(book, { pdfParserConfig: { processMemoryMb: 700 } });
    setTimeout(() => controller.abort(), 50);
    const secondRead = await second;
    check('pdf: a parse\'s abort signal, fired after it ended, cannot end the process a later parse is using', !first.error && listening === 0 && !secondRead.error && secondRead.pages === 150 && !secondRead.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${listening} listeners, ${secondRead.pages} pages ${secondRead.codes} ${secondRead.error}`);

    // A parse that ended stops reading: an aborted parse went on asking its ended process until the time
    // budget ran out, then reported the limit to a parse that had already rejected.
    const late: string[] = [];
    const abortedEarly = await parseQuiet(drawnOften, 'pdf', { abortSignal: AbortSignal.timeout(300), pdfParserConfig: { maxTimeMs: 1500 }, onWarning: (w: any) => late.push(w.code) });
    await new Promise(resolve => setTimeout(resolve, 2500));
    const nextRead = await read(testPdf);
    check('pdf: an aborted parse reports nothing after it rejects, and the next parse reads whole', /abort/i.test(abortedEarly.error) && late.length === 0 && !nextRead.error && nextRead.pages === 8, `${abortedEarly.error} ${late} ${nextRead.pages}`);

    // The time budget is the process's CPU time: a host too busy to read the process's replies (here,
    // stalled most of the time) slows a document without spending its budget. Charged the time passed,
    // 115 of 150 ordinary parses at once came back cut short.
    const small = notes(20, false);
    const stall = setInterval(() => { const until = Date.now() + 300; while (Date.now() < until) { /* the host is busy */ } }, 50);
    let stalled: Awaited<ReturnType<typeof read>>;
    try { stalled = await read(small, { pdfParserConfig: { maxTimeMs: 1000 } }); } finally { clearInterval(stall); }
    check('pdf: a host too busy to read the process\'s replies cuts no document short', !stalled.error && stalled.pages === 20 && !stalled.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED'), `${stalled.pages} pages ${stalled.codes} ${stalled.error}`);

    // Processes are pooled, one per CPU: a process per parse started 101 for 100 parses at once.
    const childProcess = require('child_process') as typeof import('child_process');
    const { syncBuiltinESMExports } = require('module') as typeof import('module');
    const spawn = childProcess.spawn;
    let running = 0, most = 0;
    (childProcess as any).spawn = (...args: any[]) => {
        const child = (spawn as any).apply(childProcess, args);
        running++;
        most = Math.max(most, running);
        child.once('exit', () => { running--; });
        return child;
    };
    syncBuiltinESMExports();
    try {
        const cpus = typeof os.availableParallelism === 'function' ? os.availableParallelism() : os.cpus().length;
        const many = await Promise.all(Array.from({ length: cpus * 3 }, () => read(testPdf, { pdfParserConfig: { processMemoryMb: 600 } })));
        check('pdf: parses at once share at most one process per CPU, and each reads whole', most > 0 && most <= cpus && many.every(r => !r.error && r.pages === 8 && !r.codes.includes('PDF_CONTENT_LIMIT_EXCEEDED')), `${most} processes for ${cpus} CPUs, pages ${[...new Set(many.map(r => r.pages))]}`);
    } finally {
        (childProcess as any).spawn = spawn;
        syncBuiltinESMExports();
    }

    // A process reading a document ends with its host: one busy drawing a long string, which it does
    // not look up from, outlived its parent by seconds (a 100 MB string, 8 seconds).
    const script = path.join(os.tmpdir(), `officeparser-pdf-exit-${process.pid}.ts`);
    const pdfFile = path.join(os.tmpdir(), `officeparser-pdf-exit-${process.pid}.pdf`);
    fs.writeFileSync(pdfFile, pdfFrom([
        '<< /Type /Catalog /Pages 2 0 R >>', '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
        flate(Buffer.concat([Buffer.from('BT /F1 10 Tf 0 50 Td ('), Buffer.alloc(20_000_000, 'a'), Buffer.from(') Tj ET\n')])), '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    ]));
    fs.writeFileSync(script, [
        `const childProcess = require('child_process');`,
        `const spawn = childProcess.spawn;`,
        `childProcess.spawn = (...args: any[]) => { const child = spawn.apply(childProcess, args); console.log('PID ' + child.pid); setTimeout(() => process.exit(0), 800); return child; };`,
        `require('module').syncBuiltinESMExports();`,
        `const { OfficeParser } = require(${JSON.stringify(path.join(__dirname, '..', '..', 'src', 'OfficeParser'))});`,
        `OfficeParser.parseOffice(${JSON.stringify(pdfFile)}, { onWarning: () => {}, pdfParserConfig: { maxTimeMs: 60000 } }).catch(() => {});`,
    ].join('\n'));
    let pids: number[] = [];
    try {
        const out = childProcess.execFileSync(process.execPath, [...process.execArgv, script], { encoding: 'utf8', timeout: 60_000 });
        pids = [...out.matchAll(/PID (\d+)/g)].map(m => Number(m[1]));
    } catch { /* checked below */ } finally {
        fs.rmSync(script, { force: true });
        fs.rmSync(pdfFile, { force: true });
    }
    // Ended as its host exits, it is gone at once, not when it next looks up from its work.
    await new Promise(resolve => setTimeout(resolve, 100));
    const survivors = pids.filter(pid => { try { process.kill(pid, 0); return true; } catch { return false; } });
    for (const pid of survivors) { try { process.kill(pid); } catch { /* gone */ } }
    check('pdf: a pdf.js process at work ends when its host exits', pids.length > 0 && survivors.length === 0, `${pids.length} started, ${survivors.length} still running`);

    // pdf.js in this process: operators still stop without costing text, and nested shapes read for
    // fonts only stop once the fonts are known.
    await operatorChecks(false);
    // Where no process can start, pdf.js runs here with a warning; and having run here, it leaves the
    // next parse in a separate process (pdf.js's own worker sets the global a caller's preloaded worker
    // is read from, and every later parse ran here, without its memory or time limit).
    const execPath = process.execPath;
    process.execPath = path.join(os.tmpdir(), 'officeparser-no-such-node');
    let fallback: Awaited<ReturnType<typeof read>>;
    try { fallback = await read(testPdf, { pdfParserConfig: { processMemoryMb: 500 } }); } finally { process.execPath = execPath; }
    check('pdf: where no process can start, pdf.js runs here with PDF_SEPARATE_PROCESS_UNAVAILABLE', !fallback.error && fallback.pages === 8 && fallback.codes.join() === 'PDF_SEPARATE_PROCESS_UNAVAILABLE', `${fallback.pages} pages ${fallback.codes} ${fallback.error}`);
    const begun = Date.now();
    const afterHere = await read(drawnOften, { pdfParserConfig: { maxTimeMs: 200 } });
    check('pdf: after pdf.js ran in this process, the next parse still runs it in a separate one', !afterHere.error && Date.now() - begun < 4000 && named(afterHere.warnings, 'maxTimeMs'), `${Date.now() - begun}ms ${afterHere.codes} ${afterHere.error}`);
}

/**
 * The HTML reader's tree building (implied ends, formatting carried across paragraphs, foster
 * parenting, declarations) and the EPUB writer's paragraph pairing, on hostile shapes: each in time
 * linear in the input, checked by timing two sizes four times apart (linear work takes about four
 * times as long, quadratic sixteen).
 */
async function htmlReadingTests() {
    console.log('- HTML reading and EPUB writing (linear time)...');
    // (Each size's faster of two runs, so a pause in a busy machine is not read as the work growing.)
    const scalesLinearly = async (label: string, run: (size: number) => Promise<unknown>, size: number) => {
        const once = async (n: number) => { const started = performance.now(); await run(n); return performance.now() - started; };
        const time = async (n: number) => Math.min(await once(n), await once(n));
        await once(size);
        const small = await time(size);
        const large = await time(size * 4);
        check(`${label} takes linear time`, large < 5000 && large < Math.max(small, 25) * 10, `${small.toFixed(0)}ms -> ${large.toFixed(0)}ms`);
    };
    const epubOf = (override: string) => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [{ type: 'text', text: 'x' }] }]), 'epub', { onNode: (n: any) => n.type === 'paragraph' ? override : undefined, onWarning: () => {} } as any);

    const html = (src: string) => OfficeParser.parseOffice(Buffer.from(src), { fileType: 'html', onWarning: () => {} } as any).catch(() => undefined);
    const distinctBolds = Array.from({ length: 40 }, (_, k) => `<b data-k="${k}">`).join('');
    for (const [label, unit, prefix, size] of [
        // Formatting carried into each paragraph: at most MAX_CARRIED_FORMATTING at a time, and a budget of them in all.
        ['paragraphs after 40 distinct unclosed formatting elements', '<p>x', `<p>${distinctBolds}`, 5000],
        ['formatting elements each closed by a paragraph', '<p><b data-a="1"><i>x', '', 5000],
        ['misnested formatting end tags', '<b>1<i>2</b>3</i>', '', 5000],
        ['end tags outside their scope', '</b></td></tr></span></li></p>', '<b><table><tr><td><span><li><p>x', 5000],
        ['list items under 200 divisions', '<li>x', '<div>'.repeat(200), 5000],
        ['cells written directly in tables', '<table><td>x</table>', '', 2500],
        ['text and paragraphs between table rows', 'x<p>y</p><tr><td>z</td></tr>', '<table>', 2500],
        ['headings in headings', '<h1>a<h2>b', '', 5000],
        ['declarations and bogus end tags', '<!x><?y?></ z><![if a]>w<![endif]>', '', 5000],
        ['comment ends', '<!-- a -- b --!><!--><!--->', '', 5000],
        ['textareas and titles', '<textarea>a</textarea><title>t</title>', '', 5000],
        ['raw text holding other end tags', '<script>a</style></scripty></script><xmp>b</textarea></xmp>', '', 5000],
        ['stray less-than signs', 'a < b <5 <', '', 10000],
        ['scripts between words', ' <script></script> <b>w</b>', '<p>', 5000],
        ['GitHub footnote references', '<sup><a data-footnote-ref href="#fn-1">1</a></sup>', '', 5000],
    ] as const) {
        const ending = label === 'GitHub footnote references' ? '<section data-footnotes><ol><li id="fn-1"><p>n <a data-footnote-backref href="#r">b</a></p></li></ol></section>' : '';
        await scalesLinearly(`html: ${label}`, n => html(prefix + unit.repeat(n) + ending), size);
    }
    // A page in the encoding it declares is decoded twice at most (once found not UTF-8).
    await scalesLinearly('html: a windows-1252 page', n => OfficeParser.parseOffice(Buffer.from('<meta charset="windows-1252"><p>' + 'caf\xe9 '.repeat(n), 'latin1'), { fileType: 'html', onWarning: () => {} } as any), 20000);

    // MHT: each space after an `=` is looked at once, whether or not it pads a soft line break.
    const { readMht } = await import('../../src/utils/mhtUtils');
    const qp = (body: string) => Buffer.from(`Content-Type: text/html\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n${body}`, 'latin1');
    await scalesLinearly('mht: equals signs each before a run of spaces', async n => readMht(qp(('=' + ' '.repeat(50)).repeat(n))), 5000);
    await scalesLinearly('mht: equals signs with spaces between', async n => readMht(qp('= '.repeat(n * 25))), 5000);

    // Writing: a srcset candidate's trailing commas (a pattern for them retried every comma), ids renamed as XML names.
    await scalesLinearly('html: a srcset candidate of commas', n => OfficeGenerator.generate(astWith([{ type: 'paragraph', children: [{ type: 'image', metadata: { url: 'https://ok/a.png' }, htmlAttributes: { srcset: 'a' + ','.repeat(n) + 'x,' } }] }]), 'html', { onWarning: () => {} } as any), 20000);
    await scalesLinearly('epub: ids each starting with a digit, and links to them', n => OfficeGenerator.generate(astWith(Array.from({ length: n }, (_, k) => ({ type: 'paragraph', metadata: { anchorIds: [`${k}`, `${k} x`] }, children: [{ type: 'text', text: 'x', metadata: { link: `#${k}`, linkType: 'internal' } }] }))), 'epub', { onWarning: () => {} } as any), 1000);

    // EPUB: each </p> is paired with its <p> in one pass, nested or not, and an unclosed comment ends the scan.
    await scalesLinearly('epub: paragraphs holding tables holding paragraphs, repeated,', n => epubOf('<p>a<table><tr><td><p>b</p></td></tr></table>c</p>'.repeat(n)), 5000);
    await scalesLinearly('epub: paragraphs nested 200 deep, repeated,', n => epubOf(('<p>'.repeat(200) + 'x' + '</p>'.repeat(200)).repeat(n)), 100);
    await scalesLinearly('epub: unclosed comments before paragraphs', n => epubOf('<!--<p>'.repeat(n)), 20000);
}

async function main() {
    console.log('Running sanitization security tests...\n');
    unitTests();
    configPollutionTests();
    await htmlTests();
    await htmlAttributeBagTests();
    await htmlSourceAttributesTests();
    await sourceCommentBreakoutTests();
    await iframePreservationTests();
    await mdInlineFormattingTests();
    await markdownTests();
    await csvTests();
    await metadataOverrideTests();
    await styleMapTests();
    await rtfUrlTests();
    await docxSanitizationTests();
    await odtSanitizationTests();
    await latexSanitizationTests();
    await latexParserTests();
    await odfRepeatExpansionTests();
    await abortSignalTests();
    await corruptArchiveTests();
    await truncatedArchiveTests();
    await missingMainPartTests();
    await incompleteArchiveWarningTests();
    await odfTypeResolutionTests();
    await configOwnershipTests();
    errorReportingTests();
    await errorRoutingTests();
    await parserHardeningTests();
    await htmlReadingTests();

    console.log(`\n${failed === 0 ? '✓' : '✗'} Sanitization tests: ${passed} passed, ${failed} failed`);
    if (failed > 0) process.exit(1);
}

main().catch(err => { console.error(err); process.exit(1); });
