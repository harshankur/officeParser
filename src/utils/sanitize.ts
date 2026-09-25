/**
 * Shared output-sanitization helpers.
 *
 * Every string in the parsed AST originates from an untrusted document, so any
 * value interpolated into generated output (HTML, XHTML, CSS, URLs, inline
 * scripts, CSV, RTF, Markdown) must be escaped for its destination context.
 * These are the single source of truth: each generator delegates to them so
 * escaping stays consistent and a gap fixed here is fixed everywhere.
 */

/**
 * Whether a string is a plain HTML attribute *name*, safe to interpolate before `="..."`.
 *
 * Escaping the value is not enough on its own: a key containing a quote or `=` closes the
 * attribute and opens another, so `x" onmouseover="alert(1)" z` yields a real event handler no
 * matter how carefully the value is escaped. This is the shape an attribute-injection payload
 * takes, and rejecting it outright is simpler and safer than trying to escape a name.
 *
 * The predicate is shared rather than restated because it is now applied at four independent
 * points (the parser's attribute collection, the generator's attribute bag, and two
 * styleMap-driven paths). Each of those still keeps its own skip-list inline: the lists are the
 * same policy expressed for different layers, and collapsing them would erase the defence in
 * depth the surrounding comments describe.
 */
export function isSafeHtmlAttributeName(name: string): boolean {
    return typeof name === 'string' && /^[a-zA-Z][a-zA-Z0-9-]*$/.test(name);
}

/**
 * Element names a `styleMap` may map a node onto.
 *
 * An allowlist rather than a pattern: a tag name is interpolated into both `<TAG …>` and
 * `</TAG>`, so it is not enough for it to *look* like a name - `script`, `style`, `iframe` and
 * friends are perfectly well-formed names that would introduce an active context the rest of the
 * generator's escaping assumes does not exist. This is the semantic set a style mapping is for:
 * block containers, headings, and the inline emphasis elements.
 */
const SAFE_STYLE_MAP_TAGS = new Set([
    'p', 'div', 'span', 'section', 'article', 'aside', 'header', 'footer', 'main', 'figure', 'figcaption',
    'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
    'blockquote', 'pre', 'code', 'q', 'cite', 'address',
    'ul', 'ol', 'li', 'dl', 'dt', 'dd',
    'b', 'strong', 'i', 'em', 'u', 'ins', 'del', 's', 'strike', 'mark', 'small', 'sub', 'sup', 'kbd', 'samp', 'var', 'abbr',
]);

/**
 * Whether a `styleMap` `output.tag` may be emitted as an element name.
 *
 * Callers must fall back to their default tag when this returns false, never emit the value.
 */
export function isSafeStyleMapTag(tag: unknown): tag is string {
    return typeof tag === 'string' && SAFE_STYLE_MAP_TAGS.has(tag.toLowerCase());
}

/**
 * Escapes text for an HTML text node or a double-quoted attribute value.
 * Includes the single quote so the result is also safe inside single-quoted
 * attributes.
 */
export function escapeHtml(text: string): string {
    if (typeof text !== 'string') return text as unknown as string;
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * Make raw text safe to wrap in `<!-- ... -->` in HTML or Markdown output. A comment closes at the first
 * `-->` or `--!>`, and a leading `>` or `->` closes an empty one, so AST-derived text containing those
 * could end the comment early and turn what follows into live markup. Each is broken by escaping its
 * `>` (entities are not decoded inside a comment, so this stays inert). Text parsed FROM a comment can
 * never contain `-->`, so real comments still round-trip byte-for-byte; only hand-built or hostile
 * nodes are altered.
 */
export function sanitizeCommentText(text: string): string {
    if (typeof text !== 'string') return '';
    let out = text.replace(/--!?>/g, (m) => `${m.slice(0, -1)}&gt;`);
    if (out.startsWith('>')) out = `&gt;${out.slice(1)}`;
    else if (out.startsWith('->')) out = `-&gt;${out.slice(2)}`;
    return out;
}

/**
 * Escapes text for an XML text node or attribute (XHTML/OPF/NCX). Same as
 * escapeHtml but emits the XML-canonical `&apos;` for the single quote.
 */
export function escapeXml(text: string): string {
    if (typeof text !== 'string') return '';
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Sanitizes a single CSS value (e.g. a color/size/font/alignment pulled from a
 * document) for placement inside a `style="prop: VALUE"` attribute.
 *
 * - Drops the whole value if it contains a resource-fetching or executing
 *   construct (`url()`, `expression()`, `@import`, `image-set()`, `javascript:`)
 *   or angle brackets that could break out of the attribute/tag.
 * - Strips characters that break out of `prop: value` (`;`, quotes), out of a
 *   `<style>` rule (`{}`), CSS escapes (`\`), and control characters.
 *
 * `rgb()/hsl()` and hex/named colors, lengths, and (unquoted) font names all
 * survive; the trade-off is that legitimately quoted font names lose their
 * quotes, which browsers tolerate.
 */
export function sanitizeCssValue(value: string): string {
    if (typeof value !== 'string') return '';
    // Strip every form of intra-token noise FIRST, then test for dangerous constructs.
    // Order matters: a payload like "u\nrl(", "url/*x*/(" or "u\rl(" would survive the test if
    // tested before removal, then reassemble into "url(" once the noise is stripped.
    //
    // The backslash strip belongs here, not after the test. CSS treats `\` as an escape a
    // browser resolves away, so `u\rl(http://evil)` IS `url(http://evil)` to a renderer -
    // stripping it downstream of the test meant the sanitizer handed back a live `url()` it
    // had just declared safe. Every construct in the denylist is reachable this way
    // (`expr\ession(`, `image\-set(`), so the fix is the ordering, not another pattern.
    const cleaned = value
        .replace(/[\x00-\x1F\x7F]/g, '')          // control chars (incl. newlines/tabs)
        .replace(/\/\*[\s\S]*?\*\//g, '')         // CSS comments used to obfuscate
        .replace(/\\/g, '');                      // CSS escapes; see above
    if (/(?:url|expression|image-set|element|-moz-binding)\s*\(|@import|javascript:|[<>]/i.test(cleaned)) {
        return '';
    }
    // Backslash is already gone above; the rest still have work to do here.
    return cleaned.replace(/[;{}"'`]/g, '').trim();
}

/**
 * Escapes a document-supplied URL for use in an href/src attribute. Beyond the
 * usual attribute escaping, this rejects script-executing schemes (javascript:,
 * vbscript:, data:, etc.) so a hyperlink extracted from an untrusted document
 * can't run code when clicked: only http(s)/mailto/tel and relative/fragment
 * URLs are passed through.
 */
export function sanitizeUrl(url: string): string {
    if (typeof url !== 'string') return '';
    const trimmed = url.trim();
    // Browsers ignore control characters when parsing a URL scheme, so strip them
    // first to catch obfuscated payloads like "java\tscript:alert(1)".
    const stripped = trimmed.replace(/[\x00-\x1F\x7F]+/g, '');
    const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(stripped);
    if (schemeMatch && !/^(https?|mailto|tel)$/i.test(schemeMatch[1])) {
        return '';
    }
    // Emit the same normalized string that was validated.
    return escapeHtml(stripped);
}

/**
 * Decide whether a non-provider `<iframe>` should be preserved, per
 * `HtmlParserConfig.preserveIframes`. `true` allows any src; an array is a hostname allowlist,
 * where an entry matches the src's host exactly or as a `.`-suffix (so `"vimeo.com"` also matches
 * `player.vimeo.com`). A relative or unparseable src is allowed only under `true`. This is a
 * preservation gate, not a sanitizer - the src is still scheme-checked with `sanitizeUrl` on
 * generation.
 */
export function iframeAllowed(src: string, preserve: boolean | string[] | undefined): boolean {
    if (preserve === true) return true;
    if (!Array.isArray(preserve) || preserve.length === 0) return false;
    let host: string;
    try {
        host = new URL(src).hostname.toLowerCase();
    } catch {
        return false;
    }
    return preserve.some(entry => {
        const e = String(entry).toLowerCase().trim();
        return e.length > 0 && (host === e || host.endsWith('.' + e));
    });
}

/**
 * Like sanitizeUrl but for an <img>/<source> src: additionally permits
 * `data:image/*` URIs (embedded document images) while still rejecting
 * script-executing schemes and non-image data URIs (e.g. data:text/html).
 */
export function sanitizeImageUrl(url: string): string {
    if (typeof url !== 'string') return '';
    const trimmed = url.trim();
    const stripped = trimmed.replace(/[\x00-\x1F\x7F]+/g, '');
    const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(stripped);
    if (schemeMatch) {
        const scheme = schemeMatch[1].toLowerCase();
        if (scheme === 'data') {
            if (!/^data:image\//i.test(stripped)) return '';
        } else if (scheme !== 'http' && scheme !== 'https') {
            return '';
        }
    }
    return escapeHtml(stripped);
}

/**
 * Serializes data for embedding inside an inline <script> block. JSON.stringify
 * alone doesn't escape "<", so a value containing "</script>" (e.g. a chart
 * label from attacker-controlled document XML) would close the script early and
 * inject markup. Also escapes the U+2028/U+2029 line separators, which are
 * invalid in JS string literals.
 */
export function serializeForInlineScript(data: unknown): string {
    // U+2028/U+2029 (line/paragraph separators) are valid in JSON but break
    // JS string literals; reference them by code point to keep the source ASCII.
    const lineSep = String.fromCharCode(0x2028);
    const paraSep = String.fromCharCode(0x2029);
    return JSON.stringify(data)
        .replace(/</g, '\\u003C')
        .replace(/>/g, '\\u003E')
        .split(lineSep).join('\\u2028')
        .split(paraSep).join('\\u2029');
}

/**
 * Formats a value for a CSV field: guards against spreadsheet formula/DDE
 * injection (CWE-1236) and applies RFC 4180 quoting.
 *
 * A cell beginning with `= + - @` (or a tab/CR that some apps treat as a
 * formula start) is prefixed with a single quote so Excel/Sheets render it as
 * literal text rather than executing it. Genuine numbers (including negatives)
 * are exempt so numeric columns are preserved.
 */
export function csvSafeCell(value: string, delimiter: string): string {
    let v = typeof value === 'string' ? value : String(value ?? '');
    // A plain signed number (e.g. "-8", "+7", "-5.3") can't be a formula, so exempt it;
    // otherwise numeric columns get quoted as text. Anything else starting with a formula
    // trigger (including "+1+1", "-1+cmd", "=", "@") is prefixed with a quote.
    const isNumber = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(v.trim());
    // Test the trimmed value: the numeric exemption above already trims, so testing the raw
    // string here meant a leading space slipped a trigger past the guard (" =1+1" was emitted
    // unprefixed). Most spreadsheet apps treat a leading-space cell as text and would not
    // evaluate it, so this is defence in depth rather than a demonstrated bypass - but the
    // asymmetry between the two tests was an accident, not a decision.
    if (!isNumber && /^[=+\-@\t\r]/.test(v.trim())) {
        v = `'${v}`;
    }
    if (v.includes(delimiter) || v.includes('"') || v.includes('\n') || v.includes('\r')) {
        return `"${v.replace(/"/g, '""')}"`;
    }
    return v;
}

/**
 * Validates and escapes a document-supplied URL for an RTF `HYPERLINK` field argument.
 *
 * Mirrors `sanitizeUrl`'s contract (validate the scheme, then encode for the destination, else
 * return `''`) but cannot reuse it: `sanitizeUrl` returns `escapeHtml(...)`, which would emit
 * `&amp;` into an RTF field. The scheme allowlist is deliberately identical to `sanitizeUrl`'s
 * and `sanitizeMarkdownUrl`'s, so all three text generators agree on what a hyperlink may point at.
 *
 * **Additionally rejects UNC paths (`\\host\share`), which the HTML allowlist does not.** In a
 * browser `\\evil.com\share` is an inert relative path; in Word it is a live UNC reference that
 * triggers an SMB fetch and an NTLM handshake on click, which is a credential-leak vector rather
 * than a rendering quirk. That asymmetry is why this is a separate function and not a flag on
 * `sanitizeUrl` - the HTML helper must NOT gain this behaviour, since there the path is harmless
 * and rejecting it would break legitimate relative links.
 *
 * Returns `''` for a rejected URL; callers emit the link text without the field wrapper, matching
 * how HTML degrades to `href=""` and Markdown to `[text]()`.
 */
export function sanitizeRtfUrl(url: string): string {
    if (typeof url !== 'string') return '';
    const trimmed = url.trim();
    // Control characters are stripped before scheme matching for the same reason as sanitizeUrl:
    // they are ignored when the target application parses the scheme.
    const stripped = trimmed.replace(/[\x00-\x1F\x7F]+/g, '');
    if (/^[\\/]{2}[^\\/]/.test(stripped)) return '';   // UNC (\\host\share, //host\share)
    const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(stripped);
    if (schemeMatch && !/^(https?|mailto|tel)$/i.test(schemeMatch[1])) {
        return '';
    }
    return escapeRtf(stripped);
}

/**
 * Validates a hyperlink URL for an office package (DOCX/ODT), returning the raw (un-escaped)
 * validated string or `''` when rejected. The caller is responsible for `escapeXml`-ing the result
 * at the sink (the relationship `Target` in DOCX, the `xlink:href` attribute in ODT) - unlike
 * {@link sanitizeRtfUrl}, which RTF-escapes because RTF has no separate attribute-quoting layer. The
 * validation policy is identical to RTF's, and for the same reason: an office-document hyperlink is
 * the same click target as an RTF field, so a UNC target (`\\host\share`) is a live SMB/NTLM
 * credential-leak vector in Word and Writer, not the inert relative path a browser sees. Only
 * `https`/`http`/`mailto`/`tel` schemes (plus relative and fragment URLs) are allowed.
 */
export function sanitizeOfficePackageUrl(url: string): string {
    if (typeof url !== 'string') return '';
    const stripped = url.trim().replace(/[\x00-\x1F\x7F]+/g, '');
    if (/^[\\/]{2}[^\\/]/.test(stripped)) return '';   // UNC
    const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(stripped);
    if (schemeMatch && !/^(https?|mailto|tel)$/i.test(schemeMatch[1])) return '';
    return stripped;
}


/**
 * Removes characters that are illegal in XML 1.0 even when escaped, so a single stray control byte
 * from a hostile source document cannot make an OOXML part unparseable (which makes Word refuse the
 * whole file). Strips C0/C1 controls except tab/newline/carriage-return, U+FFFE/U+FFFF, and lone
 * surrogates. Apply before escaping at every text and attribute sink.
 */
export function stripInvalidXmlChars(text: string): string {
    if (typeof text !== 'string') return '';
    return text
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x84\x86-\x9F\uFFFE\uFFFF]/g, '')
        .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')   // high surrogate not followed by low
        .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1'); // low surrogate not preceded by high
}

/**
 * Escapes text for RTF: neutralizes the control/group metacharacters `\ { }`
 * (which would otherwise inject RTF control words or groups), encodes the double
 * quote (so a hyperlink field argument can't be terminated early), and hex/unicode
 * encodes non-ASCII characters.
 */
export function escapeRtf(text: string): string {
    if (typeof text !== 'string') return '';
    return text
        .replace(/\\/g, '\\\\')
        .replace(/{/g, '\\{')
        .replace(/}/g, '\\}')
        .replace(/"/g, "\\'22")
        .replace(/[^\x00-\x7F]/g, (match) => {
            let code = match.charCodeAt(0);
            if (code < 256) {
                return `\\'${code.toString(16).padStart(2, '0')}`;
            }
            if (code > 32767) {
                code -= 65536;
            }
            return `{\\uc0\\u${code}}`;
        });
}

/**
 * Neutralizes a `<` that would open an HTML tag, comment or processing instruction (one
 * immediately followed by a letter, `/`, `!` or `?`, matching how browsers detect tags), so
 * document content can't inject `<script>`/`<img onerror>` when the Markdown is rendered to
 * HTML. Nothing else changes: this is the escaping for a position where nothing decodes
 * entities (math), where encoding more would change the content.
 */
export function markdownEscapeTags(text: string): string {
    if (typeof text !== 'string') return '';
    return text.replace(/<(?=[a-zA-Z/!?])/g, '&lt;');
}

/**
 * Escapes document text for a Markdown text position, one a Markdown parser decodes character
 * references in (MarkdownParser does, as CommonMark does). Besides the tag-opening `<` of
 * {@link markdownEscapeTags}, the `&` of anything that reads as a character reference
 * (`&quot;`, `&#39;`, `&#x27;`, `&copy;`, any `&name;`) becomes `&amp;`, so literal text such
 * as `&quot;` survives a round trip instead of losing one level of escaping each time. A bare
 * `&` (`Tom & Jerry`, `a && b`) and every other Markdown metacharacter are left as they are.
 * The `&` is escaped first, so the `&lt;` this writes is not escaped again. URL schemes are
 * handled by sanitizeMarkdownUrl.
 */
export function markdownEscapeText(text: string): string {
    if (typeof text !== 'string') return '';
    return markdownEscapeTags(text.replace(/&(?=#\d+;|#[xX][0-9a-fA-F]+;|[A-Za-z][A-Za-z0-9]*;)/g, '&amp;'));
}

/**
 * Sanitizes a document-supplied URL for a Markdown `[text](url)` / `![alt](url)`
 * target. Rejects script-executing schemes (returning '' → a dead link) and
 * percent-encodes the characters that would break out of the `(...)` or inject
 * markup. `&` is preserved so query strings survive, except that a Markdown renderer
 * decodes character references in a link target: the `&` of one is written `&amp;`, so the
 * renderer reads exactly this URL (a `javascript&colon;` the scheme check let through as a
 * path cannot become `javascript:`). Set `allowDataImage` for image targets so embedded
 * `data:image/*` URIs are permitted.
 */
export function sanitizeMarkdownUrl(url: string, opts?: { allowDataImage?: boolean }): string {
    if (typeof url !== 'string') return '';
    const stripped = url.trim().replace(/[\x00-\x1F\x7F]+/g, '');
    const schemeMatch = /^([a-z][a-z0-9+.-]*):/i.exec(stripped);
    if (schemeMatch) {
        const scheme = schemeMatch[1].toLowerCase();
        const ok = /^(?:https?|mailto|tel)$/.test(scheme)
            || (opts?.allowDataImage === true && /^data:image\//i.test(stripped));
        if (!ok) return '';
    }
    return stripped.replace(/[\s()<>"`\\]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'))
        .replace(/&(?=#\d+;|#[xX][0-9a-fA-F]+;|[A-Za-z][A-Za-z0-9]*;)/g, '&amp;');
}

// ─── LaTeX ──────────────────────────────────────────────────────────────────────────────────────

/**
 * Every character sequence TeX's input reader treats as the end of a line, plus the separators a
 * source document uses for a soft line break (PPTX's vertical tab, a form feed, the Unicode line and
 * paragraph separators). They are all normalized to `\n` before anything else, so the one character
 * a caller has to reason about is `\n`: a raw line end inside a comment would end the comment and
 * expose the rest of the line as live LaTeX, and a blank line inside a macro argument is a paragraph
 * break that aborts the whole run.
 */
const LATEX_LINE_BREAKS = /\r\n?|[\n\x0B\x0C\u0085\u2028\u2029]/g;

/**
 * Characters with no visible meaning in LaTeX output: the remaining C0/C1 controls, the byte-order
 * mark, and the two noncharacters. pdfTeX rejects several controls outright ("Text line contains an
 * invalid character"), which stops the run.
 */
const LATEX_STRIPPED_CHARS = /[\x00-\x08\x0E-\x1F\x7F-\x84\x86-\x9F\uFEFF\uFFFE\uFFFF]/g;

/**
 * Normalizes untrusted text before it is escaped for LaTeX: Unicode to its composed form, line
 * breaks to `\n`, invisible control characters and lone surrogates removed (a lone surrogate cannot be encoded as UTF-8, so the file
 * would not even be valid input).
 */
function normalizeLatexInput(text: string): string {
    return text
        // Composed form: pdfLaTeX has a glyph for a precomposed letter such as U+00E9 but none for
        // a combining accent, so a decomposed `e` + U+0301 would be a fatal error there.
        .normalize('NFC')
        .replace(LATEX_LINE_BREAKS, '\n')
        .replace(LATEX_STRIPPED_CHARS, '')
        .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/g, '')
        .replace(/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, '$1');
}

/**
 * Replacement for every character that is not literal text to LaTeX.
 *
 * The ten special characters (`\ { } $ & % # _ ~ ^`) are the injection surface: left alone, `\`
 * starts a command, `{`/`}` open and close groups, `%` comments out the rest of the line, `$` enters
 * math, `&`/`#` are alignment and parameter characters. The rest keep the text meaning what it said:
 * `<`, `>` and `|` print as other glyphs in some font encodings, a backtick is an opening quote, and
 * `[`/`]` are braced because a `[` right after a command such as `\item` or `\\` is read as that
 * command's optional argument. Tabs are ordinary spaces in running text.
 */
const LATEX_TEXT_ESCAPES: Record<string, string> = {
    '\\': '\\textbackslash{}',
    '{': '\\{',
    '}': '\\}',
    '$': '\\$',
    '&': '\\&',
    '%': '\\%',
    '#': '\\#',
    '_': '\\_',
    '~': '\\textasciitilde{}',
    '^': '\\textasciicircum{}',
    '<': '\\textless{}',
    '>': '\\textgreater{}',
    '|': '\\textbar{}',
    '`': '\\textasciigrave{}',
    '[': '{[}',
    ']': '{]}',
    '\t': ' ',
    '\u00A0': '~',
    '\u00AD': '\\-',
    '\u200B': '\\hspace{0pt}',
};

/**
 * Characters TeX fonts combine with an identical neighbor into a different glyph (`--` is an en
 * dash, `''` a closing double quote, `,,` a low double quote in T1). An empty group between the pair
 * keeps the two characters the source actually contained.
 */
const LATEX_LIGATURE_CHARS = new Set(['-', "'", ',']);

/**
 * Escapes document text for a LaTeX text position (running text, a macro argument, a table cell).
 *
 * Every string in the AST comes from an untrusted document, and LaTeX is a programming language:
 * unescaped text can run `\input{/etc/passwd}`, `\write18{...}`, or simply break the document's
 * structure. The result contains no active character and no command other than the fixed
 * replacements in {@link LATEX_TEXT_ESCAPES}.
 *
 * @param text - The literal text
 * @param newline - What a line break inside the text becomes. It is never passed through raw: a blank
 *   line is a paragraph break, which is an error inside most macro arguments. Defaults to a space.
 */
export function escapeLatex(text: string, newline = ' '): string {
    if (typeof text !== 'string') return '';
    const src = normalizeLatexInput(text);
    let out = '';
    for (let i = 0; i < src.length; i++) {
        const ch = src[i];
        if (ch === '\n') { out += newline; continue; }
        out += LATEX_TEXT_ESCAPES[ch] ?? ch;
        if (LATEX_LIGATURE_CHARS.has(ch) && src[i + 1] === ch) out += '{}';
    }
    return out;
}

/**
 * Formats text as LaTeX comment lines (each prefixed `% `, the whole terminated by a newline).
 *
 * A comment runs to the end of its line and TeX ignores everything in it, so the only character
 * that matters is the line end: the input is normalized so every line break TeX would honor is a
 * `\n` that gets its own `% ` prefix, and nothing after the comment can leak onto a live line. The
 * trailing newline is part of the result, since the comment has to be closed before any code that
 * follows it.
 */
export function latexComment(text: string): string {
    if (typeof text !== 'string') return '';
    return normalizeLatexInput(text).split('\n').map(line => `% ${line}`.trimEnd()).join('\n') + '\n';
}

/**
 * A source comment (`<!-- ... -->`, `CommentMetadata.sourceSyntax: 'html'`) as LaTeX comment lines:
 * `% <!--body-->`, one `%` line per line of the body. It stays a hidden note that nothing typesets,
 * in the shape the LaTeX parser restores verbatim. Unlike `latexComment`, trailing whitespace inside
 * the body is kept; every line break form still starts a new `%` line, so no text escapes the comment.
 */
export function latexSourceComment(body: string): string {
    const text = `<!--${sanitizeCommentText(typeof body === 'string' ? body : '')}-->`;
    return normalizeLatexInput(text).split('\n').map(line => (line ? `% ${line}` : '%')).join('\n') + '\n';
}

/** URL characters `\href` takes verbatim in every context (RFC 3986 unreserved/sub-delims, minus the ones below). */
const LATEX_URL_LITERAL = /[A-Za-z0-9\-.:/?@!'()*+,;=[\]]/;

/**
 * URL characters that are special to TeX and have an escaped form hyperref turns back into the plain
 * character when it writes the link. They must be escaped rather than left raw because `\href` is
 * often inside another command's argument (bold link text, a table cell), where TeX has already
 * read `#`, `&` and `_` with their special meanings before hyperref could change them.
 */
const LATEX_URL_ESCAPES: Record<string, string> = { '#': '\\#', '&': '\\&', '_': '\\_' };

/**
 * Sanitizes a document-supplied URL for the first argument of `\href`.
 *
 * The scheme policy is the office-package one ({@link sanitizeOfficePackageUrl}): a PDF viewer opens
 * a link target the same way Word does, so a UNC path is a credential-leak vector and only
 * `https`/`http`/`mailto`/`tel` (plus relative and fragment URLs) pass. Returns '' for a rejected
 * URL, and the caller renders the link text alone.
 *
 * The URL is then made inert to TeX: `#`, `&`, `_` and `%` take their escaped forms (an existing
 * `%xx` escape is kept, a stray `%` is encoded as `%25`), and everything else that is not a plain URL
 * character is percent-encoded as UTF-8, which is equivalent in a URL and leaves no TeX-special or
 * non-ASCII character in the argument.
 */
export function sanitizeLatexUrl(url: string): string {
    const safe = sanitizeOfficePackageUrl(url);
    if (!safe) return '';
    const encoder = new TextEncoder();
    let out = '';
    const chars = Array.from(safe);
    for (let i = 0; i < chars.length; i++) {
        const ch = chars[i];
        if (ch === '%') {
            out += /^[0-9A-Fa-f]{2}$/.test((chars[i + 1] ?? '') + (chars[i + 2] ?? '')) ? '\\%' : '\\%25';
        } else if (LATEX_URL_ESCAPES[ch]) {
            out += LATEX_URL_ESCAPES[ch];
        } else if (LATEX_URL_LITERAL.test(ch)) {
            out += ch;
        } else {
            for (const byte of encoder.encode(ch)) out += '%' + byte.toString(16).toUpperCase().padStart(2, '0');
        }
    }
    return out;
}

/**
 * One segment of an image path: starts with a letter, digit or `_` (so never `.`, `..`, a hidden
 * file or an option-like `-`), then letters, digits, `.`, `_`, `-` and single spaces, and does not
 * end with a space. TeX turns a run of spaces into one, so a double space would name another file.
 */
const LATEX_IMAGE_PATH_SEGMENT = '[A-Za-z0-9_](?:[A-Za-z0-9._-]| (?! ))*(?<! )';
const LATEX_IMAGE_PATH = new RegExp(`^${LATEX_IMAGE_PATH_SEGMENT}(?:/${LATEX_IMAGE_PATH_SEGMENT})*$`);

/**
 * An image path from document content, for `\includegraphics`, or null when it is not a plain
 * relative path. TeX reads the named file when the document is compiled, so a document may only
 * point inside its own folder: no scheme, no absolute, home or drive path, no `.`/`..` segments or
 * hidden files, and only letters, digits, `.`, `_`, `-`, `/` and single spaces, none of which mean
 * anything to TeX or to a shell (LaTeX reads file names with spaces since 2019). Such a path needs
 * no escaping.
 *
 * A path from HTML or Markdown is a URL reference, so its percent-escapes (`pic%20one.png`) are
 * decoded first; the decoded path is what is checked. A LaTeX path cannot contain a raw `%` (it
 * starts a comment), so decoding never misreads one.
 */
export function sanitizeLatexImagePath(path: string): string | null {
    let p = String(path ?? '').trim();
    if (p.includes('%')) {
        try { p = decodeURIComponent(p).trim(); } catch { return null; }
    }
    if (!p || p.length > 255) return null;
    return LATEX_IMAGE_PATH.test(p) ? p : null;
}

/**
 * LaTeX commands a math expression may not use, by exact name.
 *
 * Math is the one place the LaTeX generator emits document content as live LaTeX rather than
 * escaped text: `node.text` of a math node *is* LaTeX source (from `$...$` in Markdown, `data-math`
 * in HTML, or converted equations). A document author controls that source, so everything that
 * reaches outside the formula is refused: reading files (`\input`, `\openin`, `\includegraphics`,
 * which TeX Live permits for any readable path by default), writing files or running programs
 * (`\write`, `\immediate`, `\openout`, `\directlua`), changing how later input is read (`\catcode`,
 * `\scantokens`, `\verb`), redefining commands for the rest of the document (`\def`, `\let`,
 * `\newcommand`), changing global layout or counters, ending the run (`\stop`, `\endinput`), and
 * emitting links that bypass URL sanitization (`\href`, `\url`). `\cr`, `\par` and friends would
 * end the table row or paragraph the formula sits in.
 */
const LATEX_MATH_BLOCKED_COMMANDS = new Set([
    'input', 'include', 'includeonly', 'InputIfFileExists', 'IfFileExists', 'endinput',
    'openin', 'openout', 'closein', 'closeout', 'read', 'readline', 'write', 'immediate',
    'special', 'directlua', 'latelua', 'luaexec', 'luadirect', 'ShellEscape', 'shellescape', 'mdfivesum',
    'catcode', 'lccode', 'uccode', 'sfcode', 'mathcode', 'delcode',
    'def', 'edef', 'gdef', 'xdef', 'let', 'futurelet', 'global', 'long', 'outer', 'protected',
    'csname', 'endcsname', 'scantokens', 'scantextokens', 'ExplSyntaxOn', 'ExplSyntaxOff',
    'usepackage', 'RequirePackage', 'documentclass', 'LoadClass', 'makeatletter', 'makeatother',
    'verb', 'verbatiminput', 'lstinputlisting', 'inputminted', 'includegraphics', 'includepdf',
    'href', 'url', 'nolinkurl', 'hyperref', 'hyperlink', 'hypertarget', 'hyperimage', 'hyperbaseurl', 'hypersetup',
    'font', 'fontspec', 'setmainfont', 'setsansfont', 'setmonofont', 'setmathfont', 'addfontfeatures',
    'chardef', 'mathchardef', 'countdef', 'dimendef', 'skipdef', 'muskipdef', 'toksdef',
    'stop', 'bye', 'dump', 'batchmode', 'nonstopmode', 'scrollmode', 'errorstopmode', 'errmessage', 'errhelp',
    'output', 'shipout', 'afterassignment', 'aftergroup',
    'cr', 'crcr', 'tabularnewline', 'noalign', 'omit', 'span', 'par',
    'setcounter', 'addtocounter', 'stepcounter', 'refstepcounter', 'setlength', 'addtolength',
    'settowidth', 'settoheight', 'settodepth', 'pagestyle', 'thispagestyle', 'geometry',
]);

/**
 * Command-name prefixes a math expression may not use: engine primitive families (`\pdffiledump`,
 * `\luatexversion`, `\XeTeXinputencoding`, `\filemoddate`), the `\every...` token-list hooks, every
 * command/environment definition family (`\newcommand`, `\NewDocumentCommand`, `\providecommand`,
 * `\DeclareRobustCommand`, ...), `\show...` (which pauses an interactive run), and page-header
 * commands.
 */
const LATEX_MATH_BLOCKED_PREFIXES = ['pdf', 'lua', 'XeTeX', 'file', 'every', 'new', 'New', 'renew', 'Renew', 'provide', 'Provide', 'Declare', 'show', 'fancy'];

/** Environments a formula may open anywhere inside math (amsmath and the kernel). */
const LATEX_INNER_MATH_ENVIRONMENTS = new Set([
    'matrix', 'pmatrix', 'bmatrix', 'Bmatrix', 'vmatrix', 'Vmatrix', 'smallmatrix',
    'cases', 'aligned', 'alignedat', 'gathered', 'split', 'array', 'subarray',
]);

/**
 * Display environments, which cannot sit inside `\[...\]`. A block formula that is exactly one of
 * them (the common `$$\begin{align}...\end{align}$$` in Markdown) is written bare instead.
 */
const LATEX_DISPLAY_MATH_ENVIRONMENTS = new Set([
    'equation', 'equation*', 'align', 'align*', 'alignat', 'alignat*', 'gather', 'gather*',
    'multline', 'multline*', 'flalign', 'flalign*', 'eqnarray', 'eqnarray*',
]);

/** Whether a control-word name is refused inside math. */
function isBlockedLatexMathCommand(name: string): boolean {
    return LATEX_MATH_BLOCKED_COMMANDS.has(name) || LATEX_MATH_BLOCKED_PREFIXES.some(p => name.startsWith(p));
}

/**
 * The outcome of {@link sanitizeLatexMath}. On success `latex` is ready to place between the math
 * delimiters (or, when `displayEnvironment` is set, to write bare as a display environment). On
 * failure `commands` names the refused commands; it is empty when the problem is structural.
 */
export type LatexMathResult =
    | { ok: true; latex: string; displayEnvironment: boolean }
    | { ok: false; commands: string[] };

/**
 * Makes a LaTeX math expression safe to emit as live math.
 *
 * The expression is scanned the way TeX tokenizes it (control words are a backslash plus ASCII
 * letters; that is exact because every command that could change the letter set is refused), and
 * it is accepted only if all of the following hold:
 *
 * - no refused command ({@link LATEX_MATH_BLOCKED_COMMANDS}, {@link LATEX_MATH_BLOCKED_PREFIXES}),
 *   and no `^^` notation, which spells any character (including `\`) by its hex code and so would
 *   let a command past a textual scan;
 * - braces balance, and every `\begin{env}` closes with a matching `\end{env}` at the same brace
 *   depth, using only math environments, so the formula cannot close a group or environment it
 *   did not open;
 * - no `\(`, `\)`, `\[`, `\]`, and no trailing lone backslash (which would escape the closing `$`).
 *
 * Characters that would reach outside the formula are neutralized rather than refused: `%` (a
 * comment would swallow the closing delimiter) becomes `\%`, `$` becomes `\$`, `#` becomes `\#`,
 * a top-level `&` in inline math becomes `\&` (it would split a table cell), and blank lines are
 * collapsed (a paragraph break is an error in math). In inline math a top-level `\\` becomes a
 * space; in block math a top-level `&` or `\\` has the body wrapped in `aligned`/`gathered`.
 *
 * @param source - The formula as found in the AST (no delimiters)
 * @param mode - `'inline'` for `$...$`, `'block'` for display math
 */
export function sanitizeLatexMath(source: string, mode: 'inline' | 'block'): LatexMathResult {
    const text = normalizeLatexInput(typeof source === 'string' ? source : '');
    const blocked = new Set<string>();
    if (text.includes('^^')) blocked.add('^^');
    let structural = false;
    let out = '';
    let braceDepth = 0;
    const envStack: { name: string; braceDepth: number }[] = [];
    let displayEnvironment = false;
    let displayEnvironmentEnd = -1;
    let topLevelAmpersand = false;
    let topLevelRowBreak = false;

    let i = 0;
    while (i < text.length) {
        const ch = text[i];
        if (ch === '\\') {
            const next = text[i + 1];
            if (next === undefined) { structural = true; break; }
            if (/[A-Za-z]/.test(next)) {
                let j = i + 1;
                while (j < text.length && /[A-Za-z]/.test(text[j])) j++;
                const name = text.slice(i + 1, j);
                if (isBlockedLatexMathCommand(name)) blocked.add('\\' + name);
                if (name === 'begin' || name === 'end') {
                    let k = j;
                    while (k < text.length && /[ \t\n]/.test(text[k])) k++;
                    const env = /^\{([A-Za-z]+\*?)\}/.exec(text.slice(k));
                    if (!env) { structural = true; out += text.slice(i, j); i = j; continue; }
                    const envName = env[1];
                    if (name === 'begin') {
                        const opensDisplay = LATEX_DISPLAY_MATH_ENVIRONMENTS.has(envName) && mode === 'block'
                            && envStack.length === 0 && braceDepth === 0 && out.trim() === '' && !displayEnvironment;
                        if (opensDisplay) displayEnvironment = true;
                        else if (!LATEX_INNER_MATH_ENVIRONMENTS.has(envName)) structural = true;
                        envStack.push({ name: envName, braceDepth });
                    } else {
                        const open = envStack.pop();
                        if (!open || open.name !== envName || open.braceDepth !== braceDepth) structural = true;
                    }
                    out += text.slice(i, k) + env[0];
                    i = k + env[0].length;
                    if (name === 'end' && displayEnvironment && envStack.length === 0 && displayEnvironmentEnd < 0) displayEnvironmentEnd = out.length;
                    continue;
                }
                out += text.slice(i, j);
                i = j;
                continue;
            }
            if (next === '(' || next === ')' || next === '[' || next === ']') structural = true;
            if (next === '\\' && envStack.length === 0) {
                if (mode === 'inline') { out += ' '; i += 2; continue; }
                topLevelRowBreak = true;
            }
            out += '\\' + next;
            i += 2;
            continue;
        }
        switch (ch) {
            case '%': out += '\\%'; break;
            case '$': out += '\\$'; break;
            case '#': out += '\\#'; break;
            case '&':
                if (envStack.length > 0) out += '&';
                else if (mode === 'inline') out += '\\&';
                else { topLevelAmpersand = true; out += '&'; }
                break;
            case '{': braceDepth++; out += ch; break;
            case '}':
                braceDepth--;
                if (braceDepth < 0) structural = true;
                out += ch;
                break;
            case '\n': {
                // Collapse the whole whitespace run; a blank line would be a paragraph break.
                let k = i + 1;
                while (k < text.length && /[ \t\n]/.test(text[k])) k++;
                out += mode === 'inline' ? ' ' : '\n';
                i = k;
                continue;
            }
            default: out += ch;
        }
        i++;
    }

    if (braceDepth !== 0 || envStack.length > 0) structural = true;
    // A display environment must be the whole formula: `\begin{align}...\end{align} x` cannot be
    // written bare, and cannot go inside `\[...\]` either.
    if (displayEnvironment && (displayEnvironmentEnd < 0 || out.slice(displayEnvironmentEnd).trim() !== '')) structural = true;

    if (blocked.size > 0) return { ok: false, commands: [...blocked].sort() };
    if (structural) return { ok: false, commands: [] };

    let latex = out.trim();
    if (mode === 'block' && !displayEnvironment && (topLevelAmpersand || topLevelRowBreak)) {
        const env = topLevelAmpersand ? 'aligned' : 'gathered';
        latex = `\\begin{${env}}\n${latex}\n\\end{${env}}`;
    }
    return { ok: true, latex, displayEnvironment };
}
