/**
 * Shared output-sanitization helpers.
 *
 * Every string in the parsed AST originates from an untrusted document, so any
 * value interpolated into generated output (HTML, XHTML, CSS, URLs, inline
 * scripts, CSV, RTF, Markdown) must be escaped for its destination context.
 * These are the single source of truth — each generator delegates to them so
 * escaping stays consistent and a gap fixed here is fixed everywhere.
 */

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
 * can't run code when clicked — only http(s)/mailto/tel and relative/fragment
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
    // A plain signed number (e.g. "-8", "+7", "-5.3") can't be a formula, so exempt it —
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
 * Escapes document text for a Markdown text position. Markdown passes raw HTML
 * through to the renderer, so a `<` that begins an HTML tag or comment must be
 * neutralized to prevent `<script>`/`<img onerror>` injection when the Markdown
 * is later rendered to HTML.
 *
 * Deliberately narrow — only a `<` immediately followed by a letter, `/`, `!` or
 * `?` (i.e. one that actually opens a tag/comment/PI, matching how browsers
 * detect tags) is encoded. A bare `<` (e.g. `a < b`), `>`, `&`, `[]` and other
 * Markdown metacharacters are left untouched: they can't start a tag, and
 * MarkdownParser round-trips this output without decoding entities, so encoding
 * them would corrupt re-parsed content. URL schemes are handled by
 * sanitizeMarkdownUrl.
 */
export function markdownEscapeText(text: string): string {
    if (typeof text !== 'string') return '';
    return text.replace(/<(?=[a-zA-Z/!?])/g, '&lt;');
}

/**
 * Sanitizes a document-supplied URL for a Markdown `[text](url)` / `![alt](url)`
 * target. Rejects script-executing schemes (returning '' → a dead link) and
 * percent-encodes the characters that would break out of the `(...)` or inject
 * markup. `&` is preserved so query strings survive; set `allowDataImage` for
 * image targets so embedded `data:image/*` URIs are permitted.
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
    return stripped.replace(/[\s()<>"`\\]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
}
