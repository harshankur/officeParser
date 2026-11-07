/**
 * Markdown disclosures (`<details>`) checked against a reference written independently of the Markdown
 * reader: a small CommonMark block parser (with GFM pipe tables) that renders HTML, then a small HTML
 * tree builder (the WHATWG "in body" rules that matter here) over that HTML, which is what GitHub shows.
 * Quadratic in places, which is fine for the short texts the generator draws. Spaces only (the generator
 * writes no tabs); emphasis, links and entities are not rendered, as they move no word in or out of a
 * disclosure.
 *
 * The reading it encodes (the Markdown reader's groupDisclosures states the same):
 * - CommonMark decides blocks and containers. A line starting with a disclosure's or summary's tag
 *   starts an HTML block (condition 6): it interrupts a paragraph and is no lazy line, so it ends a quote
 *   or list item it is not marked or indented into. (The reader, as 8.6.0, reads a line starting with
 *   `<summary>` after a quote line, and a line after a quote's `<summary>` line, as the quote's lazy text:
 *   the generator keeps those out.)
 * - Tags are tags in an HTML block as a browser reads them (a quoted value may hold `>`, a tag left
 *   unfinished holds the rest), and inline in a paragraph, heading, list item or quote line by
 *   CommonMark's open-tag grammar, not in a code span or escaped; never in code or a pipe table cell.
 * - An end tag closes the nearest open element of its kind in scope and everything opened after it;
 *   a quote, list item, heading or cell closes at its end what opened in it.
 * - A disclosure's summary is a `<summary>` first in it, to its `</summary>` wherever that falls.
 * - A disclosure in a list item, heading or table cell, or in a summary, is that block's text (the AST
 *   holds none there).
 *
 * Cases outside what the reader reads the same are counted as out of domain and skipped: a summary not
 * first in its disclosure, a summary closed by something else in another block, and a disclosure in a
 * list item closed in a later block (the reader, as 8.6.0, reads an item's later blocks after its list).
 * Open findings, kept out the same way until the reader has them: a tag escaped or in a code span after
 * a disclosure's tag in a paragraph or heading (the text after the tag is read as blocks of its own, an
 * HTML block where a tag starts it), two summaries side by side in a block (the HTML reader runs their
 * text together), and a disclosure in a `<summary>` outside any disclosure, closed by its `</summary>`.
 * The generator keeps out the block readings the reader has from 8.6.0 that are not CommonMark's (see
 * disclosureMarkdownGenerator).
 */
import type { OfficeContentNode, OfficeParserAST } from '../../src/types';

// ---------- CommonMark blocks ----------
type Kind = 'doc' | 'quote' | 'list' | 'item' | 'para' | 'heading' | 'code' | 'html' | 'table' | 'hr';
interface Blk {
    t: Kind; open: boolean; children: Blk[]; lines: string[];
    level?: number; ordered?: boolean; marker?: string; contentCol?: number; fence?: { ch: string; len: number; indent: number };
    htmlCond?: number; indented?: boolean;
}
const mk = (t: Kind, extra: Partial<Blk> = {}): Blk => ({ t, open: true, children: [], lines: [], ...extra });
const BLOCK_TAGS = 'address|article|aside|base|basefont|blockquote|body|caption|center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|frameset|h[1-6]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|nav|noframes|ol|optgroup|option|p|param|search|section|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul';
const COND6 = new RegExp(`^<\\/?(?:${BLOCK_TAGS})(?:\\s|\\/?>|$)`, 'i');
const COND1 = /^<(?:pre|script|style|textarea)(?:\s|>|$)/i;
const ATTR = String.raw`(?:\s+[A-Za-z_:][A-Za-z0-9_.:-]*(?:\s*=\s*(?:[^\s"'=<>\x60]+|'[^']*'|"[^"]*"))?)`;
const OPEN_TAG = new RegExp(String.raw`^<[A-Za-z][A-Za-z0-9-]*${ATTR}*\s*\/?>`);
const CLOSE_TAG = /^<\/[A-Za-z][A-Za-z0-9-]*\s*>/;
const COMMENT = /^<!--[\s\S]*?-->/;
const blank = (s: string): boolean => /^\s*$/.test(s);
const ATX = /^#{1,6}(?: |$)/;
const HR = /^(?:(?:\* *){3,}|(?:- *){3,}|(?:_ *){3,})$/;
const FENCE = /^(`{3,}|~{3,})([^`]*)$/;
const ITEM = /^([-*+]|\d{1,9}[.)])( +|$)/;
const DELIM = /^\|?(?: *:?-+:? *\|)*(?: *:?-+:? *)\|?$/;

function htmlStart(rest: string, interruptsPara: boolean): number {
    if (COND1.test(rest)) return 1;
    if (rest.startsWith('<!--')) return 2;
    if (COND6.test(rest)) return 6;
    if (!interruptsPara) {
        const m = OPEN_TAG.exec(rest) ?? CLOSE_TAG.exec(rest);
        if (m && blank(rest.slice(m[0].length)) && !/^<\/?(?:pre|script|style|textarea)\b/i.test(rest)) return 7;
    }
    return 0;
}
const htmlEnds = (cond: number, line: string): boolean => cond === 1 ? /<\/(?:pre|script|style|textarea)>/i.test(line) : cond === 2 ? line.includes('-->') : false;
const cells = (row: string): string[] => {
    let r = row.trim();
    if (r.startsWith('|')) r = r.slice(1);
    if (r.endsWith('|') && !r.endsWith('\\|')) r = r.slice(0, -1);
    const out: string[] = [];
    let cur = '';
    for (let i = 0; i < r.length; i++) {
        if (r[i] === '\\' && r[i + 1] === '|') { cur += '|'; i++; } else if (r[i] === '|') { out.push(cur); cur = ''; } else cur += r[i];
    }
    out.push(cur);
    return out.map(c => c.trim());
};

function parseBlocks(src: string): Blk {
    const doc = mk('doc');
    let tip: Blk = doc;
    const parent = new Map<Blk, Blk>();
    const canContain = (p: Blk, t: Kind): boolean => p.t === 'doc' || p.t === 'quote' || p.t === 'item' ? t !== 'item' : p.t === 'list' ? t === 'item' : false;
    const finalize = (b: Blk): void => {
        b.open = false;
        if (b.t === 'para' && b.lines.length >= 2 && b.lines[0].includes('|') && DELIM.test(b.lines[1].trim()) && cells(b.lines[0]).length === cells(b.lines[1]).length) b.t = 'table';
        tip = parent.get(b) ?? doc;
    };
    const addChild = (t: Kind, extra: Partial<Blk> = {}): Blk => {
        while (!canContain(tip, t)) finalize(tip);
        const c = mk(t, extra);
        tip.children.push(c);
        parent.set(c, tip);
        tip = c;
        return c;
    };
    const lastOpenChild = (b: Blk): Blk | undefined => { const c = b.children[b.children.length - 1]; return c && c.open ? c : undefined; };
    for (const raw of src.split('\n')) {
        const ln = raw.replace(/\t/g, '    ');
        let off = 0;
        const nonspace = (): number => { let k = off; while (ln[k] === ' ') k++; return k; };
        const isBlank = (): boolean => blank(ln.slice(off));
        const oldtip = tip;
        let container: Blk = doc;
        // 1. continue the open blocks
        let fenceClosed = false;
        for (let child = lastOpenChild(container); child; child = lastOpenChild(container)) {
            container = child;
            const ns = nonspace();
            const indent = ns - off;
            let unmatched = false;
            switch (child.t) {
                case 'quote': if (indent <= 3 && ln[ns] === '>') { off = ns + 1; if (ln[off] === ' ') off++; } else unmatched = true; break;
                case 'list': break;
                case 'item':
                    if (isBlank()) { if (child.children.length === 0) unmatched = true; else off = ns; }
                    else if (indent >= child.contentCol!) off += child.contentCol!;
                    else unmatched = true;
                    break;
                case 'heading': case 'hr': unmatched = true; break;
                case 'code':
                    if (child.fence) {
                        const f = /^(`{3,}|~{3,}) *$/.exec(ln.slice(ns));
                        if (indent <= 3 && f && f[1][0] === child.fence.ch && f[1].length >= child.fence.len) { finalize(child); fenceClosed = true; break; }
                        let k = child.fence.indent;
                        while (k > 0 && ln[off] === ' ') { off++; k--; }
                    } else if (indent >= 4) off += 4; else if (isBlank()) off = ns; else unmatched = true;
                    break;
                case 'html': if (isBlank() && (child.htmlCond === 6 || child.htmlCond === 7)) unmatched = true; break;
                case 'para': case 'table': if (isBlank()) unmatched = true; break;
            }
            if (fenceClosed) break;
            if (unmatched) { container = parent.get(child)!; break; }
        }
        if (fenceClosed) continue;
        let allClosed = container === oldtip;
        const lastMatched = container;
        const closeUnmatched = (): void => {
            if (allClosed) return;
            let o: Blk = oldtip;
            while (o !== lastMatched) { const p = parent.get(o)!; finalize(o); o = p; }
            allClosed = true;
        };
        let matchedLeaf = container.t === 'code' || container.t === 'html';
        // 2. block starts
        while (!matchedLeaf) {
            const ns = nonspace();
            const indent = ns - off;
            const r = ln.slice(ns);
            if (indent < 4 && r.startsWith('>')) { off = ns + 1; if (ln[off] === ' ') off++; closeUnmatched(); container = addChild('quote'); continue; }
            if (indent < 4 && ATX.test(r)) {
                off = ln.length;
                closeUnmatched();
                addChild('heading', { level: /^#+/.exec(r)![0].length, lines: [r.replace(/^#+ */, '').replace(/(?:^| +)#+ *$/, '')] });
                finalize(tip);
                container = tip;
                matchedLeaf = true;
                break;
            }
            const f = indent < 4 ? FENCE.exec(r) : null;
            if (f && !(f[1][0] === '`' && f[2].includes('`'))) { off = ln.length; closeUnmatched(); container = addChild('code', { fence: { ch: f[1][0], len: f[1].length, indent } }); matchedLeaf = true; break; }
            const lazyPara = !allClosed && !isBlank() && tip.t === 'para';
            const hc = indent < 4 && r.startsWith('<') ? htmlStart(r, container.t === 'para' || lazyPara) : 0;
            if (hc) { closeUnmatched(); container = addChild('html', { htmlCond: hc }); matchedLeaf = true; break; }
            if (indent < 4 && HR.test(r.trimEnd())) { off = ln.length; closeUnmatched(); addChild('hr'); finalize(tip); container = tip; matchedLeaf = true; break; }
            const m = (indent < 4 || container.t === 'list') ? ITEM.exec(r) : null;
            if (m) {
                const ordered = /\d/.test(m[1]);
                const emptyItem = blank(r.slice(m[1].length));
                if (container.t === 'para' && (emptyItem || (ordered && parseInt(m[1], 10) !== 1))) { off = ns; break; }
                const sp = emptyItem ? 1 : (m[2].length > 4 ? 1 : m[2].length);
                const contentCol = indent + m[1].length + sp;
                off = Math.min(ln.length, ns + m[1].length + sp);
                closeUnmatched();
                const type = ordered ? m[1].slice(-1) : m[1];
                if (tip.t !== 'list' || tip.marker !== type) addChild('list', { ordered, marker: type });
                container = addChild('item', { contentCol });
                continue;
            }
            if (indent >= 4 && tip.t !== 'para' && !isBlank()) { off += 4; closeUnmatched(); container = addChild('code', { indented: true }); matchedLeaf = true; break; }
            off = ns;
            break;
        }
        // 3. the rest of the line
        const rest = ln.slice(off);
        if (!allClosed && !isBlank() && tip.t === 'para') { tip.lines.push(rest); continue; }
        closeUnmatched();
        const t = tip.t;
        if (t === 'code' || t === 'html') {
            tip.lines.push(rest);
            if (t === 'html' && htmlEnds(tip.htmlCond!, rest)) finalize(tip);
        } else if (t === 'para' || t === 'table') { if (!isBlank()) tip.lines.push(rest); }
        else if (t !== 'heading' && !isBlank()) addChild('para', { lines: [rest] });
    }
    while (tip !== doc) finalize(tip);
    finalize(doc);
    return doc;
}

// ---------- inline ----------
const esc = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
function inline(text: string): string {
    let out = '';
    for (let i = 0; i < text.length;) {
        const ch = text[i];
        if (ch === '\\' && i + 1 < text.length && /[!-/:-@[-`{-~]/.test(text[i + 1])) { out += esc(text[i + 1]); i += 2; continue; }
        if (ch === '`') {
            let n = 0;
            while (text[i + n] === '`') n++;
            // The next run of exactly n backticks.
            let j = i + n;
            let found = -1;
            while (j < text.length) {
                if (text[j] === '`') { let m = 0; while (text[j + m] === '`') m++; if (m === n) { found = j; break; } j += m; } else j++;
            }
            if (found === -1) { out += '`'.repeat(n); i += n; continue; }
            out += '<code>' + esc(text.slice(i + n, found)) + '</code>';
            i = found + n;
            continue;
        }
        if (ch === '<') {
            const s = text.slice(i);
            const m = COMMENT.exec(s) ?? OPEN_TAG.exec(s) ?? CLOSE_TAG.exec(s);
            if (m) { out += m[0]; i += m[0].length; continue; }
            out += '&lt;';
            i++;
            continue;
        }
        out += esc(ch);
        i++;
    }
    return out;
}

// ---------- render ----------
/** Set by markdownToHtml: a reading the Markdown reader does not have yet (see referenceDisclosureTree). */
let openReading: string | undefined;
function markdownToHtml(src: string): string {
    let blockId = 0;
    openReading = undefined;
    // A paragraph's or heading's text after a disclosure's or summary's tag in it the reader reads as blocks
    // of their own, where a line a tag starts is an HTML block: a tag escaped or in a code span after the
    // first is a tag there. An open finding, kept out until the reader reads that text inline.
    const inlineOf = (text: string): string => {
        return inline(text);
    };
    const render = (b: Blk): string => {
        const mark = (): string => `<!--B${++blockId}-->`;
        switch (b.t) {
            case 'doc': return b.children.map(render).join('\n');
            case 'quote': return '<blockquote>\n' + b.children.map(render).join('\n') + '\n</blockquote>';
            case 'list': return (b.ordered ? '<ol>' : '<ul>') + '\n' + b.children.map(render).join('\n') + '\n' + (b.ordered ? '</ol>' : '</ul>');
            case 'item': return '<li>' + b.children.map(render).join('\n') + '</li>';
            case 'para': return mark() + '<p>' + inlineOf(b.lines.map(l => l.trimStart()).join('\n').trimEnd()) + '</p>';
            case 'heading': return mark() + `<h${b.level}>` + inlineOf(b.lines[0].trim()) + `</h${b.level}>`;
            case 'code': {
                let ls = b.fence ? b.lines.slice(1) : b.lines;
                while (ls.length && blank(ls[ls.length - 1])) ls = ls.slice(0, -1);
                return mark() + '<pre><code>' + esc(ls.join('\n')) + '</code></pre>';
            }
            case 'html': {
                // The Markdown writer's one-line alignment wrapper: its content is Markdown to the reader (a
                // deviation from CommonMark, which has it as HTML, kept from 8.6.0 for the writer's own output).
                const wrapper = /^<div\s+(?:style="text-align:\s*(?:left|center|right|justify);?"|align="(?:left|center|right|justify)")>([\s\S]*)<\/div>$/i.exec(b.lines.join('\n').trim());
                if (wrapper) return mark() + '<div>' + inline(wrapper[1].trim()) + '</div>';
                return mark() + b.lines.join('\n');
            }
            case 'hr': return mark() + '<hr />';
            case 'table': {
                const head = cells(b.lines[0]);
                const n = head.length;
                const rows = b.lines.slice(2).map(cells).map(r => { const x = r.slice(0, n); while (x.length < n) x.push(''); return x; });
                return mark() + '<table><thead><tr>' + head.map(c => `<th>${inline(c)}</th>`).join('') + '</tr></thead>'
                    + (rows.length ? '<tbody>' + rows.map(r => '<tr>' + r.map(c => `<td>${inline(c)}</td>`).join('') + '</tr>').join('') + '</tbody>' : '') + '</table>';
            }
        }
    };
    return render(parseBlocks(src));
}

// ---------- HTML tree builder (the part of WHATWG "in body" that matters here) ----------
interface El { name: string; children: (El | string)[]; openBlock: number; closedBy?: string; closeBlock?: number }
type Tok = { k: 'start' | 'end'; name: string } | { k: 'text'; text: string } | { k: 'block'; id: number };
function tokenize(html: string): Tok[] {
    const toks: Tok[] = [];
    let i = 0;
    const decode = (s: string): string => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');
    while (i < html.length) {
        const lt = html.indexOf('<', i);
        if (lt === -1) { toks.push({ k: 'text', text: decode(html.slice(i)) }); break; }
        if (lt > i) toks.push({ k: 'text', text: decode(html.slice(i, lt)) });
        i = lt;
        if (html.startsWith('<!--', i)) {
            const e = html.indexOf('-->', i + 4);
            const b = /^B(\d+)$/.exec(html.slice(i + 4, e === -1 ? html.length : e));
            if (b) toks.push({ k: 'block', id: +b[1] });
            i = e === -1 ? html.length : e + 3;
            continue;
        }
        const m = /^<(\/?)([A-Za-z][A-Za-z0-9-]*)/.exec(html.slice(i, i + 64));
        if (!m) { toks.push({ k: 'text', text: '<' }); i++; continue; }
        // The tag's end, as the tokenizer finds it: a quoted attribute value may hold '>'.
        let j = i + m[0].length;
        let ended = false;
        while (j < html.length) {
            const c = html[j];
            if (c === '>') { ended = true; j++; break; }
            if (c === '=') {
                j++;
                while (/\s/.test(html[j] ?? '')) j++;
                const q = html[j];
                if (q === '"' || q === "'") { const e = html.indexOf(q, j + 1); if (e === -1) { j = html.length; break; } j = e + 1; continue; }
                continue;
            }
            j++;
        }
        // The end of the text in a tag: the rest is the unfinished tag.
        if (!ended) break;
        toks.push({ k: m[1] ? 'end' : 'start', name: m[2].toLowerCase() });
        i = j;
    }
    return toks;
}
const SPECIAL = new Set('address applet area article aside base basefont bgsound blockquote body br button caption center col colgroup dd details dir div dl dt embed fieldset figcaption figure footer form frame frameset h1 h2 h3 h4 h5 h6 head header hgroup hr html iframe img input keygen li link listing main marquee menu meta nav noembed noframes noscript object ol p param plaintext pre script search section select source style summary table tbody td template textarea tfoot th thead title tr track ul wbr xmp'.split(' '));
const SCOPE = new Set('applet caption html table td th marquee object template'.split(' '));
const BLOCKISH_START = new Set('address article aside blockquote center details dialog dir div dl fieldset figcaption figure footer header hgroup main menu nav ol p search section summary ul'.split(' '));
const BLOCKISH_END = new Set('address article aside blockquote button center details dialog dir div dl fieldset figcaption figure footer header hgroup listing main menu nav ol pre search section summary ul'.split(' '));
const IMPLIED = new Set('dd dt li optgroup option p rb rp rt rtc'.split(' '));
const VOID = new Set('br hr img input meta link area base col embed source track wbr'.split(' '));
const H = /^h[1-6]$/;
function buildTree(html: string): El {
    const root: El = { name: '#root', children: [], openBlock: 0 };
    const stack: El[] = [root];
    let block = 0;
    const cur = (): El => stack[stack.length - 1];
    const inScope = (test: (e: El) => boolean, extra: string[] = []): boolean => {
        for (let k = stack.length - 1; k > 0; k--) { const e = stack[k]; if (test(e)) return true; if (SCOPE.has(e.name) || extra.includes(e.name)) return false; }
        return false;
    };
    const pop = (by: string): El => { const e = stack.pop()!; e.closedBy = by; e.closeBlock = block; return e; };
    const popUntil = (test: (e: El) => boolean, by: string): void => { while (stack.length > 1) { const e = pop(by); if (test(e)) break; } };
    const impliedEnds = (except?: string): void => { while (IMPLIED.has(cur().name) && cur().name !== except) pop('implied'); };
    const closeP = (): void => { if (inScope(e => e.name === 'p', ['button'])) { impliedEnds('p'); popUntil(e => e.name === 'p', 'implied'); } };
    const insert = (name: string): El => { const e: El = { name, children: [], openBlock: block }; cur().children.push(e); stack.push(e); return e; };
    for (const t of tokenize(html)) {
        if (t.k === 'block') { block = t.id; continue; }
        if (t.k === 'text') { if (t.text) cur().children.push(t.text); continue; }
        const n = t.name;
        if (t.k === 'start') {
            if (BLOCKISH_START.has(n)) { closeP(); insert(n); }
            else if (H.test(n)) { closeP(); if (H.test(cur().name)) pop('implied'); insert(n); }
            else if (n === 'pre' || n === 'listing') { closeP(); insert(n); }
            else if (n === 'li') {
                for (let k = stack.length - 1; k > 0; k--) {
                    const e = stack[k];
                    if (e.name === 'li') { impliedEnds('li'); popUntil(x => x.name === 'li', 'implied'); break; }
                    if (SPECIAL.has(e.name) && !['address', 'div', 'p'].includes(e.name)) break;
                }
                closeP();
                insert('li');
            }
            else if (n === 'table') { closeP(); insert(n); }
            else if (VOID.has(n)) { if (n === 'hr') closeP(); insert(n); pop('void'); }
            else insert(n);
            continue;
        }
        if (BLOCKISH_END.has(n)) { if (!inScope(e => e.name === n)) continue; impliedEnds(); popUntil(e => e.name === n, '/' + n); }
        else if (n === 'p') { if (!inScope(e => e.name === 'p', ['button'])) insert('p'); impliedEnds('p'); popUntil(e => e.name === 'p', '/p'); }
        else if (n === 'li') { if (!inScope(e => e.name === 'li', ['ol', 'ul'])) continue; impliedEnds('li'); popUntil(e => e.name === 'li', '/li'); }
        else if (H.test(n)) { if (!inScope(e => H.test(e.name))) continue; impliedEnds(); popUntil(e => H.test(e.name), '/' + n); }
        else if (['table', 'tbody', 'thead', 'tfoot', 'tr', 'td', 'th'].includes(n)) {
            let found = false;
            for (let k = stack.length - 1; k > 0; k--) { if (stack[k].name === n) { found = true; break; } if (['html', 'table', 'template'].includes(stack[k].name)) break; }
            if (found) popUntil(e => e.name === n, '/' + n);
        }
        else {
            for (let k = stack.length - 1; k > 0; k--) {
                const e = stack[k];
                if (e.name === n) { impliedEnds(n); popUntil(x => x === e, '/' + n); break; }
                if (SPECIAL.has(e.name)) break;
            }
        }
    }
    while (stack.length > 1) pop('eof');
    return root;
}

// ---------- the disclosure tree ----------
/** Words outside any disclosure, and each disclosure's parent (-1: none), summary words and body words, in order. */
export interface DisclosureTree { top: string[]; ds: { parent: number; summary: string[]; body: string[] }[]; outOfDomain?: string }
const WORD = /\bw\d+\b/g;
const wordsOf = (s: string): string[] => s.match(WORD) ?? [];

/** The disclosure tree a browser shows: disclosures in a list item, heading or table cell, and in a summary, flattened (the AST has none there). */
function disclosureTree(root: El): DisclosureTree {
    const tree: DisclosureTree = { top: [], ds: [] };
    const walk = (node: El, d: number, summaryOf: number, flat: boolean): void => {
        for (const c of node.children) {
            if (typeof c === 'string') {
                const ws = wordsOf(c);
                if (summaryOf >= 0) tree.ds[summaryOf].summary.push(...ws);
                else if (d >= 0) tree.ds[d].body.push(...ws); else tree.top.push(...ws);
                continue;
            }
            if (c.name === 'details' && !flat && summaryOf < 0) {
                const idx = tree.ds.length;
                tree.ds.push({ parent: d, summary: [], body: [] });
                const summary = c.children.find((x): x is El => typeof x !== 'string' && x.name === 'summary');
                if (summary) {
                    const before = c.children.slice(0, c.children.indexOf(summary));
                    if (before.some(x => typeof x === 'string' ? /\S/.test(x) : true)) tree.outOfDomain ??= 'summary not first in its disclosure';
                    if (summary.closedBy !== '/summary' && summary.closeBlock !== summary.openBlock) tree.outOfDomain ??= `summary closed by ${summary.closedBy} in another block`;
                }
                for (const x of c.children) {
                    if (x === summary) walk(summary, idx, idx, flat);
                    else walk({ name: '#frag', children: [x], openBlock: 0 }, idx, -1, flat);
                }
                continue;
            }
            if (c.name === 'details' && flat && summaryOf < 0 && c.closedBy === '/details' && c.closeBlock !== c.openBlock) tree.outOfDomain ??= 'disclosure in a list item, heading or cell closed in another block';
            walk(c, d, summaryOf, flat || c.name === 'li' || H.test(c.name) || c.name === 'td' || c.name === 'th');
        }
    };
    walk(root, -1, -1, false);
    // An open finding, at the review checkout too: the HTML reader runs together the text of two summaries
    // standing side by side (white space between) in one block, but a disclosure's own summary.
    const sideBySide = (node: El): boolean => {
        let previous: El | undefined;
        for (const x of node.children) {
            if (typeof x === 'string') { if (/\S/.test(x)) previous = undefined; continue; }
            const own = node.name === 'details' && x === node.children.find(y => typeof y !== 'string' && y.name === 'summary');
            if (x.name === 'summary' && !own && previous && previous.openBlock === x.openBlock) return true;
            previous = x.name === 'summary' && !own ? x : undefined;
            if (sideBySide(x)) return true;
        }
        return false;
    };
    void sideBySide;
    return tree;
}

/** The disclosure tree GitHub shows for a Markdown text (see the file's comment). */
export function referenceDisclosureTree(markdown: string): DisclosureTree {
    const tree = disclosureTree(buildTree(markdownToHtml(markdown)));
    if (openReading) tree.outOfDomain ??= openReading;
    return tree;
}

/** The disclosure tree of a parsed document, disclosures in a list, table or heading flattened as the reference does. */
export function astDisclosureTree(content: OfficeContentNode[]): DisclosureTree {
    const tree: DisclosureTree = { top: [], ds: [] };
    const add = (d: number, ws: string[]): void => { if (d >= 0) tree.ds[d].body.push(...ws); else tree.top.push(...ws); };
    // The words of nodes in order (a summary's: disclosures in it are its text, as the reference has them).
    const wordsIn = (nodes: OfficeContentNode[] | undefined, into: string[]): void => {
        for (const n of nodes ?? []) {
            if (n.children?.length) wordsIn(n.children, into);
            else if (typeof n.text === 'string') into.push(...wordsOf(n.text));
        }
    };
    const walk = (nodes: OfficeContentNode[] | undefined, d: number, flat: boolean): void => {
        for (const n of nodes ?? []) {
            if (n.type === 'disclosure' && !flat) {
                const idx = tree.ds.length;
                tree.ds.push({ parent: d, summary: [], body: [] });
                const kids = n.children ?? [];
                const summary = kids[0]?.type === 'summary' ? kids[0] : undefined;
                if (summary) wordsIn(summary.children, tree.ds[idx].summary);
                walk(summary ? kids.slice(1) : kids, idx, flat);
                continue;
            }
            if (n.children?.length) walk(n.children, d, flat || n.type === 'list' || n.type === 'table' || n.type === 'heading');
            else if (typeof n.text === 'string') add(d, wordsOf(n.text));
        }
    };
    walk(content, -1, false);
    return tree;
}

// ---------- the generator ----------
/**
 * A generator of random Markdown mixing disclosures (opened, closed, summaries across blocks, tags
 * inline and as blocks, upper case, attributes with a quoted `>`) with quotes, nested quotes,
 * admonitions, lists, nested items, headings, pipe tables, fenced and indented code, HTML blocks and
 * tables, `<pre>` and rules, joined tight or by blank lines. Its words are w0, w1, ... in order, so
 * where each word lands tells the tree. Each call of the result draws the next text.
 */
export function disclosureMarkdownGenerator(seed: number): () => string {
    // A full-period linear congruential generator: Math.imul keeps the product's low 32 bits exact (in
    // floating point it passed 2^53 and lost them, and every seed fell into one short cycle).
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(a: T[]): T => a[Math.floor(rnd() * a.length)];
    const chance = (p: number): boolean => rnd() < p;
    let wc = 0;
    const W = (): string => `w${wc++}`;
    const words = (n = 1 + Math.floor(rnd() * 2)): string => Array.from({ length: n }, W).join(' ');

    // A <details> opening tag, sometimes with attributes, a quoted value holding `>` among them.
    const dOpen = (): string => pick(['<details>', '<details>', '<details open>', '<details title="a>b">', "<details class='x>y' open>", '<DETAILS>', '<details >']);
    const dClose = (): string => pick(['</details>', '</details>', '</DETAILS >']);
    const sOpen = (): string => pick(['<summary>', '<summary>', '<summary title="s>t">']);
    const sClose = (): string => pick(['</summary>', '</summary>', '</SUMMARY>']);

    // Inline material for a paragraph, heading, list item or table cell.
    const inlineBits = (allowTags: boolean): string => {
        const bits: string[] = [];
        const n = 1 + Math.floor(rnd() * 3);
        for (let i = 0; i < n; i++) {
            const r = rnd();
            if (!allowTags || r < 0.45) bits.push(words());
            else if (r < 0.52) bits.push('`' + pick(['<details>', '</details>', '<summary>']) + ' ' + W() + '`');
            else if (r < 0.56) bits.push('\\<details> ' + W());
            else if (r < 0.66) bits.push(dOpen());
            else if (r < 0.78) bits.push(dClose());
            else if (r < 0.88) bits.push(`${dOpen()}${sOpen()}${words(1)}${sClose()}`);
            else if (r < 0.93) bits.push(`${sOpen()}${words(1)}${sClose()}`);
            else if (r < 0.95) bits.push(pick([`<abbr title="the <details> element">${W()}</abbr>`, `<kbd title="</details>">${W()}</kbd>`, `<span title='</summary>'>${W()}</span>`, `${W()} <!-- <details> -->`, `<b data-x="<summary>">${W()}</b>`]));
            // (After a word: a line that one starts is an HTML block's, which the reader, as 8.6.0, does not end a quote's lazy lines at.)
            else if (r < 0.96) bits.push(`${W()} ${pick(['</div>', '<div>', '</section>'])}`);
            else bits.push('*' + W() + '*');
        }
        return bits.join(' ');
    };

    // Lines inside a container: text, tags alone, tags inline.
    const containerLine = (): string => {
        const r = rnd();
        if (r < 0.5) return inlineBits(true);
        if (r < 0.75) return dClose();
        if (r < 0.85) return dOpen();
        if (r < 0.92) return `${dOpen()}${sOpen()}${words(1)}${sClose()}`;
        return sClose();
    };

    const units: (() => string)[] = [
        // paragraphs
        () => inlineBits(true),
        () => inlineBits(true) + '\n' + inlineBits(true),
        () => words(),
        // headings
        () => pick(['#', '##', '###']) + ' ' + inlineBits(chance(0.5)),
        // quotes, nested quotes, admonitions, lazy lines
        () => {
            const depth = chance(0.25) ? '> > ' : '> ';
            const lines = [depth + (chance(0.2) ? '[!NOTE]' : words())];
            const n = 1 + Math.floor(rnd() * 3);
            for (let i = 0; i < n; i++) {
                const r = rnd();
                if (r < 0.6) lines.push(depth + containerLine());
                // A blank line of the quote at its depth: the reader reads nested quotes as one (as 8.6.0).
                else if (r < 0.75) lines.push(depth.trimEnd());
                else if (!/^>(?: >)?$/.test(lines[lines.length - 1])) {
                    // A lazy line, or one starting a block (an HTML block's line ends the unit: see afterHtml).
                    const lazy = chance(0.5) ? words() : containerLine();
                    // (Not a `<summary>` line, nor any line after a quote's `<summary>` line: the reader, as
                    // 8.6.0, reads it as the quote's lazy text where CommonMark ends the quote.)
                    if (/^<summary\b/i.test(lazy) || /^(?:> )+<summary\b/i.test(lines[lines.length - 1])) break;
                    lines.push(lazy);
                    if (lazy.startsWith('<')) break;
                }
            }
            return lines.join('\n');
        },
        // lists, nested items, continuation lines
        () => {
            const ordered = chance(0.3);
            const marker = (i: number): string => ordered ? `${i + 1}. ` : '- ';
            const lines: string[] = [];
            const n = 1 + Math.floor(rnd() * 3);
            for (let i = 0; i < n; i++) {
                lines.push(marker(i) + inlineBits(chance(0.6)));
                const pad = ' '.repeat(marker(i).length);
                if (chance(0.4)) lines.push(pad + containerLine());
                if (chance(0.2)) {
                    const nested = pad + '- ' + inlineBits(chance(0.5));
                    // After an HTML block's line in the item CommonMark has the nested item's line in that block,
                    // where the reader (as 8.6.0) reads a list item: kept out.
                    if (!/^(?:[-*+]|\d+\.)? *</.test(lines[lines.length - 1].trimStart())) lines.push(nested);
                }
            }
            return lines.join('\n');
        },
        // code
        () => '```\n' + words() + ' </details>\n<details>\n```',
        () => '    ' + words() + ' </details>',
        // pipe tables (cells confine their tags)
        () => `| ${words(1)} | ${inlineBits(chance(0.5))} |\n| --- | --- |\n| ${inlineBits(chance(0.5))} | ${words(1)} |`,
        // HTML blocks
        () => `<div align="center">${words()}</div>`,
        () => `<p align="center">${words()}</p>`,
        () => `<table><tr><td>${words(1)} ${pick(['</details>', '<details>', '<details><summary>' + W() + '</summary>' + W()])}</td></tr></table>`,
        () => '<pre>\n' + words() + '\n</pre>',
        () => '***',
        // disclosure tags as HTML blocks
        () => dOpen(),
        () => dOpen(),
        () => dClose(),
        () => dClose(),
        () => `${dOpen()}\n${sOpen()}${words()}${sClose()}`,
        () => `${dOpen()}${sOpen()}${words()}${sClose()}`,
        () => `${dOpen()}\n${sOpen()}\n\n${inlineBits(false)}\n\n${sClose()}`,
        () => `${dOpen()}\n${sOpen()}\n\n${pick(['###', '-', '>'])} ${words()}\n\n${sClose()}`,
        () => `${dOpen()}\n\n${sOpen()}${words()}${sClose()}`,
        () => `${dOpen()}${sOpen()}${words()}${sClose()}\n${words()}\n${dClose()}`,
        () => `${sOpen()}\n\n${words()}\n\n${sClose()}`,
        // tags in other tags' attribute values, in HTML blocks
        () => `<p title="</details>">${words()}</p>`,
        () => `<div data-x='<details>'>${words()}</div>`,
        // the Markdown writer's one-line alignment wrapper, its content Markdown
        () => `<div style="text-align: ${pick(['center', 'right'])}">${W()} ${inlineBits(true)}</div>`,
        () => `<div align="center">${W()} ${inlineBits(true)}</div>`,
        // elements enclosing disclosures: their end tags close what opened in them
        () => pick(['<div>', '<section>', '<div align="center">']),
        () => pick(['</div>', '</section>']),
        () => `<div>\n${dOpen()}\n${sOpen()}${words(1)}${sClose()}\n${words()}\n</div>`,
        () => `<div>\n${dOpen()}${sOpen()}${words(1)}${sClose()}`,
        // summaries holding blocks, links, pictures, breaks and notes; two summaries side by side
        () => `${dOpen()}\n${sOpen()}\n\n${pick(['###', '-', '>', '1.'])} ${words()}\n\n| ${W()} | ${W()} |\n| - | - |\n| ${W()} | ${W()} |\n\n${sClose()}`,
        () => `${dOpen()}\n${sOpen()}\n\n${W()} [${W()}](http://e.x/a) ![](i.png) **${W()}**[^n1]  \n${W()}\n\n${sClose()}\n\n${words()}`,
        () => `${dOpen()}\n${sOpen()}${W()}<br>${W()}${sClose()}`,
        () => `${dOpen()}\n${sOpen()}<b>${W()}</b><div>${W()}</div>${sClose()}`,
        () => `${dOpen()}\n${sOpen()}${W()}${sClose()}${sOpen()}${W()}${sClose()}`,
        () => `${dOpen()}\n${sOpen()}\n\n${dOpen()}\n${sOpen()}${W()}${sClose()}\n\n${words()}\n\n${dClose()}\n\n${sClose()}`,
    ];

    const sep = (): string => pick(['\n\n', '\n\n', '\n\n', '\n\n\n', '\n']);
    // Code, tables, rules and `<pre>` stand between blank lines: the reader lifts fenced code before it
    // reads blocks, and a tight join would make them a lazy line or an HTML block's text (a `<pre>`
    // block's too, which CommonMark ends at its `</pre>` line and the reader, as 8.6.0, at a blank line).
    // A one-line alignment wrapper too: the reader has it a block of its own where a `<div>` line follows it (as 8.6.0), CommonMark one HTML block.
    const standalone = (u: string): boolean => u.startsWith('```') || u.startsWith('    ') || u.startsWith('| ') || u === '***' || u.startsWith('<pre>')
        || /^<div (?:style|align)=[^\n]*<\/div>$/.test(u) || /^<\/?(?:div|section)[^>\n]*>$/.test(u) || u.endsWith('\n</div>') || u.startsWith('<div>\n');
    return () => {
        wc = 0;
        const k = 2 + Math.floor(rnd() * 9);
        let out = '';
        let last = '';
        // The text since the last blank line, where an HTML block that a line starts runs on.
        let run = '';
        for (let i = 0; i < k; i++) {
            const u = pick(units)();
            // Kept apart, as the reader (as 8.6.0) reads them other than CommonMark: a list, heading, quote
            // or rule after an HTML block's line, to the blank line ending it; a list item's lazy line; a
            // line after a quote's blank line; a quote line right after a paragraph's or a quote's of another
            // depth (the reader reads nested quotes as one); an HTML block's line
            // (but a disclosure's or a summary's end tag's) right after a quote, which the reader keeps in the
            // quote; any line after a quote's `<summary>` line, which the reader (as 8.6.0) keeps in the quote.
            const apart = /(?:^|\n)<[^\n]*/.test(run) && /^(?:[-#>*]|\d+\.)/.test(u)
                || /^(?:- |\d+\. )/.test(last) && !/^(?:[-#><]|\d+\.)/.test(u)
                || /\n>(?: >)?$/.test(last)
                || u.startsWith('>') && (!last.startsWith('>') || u.startsWith('> >') !== last.startsWith('> >'))
                || last.startsWith('>') && u.startsWith('<') && !/^(?:<details|<\/details|<\/summary)\b/i.test(u)
                || /(?:^|\n)(?:> )+<summary\b[^\n]*$/i.test(last);
            const join = i ? (standalone(u) || standalone(last) || apart ? '\n\n' : sep()) : '';
            out += join + u;
            run = join === '\n' ? run + join + u : u;
            last = u;
        }
        return out + (out.includes('[^n1]') ? '\n\n[^n1]: a note\n' : '\n');
    };
}

/** What checkDisclosureReference found. */
export interface DisclosureCheck { checked: number; outOfDomain: number; distinct: number; reasons: Record<string, number>; mismatches: string[]; unstable: string[] }

/**
 * `cases` random Markdown texts (see disclosureMarkdownGenerator) read by `parse`: each in domain must
 * give the reference's disclosure tree, and keep it over 3 round trips through `toMarkdown`, the
 * Markdown settled (the second and third the same).
 */
export async function checkDisclosureReference(parse: (markdown: string) => Promise<OfficeParserAST>,
    toMarkdown: (ast: OfficeParserAST) => Promise<string>, cases: number, seed = 1): Promise<DisclosureCheck> {
    const next = disclosureMarkdownGenerator(seed);
    const result: DisclosureCheck = { checked: 0, outOfDomain: 0, distinct: 0, reasons: {}, mismatches: [], unstable: [] };
    const seen = new Set<string>();
    const sig = (t: DisclosureTree): string => JSON.stringify({ top: t.top, ds: t.ds });
    for (let n = 0; n < cases; n++) {
        const source = next();
        if (!seen.has(source)) { seen.add(source); result.distinct++; }
        const ref = referenceDisclosureTree(source);
        if (ref.outOfDomain) { result.outOfDomain++; result.reasons[ref.outOfDomain] = (result.reasons[ref.outOfDomain] ?? 0) + 1; continue; }
        result.checked++;
        const ast = await parse(source);
        const got = sig(astDisclosureTree(ast.content));
        if (got !== sig(ref)) { result.mismatches.push(`${JSON.stringify(source)}\n  reference ${sig(ref)}\n  parsed    ${got}`); continue; }
        let current = ast;
        const markdown: string[] = [];
        let stable = true;
        for (let pass = 0; pass < 3 && stable; pass++) {
            markdown.push(await toMarkdown(current));
            current = await parse(markdown[pass]);
            stable = sig(astDisclosureTree(current.content)) === got;
        }
        if (!stable || markdown[1] !== markdown[2]) result.unstable.push(`${JSON.stringify(source)}\n  ${JSON.stringify(markdown)}`);
    }
    return result;
}

/** Whether a browser has a `<details>` element anywhere in what GitHub renders of a Markdown text (in a list item or heading too). */
export function referenceHasDisclosureElement(markdown: string): boolean {
    const has = (node: El): boolean => node.children.some(c => typeof c !== 'string' && (c.name === 'details' || has(c)));
    return has(buildTree(markdownToHtml(markdown)));
}
