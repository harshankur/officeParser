/**
 * The template engine's look-alike scan (`findLookalikeTags`: text that looks like a tag but whose
 * delimiters are not the delimiters) checked against a reference written plainly: braces are objects on
 * an explicit stack, each look-alike owns the braces it opens, and a missing end that left braces open is
 * reported when one it owns is still on the stack at the paragraph's end. No stops, no counters, no rank:
 * quadratic, which is fine for the short texts the property draws.
 *
 * The reading it encodes (the engine's doc comment states the same):
 * - A delimiter is matched character by character, each as stored (XML-escaped) or full-width.
 * - For delimiters of one character typed twice or more each side, a delimiter one character short is a
 *   look-alike (`missing-delimiter`); one with a full-width character is `full-width`. Whole delimiters
 *   as stored around a name are a tag (or text between the delimiters) and skipped whole.
 * - Braces (the start's first character and the end's last, as stored) are counted only for such
 *   delimiters, and only when the two characters differ: with one character for both (`$$`), none is.
 *   A look-alike opens its plain start characters, then closes its plain end characters while any is open.
 * - A missing start is not one inside braces still open, nor right after the start's character.
 * - A missing end that opens more than it closes is reported only if one of its braces is never closed.
 */

export interface Lookalike { start: number; end: number; reason: 'full-width' | 'missing-delimiter' }

const escapeXmlText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const fullWidthOf = (c: string): string | null => {
    const code = c.charCodeAt(0);
    return c.length === 1 && code >= 0x21 && code <= 0x7E ? String.fromCharCode(code + 0xFEE0) : null;
};
const doubled = (chars: string[]): boolean => chars.length >= 2 && chars.every(c => c === chars[0]);
const MAX_NEAR_NAME = 64;

/** What the scan should find in `joined` (a paragraph's text as stored) with these delimiters, in document order. */
export function referenceLookalikes(joined: string, S: string, E: string): Lookalike[] {
    const start = Array.from(S), end = Array.from(E);
    const formsOf = (c: string) => ({ plain: escapeXmlText(c), wide: fullWidthOf(c) });
    const sf = start.map(formsOf), ef = end.map(formsOf);
    const canMiss = doubled(start) && doubled(end);
    const o = sf[0], k = ef[ef.length - 1];
    const braces = canMiss && o.plain !== k.plain;
    type Box = { owner: object | null };
    const stack: Box[] = [];
    const out: Lookalike[] = [];
    const pend: { item: Lookalike; boxes: Box[] }[] = [];
    const match = (at: number, forms: typeof sf): { end: number; wide: number } | null => {
        let wide = 0;
        for (const f of forms) {
            if (joined.slice(at, at + f.plain.length) === f.plain) at += f.plain.length;
            else if (f.wide !== null && joined.slice(at, at + f.wide.length) === f.wide) { at += f.wide.length; wide++; }
            else return null;
        }
        return { end: at, wide };
    };
    let i = 0;
    while (i < joined.length) {
        let open = match(i, sf), chars = sf.length, missingStart = false;
        if (!open && canMiss) {
            const short = match(i, sf.slice(1));
            const before = joined.slice(0, i);
            const precededByOpener = i > 0 && (before.endsWith(o.plain) || (o.wide !== null && before.endsWith(o.wide)));
            if (short && !short.wide && stack.length === 0 && !precededByOpener) { open = short; chars = sf.length - 1; missingStart = true; }
        }
        let cand: { end: number; wide: number; opens: number; closes: number; missingEnd: boolean } | null = null;
        if (open) {
            let at = open.end;
            while (at < joined.length && at - open.end < MAX_NEAR_NAME && /^[\p{L}\p{N}_.\-\s]$/u.test(joined[at])) at++;
            const name = joined.slice(open.end, at).trim();
            if (name !== '' && /^[\p{L}_]/u.test(name)) {
                let close = match(at, ef), closeChars = ef.length, missingEnd = false;
                if (!close && canMiss && !missingStart) {
                    const short = match(at, ef.slice(0, -1));
                    if (short && !short.wide) { close = short; closeChars = ef.length - 1; missingEnd = true; }
                }
                if (close) cand = { end: close.end, wide: open.wide + close.wide, opens: chars - open.wide, closes: closeChars - close.wide, missingEnd };
                if (cand && !cand.wide && !missingStart && !missingEnd) { i = cand.end; continue; } // a tag: skipped whole
            }
        }
        if (!cand) {
            if (braces) {
                if (joined.slice(i, i + o.plain.length) === o.plain) stack.push({ owner: null });
                else if (joined.slice(i, i + k.plain.length) === k.plain && stack.length) stack.pop();
            }
            i++;
            continue;
        }
        const item: Lookalike = { start: i, end: cand.end, reason: cand.wide ? 'full-width' : 'missing-delimiter' };
        const boxes: Box[] = [];
        if (braces) {
            for (let n = 0; n < cand.opens; n++) { const b = { owner: item }; stack.push(b); boxes.push(b); }
            for (let n = 0; n < cand.closes && stack.length; n++) stack.pop();
        }
        if (cand.missingEnd && !cand.wide && braces && cand.opens > cand.closes) pend.push({ item, boxes });
        else out.push(item);
        i = cand.end;
    }
    for (const p of pend) if (p.boxes.some(b => stack.includes(b))) out.push(p.item);
    return out.sort((a, b) => a.start - b.start);
}

/** The scan under test: what it lists with `room` left, and how many more it counts. */
export type LookalikeScan = (joined: string, start: string, end: string, room: number) => { found: { start: number; end: number; reason: string }[]; unlisted: number };

/**
 * Random delimiter pairs (one character typed twice or more each side, the same character both sides,
 * escaped and non-ASCII characters, space edges, and arbitrary words of 1 to 5 characters) and random texts
 * built from whole, short and full-width delimiters, single braces, names and noise, half of them shaped
 * like a tag. Each text is scanned with rooms 0, 1, 2, 3 and 1,000: the scan must list the reference's
 * first `room` and count the rest exactly. Returns the failures (at most 10), described.
 */
export function checkLookalikeScan(scan: LookalikeScan, cases: number, seed = 1): { withLookalikes: number; failures: string[] } {
    // The C library generator, modulo 2^31 in integer arithmetic, which has its full period: multiplied
    // in floating point the product passed 2^53, lost its low bits, and every seed fell into one cycle
    // of 10,466 values. Divided by 2^31, a draw is under 1, so a pick stays inside its list.
    const rnd = (): number => { seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff; return seed / 0x80000000; };
    const pick = <T>(xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)];
    const CHARS = ['{', '}', '[', ']', '<', '>', '&', '$', '%', '|', '(', ')', ' ', 'a', '«', '»', '#'];
    const delimiters = (): { start: string; end: string } => {
        const r = rnd();
        const len = (): number => 1 + Math.floor(rnd() * 5);
        if (r < 0.45) {
            const [a, b] = pick([['{', '}'], ['[', ']'], ['<', '>'], ['(', ')'], ['$', '$'], ['&', '&'], ['%', '%'], ['{', '{'], ['«', '»'], ['|', '|']]);
            return { start: a.repeat(1 + len()), end: b.repeat(1 + len()) };
        }
        if (r < 0.65) {
            const [a, b] = pick([['{', '}'], ['[', ']'], ['<', '>']]);
            return { start: ' '.repeat(len()) + a.repeat(len()), end: b.repeat(len()) + ' '.repeat(len()) };
        }
        const word = (): string => Array.from({ length: len() }, () => pick(CHARS)).join('');
        const start = word();
        return { start, end: rnd() < 0.15 ? start : word() };
    };
    const text = (d: { start: string; end: string }): string => {
        const s = Array.from(d.start), e = Array.from(d.end);
        const wideOne = (cs: string[]): string => { const k = Math.floor(rnd() * cs.length); const w = fullWidthOf(cs[k]); return cs.map((c, j) => j === k && w ? w : c).join(''); };
        const tokens: (() => string)[] = [
            () => d.start, () => d.end, () => s.slice(1).join(''), () => e.slice(0, -1).join(''),
            () => wideOne(s), () => wideOne(e), () => s[0], () => e[e.length - 1], () => s[0], () => e[e.length - 1],
            () => 'a', () => ' name ', () => 'x y', () => ' ', () => ',', () => '1', () => '&', () => '<',
            () => fullWidthOf(s[0]) ?? 'b', () => fullWidthOf(e[e.length - 1]) ?? 'c',
        ];
        const shaped = (): string => pick([() => d.start, () => s.slice(1).join(''), () => wideOne(s)])() + pick(['a', ' a ', 'name', ' x y ', '', '1'])
            + pick([() => d.end, () => e.slice(0, -1).join(''), () => wideOne(e)])();
        let out = '';
        const n = 1 + Math.floor(rnd() * 14);
        for (let k = 0; k < n; k++) out += rnd() < 0.4 ? shaped() : pick(tokens)();
        return escapeXmlText(out);
    };
    let withLookalikes = 0;
    const failures: string[] = [];
    for (let k = 0; k < cases && failures.length < 10; k++) {
        const d = delimiters();
        const t = text(d);
        const want = referenceLookalikes(t, d.start, d.end);
        if (want.length) withLookalikes++;
        for (const room of [0, 1, 2, 3, 1000]) {
            const { found, unlisted } = scan(t, d.start, d.end, room);
            if (JSON.stringify(found) !== JSON.stringify(want.slice(0, room)) || unlisted !== want.length - Math.min(room, want.length)) {
                failures.push(`${JSON.stringify(d)} ${JSON.stringify(t)} room ${room}: want ${JSON.stringify(want)}, got ${JSON.stringify(found)} and ${unlisted} more`);
                break;
            }
        }
    }
    return { withLookalikes, failures };
}
