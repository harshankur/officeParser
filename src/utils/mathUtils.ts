import { isElement } from './xmlUtils.js';

/**
 * Equation markup normalization.
 *
 * Office documents ship equations in exactly two markups: OOXML's OMML (`<m:oMath>`, used by
 * DOCX and PPTX) and MathML (`<math>`, used by ODF embedded objects, HTML, and EPUB3). Neither
 * is plain text, and neither survives the generic "recurse into unknown elements and concatenate
 * their text" fallback every parser ends with: `<m:num>1</m:num><m:den>2</m:den>` collapses to
 * `12`, which still reads as a number, so nothing downstream can tell the value is wrong.
 *
 * Both markups are converted to LaTeX here rather than to a per-format ad hoc notation, because
 * `MarkdownParser` already emits LaTeX for `$...$` / `$$...$$`. Converging on it means one
 * representation reaches every generator, and a formula survives a docx -> md -> docx round trip
 * instead of degrading at each hop.
 *
 * The output is always emitted as a `code` node carrying `CodeMetadata.math`, matching the
 * contract `MarkdownParser` established - see `src/types.ts`.
 */

/**
 * Depth cap for the recursive walks below. Equation markup nests (a fraction inside a
 * superscript inside a fraction), and the nesting is attacker-controlled: a few hundred bytes of
 * hand-written XML can carry thousands of levels. The walks are recursive, so an unbounded
 * document would exhaust the stack rather than merely producing odd output. 64 is far past any
 * real equation - the deepest construct in a typical maths paper is 3 or 4 levels.
 */
const MAX_MATH_DEPTH = 64;

/** Marker substituted for a subtree that exceeded MAX_MATH_DEPTH, so truncation is never silent. */
const TRUNCATED = '\\ldots';

/**
 * LaTeX metacharacters, escaped in any literal run of document text.
 *
 * The content of `<m:t>`, `<mi>`, `<mn>` and friends is literal characters, never LaTeX source -
 * an author who types `%` into an equation means a percent sign, not a comment. Escaping is
 * therefore lossless, and it also stops document text from injecting control sequences into the
 * `$...$` span it lands in. Backslash must be replaced first or it would re-escape the
 * backslashes introduced by the later replacements.
 */
const escapeLatex = (text: string): string =>
    text.replace(/\\/g, '\\textbackslash{}')
        .replace(/([&%$#_{}])/g, '\\$1')
        .replace(/~/g, '\\textasciitilde{}')
        .replace(/\^/g, '\\textasciicircum{}');

/**
 * Wraps an expression in braces unless it is already a single token.
 *
 * `x^{2}` and `x^2` render identically, but `x^{2y}` and `x^2y` do not, so the braces cannot be
 * dropped whenever the argument is longer than one character. A lone digit or letter is the only
 * safe case, and it is by far the most common one, so special-casing it keeps ordinary output
 * readable without risking a mis-grouped exponent.
 */
const group = (latex: string): string =>
    /^[0-9a-zA-Z]$/.test(latex) ? latex : `{${latex}}`;

/** Strips any namespace prefix: `m:oMath` -> `omath`, `mml:mfrac` -> `mfrac`. */
const localName = (element: Element): string =>
    element.tagName.toLowerCase().replace(/^.*:/, '');

/** Element children only, in document order. */
const elementChildren = (element: Element): Element[] => {
    const result: Element[] = [];
    for (let i = 0; i < (element.childNodes?.length ?? 0); i++) {
        const child = element.childNodes[i];
        if (isElement(child)) result.push(child);
    }
    return result;
};

// ─── OMML (OOXML: DOCX, PPTX) ────────────────────────────────────────────────

/**
 * Property elements. These carry styling (`m:ctrlPr` even wraps a full `w:rPr`) and never
 * contribute display text, so they must be skipped rather than descended into - the generic
 * fallback descending into them is one of the ways stray formatting text reached the output.
 */
const OMML_PROPERTY_TAGS = new Set([
    'fpr', 'rpr', 'ctrlpr', 'ssubpr', 'ssuppr', 'ssubsuppr', 'dpr', 'radpr', 'narypr',
    'funcpr', 'limlowpr', 'limupppr', 'mpr', 'argpr', 'barpr', 'accpr', 'grouppr',
    'phantpr', 'boxpr', 'eqarrpr', 'spre', 'sty', 'scr', 'brk', 'aln', 'nor',
]);

/**
 * `m:scr` math alphabets. OOXML encodes `ℝ` as an ASCII `R` plus a script attribute rather than
 * as the Unicode character ODF uses, so these carry meaning and not merely styling.
 */
const OMML_MATH_ALPHABETS: Record<string, string> = {
    'double-struck': '\\mathbb',
    'script': '\\mathcal',
    'fraktur': '\\mathfrak',
    'monospace': '\\mathtt',
    'sans-serif': '\\mathsf',
    'roman': '\\mathrm',
};

/** Named OMML functions that map onto a LaTeX command of the same meaning. */
const OMML_NARY_OPERATORS: Record<string, string> = {
    '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint',
    '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap', '⋀': '\\bigwedge', '⋁': '\\bigvee',
};

/**
 * Converts an OMML subtree (`<m:oMath>` or any node within one) to LaTeX.
 *
 * Covers the constructs that actually appear in office documents: fractions, sub/superscripts,
 * delimiters, radicals, n-ary operators, functions, accents, bars, boxes and matrices. Anything
 * unrecognized falls through to concatenating its children, which is the old behaviour and the
 * right degradation for a construct that carries no grouping of its own.
 */
export const ommlToLatex = (node: Node, depth = 0): string => {
    if (!node) return '';
    if (node.nodeType === 3) return escapeLatex(node.textContent || '');
    if (!isElement(node)) return '';
    if (depth > MAX_MATH_DEPTH) return TRUNCATED;

    const element = node;
    const tag = localName(element);
    if (OMML_PROPERTY_TAGS.has(tag)) return '';

    const kids = elementChildren(element);
    /** Concatenates every child, used for containers and for unrecognized constructs. */
    const all = () => kids.map(k => ommlToLatex(k, depth + 1)).join('');
    /** The LaTeX for the first `<m:xxx>` child with the given local name, or '' when absent. */
    const part = (name: string) => {
        const found = kids.find(k => localName(k) === name);
        return found ? ommlToLatex(found, depth + 1) : '';
    };

    switch (tag) {
        case 'r': {
            // A run may carry a math alphabet on its own `m:rPr` - `ℝ` is written as a plain `R`
            // with `<m:scr m:val="double-struck"/>`, so dropping the property loses the meaning
            // of the symbol, not just its look. Both `m:rPr` and `w:rPr` reduce to the same local
            // name, so select on content rather than on the prefix.
            const body = all();
            const rPr = kids.find(k => localName(k) === 'rpr'
                && elementChildren(k).some(p => localName(p) === 'scr' || localName(p) === 'sty'));
            if (!rPr || !body) return body;
            const valueOf = (name: string) => {
                const el = elementChildren(rPr).find(p => localName(p) === name);
                return el?.getAttribute('m:val') ?? el?.getAttribute('val') ?? '';
            };
            const alphabet = OMML_MATH_ALPHABETS[valueOf('scr')]
                ?? (valueOf('sty') === 'b' ? '\\mathbf' : undefined);
            return alphabet ? `${alphabet}${group(body)}` : body;
        }

        // Containers: an equation, an argument, a base.
        case 'omath':
        case 'omathpara':
        case 'e':
        case 'num':
        case 'den':
        case 'sub':
        case 'sup':
        case 'lim':
        case 'fname':
            return all();

        case 't':
            return escapeLatex(element.textContent || '');

        case 'f': {
            // `m:type val="lin"` asks for an inline `a/b` rather than a stacked fraction.
            const fPr = kids.find(k => localName(k) === 'fpr');
            const typeEl = fPr ? elementChildren(fPr).find(k => localName(k) === 'type') : undefined;
            const linear = typeEl?.getAttribute('m:val') === 'lin' || typeEl?.getAttribute('val') === 'lin';
            const num = part('num');
            const den = part('den');
            return linear ? `${group(num)}/${group(den)}` : `\\frac${group(num)}${group(den)}`;
        }

        case 'ssup':
            return `${group(part('e'))}^${group(part('sup'))}`;
        case 'ssub':
            return `${group(part('e'))}_${group(part('sub'))}`;
        case 'ssubsup':
            return `${group(part('e'))}_${group(part('sub'))}^${group(part('sup'))}`;
        case 'spre':
            // Pre-sub/superscript: the scripts precede the base.
            return `{}_${group(part('sub'))}^${group(part('sup'))}${group(part('e'))}`;

        case 'd': {
            // Delimiters. The characters are document-supplied and default to parentheses.
            // Plain delimiters rather than \left...\right: an unbalanced \left would break the
            // whole expression, and a document can legitimately open without closing.
            const dPr = kids.find(k => localName(k) === 'dpr');
            const chr = (name: string, fallback: string) => {
                const el = dPr ? elementChildren(dPr).find(k => localName(k) === name) : undefined;
                const raw = el?.getAttribute('m:val') ?? el?.getAttribute('val');
                return raw ? escapeLatex(raw) : fallback;
            };
            const beg = chr('begchr', '(');
            const end = chr('endchr', ')');
            const sep = chr('sepchr', ',');
            const args = kids.filter(k => localName(k) === 'e').map(k => ommlToLatex(k, depth + 1));
            return `${beg}${args.join(sep)}${end}`;
        }

        case 'rad': {
            const deg = part('deg');
            const base = group(part('e'));
            return deg ? `\\sqrt[${deg}]${base}` : `\\sqrt${base}`;
        }

        case 'nary': {
            // Summation/integral and friends: operator, optional bounds, then the operand.
            const naryPr = kids.find(k => localName(k) === 'narypr');
            const chrEl = naryPr ? elementChildren(naryPr).find(k => localName(k) === 'chr') : undefined;
            const chr = chrEl?.getAttribute('m:val') ?? chrEl?.getAttribute('val') ?? '∫';
            const op = OMML_NARY_OPERATORS[chr] ?? escapeLatex(chr);
            const sub = part('sub');
            const sup = part('sup');
            return `${op}${sub ? `_${group(sub)}` : ''}${sup ? `^${group(sup)}` : ''}${part('e')}`;
        }

        case 'func':
            // `sin`, `log`, ... - the name is document text, so it cannot become a bare command.
            return `\\operatorname${group(part('fname'))}${group(part('e'))}`;

        case 'limlow':
            return `${part('e')}_${group(part('lim'))}`;
        case 'limupp':
            return `${part('e')}^${group(part('lim'))}`;

        case 'bar': {
            const barPr = kids.find(k => localName(k) === 'barpr');
            const posEl = barPr ? elementChildren(barPr).find(k => localName(k) === 'pos') : undefined;
            const pos = posEl?.getAttribute('m:val') ?? posEl?.getAttribute('val');
            return `${pos === 'top' ? '\\overline' : '\\underline'}${group(part('e'))}`;
        }

        case 'acc': {
            const accPr = kids.find(k => localName(k) === 'accpr');
            const chrEl = accPr ? elementChildren(accPr).find(k => localName(k) === 'chr') : undefined;
            const chr = chrEl?.getAttribute('m:val') ?? chrEl?.getAttribute('val') ?? '̂';
            const command = chr === '̄' ? '\\bar' : chr === '⃗' ? '\\vec' : chr === '̇' ? '\\dot' : '\\hat';
            return `${command}${group(part('e'))}`;
        }

        case 'box':
        case 'borderbox':
        case 'phant':
        case 'group':
        case 'groupchr':
            return part('e') || all();

        case 'm': {
            // Matrix: rows of cells. `\begin{matrix}` carries no delimiters of its own, which is
            // correct - a bracketed matrix wraps the `m:m` in an `m:d` that supplies them.
            const rows = kids.filter(k => localName(k) === 'mr').map(row =>
                elementChildren(row)
                    .filter(cell => localName(cell) === 'e')
                    .map(cell => ommlToLatex(cell, depth + 1))
                    .join(' & ')
            );
            return `\\begin{matrix}${rows.join(' \\\\ ')}\\end{matrix}`;
        }

        default:
            return all();
    }
};

// ─── MathML (ODF embedded objects, HTML, EPUB3) ──────────────────────────────

/** MathML operators that have a dedicated LaTeX command. */
const MATHML_OPERATORS: Record<string, string> = {
    '∑': '\\sum', '∏': '\\prod', '∫': '\\int', '∮': '\\oint', '√': '\\sqrt',
    '±': '\\pm', '∓': '\\mp', '×': '\\times', '÷': '\\div', '⋅': '\\cdot',
    '≤': '\\leq', '≥': '\\geq', '≠': '\\neq', '≈': '\\approx', '≡': '\\equiv',
    '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊆': '\\subseteq', '∪': '\\cup',
    '∩': '\\cap', '∞': '\\infty', '→': '\\to', '⇒': '\\Rightarrow', '⇔': '\\Leftrightarrow',
    '∀': '\\forall', '∃': '\\exists', '∂': '\\partial', '∇': '\\nabla', '…': '\\ldots',
    'ℝ': '\\mathbb{R}', 'ℕ': '\\mathbb{N}', 'ℤ': '\\mathbb{Z}', 'ℚ': '\\mathbb{Q}', 'ℂ': '\\mathbb{C}',
};

/** Maps a literal run through the operator table, falling back to plain escaped text. */
const mathmlToken = (raw: string): string => {
    const trimmed = raw.trim();
    return MATHML_OPERATORS[trimmed] ?? escapeLatex(raw);
};

/**
 * The subset of a parsed element every MathML source here can present.
 *
 * MathML reaches this module from two different tree types - the XML DOM (ODF embedded objects)
 * and `HtmlParser`'s own lightweight node (HTML, and EPUB via its spine items) - so the converter
 * is written against this shape and each caller adapts into it. That keeps one implementation of
 * the conversion rather than one per tree type, which is how HTML came to have no MathML support
 * at all while ODF did.
 */
export interface MathNode {
    /** Tag name, namespace prefix included or not; `undefined` marks a text node. */
    tagName?: string;
    attributes?: Record<string, string>;
    /** Literal text, for text nodes and for leaf tokens. */
    text?: string;
    children: MathNode[];
}

/** Local name of a MathNode: `mml:mfrac` -> `mfrac`, `undefined` for a text node. */
const mathNodeName = (node: MathNode): string =>
    (node.tagName || '').toLowerCase().replace(/^.*:/, '');

/** Presents an XML DOM node through the MathNode shape. */
const fromDom = (node: Node): MathNode => {
    if (node.nodeType === 3 || !isElement(node)) {
        return { text: node.textContent || '', children: [] };
    }
    const element = node;
    const attributes: Record<string, string> = {};
    for (let i = 0; i < (element.attributes?.length ?? 0); i++) {
        const attr = element.attributes[i];
        attributes[attr.name] = attr.value;
    }
    return {
        tagName: element.tagName,
        attributes,
        text: element.textContent || '',
        children: elementChildren(element).map(fromDom),
    };
};

/**
 * Converts a MathML subtree to LaTeX.
 *
 * When the document carries a TeX annotation (`<annotation encoding="application/x-tex">`), that
 * is the author's own source and is used verbatim in preference to anything reconstructed here.
 * ODF's `<annotation encoding="StarMath 5.0">` is deliberately not used - StarMath is not LaTeX,
 * and emitting it would put a second notation back into the output this module exists to unify.
 */
export const mathmlTreeToLatex = (node: MathNode, depth = 0): string => {
    if (!node) return '';
    if (!node.tagName) return mathmlToken(node.text || '');
    if (depth > MAX_MATH_DEPTH) return TRUNCATED;

    const tag = mathNodeName(node);
    const kids = node.children ?? [];
    /**
     * The literal content of a leaf token. The DOM adapter fills `text` with `textContent`, but
     * `HtmlParser` keeps an element's text in child text nodes and leaves `text` unset, so fall
     * back to gathering the children rather than emitting an empty token.
     */
    const tokenText = () => node.text || kids.map(k => k.text || '').join('');
    const all = () => kids.map(k => mathmlTreeToLatex(k, depth + 1)).join('');
    const arg = (index: number) => (kids[index] ? mathmlTreeToLatex(kids[index], depth + 1) : '');

    switch (tag) {
        case 'math':
        case 'semantics': {
            const tex = kids.find(k => mathNodeName(k) === 'annotation'
                && /tex/i.test(k.attributes?.['encoding'] || ''));
            // Same reason `tokenText` exists below: an element's text is in `text` for the DOM
            // adapter and in child text nodes for `HtmlParser`, so both have to be consulted.
            if (tex) return (tex.text || (tex.children ?? []).map(c => c.text || '').join('')).trim();
            return kids.map(k => mathmlTreeToLatex(k, depth + 1)).join('');
        }

        case 'mrow':
        case 'mstyle':
        case 'mpadded':
        case 'mphantom':
            return all();

        case 'mi':
        case 'mn':
        case 'mo':
        case 'mtext':
        case 'ms':
            return mathmlToken(tokenText());

        case 'mfrac':
            return `\\frac${group(arg(0))}${group(arg(1))}`;
        case 'msup':
            return `${group(arg(0))}^${group(arg(1))}`;
        case 'msub':
            return `${group(arg(0))}_${group(arg(1))}`;
        case 'msubsup':
            return `${group(arg(0))}_${group(arg(1))}^${group(arg(2))}`;
        case 'munder':
            return `\\underset${group(arg(1))}${group(arg(0))}`;
        case 'mover':
            return `\\overset${group(arg(1))}${group(arg(0))}`;
        case 'munderover':
            return `${group(arg(0))}_${group(arg(1))}^${group(arg(2))}`;

        case 'msqrt':
            return `\\sqrt${group(all())}`;
        case 'mroot':
            return `\\sqrt[${arg(1)}]${group(arg(0))}`;

        case 'mfenced': {
            // Deprecated in MathML 3 but still emitted by older producers.
            const open = escapeLatex(node.attributes?.['open'] ?? '(');
            const close = escapeLatex(node.attributes?.['close'] ?? ')');
            const sep = escapeLatex(node.attributes?.['separators'] ?? ',');
            return `${open}${kids.map(k => mathmlTreeToLatex(k, depth + 1)).join(sep)}${close}`;
        }

        case 'mtable':
            return `\\begin{matrix}${kids
                .filter(k => mathNodeName(k) === 'mtr')
                .map(row => (row.children ?? [])
                    .filter(cell => mathNodeName(cell) === 'mtd')
                    .map(cell => mathmlTreeToLatex(cell, depth + 1))
                    .join(' & '))
                .join(' \\\\ ')}\\end{matrix}`;
        case 'mtr':
        case 'mtd':
            return all();

        case 'mspace':
            return ' ';

        // Presentation-only wrappers and the annotations themselves carry nothing renderable:
        // `annotation-xml` duplicates the presentation tree in Content MathML, and emitting both
        // would double every formula that has one.
        case 'annotation':
        case 'annotation-xml':
        case 'maction':
            return '';

        default:
            return all();
    }
};

/** Converts a MathML subtree held in an XML DOM (ODF embedded objects) to LaTeX. */
export const mathmlToLatex = (node: Node, depth = 0): string =>
    mathmlTreeToLatex(fromDom(node), depth);

/**
 * True when a converted equation carries nothing worth emitting, so callers can drop the node
 * instead of pushing an empty `$$` into the output.
 */
export const isEmptyMath = (latex: string): boolean => latex.trim().length === 0;
