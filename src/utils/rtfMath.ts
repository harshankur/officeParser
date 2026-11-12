/**
 * RTF equations (`\mmath`), read as DOCX's OMML is.
 *
 * RTF writes Office Math as control words that map one to one onto OMML's elements: `\moMath` is
 * `m:oMath`, `\mf` is `m:f`, `\msSup` is `m:sSup`, `\mmr` (a matrix row) is `m:mr`. Two things differ: a
 * run's text sits in `\mr` itself (OMML: `m:r/m:t`), and a property's value is its group's text or the
 * word's parameter (OMML: `m:val`). The equation is built as an OMML string and read by `ommlToLatex`, so
 * an RTF equation reads as the same equation in a DOCX does.
 *
 * @module rtfMath
 */

import type { OfficeParserConfig } from '../types.js';
import type { RtfGroup, RtfNode } from '../parsers/RtfParser.js';
import { ommlToLatex } from './mathUtils.js';
import { escapeXml, stripInvalidXmlChars } from './sanitize.js';
import { parseXmlString } from './xmlUtils.js';

/** OMML's math namespace. */
const OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

/**
 * The OMML elements whose content is a value (`m:val`): in RTF, the text of their group (`{\mbegChr [}`,
 * `{\mtype lin}`) or the word's parameter (`\mdegHide1`).
 */
export const RTF_MATH_VALUE_WORDS: ReadonlySet<string> = new Set([
    'begChr', 'endChr', 'sepChr', 'chr', 'type', 'pos', 'scr', 'sty', 'limLoc', 'degHide', 'subHide', 'supHide', 'grow', 'vertJc', 'jc',
]);

/** Whether a control word is a math one (`\mXyz`), and the OMML element it names (`xyz`, its first letter lowered). */
const mathElement = (word: string | undefined): string | undefined =>
    word && /^m[a-zA-Z]+$/.test(word) ? word[1].toLowerCase() + word.slice(2) : undefined;

/** A `\uN`'s character (a negative N counts from 65536, as RTF writes characters past 32767). */
const unicodeOf = (param: number | undefined): string => {
    if (param === undefined) return '';
    const code = param < 0 ? param + 65536 : param;
    return code >= 0 && code <= 0xFFFF ? String.fromCharCode(code) : '';
};

/** A value word's value: its group's text, else its parameter. */
const valueOf = (group: RtfGroup): string => {
    let text = '';
    let param: number | undefined;
    for (const child of group.content) {
        if (child.type === 'text') text += child.value;
        else if (child.type === 'control' && child.value === 'u') text += unicodeOf(child.param);
        else if (child.type === 'control' && param === undefined && child.param !== undefined && mathElement(child.value)) param = child.param;
    }
    return text.length ? text : param !== undefined ? String(param) : '';
};

/** XML text, the characters XML cannot hold left out. */
const xmlTextOf = (text: string) => escapeXml(stripInvalidXmlChars(text));

/** A value element. */
const valueElement = (name: string, value: string) => `<m:${name} m:val="${xmlTextOf(value)}"/>`;

/**
 * The OMML of a math group's content, inside the element `name` (a run, `r`, keeps its text as its
 * `m:t`; text directly in another element is a run of its own). A group that is not math is read through
 * (its math groups and text in place); an ignorable one (`\*`) that is not math is skipped.
 */
function contentOmml(content: RtfNode[], name: string | undefined): string {
    let out = '';
    let text = '';
    const flushText = () => {
        // (White space between groups is how the file is laid out, not the equation's text.)
        if (text.length && (name === 'r' || text.trim().length)) out += name === 'r' ? `<m:t xml:space="preserve">${xmlTextOf(text)}</m:t>` : `<m:r><m:t xml:space="preserve">${xmlTextOf(text)}</m:t></m:r>`;
        text = '';
    };
    for (const child of content) {
        if (child.type === 'text') text += child.value;
        else if (child.type === 'control') {
            if (child.value === 'u') text += unicodeOf(child.param);
            else {
                // A value word written as a word with its parameter (`\mdegHide1`).
                const element = mathElement(child.value);
                if (element && RTF_MATH_VALUE_WORDS.has(element) && child.param !== undefined) { flushText(); out += valueElement(element, String(child.param)); }
            }
        } else {
            const element = mathElement(child.destination);
            if (element) {
                flushText();
                out += RTF_MATH_VALUE_WORDS.has(element) ? valueElement(element, valueOf(child)) : `<m:${element}>${contentOmml(child.content, element)}</m:${element}>`;
            } else if (!(child.content[0]?.type === 'control' && child.content[0].value === '*')) {
                flushText();
                out += contentOmml(child.content, name);
            }
        }
    }
    flushText();
    return out;
}

/**
 * An RTF equation (`\moMath` or `\moMathPara` inside `\mmath`) as LaTeX, read as DOCX's OMML is: each math
 * control word `\mXyz` is OMML's `m:xyz`, a run's text (`\mr`) its `m:t`, a property's text or parameter
 * (`\mbegChr`, `\mchr`, `\mtype`, `\mpos`, `\mscr`, `\msty` and the other value words) its `m:val`; the
 * OMML is read by `ommlToLatex`. Text directly in another element is read as a run; a group without a math
 * destination is read through (its math groups and text in place), an ignorable one (`\*`) not in math
 * skipped; other control words (fonts, sizes) ignored; `\uN` is its character.
 *
 * Every element is charged to `decompressionLimits.maxXmlElements` (through `parseXmlString`).
 *
 * @param math - The `\moMathPara` or `\moMath` group
 * @param config - The parse's config, whose element budget the equation is charged to
 * @returns the LaTeX (empty when the equation has no text) and whether it is a display equation (`\moMathPara`).
 */
export function rtfMathToLatex(math: RtfGroup, config: OfficeParserConfig): { latex: string; block: boolean } {
    const root = mathElement(math.destination) ?? 'oMath';
    const xml = `<m:${root} xmlns:m="${OMML_NS}">${contentOmml(math.content, root)}</m:${root}>`;
    const doc = parseXmlString(xml, { config });
    return { latex: doc.documentElement ? ommlToLatex(doc.documentElement) : '', block: root === 'oMathPara' };
}
