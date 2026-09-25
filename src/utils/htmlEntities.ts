/**
 * Character references (`&name;`, `&#NN;`, `&#xHH;`), decoded as HTML and CommonMark decode them.
 *
 * The named set is HTML 4.01's (Latin-1, symbols and Greek letters, and the special characters:
 * typographic quotes, dashes, spaces), plus `&apos;`: the references real documents use. A name
 * outside it stays as written.
 */

/** Code points of the named references, by name (a Map, so a name such as `constructor` finds nothing). */
const NAMED_REFERENCES = new Map<string, number>(Object.entries({
    Aacute: 0xC1, aacute: 0xE1, Acirc: 0xC2, acirc: 0xE2, acute: 0xB4, AElig: 0xC6, aelig: 0xE6,
    Agrave: 0xC0, agrave: 0xE0, alefsym: 0x2135, Alpha: 0x391, alpha: 0x3B1, amp: 0x26, and: 0x2227,
    ang: 0x2220, apos: 0x27, Aring: 0xC5, aring: 0xE5, asymp: 0x2248, Atilde: 0xC3, atilde: 0xE3,
    Auml: 0xC4, auml: 0xE4, bdquo: 0x201E, Beta: 0x392, beta: 0x3B2, brvbar: 0xA6, bull: 0x2022,
    cap: 0x2229, Ccedil: 0xC7, ccedil: 0xE7, cedil: 0xB8, cent: 0xA2, Chi: 0x3A7, chi: 0x3C7,
    circ: 0x2C6, clubs: 0x2663, cong: 0x2245, copy: 0xA9, crarr: 0x21B5, cup: 0x222A, curren: 0xA4,
    Dagger: 0x2021, dagger: 0x2020, dArr: 0x21D3, darr: 0x2193, deg: 0xB0, Delta: 0x394, delta: 0x3B4,
    diams: 0x2666, divide: 0xF7, Eacute: 0xC9, eacute: 0xE9, Ecirc: 0xCA, ecirc: 0xEA, Egrave: 0xC8,
    egrave: 0xE8, empty: 0x2205, emsp: 0x2003, ensp: 0x2002, Epsilon: 0x395, epsilon: 0x3B5,
    equiv: 0x2261, Eta: 0x397, eta: 0x3B7, ETH: 0xD0, eth: 0xF0, Euml: 0xCB, euml: 0xEB, euro: 0x20AC,
    exist: 0x2203, fnof: 0x192, forall: 0x2200, frac12: 0xBD, frac14: 0xBC, frac34: 0xBE,
    frasl: 0x2044, Gamma: 0x393, gamma: 0x3B3, ge: 0x2265, gt: 0x3E, hArr: 0x21D4, harr: 0x2194,
    hearts: 0x2665, hellip: 0x2026, Iacute: 0xCD, iacute: 0xED, Icirc: 0xCE, icirc: 0xEE, iexcl: 0xA1,
    Igrave: 0xCC, igrave: 0xEC, image: 0x2111, infin: 0x221E, int: 0x222B, Iota: 0x399, iota: 0x3B9,
    iquest: 0xBF, isin: 0x2208, Iuml: 0xCF, iuml: 0xEF, Kappa: 0x39A, kappa: 0x3BA, Lambda: 0x39B,
    lambda: 0x3BB, lang: 0x2329, laquo: 0xAB, lArr: 0x21D0, larr: 0x2190, lceil: 0x2308, ldquo: 0x201C,
    le: 0x2264, lfloor: 0x230A, lowast: 0x2217, loz: 0x25CA, lrm: 0x200E, lsaquo: 0x2039,
    lsquo: 0x2018, lt: 0x3C, macr: 0xAF, mdash: 0x2014, micro: 0xB5, middot: 0xB7, minus: 0x2212,
    Mu: 0x39C, mu: 0x3BC, nabla: 0x2207, nbsp: 0xA0, ndash: 0x2013, ne: 0x2260, ni: 0x220B, not: 0xAC,
    notin: 0x2209, nsub: 0x2284, Ntilde: 0xD1, ntilde: 0xF1, Nu: 0x39D, nu: 0x3BD, Oacute: 0xD3,
    oacute: 0xF3, Ocirc: 0xD4, ocirc: 0xF4, OElig: 0x152, oelig: 0x153, Ograve: 0xD2, ograve: 0xF2,
    oline: 0x203E, Omega: 0x3A9, omega: 0x3C9, Omicron: 0x39F, omicron: 0x3BF, oplus: 0x2295,
    or: 0x2228, ordf: 0xAA, ordm: 0xBA, Oslash: 0xD8, oslash: 0xF8, Otilde: 0xD5, otilde: 0xF5,
    otimes: 0x2297, Ouml: 0xD6, ouml: 0xF6, para: 0xB6, part: 0x2202, permil: 0x2030, perp: 0x22A5,
    Phi: 0x3A6, phi: 0x3C6, Pi: 0x3A0, pi: 0x3C0, piv: 0x3D6, plusmn: 0xB1, pound: 0xA3, Prime: 0x2033,
    prime: 0x2032, prod: 0x220F, prop: 0x221D, Psi: 0x3A8, psi: 0x3C8, quot: 0x22, radic: 0x221A,
    rang: 0x232A, raquo: 0xBB, rArr: 0x21D2, rarr: 0x2192, rceil: 0x2309, rdquo: 0x201D, real: 0x211C,
    reg: 0xAE, rfloor: 0x230B, Rho: 0x3A1, rho: 0x3C1, rlm: 0x200F, rsaquo: 0x203A, rsquo: 0x2019,
    sbquo: 0x201A, Scaron: 0x160, scaron: 0x161, sdot: 0x22C5, sect: 0xA7, shy: 0xAD, Sigma: 0x3A3,
    sigma: 0x3C3, sigmaf: 0x3C2, sim: 0x223C, spades: 0x2660, sub: 0x2282, sube: 0x2286, sum: 0x2211,
    sup: 0x2283, sup1: 0xB9, sup2: 0xB2, sup3: 0xB3, supe: 0x2287, szlig: 0xDF, Tau: 0x3A4, tau: 0x3C4,
    there4: 0x2234, Theta: 0x398, theta: 0x3B8, thetasym: 0x3D1, thinsp: 0x2009, THORN: 0xDE,
    thorn: 0xFE, tilde: 0x2DC, times: 0xD7, trade: 0x2122, Uacute: 0xDA, uacute: 0xFA, uArr: 0x21D1,
    uarr: 0x2191, Ucirc: 0xDB, ucirc: 0xFB, Ugrave: 0xD9, ugrave: 0xF9, uml: 0xA8, upsih: 0x3D2,
    Upsilon: 0x3A5, upsilon: 0x3C5, Uuml: 0xDC, uuml: 0xFC, weierp: 0x2118, Xi: 0x39E, xi: 0x3BE,
    Yacute: 0xDD, yacute: 0xFD, yen: 0xA5, Yuml: 0x178, yuml: 0xFF, Zeta: 0x396, zeta: 0x3B6,
    zwj: 0x200D, zwnj: 0x200C
}));

/** A well-formed character reference. */
export const CHARACTER_REFERENCE = /&(#\d+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);/g;

/**
 * One reference decoded, given its body (`amp`, `#39`, `#x27`): the character, or undefined when
 * the name is unknown or the number is not a code point. NUL and lone surrogates are not
 * characters, and read as U+FFFD, as HTML and CommonMark read them.
 */
export function decodeCharacterReference(body: string): string | undefined {
    if (body[0] !== '#') {
        const codePoint = NAMED_REFERENCES.get(body);
        return codePoint === undefined ? undefined : String.fromCodePoint(codePoint);
    }
    const codePoint = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
    if (!Number.isFinite(codePoint) || codePoint > 0x10FFFF) return undefined;
    return codePoint === 0 || (codePoint >= 0xD800 && codePoint <= 0xDFFF) ? '\uFFFD' : String.fromCodePoint(codePoint);
}

/**
 * Every character reference in `text` decoded, in one pass, so `&amp;quot;` becomes the literal
 * `&quot;` (the exact inverse of escaping `&` first). An unknown name or an invalid number stays
 * as written.
 */
export function decodeCharacterReferences(text: string): string {
    return text.replace(CHARACTER_REFERENCE, (full: string, body: string) => decodeCharacterReference(body) ?? full);
}
