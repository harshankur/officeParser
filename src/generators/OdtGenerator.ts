import { zipSync, Zippable } from 'fflate';
import { UniqueNames } from '../utils/uniqueNames.js';
import { layoutTableRows } from '../utils/tableLayout.js';
import { BreakMetadata, ConversionResult, GeneratorConfig, HeaderFooterMetadata, ImageMode, OdtGeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType, PageNumberFormat, PageSetup, TextFormatting } from '../types.js';
import { odfNumFormatOf } from '../utils/pageNumbers.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, isSafeStyleMapTag, sanitizeOfficePackageUrl, stripInvalidXmlChars } from '../utils/sanitize.js';
import { heldColumns } from '../utils/columnUtils.js';
import { onDocumentPage } from '../utils/pageLimits.js';
import { ADMONITION_COLOR, decodeBase64, embedUrl, encUrl, fillSheetRowGaps, hexColor, isHeaderRow, furnitureKind, lengthToPt, MIME_EXT, odfUserDefinedXml, resolvePaper, resolveZipInstant, runningFurniture, sectionFurniture, sectionFurnitureIndex, sniffImageSize, toBookmarkNameRaw, toW3CDTF, writableSpacing } from '../utils/officeGenUtils.js';
import { markerPattern, MarkerPattern, odfNumFormat } from '../utils/listMarkers.js';
import { BaseGenerator } from './BaseGenerator.js';
import { lookupTable } from '../utils/lookupUtils.js';

/**
 * The full ODF namespace set. Declared unconditionally on every part root (content.xml, styles.xml,
 * meta.xml) so no render path can emit a prefix the root never declared - the malformed-XML class
 * that broke the DOCX generator when a note carried a link or a header carried an image.
 */
const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
    + ' xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"'
    + ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
    + ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
    + ' xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"'
    + ' xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"'
    + ' xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"'
    + ' xmlns:xlink="http://www.w3.org/1999/xlink"'
    + ' xmlns:dc="http://purl.org/dc/elements/1.1/"'
    + ' xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"';

const ODT_MIMETYPE = 'application/vnd.oasis.opendocument.text';

/** Formats a number with trimmed 4-decimal precision, deterministic across platforms. */
function fmtNum(n: number): string { return (Math.round(n * 10000) / 10000).toString(); }
function fmtPt(n: number): string { return `${fmtNum(n)}pt`; }
function fmtIn(n: number): string { return `${fmtNum(n)}in`; }

/** Escapes text/attribute content after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string { return escapeXml(stripInvalidXmlChars(s ?? '')); }

/**
 * Collapses a concatenated ` name="value"` attribute string so no name appears twice - a duplicate
 * attribute (e.g. `fo:margin-left` contributed by both an admonition border and a paragraph indent)
 * is malformed XML. Last value wins; earliest position is kept, so output is deterministic per input.
 */
function dedupeXmlAttrs(s: string): string {
    if (!s) return '';
    const map = new Map<string, string>();
    for (const m of s.matchAll(/\s([\w:-]+)="([^"]*)"/g)) map.set(m[1], m[2]);
    let out = '';
    for (const [k, v] of map) out += ` ${k}="${v}"`;
    return out;
}

/** Quotes a font family for `svg:font-family` when it contains spaces. */
function quoteFontFamily(name: string): string { return /\s/.test(name) ? `'${name}'` : name; }

/**
 * Encodes a run of document text into ODF inline markup. ODF processors collapse space runs, strip
 * leading/trailing spaces, and ignore literal tab/newline in text, so those must become dedicated
 * elements (`text:s`, `text:tab`, `text:line-break`). Literal segments are XML-escaped; the encoder
 * escapes before wrapping so a literal `<text:s/>` in the source survives as text.
 */
function encodeOdfText(s: string | undefined): string {
    const norm = stripInvalidXmlChars(s ?? '').replace(/\r\n?/g, '\n');
    let out = '';
    let atSegStart = true;
    let i = 0;
    while (i < norm.length) {
        const ch = norm[i];
        if (ch === '\t') { out += '<text:tab/>'; atSegStart = true; i++; continue; }
        if (ch === '\n') { out += '<text:line-break/>'; atSegStart = true; i++; continue; }
        if (ch === ' ') {
            let k = 0; while (i + k < norm.length && norm[i + k] === ' ') k++;
            if (atSegStart) out += `<text:s text:c="${k}"/>`;
            else { out += ' '; if (k > 1) out += `<text:s text:c="${k - 1}"/>`; }
            atSegStart = false; i += k; continue;
        }
        let j = i; while (j < norm.length && norm[j] !== ' ' && norm[j] !== '\t' && norm[j] !== '\n') j++;
        out += escapeXml(norm.slice(i, j));
        atSegStart = false; i = j;
    }
    return out;
}

type StyleFamily = 'text' | 'paragraph' | 'table' | 'table-cell' | 'section';

/**
 * Interns automatic styles for one output file (content.xml or styles.xml). ODF puts run/paragraph
 * formatting into named `<style:style>` definitions and references them by name, so a bundle of
 * `TextFormatting` becomes a canonical property string that is itself the dedupe key: 500 identical
 * bold runs share one `T1`. Names are minted in first-seen (render) order, so output is byte-stable.
 * List styles and referenced fonts accumulate here too and serialize with the automatic styles.
 */
class StyleRegistry {
    private byKey = new Map<string, string>();
    private serialized: string[] = [];
    private counters: Record<string, number> = {};
    private listStyles: string[] = [];
    private fonts = new Set<string>();

    /** Each automatic style by its name: its family, properties and attributes (see derive). */
    private byName = new Map<string, { family: StyleFamily; props: string; extraAttrs: string }>();
    intern(family: StyleFamily, prefix: string, props: string, extraAttrs = ''): string {
        const key = `${family}::${extraAttrs}::${props}`;
        const existing = this.byKey.get(key);
        if (existing) return existing;
        const name = `${prefix}${(this.counters[prefix] = (this.counters[prefix] ?? 0) + 1)}`;
        this.byKey.set(key, name);
        this.byName.set(name, { family, props, extraAttrs });
        this.serialized.push(`<style:style style:name="${name}" style:family="${family}"${extraAttrs}>${props}</style:style>`);
        return name;
    }
    /**
     * A style as `name` is (an automatic style's properties and attributes; a common one, or none, as its
     * parent) with `attrs` more on its `element` properties and `extra` more on it, each in place of its
     * own of the name: an automatic style cannot be a parent.
     */
    derive(name: string, family: StyleFamily, prefix: string, element: string, attrs: Record<string, string>, extra: Record<string, string>): string {
        const xmlOf = (given: Record<string, string>) => Object.entries(given).map(([k, v]) => ` ${k}="${v}"`).join('');
        const without = (xml: string, keys: string[]) => keys.reduce((x, k) => x.split(new RegExp(`\\s${k}="[^"]*"`)).join(''), xml);
        const own = name ? this.byName.get(name) : undefined;
        if (own && own.family === family) {
            const open = new RegExp(`<${element}(?=[\\s/>])([^>]*?)(/?)>`);
            const props = open.test(own.props)
                ? own.props.replace(open, (_m, rest: string, close: string) => `<${element}${xmlOf(attrs)}${without(rest, Object.keys(attrs))}${close}>`)
                : `<${element}${xmlOf(attrs)}/>${own.props}`;
            return this.intern(family, prefix, props, without(own.extraAttrs, Object.keys(extra)) + xmlOf(extra));
        }
        const parent = name || (family === 'paragraph' ? 'Standard' : '');
        return this.intern(family, prefix, `<${element}${xmlOf(attrs)}/>`, `${parent ? ` style:parent-style-name="${parent}"` : ''}${xmlOf(extra)}`);
    }
    internListStyle(xml: string): void { this.listStyles.push(xml); }
    useFont(name: string): void { this.fonts.add(name); }
    automaticStylesXml(): string { return this.serialized.join('') + this.listStyles.join(''); }
    fontFaceDeclsXml(): string {
        return [...this.fonts].map(f => `<style:font-face style:name="${xmlText(f)}" svg:font-family="${xmlText(quoteFontFamily(f))}"/>`).join('');
    }
}

interface MediaPart { name: string; bytes: Uint8Array; ext: string; contentType: string; }
interface ListDef { name: string; registry: StyleRegistry; levels: Map<number, MarkerPattern>; startAt: Map<number, number>; }

/**
 * Generates an OpenDocument Text (`.odt`) document from any AST, as a `Uint8Array`.
 *
 * Hand-built with `fflate.zipSync` (the `EpubGenerator`/`DocxGenerator` approach), no new
 * dependencies, browser-safe, and reproducible (pinned zip mtimes, deterministic style/id names).
 * It emits the ODF dialect `OpenOfficeParser` reads back, so a generated document round-trips through
 * the parser with near-zero loss. Unlike DOCX, ODF carries formatting in named automatic styles and
 * keeps notes/comments/headers inline in the same files, so there is no per-part relationship layer.
 */
/** A section's master page: its page, and its header and footer for every page, its first page and even pages (rendered once). */
interface SectionPage {
    name: string; setup: PageSetup | undefined; key: string; numbering?: PageNumberFormat;
    header?: OfficeContentNode[]; footer?: OfficeContentNode[]; firstHeader?: OfficeContentNode[]; firstFooter?: OfficeContentNode[]; evenHeader?: OfficeContentNode[]; evenFooter?: OfficeContentNode[];
    headerXml?: string; footerXml?: string; firstHeaderXml?: string; firstFooterXml?: string; evenHeaderXml?: string; evenFooterXml?: string;
}

export class OdtGenerator extends BaseGenerator<'odt'> {
    private contentStyles = new StyleRegistry();
    private stylesFileStyles = new StyleRegistry();
    private activeStyles = this.contentStyles;
    private paraContextAttrs = '';

    private media: MediaPart[] = [];
    private mediaByAttachment = new Map<string, string>();
    /** Each attachment's intrinsic size, measured when it was first decoded, and those that could not be decoded. */
    private readonly mediaIntrinsic = new Map<string, { w: number; h: number } | null>();
    private readonly mediaUndecodable = new Set<string>();
    private manifestMedia = new Map<string, string>(); // Pictures/xxx -> contentType

    private listDefs = new Map<string, ListDef>();
    private listCounter = 0;
    private tableCounter = 0;
    private frameCounter = 0;
    private commentCounter = 0;
    private footnoteOrd = 0;
    /** Where each note was written: its id and number, in each part (by the part's style registry) it is written in. */
    private readonly writtenNotes = new Map<OfficeContentNode, { id: string; ord: number; cls: 'footnote' | 'endnote'; part: StyleRegistry }[]>();
    private endnoteOrd = 0;

    private usedBookmarkNames = new UniqueNames();
    private mathWarned = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'odt'>) {
        super('odt', ast, config);
    }

    // ── entry point ─────────────────────────────────────────────────────────────

    async generate(): Promise<ConversionResult<'odt'>> {
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);

        const metaBlock = this.config.renderMetadata ? this.renderMetadataBlock() : '';
        let body = await this.renderBlocks(this.ast.content);
        this.reportUnshownSectionPart();
        if (this.openSection) { body += '</text:section>'; this.openSection = false; }
        // (A section of columns opened by a break with nothing after it is none.)
        body = body.replace(/<text:section text:style-name="[^"]*" text:name="[^"]*"><\/text:section>$/, '');

        const headers = runningFurniture(this.ast.auxiliary?.headers, 'header'), footers = runningFurniture(this.ast.auxiliary?.footers, 'footer');
        const hasHeader = headers.length > 0;
        const hasFooter = footers.length > 0;
        this.activeStyles = this.stylesFileStyles;
        const headerXml = hasHeader ? await this.renderBlocks(headers) : '';
        const footerXml = hasFooter ? await this.renderBlocks(footers) : '';
        for (const page of this.sectionPages) {
            if (page.header) page.headerXml = await this.renderBlocks(page.header);
            if (page.footer) page.footerXml = await this.renderBlocks(page.footer);
            if (page.firstHeader) page.firstHeaderXml = page.firstHeader.length ? await this.renderBlocks(page.firstHeader) : '';
            if (page.firstFooter) page.firstFooterXml = page.firstFooter.length ? await this.renderBlocks(page.firstFooter) : '';
            if (page.evenHeader) page.evenHeaderXml = await this.renderBlocks(page.evenHeader);
            if (page.evenFooter) page.evenFooterXml = await this.renderBlocks(page.evenFooter);
        }
        // The first section's own first page's and even pages' parts, beside the running ones (see buildStylesXml).
        const bodyOwn: Record<'firstHeader' | 'firstFooter' | 'evenHeader' | 'evenFooter', string | undefined> = { firstHeader: undefined, firstFooter: undefined, evenHeader: undefined, evenFooter: undefined };
        for (const [key, kind, type] of [['firstHeader', 'header', 'first'], ['firstFooter', 'footer', 'first'], ['evenHeader', 'header', 'even'], ['evenFooter', 'footer', 'even']] as const) {
            const own = sectionFurniture(kind === 'header' ? this.ast.auxiliary?.headers : this.ast.auxiliary?.footers, kind, 0, type);
            if (own) bodyOwn[key] = await this.renderBlocks(own);
        }
        this.bodyOwnParts = bodyOwn;
        this.activeStyles = this.contentStyles;

        // Serialize list styles into their owning file's registry (learned during render).
        for (const def of this.listDefs.values()) def.registry.internListStyle(this.buildListStyleXml(def));

        const contentXml = this.buildContentXml(metaBlock + body);
        const stylesXml = this.buildStylesXml(headerXml, footerXml, hasHeader ? furnitureKind(headers, 'header') : 'default', hasFooter ? furnitureKind(footers, 'footer') : 'default');
        const metaXml = this.buildMetaXml(iso);

        const parts: Record<string, Uint8Array> = {};
        const enc = new TextEncoder();
        const manifestEntries = new Map<string, string>(); // path -> mime
        parts['content.xml'] = enc.encode(contentXml); manifestEntries.set('content.xml', 'text/xml');
        parts['styles.xml'] = enc.encode(stylesXml); manifestEntries.set('styles.xml', 'text/xml');
        parts['meta.xml'] = enc.encode(metaXml); manifestEntries.set('meta.xml', 'text/xml');
        for (const m of this.media) { parts[`Pictures/${m.name}`] = m.bytes; manifestEntries.set(`Pictures/${m.name}`, m.contentType); }
        const manifestXml = this.buildManifest(manifestEntries);

        // mimetype must be the FIRST entry and STORED (uncompressed); fflate preserves insertion order.
        const zipFiles: Zippable = {
            mimetype: [enc.encode(ODT_MIMETYPE), { level: 0 }],
            'META-INF/manifest.xml': enc.encode(manifestXml),
        };
        for (const [path, data] of Object.entries(parts)) zipFiles[path] = data;

        const bytes = zipSync(zipFiles, { mtime });
        return { value: bytes, messages: this.messages };
    }

    // ── block / inline traversal ─────────────────────────────────────────────────

    private async renderBlocks(nodes: OfficeContentNode[] | undefined): Promise<string> {
        let out = '';
        let prevPaginated: string | null = null;
        const items = nodes || [];
        // Where the run of inline nodes last found without text ends (see below), so it is looked at once.
        let inlineWithoutTextUntil = -1;
        for (let idx = 0; idx < items.length; idx++) {
            const node = items[idx];
            checkAbortSignal(this.config.abortSignal);
            // Inline nodes among the blocks (a cell's, a note's or a comment's runs, as ODS and several
            // parsers give them) are one paragraph: each was a paragraph of its own, so a cell reading
            // "Total: 5 units" with the 5 in bold was three lines. A run of pictures or breaks with no
            // text stays blocks.
            if (idx >= inlineWithoutTextUntil && isInlineNode(node)) {
                let end = idx;
                let hasText = false;
                while (end < items.length && isInlineNode(items[end])) {
                    if (items[end].type === 'text') hasText = true;
                    end++;
                }
                if (hasText) {
                    const waiting = this.takeBreak();
                    out += this.carryBreak(await this.paragraph({ type: 'paragraph', children: items.slice(idx, end) } as OfficeContentNode), waiting);
                    prevPaginated = null;
                    idx = end - 1;
                    continue;
                }
                inlineWithoutTextUntil = end;
            }
            if (node.type === 'list') {
                // Consume the maximal run of consecutive sibling lists sharing a listId, so the parser
                // rejoins them (it keys the logical list off the shared list style name). handleOnNode
                // is applied per item inside renderListRun, so every item still gets the callback.
                const listId = String((node.metadata as any)?.listId ?? 'odt-list');
                const run: OfficeContentNode[] = [];
                while (idx < items.length && items[idx].type === 'list'
                    && String((items[idx].metadata as any)?.listId ?? 'odt-list') === listId) {
                    run.push(items[idx]); idx++;
                }
                idx--;
                const waiting = this.takeBreak();
                out += this.carryBreak(await this.renderListRun(run), waiting);
                prevPaginated = null;
                continue;
            }
            const override = await this.handleOnNode(node);
            if (override === false) { prevPaginated = null; continue; }
            if (typeof override === 'string') { out += override; prevPaginated = null; continue; }
            if ((node.type === 'page' || node.type === 'slide') && prevPaginated === node.type) out += this.waitBreak({ before: 'page' });
            // (A break waits for the block after it, which carries it: see waitBreak.)
            const waiting = node.type === 'break' ? undefined : this.takeBreak();
            const block = await this.renderBlockNode(node);
            // Paragraphs and headings write their own bookmarks; every other block's go in its first paragraph.
            out += this.carryBreak(OWN_BOOKMARKS.has(node.type) ? block : this.withBookmarks(block, node), waiting);
            prevPaginated = (node.type === 'page' || node.type === 'slide') ? node.type : null;
        }
        // (A break that ends the blocks: a paragraph of its own carries it.)
        out += this.waitBreak(undefined);
        return out;
    }

    /**
     * A page or column break (and the master page a section turns to) waiting for the block after it:
     * ODF's breaks are its first paragraph's, heading's or table's style's (`fo:break-before`,
     * `style:master-page-name`). Written as an empty paragraph of their own, they were one more empty
     * paragraph on every save. Returns the break that was waiting, as a paragraph of its own (two breaks
     * in a row: the first one's page is empty).
     */
    private pendingBreak: { before: 'page' | 'column'; master?: string } | undefined;
    private waitBreak(next: { before: 'page' | 'column'; master?: string } | undefined): string {
        const before = this.pendingBreak;
        this.pendingBreak = next;
        return before ? this.breakParagraph(before) : '';
    }
    /** A break as a paragraph of its own (one that ends the blocks, or the first of two in a row). */
    private breakParagraph(brk: { before: 'page' | 'column'; master?: string }): string {
        return `<text:p text:style-name="${this.breakStyle('', 'paragraph', brk)}"/>`;
    }
    private takeBreak(): { before: 'page' | 'column'; master?: string } | undefined {
        const waiting = this.pendingBreak;
        this.pendingBreak = undefined;
        return waiting;
    }
    /** A style as `name` is, the break before it (see waitBreak). */
    private breakStyle(name: string, family: 'paragraph' | 'table', brk: { before: 'page' | 'column'; master?: string }): string {
        return this.activeStyles.derive(name, family, family === 'table' ? 'TB' : 'P', family === 'table' ? 'style:table-properties' : 'style:paragraph-properties',
            { 'fo:break-before': brk.before }, brk.master ? { 'style:master-page-name': brk.master } : {});
    }
    /** `xml` with the break that waited in its first paragraph, heading or table; it waits on where it has none. */
    private carryBreak(xml: string, waiting: { before: 'page' | 'column'; master?: string } | undefined): string {
        if (!waiting) return xml;
        const tag = /<(text:p|text:h|table:table)(?=[\s/>])([^>]*?)(\/?)>/.exec(xml);
        if (!tag) {
            // (The block wrote no paragraph: the break waits on, unless it set one of its own.)
            if (this.pendingBreak) return this.breakParagraph(waiting) + xml;
            this.pendingBreak = waiting;
            return xml;
        }
        const attr = tag[1] === 'table:table' ? 'table:style-name' : 'text:style-name';
        const own = new RegExp(`\\s${attr}="([^"]*)"`).exec(tag[2]);
        const style = this.breakStyle(own?.[1] ?? '', tag[1] === 'table:table' ? 'table' : 'paragraph', waiting);
        const attrs = own ? tag[2].replace(own[0], ` ${attr}="${style}"`) : `${tag[2]} ${attr}="${style}"`;
        return `${xml.slice(0, tag.index)}<${tag[1]}${attrs}${tag[3]}>${xml.slice(tag.index + tag[0].length)}`;
    }


    private async renderBlockNode(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'paragraph': return this.paragraph(node);
            case 'heading': return this.heading(node);
            case 'list': return this.renderListRun([node]);
            case 'table': return this.table(node);
            case 'image': return this.blockImage(node);
            case 'code': return this.codeBlock(node);
            case 'break': return this.blockBreak(node);
            case 'note': return this.standaloneNote(node);
            case 'comment': return this.standaloneComment(node);
            case 'admonition': return this.admonition(node);
            case 'chart': return this.chart(node);
            case 'embed': return this.embed(node);
            case 'definitionTerm': return this.paragraph(node, { extraAttrs: '', named: undefined, boldTerm: true });
            case 'definitionDescription': return this.paragraph(node, { extraAttrs: ' fo:margin-left="0.5in"' });
            case 'header': case 'footer': return this.renderBlocks(node.children);
            case 'sheet': return this.sheet(node);
            case 'row': case 'cell': case 'text':
                return `<text:p>${await this.renderInline([node])}</text:p>`;
            case 'page': case 'slide': case 'drawing': case 'definitionList':
                return this.renderBlocks(node.children);
            case 'slideMaster': return ''; // template furniture, dropped
            default: {
                const _exhaustive: never = node;
                void _exhaustive;
                return '';
            }
        }
    }

    /** Body blocks of a note/comment: its children, or a synthesized paragraph from its `text`. */
    private bodyBlocks(node: OfficeContentNode): OfficeContentNode[] {
        if (node.children && node.children.length) return node.children;
        return [{ type: 'paragraph', text: node.text || '' } as OfficeContentNode];
    }

    // ── paragraph family ─────────────────────────────────────────────────────────

    private async paragraph(node: OfficeContentNode, opts: { named?: string; extraAttrs?: string; boldTerm?: boolean } = {}): Promise<string> {
        const meta = node.metadata as any;
        const semantic = this.getSemanticMapping(node);
        const named = opts.named ?? this.styleForTag(semantic?.tag) ?? this.knownParaStyle(meta?.style);
        const attrs = (opts.extraAttrs ?? '') + this.paraContextAttrs + this.alignIndentAttrs(meta);
        const styleName = this.internPara(attrs, named);
        const styleAttr = styleName ? ` text:style-name="${styleName}"` : '';
        const bookmarks = this.bookmarksFor(node);
        const blockRefs = (await this.notesFor(node)) + (await this.commentsFor(node));
        let inner = await this.renderInline(node.children && node.children.length ? node.children : [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        if (opts.boldTerm) inner = `<text:span text:style-name="${this.boldStyle()}">${inner}</text:span>`;
        return `<text:p${styleAttr}>${bookmarks}${blockRefs}${inner}</text:p>`;
    }

    private async heading(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const level = Math.min(6, Math.max(1, meta?.level || 1));
        const attrs = this.paraContextAttrs + this.alignIndentAttrs(meta);
        const styleName = this.internPara(attrs, `Heading_20_${level}`);
        const bookmarks = this.bookmarksFor(node);
        const blockRefs = (await this.notesFor(node)) + (await this.commentsFor(node));
        const inner = await this.renderInline(node.children && node.children.length ? node.children : [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        return `<text:h text:outline-level="${level}" text:style-name="${styleName}">${bookmarks}${blockRefs}${inner}</text:h>`;
    }

    /** Interns an automatic paragraph style from an attribute string, optionally parented to a named style. */
    private internPara(propsAttrs: string, parentName?: string): string {
        const deduped = dedupeXmlAttrs(propsAttrs);
        if (!deduped && !parentName) return '';
        if (!deduped && parentName) return parentName;
        const props = `<style:paragraph-properties${deduped}/>`;
        const extra = parentName ? ` style:parent-style-name="${parentName}"` : '';
        return this.activeStyles.intern('paragraph', 'P', props, extra);
    }

    private alignIndentAttrs(meta: any): string {
        if (this.config.includeFormatting === false || !meta) return '';
        let a = '';
        const align = this.textAlign(meta.alignment);
        if (align) a += ` fo:text-align="${align}"`;
        const ind = meta.paragraphIndentation;
        if (ind) {
            if (ind.left) a += ` fo:margin-left="${fmtPt(ind.left / 20)}"`;
            if (ind.right) a += ` fo:margin-right="${fmtPt(ind.right / 20)}"`;
            if (ind.firstLine) a += ` fo:text-indent="${fmtPt(ind.firstLine / 20)}"`;
            else if (ind.hanging) a += ` fo:text-indent="${fmtPt(-ind.hanging / 20)}"`;
        }
        const spacing = writableSpacing(meta);
        if (spacing) {
            if (spacing.before !== undefined) a += ` fo:margin-top="${fmtPt(spacing.before)}"`;
            if (spacing.after !== undefined) a += ` fo:margin-bottom="${fmtPt(spacing.after)}"`;
            if (spacing.lineMultiple) a += ` fo:line-height="${fmtNum(spacing.lineMultiple * 100)}%"`;
            else if (spacing.lineHeight) a += spacing.lineAtLeast ? ` style:line-height-at-least="${fmtPt(spacing.lineHeight)}"` : ` fo:line-height="${fmtPt(spacing.lineHeight)}"`;
        }
        return a;
    }

    private textAlign(alignment: string | undefined): string | null {
        if (alignment === 'center') return 'center';
        if (alignment === 'right') return 'end';
        if (alignment === 'justify') return 'justify';
        return null;
    }

    private bookmarksFor(node: OfficeContentNode): string {
        if (this.config.ignoreInternalLinks) return '';
        const names: string[] = [...(((node.metadata as any)?.anchorIds) || [])];
        if (this.config.generateIds && node.type === 'heading') {
            const slug = this.slugify(node.text || this.getNodeText(node));
            if (slug && !names.includes(slug)) names.push(slug);
        }
        let out = '';
        for (const raw of names) out += `<text:bookmark text:name="${xmlText(this.mintBookmark(raw))}"/>`;
        return out;
    }

    /**
     * A block's bookmarks (see bookmarksFor) in the first paragraph it writes: a table's in its first
     * cell's, a picture's, code block's or chart's in its own. ODF has bookmarks in paragraphs only, and
     * only paragraphs and headings wrote theirs, so an internal link to a table, a picture, an equation
     * or a list item (LaTeX's `\ref{tab:main}` among them) had no target. A block writing no paragraph
     * has a paragraph of its own for them.
     */
    private withBookmarks(xml: string, node: OfficeContentNode): string {
        const marks = this.bookmarksFor(node);
        if (!marks) return xml;
        for (let at = xml.indexOf('<text:'); at >= 0; at = xml.indexOf('<text:', at + 1)) {
            const tag = xml[at + 6];
            if ((tag !== 'p' && tag !== 'h') || !/[\s/>]/.test(xml[at + 7] ?? '')) continue;
            const end = xml.indexOf('>', at);
            if (end < 0) break;
            return xml[end - 1] === '/'
                ? `${xml.slice(0, end - 1)}>${marks}</text:${tag}>${xml.slice(end + 1)}`
                : xml.slice(0, end + 1) + marks + xml.slice(end + 1);
        }
        return `<text:p>${marks}</text:p>${xml}`;
    }

    /** A node in a line with bookmarks of its own: what it writes between each bookmark's start and end. */
    private withInlineBookmarks(xml: string, node: OfficeContentNode): string {
        const raw = (node.metadata as any)?.anchorIds;
        if (this.config.ignoreInternalLinks || !Array.isArray(raw) || raw.length === 0) return xml;
        let start = '';
        let end = '';
        for (const name of raw) {
            const minted = xmlText(this.mintBookmark(String(name)));
            start += `<text:bookmark-start text:name="${minted}"/>`;
            end = `<text:bookmark-end text:name="${minted}"/>${end}`;
        }
        return start + xml + end;
    }

    /** Mints a unique bookmark name; the first claim of a base keeps it, later claims get `_2`. */
    private mintBookmark(rawName: string): string {
        const base = toBookmarkNameRaw(rawName);
        return this.usedBookmarkNames.claim(base, n => base.slice(0, 40 - `_${n}`.length) + `_${n}`);
    }
    private anchorName(rawName: string): string { return toBookmarkNameRaw(rawName); }

    // ── inline runs ──────────────────────────────────────────────────────────────

    private async renderInline(nodes: OfficeContentNode[]): Promise<string> {
        let out = '';
        let group: OfficeContentNode[] = [];
        let groupLink: any = null;
        let groupType: any = null;
        const flush = async () => { if (group.length) { out += await this.hyperlink(groupLink, groupType, group); group = []; groupLink = null; } };
        for (const node of nodes || []) {
            const override = await this.handleOnNode(node);
            if (override === false) continue;
            if (typeof override === 'string') { await flush(); out += override; continue; }
            const link = node.type === 'text' ? (node.metadata as any)?.link : undefined;
            if (link && node.type === 'text') {
                if (group.length && groupLink === link) group.push(node);
                else { await flush(); group = [node]; groupLink = link; groupType = (node.metadata as any)?.linkType; }
                continue;
            }
            await flush();
            out += await this.inlineNode(node);
        }
        await flush();
        return out;
    }

    private async inlineNode(node: OfficeContentNode): Promise<string> {
        return this.withInlineBookmarks(await this.inlineContent(node), node);
    }

    private async inlineContent(node: OfficeContentNode): Promise<string> {
        switch (node.type) {
            case 'text': return this.span(node.text || '', node.formatting) + await this.notesFor(node) + await this.commentsFor(node);
            case 'code': return this.inlineCode(node);
            case 'image': return this.imageRun(node);
            case 'break': return this.inlineBreak(node);
            case 'paragraph': case 'heading':
                return this.renderInline(node.children || [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
            default:
                return this.span(this.getNodeText(node), undefined);
        }
    }

    private span(text: string, fmt: TextFormatting | undefined): string {
        const enc = encodeOdfText(text);
        if (!enc) return '';
        const props = this.textProps(fmt);
        if (!props) return enc;
        const name = this.activeStyles.intern('text', 'T', props);
        return `<text:span text:style-name="${name}">${enc}</text:span>`;
    }

    private textProps(fmt: TextFormatting | undefined): string {
        if (!fmt || this.config.includeFormatting === false) return '';
        let a = '';
        if (fmt.bold) a += ' fo:font-weight="bold" style:font-weight-asian="bold"';
        if (fmt.italic) a += ' fo:font-style="italic" style:font-style-asian="italic"';
        if (fmt.underline) a += ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"';
        if (fmt.strikethrough) a += ' style:text-line-through-style="solid"';
        const c = hexColor(fmt.color); if (c) a += ` fo:color="#${c}"`;
        const bg = hexColor(fmt.backgroundColor); if (bg) a += ` fo:background-color="#${bg}"`;
        const sz = lengthToPt(fmt.size); const pt = sz && sz > 0 ? Math.round(sz * 10) / 10 : 0;
        if (pt > 0) a += ` fo:font-size="${fmtPt(pt)}" style:font-size-asian="${fmtPt(pt)}"`;
        if (fmt.font) { this.activeStyles.useFont(fmt.font); a += ` style:font-name="${xmlText(fmt.font)}"`; }
        if (fmt.superscript) a += ' style:text-position="super 58%"';
        else if (fmt.subscript) a += ' style:text-position="sub 58%"';
        return a ? `<style:text-properties${a}/>` : '';
    }

    private boldStyle(): string {
        return this.activeStyles.intern('text', 'T', '<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold"/>');
    }

    private async hyperlink(link: string, linkType: string | undefined, group: OfficeContentNode[]): Promise<string> {
        const spans = group.map(n => this.withInlineBookmarks(this.span(n.text || '', n.formatting), n)).join('');
        let trailing = '';
        for (const n of group) trailing += (await this.notesFor(n)) + (await this.commentsFor(n));
        const href = this.linkHref(link, linkType);
        return (href ? `<text:a xlink:type="simple" xlink:href="${href}">${spans}</text:a>` : spans) + trailing;
    }

    /**
     * The escaped `xlink:href` a link is written with: `#` and the bookmark name for an internal
     * target, the encoded URL for an external one. Empty when the link is not written: an internal
     * one under `ignoreInternalLinks`, or a target the scheme check refuses.
     */
    private linkHref(link: string, linkType: string | undefined): string {
        const internal = linkType === 'internal' || link.startsWith('#');
        if (internal) return this.config.ignoreInternalLinks ? '' : `#${xmlText(this.anchorName(link.replace(/^#/, '')))}`;
        const safe = sanitizeOfficePackageUrl(link);
        return safe ? xmlText(encUrl(safe)) : '';
    }

    private inlineCode(node: OfficeContentNode): string {
        const name = this.activeStyles.intern('text', 'T', '<style:text-properties style:font-name="Courier New"/>');
        this.activeStyles.useFont('Courier New');
        return `<text:span text:style-name="${name}">${encodeOdfText(node.text || '')}</text:span>`;
    }

    private inlineBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'section') return '';
        return '<text:line-break/>';
    }

    // ── notes & comments (inline, in content.xml) ─────────────────────────────────

    private async notesFor(node: OfficeContentNode): Promise<string> {
        if (!node.notes || !node.notes.length) return '';
        let out = '';
        for (const note of node.notes) {
            const cls: 'footnote' | 'endnote' = (note.metadata as any)?.noteType === 'endnote' ? 'endnote' : 'footnote';
            out += await this.note(note, cls);
        }
        return out;
    }

    private async note(node: OfficeContentNode, cls: 'footnote' | 'endnote'): Promise<string> {
        // A note referred to again is a reference to it (its number), its text written once: written at
        // every reference, one note a small document refers to thousands of times was copied that often.
        // Within the part it was written in only (content.xml, or styles.xml for the header and footer):
        // a reference cannot reach a note in the other part, so there it is written again, once, with
        // the same number.
        const written = this.writtenNotes.get(node);
        const writtenHere = written?.find(w => w.part === this.activeStyles);
        if (writtenHere) return `<text:note-ref text:note-class="${writtenHere.cls}" text:reference-format="text" text:ref-name="${writtenHere.id}">${writtenHere.ord}</text:note-ref>`;
        const ord = written ? written[0].ord : cls === 'footnote' ? ++this.footnoteOrd : ++this.endnoteOrd;
        const id = `${cls === 'footnote' ? 'ftn' : 'edn'}${ord}${written ? `_${written.length + 1}` : ''}`;
        const entry = { id, ord, cls, part: this.activeStyles };
        if (written) written.push(entry);
        else this.writtenNotes.set(node, [entry]);
        const body = this.styleNoteBody((await this.renderBlocks(this.bodyBlocks(node))) || '<text:p/>');
        return `<text:note text:id="${id}" text:note-class="${cls}"><text:note-citation>${ord}</text:note-citation><text:note-body>${body}</text:note-body></text:note>`;
    }

    private styleNoteBody(body: string): string {
        return body.replace(/<text:p>/g, '<text:p text:style-name="Footnote">');
    }

    private async standaloneNote(node: OfficeContentNode): Promise<string> {
        const cls: 'footnote' | 'endnote' = (node.metadata as any)?.noteType === 'endnote' ? 'endnote' : 'footnote';
        return `<text:p>${await this.note(node, cls)}</text:p>`;
    }

    private async commentsFor(node: OfficeContentNode): Promise<string> {
        if (!node.comments || !node.comments.length) return '';
        let out = '';
        for (const c of node.comments) out += await this.annotation(c);
        return out;
    }

    private async standaloneComment(node: OfficeContentNode): Promise<string> {
        const annotation = await this.annotation(node);
        return annotation ? `<text:p>${annotation}</text:p>` : '';
    }

    private async annotation(node: OfficeContentNode): Promise<string> {
        // Written once, at its first reference (see firstWriteOfComment).
        if (!this.firstWriteOfComment(node)) return '';
        const meta = node.metadata as any;
        const name = `cmt${++this.commentCounter}`;
        const creator = meta?.author ? `<dc:creator>${xmlText(meta.author)}</dc:creator>` : '';
        const date = toW3CDTF(meta?.date);
        const dateEl = date ? `<dc:date>${date}</dc:date>` : '';
        const body = (await this.renderBlocks(this.bodyBlocks(node))) || '<text:p/>';
        return `<office:annotation office:name="${name}">${creator}${dateEl}${body}</office:annotation>`;
    }

    // ── lists ─────────────────────────────────────────────────────────────────────

    private ensureListStyle(listId: string): string {
        let def = this.listDefs.get(listId);
        if (!def) {
            def = { name: `L${++this.listCounter}`, registry: this.activeStyles, levels: new Map(), startAt: new Map() };
            this.listDefs.set(listId, def);
        }
        return def.name;
    }

    private registerListLevel(listId: string, lvl: number, meta: any): void {
        const def = this.listDefs.get(listId)!;
        if (!def.levels.has(lvl)) {
            // The level numbers or bullets as its first item's marker does (`a)`, `1.1.`, `▪`).
            def.levels.set(lvl, markerPattern(meta, lvl));
            if (typeof meta?.itemIndex === 'number') def.startAt.set(lvl, meta.itemIndex + 1);
        }
    }

    private async renderListRun(items: OfficeContentNode[]): Promise<string> {
        const listId = String((items[0].metadata as any)?.listId ?? 'odt-list');
        const styleName = this.ensureListStyle(listId);
        const firstIdx = (items[0].metadata as any)?.itemIndex;
        const continueAttr = (typeof firstIdx === 'number' && firstIdx > 0) ? ' text:continue-numbering="true"' : '';
        let out = '';
        let depth = -1;
        for (const item of items) {
            const override = await this.handleOnNode(item);
            if (override === false) continue; // omit this item; the list continues
            const meta = item.metadata as any;
            const lvl = Math.max(0, Math.min(9, meta?.indentation | 0));
            this.registerListLevel(listId, lvl, meta);
            if (depth < 0) {
                out += `<text:list text:style-name="${styleName}"${continueAttr}><text:list-item>`;
                for (let d = 1; d <= lvl; d++) out += '<text:list><text:list-item>';
                depth = lvl;
            } else if (lvl > depth) {
                for (let d = depth + 1; d <= lvl; d++) out += '<text:list><text:list-item>';
                depth = lvl;
            } else if (lvl === depth) {
                out += '</text:list-item><text:list-item>';
            } else {
                for (let d = depth; d > lvl; d--) out += '</text:list-item></text:list>';
                out += '</text:list-item><text:list-item>';
                depth = lvl;
            }
            if (typeof override === 'string') { out += `<text:p>${override}</text:p>`; continue; }
            let prefix = '';
            if (meta?.isTask) prefix = this.span(meta.checked ? '☑ ' : '☐ ', undefined);
            const blockRefs = (await this.notesFor(item)) + (await this.commentsFor(item));
            const inner = await this.renderInline(item.children || [{ type: 'text', text: item.text || '' } as OfficeContentNode]);
            out += `<text:p>${this.bookmarksFor(item)}${prefix}${blockRefs}${inner}</text:p>`;
        }
        for (let d = depth; d >= 0; d--) out += '</text:list-item></text:list>';
        return out;
    }

    private buildListStyleXml(def: ListDef): string {
        let levels = '';
        for (let l = 0; l <= 9; l++) {
            const type = def.levels.get(l) ?? def.levels.get(0) ?? markerPattern(undefined, 0);
            const props = `<style:list-level-properties text:space-before="${fmtIn(0.25 * l)}" text:min-label-width="0.25in"/>`;
            if (type.kind === 'ordered') {
                const start = def.startAt.get(l) ?? 1;
                // Its parents' numbers shown where those levels are numbered in this list, else as the text
                // its marker shows (a PDF's levels, each a list of its own).
                const parentsNumbered = Array.from({ length: l }, (_, k) => k).every(k => def.levels.get(k)?.kind === 'ordered');
                const literalPrefix = type.prefix + (type.withParents && !parentsNumbered ? type.parentText ?? '' : '');
                const prefix = literalPrefix ? ` style:num-prefix="${xmlText(literalPrefix)}"` : '';
                const shown = type.withParents && parentsNumbered ? ` text:display-levels="${l + 1}"` : '';
                levels += `<text:list-level-style-number text:level="${l + 1}"${prefix} style:num-format="${odfNumFormat(type.format)}" style:num-suffix="${xmlText(type.suffix)}"${shown} text:start-value="${start}">${props}</text:list-level-style-number>`;
            } else {
                // A bullet of none is a space: ODF's bullet is one character.
                levels += `<text:list-level-style-bullet text:level="${l + 1}" text:bullet-char="${xmlText([...type.bullet][0] ?? ' ')}">${props}</text:list-level-style-bullet>`;
            }
        }
        return `<text:list-style style:name="${def.name}">${levels}</text:list-style>`;
    }

    // ── tables ───────────────────────────────────────────────────────────────────

    private async table(node: OfficeContentNode): Promise<string> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return '';
        // Each row's cells, merges and gaps laid out by occupancy, as the grid budget laid them out.
        const layout = layoutTableRows(rows)!;
        const cols = layout.cols;
        const tableName = `Table${++this.tableCounter}`;
        const tableStyle = this.ensureTableStyle(this.tableAlign((node.metadata as any)?.align));
        const colXml = `<table:table-column table:number-columns-repeated="${cols}"/>`;

        const rendered: { xml: string; header: boolean }[] = [];
        for (let ri = 0; ri < rows.length; ri++) {
            const row = rows[ri];
            const header = isHeaderRow(row, ri === 0);
            let cellsXml = '';
            for (const slot of layout.rows[ri]) {
                if (slot.kind === 'covered') { cellsXml += '<table:covered-table-cell/>'.repeat(slot.span); continue; }
                // A column the row skips (a sparse grid, or before a vertical merge still pending further right).
                if (slot.kind === 'gap') { cellsXml += '<table:table-cell><text:p/></table:table-cell>'; continue; }
                const { cell, rowSpan } = slot;
                const cmeta = cell.metadata as any;
                const colSpan = Math.min(cols, slot.colSpan);
                const cellStyle = this.ensureCellStyle(hexColor(cmeta?.backgroundColor));
                let spanAttr = '';
                if (colSpan > 1) spanAttr += ` table:number-columns-spanned="${colSpan}"`;
                if (rowSpan > 1) spanAttr += ` table:number-rows-spanned="${rowSpan}"`;
                let inner = await this.renderBlocks(cell.children);
                if (!inner.trim()) inner = '<text:p/>';
                inner = this.withBookmarks(inner, cell);
                cellsXml += `<table:table-cell table:style-name="${cellStyle}"${spanAttr}>${inner}</table:table-cell>`;
                for (let k = 1; k < colSpan; k++) cellsXml += '<table:covered-table-cell/>';
            }
            if (!cellsXml) cellsXml = '<table:table-cell><text:p/></table:table-cell>';
            rendered.push({ xml: `<table:table-row>${cellsXml}</table:table-row>`, header });
        }
        let body = '';
        let i = 0;
        if (rendered.length && rendered[0].header) {
            let hx = '';
            while (i < rendered.length && rendered[i].header) { hx += rendered[i].xml; i++; }
            body += `<table:table-header-rows>${hx}</table:table-header-rows>`;
        }
        for (; i < rendered.length; i++) body += rendered[i].xml;
        return `<table:table table:name="${tableName}" table:style-name="${tableStyle}">${colXml}${body}</table:table>`;
    }

    private tableAlign(alignment: string | undefined): string | null {
        if (alignment === 'center') return 'center';
        if (alignment === 'right') return 'right';
        if (alignment === 'left') return 'left';
        return null; // 'justify' is not a valid table alignment
    }

    private ensureTableStyle(align: string | null): string {
        const props = `<style:table-properties${align ? ` table:align="${align}"` : ''}/>`;
        return this.activeStyles.intern('table', 'TB', props);
    }

    private ensureCellStyle(bg: string | null): string {
        const props = `<style:table-cell-properties fo:border="0.5pt solid #000000" fo:padding="0.0382in"${bg ? ` fo:background-color="#${bg}"` : ''}/>`;
        return this.activeStyles.intern('table-cell', 'TC', props);
    }

    private async sheet(node: OfficeContentNode): Promise<string> {
        const name = (node.metadata as any)?.sheetName;
        const heading = name ? `<text:h text:outline-level="2" text:style-name="Heading_20_2">${encodeOdfText(name)}</text:h>` : '';
        const children = node.children || [];
        const table = await this.table({ type: 'table', children: fillSheetRowGaps(children.filter(c => c.type === 'row'), n => this.takeGridPositions(n)) } as OfficeContentNode);
        // A sheet's drawing images and charts are pushed as non-row children after the rows; render
        // them after the grid (the HTML generator does the same) rather than dropping them.
        const extras = await this.renderBlocks(children.filter(c => c.type !== 'row'));
        return heading + table + extras;
    }

    // ── images ───────────────────────────────────────────────────────────────────

    private async blockImage(node: OfficeContentNode): Promise<string> {
        const align = this.textAlign((node.metadata as any)?.align);
        const styleName = this.internPara(align ? ` fo:text-align="${align}"` : '');
        const styleAttr = styleName ? ` text:style-name="${styleName}"` : '';
        return `<text:p${styleAttr}>${await this.imageRun(node)}</text:p>`;
    }

    private async imageRun(node: OfficeContentNode): Promise<string> {
        const mode: ImageMode = this.imageMode();
        if (mode === 'none') return '';
        const meta = node.metadata as any;
        const ocr = node.text || '';
        if (mode === 'ocr-text-only') return ocr ? this.span(ocr, undefined) : '';

        const media = meta?.attachmentName ? this.mediaRef(meta.attachmentName) : null;
        let frame = '';
        if (media) {
            const { w, h } = this.imageDims(node, media.intrinsic);
            const frameName = `Frame${++this.frameCounter}`;
            const title = meta?.altText ? `<svg:title>${xmlText(meta.altText)}</svg:title>` : '';
            frame = `<draw:frame draw:name="${frameName}" text:anchor-type="as-char" svg:width="${fmtPt(w)}" svg:height="${fmtPt(h)}" draw:z-index="0">`
                + `<draw:image xlink:href="${media.href}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${title}</draw:frame>`;
            // A picture that is a link is a frame inside <draw:a>.
            const href = meta?.link ? this.linkHref(meta.link, meta.linkType) : '';
            if (href) frame = `<draw:a xlink:type="simple" xlink:href="${href}">${frame}</draw:a>`;
        } else if (meta?.url) {
            // Remote-only image: a link on the alt text, to where the picture links when it is a link,
            // and to the image otherwise.
            const safe = sanitizeOfficePackageUrl(meta.url);
            const href = (meta.link ? this.linkHref(meta.link, meta.linkType) : '') || (safe ? xmlText(encUrl(safe)) : '');
            if (safe) frame = `<text:a xlink:type="simple" xlink:href="${href}">${this.span(meta.altText || safe, undefined)}</text:a>`;
        }
        if (!frame) { const fb = meta?.altText || ocr; return fb ? this.span(fb, undefined) : ''; }
        if (mode === 'image+ocr-text' && ocr) return frame + this.span('\n' + ocr, undefined);
        return frame;
    }

    private mediaRef(attachmentName: string): { href: string; intrinsic: { w: number; h: number } | null } | null {
        const att = this.getAttachment(attachmentName);
        if (!att || !att.data) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'missing attachment' }); return null; }
        const ext = MIME_EXT[(att.mimeType || '').toLowerCase()];
        if (!ext) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'unsupported mime' }); return null; }
        // Each attachment decoded and measured once, at the first picture showing it: decoded again at
        // every picture before the media was looked up, one 3 MB picture shown 20,000 times took 105 s.
        if (this.mediaUndecodable.has(attachmentName)) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName }); return null; }
        let name = this.mediaByAttachment.get(attachmentName);
        if (!name) {
            let bytes: Uint8Array;
            try { bytes = decodeBase64(att.data); } catch {
                this.mediaUndecodable.add(attachmentName);
                this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName });
                return null;
            }
            this.mediaIntrinsic.set(attachmentName, sniffImageSize(bytes));
            name = `image${this.media.length + 1}.${ext}`;
            const contentType = att.mimeType || 'image/png';
            this.media.push({ name, bytes, ext, contentType });
            this.mediaByAttachment.set(attachmentName, name);
            this.manifestMedia.set(`Pictures/${name}`, contentType);
        }
        return { href: `Pictures/${name}`, intrinsic: this.mediaIntrinsic.get(attachmentName) ?? null };
    }

    /** Image size in points, priority: explicit width -> PDF bounds -> intrinsic@96dpi -> 3x2.25in, capped to content width. */
    private imageDims(node: OfficeContentNode, intrinsic: { w: number; h: number } | null): { w: number; h: number } {
        const contentWidthPt = this.contentWidthPt();
        const meta = node.metadata as any;
        if (meta?.width) {
            const pct = /^\s*([\d.]+)\s*%\s*$/.exec(meta.width);
            let w: number | null = null;
            if (pct) w = contentWidthPt * Math.min(100, parseFloat(pct[1])) / 100;
            else { const pt = lengthToPt(meta.width); if (pt) w = pt; }
            if (w) { const ar = intrinsic ? intrinsic.h / intrinsic.w : 0.75; return this.capDims(w, w * ar, contentWidthPt); }
        }
        if (node.bounds && node.bounds.width > 0 && node.bounds.height > 0) {
            return this.capDims(node.bounds.width, node.bounds.height, contentWidthPt);
        }
        if (intrinsic) return this.capDims(intrinsic.w / 96 * 72, intrinsic.h / 96 * 72, contentWidthPt);
        return { w: 3 * 72, h: 2.25 * 72 };
    }
    private capDims(w: number, h: number, maxW: number): { w: number; h: number } {
        if (w > maxW && w > 0) { h = h * maxW / w; w = maxW; }
        return { w: Math.max(1, w), h: Math.max(1, h) };
    }

    private contentWidthPt(): number {
        const cfg = this.config.odtConfig;
        const paper = resolvePaper(cfg.format, cfg.landscape, this.ast.metadata?.pageSetup, cfg.margin);
        return Math.max(36, paper.w - paper.margin.left - paper.margin.right);
    }

    // ── breaks / code / admonition / chart / embed ────────────────────────────────


    /** Whether a section of columns of its own is open, and how many there have been. */
    private openSection = false;
    private sections = 0;
    /**
     * The pages of sections set on other pages than the first's (a landscape section, other margins), each
     * a master page of its own, with its columns; and whether a section returns to the first's (`Body`).
     */
    private sectionPages: SectionPage[] = [];
    private readonly sectionMasterByKey = new Map<string, string>();
    private readonly sectionPageByName = new Map<string, SectionPage>();
    private bodyMasterUsed = false;
    /** The first section's own first page's and even pages' parts, rendered (see generate). */
    private bodyOwnParts: Partial<Record<'firstHeader' | 'firstFooter' | 'evenHeader' | 'evenFooter', string>> = {};
    private onSectionPage = false;
    /** The section being written (counted from 0, as BreakMetadata.section counts), its page, and the master its pages are on ('' the first's). */
    private currentSection = 0;
    private currentSetup: PageSetup | undefined;
    private activeMaster = '';

    /**
     * A section's header or footer of a type: its own, else the nearest earlier section's own of that type
     * (as Word gives it), the first section's included (for a first page's or even pages', the running part
     * where it is of that type); undefined where none is, the every-page one then being the first section's.
     */
    private sectionPart(kind: 'header' | 'footer', section: number, type: 'default' | 'first' | 'even' = 'default'): { key: string; nodes: OfficeContentNode[] } | undefined {
        const nodes = kind === 'header' ? this.ast.auxiliary?.headers : this.ast.auxiliary?.footers;
        if (!nodes?.length || section < (type === 'default' ? 1 : 0)) return undefined;
        // The latest section at or before this one with its own (found by halving: a document of thousands of sections).
        const { bySection, sections } = sectionFurnitureIndex(nodes, kind, type);
        let lo = 0, hi = sections.length - 1, found = -1;
        while (lo <= hi) { const mid = (lo + hi) >> 1; if (sections[mid] <= section) { found = sections[mid]; lo = mid + 1; } else hi = mid - 1; }
        if (found >= (type === 'default' ? 1 : 0)) return { key: `${kind}${type}${found}`, nodes: [bySection.get(found)!] };
        if (type === 'default') return undefined;
        const running = runningFurniture(nodes, kind);
        return running.length && furnitureKind(running, kind) === type ? { key: `${kind}${type}running`, nodes: running } : undefined;
    }

    /**
     * The master page a page of the current section is on: one of its own where its page (size, margins)
     * or its header or footer differ from the first section's, else the first's (none, or `Body` once
     * another master was turned to); and the paragraph a page break turns to it with, where it differs from
     * the one the pages before were on.
     */
    private pageStart(): string {
        this.waitingSection = undefined;
        const setup = this.currentSetup;
        const first = this.paperOf(this.ast.metadata?.pageSetup), own = setup ? this.paperOf(setup) : first;
        const samePaper = Math.abs(first.w - own.w) < 0.5 && Math.abs(first.h - own.h) < 0.5
            && (['top', 'right', 'bottom', 'left'] as const).every(k => Math.abs(first.margin[k] - own.margin[k]) < 0.5);
        const section = this.currentSection;
        const header = this.sectionPart('header', section), footer = this.sectionPart('footer', section);
        // Its first page's (where it has a title page: its own first page's part, or its page's
        // differentFirstPage; the part in force there, else none) and its even pages' (the first section's
        // are the body's own, see buildStylesXml).
        const lists = { header: this.ast.auxiliary?.headers, footer: this.ast.auxiliary?.footers };
        const titled = !!(sectionFurniture(lists.header, 'header', section, 'first') || sectionFurniture(lists.footer, 'footer', section, 'first') || setup?.differentFirstPage);
        const none = { key: 'none', nodes: [] as OfficeContentNode[] };
        const parts = section >= 1 ? {
            firstHeader: titled ? this.sectionPart('header', section, 'first') ?? none : undefined, firstFooter: titled ? this.sectionPart('footer', section, 'first') ?? none : undefined,
            evenHeader: this.sectionPart('header', section, 'even'), evenFooter: this.sectionPart('footer', section, 'even'),
        } : {};
        const more = Object.values(parts).some(Boolean);
        // (Numbered its own way, it is a page style of its own: ODF numbers a page style's pages.)
        const numbering = (setup ?? this.ast.metadata?.pageSetup)?.pageNumberFormat ?? 'decimal';
        const ownNumbering = numbering !== (this.ast.metadata?.pageSetup?.pageNumberFormat ?? 'decimal');
        let master = '', page: SectionPage | undefined;
        if (!samePaper || header || footer || more || ownNumbering) {
            const columns = samePaper ? this.ast.metadata?.pageSetup?.columns : setup?.columns;
            const key = JSON.stringify([own.w, own.h, own.margin, columns ?? null, numbering, header?.key ?? null, footer?.key ?? null, ...Object.values(parts).map(p => p?.key ?? null)]);
            master = this.sectionMasterByKey.get(key) ?? '';
            page = this.sectionPageByName.get(master);
            if (!master) {
                master = `Section${this.sectionPages.length + 1}`;
                page = { name: master, setup: samePaper ? this.ast.metadata?.pageSetup : setup, numbering, key, header: header?.nodes, footer: footer?.nodes,
                    firstHeader: parts.firstHeader?.nodes, firstFooter: parts.firstFooter?.nodes, evenHeader: parts.evenHeader?.nodes, evenFooter: parts.evenFooter?.nodes };
                this.sectionPages.push(page);
                this.sectionPageByName.set(master, page);
                this.sectionMasterByKey.set(key, master);
            }
        }
        // The columns the page itself is set in: a section master's own, else the body's.
        this.onSectionPage = !samePaper;
        this.pageColumns = samePaper ? this.ast.metadata?.pageSetup?.columns : setup?.columns;
        // A section of a first page of its own starts on that page's master, which hands over to its own.
        const startsOnFirst = !!(page && (page.firstHeader || page.firstFooter));
        if (master === this.activeMaster && !startsOnFirst) return this.waitBreak({ before: 'page' });
        const name = startsOnFirst ? `${master}First` : master || 'Body';
        if (!master) this.bodyMasterUsed = true;
        this.activeMaster = master;
        return this.waitBreak({ before: 'page', master: name });
    }
    /** A section started on the same page with a header or footer of its own, not yet shown (see blockBreak). */
    private waitingSection: number | undefined;
    /** Whether a section has a header or footer part of its own. */
    private hasOwnParts(section: number): boolean {
        return (['header', 'footer'] as const).some(kind => (['default', 'first', 'even'] as const).some(type => !!sectionFurniture(kind === 'header' ? this.ast.auxiliary?.headers : this.ast.auxiliary?.footers, kind, section, type)));
    }
    /** Reports a waiting section's own header or footer, which no page break showed. */
    private reportUnshownSectionPart(): void {
        if (this.waitingSection === undefined) return;
        this.waitingSection = undefined;
        if (!this.unshownReported) { this.unshownReported = true; this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'odt', feature: 'a header or footer of a section starting mid-page, with no page break before it ends' }); }
    }
    private unshownReported = false;
    /** The columns the current page is set in (null: the body's, before any section page). */
    private pageColumns: PageSetup['columns'] | null = null;

    /** The section page's paper as the document's first page's is resolved (format, landscape, margins; one that sets no page the document's). */
    private paperOf(setup: PageSetup | undefined) {
        const cfg = this.config.odtConfig;
        return resolvePaper(cfg.format, cfg.landscape, onDocumentPage(setup, this.ast.metadata?.pageSetup), cfg.margin);
    }

    /**
     * A page layout of a paper and columns, as ODF writes it; where the document has a header (or footer)
     * and the page its distance from the edge, ODF's way: the page's margin is that distance, and the
     * header's height and its gap to the body the rest of the body's margin (the header and footer stood
     * at the body's margin, a header distance lost).
     */
    private pageLayoutXml(name: string, paper: ReturnType<typeof resolvePaper>, columns: PageSetup['columns'] | undefined, numbering?: PageNumberFormat): string {
        const mar = { ...paper.margin };
        // How its pages are numbered (`style:num-format`; 1 is ODF's own, written as nothing).
        const numFormat = numbering && numbering !== 'decimal' ? odfNumFormatOf(numbering) : undefined;
        if (numbering && numbering !== 'decimal' && !numFormat && !this.unwrittenNumbering.has(numbering)) {
            this.unwrittenNumbering.add(numbering);
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'odt', feature: `page numbers in ${numbering} (written in digits)` });
        }
        let styles = '';
        const hd = paper.headerDistance, fd = paper.footerDistance;
        if (this.writesHeader && hd !== undefined && hd < mar.top) {
            styles += `<style:header-style><style:header-footer-properties fo:min-height="0pt" fo:margin-bottom="${fmtPt(mar.top - hd)}" style:dynamic-spacing="true"/></style:header-style>`;
            mar.top = hd;
        }
        if (this.writesFooter && fd !== undefined && fd < mar.bottom) {
            styles += `<style:footer-style><style:header-footer-properties fo:min-height="0pt" fo:margin-top="${fmtPt(mar.bottom - fd)}" style:dynamic-spacing="true"/></style:footer-style>`;
            mar.bottom = fd;
        }
        return `<style:page-layout style:name="${name}"><style:page-layout-properties`
            + ` fo:page-width="${fmtIn(paper.w / 72)}" fo:page-height="${fmtIn(paper.h / 72)}" style:print-orientation="${paper.landscape ? 'landscape' : 'portrait'}"`
            + ` fo:margin-top="${fmtPt(mar.top)}" fo:margin-bottom="${fmtPt(mar.bottom)}"`
            + ` fo:margin-left="${fmtPt(mar.left)}" fo:margin-right="${fmtPt(mar.right)}"${numFormat ? ` style:num-format="${numFormat}"` : ''}${(columns?.count ?? 1) > 1 ? `>${this.columnsXml(columns, paper)}</style:page-layout-properties>` : '/>'}${styles}</style:page-layout>`;
    }
    /** The page number formats a section had that ODF has none of, each reported once. */
    private unwrittenNumbering = new Set<string>();
    /** Whether the document writes a header (a footer) anywhere: its pages' layouts then hold its distance. */
    private writesHeader = false;
    private writesFooter = false;

    /** Columns as ODF writes them (`style:columns`): their number, gap, widths where they differ, a line between. */
    private columnsXml(given: PageSetup['columns'] | undefined, paper = this.paperOf(this.ast.metadata?.pageSetup)): string {
        // (Held to the page: a gap of millions made a file LibreOffice would not open, see heldColumns.)
        const columns = heldColumns(given, paper.w, paper.w - paper.margin.left - paper.margin.right);
        if (!columns) return '<style:columns fo:column-count="1" fo:column-gap="0pt"/>';
        const sep = columns.separator ? '<style:column-sep style:width="0.5pt" style:color="#000000" style:height="100%" style:style="solid"/>' : '';
        const own = columns.widths?.length === columns.count
            ? columns.widths.map((w, i) => {
                // A column's relative width holds its indents (the halves of the gaps beside it), as office suites read it.
                const start = i ? columns.gap / 2 : 0, end = i < columns.count - 1 ? columns.gap / 2 : 0;
                return `<style:column style:rel-width="${Math.max(1, Math.round((w + start + end) * 20))}*" fo:start-indent="${fmtPt(start)}" fo:end-indent="${fmtPt(end)}"/>`;
            }).join('')
            : '';
        return `<style:columns fo:column-count="${columns.count}" fo:column-gap="${fmtPt(columns.gap)}">${sep}${own}</style:columns>`;
    }

    private blockBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        const sectionStart = t === 'section' || (t === 'page' && typeof (node.metadata as BreakMetadata | undefined)?.section === 'number');
        if (sectionStart) {
            // A section starting on this page (or on the next, after its page break), in its own columns:
            // a text:section of them, where they are not the page's (which the page layout has).
            const meta = node.metadata as BreakMetadata | undefined;
            const setup = meta?.pageSetup;
            const columns = setup?.columns;
            let out = this.openSection ? '</text:section>' : '';
            // (As the break numbers it: one numbering only some breaks leaves the others in the section before.)
            // A section starting on the same page whose header or footer is its own shows it from the next page
            // break (ODF turns to another master page only there); one that ends before any is reported.
            this.reportUnshownSectionPart();
            // (One on pages of another size, or turned, starts on a page of its own, as a word processor sets it:
            // it went on on the pages before, of their size.)
            const before = this.paperOf(this.currentSetup ?? this.ast.metadata?.pageSetup), after = setup ? this.paperOf(setup) : before;
            const newPage = t === 'page' || Math.abs(before.w - after.w) > 0.5 || Math.abs(before.h - after.h) > 0.5;
            if (typeof meta?.section === 'number') this.currentSection = meta.section;
            if (setup) this.currentSetup = setup;
            if (!newPage && typeof meta?.section === 'number' && this.hasOwnParts(meta.section)) this.waitingSection = meta.section;
            // A section on a page of its own size or margins, or with a header or footer of its own: a master
            // page of them (and its columns), which the break's paragraph turns to (and the first's after it).
            // A section starting on the same page takes them at its next page.
            if (newPage) out += this.pageStart();
            const pageColumns = this.pageColumns === null ? this.ast.metadata?.pageSetup?.columns : this.pageColumns;
            this.openSection = (columns?.count ?? 1) !== (pageColumns?.count ?? 1) || (columns?.gap ?? 0) !== (pageColumns?.gap ?? 0);
            if (this.openSection) {
                const style = this.activeStyles.intern('section', 'Sect', `<style:section-properties>${this.columnsXml(columns)}</style:section-properties>`);
                out += `<text:section text:style-name="${style}" text:name="Section${++this.sections}">`;
            }
            return out;
        }
        if (t === 'page') return this.pageStart();
        if (t === 'column') return this.waitBreak({ before: 'column' });
        if (t === 'thematic') return `<text:p text:style-name="${this.internPara(' fo:border-bottom="0.5pt solid #808080" fo:padding-bottom="1pt"')}"/>`;
        if (t === 'lastRenderedPage') return '';
        return '<text:p/>';
    }

    private codeBlock(node: OfficeContentNode): string {
        if ((node.metadata as any)?.math && !this.mathWarned) { this.mathWarned = true; this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'odt', feature: 'math' }); }
        const text = node.text || this.getNodeText(node);
        return `<text:p text:style-name="Preformatted_20_Text">${encodeOdfText(text)}</text:p>`;
    }

    private async admonition(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const color = ADMONITION_COLOR[meta?.admonitionType] || '0969DA';
        const title = meta?.title || (meta?.admonitionType ? meta.admonitionType[0].toUpperCase() + meta.admonitionType.slice(1) : 'Note');
        const borderAttrs = ` fo:border-left="2.25pt solid #${color}" fo:margin-left="0.14in" fo:padding-left="0.1in"`;
        const titleStyle = this.internPara(borderAttrs);
        const titleRun = this.activeStyles.intern('text', 'T', `<style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" fo:color="#${color}"/>`);
        const titleP = `<text:p text:style-name="${titleStyle}"><text:span text:style-name="${titleRun}">${encodeOdfText(title)}</text:span></text:p>`;
        const prev = this.paraContextAttrs;
        this.paraContextAttrs = borderAttrs;
        const body = await this.renderBlocks(node.children);
        this.paraContextAttrs = prev;
        return titleP + body;
    }

    private async chart(node: OfficeContentNode): Promise<string> {
        if (this.config.includeCharts === false) return '';
        const meta = node.metadata as any;
        const att = this.getAttachment(meta?.attachmentName);
        const data = att?.chartData;
        if (!data) return `<text:p>${encodeOdfText(`[Chart: ${meta?.attachmentName || ''}]`)}</text:p>`;
        const caption = data.title ? `<text:p><text:span text:style-name="${this.boldStyle()}">${encodeOdfText(data.title)}</text:span></text:p>` : '';
        return caption + await this.table(this.chartTable(data, cellOf));
    }

    private async embed(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const rawUrl = embedUrl(meta);
        const url = rawUrl ? sanitizeOfficePackageUrl(rawUrl) : '';
        if (!url) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'odt', feature: 'embed' });
            const fb = meta?.label || node.text || this.getNodeText(node);
            return fb ? `<text:p>${this.span(fb, undefined)}</text:p>` : '';
        }
        return `<text:p><text:a xlink:type="simple" xlink:href="${xmlText(encUrl(url))}">${this.span(meta?.label || url, undefined)}</text:a></text:p>`;
    }

    // ── styleMap / named-style helpers ────────────────────────────────────────────

    private styleForTag(tag: string | undefined): string | undefined {
        if (!tag || !isSafeStyleMapTag(tag)) return undefined;
        const map: Record<string, string> = lookupTable({
            h1: 'Heading_20_1', h2: 'Heading_20_2', h3: 'Heading_20_3', h4: 'Heading_20_4', h5: 'Heading_20_5', h6: 'Heading_20_6',
            blockquote: 'Quotations', pre: 'Preformatted_20_Text',
        });
        return map[tag.toLowerCase()];
    }
    private knownParaStyle(style: string | undefined): string | undefined {
        const map: Record<string, string> = lookupTable({ Quote: 'Quotations', IntenseQuote: 'Quotations', Title: 'Title', Code: 'Preformatted_20_Text' });
        return style ? map[style] : undefined;
    }

    // ── metadata block ────────────────────────────────────────────────────────────

    private renderMetadataBlock(): string {
        const m = this.effectiveMetadata;
        let out = '';
        if (m.title) out += `<text:p text:style-name="Title">${encodeOdfText(m.title)}</text:p>`;
        const modIso = toW3CDTF(m.modified);
        const bits = [m.author, modIso ? modIso.slice(0, 10) : undefined].filter(Boolean).join(' · ');
        if (bits) out += `<text:p><text:span text:style-name="${this.activeStyles.intern('text', 'T', '<style:text-properties fo:font-style="italic" style:font-style-asian="italic"/>')}">${encodeOdfText(bits)}</text:span></text:p>`;
        return out;
    }

    // ── part builders ─────────────────────────────────────────────────────────────

    private buildContentXml(body: string): string {
        const bodyContent = body.trim() ? body : '<text:p/>';
        return `<?xml version="1.0" encoding="UTF-8"?>\n`
            + `<office:document-content ${ODF_NS} office:version="1.2">`
            + `<office:font-face-decls>${this.contentStyles.fontFaceDeclsXml()}</office:font-face-decls>`
            + `<office:automatic-styles>${this.contentStyles.automaticStylesXml()}</office:automatic-styles>`
            + `<office:body><office:text>${bodyContent}</office:text></office:body>`
            + `</office:document-content>`;
    }

    private buildStylesXml(headerXml: string, footerXml: string, headerKind: 'default' | 'first' | 'even' = 'default', footerKind: 'default' | 'first' | 'even' = 'default'): string {
        // The Preformatted Text named style always references Courier New, so styles.xml must always
        // declare that font-face. Register it in the font set (deduplicated) rather than prepending a
        // literal decl: a header/footer that also used inline code would otherwise declare it twice.
        this.stylesFileStyles.useFont('Courier New');
        const cfg = this.config.odtConfig;
        const paper = resolvePaper(cfg.format, cfg.landscape, this.ast.metadata?.pageSetup, cfg.margin);
        this.writesHeader = !!this.ast.auxiliary?.headers?.length;
        this.writesFooter = !!this.ast.auxiliary?.footers?.length;
        const pageLayout = this.pageLayoutXml('pm1', paper, this.ast.metadata?.pageSetup?.columns, this.ast.metadata?.pageSetup?.pageNumberFormat);
        // A part on every page, on even pages alone (`style:header-left`, its odd pages' header an empty
        // paragraph: an empty `style:header` is read as none, and its left pages' with it), or
        // on the first page alone: then the first page's master (the document starts on `Standard`)
        // hands over to one for the rest (`style:next-style-name`), as a word processor's first-page style.
        // Where the document's even pages have parts of their own (Word's even and odd headers), a kind with
        // none of its own shows none there, as Word shows them: an empty left page's part, not the odd pages'.
        const facing = [...(this.ast.auxiliary?.headers ?? []), ...(this.ast.auxiliary?.footers ?? [])]
            .some(n => (n.metadata as HeaderFooterMetadata | undefined)?.type === 'even' && !(n.metadata as HeaderFooterMetadata | undefined)?.hidden);
        const noLeft = (tag: 'header' | 'footer') => facing ? `<style:${tag}-left><text:p/></style:${tag}-left>` : '';
        // A section's running part where it has none of its own: on every page, or on even pages alone (a first page's shows on none).
        const part = (tag: 'header' | 'footer', xml: string, kind: string) => !xml || kind === 'first' ? ''
            : kind === 'even' ? `<style:${tag}><text:p/></style:${tag}><style:${tag}-left>${xml}</style:${tag}-left>` : `<style:${tag}>${xml}</style:${tag}>${noLeft(tag)}`;
        // A master's parts from a section's own: every page's (an empty one written as an empty paragraph:
        // an empty part is read as none), even pages' (`-left`), and its first page's where it is that page's.
        // (A first page's master shows its part on whichever side it falls: no left page's part of its own.)
        const own = (tag: 'header' | 'footer', every: string | undefined, even: string | undefined, firstPage = false) => (every === undefined && even === undefined ? ''
            : `<style:${tag}>${every || '<text:p/>'}</style:${tag}>`) + (even !== undefined ? `<style:${tag}-left>${even || '<text:p/>'}</style:${tag}-left>` : every !== undefined && !firstPage ? noLeft(tag) : '');
        // The first section's parts of each kind: the running one where it is of its type, else its own.
        const body = this.bodyOwnParts;
        const ofKind = (kind: string, xml: string, type: 'first' | 'even', own: string | undefined) => kind === type ? xml : own;
        const first = { header: ofKind(headerKind, headerXml, 'first', body.firstHeader), footer: ofKind(footerKind, footerXml, 'first', body.firstFooter) };
        const even = { header: ofKind(headerKind, headerXml, 'even', body.evenHeader), footer: ofKind(footerKind, footerXml, 'even', body.evenFooter) };
        const every = { header: headerKind === 'default' ? headerXml || undefined : undefined, footer: footerKind === 'default' ? footerXml || undefined : undefined };
        const firstOwn = first.header !== undefined || first.footer !== undefined || !!this.ast.metadata?.pageSetup?.differentFirstPage;
        const masters = (page: 'first' | 'rest') => page === 'first'
            // (A title page shows its own first page's parts, and none of a type it has none of, as Word does.)
            ? own('header', first.header ?? (this.writesHeader ? '' : undefined), undefined, true) + own('footer', first.footer ?? (this.writesFooter ? '' : undefined), undefined, true)
            : own('header', every.header, even.header) + own('footer', every.footer, even.footer);
        const masterPage = firstOwn
            ? `<style:master-page style:name="Standard" style:page-layout-name="pm1" style:next-style-name="Rest">${masters('first')}</style:master-page>`
                + `<style:master-page style:name="Rest" style:page-layout-name="pm1">${masters('rest')}</style:master-page>`
            : `<style:master-page style:name="Standard" style:page-layout-name="pm1">${masters('rest')}</style:master-page>`;
        // Sections' own pages, and the body's page to return to after them, each with the body's header and footer.
        const sectionLayouts = this.sectionPages.map((p, i) => this.pageLayoutXml(`pm${i + 2}`, this.paperOf(p.setup), p.setup?.columns, p.numbering)).join('');
        // A section's own header and footer, else the first section's (as on the body's pages).
        // Even pages' and a first page's of its own too: its first page on a master of its own (`<name>First`)
        // that hands over to the section's.
        const sectionSide = (tag: 'header' | 'footer', every: string | undefined, even: string | undefined, kind: string, running: string) =>
            every !== undefined || even !== undefined ? own(tag, every ?? (kind === 'default' ? running || undefined : undefined), even) : part(tag, running, kind);
        const sectionParts = (p: SectionPage) => sectionSide('header', p.headerXml, p.evenHeaderXml, headerKind, headerXml) + sectionSide('footer', p.footerXml, p.evenFooterXml, footerKind, footerXml);
        // (A section's title page: the first page's parts in force, none of a type it has none of, as Word does.)
        // (An empty one where the document has them elsewhere: a page without one stood its body higher.)
        const firstParts = (p: SectionPage) => (p.firstHeaderXml !== undefined && (p.firstHeaderXml || this.writesHeader) ? own('header', p.firstHeaderXml, undefined, true) : '')
            + (p.firstFooterXml !== undefined && (p.firstFooterXml || this.writesFooter) ? own('footer', p.firstFooterXml, undefined, true) : '');
        const sectionMasters = this.sectionPages.map((p, i) => (p.firstHeaderXml !== undefined || p.firstFooterXml !== undefined
            ? `<style:master-page style:name="${p.name}First" style:page-layout-name="pm${i + 2}" style:next-style-name="${p.name}">${firstParts(p)}</style:master-page>` : '')
            + `<style:master-page style:name="${p.name}" style:page-layout-name="pm${i + 2}">${sectionParts(p)}</style:master-page>`).join('')
            + (this.bodyMasterUsed ? `<style:master-page style:name="Body" style:page-layout-name="pm1">${masters('rest')}</style:master-page>` : '');
        return `<?xml version="1.0" encoding="UTF-8"?>\n`
            + `<office:document-styles ${ODF_NS} office:version="1.2">`
            + `<office:font-face-decls>${this.stylesFileStyles.fontFaceDeclsXml()}</office:font-face-decls>`
            + `<office:styles>${NAMED_STYLES}</office:styles>`
            + `<office:automatic-styles>${pageLayout}${sectionLayouts}${this.stylesFileStyles.automaticStylesXml()}</office:automatic-styles>`
            + `<office:master-styles>${masterPage}${sectionMasters}</office:master-styles>`
            + `</office:document-styles>`;
    }

    private buildMetaXml(iso: string): string {
        const m = this.effectiveMetadata;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        const created = toW3CDTF(m.created);
        let meta = el('dc:title', m.title) + el('meta:initial-creator', m.author) + el('dc:creator', m.lastModifiedBy || m.author)
            + el('dc:description', m.description) + el('dc:subject', m.subject) + el('dc:language', (m as any).language);
        if (m.keywords) for (const k of String(m.keywords).split(',').map(s => s.trim()).filter(Boolean)) meta += `<meta:keyword>${xmlText(k)}</meta:keyword>`;
        if (created) meta += `<meta:creation-date>${created}</meta:creation-date>`;
        meta += `<dc:date>${iso}</dc:date><meta:generator>officeParser</meta:generator>`;
        meta += odfUserDefinedXml(m.customProperties);
        return `<?xml version="1.0" encoding="UTF-8"?>\n`
            + `<office:document-meta ${ODF_NS} office:version="1.2"><office:meta>${meta}</office:meta></office:document-meta>`;
    }

    private buildManifest(entries: Map<string, string>): string {
        let files = `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="${ODT_MIMETYPE}"/>`;
        for (const [path, mime] of entries) files += `<manifest:file-entry manifest:full-path="${xmlText(path)}" manifest:media-type="${xmlText(mime)}"/>`;
        return `<?xml version="1.0" encoding="UTF-8"?>\n`
            + `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">${files}</manifest:manifest>`;
    }
}

/**
 * Blocks that write their bookmarks themselves: through `paragraph()` or `heading()`, a list's items,
 * and a run, row or cell among blocks, written in a line (see withInlineBookmarks).
 */
const OWN_BOOKMARKS: ReadonlySet<string> = new Set(['paragraph', 'heading', 'definitionTerm', 'definitionDescription', 'list', 'text', 'row', 'cell']);

/** Whether a node among blocks belongs in a paragraph's line: a run, a picture, a line break, inline math. */
function isInlineNode(node: OfficeContentNode): boolean {
    if (node.type === 'text' || node.type === 'image') return true;
    if (node.type === 'break') return !['page', 'column', 'thematic', 'lastRenderedPage', 'section'].includes((node.metadata as any)?.breakType);
    return node.type === 'code' && (node.metadata as any)?.math === 'inline';
}

function cellOf(text: string): OfficeContentNode {
    return { type: 'cell', text, children: [{ type: 'paragraph', children: [{ type: 'text', text } as OfficeContentNode] } as OfficeContentNode], metadata: { row: 0, col: 0 } as any };
}

/** Static named paragraph styles (office:styles in styles.xml). A constant string keeps baselines byte-stable. */
const NAMED_STYLES = (() => {
    const heading = (n: number, pt: number) => `<style:style style:name="Heading_20_${n}" style:display-name="Heading ${n}" style:family="paragraph" style:parent-style-name="Heading" style:next-style-name="Text_20_body" style:default-outline-level="${n}"><style:text-properties fo:font-weight="bold" style:font-weight-asian="bold" fo:font-size="${pt}pt" style:font-size-asian="${pt}pt"/></style:style>`;
    return `<style:style style:name="Standard" style:family="paragraph" style:class="text"/>`
        + `<style:style style:name="Text_20_body" style:display-name="Text body" style:family="paragraph" style:parent-style-name="Standard"/>`
        + `<style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Text_20_body"><style:text-properties fo:font-weight="bold"/></style:style>`
        + heading(1, 18) + heading(2, 16) + heading(3, 14) + heading(4, 13) + heading(5, 12) + heading(6, 12)
        + `<style:style style:name="Title" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-weight="bold" fo:font-size="28pt" style:font-size-asian="28pt"/></style:style>`
        + `<style:style style:name="Quotations" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="0.5in" fo:margin-right="0.5in"/><style:text-properties fo:font-style="italic" style:font-style-asian="italic"/></style:style>`
        + `<style:style style:name="Preformatted_20_Text" style:display-name="Preformatted Text" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:background-color="#f5f5f5"/><style:text-properties style:font-name="Courier New"/></style:style>`
        + `<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-size="10pt" style:font-size-asian="10pt"/></style:style>`;
})();
