import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, ImageMode, OdtGeneratorConfig, OfficeContentNode, OfficeParserAST, OfficeWarningType, TextFormatting } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, isSafeStyleMapTag, sanitizeOfficePackageUrl, stripInvalidXmlChars } from '../utils/sanitize.js';
import { ADMONITION_COLOR, decodeBase64, embedUrl, encUrl, fillSheetRowGaps, hexColor, isHeaderRow, lengthToPt, marginPt, MIME_EXT, paperSizePt, resolveZipInstant, sniffImageSize, toBookmarkNameRaw, toW3CDTF } from '../utils/officeGenUtils.js';
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

type StyleFamily = 'text' | 'paragraph' | 'table' | 'table-cell';

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

    intern(family: StyleFamily, prefix: string, props: string, extraAttrs = ''): string {
        const key = `${family}::${extraAttrs}::${props}`;
        const existing = this.byKey.get(key);
        if (existing) return existing;
        const name = `${prefix}${(this.counters[prefix] = (this.counters[prefix] ?? 0) + 1)}`;
        this.byKey.set(key, name);
        this.serialized.push(`<style:style style:name="${name}" style:family="${family}"${extraAttrs}>${props}</style:style>`);
        return name;
    }
    internListStyle(xml: string): void { this.listStyles.push(xml); }
    useFont(name: string): void { this.fonts.add(name); }
    automaticStylesXml(): string { return this.serialized.join('') + this.listStyles.join(''); }
    fontFaceDeclsXml(): string {
        return [...this.fonts].map(f => `<style:font-face style:name="${xmlText(f)}" svg:font-family="${xmlText(quoteFontFamily(f))}"/>`).join('');
    }
}

interface MediaPart { name: string; bytes: Uint8Array; ext: string; contentType: string; }
interface ListDef { name: string; registry: StyleRegistry; levels: Map<number, 'ordered' | 'unordered'>; startAt: Map<number, number>; }

/**
 * Generates an OpenDocument Text (`.odt`) document from any AST, as a `Uint8Array`.
 *
 * Hand-built with `fflate.zipSync` (the `EpubGenerator`/`DocxGenerator` approach), no new
 * dependencies, browser-safe, and reproducible (pinned zip mtimes, deterministic style/id names).
 * It emits the ODF dialect `OpenOfficeParser` reads back, so a generated document round-trips through
 * the parser with near-zero loss. Unlike DOCX, ODF carries formatting in named automatic styles and
 * keeps notes/comments/headers inline in the same files, so there is no per-part relationship layer.
 */
export class OdtGenerator extends BaseGenerator<'odt'> {
    private contentStyles = new StyleRegistry();
    private stylesFileStyles = new StyleRegistry();
    private activeStyles = this.contentStyles;
    private paraContextAttrs = '';

    private media: MediaPart[] = [];
    private mediaByAttachment = new Map<string, string>();
    private manifestMedia = new Map<string, string>(); // Pictures/xxx -> contentType

    private listDefs = new Map<string, ListDef>();
    private listCounter = 0;
    private tableCounter = 0;
    private frameCounter = 0;
    private commentCounter = 0;
    private footnoteOrd = 0;
    private readonly writtenNotes = new Map<OfficeContentNode, { id: string; ord: number; cls: 'footnote' | 'endnote' }>();
    private endnoteOrd = 0;

    private usedBookmarkNames = new Set<string>();
    private mathWarned = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'odt'>) {
        super('odt', ast, config);
    }

    // ── entry point ─────────────────────────────────────────────────────────────

    async generate(): Promise<ConversionResult<'odt'>> {
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);

        const metaBlock = this.config.renderMetadata ? this.renderMetadataBlock() : '';
        const body = await this.renderBlocks(this.ast.content);

        const hasHeader = !!(this.ast.auxiliary?.headers?.length);
        const hasFooter = !!(this.ast.auxiliary?.footers?.length);
        this.activeStyles = this.stylesFileStyles;
        const headerXml = hasHeader ? await this.renderBlocks(this.ast.auxiliary!.headers!) : '';
        const footerXml = hasFooter ? await this.renderBlocks(this.ast.auxiliary!.footers!) : '';
        this.activeStyles = this.contentStyles;

        // Serialize list styles into their owning file's registry (learned during render).
        for (const def of this.listDefs.values()) def.registry.internListStyle(this.buildListStyleXml(def));

        const contentXml = this.buildContentXml(metaBlock + body);
        const stylesXml = this.buildStylesXml(headerXml, footerXml);
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
        for (let idx = 0; idx < items.length; idx++) {
            const node = items[idx];
            checkAbortSignal(this.config.abortSignal);
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
                out += await this.renderListRun(run);
                prevPaginated = null;
                continue;
            }
            const override = await this.handleOnNode(node);
            if (override === false) { prevPaginated = null; continue; }
            if (typeof override === 'string') { out += override; prevPaginated = null; continue; }
            if ((node.type === 'page' || node.type === 'slide') && prevPaginated === node.type) {
                out += `<text:p text:style-name="${this.pageBreakStyle()}"/>`;
            }
            out += await this.renderBlockNode(node);
            prevPaginated = (node.type === 'page' || node.type === 'slide') ? node.type : null;
        }
        return out;
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

    /** Mints a unique bookmark name; the first claim of a base keeps it, later claims get `_2`. */
    private mintBookmark(rawName: string): string {
        const base = toBookmarkNameRaw(rawName);
        let name = base, i = 2;
        while (this.usedBookmarkNames.has(name)) { const suffix = `_${i++}`; name = base.slice(0, 40 - suffix.length) + suffix; }
        this.usedBookmarkNames.add(name);
        return name;
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
        const spans = group.map(n => this.span(n.text || '', n.formatting)).join('');
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
        if (t === 'carriageReturn') return '<text:line-break/>';
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
        const written = this.writtenNotes.get(node);
        if (written) return `<text:note-ref text:note-class="${written.cls}" text:reference-format="text" text:ref-name="${written.id}">${written.ord}</text:note-ref>`;
        const ord = cls === 'footnote' ? ++this.footnoteOrd : ++this.endnoteOrd;
        const id = `${cls === 'footnote' ? 'ftn' : 'edn'}${ord}`;
        this.writtenNotes.set(node, { id, ord, cls });
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
            def.levels.set(lvl, meta?.listType === 'ordered' ? 'ordered' : 'unordered');
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
            out += `<text:p>${prefix}${blockRefs}${inner}</text:p>`;
        }
        for (let d = depth; d >= 0; d--) out += '</text:list-item></text:list>';
        return out;
    }

    private buildListStyleXml(def: ListDef): string {
        let levels = '';
        for (let l = 0; l <= 9; l++) {
            const type = def.levels.get(l) ?? def.levels.get(0) ?? 'unordered';
            const props = `<style:list-level-properties text:space-before="${fmtIn(0.25 * l)}" text:min-label-width="0.25in"/>`;
            if (type === 'ordered') {
                const start = def.startAt.get(l) ?? 1;
                levels += `<text:list-level-style-number text:level="${l + 1}" style:num-format="1" style:num-suffix="." text:start-value="${start}">${props}</text:list-level-style-number>`;
            } else {
                levels += `<text:list-level-style-bullet text:level="${l + 1}" text:bullet-char="•">${props}</text:list-level-style-bullet>`;
            }
        }
        return `<text:list-style style:name="${def.name}">${levels}</text:list-style>`;
    }

    // ── tables ───────────────────────────────────────────────────────────────────

    private async table(node: OfficeContentNode): Promise<string> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return '';
        const cols = this.gridWidth(rows);
        const tableName = `Table${++this.tableCounter}`;
        const tableStyle = this.ensureTableStyle(this.tableAlign((node.metadata as any)?.align));
        const colXml = `<table:table-column table:number-columns-repeated="${cols}"/>`;

        const active = new Map<number, number>();
        const rendered: { xml: string; header: boolean }[] = [];
        for (let ri = 0; ri < rows.length; ri++) {
            const row = rows[ri];
            const cells = (row.children || []).filter(c => c.type === 'cell');
            const header = isHeaderRow(row, ri === 0);
            let cellsXml = '';
            let col = 0, ci = 0;
            while (ci < cells.length || [...active.keys()].some(c => c >= col)) {
                if ((active.get(col) || 0) > 0) {
                    cellsXml += '<table:covered-table-cell/>';
                    active.set(col, active.get(col)! - 1);
                    if (active.get(col)! <= 0) active.delete(col);
                    col++;
                    continue;
                }
                if (ci >= cells.length) {
                    // No explicit cells left, but a vertical merge is still pending at a later column:
                    // fill this gap with an empty cell and advance, so the covered cell lands correctly.
                    if (![...active.keys()].some(c => c > col)) break;
                    cellsXml += '<table:table-cell><text:p/></table:table-cell>';
                    col++;
                    continue;
                }
                // Sparse source grid: ExcelParser emits only the non-empty cells, each carrying its own
                // column index. Fill the skipped columns with empty cells so a value in D1 lands in
                // column 4, rather than sliding left to whatever the running cursor happened to reach.
                const nextCol = (cells[ci].metadata as any)?.col;
                if (typeof nextCol === 'number' && nextCol > col && col < cols) {
                    cellsXml += '<table:table-cell><text:p/></table:table-cell>';
                    col++;
                    continue;
                }
                const cell = cells[ci++];
                const cmeta = cell.metadata as any;
                const colSpan = Math.max(1, Math.min(cols, cmeta?.colSpan || 1));
                const rowSpan = Math.max(1, Math.min(1000, cmeta?.rowSpan || 1));
                const cellStyle = this.ensureCellStyle(hexColor(cmeta?.backgroundColor));
                let spanAttr = '';
                if (colSpan > 1) spanAttr += ` table:number-columns-spanned="${colSpan}"`;
                if (rowSpan > 1) { spanAttr += ` table:number-rows-spanned="${rowSpan}"`; for (let k = 0; k < colSpan; k++) active.set(col + k, rowSpan - 1); }
                let inner = await this.renderBlocks(cell.children);
                if (!inner.trim()) inner = '<text:p/>';
                cellsXml += `<table:table-cell table:style-name="${cellStyle}"${spanAttr}>${inner}</table:table-cell>`;
                for (let k = 1; k < colSpan; k++) cellsXml += '<table:covered-table-cell/>';
                col += colSpan;
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

    private gridWidth(rows: OfficeContentNode[]): number {
        let max = 1;
        const active = new Map<number, number>();
        for (const row of rows) {
            let col = 0;
            const cells = (row.children || []).filter(c => c.type === 'cell');
            let ci = 0;
            while (ci < cells.length || [...active.keys()].some(c => c >= col)) {
                if ((active.get(col) || 0) > 0) { active.set(col, active.get(col)! - 1); if (active.get(col)! <= 0) active.delete(col); col++; continue; }
                if (ci >= cells.length) {
                    if (![...active.keys()].some(c => c > col)) break;
                    col++; // gap column before a still-pending vertical merge
                    continue;
                }
                // Mirror table()'s sparse-grid placement, or the grid would be narrower than the rows.
                const nextCol = (cells[ci].metadata as any)?.col;
                if (typeof nextCol === 'number' && nextCol > col && col < 1000) { col++; continue; }
                const cmeta = cells[ci++].metadata as any;
                const colSpan = Math.max(1, Math.min(1000, cmeta?.colSpan || 1));
                const rowSpan = Math.max(1, Math.min(1000, cmeta?.rowSpan || 1));
                if (rowSpan > 1) for (let k = 0; k < colSpan; k++) active.set(col + k, rowSpan - 1);
                col += colSpan;
            }
            max = Math.max(max, col);
        }
        return Math.min(1000, max);
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
        const table = await this.table({ type: 'table', children: fillSheetRowGaps(children.filter(c => c.type === 'row')) } as OfficeContentNode);
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
        let bytes: Uint8Array;
        try { bytes = decodeBase64(att.data); } catch { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName }); return null; }
        let name = this.mediaByAttachment.get(attachmentName);
        if (!name) {
            name = `image${this.media.length + 1}.${ext}`;
            const contentType = att.mimeType || 'image/png';
            this.media.push({ name, bytes, ext, contentType });
            this.mediaByAttachment.set(attachmentName, name);
            this.manifestMedia.set(`Pictures/${name}`, contentType);
        }
        return { href: `Pictures/${name}`, intrinsic: sniffImageSize(bytes) };
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
        const size = paperSizePt(cfg.format);
        const widthPt = cfg.landscape ? size.h : size.w;
        const left = marginPt(cfg.margin.left);
        const right = marginPt(cfg.margin.right);
        return Math.max(36, widthPt - left - right);
    }

    // ── breaks / code / admonition / chart / embed ────────────────────────────────

    private pageBreakStyle(): string { return this.internPara(' fo:break-before="page"'); }

    private blockBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return `<text:p text:style-name="${this.pageBreakStyle()}"/>`;
        if (t === 'column') return `<text:p text:style-name="${this.internPara(' fo:break-before="column"')}"/>`;
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
        const rows: OfficeContentNode[] = [{ type: 'row', children: [cellOf(''), ...data.dataSets.map(d => cellOf(d.name || ''))] } as OfficeContentNode];
        (data.labels || []).forEach((label, i) => {
            rows.push({ type: 'row', children: [cellOf(label), ...data.dataSets.map(d => cellOf(String(d.values?.[i] ?? '')))] } as OfficeContentNode);
        });
        return caption + await this.table({ type: 'table', children: rows } as OfficeContentNode);
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

    private buildStylesXml(headerXml: string, footerXml: string): string {
        // The Preformatted Text named style always references Courier New, so styles.xml must always
        // declare that font-face. Register it in the font set (deduplicated) rather than prepending a
        // literal decl: a header/footer that also used inline code would otherwise declare it twice.
        this.stylesFileStyles.useFont('Courier New');
        const cfg = this.config.odtConfig;
        const size = paperSizePt(cfg.format);
        const pw = (cfg.landscape ? size.h : size.w) / 72; // inches
        const ph = (cfg.landscape ? size.w : size.h) / 72;
        const orient = cfg.landscape ? 'landscape' : 'portrait';
        const mar = cfg.margin;
        const pageLayout = `<style:page-layout style:name="pm1"><style:page-layout-properties`
            + ` fo:page-width="${fmtIn(pw)}" fo:page-height="${fmtIn(ph)}" style:print-orientation="${orient}"`
            + ` fo:margin-top="${fmtPt(marginPt(mar.top))}" fo:margin-bottom="${fmtPt(marginPt(mar.bottom))}"`
            + ` fo:margin-left="${fmtPt(marginPt(mar.left))}" fo:margin-right="${fmtPt(marginPt(mar.right))}"/></style:page-layout>`;
        const hf = (headerXml ? `<style:header>${headerXml}</style:header>` : '') + (footerXml ? `<style:footer>${footerXml}</style:footer>` : '');
        const masterPage = `<style:master-page style:name="Standard" style:page-layout-name="pm1">${hf}</style:master-page>`;
        return `<?xml version="1.0" encoding="UTF-8"?>\n`
            + `<office:document-styles ${ODF_NS} office:version="1.2">`
            + `<office:font-face-decls>${this.stylesFileStyles.fontFaceDeclsXml()}</office:font-face-decls>`
            + `<office:styles>${NAMED_STYLES}</office:styles>`
            + `<office:automatic-styles>${pageLayout}${this.stylesFileStyles.automaticStylesXml()}</office:automatic-styles>`
            + `<office:master-styles>${masterPage}</office:master-styles>`
            + `</office:document-styles>`;
    }

    private buildMetaXml(iso: string): string {
        const m = this.effectiveMetadata;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        const created = toW3CDTF(m.created);
        let meta = el('dc:title', m.title) + el('meta:initial-creator', m.author) + el('dc:creator', m.author)
            + el('dc:description', m.description) + el('dc:subject', m.subject) + el('dc:language', (m as any).language);
        if (m.keywords) for (const k of String(m.keywords).split(',').map(s => s.trim()).filter(Boolean)) meta += `<meta:keyword>${xmlText(k)}</meta:keyword>`;
        if (created) meta += `<meta:creation-date>${created}</meta:creation-date>`;
        meta += `<dc:date>${iso}</dc:date><meta:generator>officeParser</meta:generator>`;
        const props = m.customProperties || {};
        for (const [key, value] of Object.entries(props)) {
            let type = 'string', v: string;
            if (typeof value === 'boolean') { type = 'boolean'; v = value ? 'true' : 'false'; }
            else if (typeof value === 'number') { type = 'float'; v = String(value); }
            else if (value instanceof Date) { type = 'date'; v = value.toISOString().replace(/\.\d+Z$/, 'Z'); }
            else v = String(value);
            meta += `<meta:user-defined meta:name="${xmlText(key)}" meta:value-type="${type}">${xmlText(v)}</meta:user-defined>`;
        }
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
