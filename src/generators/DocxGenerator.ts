import { zipSync, Zippable } from 'fflate';
import { ConversionResult, DocxGeneratorConfig, GeneratorConfig, ImageMode, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, TextFormatting } from '../types.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { escapeXml, isSafeStyleMapTag, sanitizeOfficePackageUrl, stripInvalidXmlChars } from '../utils/sanitize.js';
import { ADMONITION_COLOR, decodeBase64, embedUrl, encUrl, fillSheetRowGaps, hexColor, isHeaderRow, lengthToPt, marginPt, MIME_EXT, paperSizePt, resolveZipInstant, sniffImageSize, toBookmarkNameRaw, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { lookupTable } from '../utils/lookupUtils.js';

const EMU_PER_PT = 12700;
const EMU_PER_IN = 914400;
const TWIPS_PER_PT = 20;

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/**
 * The full WML namespace set. Every part whose body can contain runs, hyperlinks (`r:`) or drawings
 * (`wp:`/`a:`/`pic:`) - document, notes, comments, headers, footers - must declare all of them on its
 * root, or an `r:id`/`wp:inline` inside it references an undeclared prefix and the part is malformed.
 */
const WML_NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" `
    + `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" `
    + `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" `
    + `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `
    + `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;

interface Rel { id: string; type: string; target: string; mode?: string; }
interface MediaPart { name: string; bytes: Uint8Array; ext: string; contentType: string; }
interface NoteEntry { key: string; kind: 'footnote' | 'endnote'; node: OfficeContentNode; }
interface CommentEntry { id: number; node: OfficeContentNode; }

/**
 * Generates a Word (.docx / WordprocessingML) document from any AST, as a `Uint8Array`.
 *
 * The package is hand-built with `fflate.zipSync` (the same approach as `EpubGenerator`), with no
 * new dependencies, browser-safe, and reproducible (pinned zip mtimes, deterministic ids). It emits
 * exactly the WML dialect `WordParser` reads back, so a generated document round-trips through the
 * parser with near-zero loss (the basis of its test strategy).
 */
export class DocxGenerator extends BaseGenerator<'docx'> {
    // Relationships are per owning part: an r:id inside footnotes.xml/header1.xml resolves against
    // that part's own `.rels`, not document.xml.rels. `currentRelOwner` names the part being rendered.
    private partRels = new Map<string, Rel[]>();
    private relDedupe = new Map<string, string>();
    private currentRelOwner = 'word/document.xml';
    private media: MediaPart[] = [];
    private mediaByAttachment = new Map<string, string>();
    private usedExtensions = new Set<string>();
    private drawingCounter = 0;
    private usedBookmarkNames = new Set<string>();
    private bookmarkCounter = 0;
    private footnotes: NoteEntry[] = [];
    private endnotes: NoteEntry[] = [];
    private comments: CommentEntry[] = [];
    private commentCounter = 0;
    private numbering: { listId: string; numId: number; levels: Map<number, 'ordered' | 'unordered'>; startAt: Map<number, number> }[] = [];
    private numByListId = new Map<string, number>();
    private needsNumbering = false;
    private syntheticListRun = 0; // per-run id counter for lists that carry no listId
    private mathWarned = false;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'docx'>) {
        super('docx', ast, config);
    }

    // ── registries ────────────────────────────────────────────────────────────

    /** Adds a relationship to the part currently being rendered and returns its part-local r:id. */
    private addRel(type: string, target: string, mode?: string): string {
        const owner = this.currentRelOwner;
        const dedupeKey = `${owner}::${type}::${target}::${mode ?? ''}`;
        const existing = this.relDedupe.get(dedupeKey);
        if (existing) return existing;
        let bucket = this.partRels.get(owner);
        if (!bucket) { bucket = []; this.partRels.set(owner, bucket); }
        const id = `rId${bucket.length + 1}`;
        bucket.push({ id, type, target, mode });
        this.relDedupe.set(dedupeKey, id);
        return id;
    }

    /** Runs `fn` with relationships routed to `owner`'s `.rels`, restoring the previous owner after. */
    private async withRelOwner<T>(owner: string, fn: () => Promise<T>): Promise<T> {
        const prev = this.currentRelOwner;
        this.currentRelOwner = owner;
        try { return await fn(); } finally { this.currentRelOwner = prev; }
    }

    /**
     * Mints a fresh, unique bookmark for one anchor. The first claim of a base keeps the base name
     * (so an internal link to it, resolved via {@link anchorName}, matches); later claims of the same
     * base get a `_2`/`_3` suffix and a distinct id, so no two bookmarks share a name or id.
     */
    private mintBookmark(rawName: string): { id: number; name: string } {
        const base = toBookmarkNameRaw(rawName);
        let name = base, i = 2;
        while (this.usedBookmarkNames.has(name)) { const suffix = `_${i++}`; name = base.slice(0, 40 - suffix.length) + suffix; }
        this.usedBookmarkNames.add(name);
        return { id: this.bookmarkCounter++, name };
    }

    /** The bookmark name an internal link resolves to: the sanitized base (matches the first claim). */
    private anchorName(rawName: string): string {
        return toBookmarkNameRaw(rawName);
    }

    /** Packages an image attachment once, returning the relationship id or null (skipped). */
    private mediaRel(attachmentName: string): { rid: string; cx: number; cy: number; intrinsic: { w: number; h: number } | null } | null {
        const att = this.getAttachment(attachmentName);
        if (!att || !att.data) {
            // Warn like the ODT and native-PDF generators do, rather than degrading to alt text silently.
            this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: att ? 'attachment has no data' : 'missing attachment' });
            return null;
        }
        const ext = MIME_EXT[(att.mimeType || '').toLowerCase()];
        if (!ext) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName, reason: 'unsupported mime' }); return null; }
        // Word does not render an SVG referenced by a bare a:blip (it needs the asvg extension + a
        // raster fallback); emit nothing here so the caller degrades to alt text instead of a broken image.
        if (ext === 'svg') { this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'docx', feature: 'svg image' }); return null; }
        let bytes: Uint8Array;
        try { bytes = decodeBase64(att.data); } catch { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: attachmentName }); return null; }
        let name = this.mediaByAttachment.get(attachmentName);
        if (!name) {
            name = `image${this.media.length + 1}.${ext}`;
            const contentType = (att.mimeType || 'image/png');
            this.media.push({ name, bytes, ext, contentType });
            this.mediaByAttachment.set(attachmentName, name);
            this.usedExtensions.add(ext);
        }
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', `media/${name}`);
        return { rid, cx: 0, cy: 0, intrinsic: sniffImageSize(bytes) };
    }

    // ── metadata / reproducibility ─────────────────────────────────────────────

    private resolveModified(): { iso: string; mtime: Date } {
        return resolveZipInstant(this.effectiveMetadata.modified);
    }

    // ── entry point ─────────────────────────────────────────────────────────────

    async generate(): Promise<ConversionResult<'docx'>> {
        this.warnUnrepresentableCustomMetadata('docx');
        const { iso, mtime } = this.resolveModified();

        // Render body blocks (this also populates rels, media, notes, comments, bookmarks, numbering).
        const metaBlock = this.config.renderMetadata ? this.renderMetadataBlock() : '';
        const body = await this.renderBlocks(this.ast.content);

        const hasHeader = !!(this.ast.auxiliary?.headers?.length);
        const hasFooter = !!(this.ast.auxiliary?.footers?.length);
        const headerXml = hasHeader ? await this.withRelOwner('word/header1.xml', () => this.renderHdrFtr('hdr', this.ast.auxiliary!.headers!)) : '';
        const footerXml = hasFooter ? await this.withRelOwner('word/footer1.xml', () => this.renderHdrFtr('ftr', this.ast.auxiliary!.footers!)) : '';
        const headerRid = hasHeader ? this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml') : '';
        const footerRid = hasFooter ? this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml') : '';

        const documentXml = this.buildDocumentXml(metaBlock + body, headerRid, footerRid);
        const parts = this.assembleParts({ documentXml, iso, hasHeader, hasFooter, headerXml, footerXml });

        const zipFiles: Zippable = {};
        for (const [path, content] of Object.entries(parts)) zipFiles[path] = new TextEncoder().encode(content);
        for (const m of this.media) zipFiles[`word/media/${m.name}`] = m.bytes;

        const bytes = zipSync(zipFiles, { mtime });
        return { value: bytes, messages: this.messages };
    }

    // ── block / inline traversal ─────────────────────────────────────────────────

    private async renderBlocks(nodes: OfficeContentNode[] | undefined): Promise<string> {
        let out = '';
        let prevPaginated: string | null = null;
        let prevBareList = false; // previous block was a list carrying no listId
        for (const node of nodes || []) {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) { prevBareList = false; continue; }
            if (typeof override === 'string') { out += override; prevPaginated = null; prevBareList = false; continue; }
            // Preserve source pagination: a page break between consecutive page (or slide) nodes.
            if ((node.type === 'page' || node.type === 'slide') && prevPaginated === node.type) {
                out += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
            }
            // A run of consecutive lists that carry no listId of their own gets one synthetic id, so a
            // later separate list starts fresh instead of continuing the first one's numbering.
            const isBareList = node.type === 'list' && !(node.metadata as any)?.listId;
            if (isBareList && !prevBareList) this.syntheticListRun++;
            out += await this.renderBlockNode(node, isBareList ? `__run${this.syntheticListRun}` : undefined);
            prevPaginated = (node.type === 'page' || node.type === 'slide') ? node.type : null;
            prevBareList = isBareList;
        }
        return out;
    }

    private async renderBlockNode(node: OfficeContentNode, bareListId?: string): Promise<string> {
        switch (node.type) {
            case 'paragraph': return this.paragraph(node);
            case 'heading': return this.paragraph(node, this.headingStyle(node));
            case 'list': return this.listItem(node, bareListId);
            case 'table': return this.table(node);
            case 'image': { const r = await this.imageRun(node); return r ? this.wrapP(r, this.jc((node.metadata as any)?.align)) : ''; }
            case 'code': return this.codeBlock(node);
            case 'break': return this.blockBreak(node);
            case 'note': return this.standaloneNote(node); // orphan note definition: keep it as a trailing note
            case 'comment': return this.standaloneComment(node);
            case 'admonition': return this.admonition(node);
            case 'chart': return this.chart(node);
            case 'embed': return this.embed(node);
            case 'definitionTerm': return this.paragraph(node, 'DefinitionTerm');
            case 'definitionDescription': return this.paragraphIndented(node, 720);
            case 'header': case 'footer': return this.renderBlocks(node.children); // inline page furniture
            case 'sheet': return this.sheet(node);
            case 'row': case 'cell': case 'text': // stray inline/table-part at block level: wrap
                return this.wrapP(await this.renderInline([node]));
            case 'page': case 'slide': case 'drawing': case 'definitionList':
                return this.renderBlocks(node.children); // containers: flatten (page/slide breaks handled by caller sequence)
            case 'slideMaster': return ''; // template furniture, dropped
            default: {
                // Exhaustiveness guard: a new OfficeContentNodeType must be classified here.
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

    private async paragraph(node: OfficeContentNode, styleOverride?: string): Promise<string> {
        const meta = node.metadata as any;
        const semantic = this.getSemanticMapping(node);
        const style = styleOverride ?? this.styleForTag(semantic?.tag) ?? this.knownStyle(meta?.style);
        const pPr = this.buildPPr({ style, meta });
        const bookmarks = this.bookmarksFor(node);
        // Notes/comments the parser attached to the block itself (their ref preceded any run).
        const blockRefs = (await this.noteRefs(node)) + (await this.commentRefs(node));
        const inner = await this.renderInline(node.children && node.children.length ? node.children : [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        return `<w:p>${pPr}${bookmarks.start}${blockRefs}${inner}${bookmarks.end}</w:p>`;
    }

    private async paragraphIndented(node: OfficeContentNode, leftTwips: number): Promise<string> {
        const inner = await this.renderInline(node.children || []);
        return `<w:p><w:pPr><w:ind w:left="${leftTwips}"/></w:pPr>${inner}</w:p>`;
    }

    private buildPPr(opts: { style?: string; meta?: any; numPr?: string; extra?: string }): string {
        const { style, meta, numPr, extra } = opts;
        if (this.config.includeFormatting === false && !style && !numPr && !extra) return '';
        let s = '';
        if (style) s += `<w:pStyle w:val="${escapeXml(style)}"/>`;
        if (numPr) s += numPr;
        if (extra) s += extra;
        if (this.config.includeFormatting !== false && meta) {
            const ind = meta.paragraphIndentation;
            if (ind && (ind.left || ind.right || ind.firstLine || ind.hanging)) {
                const attrs = [
                    ind.left ? `w:left="${ind.left | 0}"` : '',
                    ind.right ? `w:right="${ind.right | 0}"` : '',
                    ind.firstLine ? `w:firstLine="${ind.firstLine | 0}"` : '',
                    ind.hanging ? `w:hanging="${ind.hanging | 0}"` : '',
                ].filter(Boolean).join(' ');
                if (attrs) s += `<w:ind ${attrs}/>`;
            }
            const j = this.jcVal(meta.alignment);
            if (j) s += `<w:jc w:val="${j}"/>`;
        }
        return s ? `<w:pPr>${s}</w:pPr>` : '';
    }

    private bookmarksFor(node: OfficeContentNode): { start: string; end: string } {
        if (this.config.ignoreInternalLinks) return { start: '', end: '' };
        const names: string[] = [...(((node.metadata as any)?.anchorIds) || [])];
        if (this.config.generateIds && node.type === 'heading') {
            const slug = this.slugify(node.text || this.getNodeText(node));
            if (slug && !names.includes(slug)) names.push(slug);
        }
        let start = '', end = '';
        for (const raw of names) {
            const { id, name } = this.mintBookmark(raw);
            start += `<w:bookmarkStart w:id="${id}" w:name="${escapeXml(name)}"/>`;
            end += `<w:bookmarkEnd w:id="${id}"/>`;
        }
        return { start, end };
    }

    private headingStyle(node: OfficeContentNode): string {
        const level = Math.min(6, Math.max(1, (node.metadata as any)?.level || 1));
        return `Heading${level}`;
    }

    private jcVal(alignment: string | undefined): string | null {
        if (alignment === 'center') return 'center';
        if (alignment === 'right') return 'right';
        if (alignment === 'justify') return 'both';
        return null;
    }
    private jc(alignment: string | undefined): string {
        const v = this.jcVal(alignment);
        return v ? `<w:pPr><w:jc w:val="${v}"/></w:pPr>` : '';
    }

    private wrapP(inner: string, pPr = ''): string {
        return `<w:p>${pPr}${inner}</w:p>`;
    }

    // ── inline runs ──────────────────────────────────────────────────────────────

    private async renderInline(nodes: OfficeContentNode[]): Promise<string> {
        // Group consecutive text children that share the same link into one hyperlink wrapper, while
        // still running handleOnNode for every node (a grouped member can be overridden or skipped).
        let out = '';
        let group: OfficeContentNode[] = [];
        let groupLink: any = null;
        let groupType: any = null;
        const flush = async () => { if (group.length) { out += await this.hyperlink(groupLink, groupType, group); group = []; groupLink = null; } };
        for (const node of nodes || []) {
            const override = await this.handleOnNode(node);
            if (override === false) { await flush(); continue; }
            if (typeof override === 'string') { await flush(); out += override; continue; }
            const link = (node.type === 'text') ? (node.metadata as any)?.link : undefined;
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
            case 'text': return this.textRuns(node) + await this.noteRefs(node) + await this.commentRefs(node);
            case 'code': return this.inlineCode(node);
            case 'image': return await this.imageRun(node);
            case 'break': return this.inlineBreak(node);
            case 'paragraph': case 'heading': // block content nested inline (HTML): render its runs
                return this.renderInline(node.children || [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
            default:
                // Anything else nested inline: fall back to its plain text as a run.
                return this.run(this.getNodeText(node), undefined);
        }
    }

    /** Text as run content: newlines become `<w:br/>` and tabs `<w:tab/>`, literal segments escaped. */
    private runContent(text: string): string {
        return text.split(/\r\n|\r|\n/).map((line, li) => {
            const parts = line.split('\t').map(seg => seg ? `<w:t xml:space="preserve">${xmlText(seg)}</w:t>` : '').join('<w:tab/>');
            return (li > 0 ? '<w:br/>' : '') + parts;
        }).join('');
    }

    private run(text: string, fmt: TextFormatting | undefined): string {
        if (!text) return '';
        return `<w:r>${this.buildRPr(fmt)}${this.runContent(text)}</w:r>`;
    }

    private textRuns(node: OfficeContentNode): string {
        return this.run(node.text || '', node.formatting);
    }

    private buildRPr(fmt: TextFormatting | undefined, styleId?: string): string {
        let s = '';
        if (styleId) s += `<w:rStyle w:val="${styleId}"/>`;
        if (fmt && this.config.includeFormatting !== false) {
            if (fmt.font) { const f = escapeXml(stripInvalidXmlChars(fmt.font)); s += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`; }
            if (fmt.bold) s += '<w:b/>';
            if (fmt.italic) s += '<w:i/>';
            if (fmt.strikethrough) s += '<w:strike/>';
            const c = hexColor(fmt.color);
            if (c) s += `<w:color w:val="${c}"/>`;
            const sz = lengthToPt(fmt.size);
            const halfPts = sz && sz > 0 ? Math.round(sz * 2) : 0;
            if (halfPts > 0) s += `<w:sz w:val="${halfPts}"/>`;
            if (fmt.underline) s += '<w:u w:val="single"/>';
            const bg = hexColor(fmt.backgroundColor);
            if (bg) s += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
            if (fmt.superscript) s += '<w:vertAlign w:val="superscript"/>';
            else if (fmt.subscript) s += '<w:vertAlign w:val="subscript"/>';
        }
        return s ? `<w:rPr>${s}</w:rPr>` : '';
    }

    private async hyperlink(link: string, linkType: string | undefined, group: OfficeContentNode[]): Promise<string> {
        const runs = group.map(n => this.styledRun(n.text || '', n.formatting, 'Hyperlink')).join('');
        // Notes/comments anchored on linked text still have to be emitted; place them after the link.
        let trailing = '';
        for (const n of group) trailing += await this.noteRefs(n) + await this.commentRefs(n);
        return (this.hyperlinkAround(link, linkType, runs) ?? group.map(n => this.textRuns(n)).join('')) + trailing;
    }

    /**
     * `runs` inside a hyperlink to `link`: a bookmark for an internal target, a relationship for an
     * external one. Null when the link is not written: an internal one under `ignoreInternalLinks`,
     * or a target the scheme check refuses.
     */
    private hyperlinkAround(link: string, linkType: string | undefined, runs: string): string | null {
        const internal = linkType === 'internal' || link.startsWith('#');
        if (internal) {
            if (this.config.ignoreInternalLinks) return null;
            return `<w:hyperlink w:anchor="${escapeXml(this.anchorName(link.replace(/^#/, '')))}">${runs}</w:hyperlink>`;
        }
        const safe = sanitizeOfficePackageUrl(link);
        if (!safe) return null;
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(encUrl(safe)), 'External');
        return `<w:hyperlink r:id="${rid}">${runs}</w:hyperlink>`;
    }

    private styledRun(text: string, fmt: TextFormatting | undefined, styleId: string): string {
        if (!text) return '';
        return `<w:r>${this.buildRPr(fmt, styleId)}${this.runContent(text)}</w:r>`;
    }

    // ── notes & comments ─────────────────────────────────────────────────────────

    private async noteRefs(node: OfficeContentNode): Promise<string> {
        if (!node.notes || !node.notes.length) return '';
        let out = '';
        for (const note of node.notes) {
            const kind: 'footnote' | 'endnote' = (note.metadata as any)?.noteType === 'endnote' ? 'endnote' : 'footnote';
            out += await this.registerNote(note, kind);
        }
        return out;
    }

    /**
     * Registers a note (once, keyed by content) into the footnotes/endnotes part, rendering its body
     * with relationships routed to that part, and returns the in-text reference run.
     */
    private async registerNote(note: OfficeContentNode, kind: 'footnote' | 'endnote'): Promise<string> {
        const key = this.getFootnoteKey(note);
        if (!this.noteBodies.has(key)) {
            (kind === 'endnote' ? this.endnotes : this.footnotes).push({ key, kind, node: note });
            const owner = kind === 'endnote' ? 'word/endnotes.xml' : 'word/footnotes.xml';
            const body = (await this.withRelOwner(owner, () => this.renderBlocks(this.bodyBlocks(note)))) || '<w:p/>';
            this.noteBodies.set(key, this.withNoteMarker(this.styleNoteBody(body), kind));
        }
        const id = this.noteId(key);
        const tag = kind === 'endnote' ? 'w:endnoteReference' : 'w:footnoteReference';
        return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><${tag} w:id="${id}"/></w:r>`;
    }

    private styleNoteBody(body: string): string {
        // Give each top-level body paragraph the FootnoteText style (only when it has no pPr already).
        return body.replace(/<w:p>(?!<w:pPr>)/g, '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>');
    }

    /** Prepends the numbered marker run (w:footnoteRef/w:endnoteRef) to the note body's first paragraph. */
    private withNoteMarker(body: string, kind: 'footnote' | 'endnote'): string {
        const ref = `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:${kind}Ref/></w:r>`;
        const m = /^<w:p>(<w:pPr>[\s\S]*?<\/w:pPr>)?/.exec(body);
        if (m) return body.slice(0, m[0].length) + ref + body.slice(m[0].length);
        // The body does not open with a paragraph (a note whose first block is a table). A run is not
        // valid block content under <w:footnote>/<w:endnote>, so give the marker a paragraph of its own
        // rather than emitting a bare <w:r> there.
        return `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${ref}</w:p>` + body;
    }

    private noteIdMap = new Map<string, number>();
    private noteIdCounter = 2; // ids -1/0/1 reserved by the separators
    private noteId(key: string): number {
        let id = this.noteIdMap.get(key);
        if (id === undefined) { id = this.noteIdCounter++; this.noteIdMap.set(key, id); }
        return id;
    }

    private async commentRefs(node: OfficeContentNode): Promise<string> {
        if (!node.comments || !node.comments.length) return '';
        let out = '';
        for (const c of node.comments) out += await this.registerComment(c, false);
        return out;
    }

    private async standaloneComment(node: OfficeContentNode): Promise<string> {
        return this.registerComment(node, true);
    }

    /** Registers a comment into comments.xml (body rels routed there) and returns its reference run. */
    private async registerComment(node: OfficeContentNode, ownParagraph: boolean): Promise<string> {
        const id = this.commentCounter++;
        this.comments.push({ id, node });
        const body = (await this.withRelOwner('word/comments.xml', () => this.renderBlocks(this.bodyBlocks(node)))) || '<w:p/>';
        this.commentBodies.set(id, body);
        const ref = `<w:r><w:commentReference w:id="${id}"/></w:r>`;
        return ownParagraph ? `<w:p>${ref}</w:p>` : ref;
    }

    /**
     * An orphan note node (e.g. an unreferenced footnote definition) has no in-text anchor. Keep it by
     * registering it and emitting a trailing paragraph carrying its reference, so it survives round-trip.
     */
    private async standaloneNote(node: OfficeContentNode): Promise<string> {
        const kind: 'footnote' | 'endnote' = (node.metadata as any)?.noteType === 'endnote' ? 'endnote' : 'footnote';
        const ref = await this.registerNote(node, kind);
        return `<w:p>${ref}</w:p>`;
    }

    // ── lists / numbering ────────────────────────────────────────────────────────

    /**
     * Registers the numbering definition for a list item at render time and returns its numId. Done
     * lazily during rendering (not a pre-pass) so lists inside notes, comments, headers and footers -
     * which the content walk never reaches - still get a numbering entry and a numbering.xml part.
     */
    private ensureNumbering(meta: any, bareListId?: string): number {
        const listId = String(meta?.listId ?? bareListId ?? 'pdf-list');
        let numId = this.numByListId.get(listId);
        let entry: { listId: string; numId: number; levels: Map<number, 'ordered' | 'unordered'>; startAt: Map<number, number> };
        if (numId === undefined) {
            numId = this.numbering.length + 1;
            entry = { listId, numId, levels: new Map(), startAt: new Map() };
            this.numbering.push(entry);
            this.numByListId.set(listId, numId);
            this.needsNumbering = true;
        } else {
            entry = this.numbering.find(n => n.numId === numId)!;
        }
        const lvl = Math.max(0, Math.min(8, meta?.indentation | 0));
        if (!entry.levels.has(lvl)) {
            entry.levels.set(lvl, meta?.listType === 'ordered' ? 'ordered' : 'unordered');
            if (typeof meta?.itemIndex === 'number') entry.startAt.set(lvl, meta.itemIndex + 1);
        }
        return numId;
    }

    private async listItem(node: OfficeContentNode, bareListId?: string): Promise<string> {
        const meta = node.metadata as any;
        const numId = this.ensureNumbering(meta, bareListId);
        const lvl = Math.max(0, Math.min(8, meta?.indentation | 0));
        const numPr = `<w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
        const pPr = this.buildPPr({ style: 'ListParagraph', numPr, meta });
        let prefix = '';
        if (meta?.isTask) prefix = this.run(meta.checked ? '☑ ' : '☐ ', undefined);
        const blockRefs = (await this.noteRefs(node)) + (await this.commentRefs(node));
        const inner = await this.renderInline(node.children || [{ type: 'text', text: node.text || '' } as OfficeContentNode]);
        return `<w:p>${pPr}${prefix}${blockRefs}${inner}</w:p>`;
    }

    private buildNumberingXml(): string {
        let abstracts = '', nums = '';
        for (const entry of this.numbering) {
            const aId = entry.numId - 1;
            let levels = '';
            for (let l = 0; l <= 8; l++) {
                const type = entry.levels.get(l) ?? entry.levels.get(0) ?? 'unordered';
                const fmt = type === 'ordered' ? 'decimal' : 'bullet';
                const text = type === 'ordered' ? `%${l + 1}.` : '•';
                levels += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>`
                    + `<w:lvlText w:val="${escapeXml(text)}"/><w:lvlJc w:val="left"/>`
                    + `<w:pPr><w:ind w:left="${720 * (l + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
            }
            abstracts += `<w:abstractNum w:abstractNumId="${aId}"><w:multiLevelType w:val="multilevel"/>${levels}</w:abstractNum>`;
            let overrides = '';
            for (const [lvl, start] of entry.startAt) {
                if (start > 1) overrides += `<w:lvlOverride w:ilvl="${lvl}"><w:startOverride w:val="${start}"/></w:lvlOverride>`;
            }
            nums += `<w:num w:numId="${entry.numId}"><w:abstractNumId w:val="${aId}"/>${overrides}</w:num>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstracts}${nums}</w:numbering>`;
    }

    // ── tables ───────────────────────────────────────────────────────────────────

    private async table(node: OfficeContentNode): Promise<string> {
        const rows = (node.children || []).filter(r => r.type === 'row');
        if (!rows.length) return '';
        // Grid-occupancy pass: compute grid width and per-row rendered cells with synthesized merges.
        const cols = this.gridWidth(rows);
        const contentWidth = this.contentWidthTwips();
        const colW = Math.max(1, Math.floor(contentWidth / Math.max(1, cols)));
        const tblGrid = `<w:tblGrid>${Array.from({ length: cols }, () => `<w:gridCol w:w="${colW}"/>`).join('')}</w:tblGrid>`;
        // Table-level justification only accepts left/center/right; 'both' (justify) is not valid here.
        const tableAlign = this.jcVal((node.metadata as any)?.align) === 'both' ? null : this.jcVal((node.metadata as any)?.align);
        const tblPr = `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>`
            + (tableAlign ? `<w:jc w:val="${tableAlign}"/>` : '')
            + `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('')}</w:tblBorders></w:tblPr>`;

        // grid left-col -> a pending vertical merge (rows still to cover, and the merge's column span).
        const active = new Map<number, { remaining: number; span: number }>();
        let trs = '';
        for (let ri = 0; ri < rows.length; ri++) {
            const row = rows[ri];
            const cells = (row.children || []).filter(c => c.type === 'cell');
            const trPr = isHeaderRow(row, ri === 0) ? '<w:trPr><w:tblHeader/></w:trPr>' : '';
            let tcs = '';
            let col = 0, ci = 0;
            while (ci < cells.length || [...active.keys()].some(c => c >= col)) {
                const act = active.get(col);
                if (act) {
                    // Continuation cell for an active vertical merge. Carry the origin's gridSpan so a cell
                    // merged BOTH across columns and down emits one spanning continuation, not one narrow
                    // vMerge per column (which made the merge cover only its first column in Word).
                    const gs = act.span > 1 ? `<w:gridSpan w:val="${act.span}"/>` : '';
                    tcs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>${gs}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
                    act.remaining--;
                    if (act.remaining <= 0) active.delete(col);
                    col += act.span;
                    continue;
                }
                if (ci >= cells.length) {
                    // No explicit cells left. A vertical merge is still pending at a later column, so
                    // fill this gap column with an empty cell and advance until the merges are placed;
                    // breaking here would drop the continuation and shift the grid a row down.
                    if (![...active.keys()].some(c => c > col)) break;
                    tcs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p/></w:tc>`;
                    col++;
                    continue;
                }
                // Sparse source grid: ExcelParser emits only the non-empty cells, each carrying its own
                // column index. Fill the skipped columns with empty cells so a value in D1 lands in
                // column 4, rather than sliding left to whatever the running cursor happened to reach.
                const nextCol = (cells[ci].metadata as any)?.col;
                if (typeof nextCol === 'number' && nextCol > col && col < cols) {
                    tcs += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p/></w:tc>`;
                    col++;
                    continue;
                }
                const cell = cells[ci++];
                const cmeta = cell.metadata as any;
                const colSpan = Math.max(1, Math.min(cols, cmeta?.colSpan || 1));
                const rowSpan = Math.max(1, Math.min(1000, cmeta?.rowSpan || 1));
                let tcPr = `<w:tcW w:w="0" w:type="auto"/>`;
                if (colSpan > 1) tcPr += `<w:gridSpan w:val="${colSpan}"/>`;
                if (rowSpan > 1) { tcPr += `<w:vMerge w:val="restart"/>`; active.set(col, { remaining: rowSpan - 1, span: colSpan }); }
                const bg = hexColor(cmeta?.backgroundColor);
                if (bg) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
                let inner = await this.renderBlocks(cell.children);
                // A w:tc must end with a w:p: append one only when empty or ending in a nested table.
                const trimmed = inner.trimEnd();
                if (!trimmed) inner = '<w:p/>';
                else if (trimmed.endsWith('</w:tbl>')) inner = trimmed + '<w:p/>';
                tcs += `<w:tc><w:tcPr>${tcPr}</w:tcPr>${inner}</w:tc>`;
                col += colSpan;
            }
            // A w:tr must contain at least one w:tc.
            if (!tcs) tcs = `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p/></w:tc>`;
            trs += `<w:tr>${trPr}${tcs}</w:tr>`;
        }
        return `<w:tbl>${tblPr}${tblGrid}${trs}</w:tbl>`;
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

    private async sheet(node: OfficeContentNode): Promise<string> {
        const name = (node.metadata as any)?.sheetName;
        const heading = name ? `<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t xml:space="preserve">${xmlText(name)}</w:t></w:r></w:p>` : '';
        const children = node.children || [];
        const table = await this.table({ type: 'table', children: fillSheetRowGaps(children.filter(c => c.type === 'row')) } as OfficeContentNode);
        // A sheet's drawing images and charts are pushed as non-row children after the rows; render
        // them after the grid (the HTML generator does the same) rather than dropping them.
        const extras = await this.renderBlocks(children.filter(c => c.type !== 'row'));
        return heading + table + extras;
    }

    // ── images ───────────────────────────────────────────────────────────────────

    private async imageRun(node: OfficeContentNode): Promise<string> {
        const mode: ImageMode = this.imageMode();
        if (mode === 'none') return '';
        const meta = node.metadata as any;
        const ocr = node.text || '';
        if (mode === 'ocr-text-only') return ocr ? this.run(ocr, undefined) : '';

        const rel = meta?.attachmentName ? this.mediaRel(meta.attachmentName) : null;
        let drawing = '';
        if (rel) {
            const { cx, cy } = this.imageEmu(node, rel.intrinsic);
            const id = ++this.drawingCounter; // unique per emitted drawing (image reuse must not repeat ids)
            const alt = xmlText(meta?.altText || '');
            drawing = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
                + `<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="image${id}" descr="${alt}"/>`
                + `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`
                + `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`
                + `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`
                + `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr>`
                + `<pic:blipFill><a:blip r:embed="${rel.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
                + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
                + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
            // A picture that is itself a link (ImageMetadata.link) wraps in a hyperlink run container,
            // exactly as ODT wraps the frame in <draw:a>.
            if (meta?.link) drawing = this.hyperlinkAround(meta.link, meta.linkType, drawing) ?? drawing;
        } else if (meta?.url) {
            // Remote-only image: degrade to a link on the alt text (never fetch bytes: SSRF), to where
            // the picture links when it is a link, and to the image otherwise.
            const safe = sanitizeOfficePackageUrl(meta.url);
            if (safe) {
                const run = this.styledRun(meta.altText || safe, undefined, 'Hyperlink');
                drawing = (meta.link ? this.hyperlinkAround(meta.link, meta.linkType, run) : null)
                    ?? `<w:hyperlink r:id="${this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(encUrl(safe)), 'External')}">${run}</w:hyperlink>`;
            }
        }
        // No renderable image (unresolvable/unsupported attachment, no url): keep the alt text or OCR
        // so the content is not silently lost.
        if (!drawing) { const fallback = meta?.altText || ocr; return fallback ? this.run(fallback, undefined) : ''; }
        if (mode === 'image+ocr-text' && ocr) return drawing + this.run('\n' + ocr, undefined);
        return drawing;
    }

    private imageEmu(node: OfficeContentNode, intrinsic: { w: number; h: number } | null): { cx: number; cy: number } {
        const contentWidthEmu = Math.round(this.contentWidthTwips() / TWIPS_PER_PT * EMU_PER_PT);
        const meta = node.metadata as any;
        // 1. explicit width
        if (meta?.width) {
            const pct = /^\s*([\d.]+)\s*%\s*$/.exec(meta.width);
            let cx: number | null = null;
            if (pct) cx = Math.round(contentWidthEmu * Math.min(100, parseFloat(pct[1])) / 100);
            else { const pt = lengthToPt(meta.width); if (pt) cx = Math.round(pt * EMU_PER_PT); }
            if (cx) { const ar = intrinsic ? intrinsic.h / intrinsic.w : 0.75; return this.capEmu(cx, Math.round(cx * ar), contentWidthEmu); }
        }
        // 2. PDF-sourced on-page bounds (points)
        if (node.bounds && node.bounds.width > 0 && node.bounds.height > 0) {
            return this.capEmu(Math.round(node.bounds.width * EMU_PER_PT), Math.round(node.bounds.height * EMU_PER_PT), contentWidthEmu);
        }
        // 3. intrinsic pixels at 96dpi
        if (intrinsic) {
            const cx = Math.round(intrinsic.w / 96 * EMU_PER_IN);
            const cy = Math.round(intrinsic.h / 96 * EMU_PER_IN);
            return this.capEmu(cx, cy, contentWidthEmu);
        }
        // 4. fallback 3in x 2.25in
        return { cx: 3 * EMU_PER_IN, cy: Math.round(2.25 * EMU_PER_IN) };
    }
    private capEmu(cx: number, cy: number, maxCx: number): { cx: number; cy: number } {
        if (cx > maxCx && cx > 0) { cy = Math.round(cy * maxCx / cx); cx = maxCx; }
        return { cx: Math.max(1, cx), cy: Math.max(1, cy) };
    }

    private contentWidthTwips(): number {
        const cfg = this.config.docxConfig;
        const size = paperSizePt(cfg.format);
        const w = Math.round((cfg.landscape ? size.h : size.w) * TWIPS_PER_PT);
        const left = Math.round(marginPt(cfg.margin.left) * TWIPS_PER_PT);
        const right = Math.round(marginPt(cfg.margin.right) * TWIPS_PER_PT);
        return Math.max(720, w - left - right);
    }

    // ── breaks / code / admonition / chart / embed ────────────────────────────────

    private blockBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
        if (t === 'column') return `<w:p><w:r><w:br w:type="column"/></w:r></w:p>`;
        if (t === 'thematic') return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
        if (t === 'lastRenderedPage') return '';
        return `<w:p><w:r><w:br/></w:r></w:p>`;
    }
    private inlineBreak(node: OfficeContentNode): string {
        const t = (node.metadata as any)?.breakType;
        if (t === 'page') return `<w:r><w:br w:type="page"/></w:r>`;
        if (t === 'column') return `<w:r><w:br w:type="column"/></w:r>`;
        if (t === 'carriageReturn') return `<w:r><w:cr/></w:r>`;
        return `<w:r><w:br/></w:r>`;
    }

    private codeBlock(node: OfficeContentNode): string {
        if ((node.metadata as any)?.math && !this.mathWarned) { this.mathWarned = true; this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'docx', feature: 'math' }); }
        const text = node.text || this.getNodeText(node);
        const segs = text.split(/\r\n|\r|\n/);
        const runs = segs.map((seg, i) => `${i > 0 ? '<w:br/>' : ''}${seg ? `<w:t xml:space="preserve">${xmlText(seg)}</w:t>` : ''}`).join('');
        return `<w:p><w:pPr><w:pStyle w:val="Code"/></w:pPr><w:r>${runs}</w:r></w:p>`;
    }
    private inlineCode(node: OfficeContentNode): string {
        return `<w:r><w:rPr><w:rStyle w:val="CodeChar"/></w:rPr><w:t xml:space="preserve">${xmlText(node.text || '')}</w:t></w:r>`;
    }

    private async admonition(node: OfficeContentNode): Promise<string> {
        const meta = node.metadata as any;
        const color = ADMONITION_COLOR[meta?.admonitionType] || '0969DA';
        const title = meta?.title || (meta?.admonitionType ? meta.admonitionType[0].toUpperCase() + meta.admonitionType.slice(1) : 'Note');
        const titleP = `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="${color}"/></w:pBdr><w:ind w:left="360"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="${color}"/></w:rPr><w:t xml:space="preserve">${xmlText(title)}</w:t></w:r></w:p>`;
        const body = await this.renderBlocks(node.children);
        // Indent the body paragraphs with a left border to visually box the admonition.
        const boxed = body.replace(/<w:p>(?!<w:pPr>)/g, `<w:p><w:pPr><w:pBdr><w:left w:val="single" w:sz="18" w:space="4" w:color="${color}"/></w:pBdr><w:ind w:left="360"/></w:pPr>`);
        return titleP + boxed;
    }

    private async chart(node: OfficeContentNode): Promise<string> {
        if (this.config.includeCharts === false) return '';
        const meta = node.metadata as any;
        const att = this.getAttachment(meta?.attachmentName);
        const data = att?.chartData;
        if (!data) return `<w:p><w:r><w:t xml:space="preserve">[Chart: ${xmlText(meta?.attachmentName || '')}]</w:t></w:r></w:p>`;
        const caption = data.title ? `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${xmlText(data.title)}</w:t></w:r></w:p>` : '';
        // Build a table: header = series names, first col = labels.
        const rows: OfficeContentNode[] = [];
        const header: OfficeContentNode = { type: 'row', children: [cellOf(''), ...data.dataSets.map(d => cellOf(d.name || ''))] };
        rows.push(header);
        (data.labels || []).forEach((label, i) => {
            rows.push({ type: 'row', children: [cellOf(label), ...data.dataSets.map(d => cellOf(String(d.values?.[i] ?? '')))] });
        });
        const table = await this.table({ type: 'table', children: rows } as OfficeContentNode);
        return caption + table;
    }

    private embed(node: OfficeContentNode): string {
        const meta = node.metadata as any;
        const rawUrl = embedUrl(meta);
        const url = rawUrl ? sanitizeOfficePackageUrl(rawUrl) : '';
        if (!url) {
            this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { format: 'docx', feature: 'embed' });
            const fallback = meta?.label || node.text || this.getNodeText(node);
            return fallback ? `<w:p>${this.styledRun(fallback, undefined, '')}</w:p>` : '';
        }
        const rid = this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', escapeXml(encUrl(url)), 'External');
        return `<w:p><w:hyperlink r:id="${rid}">${this.styledRun(meta?.label || url, undefined, 'Hyperlink')}</w:hyperlink></w:p>`;
    }

    // ── styleMap helpers ─────────────────────────────────────────────────────────

    private styleForTag(tag: string | undefined): string | undefined {
        if (!tag || !isSafeStyleMapTag(tag)) return undefined;
        const map: Record<string, string> = lookupTable({
            h1: 'Heading1', h2: 'Heading2', h3: 'Heading3', h4: 'Heading4', h5: 'Heading5', h6: 'Heading6',
            blockquote: 'Quote', pre: 'Code',
        });
        return map[tag.toLowerCase()];
    }
    private knownStyle(style: string | undefined): string | undefined {
        const known = new Set(['Quote', 'IntenseQuote', 'Title', 'ListParagraph', 'Code', 'DefinitionTerm']);
        return style && known.has(style) ? style : undefined;
    }

    // ── header / footer / metadata block ─────────────────────────────────────────

    private async renderHdrFtr(kind: 'hdr' | 'ftr', nodes: OfficeContentNode[]): Promise<string> {
        const inner = await this.renderBlocks(nodes) || '<w:p/>';
        const tag = kind === 'hdr' ? 'w:hdr' : 'w:ftr';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<${tag} ${WML_NS}>${inner}</${tag}>`;
    }

    private renderMetadataBlock(): string {
        const m = this.effectiveMetadata;
        let out = '';
        if (m.title) out += `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t xml:space="preserve">${xmlText(m.title)}</w:t></w:r></w:p>`;
        const modIso = toW3CDTF(m.modified);
        const bits = [m.author, modIso ? modIso.slice(0, 10) : undefined].filter(Boolean).join(' · ');
        if (bits) out += `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">${xmlText(bits)}</w:t></w:r></w:p>`;
        return out;
    }

    // ── part assembly ─────────────────────────────────────────────────────────────

    private buildDocumentXml(body: string, headerRid: string, footerRid: string): string {
        const cfg = this.config.docxConfig;
        const size = paperSizePt(cfg.format);
        const pw = Math.round((cfg.landscape ? size.h : size.w) * TWIPS_PER_PT);
        const ph = Math.round((cfg.landscape ? size.w : size.h) * TWIPS_PER_PT);
        const orient = cfg.landscape ? ' w:orient="landscape"' : '';
        const mar = cfg.margin;
        const pgMar = `<w:pgMar w:top="${Math.round(marginPt(mar.top) * TWIPS_PER_PT)}" w:right="${Math.round(marginPt(mar.right) * TWIPS_PER_PT)}" `
            + `w:bottom="${Math.round(marginPt(mar.bottom) * TWIPS_PER_PT)}" w:left="${Math.round(marginPt(mar.left) * TWIPS_PER_PT)}" w:header="708" w:footer="708" w:gutter="0"/>`;
        const sectPr = `<w:sectPr>`
            + (headerRid ? `<w:headerReference w:type="default" r:id="${headerRid}"/>` : '')
            + (footerRid ? `<w:footerReference w:type="default" r:id="${footerRid}"/>` : '')
            + `<w:pgSz w:w="${pw}" w:h="${ph}"${orient}/>${pgMar}</w:sectPr>`;
        const bodyContent = body.trim() ? body : '<w:p/>';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<w:document ${WML_NS}>`
            + `<w:body>${bodyContent}${sectPr}</w:body></w:document>`;
    }

    private assembleParts(o: { documentXml: string; iso: string; hasHeader: boolean; hasFooter: boolean; headerXml: string; footerXml: string }): Record<string, string> {
        const parts: Record<string, string> = {};
        const hasFootnotes = this.footnotes.length > 0;
        const hasEndnotes = this.endnotes.length > 0;
        const hasComments = this.comments.length > 0;

        // Fixed-part relationships (styles + numbering + notes + comments), added to the doc rels.
        this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'styles.xml');
        if (this.needsNumbering) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', 'numbering.xml');
        if (hasFootnotes) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes', 'footnotes.xml');
        if (hasEndnotes) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes', 'endnotes.xml');
        if (hasComments) this.addRel('http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments', 'comments.xml');

        const custom = this.customProps();
        parts['[Content_Types].xml'] = this.contentTypes({ hasFootnotes, hasEndnotes, hasComments, hasCustom: !!custom });
        parts['_rels/.rels'] = this.rootRels(!!custom);
        parts['word/document.xml'] = o.documentXml;
        parts['word/styles.xml'] = STYLES_XML;
        if (this.needsNumbering) parts['word/numbering.xml'] = this.buildNumberingXml();
        if (hasFootnotes) parts['word/footnotes.xml'] = this.buildNotesXml('footnote', this.footnotes);
        if (hasEndnotes) parts['word/endnotes.xml'] = this.buildNotesXml('endnote', this.endnotes);
        if (hasComments) parts['word/comments.xml'] = this.buildCommentsXml();
        if (o.hasHeader) parts['word/header1.xml'] = o.headerXml;
        if (o.hasFooter) parts['word/footer1.xml'] = o.footerXml;
        parts['docProps/core.xml'] = this.coreProps(o.iso);
        parts['docProps/app.xml'] = APP_XML;
        if (custom) parts['docProps/custom.xml'] = custom;
        // One `.rels` per part that actually carries relationships (r:id/r:embed resolve per-part).
        for (const [owner, bucket] of this.partRels) {
            if (bucket.length) parts[this.relsPathFor(owner)] = this.relsXml(bucket);
        }
        return parts;
    }

    /** Maps an owning part path (`word/footnotes.xml`) to its relationships part (`word/_rels/footnotes.xml.rels`). */
    private relsPathFor(owner: string): string {
        const slash = owner.lastIndexOf('/');
        const dir = slash >= 0 ? owner.slice(0, slash) : '';
        const base = slash >= 0 ? owner.slice(slash + 1) : owner;
        return `${dir ? dir + '/' : ''}_rels/${base}.rels`;
    }

    private relsXml(bucket: Rel[]): string {
        const rels = bucket.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${r.target}"${r.mode ? ` TargetMode="${r.mode}"` : ''}/>`).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
    }

    /** Package-level relationships. The custom-properties relationship is included only when that part exists. */
    private rootRels(hasCustom: boolean): string {
        let rels = `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`
            + `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
            + `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>`;
        if (hasCustom) rels += `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>`;
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
    }

    private contentTypes(o: { hasFootnotes: boolean; hasEndnotes: boolean; hasComments: boolean; hasCustom: boolean }): string {
        const wml = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
        let defaults = `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`;
        for (const ext of this.usedExtensions) {
            const ct = ext === 'jpg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : ext === 'tiff' ? 'image/tiff' : `image/${ext}`;
            defaults += `<Default Extension="${ext}" ContentType="${ct}"/>`;
        }
        let overrides = `<Override PartName="/word/document.xml" ContentType="${wml}.document.main+xml"/>`
            + `<Override PartName="/word/styles.xml" ContentType="${wml}.styles+xml"/>`;
        if (this.needsNumbering) overrides += `<Override PartName="/word/numbering.xml" ContentType="${wml}.numbering+xml"/>`;
        if (o.hasFootnotes) overrides += `<Override PartName="/word/footnotes.xml" ContentType="${wml}.footnotes+xml"/>`;
        if (o.hasEndnotes) overrides += `<Override PartName="/word/endnotes.xml" ContentType="${wml}.endnotes+xml"/>`;
        if (o.hasComments) overrides += `<Override PartName="/word/comments.xml" ContentType="${wml}.comments+xml"/>`;
        if (this.ast.auxiliary?.headers?.length) overrides += `<Override PartName="/word/header1.xml" ContentType="${wml}.header+xml"/>`;
        if (this.ast.auxiliary?.footers?.length) overrides += `<Override PartName="/word/footer1.xml" ContentType="${wml}.footer+xml"/>`;
        overrides += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`
            + `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`;
        if (o.hasCustom) overrides += `<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`;
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`;
    }

    private buildNotesXml(kind: 'footnote' | 'endnote', notes: NoteEntry[]): string {
        const tag = kind === 'footnote' ? 'w:footnote' : 'w:endnote';
        const root = kind === 'footnote' ? 'w:footnotes' : 'w:endnotes';
        const sepTag = kind === 'footnote' ? 'w:footnote' : 'w:endnote';
        let entries = `<${sepTag} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></${sepTag}>`
            + `<${sepTag} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></${sepTag}>`;
        // Render each note body synchronously via a cached string (children already rendered in a pre-pass not available;
        // so render here). Since renderBlocks is async, we handle it in the caller; store pre-rendered bodies.
        for (const n of notes) {
            const id = this.noteId(n.key);
            const body = (this.noteBodies.get(n.key)) ?? '<w:p/>';
            entries += `<${tag} w:id="${id}">${body}</${tag}>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${root} ${WML_NS}>${entries}</${root}>`;
    }

    private noteBodies = new Map<string, string>();
    private buildCommentsXml(): string {
        let entries = '';
        for (const c of this.comments) {
            const meta = c.node.metadata as any;
            const body = this.commentBodies.get(c.id) ?? '<w:p/>';
            // w:author is required; w:date must be a valid dateTime (omit if absent/garbage); w:initials optional.
            const dateVal = toW3CDTF(meta?.date);
            const dateAttr = dateVal ? ` w:date="${dateVal}"` : '';
            const initialsAttr = meta?.initials ? ` w:initials="${xmlText(String(meta.initials))}"` : '';
            entries += `<w:comment w:id="${c.id}" w:author="${xmlText(meta?.author || '')}"${dateAttr}${initialsAttr}>${body}</w:comment>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:comments ${WML_NS}>${entries}</w:comments>`;
    }
    private commentBodies = new Map<number, string>();

    private coreProps(iso: string): string {
        const m = this.effectiveMetadata;
        const created = toW3CDTF(m.created) ?? iso;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" `
            + `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
            + el('dc:title', m.title) + el('dc:creator', m.author) + el('dc:subject', m.subject)
            + el('dc:description', m.description) + el('dc:language', (m as any).language)
            + el('cp:keywords', m.keywords) + el('cp:lastModifiedBy', m.lastModifiedBy)
            + `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>`
            + `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`;
    }

    private customProps(): string | null {
        const props = this.effectiveMetadata.customProperties;
        if (!props || !Object.keys(props).length) return null;
        let entries = '', pid = 2;
        for (const [key, value] of Object.entries(props)) {
            let v: string;
            if (typeof value === 'boolean') v = `<vt:bool>${value ? 'true' : 'false'}</vt:bool>`;
            else if (typeof value === 'number') v = `<vt:r8>${value}</vt:r8>`;
            else if (value instanceof Date) v = `<vt:filetime>${value.toISOString().replace(/\.\d+Z$/, 'Z')}</vt:filetime>`;
            else v = `<vt:lpwstr>${xmlText(String(value))}</vt:lpwstr>`;
            entries += `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${pid++}" name="${xmlText(key)}">${v}</property>`;
        }
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" `
            + `xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${entries}</Properties>`;
    }
}

function cellOf(text: string): OfficeContentNode {
    return { type: 'cell', text, children: [{ type: 'paragraph', children: [{ type: 'text', text } as OfficeContentNode] } as OfficeContentNode], metadata: { row: 0, col: 0 } as any };
}

const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>officeParser</Application></Properties>`;

const STYLES_XML = buildStyles();
function buildStyles(): string {
    const heading = (n: number, sz: number) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
        + `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`
        + `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>`
        + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`
        + heading(1, 32) + heading(2, 28) + heading(3, 26) + heading(4, 24) + heading(5, 22) + heading(6, 22)
        + `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="IntenseQuote"><w:name w:val="Intense Quote"/><w:basedOn w:val="Quote"/><w:qFormat/><w:rPr><w:i/><w:b/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:contextualSpacing/></w:pPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="CodeChar"><w:name w:val="Code Char"/><w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="DefinitionTerm"><w:name w:val="Definition Term"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`
        + `<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="Footnote Reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="Footnote Text"/><w:basedOn w:val="Normal"/></w:style>`
        + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>`
        + `</w:styles>`;
}
