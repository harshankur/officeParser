import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { PRESET_TO_ODF } from '../utils/shapeNames.js';
import { decodeBase64, encUrl, hexColor, odfUserDefinedXml, resolveZipInstant, sniffImageSize, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildDeck, bulletMargin, Box, DeckItem, DeckParagraph, DeckRun, DeckSlide, unseenTable } from './slides/deck.js';


/** The pictures ODP output holds, by their types' names (EMF, WMF, TIFF and SVG too, as LibreOffice shows them), and each kind's type. */
const PICTURE_TYPES = new Map<string, string>([['image/png', 'png'], ['image/jpeg', 'jpg'], ['image/jpg', 'jpg'], ['image/gif', 'gif'], ['image/bmp', 'bmp'], ['image/tiff', 'tiff'],
    ['image/x-emf', 'emf'], ['image/emf', 'emf'], ['image/x-wmf', 'wmf'], ['image/wmf', 'wmf'], ['image/svg+xml', 'svg']]);
const PICTURE_MIME = new Map<string, string>([['png', 'image/png'], ['jpg', 'image/jpeg'], ['gif', 'image/gif'], ['bmp', 'image/bmp'], ['tiff', 'image/tiff'], ['emf', 'image/x-emf'], ['wmf', 'image/x-wmf'], ['svg', 'image/svg+xml']]);
/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** A length in points, as ODF writes one. */
/** Points as ODF writes them: finite, within a million points either way (a page's box no slide holds is held to it). */
const pt = (n: number) => `${Number.isFinite(n) ? Math.round(Math.max(-1e6, Math.min(1e6, n)) * 100) / 100 : 0}pt`;

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
    + 'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
    + 'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" '
    + 'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" '
    + 'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" '
    + 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

/** DrawingML's preset outlines ODF draws as its own shapes. */

/**
 * Generates ODP presentations from an AST, as PptxGenerator does PPTX (see buildDeck): a deck's slides
 * as they are, a PDF's pages a slide each with each paragraph a text frame in its place (and a
 * background picture a page where `odpConfig.backgrounds` gives one), a flowing document's headings a
 * slide each.
 */
export class OdpGenerator extends BaseGenerator<'odp'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'odp'>) {
        super('odp', ast, config);
    }

    private readonly styles = new Map<string, { name: string; xml: string }>();
    private readonly media: { path: string; bytes: Uint8Array; mime: string }[] = [];
    private readonly mediaByKey = new Map<string, number>();
    private readonly attachmentsByName = new Map<string, OfficeAttachment>();

    /** Reports, once, a table's cells that found no place in its grid. */
    private pastGridReported = false;
    private cellsPastGrid(): void {
        if (this.pastGridReported) return;
        this.pastGridReported = true;
        this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'table cells past their table\'s columns', format: 'odp' });
    }

    private styleNamed(key: string, prefix: string, xml: (name: string) => string): string {
        let style = this.styles.get(key);
        if (!style) {
            const name = `${prefix}${this.styles.size + 1}`;
            style = { name, xml: xml(name) };
            this.styles.set(key, style);
        }
        return style.name;
    }

    async generate(): Promise<ConversionResult<'odp'>> {
        // onNode is asked once for each node, however often the deck consults it.
        this.onNodeMemo = new WeakMap();
        try { return await this.generateDeck(); } finally { this.onNodeMemo = null; }
    }

    private async generateDeck(): Promise<ConversionResult<'odp'>> {
        for (const a of this.ast.attachments ?? []) if (a?.name) this.attachmentsByName.set(a.name, a);
        const deck = await buildDeck(this.ast.content, this.ast.metadata ?? {}, {
            // Every node a slide takes is asked for here, so the signal is heard node by node.
            keep: async node => { checkAbortSignal(this.config.abortSignal); const v = await this.handleOnNode(node); return v === false ? false : typeof v === 'string' ? v : true; },
            includeCharts: this.config.includeCharts !== false,
            imageMode: this.imageMode(),
            includeFormatting: this.config.includeFormatting !== false,
            chartTable: node => {
                const data = this.attachmentsByName.get((node.metadata as { attachmentName?: string } | undefined)?.attachmentName ?? '')?.chartData;
                return data ? this.chartTable(data, text => ({ type: 'cell', text, children: [{ type: 'text', text }] } as OfficeContentNode), true) : undefined;
            },
        }, { backgrounds: this.config.odpConfig.backgrounds ?? undefined });
        for (const feature of deck.unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'odp' });
        const pages = deck.slides.map((slide, i) => this.pageXml(slide, i, deck.width, deck.height)).join('');
        const enc = (s: string) => new TextEncoder().encode(s);
        const m = this.effectiveMetadata;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        const { iso, mtime } = resolveZipInstant(m.modified);
        const files: Zippable = {
            mimetype: [enc('application/vnd.oasis.opendocument.presentation'), { level: 0 }],
            'content.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${ODF_NS} office:version="1.2">`
                + `<office:automatic-styles>${[...this.styles.values()].map(s => s.xml).join('')}</office:automatic-styles>`
                + `<office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`),
            'styles.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${ODF_NS} office:version="1.2">`
                // Text in a sans face at 18pt where it says no other (a reader's default is a serif).
                + `<office:styles><style:default-style style:family="graphic"><style:text-properties fo:font-family="Calibri" style:font-family-generic="swiss" fo:font-size="18pt" style:font-size-asian="18pt" style:font-size-complex="18pt"/></style:default-style>`
                + `<style:style style:name="standard" style:family="graphic"><style:graphic-properties draw:stroke="none" draw:fill="none"/></style:style></office:styles>`
                + `<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0pt" fo:margin-bottom="0pt" fo:margin-left="0pt" fo:margin-right="0pt" fo:page-width="${pt(deck.width)}" fo:page-height="${pt(deck.height)}" style:print-orientation="${deck.width > deck.height ? 'landscape' : 'portrait'}"/></style:page-layout>`
                + `<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:background-size="border" draw:fill="none"/></style:style></office:automatic-styles>`
                + `<office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1"/></office:master-styles></office:document-styles>`),
            'meta.xml': enc(`<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta ${ODF_NS} office:version="1.2"><office:meta>`
                + el('dc:title', m.title) + el('meta:initial-creator', m.author) + el('dc:creator', m.lastModifiedBy || m.author) + el('dc:subject', m.subject) + el('dc:description', m.description)
                + el('dc:language', m.language) + String(m.keywords ?? '').split(',').map(k => k.trim()).filter(Boolean).map(k => `<meta:keyword>${xmlText(k)}</meta:keyword>`).join('')
                + odfUserDefinedXml(m.customProperties)
                + `<meta:creation-date>${toW3CDTF(m.created) ?? iso}</meta:creation-date><dc:date>${iso}</dc:date><meta:generator>officeParser</meta:generator></office:meta></office:document-meta>`),
        };
        for (const pic of this.media) files[pic.path] = pic.bytes;
        files['META-INF/manifest.xml'] = enc(`<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">`
            + `<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.presentation"/>`
            + ['content.xml', 'styles.xml', 'meta.xml'].map(p => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="text/xml"/>`).join('')
            + this.media.map(pic => `<manifest:file-entry manifest:full-path="${pic.path}" manifest:media-type="${pic.mime}"/>`).join('')
            + `</manifest:manifest>`);
        return { value: zipSync(files, { mtime }), messages: this.messages };
    }

    /** A picture's place in the package, added once, or undefined where it is none a slide shows. */
    private mediaFor(key: string, bytes: Uint8Array | undefined, mime: string | undefined): string | undefined {
        let at = this.mediaByKey.get(key);
        if (at === undefined) {
            if (!bytes || !bytes.length) return undefined;
            // (Its bytes say what it is first: EMF by its signature at 40, a placeable WMF by its key, TIFF either
            // way round, SVG by its element.)
            const head = String.fromCharCode(...bytes.subarray(0, 256));
            const sniffed = bytes[0] === 0x89 && bytes[1] === 0x50 ? 'png' : bytes[0] === 0xff && bytes[1] === 0xd8 ? 'jpg' : bytes[0] === 0x47 && bytes[1] === 0x49 ? 'gif' : bytes[0] === 0x42 && bytes[1] === 0x4d ? 'bmp'
                : (bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0) || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a) ? 'tiff'
                : bytes[0] === 1 && bytes[1] === 0 && bytes[2] === 0 && bytes[3] === 0 && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46 ? 'emf'
                : bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a ? 'wmf'
                : /^\s*(?:<\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*(?:<!DOCTYPE[^>]*>\s*)?<svg[\s>]/i.test(head) ? 'svg' : undefined;
            // (A type named by the attachment, looked up as a type alone: one that is a name every object has is none.)
            const ext = sniffed ?? (mime ? PICTURE_TYPES.get(mime.toLowerCase()) : undefined);
            if (!ext) return undefined;
            at = this.media.length;
            this.media.push({ path: `Pictures/image${at + 1}.${ext}`, bytes, mime: PICTURE_MIME.get(ext)! });
            this.mediaByKey.set(key, at);
        }
        return this.media[at].path;
    }

    /** A slide: its page (background colour or picture), its frames, its notes. */
    private pageXml(slide: DeckSlide, index: number, width: number, height: number): string {
        let pageStyle = '', under = '';
        if (slide.picture) {
            // A background picture: a frame the page's size, under the rest.
            const href = this.mediaFor(`background:${index}`, slide.picture, undefined);
            if (href) under = `<draw:frame svg:x="0pt" svg:y="0pt" svg:width="${pt(width)}" svg:height="${pt(height)}"><draw:image xlink:href="${xmlText(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
            else this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { reason: 'a background picture of a kind ODP output does not hold' });
        } else if (slide.background && hexColor(slide.background)) {
            pageStyle = this.styleNamed(`dp|${hexColor(slide.background)}`, 'dp', name => `<style:style style:name="${name}" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#${hexColor(slide.background)}" draw:background-size="border"/></style:style>`);
        }
        const items = slide.items.map(item => this.itemXml(item)).join('');
        const notes = slide.notes?.length ? `<presentation:notes><draw:frame svg:x="2cm" svg:y="14cm" svg:width="17cm" svg:height="10cm" presentation:class="notes"><draw:text-box>${slide.notes.map(p => this.paragraphXml(p)).join('')}</draw:text-box></draw:frame></presentation:notes>` : '';
        return `<draw:page draw:name="page${index + 1}" draw:master-page-name="Default"${pageStyle ? ` draw:style-name="${pageStyle}"` : ''}>${under}${items}${notes}</draw:page>`;
    }

    private frameAttrs(box: Box, rotation?: number): string {
        if (!rotation || !Number.isFinite(rotation)) return ` svg:x="${pt(box.x)}" svg:y="${pt(box.y)}" svg:width="${pt(Math.max(1, box.w))}" svg:height="${pt(Math.max(1, box.h))}"`;
        // Turned about its middle: ODF's transform turns about the frame's corner, then moves it.
        const r = -rotation * Math.PI / 180, cx = box.x + box.w / 2, cy = box.y + box.h / 2;
        const ox = cx - (box.w / 2) * Math.cos(r) - (box.h / 2) * Math.sin(r), oy = cy + (box.w / 2) * Math.sin(r) - (box.h / 2) * Math.cos(r);
        return ` svg:width="${pt(Math.max(1, box.w))}" svg:height="${pt(Math.max(1, box.h))}" draw:transform="rotate (${Math.round(r * 1e6) / 1e6}) translate (${pt(ox)} ${pt(oy)})"`;
    }

    /**
     * A shape's or frame's graphic style: its fill, outline, text anchor and insets. A page's paragraph placed on its
     * baselines (`onBaselines`) has its lines spaced independently of its font's metrics, as LibreOffice reads PPTX,
     * so they stand on the baselines the deck placed them on (see baselineSet) whatever font a reader has for it.
     */
    private graphicStyle(fill?: string, line?: { color?: string; width?: number }, anchor?: string, insets?: { top: number; right: number; bottom: number; left: number }, autoGrow = false, unwrapped = false, onBaselines = false): string {
        const f = fill && hexColor(fill) ? ` draw:fill="solid" draw:fill-color="#${hexColor(fill)}"` : ' draw:fill="none"';
        const l = line && (line.color || line.width) ? ` draw:stroke="solid" svg:stroke-color="#${hexColor(line.color ?? '#000000') ?? '000000'}" svg:stroke-width="${pt(line.width ?? 0.75)}"` : ' draw:stroke="none"';
        const a = ` draw:textarea-vertical-align="${anchor === 'middle' ? 'middle' : anchor === 'bottom' ? 'bottom' : 'top'}"`;
        const i = insets ? ` fo:padding-top="${pt(insets.top)}" fo:padding-right="${pt(insets.right)}" fo:padding-bottom="${pt(insets.bottom)}" fo:padding-left="${pt(insets.left)}"` : '';
        const key = `gr|${f}|${l}|${a}|${i}|${autoGrow}|${unwrapped}|${onBaselines}`;
        return this.styleNamed(key, 'gr', name => `<style:style style:name="${name}" style:family="graphic" style:parent-style-name="standard"><style:graphic-properties${f}${l}${a}${i} draw:auto-grow-height="${autoGrow}" draw:auto-grow-width="${unwrapped}"${unwrapped ? ' fo:wrap-option="no-wrap" draw:textarea-horizontal-align="left"' : ''}/>${onBaselines ? '<style:paragraph-properties style:font-independent-line-spacing="true"/>' : ''}</style:style>`);
    }

    /** A shape's ODF name: its preset's own, else a rectangle, reported once a name. */
    private readonly unknownShapes = new Set<string>();
    private odfShape(geometry: string | undefined): string {
        if (!geometry) return 'rectangle';
        if (Object.prototype.hasOwnProperty.call(PRESET_TO_ODF, geometry)) return PRESET_TO_ODF[geometry];
        if (!this.unknownShapes.has(geometry)) { this.unknownShapes.add(geometry); this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `a ${geometry.slice(0, 40)} shape (drawn as a rectangle)`, format: 'odp' }); }
        return 'rectangle';
    }

    private itemXml(item: DeckItem): string {
        if (item.kind === 'shape') {
            const s = item.shape;
            const kind = this.odfShape(s.geometry);
            const style = this.graphicStyle(s.fill, s.line);
            if (s.geometry === 'line') return `<draw:line draw:style-name="${style}" svg:x1="${pt(item.box.x)}" svg:y1="${pt(item.box.y)}" svg:x2="${pt(item.box.x + item.box.w)}" svg:y2="${pt(item.box.y + item.box.h)}"/>`;
            return `<draw:custom-shape draw:style-name="${style}"${this.frameAttrs(item.box, s.rotation)}><draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="${kind}"/></draw:custom-shape>`;
        }
        if (item.kind === 'picture') {
            const a = this.attachmentsByName.get(item.attachment);
            const bytes = a?.data ? decodeBase64(a.data) : undefined;
            const href = this.mediaFor(item.attachment, bytes ?? undefined, a?.mimeType);
            if (!href) {
                // (Its alternative text in its place.)
                this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: item.attachment, reason: a ? 'a picture of a kind ODP output does not hold (PNG, JPEG, GIF, BMP, TIFF, EMF, WMF and SVG it does)' : 'missing attachment' });
                return item.alt ? this.itemXml({ kind: 'text', box: item.box, paragraphs: [{ runs: [{ text: item.alt }] }] }) : '';
            }
            const size = bytes ? sniffImageSize(bytes) : null;
            let box = item.box;
            if (!item.fill && size && size.w > 0 && size.h > 0) {
                const scale = Math.min(box.w / size.w, box.h / size.h);
                box = { x: box.x + (box.w - size.w * scale) / 2, y: box.y + (box.h - size.h * scale) / 2, w: size.w * scale, h: size.h * scale };
            }
            return `<draw:frame${this.frameAttrs(box, item.rotation)}><draw:image xlink:href="${xmlText(href)}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${item.alt ? `<svg:desc>${xmlText(item.alt)}</svg:desc>` : ''}</draw:frame>`;
        }
        if (item.kind === 'table') return this.tableXml(item);
        const s = item.shape;
        const insets = item.insets ?? (item.exact ? { top: 0, right: 0, bottom: 0, left: 0 } : undefined);
        const size = item.role === 'body' ? 18 : undefined;
        const paragraphs = item.paragraphs.map(p => this.paragraphXml(p, size)).join('') || '<text:p/>';
        if (s && (s.geometry && s.geometry !== 'rect' || s.fill || s.line)) {
            // A shape with its text.
            const kind = this.odfShape(s.geometry);
            return `<draw:custom-shape draw:style-name="${this.graphicStyle(s.fill, s.line, item.anchor, insets)}"${this.frameAttrs(item.box, item.rotation ?? s.rotation)}>${paragraphs}<draw:enhanced-geometry svg:viewBox="0 0 21600 21600" draw:type="${kind}"/></draw:custom-shape>`;
        }
        return `<draw:frame draw:style-name="${this.graphicStyle(undefined, undefined, item.anchor, insets, !item.exact, item.exact, !!item.exact && item.paragraphs.some(p => p.baselineBox))}"${this.frameAttrs(item.box, item.rotation)}><draw:text-box>${paragraphs}</draw:text-box></draw:frame>`;
    }

    /** A table in a frame: its columns' widths, its cells (merges covering others), thin rules. */
    private tableXml(item: Extract<DeckItem, { kind: 'table' }>): string {
        const cols = Math.max(1, item.widths.length);
        const total = item.widths.reduce((a, b) => a + b, 0) || cols;
        // A measured table's cells keep its text close to their rules, as its source set them; a shaded cell its colour.
        const pad = item.heights ? ` fo:padding-top="0pt" fo:padding-bottom="0pt" fo:padding-left="${pt(3)}" fo:padding-right="${pt(3)}"` : '';
        // (No rules for a table whose text is all invisible: a scan's recognised table, its rules in the scan.)
        const border = unseenTable(item) ? 'none' : '0.5pt solid #BFBFBF';
        const cellStyleOf = (fill?: string) => {
            const hex = fill && hexColor(fill);
            return this.styleNamed(`cell|${border}|${hex ?? ''}|${pad}`, 'ce', name => `<style:style style:name="${name}" style:family="table-cell"><style:graphic-properties ${hex ? `draw:fill="solid" draw:fill-color="#${hex}"` : 'draw:fill="none"'}${pad}/><style:paragraph-properties fo:border="${border}"/></style:style>`);
        };
        const cellStyle = cellStyleOf();
        const hTotal = item.heights?.reduce((a, b) => a + b, 0);
        const rowStyle = (r: number) => {
            if (!item.heights || !hTotal) return '';
            const h = item.box.h * item.heights[r] / hTotal;
            return ` table:style-name="${this.styleNamed(`row|${h.toFixed(2)}`, 'ro', name => `<style:style style:name="${name}" style:family="table-row"><style:table-row-properties style:row-height="${pt(h)}"/></style:style>`)}"`;
        };
        const columns = item.widths.map(w => {
            const style = this.styleNamed(`col|${Math.round(item.box.w * w / total)}`, 'co', name => `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties style:column-width="${pt(item.box.w * w / total)}"/></style:style>`);
            return `<table:table-column table:style-name="${style}"/>`;
        }).join('');
        const rows = item.rows.length;
        const covered = new Set<string>();
        const rowsXml = item.rows.map((row, r) => {
            let col = 0;
            const out: string[] = [];
            for (const cell of row.cells) {
                while (col < cols && covered.has(`${r},${col}`)) { out.push('<table:covered-table-cell/>'); col++; }
                // (A cell past the grid is reported: none of its rows' merges left it a place.)
                if (col >= cols) { this.cellsPastGrid(); break; }
                let span = Math.min(cell.colSpan, cols - col), down = Math.min(cell.rowSpan, rows - r);
                // Its merge held to the positions no earlier cell's covers: a cell inside another's merge is none a reader keeps.
                for (let dc = 1; dc < span; dc++) if (covered.has(`${r},${col + dc}`)) { span = dc; break; }
                for (let dr = 1; dr < down; dr++) for (let dc = 0; dc < span; dc++) if (covered.has(`${r + dr},${col + dc}`)) { down = dr; break; }
                for (let dr = 0; dr < down; dr++) for (let dc = 0; dc < span; dc++) if (dr || dc) covered.add(`${r + dr},${col + dc}`);
                const paragraphs = cell.paragraphs.map(p => this.paragraphXml(cell.bold ? { ...p, runs: p.runs.map(x => ({ ...x, bold: true })) } : p, 12)).join('') || '<text:p/>';
                out.push(`<table:table-cell table:style-name="${cellStyleOf(cell.fill)}"${span > 1 ? ` table:number-columns-spanned="${span}"` : ''}${down > 1 ? ` table:number-rows-spanned="${down}"` : ''}>${paragraphs}</table:table-cell>`);
                col++;
                for (let dc = 1; dc < span; dc++) { out.push('<table:covered-table-cell/>'); col++; }
            }
            while (col < cols) { out.push(covered.has(`${r},${col}`) ? '<table:covered-table-cell/>' : `<table:table-cell table:style-name="${cellStyle}"/>`); col++; }
            return `<table:table-row${rowStyle(r)}>${out.join('')}</table:table-row>`;
        }).join('');
        return `<draw:frame${this.frameAttrs(item.box)}><table:table>${columns}${rowsXml}</table:table></draw:frame>`;
    }

    /** A paragraph: its alignment and bullet (a list of one level, as ODF writes a bullet), its spans. */
    private paragraphXml(p: DeckParagraph, defaultSize?: number): string {
        const align = p.align === 'center' ? 'center' : p.align === 'right' ? 'end' : p.align === 'justify' ? 'justify' : undefined;
        const first = !p.bullet && p.firstIndent ? p.firstIndent : 0;
        // Lines as far apart as the source set them.
        const spacing = p.lineSpacing && p.lineSpacing > 0 ? Math.min(1584, p.lineSpacing) : 0;
        const pStyle = align || p.bullet || first || spacing ? this.styleNamed(`p|${align}|${p.bullet?.level ?? ''}|${first.toFixed(2)}|${spacing.toFixed(2)}`, 'P', name => `<style:style style:name="${name}" style:family="paragraph"><style:paragraph-properties${align ? ` fo:text-align="${align}"` : ''}${spacing ? ` fo:line-height="${pt(spacing)}"` : ''}${p.bullet ? ` fo:margin-left="${pt(bulletMargin(p.bullet.level))}" fo:text-indent="${pt(-14)}"` : first ? ` fo:margin-left="${pt(Math.max(0, -first))}" fo:text-indent="${pt(first)}"` : ''}/></style:style>`) : undefined;
        const spans = p.runs.map(r => this.spanXml(r, defaultSize, p.mono)).join('');
        // A number is written as the item shows it, as a bullet's character is (an invisible item's invisible with it).
        const shown = p.bullet ? (p.bullet.numbered ? p.bullet.text ?? '' : p.bullet.char ?? '') : '';
        const lead = p.runs.find(r => r.text && !r.lineBreak);
        // (An item is invisible where all its text is: one of shown text after an invisible word keeps its marker.)
        const unseen = p.runs.some(r => r.text.trim()) && p.runs.every(r => !r.text.trim() || r.invisible);
        const marker = !p.bullet || (p.bullet.numbered && !p.bullet.text) ? ''
            : unseen ? this.spanXml({ text: `${shown}\t`, invisible: true, ...(lead?.color ? { color: lead.color } : {}) }, defaultSize)
                : `${xmlText(shown)}<text:tab/>`;
        return `<text:p${pStyle ? ` text:style-name="${pStyle}"` : ''}>${marker}${spans}</text:p>`;
    }

    private spanXml(r: DeckRun, defaultSize?: number, mono?: boolean): string {
        if (r.lineBreak) return '<text:line-break/>';
        if (!r.text) return '';
        const size = r.size ?? defaultSize;
        // Size, weight, slant and font for every script (Latin, Asian, complex): CJK, Arabic and Indic text
        // took the default's alone, and overflowed its box.
        const all = (prop: string, value: string) => ` fo:${prop}="${value}" style:${prop}-asian="${value}" style:${prop}-complex="${value}"`;
        const face = (name: string) => ` style:font-name="${xmlText(name)}" fo:font-family="${xmlText(name)}" style:font-family-asian="${xmlText(name)}" style:font-family-complex="${xmlText(name)}"`;
        const props = (size ? all('font-size', pt(size)) : '') + (r.bold ? all('font-weight', 'bold') : '') + (r.italic ? all('font-style', 'italic') : '')
            + (r.underline ? ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"' : '')
            + (r.strike ? ' style:text-line-through-style="solid"' : '') + (r.superscript ? ' style:text-position="33% 58%"' : r.subscript ? ' style:text-position="-33% 58%"' : '')
            // Invisible text in its colour made fully transparent (LibreOffice's opacity, which it reads with a colour
            // of the text's own): not seen, still there to select and search.
            + (r.invisible ? ` fo:color="#${(r.color && hexColor(r.color)) || '000000'}" loext:opacity="0%"` : r.color && hexColor(r.color) ? ` fo:color="#${hexColor(r.color)}"` : '') + (r.font || mono ? face(r.font || 'Courier New') : '');
        const text = xmlText(r.text).replace(/ {2,}/g, m => ` <text:s text:c="${m.length - 1}"/>`).replace(/\t/g, '<text:tab/>');
        const span = props ? `<text:span text:style-name="${this.styleNamed(`t|${props}`, 'T', name => `<style:style style:name="${name}" style:family="text"><style:text-properties${props}/></style:style>`)}">${text}</text:span>` : text;
        return r.link ? `<text:a xlink:type="simple" xlink:href="${xmlText(encUrl(r.link))}">${span}</text:a>` : span;
    }
}
