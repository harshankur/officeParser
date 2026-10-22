import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { decodeBase64, hexColor, resolveZipInstant, sniffImageSize, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildDeck, Box, DeckItem, DeckParagraph, DeckRun, DeckSlide } from './slides/deck.js';

/** Escapes text content for an XML part after stripping XML-illegal characters. */
function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** Points as English Metric Units (12,700 a point), whole. */
const emu = (pt: number) => Math.round(pt * 12700);

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
const GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/** The picture formats a slide holds, by their media types. */
const PICTURE_TYPES: Record<string, string> = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/jpg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/tiff': 'tiff' };

/** A theme with Office's colours and fonts, as a presentation needs one. */
function themeXml(name: string): string {
    const colors = [['dk1', '<a:sysClr val="windowText" lastClr="000000"/>'], ['lt1', '<a:sysClr val="window" lastClr="FFFFFF"/>'], ['dk2', '<a:srgbClr val="44546A"/>'], ['lt2', '<a:srgbClr val="E7E6E6"/>'],
        ['accent1', '<a:srgbClr val="4472C4"/>'], ['accent2', '<a:srgbClr val="ED7D31"/>'], ['accent3', '<a:srgbClr val="A5A5A5"/>'], ['accent4', '<a:srgbClr val="FFC000"/>'],
        ['accent5', '<a:srgbClr val="5B9BD5"/>'], ['accent6', '<a:srgbClr val="70AD47"/>'], ['hlink', '<a:srgbClr val="0563C1"/>'], ['folHlink', '<a:srgbClr val="954F72"/>']]
        .map(([n, v]) => `<a:${n}>${v}</a:${n}>`).join('');
    const fill = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    const line = (w: number) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements>`
        + `<a:clrScheme name="Office">${colors}</a:clrScheme>`
        + `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>`
        + `<a:fmtScheme name="Office"><a:fillStyleLst>${fill}${fill}${fill}</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>`
        + `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>`
        + `<a:bgFillStyleLst>${fill}${fill}${fill}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

/**
 * Generates PPTX presentations from an AST (see buildDeck): a deck's slides as they are, a PDF's pages a
 * slide each with each paragraph a text box in its place (and a background picture a page where
 * `pptxConfig.backgrounds` gives one), a flowing document's headings a slide each. What a slide cannot
 * hold (a chart, an SVG picture) is reported.
 */
export class PptxGenerator extends BaseGenerator<'pptx'> {
    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'pptx'>) {
        super('pptx', ast, config);
    }

    private readonly media: { path: string; bytes: Uint8Array; ext: string }[] = [];
    private readonly mediaByAttachment = new Map<string, number>();
    private readonly attachmentsByName = new Map<string, OfficeAttachment>();

    async generate(): Promise<ConversionResult<'pptx'>> {
        for (const a of this.ast.attachments ?? []) if (a?.name) this.attachmentsByName.set(a.name, a);
        const deck = await buildDeck(this.ast.content, this.ast.metadata ?? {}, {
            // Every node a slide takes is asked for here, so the signal is heard node by node.
            keep: async node => { checkAbortSignal(this.config.abortSignal); const v = await this.handleOnNode(node); return v === false ? false : typeof v === 'string' ? v : true; },
            chartTable: node => {
                if (this.config.includeCharts === false) return undefined;
                const data = this.attachmentsByName.get((node.metadata as { attachmentName?: string } | undefined)?.attachmentName ?? '')?.chartData;
                return data ? this.chartTable(data, text => ({ type: 'cell', text, children: [{ type: 'text', text }] } as OfficeContentNode)) : undefined;
            },
        }, { backgrounds: this.config.pptxConfig.backgrounds ?? undefined });
        for (const feature of deck.unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'pptx' });
        const files: Zippable = {};
        const enc = (s: string) => new TextEncoder().encode(s);
        const slideXml: string[] = [], slideRels: string[] = [], notesXml: (string | undefined)[] = [];
        deck.slides.forEach((slide, i) => {
            const rels: string[] = [`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`];
            const rel = (type: string, target: string, external = false) => { const id = `rId${rels.length + 1}`; rels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${xmlText(target)}"${external ? ' TargetMode="External"' : ''}/>`); return id; };
            slideXml.push(this.slideXml(slide, rel));
            if (slide.notes?.length) { rel('notesSlide', `../notesSlides/notesSlide${i + 1}.xml`); notesXml[i] = this.notesXml(slide.notes); }
            slideRels.push(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`);
        });
        const hasNotes = notesXml.some(Boolean);
        const n = deck.slides.length;
        // The presentation's relationships: its master, slides, notes master, theme, properties.
        const presRels = [`<Relationship Id="rId1" Type="${REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>`,
            ...deck.slides.map((_, i) => `<Relationship Id="rId${i + 2}" Type="${REL}/slide" Target="slides/slide${i + 1}.xml"/>`),
            `<Relationship Id="rId${n + 2}" Type="${REL}/theme" Target="theme/theme1.xml"/>`,
            `<Relationship Id="rId${n + 3}" Type="${REL}/presProps" Target="presProps.xml"/>`,
            `<Relationship Id="rId${n + 4}" Type="${REL}/viewProps" Target="viewProps.xml"/>`,
            `<Relationship Id="rId${n + 5}" Type="${REL}/tableStyles" Target="tableStyles.xml"/>`,
            ...(hasNotes ? [`<Relationship Id="rId${n + 6}" Type="${REL}/notesMaster" Target="notesMasters/notesMaster1.xml"/>`] : [])];
        files['ppt/presentation.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentation ${NS} saveSubsetFonts="1">`
            + `<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>`
            + (hasNotes ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${n + 6}"/></p:notesMasterIdLst>` : '')
            + `<p:sldIdLst>${deck.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst>`
            + `<p:sldSz cx="${Math.min(51206400, Math.max(914400, emu(deck.width)))}" cy="${Math.min(51206400, Math.max(914400, emu(deck.height)))}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`);
        files['ppt/_rels/presentation.xml.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${presRels.join('')}</Relationships>`);
        files['ppt/slideMasters/slideMaster1.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GROUP}</p:spTree></p:cSld>${CLR_MAP}<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`);
        files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL}/theme" Target="../theme/theme1.xml"/></Relationships>`);
        files['ppt/slideLayouts/slideLayout1.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${GROUP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
        files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`);
        files['ppt/theme/theme1.xml'] = enc(themeXml('Office Theme'));
        files['ppt/presProps.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:presentationPr ${NS}/>`);
        files['ppt/viewProps.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr ${NS}><p:normalViewPr/><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`);
        files['ppt/tableStyles.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`);
        slideXml.forEach((xml, i) => {
            files[`ppt/slides/slide${i + 1}.xml`] = enc(xml);
            files[`ppt/slides/_rels/slide${i + 1}.xml.rels`] = enc(slideRels[i]);
        });
        if (hasNotes) {
            files['ppt/notesMasters/notesMaster1.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notesMaster ${NS}><p:cSld><p:spTree>${GROUP}</p:spTree></p:cSld>${CLR_MAP}</p:notesMaster>`);
            files['ppt/notesMasters/_rels/notesMaster1.xml.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/theme" Target="../theme/theme2.xml"/></Relationships>`);
            files['ppt/theme/theme2.xml'] = enc(themeXml('Notes Theme'));
            notesXml.forEach((xml, i) => {
                if (!xml) return;
                files[`ppt/notesSlides/notesSlide${i + 1}.xml`] = enc(xml);
                files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="../slides/slide${i + 1}.xml"/></Relationships>`);
            });
        }
        for (const m of this.media) files[m.path] = m.bytes;
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);
        files['docProps/core.xml'] = enc(this.coreProps(iso));
        files['docProps/app.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>officeParser</Application><Slides>${n}</Slides></Properties>`);
        files['_rels/.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
            + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`
            + `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
            + `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`);
        const exts = [...new Set(this.media.map(m => m.ext))];
        const pml = 'application/vnd.openxmlformats-officedocument.presentationml';
        files['[Content_Types].xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
            + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
            + exts.map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('')
            + `<Override PartName="/ppt/presentation.xml" ContentType="${pml}.presentation.main+xml"/>`
            + `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${pml}.slideMaster+xml"/>`
            + `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${pml}.slideLayout+xml"/>`
            + deck.slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${pml}.slide+xml"/>`).join('')
            + notesXml.map((x, i) => x ? `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="${pml}.notesSlide+xml"/>` : '').join('')
            + (hasNotes ? `<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${pml}.notesMaster+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` : '')
            + `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
            + `<Override PartName="/ppt/presProps.xml" ContentType="${pml}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${pml}.viewProps+xml"/>`
            + `<Override PartName="/ppt/tableStyles.xml" ContentType="${pml}.tableStyles+xml"/>`
            + `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
        return { value: zipSync(files, { mtime }), messages: this.messages };
    }

    /** A picture's part in the package, added once for each attachment (or background), or undefined where it cannot be held. */
    private mediaFor(key: string, bytes: Uint8Array | undefined, mime: string | undefined): string | undefined {
        let at = this.mediaByAttachment.get(key);
        if (at === undefined) {
            if (!bytes || !bytes.length) return undefined;
            const sniffed = bytes[0] === 0x89 && bytes[1] === 0x50 ? 'png' : bytes[0] === 0xff && bytes[1] === 0xd8 ? 'jpeg' : bytes[0] === 0x47 && bytes[1] === 0x49 ? 'gif' : bytes[0] === 0x42 && bytes[1] === 0x4d ? 'bmp' : undefined;
            const ext = sniffed ?? (mime ? PICTURE_TYPES[mime.toLowerCase()] : undefined);
            if (!ext) return undefined;
            at = this.media.length;
            this.media.push({ path: `ppt/media/image${at + 1}.${ext}`, bytes, ext });
            this.mediaByAttachment.set(key, at);
        }
        return `../media/image${at + 1}.${this.media[at].ext}`;
    }

    /** A slide's part: its background (a colour, or a picture), then each item in its place. */
    private slideXml(slide: DeckSlide, rel: (type: string, target: string, external?: boolean) => string): string {
        let id = 1;
        const next = () => ++id;
        let bg = '';
        if (slide.picture) {
            const target = this.mediaFor(`background:${this.media.length}`, slide.picture, undefined);
            if (target) bg = `<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="${rel('image', target)}"/><a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
            else this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { reason: 'a background picture that is not PNG, JPEG, GIF or BMP' });
        } else if (slide.background && hexColor(slide.background)) {
            bg = `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${hexColor(slide.background)}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
        }
        const items = slide.items.map(item => this.itemXml(item, next, rel)).join('');
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld ${NS}><p:cSld>${bg}<p:spTree>${GROUP}${items}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
    }

    private xfrm(box: Box, rotation?: number, flipH?: boolean, flipV?: boolean, tag = 'a:xfrm'): string {
        const rot = rotation ? ` rot="${Math.round((((rotation % 360) + 360) % 360) * 60000)}"` : '';
        return `<${tag}${rot}${flipH ? ' flipH="1"' : ''}${flipV ? ' flipV="1"' : ''}><a:off x="${emu(box.x)}" y="${emu(box.y)}"/><a:ext cx="${emu(Math.max(box.w, 1))}" cy="${emu(Math.max(box.h, 1))}"/></${tag}>`;
    }

    /** A shape's fill and outline as DrawingML. */
    private looks(fill?: string, line?: { color?: string; width?: number }): string {
        const f = fill && hexColor(fill) ? `<a:solidFill><a:srgbClr val="${hexColor(fill)}"/></a:solidFill>` : '<a:noFill/>';
        const l = line && (line.color || line.width) ? `<a:ln w="${emu(line.width ?? 0.75)}"><a:solidFill><a:srgbClr val="${hexColor(line.color ?? '#000000') ?? '000000'}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>';
        return f + l;
    }

    private itemXml(item: DeckItem, next: () => number, rel: (type: string, target: string, external?: boolean) => string): string {
        if (item.kind === 'shape') {
            const s = item.shape;
            const id = next();
            return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${this.xfrm(item.box, s.rotation, s.flipH, s.flipV)}<a:prstGeom prst="${xmlText(s.geometry || 'rect')}"><a:avLst/></a:prstGeom>${this.looks(s.fill, s.line)}</p:spPr></p:sp>`;
        }
        if (item.kind === 'picture') {
            const a = this.attachmentsByName.get(item.attachment);
            const bytes = a?.data ? decodeBase64(a.data) : undefined;
            const target = this.mediaFor(item.attachment, bytes ?? undefined, a?.mimeType);
            if (!target) { this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: item.attachment, reason: a ? 'a picture a slide cannot hold (PNG, JPEG, GIF and BMP it can)' : 'missing attachment' }); return ''; }
            // Fitted into its box, keeping its shape, centred.
            const size = bytes ? sniffImageSize(bytes) : null;
            let box = item.box;
            if (size && size.w > 0 && size.h > 0) {
                const scale = Math.min(box.w / size.w, box.h / size.h);
                const w = size.w * scale, h = size.h * scale;
                box = { x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h };
            }
            const id = next();
            return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rel('image', target)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${this.xfrm(box, item.rotation)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
        }
        if (item.kind === 'table') return this.tableXml(item, next, rel);
        // A text box: exactly placed (a page's paragraph: no insets, no fitting) or a body or title, or a shape's text.
        const shape = item.shape;
        const ins = item.insets ?? (item.exact ? { top: 0, right: 0, bottom: 0, left: 0 } : undefined);
        // (A page's paragraph keeps the lines it had, unwrapped: wrapped again in other fonts, it ran into the next.)
        const body = `<a:bodyPr wrap="${item.exact ? 'none' : 'square'}"${ins ? ` lIns="${emu(ins.left)}" tIns="${emu(ins.top)}" rIns="${emu(ins.right)}" bIns="${emu(ins.bottom)}"` : ''} anchor="${item.anchor === 'middle' ? 'ctr' : item.anchor === 'bottom' ? 'b' : 't'}" rtlCol="0">${item.exact ? '<a:noAutofit/>' : '<a:normAutofit/>'}</a:bodyPr>`;
        const id = next();
        const size = item.role === 'body' ? 18 : undefined;
        return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${item.role === 'title' ? 'Title' : 'Text'} ${id}"/><p:cNvSpPr${shape ? '' : ' txBox="1"'}/><p:nvPr/></p:nvSpPr>`
            + `<p:spPr>${this.xfrm(item.box, item.rotation ?? shape?.rotation, shape?.flipH, shape?.flipV)}<a:prstGeom prst="${xmlText(shape?.geometry || 'rect')}"><a:avLst/></a:prstGeom>${this.looks(shape?.fill, shape?.line)}</p:spPr>`
            + `<p:txBody>${body}<a:lstStyle/>${item.paragraphs.map(p => this.paragraphXml(p, rel, size)).join('') || '<a:p/>'}</p:txBody></p:sp>`;
    }

    /** A table: a column grid, rows of cells, merged cells' covered positions marked as merged, thin rules. */
    private tableXml(item: Extract<DeckItem, { kind: 'table' }>, next: () => number, rel: (type: string, target: string, external?: boolean) => string): string {
        const cols = Math.max(1, item.widths.length);
        const total = item.widths.reduce((a, b) => a + b, 0) || cols;
        const widths = item.widths.map(w => emu(item.box.w * w / total));
        // Rows as the source measured them (in proportion, filling the box), else of one height.
        const hTotal = item.heights?.reduce((a, b) => a + b, 0);
        const rowH = (r: number) => emu(item.heights && hTotal ? item.box.h * item.heights[r] / hTotal : Math.max(12, item.box.h / Math.max(1, item.rows.length)));
        // A measured table's cells keep its text close to their rules, as its source set them.
        const margins = item.heights ? ` marL="${emu(3)}" marR="${emu(3)}" marT="0" marB="0"` : '';
        // The grid: each cell at the first free position of its row, the positions it spans over covered
        // (merged across from it, merged down from it, or both).
        const rows = item.rows.length;
        const grid: ({ cell: typeof item.rows[number]['cells'][number]; span: number; down: number } | { h: boolean; v: boolean } | undefined)[][] = Array.from({ length: rows }, () => new Array(cols).fill(undefined));
        item.rows.forEach((row, r) => {
            let col = 0;
            for (const cell of row.cells) {
                while (col < cols && grid[r][col]) col++;
                if (col >= cols) break;
                const span = Math.min(cell.colSpan, cols - col), down = Math.min(cell.rowSpan, rows - r);
                for (let dr = 0; dr < down; dr++) for (let dc = 0; dc < span; dc++) if ((dr || dc) && !grid[r + dr][col + dc]) grid[r + dr][col + dc] = { h: dc > 0, v: dr > 0 };
                grid[r][col] = { cell, span, down };
                col += span;
            }
        });
        const rule = '<a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill>';
        const edges = ['lnL', 'lnR', 'lnT', 'lnB'].map(e => `<a:${e} w="9525">${rule}</a:${e}>`).join('');
        const empty = (attrs: string) => `<a:tc${attrs}><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr${margins}>${edges}</a:tcPr></a:tc>`;
        const rowsXml = grid.map((row, r) => `<a:tr h="${rowH(r)}">${row.map(at => {
            if (!at) return empty('');
            if (!('cell' in at)) return empty(`${at.h ? ' hMerge="1"' : ''}${at.v ? ' vMerge="1"' : ''}`);
            const { cell, span, down } = at;
            const paragraphs = cell.paragraphs.map(p => this.paragraphXml(cell.bold ? { ...p, runs: p.runs.map(x => ({ ...x, bold: true })) } : p, rel, 12)).join('') || '<a:p/>';
            const fill = cell.fill && hexColor(cell.fill) ? `<a:solidFill><a:srgbClr val="${hexColor(cell.fill)}"/></a:solidFill>` : '';
            return `<a:tc${span > 1 ? ` gridSpan="${span}"` : ''}${down > 1 ? ` rowSpan="${down}"` : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</a:txBody><a:tcPr${margins}>${edges}${fill}</a:tcPr></a:tc>`;
        }).join('')}</a:tr>`).join('');
        const id = next();
        return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
            + this.xfrm(item.box, undefined, undefined, undefined, 'p:xfrm')
            + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid>${widths.map(w => `<a:gridCol w="${w}"/>`).join('')}</a:tblGrid>${rowsXml}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
    }

    /** A paragraph: its alignment, bullet (a character, a number, or none) and runs. */
    private paragraphXml(p: DeckParagraph, rel: (type: string, target: string, external?: boolean) => string, defaultSize?: number): string {
        const algn = p.align === 'center' ? ' algn="ctr"' : p.align === 'right' ? ' algn="r"' : p.align === 'justify' ? ' algn="just"' : '';
        const b = p.bullet;
        const first = p.firstIndent;
        const indent = b ? ` marL="${emu(18 * (b.level + 1))}" indent="${emu(-14)}"${b.level ? ` lvl="${Math.min(8, b.level)}"` : ''}`
            : first ? ` marL="${emu(Math.max(0, -first))}" indent="${emu(first)}"` : '';
        const bullet = !b ? '<a:buNone/>' : b.numbered ? '<a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/>' : b.char ? `<a:buFont typeface="Arial"/><a:buChar char="${xmlText([...b.char][0])}"/>` : '<a:buNone/>';
        const runs = p.runs.map(r => this.runXml(r, rel, defaultSize, p.mono)).join('');
        const size = p.runs.find(r => r.size)?.size ?? defaultSize;
        return `<a:p><a:pPr${algn}${indent}>${bullet}</a:pPr>${runs}<a:endParaRPr lang="en-US"${size ? ` sz="${Math.round(size * 100)}"` : ''} dirty="0"/></a:p>`;
    }

    private runXml(r: DeckRun, rel: (type: string, target: string, external?: boolean) => string, defaultSize?: number, mono?: boolean): string {
        if (r.lineBreak) return '<a:br><a:rPr lang="en-US" dirty="0"/></a:br>';
        if (!r.text) return '';
        const size = r.size ?? defaultSize;
        const attrs = `lang="en-US"${size ? ` sz="${Math.max(100, Math.min(400000, Math.round(size * 100)))}"` : ''}${r.bold ? ' b="1"' : ''}${r.italic ? ' i="1"' : ''}${r.underline ? ' u="sng"' : ''}${r.strike ? ' strike="sngStrike"' : ''}${r.superscript ? ' baseline="30000"' : r.subscript ? ' baseline="-25000"' : ''} dirty="0"`;
        const color = r.color && hexColor(r.color) ? `<a:solidFill><a:srgbClr val="${hexColor(r.color)}"/></a:solidFill>` : '';
        const font = r.font || mono ? `<a:latin typeface="${xmlText(r.font || 'Courier New')}"/>` : '';
        const link = r.link ? `<a:hlinkClick r:id="${rel('hyperlink', r.link, true)}"/>` : '';
        return `<a:r><a:rPr ${attrs}>${color}${font}${link}</a:rPr><a:t>${xmlText(r.text)}</a:t></a:r>`;
    }

    /** A notes page: the slide's notes in its body. */
    private notesXml(notes: DeckParagraph[]): string {
        const none = () => '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notes ${NS}><p:cSld><p:spTree>${GROUP}`
            + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>`
            + `<p:txBody><a:bodyPr/><a:lstStyle/>${notes.map(p => this.paragraphXml({ ...p, runs: p.runs.map(r => ({ ...r, link: undefined })) }, none)).join('') || '<a:p/>'}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
    }

    private coreProps(iso: string): string {
        const m = this.effectiveMetadata;
        const created = toW3CDTF(m.created) ?? iso;
        const el = (tag: string, val: string | undefined) => val ? `<${tag}>${xmlText(val)}</${tag}>` : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`
            + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" `
            + `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
            + el('dc:title', m.title) + el('dc:creator', m.author) + el('dc:subject', m.subject)
            + el('dc:description', m.description) + el('cp:keywords', m.keywords) + el('cp:lastModifiedBy', m.lastModifiedBy)
            + `<dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>`
            + `<dcterms:modified xsi:type="dcterms:W3CDTF">${iso}</dcterms:modified></cp:coreProperties>`;
    }
}
