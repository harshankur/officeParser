import { zipSync, Zippable } from 'fflate';
import { ConversionResult, GeneratorConfig, OfficeAttachment, OfficeContentNode, OfficeParserAST, OfficeWarningType } from '../types.js';
import { escapeXml, stripInvalidXmlChars } from '../utils/sanitize.js';
import { checkAbortSignal } from '../utils/errorUtils.js';
import { DRAWINGML_PRESETS } from '../utils/shapeNames.js';
import { decodeBase64, encUrl, hexColor, ooxmlCustomPropertiesXml, resolveZipInstant, sniffImageSize, toW3CDTF } from '../utils/officeGenUtils.js';
import { BaseGenerator } from './BaseGenerator.js';
import { buildDeck, bulletMargin, Box, DeckItem, DeckParagraph, DeckRun, DeckSlide } from './slides/deck.js';

/** Escapes text content for an XML part after stripping XML-illegal characters. */
/** A size in points as a text size PowerPoint takes (hundredths of a point, 1 to 4,000 points). */
const pptSize = (pt: number) => Math.max(100, Math.min(400000, Math.round(pt * 100)));

function xmlText(s: string | undefined): string {
    return escapeXml(stripInvalidXmlChars(s ?? ''));
}

/** Points as English Metric Units (12,700 a point), whole. */
/** Points as EMUs: whole, finite, within a million points either way (a page's box no slide holds is held to it). */
const emu = (pt: number) => Number.isFinite(pt) ? Math.round(Math.max(-1e6, Math.min(1e6, pt)) * 12700) : 0;

const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CLR_MAP = '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>';
const GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/** The picture formats a slide holds, by their media types. */
const PICTURE_TYPES = new Map<string, string>([['image/png', 'png'], ['image/jpeg', 'jpeg'], ['image/jpg', 'jpeg'], ['image/gif', 'gif'], ['image/bmp', 'bmp'], ['image/tiff', 'tiff'],
    ['image/x-emf', 'emf'], ['image/emf', 'emf'], ['image/x-wmf', 'wmf'], ['image/wmf', 'wmf']]);
/** A picture part's content type by its extension (EMF's and WMF's are not `image/<ext>`). */
const PICTURE_CONTENT_TYPES = new Map<string, string>([['emf', 'image/x-emf'], ['wmf', 'image/x-wmf']]);

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
        // onNode is asked once for each node, however often the deck consults it.
        this.onNodeMemo = new WeakMap();
        try { return await this.generateDeck(); } finally { this.onNodeMemo = null; }
    }

    private async generateDeck(): Promise<ConversionResult<'pptx'>> {
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
        }, { backgrounds: this.config.pptxConfig.backgrounds ?? undefined });
        for (const feature of deck.unrepresentable) this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature, format: 'pptx' });
        const files: Zippable = {};
        const enc = (s: string) => new TextEncoder().encode(s);
        const slideXml: string[] = [], slideRels: string[] = [], notesXml: (string | undefined)[] = [], notesRels: string[][] = [];
        deck.slides.forEach((slide, i) => {
            const rels: string[] = [`<Relationship Id="rId1" Type="${REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>`];
            const rel = (type: string, target: string, external = false) => { const id = `rId${rels.length + 1}`; rels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${xmlText(target)}"${external ? ' TargetMode="External"' : ''}/>`); return id; };
            slideXml.push(this.slideXml(slide, rel));
            if (slide.notes?.length) { rel('notesSlide', `../notesSlides/notesSlide${i + 1}.xml`); const n = this.notesXml(slide.notes); notesXml[i] = n.xml; notesRels[i] = n.rels; }
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
                files[`ppt/notesSlides/_rels/notesSlide${i + 1}.xml.rels`] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/notesMaster" Target="../notesMasters/notesMaster1.xml"/><Relationship Id="rId2" Type="${REL}/slide" Target="../slides/slide${i + 1}.xml"/>${(notesRels[i] ?? []).join('')}</Relationships>`);
            });
        }
        for (const m of this.media) files[m.path] = m.bytes;
        const { iso, mtime } = resolveZipInstant(this.effectiveMetadata.modified);
        files['docProps/core.xml'] = enc(this.coreProps(iso));
        const custom = ooxmlCustomPropertiesXml(this.effectiveMetadata.customProperties);
        if (custom) files['docProps/custom.xml'] = enc(custom);
        files['docProps/app.xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>officeParser</Application><Slides>${n}</Slides></Properties>`);
        files['_rels/.rels'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
            + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="ppt/presentation.xml"/>`
            + `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>`
            + `<Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/>`
            + (custom ? `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/>` : '') + '</Relationships>');
        const exts = [...new Set(this.media.map(m => m.ext))];
        const pml = 'application/vnd.openxmlformats-officedocument.presentationml';
        files['[Content_Types].xml'] = enc(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
            + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
            + exts.map(e => `<Default Extension="${e}" ContentType="${PICTURE_CONTENT_TYPES.get(e) ?? `image/${e}`}"/>`).join('')
            + `<Override PartName="/ppt/presentation.xml" ContentType="${pml}.presentation.main+xml"/>`
            + `<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="${pml}.slideMaster+xml"/>`
            + `<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="${pml}.slideLayout+xml"/>`
            + deck.slides.map((_, i) => `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="${pml}.slide+xml"/>`).join('')
            + notesXml.map((x, i) => x ? `<Override PartName="/ppt/notesSlides/notesSlide${i + 1}.xml" ContentType="${pml}.notesSlide+xml"/>` : '').join('')
            + (hasNotes ? `<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="${pml}.notesMaster+xml"/><Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` : '')
            + `<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`
            + `<Override PartName="/ppt/presProps.xml" ContentType="${pml}.presProps+xml"/><Override PartName="/ppt/viewProps.xml" ContentType="${pml}.viewProps+xml"/>`
            + `<Override PartName="/ppt/tableStyles.xml" ContentType="${pml}.tableStyles+xml"/>`
            + `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`
            + (custom ? '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>' : '') + '</Types>');
        return { value: zipSync(files, { mtime }), messages: this.messages };
    }

    /** A picture's part in the package, added once for each attachment (or background), or undefined where it cannot be held. */
    private mediaFor(key: string, bytes: Uint8Array | undefined, mime: string | undefined): string | undefined {
        let at = this.mediaByAttachment.get(key);
        if (at === undefined) {
            if (!bytes || !bytes.length) return undefined;
            // (Its bytes say what it is first: EMF by its signature at 40, a placeable WMF by its key.)
            const sniffed = bytes[0] === 0x89 && bytes[1] === 0x50 ? 'png' : bytes[0] === 0xff && bytes[1] === 0xd8 ? 'jpeg' : bytes[0] === 0x47 && bytes[1] === 0x49 ? 'gif' : bytes[0] === 0x42 && bytes[1] === 0x4d ? 'bmp'
                : bytes[0] === 1 && bytes[1] === 0 && bytes[2] === 0 && bytes[3] === 0 && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46 ? 'emf'
                : bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a ? 'wmf' : undefined;
            // (A type named by the attachment, looked up as a type alone: one that is a name every object has, `constructor`, is none.)
            const ext = sniffed ?? (mime ? PICTURE_TYPES.get(mime.toLowerCase()) : undefined);
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
        const rot = rotation && Number.isFinite(rotation) ? ` rot="${Math.round((((rotation % 360) + 360) % 360) * 60000)}"` : '';
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
            return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${this.xfrm(item.box, s.rotation, s.flipH, s.flipV)}<a:prstGeom prst="${this.preset(s.geometry)}"><a:avLst/></a:prstGeom>${this.looks(s.fill, s.line)}</p:spPr></p:sp>`;
        }
        if (item.kind === 'picture') {
            const a = this.attachmentsByName.get(item.attachment);
            const bytes = a?.data ? decodeBase64(a.data) : undefined;
            const target = this.mediaFor(item.attachment, bytes ?? undefined, a?.mimeType);
            if (!target) {
                // (An SVG picture PowerPoint shows only beside a PNG copy, which is not made here; another kind PPTX
                // holds none of. Its alternative text in its place.)
                const svg = /svg/i.test(a?.mimeType ?? '') || /\.svg$/i.test(item.attachment);
                this.warn(OfficeWarningType.IMAGE_PROCESSING_FAILED, { name: item.attachment, reason: !a ? 'missing attachment' : svg ? 'an SVG picture (PPTX output holds one only with a PNG copy of it)' : 'a picture of a kind PPTX output does not hold (PNG, JPEG, GIF, BMP, TIFF, EMF and WMF it does)' });
                return item.alt ? this.itemXml({ kind: 'text', box: item.box, paragraphs: [{ runs: [{ text: item.alt }] }] }, next, rel) : '';
            }
            // Fitted into its box, keeping its shape, centred (unless it fills it, as its page drew it).
            const size = bytes ? sniffImageSize(bytes) : null;
            let box = item.box;
            if (!item.fill && size && size.w > 0 && size.h > 0) {
                const scale = Math.min(box.w / size.w, box.h / size.h);
                const w = size.w * scale, h = size.h * scale;
                box = { x: box.x + (box.w - w) / 2, y: box.y + (box.h - h) / 2, w, h };
            }
            const id = next();
            return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"${item.alt ? ` descr="${xmlText(item.alt)}"` : ''}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rel('image', target)}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${this.xfrm(box, item.rotation)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
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
            + `<p:spPr>${this.xfrm(item.box, item.rotation ?? shape?.rotation, shape?.flipH, shape?.flipV)}<a:prstGeom prst="${this.preset(shape?.geometry)}"><a:avLst/></a:prstGeom>${this.looks(shape?.fill, shape?.line)}</p:spPr>`
            + `<p:txBody>${body}<a:lstStyle/>${item.paragraphs.map(p => this.paragraphXml(p, rel, size)).join('') || '<a:p/>'}</p:txBody></p:sp>`;
    }

    /** Reports, once, a table's cells that found no place in its grid. */
    private pastGridReported = false;
    private cellsPastGrid(): void {
        if (this.pastGridReported) return;
        this.pastGridReported = true;
        this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: 'table cells past their table\'s columns', format: 'pptx' });
    }

    /** A shape's preset: its own where DrawingML has it, else a rectangle, reported once a name. */
    private readonly unknownShapes = new Set<string>();
    private preset(geometry: string | undefined): string {
        if (!geometry || DRAWINGML_PRESETS.has(geometry)) return geometry || 'rect';
        if (!this.unknownShapes.has(geometry)) { this.unknownShapes.add(geometry); this.warn(OfficeWarningType.CONTENT_NOT_REPRESENTABLE, { feature: `a ${geometry.slice(0, 40)} shape (drawn as a rectangle)`, format: 'pptx' }); }
        return 'rect';
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
                // (A cell past the grid is reported: none of its rows' merges left it a place.)
                if (col >= cols) { this.cellsPastGrid(); break; }
                let span = Math.min(cell.colSpan, cols - col), down = Math.min(cell.rowSpan, rows - r);
                // Its merge held to the positions no earlier cell's holds: two that overlap, PowerPoint repairs.
                for (let dc = 1; dc < span; dc++) if (grid[r][col + dc]) { span = dc; break; }
                for (let dr = 1; dr < down; dr++) if (grid[r + dr].slice(col, col + span).some(Boolean)) { down = dr; break; }
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
        const indent = b ? ` marL="${emu(bulletMargin(b.level))}" indent="${emu(-14)}"${b.level ? ` lvl="${Math.min(8, b.level)}"` : ''}`
            : first ? ` marL="${emu(Math.max(0, -first))}" indent="${emu(first)}"` : '';
        // A number is written as the item shows it (each item a paragraph or a box of its own, an automatic
        // number restarted at 1 in each).
        // (A marker of more than one character, `.1` as a right-to-left list sets it, is text too: a bullet is one
        // character, and the rest was lost.)
        const long = !!b && !b.numbered && !!b.char && [...b.char].length > 1;
        const bullet = !b || b.numbered || long ? '<a:buNone/>' : b.char ? `<a:buFont typeface="Arial"/><a:buChar char="${xmlText(b.char)}"/>` : '<a:buNone/>';
        const lead = p.runs.find(r => r.text && !r.lineBreak);
        const shown = b?.numbered ? b.text : long ? b!.char : undefined;
        const number: DeckRun[] = shown ? [{ text: `${shown}\t`, ...(lead?.size ? { size: lead.size } : {}), ...(lead?.font ? { font: lead.font } : {}), ...(lead?.color ? { color: lead.color } : {}) }] : [];
        const runs = [...number, ...p.runs].map(r => this.runXml(r, rel, defaultSize, p.mono)).join('');
        const size = p.runs.find(r => r.size)?.size ?? defaultSize;
        // Lines as far apart as the source set them (spcPts holds at most 1,584 points).
        const spacing = p.lineSpacing && p.lineSpacing > 0 ? `<a:lnSpc><a:spcPts val="${Math.round(Math.min(1584, p.lineSpacing) * 100)}"/></a:lnSpc>` : '';
        return `<a:p><a:pPr${algn}${indent}>${spacing}${bullet}</a:pPr>${runs}<a:endParaRPr lang="en-US"${size ? ` sz="${pptSize(size)}"` : ''} dirty="0"/></a:p>`;
    }

    private runXml(r: DeckRun, rel: (type: string, target: string, external?: boolean) => string, defaultSize?: number, mono?: boolean): string {
        if (r.lineBreak) return '<a:br><a:rPr lang="en-US" dirty="0"/></a:br>';
        if (!r.text) return '';
        const size = r.size ?? defaultSize;
        const attrs = `lang="en-US"${size ? ` sz="${pptSize(size)}"` : ''}${r.bold ? ' b="1"' : ''}${r.italic ? ' i="1"' : ''}${r.underline ? ' u="sng"' : ''}${r.strike ? ' strike="sngStrike"' : ''}${r.superscript ? ' baseline="30000"' : r.subscript ? ' baseline="-25000"' : ''} dirty="0"`;
        const color = r.color && hexColor(r.color) ? `<a:solidFill><a:srgbClr val="${hexColor(r.color)}"/></a:solidFill>` : '';
        const font = r.font || mono ? `<a:latin typeface="${xmlText(r.font || 'Courier New')}"/>` : '';
        const link = r.link ? `<a:hlinkClick r:id="${rel('hyperlink', encUrl(r.link), true)}"/>` : '';
        return `<a:r><a:rPr ${attrs}>${color}${font}${link}</a:rPr><a:t>${xmlText(r.text)}</a:t></a:r>`;
    }

    /** A notes page: the slide's notes in its body, and the relationships of their links (after its master's and slide's). */
    private notesXml(notes: DeckParagraph[]): { xml: string; rels: string[] } {
        const rels: string[] = [];
        const rel = (type: string, target: string, external = false) => {
            const id = `rId${rels.length + 3}`;
            rels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${xmlText(target)}"${external ? ' TargetMode="External"' : ''}/>`);
            return id;
        };
        const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:notes ${NS}><p:cSld><p:spTree>${GROUP}`
            + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>`
            + `<p:txBody><a:bodyPr/><a:lstStyle/>${notes.map(p => this.paragraphXml(p, rel)).join('') || '<a:p/>'}</p:txBody></p:sp></p:spTree></p:cSld></p:notes>`;
        return { xml, rels };
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
