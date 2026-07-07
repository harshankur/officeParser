import { EmbedMetadata, FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { parseOfficeDate } from '../utils/dateUtils.js';
import { checkAbortSignal, getOfficeError } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { getAttribute, getElementsByTagName, getFirstElementByTagName, parseXmlString } from '../utils/xmlUtils.js';
import { extractFiles } from '../utils/zipUtils.js';
import { parseHtml } from './HtmlParser.js';

/**
 * Resolves a manifest-relative href against the OPF file's directory, collapsing
 * `./` and `../` segments the way a normal filesystem path resolver would.
 */
const resolveOpfPath = (opfDir: string, href: string): string => {
    const parts = (opfDir + href).split('/');
    const resolved: string[] = [];
    for (const part of parts) {
        if (part === '.' || part === '') continue;
        if (part === '..') resolved.pop();
        else resolved.push(part);
    }
    return resolved.join('/');
};

/**
 * Parses an EPUB file (a ZIP archive of XHTML content plus an OPF manifest) into the
 * unified OfficeParserAST. Each spine item is parsed via the existing `HtmlParser` and
 * the resulting content/attachments are concatenated in reading order - EPUB is
 * essentially a sequence of XHTML documents, so there's no need for a bespoke content model.
 */
export const parseEpub = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);

    const files = await extractFiles(
        buffer,
        (path) => /META-INF\/container\.xml$/i.test(path)
            || /\.opf$/i.test(path)
            || /\.(xhtml|html|htm)$/i.test(path)
            || (!!config.extractAttachments && /\.(png|jpe?g|gif|svg|webp)$/i.test(path)),
        config.decompressionLimits
    );

    // The OPF path is authoritative via META-INF/container.xml; fall back to scanning
    // for any .opf file for malformed archives that skip the container manifest.
    let opfPath: string | undefined;
    const containerFile = files.find(f => /META-INF\/container\.xml$/i.test(f.path));
    if (containerFile) {
        const containerXml = parseXmlString(containerFile.content.toString('utf-8'));
        const rootfile = getFirstElementByTagName(containerXml, 'rootfile');
        opfPath = rootfile ? getAttribute(rootfile, 'full-path') : undefined;
    }
    const opfFile = (opfPath && files.find(f => f.path === opfPath)) || files.find(f => /\.opf$/i.test(f.path));
    if (!opfFile) {
        throw getOfficeError(OfficeErrorType.FILE_CORRUPTED, config, 'epub (no OPF manifest found)');
    }

    const opfDir = opfFile.path.includes('/') ? opfFile.path.substring(0, opfFile.path.lastIndexOf('/') + 1) : '';
    const opfXml = parseXmlString(opfFile.content.toString('utf-8'));

    // ─── Metadata (Dublin Core) ─────────────────────────────────────────────
    const metadata: OfficeMetadata = {};
    const metadataEl = getFirstElementByTagName(opfXml, 'metadata');
    if (metadataEl) {
        const nativeProps: Record<string, any> = {};
        const dcText = (tag: string): string | undefined => getElementsByTagName(metadataEl, tag)[0]?.textContent || undefined;

        const title = dcText('dc:title');
        if (title) { metadata.title = title; nativeProps.title = title; }
        const creator = dcText('dc:creator');
        if (creator) { metadata.author = creator; nativeProps.creator = creator; }
        const description = dcText('dc:description');
        if (description) { metadata.description = description; nativeProps.description = description; }
        const subject = dcText('dc:subject');
        if (subject) { metadata.subject = subject; nativeProps.subject = subject; }
        const dateStr = dcText('dc:date');
        if (dateStr) {
            nativeProps.date = dateStr;
            metadata.created = parseOfficeDate(dateStr) || (isNaN(Date.parse(dateStr)) ? undefined : new Date(dateStr));
        }
        const publisher = dcText('dc:publisher');
        if (publisher) nativeProps.publisher = publisher;
        const language = dcText('dc:language');
        if (language) nativeProps.language = language;
        const identifier = dcText('dc:identifier');
        if (identifier) nativeProps.identifier = identifier;

        // Calibre/EPUB2-style <meta name="..." content="..."> refinements
        for (const metaTag of getElementsByTagName(metadataEl, 'meta')) {
            const name = getAttribute(metaTag, 'name');
            const content = getAttribute(metaTag, 'content');
            if (name && content) nativeProps[name] = content;
        }

        if (Object.keys(nativeProps).length > 0) metadata.nativeProperties = nativeProps;
    }

    // ─── Manifest: id -> {href, mediaType} ──────────────────────────────────
    const manifest = new Map<string, { href: string; mediaType: string }>();
    let coverImageId: string | undefined;
    for (const item of getElementsByTagName(opfXml, 'item')) {
        const id = getAttribute(item, 'id');
        const href = getAttribute(item, 'href');
        const mediaType = getAttribute(item, 'media-type') || '';
        if (id && href) manifest.set(id, { href, mediaType });
        if ((getAttribute(item, 'properties') || '').split(/\s+/).includes('cover-image')) coverImageId = id;
    }
    if (!coverImageId) {
        // EPUB2-style cover declaration: <meta name="cover" content="{manifest id}">
        const coverMeta = metadataEl && getElementsByTagName(metadataEl, 'meta').find(m => getAttribute(m, 'name') === 'cover');
        coverImageId = coverMeta ? getAttribute(coverMeta, 'content') : undefined;
    }

    // ─── Spine: ordered reading order of XHTML documents ────────────────────
    const spineHrefs: string[] = [];
    for (const itemref of getElementsByTagName(opfXml, 'itemref')) {
        const idref = getAttribute(itemref, 'idref');
        const item = idref ? manifest.get(idref) : undefined;
        if (item && /html/i.test(item.mediaType)) spineHrefs.push(item.href);
    }

    const content: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];

    // Map each in-zip image resource by its resolved path, so inline <img> references can
    // be resolved to real bytes (EPUB images are separate files referenced by relative
    // path, unlike DOCX's embedded parts).
    const imageByPath = new Map<string, { content: Buffer; mediaType: string }>();
    if (config.extractAttachments) {
        for (const [, item] of manifest) {
            if (!item.mediaType.startsWith('image/')) continue;
            const p = resolveOpfPath(opfDir, item.href);
            const f = files.find(ff => ff.path === p);
            if (f) imageByPath.set(p, { content: f.content, mediaType: item.mediaType });
        }
    }
    const referencedImagePaths = new Set<string>();

    for (const href of spineHrefs) {
        checkAbortSignal(config.abortSignal);
        const xhtmlPath = resolveOpfPath(opfDir, href.split('#')[0]);
        const xhtmlFile = files.find(f => f.path === xhtmlPath);
        if (!xhtmlFile) continue;

        let xhtml = xhtmlFile.content.toString('utf-8');
        if (config.extractAttachments && imageByPath.size > 0) {
            // Inline each referenced image as a data URI so HtmlParser extracts it as an
            // attachment (with a real image node linked by name) - the same treatment
            // DOCX images get, and what makes the image survive conversion to any format.
            const xhtmlDir = xhtmlPath.includes('/') ? xhtmlPath.substring(0, xhtmlPath.lastIndexOf('/') + 1) : '';
            xhtml = xhtml.replace(/(<img\b[^>]*\bsrc=")([^"]+)(")/gi, (full, pre, src, post) => {
                if (/^(data:|https?:|\/\/)/i.test(src)) return full;
                const resolved = resolveOpfPath(xhtmlDir, src.split('#')[0].split('?')[0]);
                const img = imageByPath.get(resolved);
                if (!img) return full;
                referencedImagePaths.add(resolved);
                return `${pre}data:${img.mediaType};base64,${img.content.toString('base64')}${post}`;
            });
        }

        const chapterAst = await parseHtml(Buffer.from(xhtml, 'utf-8'), config);
        content.push(...chapterAst.content);
        attachments.push(...chapterAst.attachments);
    }

    // Keep manifest images that were NOT referenced inline (e.g. cover art, or images used
    // only as CSS list-style bullets) as attachments so the raw assets aren't lost - DOCX
    // likewise exposes such images as attachments even without an inline image node.
    if (config.extractAttachments) {
        const customProperties: Record<string, string> = {};
        for (const [id, item] of manifest) {
            if (!item.mediaType.startsWith('image/')) continue;
            const p = resolveOpfPath(opfDir, item.href);
            const img = imageByPath.get(p);
            if (!img || referencedImagePaths.has(p)) continue;
            const attachment = createAttachment(item.href.split('/').pop() || item.href, img.content);
            attachments.push(attachment);
            if (id === coverImageId) customProperties.coverImageName = attachment.name;
        }
        if (Object.keys(customProperties).length > 0) {
            metadata.customProperties = { ...metadata.customProperties, ...customProperties };
        }
    }

    const toTextSync = () => content.map(n => {
        const getText = (node: OfficeContentNode): string => {
            if (node.type === 'text' || node.type === 'code') return node.text || '';
            if (node.type === 'break') return '\n';
            if (node.type === 'embed') return (node.metadata as EmbedMetadata)?.url || '';
            if (node.type === 'image') return (node.metadata as ImageMetadata)?.altText || '';
            if (node.children) {
                const isBlock = ['table', 'row', 'list', 'sheet', 'slide', 'admonition'].includes(node.type);
                return node.children.map(getText).join(isBlock ? config.newlineDelimiter : '');
            }
            return '';
        };
        return getText(n);
    }).join(config.newlineDelimiter)
        .replace(/\n{3,}/g, '\n\n');

    return createAST('epub', metadata, content, attachments, config, undefined, toTextSync);
};
