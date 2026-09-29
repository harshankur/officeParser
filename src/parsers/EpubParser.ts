import { EmbedMetadata, FullOfficeParserConfig, ImageMetadata, OfficeAttachment, OfficeContentNode, OfficeErrorType, OfficeMetadata, OfficeParserAST, OfficeWarningType } from '../types.js';
import { createAST } from '../utils/astUtils.js';
import { parseOfficeDate } from '../utils/dateUtils.js';
import { checkAbortSignal, getOfficeError, logWarning } from '../utils/errorUtils.js';
import { createAttachment } from '../utils/imageUtils.js';
import { getAttribute, getElementsByTagName, getFirstElementByTagName, parseXmlString, takeXmlElements } from '../utils/xmlUtils.js';
import { extractFiles } from '../utils/zipUtils.js';
import { parseHtml } from './HtmlParser.js';
import { setOwn } from '../utils/lookupUtils.js';
import { appendAll } from '../utils/nodeListUtils.js';

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
 * An href (a URL, as a manifest's and a chapter's are) as the path of a file in the archive: its
 * percent-escapes decoded (`chapter%201.xhtml` is `chapter 1.xhtml`), and as written when they do not
 * decode. Read as written, a chapter or picture whose name held a space was not found.
 */
const hrefPath = (href: string): string => {
    try {
        return decodeURIComponent(href);
    } catch {
        return href;
    }
};

/**
 * Encryption methods that obfuscate fonts only (IDPF's and Adobe's): a book listing its fonts under them
 * is readable. Any other method a book lists a file under is encryption (DRM) that file cannot be read through.
 */
const FONT_OBFUSCATION = new Set(['http://www.idpf.org/2008/embedding', 'http://ns.adobe.com/pdf/enc#RC']);

/**
 * Parses an EPUB file (a ZIP archive of XHTML content plus an OPF manifest) into the
 * unified OfficeParserAST. Each spine item is parsed via the existing `HtmlParser` and
 * the resulting content/attachments are concatenated in reading order - EPUB is
 * essentially a sequence of XHTML documents, so there's no need for a bespoke content model.
 */
export const parseEpub = async (buffer: Buffer, config: FullOfficeParserConfig): Promise<OfficeParserAST> => {
    checkAbortSignal(config.abortSignal);

    // Chapters by the extensions books give them (`.xml` too); a chapter the spine lists under another
    // is reported as not read (see below), not skipped silently.
    const files = await extractFiles(
        buffer,
        (path) => /META-INF\/(?:container|encryption)\.xml$/i.test(path)
            || /\.opf$/i.test(path)
            || /\.(xhtml|html|htm|xht|xml)$/i.test(path)
            || (!!config.extractAttachments && /\.(png|jpe?g|gif|svg|webp)$/i.test(path)),
        config.decompressionLimits,
        config
    );

    // The OPF path is authoritative via META-INF/container.xml; fall back to scanning
    // for any .opf file for malformed archives that skip the container manifest.
    let opfPath: string | undefined;
    const containerFile = files.find(f => /META-INF\/container\.xml$/i.test(f.path));
    if (containerFile) {
        const containerXml = parseXmlString(containerFile.content.toString('utf-8'), { config });
        const rootfile = getFirstElementByTagName(containerXml, 'rootfile');
        opfPath = rootfile ? getAttribute(rootfile, 'full-path') : undefined;
    }
    const opfFile = (opfPath && files.find(f => f.path === opfPath)) || files.find(f => /\.opf$/i.test(f.path));
    if (!opfFile) {
        throw getOfficeError(OfficeErrorType.REQUIRED_PART_MISSING, config,
            { fileType: 'epub', part: 'OPF package document (.opf)' });
    }

    const opfDir = opfFile.path.includes('/') ? opfFile.path.substring(0, opfFile.path.lastIndexOf('/') + 1) : '';
    const opfXml = parseXmlString(opfFile.content.toString('utf-8'), { config });

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
        // `metadata.language` is where every parser puts the language; the native field stays too.
        if (language) { nativeProps.language = language; metadata.language = language; }
        const identifier = dcText('dc:identifier');
        if (identifier) nativeProps.identifier = identifier;

        // Calibre/EPUB2-style <meta name="..." content="..."> refinements
        for (const metaTag of getElementsByTagName(metadataEl, 'meta')) {
            const name = getAttribute(metaTag, 'name');
            const content = getAttribute(metaTag, 'content');
            if (name && content) setOwn(nativeProps, name, content);
        }

        if (Object.keys(nativeProps).length > 0) metadata.nativeProperties = nativeProps;
    }

    // ─── Manifest: id -> {href, mediaType} ──────────────────────────────────
    // Each item's file: its href (a URL, relative to the package document) as a path in the archive.
    const manifest = new Map<string, { href: string; path: string; mediaType: string }>();
    let coverImageId: string | undefined;
    for (const item of getElementsByTagName(opfXml, 'item')) {
        const id = getAttribute(item, 'id');
        const href = getAttribute(item, 'href');
        const mediaType = getAttribute(item, 'media-type') || '';
        if (id && href) manifest.set(id, { href, path: resolveOpfPath(opfDir, hrefPath(href.split('#')[0])), mediaType });
        if ((getAttribute(item, 'properties') || '').split(/\s+/).includes('cover-image')) coverImageId = id;
    }
    if (!coverImageId) {
        // EPUB2-style cover declaration: <meta name="cover" content="{manifest id}">
        const coverMeta = metadataEl && getElementsByTagName(metadataEl, 'meta').find(m => getAttribute(m, 'name') === 'cover');
        coverImageId = coverMeta ? getAttribute(coverMeta, 'content') : undefined;
    }

    // ─── Spine: ordered reading order of XHTML documents ────────────────────
    const spine: { href: string; path: string }[] = [];
    for (const itemref of getElementsByTagName(opfXml, 'itemref')) {
        const idref = getAttribute(itemref, 'idref');
        const item = idref ? manifest.get(idref) : undefined;
        if (item && /html/i.test(item.mediaType)) spine.push(item);
    }

    // ─── Encryption: the files a book's DRM encrypts (META-INF/encryption.xml) ──
    // Read as they are, an encrypted chapter was its ciphertext decoded as text, with no word of why.
    const encrypted = new Set<string>();
    const encryptionFile = files.find(f => /^META-INF\/encryption\.xml$/i.test(f.path));
    if (encryptionFile) {
        const encryptionXml = parseXmlString(encryptionFile.content.toString('utf-8'), { config });
        // Each EncryptedData names its method and the file it encrypts (by local name, whatever prefix
        // the XML Encryption namespace is given).
        let method: string | undefined;
        for (const element of Array.from(encryptionXml.getElementsByTagName('*')) as Element[]) {
            if (element.localName === 'EncryptedData') method = undefined;
            else if (element.localName === 'EncryptionMethod') method = getAttribute(element, 'Algorithm');
            else if (element.localName === 'CipherReference') {
                const uri = getAttribute(element, 'URI');
                if (uri && !FONT_OBFUSCATION.has(method ?? '')) encrypted.add(resolveOpfPath('', hrefPath(uri)));
            }
        }
    }
    const readable = spine.filter(item => !encrypted.has(item.path));
    if (spine.length && !readable.length) {
        throw getOfficeError(OfficeErrorType.DOCUMENT_DECRYPTION_FAILED, config,
            'the EPUB\'s content is encrypted (DRM, listed in META-INF/encryption.xml); it can only be read by the reading system it is licensed to');
    }

    const content: OfficeContentNode[] = [];
    const attachments: OfficeAttachment[] = [];

    // The package's files by path: each image the manifest lists and each chapter the spine lists is
    // found through it (a scan of every file per item took items x files).
    const fileByPath = new Map<string, (typeof files)[number]>();
    for (const f of files) if (!fileByPath.has(f.path)) fileByPath.set(f.path, f);

    // Map each in-zip image resource by its resolved path, so inline <img> references can
    // be resolved to real bytes (EPUB images are separate files referenced by relative
    // path, unlike DOCX's embedded parts).
    const imageByPath = new Map<string, { content: Buffer; mediaType: string }>();
    if (config.extractAttachments) {
        for (const [, item] of manifest) {
            if (!item.mediaType.startsWith('image/') || encrypted.has(item.path)) continue;
            const f = fileByPath.get(item.path);
            if (f) imageByPath.set(item.path, { content: f.content, mediaType: item.mediaType });
        }
    }
    // Attachment names, each used once in the book: a file's own name, else that name numbered. The
    // numbering resumes where it left off for each name: tried from 2 each time, the pictures every
    // chapter names image_1.png, image_2.png, ... took chapters squared (1.5 MB, seven minutes).
    const usedNames = new Set<string>();
    const nextNumber = new Map<string, number>();
    const uniqueName = (name: string): string => {
        if (!usedNames.has(name)) { usedNames.add(name); return name; }
        const dot = name.lastIndexOf('.');
        const [stem, extension] = dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ''];
        let n = nextNumber.get(name) ?? 2;
        let candidate = `${stem}-${n}${extension}`;
        while (usedNames.has(candidate)) candidate = `${stem}-${++n}${extension}`;
        nextNumber.set(name, n + 1);
        usedNames.add(candidate);
        return candidate;
    };
    // The attachment each of the book's images is, made the first time a chapter shows the image.
    const imageAttachmentNames = new Map<string, string>();

    // Each chapter once, in the order the spine first lists it: a spine listing one chapter many
    // times (which EPUB does not allow) read it, and wrote its content, again for each.
    const readChapters = new Set<string>();
    for (const { href, path: xhtmlPath } of spine) {
        checkAbortSignal(config.abortSignal);
        if (readChapters.has(xhtmlPath)) continue;
        readChapters.add(xhtmlPath);
        // A chapter that cannot be read is reported, not left out without a word.
        if (encrypted.has(xhtmlPath)) {
            logWarning(OfficeWarningType.CONTENT_PART_NOT_READ, config, { part: href, reason: 'the chapter is encrypted (DRM, listed in META-INF/encryption.xml)' });
            continue;
        }
        const xhtmlFile = fileByPath.get(xhtmlPath);
        if (!xhtmlFile) {
            logWarning(OfficeWarningType.CONTENT_PART_NOT_READ, config, { part: href, reason: 'the book lists the chapter, but its file is not in the archive (or has an extension other than .xhtml, .html, .htm, .xht or .xml)' });
            continue;
        }

        // A picture showing one of the book's images is linked to that image's attachment (made
        // once), so the image survives conversion to any format, as a DOCX image does.
        const xhtmlDir = xhtmlPath.includes('/') ? xhtmlPath.substring(0, xhtmlPath.lastIndexOf('/') + 1) : '';
        const imageAttachment = (src: string): string | undefined => {
            if (!config.extractAttachments || /^(data:|https?:|\/\/)/i.test(src)) return undefined;
            const resolved = resolveOpfPath(xhtmlDir, hrefPath(src.split('#')[0].split('?')[0]));
            const img = imageByPath.get(resolved);
            if (!img) return undefined;
            let name = imageAttachmentNames.get(resolved);
            if (!name) {
                const attachment = createAttachment(uniqueName(resolved.split('/').pop() || resolved), img.content);
                attachments.push(attachment);
                name = attachment.name;
                imageAttachmentNames.set(resolved, name);
            }
            return name;
        };

        // A chapter's elements count against the document's budget, as its package XML does (see
        // maxXmlElements): read by the HTML parser, they were not counted, and a chapter of 10 million
        // empty elements (118 KB of EPUB) ran the process out of memory.
        takeXmlElements(xhtmlFile.content.toString('utf8'), config);
        // A chapter's own inline pictures (data URIs) are numbered from 1 in each chapter: each is named
        // as it is read, so no two of the book's attachments share a name. (Named again afterwards, by
        // name, a picture of the book's own named `image_1.png` was renamed with the chapter's first.)
        const chapterAst = await parseHtml(xhtmlFile.content, config, { imageAttachment, attachmentName: uniqueName });
        appendAll(attachments, chapterAst.attachments);
        appendAll(content, chapterAst.content);
    }

    // Keep manifest images that were NOT referenced inline (e.g. cover art, or images used
    // only as CSS list-style bullets) as attachments so the raw assets aren't lost - DOCX
    // likewise exposes such images as attachments even without an inline image node.
    if (config.extractAttachments) {
        const customProperties: Record<string, string> = {};
        for (const [id, item] of manifest) {
            if (!item.mediaType.startsWith('image/')) continue;
            const img = imageByPath.get(item.path);
            if (!img) continue;
            const shown = imageAttachmentNames.get(item.path);
            if (shown !== undefined) {
                if (id === coverImageId) customProperties.coverImageName = shown;
                continue;
            }
            const attachment = createAttachment(uniqueName(item.path.split('/').pop() || item.path), img.content);
            attachments.push(attachment);
            if (id === coverImageId) customProperties.coverImageName = attachment.name;
        }
        if (Object.keys(customProperties).length > 0) {
            metadata.customProperties = { ...metadata.customProperties, ...customProperties };
        }
    }

    return createAST('epub', metadata, content, attachments, config, undefined);
};
