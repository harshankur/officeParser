/**
 * Exhaustive test suite for officeParser.
 * Covers every AST node type, metadata field, text formatting flag,
 * and round-trip output correctness for Markdown, HTML, CSV, and RTF formats.
 */

import { OfficeParser } from '../src/OfficeParser';
import { OfficeGenerator } from '../src/OfficeGenerator';
import { OfficeTemplate } from '../src/OfficeTemplate';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import * as assert from 'assert';
import * as path from 'path';
import * as fs from 'fs';
import type { OfficeContentNode, OfficeParserAST } from '../src/types';
import { parseXmlString } from '../src/utils/xmlUtils';
import { decodeBase64, hexColor, isHeaderRow, lengthToPt, resolveZipInstant, sniffImageSize, toBookmarkNameRaw } from '../src/utils/officeGenUtils';

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Recursively collect every node in the AST (content + notes). */
function collectAllNodes(ast: OfficeParserAST): OfficeContentNode[] {
    const result: OfficeContentNode[] = [];
    const walk = (nodes: OfficeContentNode[]) => {
        for (const node of nodes) {
            result.push(node);
            if (node.children) walk(node.children);
            if (node.notes) walk(node.notes);
            if (node.comments) walk(node.comments);
        }
    };
    walk(ast.content);
    return result;
}

function assertExists<T>(
    items: T[],
    predicate: (item: T) => boolean,
    message: string
): T {
    const found = items.find(predicate);
    if (!found) {
        throw new assert.AssertionError({ message: `FAIL: ${message}` });
    }
    return found;
}

// ─── Markdown ────────────────────────────────────────────────────────────────

async function testMarkdown(): Promise<void> {
    console.log('\n=== Running Exhaustive Markdown Tests ===');
    const filePath = path.join(__dirname, 'files/exhaustive/markdown.md');
    const ast = await OfficeParser.parseOffice(filePath);
    const nodes = collectAllNodes(ast);

    // ── Metadata / YAML Frontmatter ──────────────────────────────────────────
    assert.strictEqual(ast.metadata.title, 'Exhaustive Markdown Test', 'MD: metadata.title');
    assert.strictEqual(ast.metadata.author, 'Test Author', 'MD: metadata.author');
    assert.strictEqual(ast.metadata.description, 'Tests every markdown feature', 'MD: metadata.description');

    // customProperties.tags must be an Array
    const tags = ast.metadata.customProperties?.['tags'];
    assert.ok(Array.isArray(tags), 'MD: customProperties.tags is an array');
    assert.ok((tags as string[]).length >= 2, 'MD: tags array has at least 2 items');

    // nativeProperties must contain all front-matter keys
    assert.ok(ast.metadata.nativeProperties?.['tags'] !== undefined, 'MD: nativeProperties.tags');
    assert.ok(ast.metadata.nativeProperties?.['version'] !== undefined, 'MD: nativeProperties.version');

    // ── Headings H1–H6 ───────────────────────────────────────────────────────
    const headings = nodes.filter(n => n.type === 'heading');
    assert.ok(headings.length >= 6, `MD: At least 6 headings, got ${headings.length}`);
    for (let level = 1; level <= 6; level++) {
        assertExists(headings, n => (n.metadata as any)?.level === level, `MD: heading level ${level}`);
    }
    // H1 has anchorIds from {#h1-anchor}
    const h1 = assertExists(headings, n => (n.metadata as any)?.level === 1, 'MD: H1 heading');
    assert.ok(
        Array.isArray((h1.metadata as any)?.anchorIds) && (h1.metadata as any).anchorIds.length > 0,
        'MD: H1 has anchorIds'
    );

    // ── Paragraphs ────────────────────────────────────────────────────────────
    const paragraphs = nodes.filter(n => n.type === 'paragraph');
    assert.ok(paragraphs.length >= 1, 'MD: Has paragraphs');

    // Right-aligned paragraph
    assertExists(
        paragraphs,
        n => (n.metadata as any)?.alignment === 'right',
        'MD: paragraph with right alignment'
    );

    // ── Text formatting ───────────────────────────────────────────────────────
    const textNodes = nodes.filter(n => n.type === 'text');
    assertExists(textNodes, n => n.formatting?.bold === true, 'MD: bold text node');
    assertExists(textNodes, n => n.formatting?.italic === true, 'MD: italic text node');
    assertExists(textNodes, n => n.formatting?.strikethrough === true, 'MD: strikethrough text node');
    assertExists(textNodes, n => n.formatting?.underline === true, 'MD: underline text node');
    assertExists(textNodes, n => n.formatting?.subscript === true, 'MD: subscript text node');
    assertExists(textNodes, n => n.formatting?.superscript === true, 'MD: superscript text node');
    // Inline code → font: 'monospace'
    assertExists(textNodes, n => n.formatting?.font === 'monospace', 'MD: monospace (inline code) text node');

    // ── Lists ─────────────────────────────────────────────────────────────────
    const listNodes = nodes.filter(n => n.type === 'list');
    assert.ok(listNodes.length >= 6, `MD: At least 6 list nodes, got ${listNodes.length}`);
    assertExists(listNodes, n => (n.metadata as any)?.listType === 'unordered', 'MD: unordered list');
    assertExists(listNodes, n => (n.metadata as any)?.listType === 'ordered', 'MD: ordered list');
    // Nested list indentation
    assertExists(listNodes, n => (n.metadata as any)?.indentation >= 1, 'MD: nested list (indentation>=1)');
    // Task lists
    assertExists(listNodes, n => (n.metadata as any)?.isTask === true && (n.metadata as any)?.checked === true, 'MD: checked task list item');
    assertExists(listNodes, n => (n.metadata as any)?.isTask === true && (n.metadata as any)?.checked === false, 'MD: unchecked task list item');
    // itemIndex is a number
    assert.ok(listNodes.every(n => typeof (n.metadata as any)?.itemIndex === 'number'), 'MD: all list items have itemIndex');
    // Exact itemIndex values, so a nested-list counter regression (e.g. a level-1 counter
    // leaking across level-0 siblings) is actually caught rather than merely "is a number".
    const findListItem = (text: string) => assertExists(listNodes, n => n.text === text, `MD: list item "${text}"`);
    assert.strictEqual((findListItem('Unordered item A').metadata as any)?.itemIndex, 0, 'MD: "Unordered item A" itemIndex 0');
    assert.strictEqual((findListItem('Unordered item B').metadata as any)?.itemIndex, 1, 'MD: "Unordered item B" itemIndex 1');
    assert.strictEqual((findListItem('Nested unordered item').metadata as any)?.itemIndex, 0, 'MD: "Nested unordered item" itemIndex 0');
    assert.strictEqual((findListItem('Unordered item C').metadata as any)?.itemIndex, 2, 'MD: "Unordered item C" itemIndex 2');
    assert.strictEqual((findListItem('Ordered item 1').metadata as any)?.itemIndex, 0, 'MD: "Ordered item 1" itemIndex 0');
    assert.strictEqual((findListItem('Ordered item 2').metadata as any)?.itemIndex, 1, 'MD: "Ordered item 2" itemIndex 1');
    assert.strictEqual((findListItem('Nested ordered item').metadata as any)?.itemIndex, 0, 'MD: "Nested ordered item" itemIndex 0');
    assert.strictEqual((findListItem('Ordered item 3').metadata as any)?.itemIndex, 2, 'MD: "Ordered item 3" itemIndex 2');

    // ── Definition lists ──────────────────────────────────────────────────────
    const defLists = nodes.filter(n => n.type === 'definitionList');
    assert.ok(defLists.length >= 1, 'MD: Has definitionList nodes');
    const defTerms = nodes.filter(n => n.type === 'definitionTerm');
    assert.ok(defTerms.length >= 2, `MD: At least 2 definitionTerm nodes, got ${defTerms.length}`);
    const defDescs = nodes.filter(n => n.type === 'definitionDescription');
    assert.ok(defDescs.length >= 2, `MD: At least 2 definitionDescription nodes, got ${defDescs.length}`);

    // ── Admonitions ───────────────────────────────────────────────────────────
    const admonitions = nodes.filter(n => n.type === 'admonition');
    // 5 GitHub-style + 1 GLFM :::danger = 6 total
    assert.ok(admonitions.length >= 6, `MD: At least 6 admonitions (5 GH + 1 GLFM), got ${admonitions.length}`);
    for (const adType of ['note', 'tip', 'important', 'warning', 'caution'] as const) {
        assertExists(admonitions, n => (n.metadata as any)?.admonitionType === adType, `MD: admonition type '${adType}'`);
    }
    // GLFM :::danger maps to 'caution' - we should have at least 2 'caution' entries
    const cautionCount = admonitions.filter(n => (n.metadata as any)?.admonitionType === 'caution').length;
    assert.ok(cautionCount >= 2, `MD: At least 2 'caution' admonitions (one GH, one GLFM danger), got ${cautionCount}`);
    // sourceSyntax provenance: GitHub `> [!TYPE]` vs GLFM `:::type` must be distinguishable
    assertExists(admonitions, n => (n.metadata as any)?.admonitionType === 'note' && (n.metadata as any)?.sourceSyntax === 'github', 'MD: GitHub admonition has sourceSyntax "github"');
    assertExists(admonitions, n => (n.metadata as any)?.admonitionType === 'caution' && (n.metadata as any)?.sourceSyntax === 'gitlab', 'MD: GLFM :::danger admonition has sourceSyntax "gitlab"');

    // ── Code blocks ───────────────────────────────────────────────────────────
    const codeNodes = nodes.filter(n => n.type === 'code');
    assert.ok(codeNodes.length >= 3, `MD: At least 3 code nodes (2 fenced + 1 inline math + 1 block math), got ${codeNodes.length}`);
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'typescript', 'MD: code block with typescript language');
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'python', 'MD: code block with python language');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'inline', 'MD: inline math code node');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'block', 'MD: block math code node');

    // ── Tables ────────────────────────────────────────────────────────────────
    const tables = nodes.filter(n => n.type === 'table');
    assert.ok(tables.length >= 2, `MD: At least 2 tables (pipe + HTML), got ${tables.length}`);
    // HTML table with data-align="center"
    assertExists(tables, n => (n.metadata as any)?.align === 'center', 'MD: table with align=center');

    const rows = nodes.filter(n => n.type === 'row');
    assert.ok(rows.length >= 4, `MD: At least 4 rows, got ${rows.length}`);

    const cells = nodes.filter(n => n.type === 'cell');
    assert.ok(cells.length >= 6, `MD: At least 6 cells, got ${cells.length}`);
    // HTML table cells with colspan and rowspan
    assertExists(cells, n => (n.metadata as any)?.colSpan >= 2, 'MD: cell with colSpan>=2');
    assertExists(cells, n => (n.metadata as any)?.rowSpan >= 2, 'MD: cell with rowSpan>=2');

    // ── Image ─────────────────────────────────────────────────────────────────
    const images = nodes.filter(n => n.type === 'image');
    assert.ok(images.length >= 1, 'MD: Has image nodes');
    const img = assertExists(images, n => (n.metadata as any)?.url?.includes('example.com'), 'MD: image with url');
    assert.ok((img.metadata as any)?.altText, 'MD: image has altText');
    assert.ok((img.metadata as any)?.width, 'MD: image has width');
    assert.ok((img.metadata as any)?.align, 'MD: image has align');

    // ── Embed (YouTube) ───────────────────────────────────────────────────────
    const embeds = nodes.filter(n => n.type === 'embed');
    assert.ok(embeds.length >= 1, 'MD: Has embed nodes');
    const embed = assertExists(embeds, n => (n.metadata as any)?.embedType === 'youtube', 'MD: youtube embed');
    assert.ok((embed.metadata as any)?.videoId, 'MD: embed has videoId');
    assert.ok((embed.metadata as any)?.width, 'MD: embed has width');

    // ── Text metadata: links ───────────────────────────────────────────────
    // MD parser always gives linkType='external' for [text](url) links (even #anchor ones)
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'external' && (n.metadata as any)?.link?.startsWith('https://'), 'MD: external https link text node');
    // #anchor links also get linkType=external in the MD parser
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'external' && (n.metadata as any)?.link?.startsWith('#'), 'MD: anchor (#) link text node');
    // wikilinks always get linkType='internal'
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'internal' && (n.metadata as any)?.wikilink === true, 'MD: wikilink has linkType=internal');

    // ── Wikilinks ─────────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.wikilink === true, 'MD: wikilink text node');
    // Both bare [[WikiPage]] and [[WikiPage|Alias Text]]
    const wikilinks = textNodes.filter(n => (n.metadata as any)?.wikilink === true);
    assert.ok(wikilinks.length >= 2, `MD: At least 2 wikilinks, got ${wikilinks.length}`);

    // ── Citations ─────────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.citationKey !== undefined, 'MD: citation text node with citationKey');

    // ── Footnotes ─────────────────────────────────────────────────────────────
    const noteNodes = nodes.filter(n => n.type === 'note');
    assert.ok(noteNodes.length >= 1, 'MD: Has note nodes');
    assertExists(noteNodes, n => (n.metadata as any)?.noteType === 'footnote', 'MD: footnote note node');
    // Multi-line definition: the indented continuation lines fold into one note rather than
    // splitting off as stray blocks.
    const mlNote = assertExists(noteNodes, n => (n.metadata as any)?.noteId === 'fnML', 'MD: multi-line footnote node');
    for (const frag of ['First line', 'Second line', 'Third line']) {
        assert.ok((mlNote.text || '').includes(frag), `MD: multi-line footnote kept "${frag}" in one definition`);
    }

    // ── Abbreviations ─────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.abbreviationTitle !== undefined, 'MD: abbreviation text node');

    // ── Break (horizontal rule) ───────────────────────────────────────────────
    const breaks = nodes.filter(n => n.type === 'break');
    assert.ok(breaks.length >= 1, 'MD: Has break nodes');

    // ── Blockquote paragraph ──────────────────────────────────────────────────
    assertExists(paragraphs, n => (n.metadata as any)?.style === 'Quote', 'MD: blockquote paragraph with style=Quote');

    // ── Nested blockquotes (2-level, 3-level) ────────────────────────────────
    const quoteParas = paragraphs.filter(n => (n.metadata as any)?.style === 'Quote');
    const quoteText = quoteParas.map(p => (p.children || []).map((c: any) => c.text || '').join('')).join(' ');
    assert.ok(quoteText.includes('Two-level nested blockquote'), 'MD: 2-level nested blockquote text present');
    assert.ok(quoteText.includes('Three-level nested blockquote'), 'MD: 3-level nested blockquote text present');
    assert.ok(!quoteText.includes('>'), 'MD: nested blockquotes fully unwrapped (no literal ">")');

    // ── Paren-marker (')') ordered list ──────────────────────────────────────
    assertExists(listNodes, n => n.text === 'Paren-marker ordered item one' && (n.metadata as any)?.listType === 'ordered', 'MD: ")"-marker ordered list item');

    // ── Short-cell table separator (|-|-|) ───────────────────────────────────
    assertExists(tables, n => (n.children || []).some((row: any) => (row.children || []).some((cell: any) => (cell.children || []).some((c: any) => c.text === 'C1'))), 'MD: short-cell-separator table parsed');

    // ── Tilde-fenced code block ───────────────────────────────────────────────
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'javascript' && n.text?.includes('tilde fence') === true, 'MD: tilde-fenced code block');

    // ── HR-anchoring edge case (trailing text after hyphens is NOT an Hr) ────
    assert.ok(paragraphs.some(n => (n.children || []).some((c: any) => c.text?.includes('not actually a horizontal rule'))), 'MD: "----- trailing text" parsed as paragraph, not Hr');

    // ── Nested-list itemIndex leak regression (a sibling's nested counter must not
    //    carry over to the next sibling's own nested list) ────────────────────
    assert.strictEqual((findListItem('Sibling parent Alpha').metadata as any)?.itemIndex, 0, 'MD: "Sibling parent Alpha" itemIndex 0');
    assert.strictEqual((findListItem('Alpha child one').metadata as any)?.itemIndex, 0, 'MD: "Alpha child one" itemIndex 0');
    assert.strictEqual((findListItem('Alpha child two').metadata as any)?.itemIndex, 1, 'MD: "Alpha child two" itemIndex 1');
    assert.strictEqual((findListItem('Sibling parent Beta').metadata as any)?.itemIndex, 1, 'MD: "Sibling parent Beta" itemIndex 1');
    assert.strictEqual((findListItem('Beta child one').metadata as any)?.itemIndex, 0, 'MD: "Beta child one" itemIndex 0 (must NOT continue Alpha\'s children counter)');

    // ── Backslash escapes ────────────────────────────────────────────────────
    const escapeText = paragraphs.map(p => (p.children || []).map((c: any) => c.text || '').join('')).join(' ');
    assert.ok(escapeText.includes('*not bold*') && escapeText.includes('_not italic_') && escapeText.includes('`not code`') && escapeText.includes('[not a link]'), 'MD: backslash-escaped punctuation renders literally');

    // ── Reference-style links/images ─────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.link === 'https://example.com/reference', 'MD: explicit reference link resolved');
    assertExists(textNodes, n => (n.metadata as any)?.link === 'https://example.com/shortcut', 'MD: shortcut reference link resolved');
    assertExists(images, n => (n.metadata as any)?.url === 'https://example.com/ref-image.png', 'MD: reference-style image resolved');
    assert.ok(escapeText.includes('[unresolved reference][nowhere]') || nodes.some(n => (n.text || '').includes('[unresolved reference][nowhere]')), 'MD: unresolved explicit reference falls back to literal text');
    assert.ok(nodes.some(n => (n.text || '').includes('[bare bracket]')), 'MD: unresolved shortcut reference falls back to literal text');

    // ── Underscore emphasis ───────────────────────────────────────────────────
    assertExists(textNodes, n => n.formatting?.italic === true && n.text === 'underscore italic', 'MD: underscore italic');
    assertExists(textNodes, n => n.formatting?.bold === true && n.text === 'underscore bold', 'MD: underscore bold');

    // ── Multi-backtick inline code span ──────────────────────────────────────
    assertExists(textNodes, n => n.formatting?.font === 'monospace' && n.text === 'code with a ` backtick inside', 'MD: multi-backtick code span preserves embedded backtick');

    // ── HTML entity decoding ──────────────────────────────────────────────────
    assert.ok(escapeText.includes('Fish & Chips') && escapeText.includes('Q&A'), 'MD: bare "&" in ordinary text left untouched');
    assert.ok(escapeText.includes('& < > \'') && escapeText.includes('❤'), 'MD: named and numeric/hex entities decoded');
    assert.ok(escapeText.includes('&#999999999;') && escapeText.includes('&#x999999999;'), 'MD: out-of-bounds entity references preserved raw');

    // ── Hard vs soft line break ───────────────────────────────────────────────
    assertExists(nodes, n => n.type === 'break' && (n.metadata as any)?.breakType === 'carriageReturn', 'MD: hard line break emits a break node');

    // ── Setext headings ───────────────────────────────────────────────────────
    assertExists(headings, n => (n.metadata as any)?.level === 1 && n.text === 'Setext Heading One', 'MD: setext H1 (=== underline)');
    assertExists(headings, n => (n.metadata as any)?.level === 2 && n.text === 'Setext Heading Two', 'MD: setext H2 (--- underline)');

    // ── <url> autolink ────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.link === 'https://example.com/autolink', 'MD: <url> autolink resolved');

    // ── List-item continuation line ──────────────────────────────────────────
    assertExists(listNodes, n => n.text === 'Continuation parent item continuation text merged into the parent item', 'MD: list-item continuation line merged into item text');

    // ── Standalone indented code block ───────────────────────────────────────
    assertExists(codeNodes, n => !(n.metadata as any)?.language && !(n.metadata as any)?.math && n.text === 'indented code block line one\nindented code block line two', 'MD: standalone 4-space-indented code block');

    // ── Roundtrip: generate to MD ────────────────────────────────────────────
    const result = await OfficeGenerator.generate(ast, 'md');
    const mdOutput = result.value as string;
    assert.ok(mdOutput.includes('> [!NOTE]'), 'MD roundtrip: note admonition preserved');
    assert.ok(mdOutput.includes('> [!TIP]'), 'MD roundtrip: tip admonition preserved');
    assert.ok(mdOutput.includes('> [!IMPORTANT]'), 'MD roundtrip: important admonition preserved');
    assert.ok(mdOutput.includes('> [!WARNING]'), 'MD roundtrip: warning admonition preserved');
    assert.ok(mdOutput.includes('> [!CAUTION]'), 'MD roundtrip: caution admonition preserved');
    assert.ok(mdOutput.includes('**'), 'MD roundtrip: bold marker');
    assert.ok(mdOutput.includes('*'), 'MD roundtrip: italic marker');
    assert.ok(mdOutput.includes('```'), 'MD roundtrip: fenced code');
    assert.ok(mdOutput.includes('|'), 'MD roundtrip: pipe table');
    assert.ok(mdOutput.includes('[[WikiPage]]'), 'MD roundtrip: wikilink');
    assert.ok(mdOutput.includes('[@smith2023]'), 'MD roundtrip: citation');

    // Delimiter-adjacent constructs now escape or allowlist their content, so assert the
    // *legitimate* forms survive intact. Each of these is a shape a naive guard would break:
    // dropping `<` would corrupt the comparison, stripping the attribute list would lose the
    // width, and validating the footnote id too tightly would renumber a real label.
    assert.ok(mdOutput.includes('$E=mc^2$'), 'MD roundtrip: inline math preserved');
    assert.ok(mdOutput.includes('$a < b$'), 'MD roundtrip: LaTeX comparison operator not escaped away');
    assert.ok(mdOutput.includes('a^2 + b^2 = c^2'), 'MD roundtrip: block math content preserved');
    assert.ok(mdOutput.includes('*[ABBR]: Abbreviation Full Title'), 'MD roundtrip: abbreviation definition preserved');
    assert.ok(mdOutput.includes('{width=50px align=center}'), 'MD roundtrip: image attribute list preserved');
    assert.ok(mdOutput.includes('[^fn1]'), 'MD roundtrip: label-shaped footnote id preserved, not renumbered');
    assert.ok(mdOutput.includes('[[WikiPage|Alias Text]]'), 'MD roundtrip: wikilink alias preserved');
    // The multi-line footnote must re-parse as one note (the generator indents continuation lines,
    // so a bare newline can't end the definition early).
    const reAst = await OfficeParser.parseOffice(Buffer.from(mdOutput), { fileType: 'md' });
    const reNote = collectAllNodes(reAst).find(n => n.type === 'note' && (n.metadata as any)?.noteId === 'fnML');
    assert.ok(reNote && ['First line', 'Second line', 'Third line'].every(f => (reNote!.text || '').includes(f)),
        'MD roundtrip: multi-line footnote survives generate -> reparse as one definition');

    // ── Roundtrip: the bug-fix-pass additions survive generate() ────────────
    assert.ok(mdOutput.includes('  \n'), 'MD roundtrip: hard line break emits two trailing spaces');
    assert.ok(mdOutput.includes('https://example.com/reference'), 'MD roundtrip: resolved reference-link URL preserved');
    assert.ok(mdOutput.includes('https://example.com/ref-image.png'), 'MD roundtrip: resolved reference-image URL preserved');
    assert.ok(mdOutput.includes('https://example.com/autolink'), 'MD roundtrip: autolink URL preserved');
    assert.ok(mdOutput.includes('code with a ` backtick inside'), 'MD roundtrip: multi-backtick code content preserved');
    assert.ok(mdOutput.includes('underscore italic') && mdOutput.includes('underscore bold'), 'MD roundtrip: underscore-emphasized text preserved');
    assert.ok(mdOutput.includes('not bold') && mdOutput.includes('not code'), 'MD roundtrip: decoded escaped-punctuation text preserved');
    assert.ok(mdOutput.includes('❤'), 'MD roundtrip: decoded HTML entity character preserved');
    assert.ok(mdOutput.includes('Continuation parent item') && mdOutput.includes('continuation text merged into the parent item'), 'MD roundtrip: list-item continuation text preserved');
    assert.ok(mdOutput.includes('indented code block line one'), 'MD roundtrip: indented-code-block content preserved');
    assert.ok(mdOutput.includes('Setext Heading One') && mdOutput.includes('Setext Heading Two'), 'MD roundtrip: setext heading text preserved');
    assert.ok(mdOutput.includes('Two-level nested blockquote') && mdOutput.includes('Three-level nested blockquote'), 'MD roundtrip: nested-blockquote text preserved');
    assert.ok(mdOutput.includes('Paren-marker ordered item one'), 'MD roundtrip: ")"-marker ordered list item preserved');

    console.log('  Markdown: All assertions passed ✓');
}

// ─── HTML ────────────────────────────────────────────────────────────────────

async function testHtml(): Promise<void> {
    console.log('\n=== Running Exhaustive HTML Tests ===');
    const filePath = path.join(__dirname, 'files/exhaustive/html.html');
    const ast = await OfficeParser.parseOffice(filePath);
    const nodes = collectAllNodes(ast);

    // ── Metadata ──────────────────────────────────────────────────────────────
    assert.strictEqual(ast.metadata.title, 'Exhaustive HTML Test', 'HTML: metadata.title');
    assert.strictEqual(ast.metadata.author, 'Test Author', 'HTML: metadata.author');
    assert.strictEqual(ast.metadata.description, 'Exhaustive HTML test description', 'HTML: metadata.description');
    assert.ok(ast.metadata.nativeProperties?.['author'] !== undefined, 'HTML: nativeProperties.author');
    // Custom meta properties
    const customProps = ast.metadata.customProperties;
    assert.ok(customProps !== undefined, 'HTML: Has customProperties');
    assert.strictEqual(customProps?.['version'], 1, 'HTML: customProperties.version === 1 (number)');
    assert.strictEqual(customProps?.['reviewed'], true, 'HTML: customProperties.reviewed === true (boolean)');

    // ── Headings H1–H6 ────────────────────────────────────────────────────────
    const headings = nodes.filter(n => n.type === 'heading');
    assert.ok(headings.length >= 6, `HTML: At least 6 headings, got ${headings.length}`);
    for (let level = 1; level <= 6; level++) {
        assertExists(headings, n => (n.metadata as any)?.level === level, `HTML: heading level ${level}`);
    }
    // H1 with id="heading-1" → anchorIds
    const h1 = assertExists(headings, n => (n.metadata as any)?.level === 1, 'HTML: H1 heading');
    assert.ok(
        Array.isArray((h1.metadata as any)?.anchorIds) && (h1.metadata as any).anchorIds[0] === 'heading-1',
        'HTML: H1 anchorId === "heading-1"'
    );

    // ── Paragraphs ────────────────────────────────────────────────────────────
    const paragraphs = nodes.filter(n => n.type === 'paragraph');
    assert.ok(paragraphs.length >= 3, `HTML: At least 3 paragraphs, got ${paragraphs.length}`);
    // center alignment via align attribute
    assertExists(paragraphs, n => (n.metadata as any)?.alignment === 'center', 'HTML: center-aligned paragraph');
    // right alignment via style
    assertExists(paragraphs, n => (n.metadata as any)?.alignment === 'right', 'HTML: right-aligned paragraph');

    // ── Text formatting ───────────────────────────────────────────────────────
    const textNodes = nodes.filter(n => n.type === 'text');
    assertExists(textNodes, n => n.formatting?.bold === true, 'HTML: bold text');
    assertExists(textNodes, n => n.formatting?.italic === true, 'HTML: italic text');
    assertExists(textNodes, n => n.formatting?.underline === true, 'HTML: underline text');
    assertExists(textNodes, n => n.formatting?.strikethrough === true, 'HTML: strikethrough text');
    assertExists(textNodes, n => n.formatting?.subscript === true, 'HTML: subscript text');
    assertExists(textNodes, n => n.formatting?.superscript === true, 'HTML: superscript text');

    // ── Break (<br>) ──────────────────────────────────────────────────────────
    const breaks = nodes.filter(n => n.type === 'break');
    assert.ok(breaks.length >= 1, 'HTML: Has break nodes');
    // A <br> is a hard line break -> carriageReturn, so the md generator emits `  \n` (which
    // re-imports as a <br>) rather than a bare `\n` that would collapse to a space (8.H).
    assertExists(breaks, n => (n.metadata as any)?.breakType === 'carriageReturn', 'HTML: <br> is a carriageReturn (hard) break');

    // ── Lists (unordered/ordered) ─────────────────────────────────────────────
    const listNodes = nodes.filter(n => n.type === 'list');
    assert.ok(listNodes.length >= 4, `HTML: At least 4 list nodes, got ${listNodes.length}`);
    assertExists(listNodes, n => (n.metadata as any)?.listType === 'unordered', 'HTML: unordered list');
    assertExists(listNodes, n => (n.metadata as any)?.listType === 'ordered', 'HTML: ordered list');
    // Nested list
    assertExists(listNodes, n => (n.metadata as any)?.indentation >= 1, 'HTML: nested list (indentation>=1)');

    // ── Task lists ────────────────────────────────────────────────────────────
    assertExists(listNodes, n => (n.metadata as any)?.isTask === true && (n.metadata as any)?.checked === true, 'HTML: checked task list item');
    assertExists(listNodes, n => (n.metadata as any)?.isTask === true && (n.metadata as any)?.checked === false, 'HTML: unchecked task list item');

    // ── Definition lists ──────────────────────────────────────────────────────
    const defLists = nodes.filter(n => n.type === 'definitionList');
    assert.ok(defLists.length >= 1, 'HTML: Has definitionList nodes');
    const defTerms = nodes.filter(n => n.type === 'definitionTerm');
    assert.ok(defTerms.length >= 1, `HTML: At least 1 definitionTerm, got ${defTerms.length}`);
    const defDescs = nodes.filter(n => n.type === 'definitionDescription');
    assert.ok(defDescs.length >= 1, `HTML: At least 1 definitionDescription, got ${defDescs.length}`);

    // ── Code blocks ───────────────────────────────────────────────────────────
    const codeNodes = nodes.filter(n => n.type === 'code');
    assert.ok(codeNodes.length >= 2, `HTML: At least 2 code nodes, got ${codeNodes.length}`);
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'javascript', 'HTML: javascript code block');
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'python', 'HTML: python code block');

    // ── Mermaid (attribute-driven) ────────────────────────────────────────────
    // div[data-mermaid] / div.mermaid / pre.mermaid all map to a mermaid-language code node;
    // previously the div flattened to paragraph text.
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'mermaid' && n.text === 'graph TD; A-->B;', 'HTML: mermaid div (class + attr + text content)');
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'mermaid' && n.text === 'pie showData', 'HTML: mermaid div[data-mermaid] with empty body falls back to the attribute');
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'mermaid' && (n.text || '').includes('flowchart'), 'HTML: pre.mermaid maps to a mermaid code node');
    // A `class="mermaid"` div with no diagram source (nested elements only) must NOT become an
    // empty mermaid code node; its content falls through to generic handling and survives.
    assert.ok(!codeNodes.some(n => (n.metadata as any)?.language === 'mermaid' && !(n.text || '').trim()), 'HTML: no empty mermaid code node from a styling-only .mermaid div');
    assert.ok(nodes.some(n => n.type === 'text' && (n.text || '').includes('Not a diagram')), 'HTML: styling-only .mermaid div content preserved (fell through)');

    // ── Math ─────────────────────────────────────────────────────────────────
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'inline', 'HTML: inline math code node');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'block', 'HTML: block math code node');
    // Attribute-driven math: raw LaTeX in data-math, mode from the class token, undelimited body.
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'inline' && n.text === '\\alpha+\\beta', 'HTML: attribute-driven inline math (latex in data-math, class names the mode)');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'block' && n.text === '\\sum_{i=0}^n i', 'HTML: attribute-driven block math');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'inline' && n.text === '\\delta', 'HTML: attribute-driven math still strips $ delimiters from the text content');

    // Native MathML, which reaches every EPUB too since EpubParser parses each spine item with
    // HtmlParser. Each assertion below names the construct whose *structure* is the point: it is
    // not enough that the digits survive, because the bug being locked out here was structure
    // being silently flattened away while the characters came through looking fine.
    const mathNodes = codeNodes.filter(n => (n.metadata as any)?.math);
    assertExists(mathNodes, n => n.text === '\\frac12', 'HTML: MathML mfrac becomes \\frac, not the concatenation "12"');
    assertExists(mathNodes, n => n.text === 'x^2', 'HTML: MathML msup becomes x^2, not the concatenation "x2"');
    assertExists(mathNodes, n => n.text === '\\sqrt9', 'HTML: MathML msqrt becomes \\sqrt9');
    // An author-written TeX annotation is the source of truth and must be preferred verbatim
    // over anything reconstructed from the presentation tree beside it.
    assertExists(mathNodes, n => n.text === '\\gamma_{0}', 'HTML: TeX annotation wins over the presentation MathML');
    // `display="block"` is MathML's own way of marking a display equation.
    assertExists(mathNodes, n => n.text === 'a_1+b' && (n.metadata as any)?.math === 'block',
        'HTML: MathML display="block" yields a block math node with its subscript intact');
    // The whole point of the conversion: no math node may be a bare run of the digits and letters
    // its markup happened to contain, which is exactly what flattening produced.
    for (const n of mathNodes) {
        assert.ok(!/^[0-9]+$/.test(n.text || ''), `HTML: math node "${n.text}" flattened to bare digits`);
    }

    // ── Tables ────────────────────────────────────────────────────────────────
    const tables = nodes.filter(n => n.type === 'table');
    assert.ok(tables.length >= 1, 'HTML: Has table nodes');
    // Table with data-align="center"
    assertExists(tables, n => (n.metadata as any)?.align === 'center', 'HTML: table align=center');

    const rows = nodes.filter(n => n.type === 'row');
    assert.ok(rows.length >= 3, `HTML: At least 3 rows, got ${rows.length}`);

    const cells = nodes.filter(n => n.type === 'cell');
    assert.ok(cells.length >= 5, `HTML: At least 5 cells, got ${cells.length}`);
    // colspan and rowspan
    assertExists(cells, n => (n.metadata as any)?.colSpan >= 2, 'HTML: cell with colSpan>=2');
    assertExists(cells, n => (n.metadata as any)?.rowSpan >= 2, 'HTML: cell with rowSpan>=2');

    // ── Admonitions (all 5 types) ─────────────────────────────────────────────
    const admonitions = nodes.filter(n => n.type === 'admonition');
    assert.ok(admonitions.length >= 5, `HTML: At least 5 admonitions, got ${admonitions.length}`);
    for (const adType of ['note', 'tip', 'important', 'warning', 'caution'] as const) {
        assertExists(admonitions, n => (n.metadata as any)?.admonitionType === adType, `HTML: admonition type '${adType}'`);
    }

    // ── Image ─────────────────────────────────────────────────────────────────
    const images = nodes.filter(n => n.type === 'image');
    assert.ok(images.length >= 1, 'HTML: Has image nodes');
    const img = assertExists(images, n => (n.metadata as any)?.url?.includes('example.com'), 'HTML: image with url');
    assert.ok((img.metadata as any)?.altText, 'HTML: image has altText');
    assert.ok((img.metadata as any)?.width, 'HTML: image has width');
    assert.ok((img.metadata as any)?.align === 'center', 'HTML: image align=center');

    // ── Embed (YouTube) ───────────────────────────────────────────────────────
    const embeds = nodes.filter(n => n.type === 'embed');
    assert.ok(embeds.length >= 1, 'HTML: Has embed nodes');
    const embed = assertExists(embeds, n => (n.metadata as any)?.embedType === 'youtube', 'HTML: youtube embed');
    assert.ok((embed.metadata as any)?.videoId, 'HTML: embed has videoId');

    // ── Links ─────────────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'external', 'HTML: external link');
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'internal' && (n.metadata as any)?.wikilink !== true, 'HTML: internal anchor link');

    // ── Wikilink ─────────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.wikilink === true, 'HTML: wikilink text node');
    // Attribute-driven shape: page in data-target, display text from the body, or synthesized
    // from data-alias when the anchor is empty. data-wikilink-page keeps precedence over this.
    assertExists(textNodes, n => (n.metadata as any)?.wikilink === true && (n.metadata as any)?.link === 'Target Page' && n.text === 'Alias Text', 'HTML: attribute-driven aliased wikilink');
    assertExists(textNodes, n => (n.metadata as any)?.wikilink === true && (n.metadata as any)?.link === 'Empty Target' && n.text === 'Empty Alias', 'HTML: attribute-driven childless wikilink synthesizes display text from data-alias');

    // ── Citation (attribute-driven span shape) ────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.citationKey === 'doe2021', 'HTML: span.citation with data-key becomes a citation text node');

    // ── Abbreviation ─────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.abbreviationTitle !== undefined, 'HTML: abbreviation text node');

    // ── Footnote ─────────────────────────────────────────────────────────────
    const noteNodes = nodes.filter(n => n.type === 'note');
    assert.ok(noteNodes.length >= 1, 'HTML: Has note nodes');
    assertExists(noteNodes, n => (n.metadata as any)?.noteType === 'footnote', 'HTML: footnote note node');

    // ── Inline-style declaration parsing ─────────────────────────────────────
    // Assert the ABSENCE of the spurious field, not just the presence of the right one: the bug
    // was that `color:` matched inside `background-color:`, so a passing `backgroundColor` check
    // alone would not have caught it.
    const bgOnly = assertExists(textNodes, n => n.formatting?.backgroundColor === 'yellow',
        'HTML: background-color parsed');
    assert.strictEqual(bgOnly.formatting?.color, undefined,
        'HTML: background-color does not leak into color');
    // The wrong-value case: an unanchored regex matched the substring inside background-color and
    // returned red for a run whose text colour is blue.
    assertExists(textNodes, n => n.formatting?.backgroundColor === 'red' && n.formatting?.color === 'blue',
        'HTML: color reads the real declaration, not one nested in another property name');

    // Regression guards - substring matching caught these by accident, exact matching must not lose them.
    assert.ok(textNodes.some(n => n.formatting?.bold === true), 'HTML: font-weight bold variants detected');
    assertExists(textNodes, n => n.formatting?.underline === true, 'HTML: vendor-prefixed text-decoration detected');
    // text-decoration is a shorthand: both keywords have to land, not just the first.
    assertExists(textNodes, n => n.formatting?.underline === true && n.formatting?.strikethrough === true,
        'HTML: combined text-decoration sets both flags');
    // Quote-aware splitting: a semicolon inside a quoted font stack must not split the declaration.
    assertExists(textNodes, n => n.formatting?.font === 'Fira, A;B',
        'HTML: semicolon inside a quoted font-family survives splitting');

    const htmlImages = nodes.filter(n => n.type === 'image');
    // max-width constrains rendering; it is not an author-declared width.
    assertExists(htmlImages, n => (n.metadata as any)?.altText === 'Responsive image' && (n.metadata as any)?.width === undefined,
        'HTML: max-width is not read as an explicit width');
    assertExists(htmlImages, n => (n.metadata as any)?.altText === 'Centered image' && (n.metadata as any)?.align === 'center',
        'HTML: margin shorthand centering is recognised');
    // A 0.5rem left margin is not "left aligned" - the old check substring-matched "margin-left: 0".
    assertExists(htmlImages, n => (n.metadata as any)?.altText === 'Indented image' && (n.metadata as any)?.align === undefined,
        'HTML: a non-zero left margin is not read as left alignment');

    // ── Attribute pass-through (htmlParserConfig.preserveAttributes) ─────────
    // Default OFF is a compatibility guarantee, not a preference: with the flag unset the AST must
    // be byte-identical to previous releases, so assert absence before asserting capture.
    assert.ok(nodes.every(n => n.htmlAttributes === undefined),
        'HTML: htmlAttributes absent by default (no observable AST change)');

    const preserved = await OfficeParser.parseOffice(filePath, { htmlParserConfig: { preserveAttributes: true } } as any);
    const preservedNodes = collectAllNodes(preserved);

    const bagged = assertExists(preservedNodes, n => n.htmlAttributes?.['data-custom'] === 'kept',
        'HTML: preserveAttributes captures an unconsumed data-* attribute');
    assert.strictEqual(bagged.htmlAttributes?.['data-tracking-id'], 'abc123', 'HTML: captures every unconsumed attribute');
    // `class` is deliberately carried (the generator builds its class attribute from style-mapping
    // only, so without this a plain class="lead" is lost outright).
    assert.strictEqual(bagged.htmlAttributes?.['class'], 'lead', 'HTML: class is carried, not dropped');
    // `style` and `id` are consumed into formatting/anchorIds, so they must NOT be duplicated here.
    assert.ok(!('style' in (bagged.htmlAttributes || {})) && !('id' in (bagged.htmlAttributes || {})),
        'HTML: generator-owned attributes (style/id) are not carried');

    // Regression: the attribute-name pattern used to split on any character outside
    // [a-zA-Z0-9-:], so `data_under_score="x"` yielded TWO attributes - `data` plus an invented
    // `under_score`. Assert the invented one is gone; the real name is not a bare-name match so it
    // is filtered rather than carried, which is the safe outcome.
    assert.ok(!preservedNodes.some(n => Object.keys(n.htmlAttributes || {}).some(k => k === 'data' || k.includes('under_score'))),
        'HTML: an underscored attribute name is never split into an invented attribute');

    // ── Roundtrip: generate to HTML ──────────────────────────────────────────
    const result = await OfficeGenerator.generate(ast, 'html');
    const htmlOutput = result.value as string;
    assert.ok(htmlOutput.includes('<h1'), 'HTML roundtrip: h1 tag');
    assert.ok(htmlOutput.includes('<ul') || htmlOutput.includes('<ol'), 'HTML roundtrip: list tag');
    assert.ok(htmlOutput.includes('<table'), 'HTML roundtrip: table tag');
    assert.ok(htmlOutput.includes('<ol') || htmlOutput.includes('<ul'), 'HTML roundtrip: list');

    // Preserved attributes must survive back out, and merging must not produce a duplicate
    // `class` - which is merely invalid in HTML but a *fatal* well-formedness error in the XHTML
    // EpubGenerator emits, i.e. an EPUB that refuses to open.
    const preservedOut = String((await OfficeGenerator.generate(preserved, 'html')).value);
    assert.ok(preservedOut.includes('data-custom="kept"'), 'HTML roundtrip: preserved attribute re-emitted');
    assert.ok(/class="[^"]*\blead\b[^"]*"/.test(preservedOut), 'HTML roundtrip: source class merged into the class attribute');
    for (const tag of preservedOut.match(/<[a-zA-Z][^>]*>/g) || []) {
        const attrNames = [...tag.matchAll(/\s([a-zA-Z_:][\w:.-]*)\s*=/g)].map(m => m[1].toLowerCase());
        assert.strictEqual(new Set(attrNames).size, attrNames.length,
            `HTML roundtrip: no duplicate attribute in ${tag.slice(0, 80)}`);
    }

    console.log('  HTML: All assertions passed ✓');
}

// ─── CSV ─────────────────────────────────────────────────────────────────────

async function testCsv(): Promise<void> {
    console.log('\n=== Running Exhaustive CSV Tests ===');
    const filePath = path.join(__dirname, 'files/exhaustive/csv.csv');
    const ast = await OfficeParser.parseOffice(filePath);
    const nodes = collectAllNodes(ast);

    // ── Sheet node ────────────────────────────────────────────────────────────
    const sheets = nodes.filter(n => n.type === 'sheet');
    assert.ok(sheets.length >= 1, 'CSV: Has sheet node');
    assert.strictEqual((sheets[0].metadata as any)?.sheetName, 'Sheet1', 'CSV: sheet name is Sheet1');

    // ── Comment rows ──────────────────────────────────────────────────────────
    const comments = nodes.filter(n => n.type === 'comment');
    assert.ok(comments.length >= 2, `CSV: At least 2 comment rows, got ${comments.length}`);
    assert.ok(comments.every(c => (c.text || '').startsWith('#')), 'CSV: All comments start with #');

    // ── Rows ──────────────────────────────────────────────────────────────────
    const rows = nodes.filter(n => n.type === 'row');
    // Header row + 5 data rows = 6 rows
    assert.ok(rows.length >= 6, `CSV: At least 6 rows (1 header + 5 data), got ${rows.length}`);

    // ── Cells ─────────────────────────────────────────────────────────────────
    const cells = nodes.filter(n => n.type === 'cell');
    assert.ok(cells.length >= 20, `CSV: At least 20 cells, got ${cells.length}`);

    // Cell with positional metadata
    const cellsWithMeta = cells.filter(n => n.metadata !== undefined);
    assert.ok(cellsWithMeta.length > 0, 'CSV: Cells have metadata (row/col)');
    const firstDataCell = cellsWithMeta.find(n => (n.metadata as any)?.row !== undefined);
    assert.ok(firstDataCell !== undefined, 'CSV: Cell has metadata.row');
    assert.ok(typeof (firstDataCell!.metadata as any)?.col === 'number', 'CSV: Cell has metadata.col');

    // ── Cell with comma inside ────────────────────────────────────────────────
    assertExists(cells, n => (n.text || '').includes(','), 'CSV: cell containing comma');

    // ── Cell with escaped double-quotes ───────────────────────────────────────
    assertExists(cells, n => (n.text || '').includes('"'), 'CSV: cell with escaped double-quotes');

    // ── Cell with newline (multiline) ─────────────────────────────────────────
    assertExists(cells, n => (n.text || '').includes('\n'), 'CSV: multiline cell');

    // ── Roundtrip: generate to CSV ───────────────────────────────────────────
    const result = await OfficeGenerator.generate(ast, 'csv');
    const csvOutput = result.value as string;
    // Comma-containing values should be quoted
    assert.ok(csvOutput.includes('"Value with, a comma"'), 'CSV roundtrip: comma-value quoted');
    // Escaped quotes
    assert.ok(csvOutput.includes('""'), 'CSV roundtrip: escaped double-quotes');
    // Header row preserved
    assert.ok(csvOutput.includes('id'), 'CSV roundtrip: header column "id"');
    assert.ok(csvOutput.includes('name'), 'CSV roundtrip: header column "name"');

    console.log('  CSV: All assertions passed ✓');
}

// ─── RTF ─────────────────────────────────────────────────────────────────────

async function testRtf(): Promise<void> {
    console.log('\n=== Running Exhaustive RTF Tests ===');
    const filePath = path.join(__dirname, 'files/exhaustive/rtf.rtf');
    const ast = await OfficeParser.parseOffice(filePath);
    const nodes = collectAllNodes(ast);

    // ── Paragraphs ────────────────────────────────────────────────────────────
    const paragraphs = nodes.filter(n => n.type === 'paragraph');
    assert.ok(paragraphs.length > 0, `RTF: Has paragraphs, got ${paragraphs.length}`);

    // ── Text nodes ────────────────────────────────────────────────────────────
    const textNodes = nodes.filter(n => n.type === 'text');
    assert.ok(textNodes.length > 0, 'RTF: Has text nodes');

    // ── Formatting flags (bold/italic/underline) ──────────────────────────────
    // The RTF test file should have some formatted text
    const boldNodes = textNodes.filter(n => n.formatting?.bold === true);
    const italicNodes = textNodes.filter(n => n.formatting?.italic === true);
    const underlineNodes = textNodes.filter(n => n.formatting?.underline === true);
    // At least one should be present (the test.rtf is a large file with formatting)
    assert.ok(
        boldNodes.length > 0 || italicNodes.length > 0 || underlineNodes.length > 0,
        'RTF: Has at least one formatted text node (bold/italic/underline)'
    );

    // ── Roundtrip: generate to RTF ────────────────────────────────────────────
    const result = await OfficeGenerator.generate(ast, 'rtf');
    const rtfOutput = result.value as string;
    assert.ok(rtfOutput.includes('{\\rtf1'), 'RTF roundtrip: output starts with {\\rtf1');
    assert.ok(rtfOutput.includes('\\par'), 'RTF roundtrip: has \\par paragraph marker');

    console.log('  RTF: All assertions passed ✓');
}

/**
 * The full interop loop: externally-authored editor HTML -> HtmlParser -> AST ->
 * MarkdownGenerator -> .md -> MarkdownParser -> AST -> HtmlGenerator -> HTML. Proves every rich
 * construct survives all four hops, that the `sourceAttributes` emission re-expresses each one as
 * a data-* attribute the parser reads back, and that the default (flag off) emission is unchanged.
 */
async function testAttributeRoundtrip(): Promise<void> {
    console.log('\n=== Running Attribute-Driven Round-Trip Tests ===');

    const editorHtml = [
        '<p><a data-wikilink="true" data-target="Target Page" data-alias="Alias Text">Alias Text</a></p>',
        '<p><a data-wikilink="true" data-target="Bare Page">Bare Page</a></p>',
        '<p><span class="citation cursor-help text-emerald-600" data-key="doe2021" data-label="Doe 2021" title="Doe, J. (2021)">[Doe 2021]</span></p>',
        '<span data-math="E=mc^2" class="math-inline">E=mc^2</span>',
        '<div data-math="a^2+b^2=c^2" class="math-block">a^2+b^2=c^2</div>',
        // Multi-line, as real diagrams are: single-line code round-trips as inline code (no fence),
        // which is the generator's content-based block/inline rule for all code, not mermaid-specific.
        '<div class="mermaid" data-mermaid="graph TD;\n    A--&gt;B;">graph TD;\n    A--&gt;B;</div>',
    ].join('\n');

    // Hop 1: editor HTML -> AST (widened parser).
    const ast1 = await OfficeParser.parseOffice(Buffer.from(editorHtml), { fileType: 'html' });
    const n1 = collectAllNodes(ast1);
    assertExists(n1, n => (n.metadata as any)?.wikilink === true && (n.metadata as any)?.link === 'Target Page' && n.text === 'Alias Text', 'RT hop1: aliased wikilink');
    assertExists(n1, n => (n.metadata as any)?.citationKey === 'doe2021', 'RT hop1: citation');
    assertExists(n1, n => (n.metadata as any)?.math === 'inline' && n.text === 'E=mc^2', 'RT hop1: inline math from data-math');
    assertExists(n1, n => (n.metadata as any)?.math === 'block' && n.text === 'a^2+b^2=c^2', 'RT hop1: block math from data-math');
    assertExists(n1, n => (n.metadata as any)?.language === 'mermaid' && (n.text || '').includes('graph TD') && (n.text || '').includes('A-->B'), 'RT hop1: mermaid');

    // Hop 2: AST -> Markdown (defaults).
    const md = String((await OfficeGenerator.generate(ast1, 'md')).value);
    assert.ok(md.includes('[[Target Page|Alias Text]]'), 'RT hop2: aliased wikilink -> [[page|alias]]');
    assert.ok(md.includes('[@doe2021]'), 'RT hop2: citation -> [@key]');
    assert.ok(md.includes('$E=mc^2$'), 'RT hop2: inline math -> $...$');
    assert.ok(md.includes('a^2+b^2=c^2'), 'RT hop2: block math content preserved');
    assert.ok(/```mermaid[\s\S]*graph TD/.test(md), 'RT hop2: mermaid -> ```mermaid fence');

    // Hop 3: Markdown -> AST.
    const ast2 = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });

    // Hop 4a: AST -> HTML with sourceAttributes ON - every data-* survives.
    const htmlOn = String((await OfficeGenerator.generate(ast2, 'html', { htmlConfig: { sourceAttributes: true, standalone: false } })).value);
    assert.ok(htmlOn.includes('data-wikilink="true"') && htmlOn.includes('data-target="Target Page"') && htmlOn.includes('data-alias="Alias Text"'), 'RT hop4 (on): wikilink data-* survive');
    assert.ok(htmlOn.includes('class="citation"') && htmlOn.includes('data-key="doe2021"'), 'RT hop4 (on): citation span with data-key');
    assert.ok(htmlOn.includes('data-math="E=mc^2"'), 'RT hop4 (on): LaTeX in data-math, undelimited');
    assert.ok(htmlOn.includes('class="mermaid"') && htmlOn.includes('data-mermaid="graph TD;') && htmlOn.includes('A--&gt;B'), 'RT hop4 (on): mermaid div with data-mermaid');

    // Hop 4b: same AST -> HTML with sourceAttributes OFF (default) - legacy shapes, locking defaults.
    const htmlOff = String((await OfficeGenerator.generate(ast2, 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(htmlOff.includes('<cite data-citation-key="doe2021"'), 'RT hop4 (off): citation is <cite>');
    assert.ok(htmlOff.includes('data-math="inline"') && htmlOff.includes('$E=mc^2$'), 'RT hop4 (off): math is delimited data-math="inline"');
    assert.ok(htmlOff.includes('class="language-mermaid"'), 'RT hop4 (off): mermaid is <code class="language-mermaid">');
    assert.ok(!htmlOff.includes('data-wikilink="true"'), 'RT hop4 (off): no attribute-driven wikilink attrs by default');

    console.log('  Attribute round-trip: All assertions passed ✓');
}

/**
 * parseOffice/convert accept a web Blob/File (or any object with `arrayBuffer()`), so browser
 * callers don't have to convert first. A filename, when present, drives extension-based type
 * detection; a nameless blob still resolves through magic-byte sniffing.
 */
async function testBlobInput(): Promise<void> {
    console.log('\n=== Running Blob/File Input Tests ===');
    const filePath = path.join(__dirname, 'files/exhaustive/html.html');
    const pathAst = await OfficeParser.parseOffice(filePath);
    const pathText = collectAllNodes(pathAst).filter(n => n.type === 'text').length;
    const bytes = fs.readFileSync(filePath);

    // Web Blob (global since Node 18), parsed with an explicit fileType.
    if (typeof Blob !== 'undefined') {
        const blob = new Blob([bytes]);
        const blobAst = await OfficeParser.parseOffice(blob as any, { fileType: 'html' });
        assert.strictEqual(collectAllNodes(blobAst).filter(n => n.type === 'text').length, pathText, 'Blob: text-node count matches the path parse');
    } else {
        console.log('  (global Blob unavailable, skipping the Blob case)');
    }

    // A structural BlobLike carrying a filename: the extension drives type detection (no fileType).
    const fileLike = { arrayBuffer: async () => new Uint8Array(bytes).buffer, name: 'document.html' };
    const fileAst = await OfficeParser.parseOffice(fileLike as any);
    assert.strictEqual(fileAst.type, 'html', 'BlobLike: filename extension drives type detection');
    assert.strictEqual(collectAllNodes(fileAst).filter(n => n.type === 'text').length, pathText, 'BlobLike: text-node count matches the path parse');

    console.log('  Blob/File input: All assertions passed ✓');
}

/**
 * Issue #109: paragraph-mark run properties (`<w:pPr><w:rPr>`) format only the paragraph mark
 * per OOXML ISO 29500 §17.3.1.29 - they must not bleed onto the paragraph's text runs. A DOCX is
 * built in-memory: paragraph 1's mark is bold+italic; paragraph 2 uses a bold paragraph *style*
 * to prove real style inheritance still reaches runs.
 */
async function testWordParagraphMarkFormatting(): Promise<void> {
    console.log('\n=== Running Word Paragraph-Mark Formatting Tests (issue #109) ===');

    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:rPr><w:b/><w:i/></w:rPr></w:pPr>
      <w:r><w:t>PlainRun</w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>BoldRun</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Strong1"/></w:pPr>
      <w:r><w:t>StyledRun</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Strong1"><w:rPr><w:b/></w:rPr></w:style>
</w:styles>`;
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

    const zip = zipSync({
        '[Content_Types].xml': strToU8(contentTypes),
        '_rels/.rels': strToU8(rels),
        'word/document.xml': strToU8(documentXml),
        'word/styles.xml': strToU8(stylesXml),
    });

    const ast = await OfficeParser.parseOffice(Buffer.from(zip), { fileType: 'docx' });
    const textNodes = collectAllNodes(ast).filter(n => n.type === 'text');
    const byText = (t: string) => textNodes.find(n => (n.text || '').includes(t));

    const plain = byText('PlainRun');
    assert.ok(plain, 'Word #109: PlainRun text node exists');
    assert.ok(!plain!.formatting?.bold, 'Word #109: a run with no rPr must NOT inherit the paragraph mark\'s bold');
    assert.ok(!plain!.formatting?.italic, 'Word #109: a run with no rPr must NOT inherit the paragraph mark\'s italic');

    const boldRun = byText('BoldRun');
    assert.ok(boldRun, 'Word #109: BoldRun text node exists');
    assert.ok(boldRun!.formatting?.bold === true, 'Word #109: a run with its own <w:b/> is still bold');
    assert.ok(!boldRun!.formatting?.italic, 'Word #109: the paragraph mark\'s italic does not bleed onto an explicitly-bold run');

    const styled = byText('StyledRun');
    assert.ok(styled, 'Word #109: StyledRun text node exists');
    assert.ok(styled!.formatting?.bold === true, 'Word #109: paragraph-style formatting still reaches its runs (style chain intact)');

    console.log('  Word paragraph-mark formatting: All assertions passed ✓');
}

/**
 * Excel stores a date or a time as a day count and shows it through the cell's number format, and
 * a boolean as 1 or 0. An XLSX is built in-memory with built-in and custom date/time formats, number
 * formats that only look like dates, and both date systems; every cell must read as the sheet shows it.
 */
async function testExcelDates(): Promise<void> {
    console.log('\n=== Running Excel Date/Time/Boolean Cell Tests ===');

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="12">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd h:mm:ss"/>
    <numFmt numFmtId="166" formatCode="0.00"/>
    <numFmt numFmtId="167" formatCode="0.0 &quot;days&quot;"/>
    <numFmt numFmtId="168" formatCode="[$-x-systime]h:mm:ss AM/PM"/>
    <numFmt numFmtId="169" formatCode="DD.MM.YYYY"/>
    <numFmt numFmtId="170" formatCode="mmm yyyy"/>
    <numFmt numFmtId="171" formatCode="d-mmm"/>
    <numFmt numFmtId="172" formatCode="mm:ss"/>
    <numFmt numFmtId="173" formatCode="h:mm"/>
    <numFmt numFmtId="174" formatCode="0.0\\ \\h"/>
    <numFmt numFmtId="175" formatCode="[h]:mm"/>
  </numFmts>
  <cellXfs count="17">
    <xf numFmtId="0"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="20"/><xf numFmtId="46"/>
    <xf numFmtId="14"/><xf numFmtId="22"/><xf numFmtId="166"/><xf numFmtId="167"/><xf numFmtId="168"/>
    <xf numFmtId="169"/><xf numFmtId="170"/><xf numFmtId="171"/><xf numFmtId="172"/><xf numFmtId="173"/>
    <xf numFmtId="174"/><xf numFmtId="175"/>
  </cellXfs>
</styleSheet>`;
    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" s="1"><v>45565</v></c>
      <c r="B1" s="2"><v>45565.586804398</v></c>
      <c r="C1" s="3"><v>0.25</v></c>
      <c r="D1" s="4"><v>1.5</v></c>
      <c r="E1" s="5"><v>45292</v></c>
      <c r="F1" s="6"><v>45292.395833333336</v></c>
      <c r="G1" s="10"><v>45565</v></c>
      <c r="H1" s="1"><f>A1+30</f><v>45595</v></c>
      <c r="I1" s="11"><v>45566</v></c>
      <c r="J1" s="12"><v>45565</v></c>
      <c r="K1" s="13"><v>0.0625</v></c>
      <c r="L1" s="14"><v>0.75</v></c>
      <c r="M1" s="16"><v>1.25</v></c>
    </row>
    <row r="2">
      <c r="A2" s="7"><v>12.5</v></c>
      <c r="B2" s="8"><v>3.5</v></c>
      <c r="C2" s="9"><v>0.5</v></c>
      <c r="D2" s="1"><v>59</v></c>
      <c r="E2" s="1"><v>61</v></c>
      <c r="F2" s="1"><v>3000000</v></c>
      <c r="G2" s="1" t="inlineStr"><is><t>not a date</t></is></c>
      <c r="H2" t="b"><v>1</v></c>
      <c r="I2" t="b"><v>0</v></c>
      <c r="J2"><v>45565</v></c>
      <c r="K2" s="15"><v>7.5</v></c>
      <c r="L2" s="1"><v>-1</v></c>
      <c r="M2" s="1"><v></v></c>
      <c r="N2" s="1" t="str"><f>TEXT(A1,"yyyy")</f><v>2024</v></c>
    </row>
  </sheetData>
</worksheet>`;
    const workbook = (date1904?: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${date1904 ? `<workbookPr date1904="${date1904}"/>` : '<workbookPr/>'}
  <sheets><sheet name="Dates" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
    const parseCells = async (date1904?: string): Promise<string[]> => {
        const zip = zipSync({
            '[Content_Types].xml': strToU8(contentTypes),
            '_rels/.rels': strToU8(rels),
            'xl/workbook.xml': strToU8(workbook(date1904)),
            'xl/_rels/workbook.xml.rels': strToU8(workbookRels),
            'xl/styles.xml': strToU8(stylesXml),
            'xl/worksheets/sheet1.xml': strToU8(sheetXml),
        });
        const ast = await OfficeParser.parseOffice(Buffer.from(zip), { fileType: 'xlsx' });
        return collectAllNodes(ast).filter(n => n.type === 'cell').map(n => n.text);
    };

    assert.deepStrictEqual(await parseCells(), [
        // Date, rounded date-time, built-in time and elapsed duration, built-in date and date-time,
        // an upper-case format, a formula result in a date style, a year without a day, a day
        // without a year, minutes and seconds, hours and minutes, and a custom elapsed duration
        '2024-09-30', '2024-09-30 14:05:00', '06:00:00', '36:00:00', '2024-01-01', '2024-01-01 09:30:00',
        '2024-09-30', '2024-10-30', '2024-10-01', '2024-09-30', '01:30:00', '18:00:00', '30:00:00',
        // Number formats (one with "days" in quoted text), a time whose [$-x-systime] section contains
        // a "y", the 1900 system around its fictitious 1900-02-29, a serial past 9999-12-31, text in a
        // date style, booleans, a number with no style, a number with an escaped unit, a negative
        // serial, and a text formula result in a date style. The empty value makes no cell.
        '12.5', '3.5', '12:00:00', '1900-02-28', '1900-03-01', '3000000', 'not a date', 'TRUE', 'FALSE',
        '45565', '7.5', '-1', '2024',
    ], 'Excel: cells read as the sheet shows them');

    for (const date1904 of ['1', 'true']) {
        const cells = await parseCells(date1904);
        assert.strictEqual(cells[0], '2028-10-01', `Excel: date1904="${date1904}" counts from 1904-01-01`);
        assert.strictEqual(cells[2], '06:00:00', 'Excel: a time of day does not depend on the date system');
    }

    console.log('  Excel date/time/boolean cells: All assertions passed ✓');
}

/**
 * Round 3: generated-output assertions. The suite historically asserted parse results but almost
 * never what the generators emit, which is how the frontmatter and footnote-markup bugs shipped
 * unnoticed. Covers chunking text retention (3.A), empty-frontmatter round trip (3.B), footnote
 * definition markup (3.C), opt-in inline formatting through `.md` (3.D), and generated
 * dl/dt/dd/abbr/taskList/frontmatter (3.F).
 */
async function testGeneratedOutput(): Promise<void> {
    console.log('\n=== Running Generated-Output Tests (round 3) ===');
    const parseHtml = (s: string) => OfficeParser.parseOffice(Buffer.from(s), { fileType: 'html' });
    const parseMd = (s: string) => OfficeParser.parseOffice(Buffer.from(s), { fileType: 'md' });
    const strip = (s: string) => (s || '').replace(/\s/g, '');

    // 3.A: chunking retains text from HTML- and MD-origin ASTs (their paragraphs are children-only).
    for (const origin of ['html', 'md'] as const) {
        const ast = origin === 'html'
            ? await parseHtml('<h1>Chapter</h1><p>First paragraph.</p><p>Second one here.</p>')
            : await parseMd('# Chapter\n\nFirst paragraph.\n\nSecond one here.');
        const chunks = (await OfficeGenerator.generate(ast, 'chunks')).value as any[];
        assert.ok(chunks.length > 0, `chunking (${origin}): produces chunks, not []`);
        const chunkChars = strip(chunks.map(c => c.text).join(' ')).length;
        const plainChars = strip(((await ast.to('text')).value as string) || '').length;
        assert.ok(chunkChars >= plainChars * 0.9, `chunking (${origin}): retains >=90% of .to('text') chars (${chunkChars}/${plainChars})`);
    }

    // 3.B: empty metadata emits no frontmatter fence, and doesn't reparse into a `## ---` heading.
    const emptyMd = String((await OfficeGenerator.generate(await parseHtml('<p>Body only, no head.</p>'), 'md')).value);
    assert.ok(!emptyMd.startsWith('---'), '3.B: empty metadata emits no frontmatter fence');
    assert.ok(!collectAllNodes(await parseMd(emptyMd)).some(n => n.type === 'heading' && (n.text || '').includes('---')), '3.B: no bogus "## ---" heading on reparse');
    // Also parse the raw broken shape directly: it must not become a heading.
    assert.ok(!collectAllNodes(await parseMd('---\n---\n\nJust body.')).some(n => n.type === 'heading'), '3.B: a raw empty `---\\n---` block is not misread as a heading');

    // 3.F: a metadata-bearing AST still emits a real frontmatter block.
    assert.ok(/^---\ntitle: /.test(String((await OfficeGenerator.generate(await parseMd('---\ntitle: T\n---\n\nBody'), 'md')).value)), '3.F: metadata emits a frontmatter block with title');

    // 3.C: footnote definition markup is <div data-footnote-id>, not a <p> wrapping block content.
    const fnHtml = String((await OfficeGenerator.generate(await parseMd('Text[^1].\n\n[^1]: A footnote.'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<div[^>]*data-footnote-id/.test(fnHtml), '3.C: footnote definition is a <div data-footnote-id>');
    assert.ok(!/<p[^>]*data-footnote-id/.test(fnHtml), '3.C: footnote definition is not a <p>');

    // 3.F: generated dl/dt/dd/abbr/taskList from the exhaustive markdown fixture.
    const genHtml = String((await OfficeGenerator.generate(await OfficeParser.parseOffice(path.join(__dirname, 'files/exhaustive/markdown.md')), 'html', { htmlConfig: { standalone: false } })).value);
    for (const frag of ['<dl>', '<dt>', '<dd>', '<abbr', 'data-type="taskList"']) {
        assert.ok(genHtml.includes(frag), `3.F: generated HTML contains ${frag}`);
    }

    // 3.D: inline color/highlight/size survive `.md` only when opted in; default output has no span.
    const colored = await parseHtml('<p>plain <span style="color:#cc0000">red</span> <mark>hi</mark></p>');
    assert.ok(!String((await OfficeGenerator.generate(colored, 'md')).value).includes('<span style'), '3.D: default .md output has no inline-formatting span');
    const withSpan = String((await OfficeGenerator.generate(colored, 'md', { mdConfig: { fallbackToHtml: { inlineFormatting: true } } } as any)).value);
    assert.ok(withSpan.includes('<span style="color:'), '3.D: inlineFormatting emits a styled span');
    const reColored = collectAllNodes(await parseMd(withSpan)).filter(n => n.type === 'text');
    assert.ok(reColored.some(t => (t.formatting?.color || '').toLowerCase().includes('cc0000')), '3.D: text color survives the .md round trip');
    assert.ok(reColored.some(t => !!t.formatting?.backgroundColor), '3.D/mark: <mark> highlight survives into the AST');
    assert.ok(collectAllNodes(await parseHtml('<mark data-color="#00ff00">x</mark>')).some(n => n.type === 'text' && (n.formatting?.backgroundColor || '') === '#00ff00'), 'mark: data-color drives the highlight color');

    // --- Review-round fixes (regression guards) ---

    // Fix: the tokenizer must not truncate a tag at a literal `>` inside an attribute value; a real
    // editor serializes a mermaid diagram's `-->` unescaped in data-mermaid.
    const merReal = collectAllNodes(await parseHtml('<div class="mermaid" data-mermaid="graph TD; A-->B; B-->C;">graph TD; A-->B; B-->C;</div>')).find(n => n.type === 'code' && (n.metadata as any)?.language === 'mermaid');
    assert.ok(merReal && merReal.text === 'graph TD; A-->B; B-->C;', 'fix: literal > inside an attribute value does not truncate the tag');

    // Fix: the quote-aware tag scan must not, on a stray unescaped `<` in prose followed by an
    // unbalanced quote (an apostrophe is enough), swallow the rest of the document into one text
    // node. It falls back to the next literal `>`, degrading like the pre-widening naive scan.
    const strayLt = collectAllNodes(await parseHtml("<p>score a < b's weight > c <strong>bold</strong> end</p>"));
    assert.ok(strayLt.some(n => n.type === 'text' && n.text === 'bold' && n.formatting?.bold), 'fix: an unbalanced quote after a stray < does not swallow following elements');
    assert.ok(!strayLt.some(n => (n.text || '').includes('</strong>')), 'fix: literal markup does not leak into text when a stray < has an unbalanced quote');

    // Fix: `<pre><code>` decodes entities (mermaid arrows, and `<`/`>`/`&` in code snippets).
    const preCode = collectAllNodes(await parseHtml('<pre><code class="language-js">a &lt; b &amp;&amp; c &gt; d</code></pre>')).find(n => n.type === 'code');
    assert.ok(preCode && preCode.text === 'a < b && c > d', 'fix: <pre><code> entities are decoded');

    // Fix: `<div data-math="latex">$x$</div>` reads as inline (delimiter over div-tag) with the
    // `$` delimiters stripped, not block with the delimiters retained.
    const dm = collectAllNodes(await parseHtml('<div data-math="whatever">$x+y$</div>')).find(n => n.type === 'code' && (n.metadata as any)?.math);
    assert.ok(dm && (dm.metadata as any).math === 'inline' && dm.text === 'x+y', 'fix: $-delimited data-math div is inline with delimiters stripped');

    // Fix: chunking must not merge words across block-level children of one node.
    const merged = ((await OfficeGenerator.generate(await parseHtml('<ul><li><p>First para</p><p>Second para</p></li></ul>'), 'chunks')).value as any[]).map(c => c.text).join(' ');
    assert.ok(!merged.includes('paraSecond') && merged.includes('First para'), 'fix: chunking does not merge words across block children');

    // ── Round 4 (release-blocker + editor/RAG gaps) ──────────────────────────
    const mdCycle = async (md: string) => String((await OfficeGenerator.generate(await parseMd(md), 'md')).value);

    // 4.A: a footnote-bearing .md is byte-stable after cycle 1 - no `### Notes` heading accumulates,
    // and the reference marker stays before the period (aligned with the HTML generator).
    const fnC1 = await mdCycle('Body[^1].\n\n[^1]: Def body.');
    const fnC2 = await mdCycle(fnC1);
    const fnC3 = await mdCycle(fnC2);
    assert.strictEqual(fnC2, fnC1, '4.A: footnote .md is byte-stable after cycle 1');
    assert.strictEqual(fnC3, fnC2, '4.A: footnote .md stays byte-stable across further cycles');
    assert.ok(!/###\s*Notes/.test(fnC1), '4.A: no "### Notes" heading is emitted before the definitions');
    assert.ok(/Body\[\^1\]\./.test(fnC1), '4.A: footnote marker stays before the period, aligned with HTML');

    // 4.F.5 / 6.E.4: cycle-stability across construct types (the shape that catches the 3.B/4.A
    // class). 6.E.4 broadened the sweep to headings/lists/code/links/images/abbr/frontmatter, the
    // 6.A thematic break, and a legacy `### Notes` document - whose `---` used to evaporate and force
    // a cycle-2 settle, and which the 6.A fix makes stable from cycle 1.
    for (const [label, seed] of [
        ['task list', '- [x] done\n- [ ] todo'],
        ['admonition', '> [!NOTE]\n> heads up'],
        ['table', '| a | b |\n| --- | --- |\n| 1 | 2 |'],
        ['definition list', 'Term\n: Definition'],
        ['thematic break', 'Above.\n\n---\n\nBelow.'],
        ['heading', '# Title\n\nBody paragraph.'],
        ['unordered list', '- one\n- two\n- three'],
        ['ordered list', '1. one\n2. two'],
        ['nested list', '- a\n    - b'],
        ['fenced code', '```js\nconst x = 1;\nconst y = 2;\n```'],
        ['link', '[text](https://example.com)'],
        ['image', '![alt](https://example.com/i.png)'],
        ['abbreviation', 'The HTML spec.\n\n*[HTML]: HyperText Markup Language'],
        ['frontmatter', '---\ntitle: Doc\n---\n\nBody.'],
        ['legacy ### Notes', '## Heading\n\nBody[^1].\n\n---\n\n### Notes\n\n[^1]: note body'],
    ] as const) {
        const a = await mdCycle(seed);
        const b = await mdCycle(a);
        assert.strictEqual(b, a, `4.F.5/6.E.4: ${label} .md is cycle-stable after cycle 1`);
    }

    // 4.B: a highlight emits <mark> (Tiptap's Highlight extension parseHTML matches exactly `mark`),
    // not a <span style="background-color">, and the generated <mark> re-parses as a highlight.
    const hlHtml = String((await OfficeGenerator.generate(await parseHtml('<p><span style="background-color:#ffff00">hi</span></p>'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<mark[^>]*background-color/.test(hlHtml), '4.B: highlight emits <mark> carrying the background-color');
    assert.ok(!/<span[^>]*background-color/.test(hlHtml), '4.B: highlight is not a background-color <span>');
    assert.ok(collectAllNodes(await parseHtml(hlHtml)).some(n => !!n.formatting?.backgroundColor), '4.B: generated <mark> round-trips back to a highlight');

    // 4.C: a footnote body is searchable in RAG chunks (folded into the referencing node's text).
    const fnChunks = ((await OfficeGenerator.generate(await parseMd('Para[^1].\n\n[^1]: Searchable footnote body.'), 'chunks')).value as any[]).map(c => c.text).join('  ');
    assert.ok(/Searchable footnote body/.test(fnChunks), '4.C: footnote body reaches the RAG chunks');

    // 4.D: a quoted frontmatter scalar stays a string across a save/reload cycle; unquoted coerces.
    const fmMd = String((await OfficeGenerator.generate(await parseMd('---\nversion: "123"\ncount: 5\n---\n\nBody'), 'md')).value);
    assert.ok(/version:\s*"123"/.test(fmMd), '4.D: a quoted "123" stays a quoted string across the cycle');
    assert.ok(/^count:\s*5\s*$/m.test(fmMd), '4.D: an unquoted 5 stays an unquoted number');

    // 4.E: an orphan footnote definition (defined, never referenced) is preserved, not dropped.
    const orphanMd = String((await OfficeGenerator.generate(await parseMd('Body text.\n\n[^x]: Orphan definition.'), 'md')).value);
    assert.ok(/\[\^x\]:\s*Orphan definition/.test(orphanMd), '4.E: an orphan footnote definition is preserved');

    // 4.F.1: generated footnote reference (sup[data-footnote-ref]) and container (section[data-footnotes]).
    const refHtml = String((await OfficeGenerator.generate(await parseMd('Cite[^1].\n\n[^1]: Note body.'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<sup[^>]*data-footnote-ref/.test(refHtml), '4.F.1: footnote reference is a sup[data-footnote-ref]');
    assert.ok(/<section[^>]*data-footnotes/.test(refHtml), '4.F.1: footnotes live in a section[data-footnotes]');

    // 4.F.2: generated task items carry li[data-checked] in both checked states.
    const taskHtml = String((await OfficeGenerator.generate(await parseMd('- [x] done\n- [ ] todo'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<li[^>]*data-checked="true"/.test(taskHtml), '4.F.2: checked task item is li[data-checked="true"]');
    assert.ok(/<li[^>]*data-checked="false"/.test(taskHtml), '4.F.2: unchecked task item is li[data-checked="false"]');

    // 4.F.3: generated footnote HTML fed back through HtmlParser (export-side round trip) keeps the note.
    assert.ok(
        collectAllNodes(await parseHtml(refHtml)).some(n => (n.metadata as any)?.noteType === 'footnote' || (n.notes || []).some(nt => (nt.metadata as any)?.noteType === 'footnote')),
        '4.F.3: generated footnote HTML re-parses into a footnote note',
    );

    // ── Round 5 (residual-gap fixes) ─────────────────────────────────────────
    // 5.B: an orphan footnote definition survives markdownwriter's editor LOAD path
    // (md -> HTML -> md), landing inside section[data-footnotes] with no dangling back-link,
    // instead of rendering outside the section and vanishing on the return trip.
    const orphanHtml = String((await OfficeGenerator.generate(await parseMd('Body text.\n\n[^x]: Orphan definition.'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<section[^>]*data-footnotes[^>]*>[\s\S]*data-footnote-id="x"[\s\S]*<\/section>/.test(orphanHtml), '5.B: orphan definition renders inside section[data-footnotes]');
    assert.ok(!/data-footnote-id="x">[\s\S]*?#footnote-ref-x/.test(orphanHtml), '5.B: orphan definition has no dangling back-link');
    assert.ok(collectAllNodes(await parseHtml(orphanHtml)).some(n => (n.metadata as any)?.noteType === 'footnote' && (n.metadata as any)?.unreferenced), '5.B: HtmlParser recovers the orphan definition as an unreferenced note');
    const orphanRT = String((await OfficeGenerator.generate(await parseHtml(orphanHtml), 'md')).value);
    assert.ok(/\[\^x\]:\s*Orphan definition/.test(orphanRT), '5.B: orphan definition survives md -> HTML -> md');
    assert.ok(!orphanRT.includes('↩'), '5.B: no stray return-arrow leaks into the round-tripped .md');

    // 5.C: an office-origin (DOCX) footnote/endnote body reaches the chunks. DOCX/ODT/RTF set
    // `.text` on the paragraph and hang the note off a nested child, which the `.text` fast-path in
    // collectNodeText used to skip - so the body was in `.to('text')` but absent from every chunk.
    const docxAst = await OfficeParser.parseOffice(path.join(__dirname, 'files/test.docx'));
    const docxChunks = ((await OfficeGenerator.generate(docxAst, 'chunks')).value as any[]).map(c => c.text).join('\n');
    assert.ok(/clickable endnotes/i.test(docxChunks), '5.C: a DOCX footnote/endnote body appears in the chunks');

    // ── Round 6 (pre-existing bugs surfaced by the round-5 sweep) ─────────────
    // 6.A: a thematic break (`---` / `<hr>`) is no longer lost on save. It stays `---` through a
    // Markdown save, emits a plain `<hr>` in HTML, and an HTML `<hr>` comes back as `---`; an office
    // page break (`<hr class="page-break">`) is kept distinct from a thematic one.
    const tbMd = String((await OfficeGenerator.generate(await parseMd('Above.\n\n---\n\nBelow.'), 'md')).value);
    assert.ok(/Above\.\n\n---\n\nBelow\./.test(tbMd), '6.A: a Markdown thematic break survives a save');
    const tbHtml = String((await OfficeGenerator.generate(await parseMd('A\n\n---\n\nB'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(tbHtml.includes('<hr>') && !/hr class="page-break"/.test(tbHtml), '6.A: a thematic break emits a plain <hr>, not a page-break hr');
    const hrBack = String((await OfficeGenerator.generate(await parseHtml('<p>A</p><hr><p>B</p>'), 'md')).value);
    assert.ok(/A\n\n---\n\nB/.test(hrBack), '6.A: an HTML <hr> round-trips to a Markdown ---');
    assert.ok(collectAllNodes(await parseHtml('<hr>')).some(n => n.type === 'break' && (n.metadata as any)?.breakType === 'thematic'), '6.A: <hr> parses to a thematic break');
    assert.ok(collectAllNodes(await parseHtml('<hr class="page-break">')).some(n => n.type === 'break' && (n.metadata as any)?.breakType === 'page'), '6.A: <hr class="page-break"> stays a page break');

    // 6.B: a footnote referenced inside a table cell is defined exactly once. Header-row detection
    // (and sparse-column rows) re-process cells after `childrenOutput` already did, which used to
    // push the referenced note into the collected footnotes twice.
    const cellFnHtml = String((await OfficeGenerator.generate(await parseMd('| **H** | K |\n| --- | --- |\n| x[^1] | y |\n\n[^1]: cell note'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.strictEqual((cellFnHtml.match(/id="footnote-1"/g) || []).length, 1, '6.B: a table-cell footnote is defined once, not twice');
    assert.strictEqual((cellFnHtml.match(/href="#footnote-ref-1"/g) || []).length, 1, '6.B: a table-cell footnote emits a single back-link');

    // 6.C: a note nested as a CHILD of a container (a consumer-built shape; no shipped parser emits
    // it) is not hoisted into section[data-footnotes] with a back-link whose citation anchor does
    // not exist. The hoist is now gated on the `unreferenced` flag, matching MarkdownGenerator
    // (which only hoists at the top level), so the two generators agree at depth.
    const nestedAst: any = await parseHtml('<p>Parent text</p>');
    const paraNode = nestedAst.content.find((n: any) => n.children && n.children.length) || nestedAst.content[0];
    paraNode.children.push({ type: 'note', metadata: { noteType: 'footnote', noteId: '9' }, children: [{ type: 'text', text: 'nested note body' }] });
    const nestedHtml = String((await OfficeGenerator.generate(nestedAst, 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(!/<section[^>]*data-footnotes/.test(nestedHtml), '6.C: a depth-nested non-orphan note is not hoisted into a footnotes section');

    // 6.D: two references to the same footnote id are one shared footnote. They both stay `[^1]`
    // with a single definition, instead of renumbering to `[^1]`/`[^2]` with a duplicated body.
    const dupRefMd = String((await OfficeGenerator.generate(await parseMd('See[^1] and again[^1].\n\n[^1]: shared note.'), 'md')).value);
    assert.ok(/See\[\^1\] and again\[\^1\]\./.test(dupRefMd), '6.D: repeated references to one id both stay [^1]');
    assert.strictEqual((dupRefMd.match(/^\[\^1\]:/gm) || []).length, 1, '6.D: a shared footnote is defined exactly once');
    const dupRefHtml = String((await OfficeGenerator.generate(await parseMd('See[^1] and again[^1].\n\n[^1]: shared note.'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.strictEqual((dupRefHtml.match(/id="footnote-1"/g) || []).length, 1, '6.D: a shared footnote renders one HTML definition');

    // 6.E.2: office-origin footnote/endnote bodies reach the chunks for ODT and RTF too (5.C pinned
    // only DOCX). All three fixtures carry the same footnote body.
    for (const f of ['files/test.odt', 'files/test.rtf'] as const) {
        const officeAst = await OfficeParser.parseOffice(path.join(__dirname, f));
        const officeChunks = ((await OfficeGenerator.generate(officeAst, 'chunks')).value as any[]).map(c => c.text).join('\n');
        assert.ok(/In paged media, footnotes/i.test(officeChunks), `6.E.2: a ${f} footnote body reaches the chunks`);
    }

    // 7.E: a raw inline <br> round-trips symmetrically. MarkdownGenerator emits a raw <br> for a
    // line break in a table cell (a GFM pipe cell cannot hold a newline); MarkdownParser must read
    // it back as a break node rather than escaping it to literal `&lt;br&gt;` text and destroying it.
    for (const brForm of ['a<br>b', 'a<br/>b', 'a<br />b']) {
        const h = String((await OfficeGenerator.generate(await parseMd(brForm), 'html', { htmlConfig: { standalone: false } })).value).replace(/\n/g, '');
        assert.ok(/a<br\s*\/?>b/.test(h) && !/&lt;br/.test(h), `7.E: raw inline ${brForm} becomes a real <br>, not escaped text`);
    }
    const cellBrHtml = String((await OfficeGenerator.generate(await parseMd('| a<br>b | c |\n| --- | --- |\n| x | y |'), 'html', { htmlConfig: { standalone: false } })).value);
    const cellBrRoundtrip = String((await OfficeGenerator.generate(await parseHtml(cellBrHtml), 'md')).value);
    assert.ok(/a<br>b/.test(cellBrRoundtrip), '7.E: a <br> inside a table cell survives md -> html -> md');

    // 7.A: a single-line code node with a language stays a fenced block. The inline-vs-fenced
    // decision used to key only off a newline, so a one-line code block with a language collapsed
    // to an inline span, silently dropping the language and its block-ness. A `code` node is always
    // block-level (inline code is a monospace text node), so a language always forces a fence.
    const jsBlock = String((await OfficeGenerator.generate(await parseMd('```js\nconst x = 1;\n```'), 'md')).value);
    assert.ok(/```js\nconst x = 1;\n```/.test(jsBlock), '7.A: a single-line ```js block stays fenced, keeping its language');
    const merBlock = String((await OfficeGenerator.generate(await parseHtml('<div class="mermaid" data-mermaid="graph TD; A--&gt;B"></div>'), 'md')).value);
    assert.ok(/```mermaid\ngraph TD; A-->B\n```/.test(merBlock), '7.A: a single-line mermaid diagram stays a fenced ```mermaid block');

    // Inline code keeps its backticks. Inline code parses to a monospace text node, which had no
    // backtick emission in MarkdownGenerator, so every inline `code` (and inline <code> from HTML)
    // degraded to plain text on md->md and html->md. It is now re-wrapped and fence-sized.
    assert.ok(/use `x` here/.test(String((await OfficeGenerator.generate(await parseMd('use `x` here'), 'md')).value)), 'inline code keeps its backticks on a md round trip');
    assert.ok(/use `x` here/.test(String((await OfficeGenerator.generate(await parseHtml('<p>use <code>x</code> here</p>'), 'md')).value)), 'inline <code> keeps its backticks on html -> md');
    assert.ok(/``x`y``/.test(String((await OfficeGenerator.generate(await parseMd('a ``x`y`` b'), 'md')).value)), 'inline code fence grows past an embedded backtick');

    // 7.D: a generated table header is valid HTML and self-idempotent. The header cells used to sit
    // as bare <th> directly under <thead> (no <tr>), which HtmlParser could not read back - so a
    // md -> HTML -> md round trip lost the header content. It is now <thead><tr><th>.
    const tblHtml = String((await OfficeGenerator.generate(await parseMd('| Feature | Status |\n| --- | --- |\n| A | ok |'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<thead>\s*<tr>\s*<th/.test(tblHtml), '7.D: a table header row is wrapped in <tr> (valid <thead><tr><th>)');
    const tblBack = String((await OfficeGenerator.generate(await parseHtml(tblHtml), 'md')).value);
    assert.ok(/\|\s*Feature\s*\|\s*Status\s*\|/.test(tblBack), '7.D: a generated table survives md -> HTML -> md with its header intact');

    // 8.D: a nested list survives html -> md. An HTML <li> with a <p> child used to emit
    // `- a\n\n\n    - a1` (the paragraph's trailing blank line), which split the list apart and
    // was then flattened on reparse. The item is now a single tight line, and the nesting and
    // shared listId survive the reparse.
    const nestedMd = String((await OfficeGenerator.generate(await parseHtml('<ul><li><p>a</p><ul><li><p>a1</p></li></ul></li></ul>'), 'md')).value).replace(/\n+$/, '');
    assert.strictEqual(nestedMd, '- a\n    - a1', '8.D: nested html list exports to a tight `- a\\n    - a1`');
    const nestedRe = (await parseMd(nestedMd)).content.filter(n => n.type === 'list');
    assert.deepStrictEqual(nestedRe.map(n => (n.metadata as any).indentation), [0, 1], '8.D: reparsed nested list keeps indentations [0, 1]');
    assert.strictEqual((nestedRe[0].metadata as any).listId, (nestedRe[1].metadata as any).listId, '8.D: reparsed nested items share one listId');

    // 8.D: the loose shape a buggy older generator (or a foreign editor) wrote - a blank line
    // between a parent item and its indented child - is re-joined so the child nests again.
    const looseRe = (await parseMd('- a\n\n\n    - a1')).content.filter(n => n.type === 'list');
    assert.deepStrictEqual(looseRe.map(n => (n.metadata as any).indentation), [0, 1], '8.D: loose `- a\\n\\n\\n    - a1` reparses as nested [0, 1]');
    assert.strictEqual((looseRe[0].metadata as any).listId, (looseRe[1].metadata as any).listId, '8.D: re-joined loose list shares one listId');

    // 8.D: an unindented sibling after a blank line stays a separate (flat) list - the merge is
    // deliberately conservative and only pulls in indented children.
    const flatRe = (await parseMd('- a\n\n- b')).content.filter(n => n.type === 'list');
    assert.deepStrictEqual(flatRe.map(n => (n.metadata as any).indentation), [0, 0], '8.D: an unindented loose sibling stays flat [0, 0]');

    // 8.D: a multi-paragraph item collapses its internal break to the item line. `<br>` when the
    // fallback is on (default), a space when off - mirroring table cells under `cellLineBreaks`.
    const multiP = await parseHtml('<ul><li><p>F</p><p>S</p></li></ul>');
    assert.ok(/^- F<br>S/.test(String((await OfficeGenerator.generate(multiP, 'md')).value)), '8.D: multi-paragraph item joins with <br> by default');
    assert.ok(/^- F S/.test(String((await OfficeGenerator.generate(multiP, 'md', { mdConfig: { fallbackToHtml: { itemLineBreaks: false } } })).value)), '8.D: itemLineBreaks:false joins with a space');

    // 8.D: the generated HTML nests spec-validly - a nested list sits INSIDE its parent's still-open
    // <li>, and no <ul>/<ol> directly contains another (the old invalid sibling shape).
    const nestedListHtml = String((await OfficeGenerator.generate(await parseHtml('<ul><li><p>a</p><ul><li><p>a1</p></li></ul></li></ul>'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<li[^>]*>(?:(?!<\/li>)[\s\S])*?<ul/.test(nestedListHtml), '8.D: nested <ul> sits inside an open <li>');
    assert.ok(!/<[uo]l>\s*<[uo]l/.test(nestedListHtml), '8.D: no list directly contains another list');

    // 8.G: GFM per-column table alignment survives md -> HTML -> md. Alignment lives on
    // CellMetadata.align; HtmlGenerator emits it as `text-align` on each <th>/<td> and HtmlParser
    // reads it back, so the `:---`/`:---:`/`---:` markers are not lost when a table passes through
    // HTML (the markdownwriter import path). There was no HTML-round-trip case before md<->md - which
    // is exactly why 8.G slipped.
    const alignHtml = String((await OfficeGenerator.generate(await parseMd('| A | B | C |\n|:--|:-:|--:|\n| 1 | 2 | 3 |'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/<t[hd][^>]*style="[^"]*text-align:\s*left/.test(alignHtml), '8.G: left column emits text-align:left');
    assert.ok(/<t[hd][^>]*style="[^"]*text-align:\s*center/.test(alignHtml), '8.G: center column emits text-align:center');
    assert.ok(/<t[hd][^>]*style="[^"]*text-align:\s*right/.test(alignHtml), '8.G: right column emits text-align:right');
    const alignBack = String((await OfficeGenerator.generate(await parseHtml(alignHtml), 'md', { mdConfig: { dialect: 'extended' } })).value);
    assert.ok(/\|\s*:---\s*\|\s*:---:\s*\|\s*---:\s*\|/.test(alignBack), '8.G: md -> HTML -> md preserves | :--- | :---: | ---: |');

    // 8.G: an unaligned table injects no text-align (byte-identical to before this change).
    const plainHtml = String((await OfficeGenerator.generate(await parseMd('| A | B |\n| --- | --- |\n| 1 | 2 |'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(!/text-align/.test(plainHtml), '8.G: an unaligned table emits no text-align');

    // 8.G: html -> md reads BOTH the new per-cell `text-align` form and the existing table-level
    // `<table data-align>` form.
    const fromCells = String((await OfficeGenerator.generate(await parseHtml('<table><thead><tr><th style="text-align:right">A</th></tr></thead><tbody><tr><td style="text-align:right">1</td></tr></tbody></table>'), 'md', { mdConfig: { dialect: 'extended' } })).value);
    assert.ok(/\|\s*---:\s*\|/.test(fromCells), '8.G: html -> md reads per-cell text-align into the separator');
    const fromTableAlign = String((await OfficeGenerator.generate(await parseHtml('<table data-align="center"><thead><tr><th>A</th></tr></thead><tbody><tr><td>1</td></tr></tbody></table>'), 'md', { mdConfig: { dialect: 'extended' } })).value);
    assert.ok(/\|\s*:---:\s*\|/.test(fromTableAlign), '8.G: html -> md still reads the table-level data-align form');

    // ── Round 9: embeds (leaf directive, dialect.embeds modes, parser parity, gated contract) ──
    const embedMeta = (ast: OfficeParserAST) => collectAllNodes(ast).find(n => n.type === 'embed')?.metadata as any;

    // 9.C: the same YouTube iframe yields the same 'youtube' embed from both parsers (was: md gave
    // a generic 'iframe' with no videoId, gated behind preserveIframes; html gave 'youtube').
    const ytIframe = '<iframe src="https://www.youtube.com/embed/dQw4w9WgXcQ" width="560" height="315"></iframe>';
    const mdEmb = embedMeta(await parseMd(ytIframe));
    const htmlEmb = embedMeta(await parseHtml(ytIframe));
    assert.strictEqual(mdEmb?.embedType, 'youtube', '9.C: a YouTube iframe in .md parses as a youtube embed');
    assert.strictEqual(mdEmb?.videoId, 'dQw4w9WgXcQ', '9.C: the md youtube embed carries the videoId');
    assert.deepStrictEqual([htmlEmb?.embedType, htmlEmb?.videoId], [mdEmb?.embedType, mdEmb?.videoId], '9.C: md/html YouTube parity');

    // 9.A/9.B: an embed leaf directive round-trips md -> AST -> md under dialect.embeds:'directive'.
    const ytDir = '::youtube[Rick Astley]{id=dQw4w9WgXcQ width=80% align=center}';
    assert.strictEqual(String((await OfficeGenerator.generate(await parseMd(ytDir), 'md', { mdConfig: { dialect: { embeds: 'directive' } } })).value).trim(), ytDir, '9.A/9.B: ::youtube directive round-trips stably');

    // 9.B: default (html) embed output is byte-identical to before; the other modes emit their form.
    const ytAst = await parseHtml('<div data-youtube-video="dQw4w9WgXcQ"></div>');
    assert.strictEqual(String((await OfficeGenerator.generate(ytAst, 'md')).value).trim(), '<div data-youtube-video="dQw4w9WgXcQ"></div>', '9.B: default embed md output unchanged (html mode)');
    assert.ok(/^\[YouTube\]\(https:\/\//.test(String((await OfficeGenerator.generate(ytAst, 'md', { mdConfig: { dialect: { embeds: 'link' } } })).value).trim()), '9.B: link mode emits a plain link');
    assert.ok(/^\[!\[YouTube\]\(https:\/\/img\.youtube\.com/.test(String((await OfficeGenerator.generate(ytAst, 'md', { mdConfig: { dialect: { embeds: 'thumbnail' } } })).value).trim()), '9.B: thumbnail mode emits a clickable thumbnail');
    assert.ok(/^\[YouTube\]\(https:\/\//.test(String((await OfficeGenerator.generate(ytAst, 'md', { mdConfig: { fallbackToHtml: { embeds: false } } })).value).trim()), '9.B: deprecated fallbackToHtml.embeds:false still maps to link');

    // 9.A security: ::embed is gated behind preserveIframes (trust input); a hostile src stays inert.
    const embDirVal = '::embed[App]{src=https://app.example.com/x width=100% height=400px}';
    assert.strictEqual(embedMeta(await parseMd(embDirVal)), undefined, '9.A: ::embed is not parsed without preserveIframes (stays literal text)');
    const embDirTrust = embedMeta(await OfficeParser.parseOffice(Buffer.from(embDirVal), { fileType: 'md', htmlParserConfig: { preserveIframes: true } }));
    assert.strictEqual(embDirTrust?.embedType, 'iframe', '9.A: ::embed under preserveIframes parses to an iframe embed');
    assert.strictEqual(embDirTrust?.url, 'https://app.example.com/x', '9.A: ::embed carries its src');

    // 9.F: gatedEmbeds emits an inert placeholder that round-trips back to the same embed; a hostile
    // src is dropped on emit; the default (a live <iframe>) is unchanged.
    const genIframeAst = await OfficeParser.parseOffice(Buffer.from('<iframe src="https://app.example.com/x" width="100%" height="400"></iframe>'), { fileType: 'html', htmlParserConfig: { preserveIframes: true } });
    assert.ok(/<iframe src="https:\/\/app\.example\.com\/x"/.test(String((await OfficeGenerator.generate(genIframeAst, 'html', { htmlConfig: { standalone: false } })).value)), '9.F: default keeps a live <iframe>');
    const gatedHtml = String((await OfficeGenerator.generate(genIframeAst, 'html', { htmlConfig: { standalone: false, gatedEmbeds: true } })).value);
    assert.ok(/<div data-embed-gated data-embed-src="https:\/\/app\.example\.com\/x"/.test(gatedHtml), '9.F: gatedEmbeds emits an inert placeholder div');
    assert.strictEqual(embedMeta(await parseHtml(gatedHtml))?.url, 'https://app.example.com/x', '9.F: the gated placeholder round-trips back to an embed node');
    const hostileGated = await parseHtml('<div data-embed-gated data-embed-src="javascript:alert(1)"></div>');
    assert.ok(!/javascript:/.test(String((await OfficeGenerator.generate(hostileGated, 'html', { htmlConfig: { standalone: false, gatedEmbeds: true } })).value)), '9.F: a hostile gated src is dropped on emit');

    // 9.G: opt-in folk-form import (off by default so a genuine image link is never mangled).
    const obsForm = '![Rick](https://www.youtube.com/watch?v=dQw4w9WgXcQ)';
    assert.strictEqual(embedMeta(await parseMd(obsForm)), undefined, '9.G: an Obsidian youtube-image is not upgraded by default');
    const obsEmbed = embedMeta(await OfficeParser.parseOffice(Buffer.from(obsForm), { fileType: 'md', htmlParserConfig: { embedFolkForms: true } }));
    assert.strictEqual(obsEmbed?.embedType, 'youtube', '9.G: embedFolkForms upgrades an Obsidian youtube-image to a youtube embed');
    assert.strictEqual(obsEmbed?.videoId, 'dQw4w9WgXcQ', '9.G: the folk-form embed carries the videoId');
    const thumbForm = '[![Rick](https://img.youtube.com/vi/dQw4w9WgXcQ/hqdefault.jpg)](https://www.youtube.com/watch?v=dQw4w9WgXcQ)';
    assert.strictEqual(embedMeta(await OfficeParser.parseOffice(Buffer.from(thumbForm), { fileType: 'md', htmlParserConfig: { embedFolkForms: true } }))?.videoId, 'dQw4w9WgXcQ', '9.G: embedFolkForms upgrades a thumbnail-link to a youtube embed');
    const realImgNodes = collectAllNodes(await OfficeParser.parseOffice(Buffer.from('![pic](https://example.com/p.png)'), { fileType: 'md', htmlParserConfig: { embedFolkForms: true } }));
    assert.ok(realImgNodes.some(n => n.type === 'image') && !realImgNodes.some(n => n.type === 'embed'), '9.G: a non-YouTube image stays an image even under embedFolkForms');

    // 10.A: an inline link inside a paragraph is not fenced by blank lines on MD -> HTML, so a
    // punctuation-adjacent link no longer gains a stray space on the round trip.
    const linkHtml = String((await OfficeGenerator.generate(await parseMd('See this [video](https://ex.com/x).'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/>video<\/a>\./.test(linkHtml), '10.A: an inline <a> sits tight against the following punctuation (no blank line)');
    assert.strictEqual(String((await OfficeGenerator.generate(await parseHtml(linkHtml), 'md')).value).trim(), 'See this [video](https://ex.com/x).', '10.A: md -> HTML -> md keeps the link tight to punctuation (no stray space)');

    // 10.B: a youtube embed's label survives AST -> editor-HTML -> AST, at parity with the gated path.
    const ytLabelHtml = String((await OfficeGenerator.generate(await parseMd('::youtube[Carl Sagan]{id=dQw4w9WgXcQ}'), 'html', { htmlConfig: { standalone: false } })).value);
    assert.ok(/data-embed-label="Carl Sagan"/.test(ytLabelHtml), '10.B: youtube editor-HTML carries data-embed-label');
    assert.strictEqual(embedMeta(await parseHtml(ytLabelHtml))?.label, 'Carl Sagan', '10.B: the youtube label round-trips through editor-HTML');
    assert.ok(!/data-embed-label/.test(String((await OfficeGenerator.generate(await parseHtml('<div data-youtube-video="abc"></div>'), 'html', { htmlConfig: { standalone: false } })).value)), '10.B: an unlabeled youtube embed emits no label attribute (byte-identical)');

    // Fable review: the native PDF engine (engine:'native') must draw footnote/endnote bodies, which
    // parsers attach to the TEXT RUN (not the paragraph) - a run never passes through render(), so a
    // render()-level sweep dropped them. Build a doc whose note hangs off a text child, render a native
    // PDF, re-parse it, and assert the note body is present.
    const noteAst: any = {
        type: 'docx', metadata: {}, config: {}, attachments: [],
        content: [{ type: 'paragraph', children: [
            { type: 'text', text: 'Body text with a marker', notes: [{ type: 'note', children: [{ type: 'text', text: 'UNIQUE_FOOTNOTE_BODY_42' }] }] },
        ] }],
    };
    const nativePdf = (await OfficeGenerator.generate(noteAst, 'pdf', { pdfConfig: { engine: 'native' } })).value as Uint8Array;
    const nativePdfText = (await (await OfficeParser.parseOffice(Buffer.from(nativePdf))).to('text')).value as string;
    assert.ok(nativePdfText.includes('UNIQUE_FOOTNOTE_BODY_42'), 'native PDF draws footnote bodies attached to text runs');

    // Fable review: a self-contained (standalone) HTML document must inline every image regardless of
    // maxInlineImageBytes (a name reference there is a broken image); a fragment keeps the cap but must
    // warn (IMAGE_NOT_INLINED) rather than silently emit a bare src.
    const bigImg = () => ({ type: 'docx', metadata: {}, config: {}, content: [{ type: 'image', metadata: { attachmentName: 'big.png' } }], attachments: [{ name: 'big.png', type: 'image', data: 'A'.repeat(2_200_000), mimeType: 'image/png', extension: 'png' }] } as any);
    const bigStandalone = String((await OfficeGenerator.generate(bigImg(), 'html', { htmlConfig: { standalone: true } })).value);
    assert.ok(/data:image\/png;base64,AAAA/.test(bigStandalone), 'standalone HTML inlines an image larger than maxInlineImageBytes');
    let imgWarned = false;
    const bigFragment = String((await OfficeGenerator.generate(bigImg(), 'html', { htmlConfig: { standalone: false }, onWarning: (w: any) => { if (w.code === 'IMAGE_NOT_INLINED') imgWarned = true; } })).value);
    assert.ok(/src="big\.png"/.test(bigFragment) && !/data:image/.test(bigFragment), 'fragment HTML references an over-cap image by name');
    assert.ok(imgWarned, 'over-cap image in a fragment emits an IMAGE_NOT_INLINED warning');

    console.log('  Generated output: All assertions passed ✓');
}

// ─── Entry point ─────────────────────────────────────────────────────────────

/**
 * ODG (OpenDocument Graphics / LibreOffice Draw): parse-only support. Verifies each draw:page
 * becomes a `page` node, shape text (incl. custom-shapes and groups) and embedded tables/images are
 * extracted, and that ODG is not misrouted through the PDF-specific HTML styling.
 */
async function testOdg(): Promise<void> {
    const filePath = path.join(__dirname, 'files/test.odg');
    const ast = await OfficeParser.parseOffice(filePath, { extractAttachments: true });
    assert.strictEqual(ast.type, 'odg', 'ODG: ast.type is odg');

    const all: any[] = [];
    const walk = (n: any) => { all.push(n); (n.children || []).forEach(walk); };
    ast.content.forEach(walk);

    const pages = all.filter(n => n.type === 'page');
    assert.strictEqual(pages.length, 2, `ODG: 2 page nodes, got ${pages.length}`);
    assert.strictEqual((pages[0].metadata as any)?.pageName, 'Title Page', 'ODG: first page draw:name');
    assert.strictEqual((pages[1].metadata as any)?.pageName, 'Data Page', 'ODG: second page draw:name');

    const flat = all.filter(n => n.type === 'text').map(n => n.text).join(' ');
    for (const w of ['Flowchart Demonstration', 'Start: begin', 'Decision: is the condition',
        'Grouped shape A', 'Grouped shape B', 'Name', 'Value', 'Description', 'Alpha', '42', 'Beta']) {
        assert.ok(flat.includes(w), `ODG: shape/table text "${w}" extracted`);
    }

    const tables = all.filter(n => n.type === 'table');
    assert.strictEqual(tables.length, 1, `ODG: 1 embedded table, got ${tables.length}`);
    const images = all.filter(n => n.type === 'image');
    assert.strictEqual(images.length, 1, `ODG: 1 embedded image, got ${images.length}`);
    assert.ok((ast.attachments?.length ?? 0) >= 1, 'ODG: image captured as attachment');
    assert.strictEqual(ast.metadata.title, 'ODG Test Drawing', 'ODG: metadata.title from meta.xml');

    // ODG has `page` nodes but must NOT be styled as a PDF (the isPdf inference keys off ast.type).
    const html = String((await ast.to('html', { htmlConfig: { standalone: false } })).value);
    assert.ok(!html.includes('pdf-container'), 'ODG: HTML is not given the PDF container class');
    for (const w of ['Flowchart Demonstration', 'Alpha']) {
        assert.ok(html.includes(w), `ODG: HTML output contains "${w}"`);
    }
}

/**
 * ODF comments live at three placements: inline in a text paragraph (any ODF type), on the spreadsheet
 * cell (`.ods`), and on the slide/page (`.odp`/`.odg`). All three must land on `.comments`, none must
 * leak the comment body into the cell/slide text, and `ignoreComments` must suppress them.
 */
async function testOdfComments(): Promise<void> {
    const NS = `xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"`;
    const manifest = (mt: string) => `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mt}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>`;
    const pkg = (mt: string, content: string) => Buffer.from(zipSync({ mimetype: strToU8(mt), 'content.xml': strToU8(content), 'META-INF/manifest.xml': strToU8(manifest(mt)) }, { level: 0 }));

    const collectComments = (ast: OfficeParserAST): OfficeContentNode[] => {
        const out: OfficeContentNode[] = [];
        const walk = (n: OfficeContentNode) => { if (n.comments) out.push(...n.comments); (n.children || []).forEach(walk); };
        ast.content.forEach(walk);
        return out;
    };

    // ODS: cell note on a cell that also carries a value.
    const odsMt = 'application/vnd.oasis.opendocument.spreadsheet';
    const odsContent = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><office:annotation><dc:creator>Alice</dc:creator><dc:date>2026-01-01</dc:date><text:p>Cell note here</text:p></office:annotation><text:p>CellValue</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const odsAst = await OfficeParser.parseOffice(pkg(odsMt, odsContent), { fileType: 'ods' });
    const odsComments = collectComments(odsAst);
    assert.strictEqual(odsComments.length, 1, `ODS: 1 cell note, got ${odsComments.length}`);
    assert.strictEqual((odsComments[0].metadata as any)?.author, 'Alice', 'ODS: cell note author');
    assert.ok((odsComments[0].text || '').includes('Cell note here'), 'ODS: cell note text');
    const odsBody = JSON.stringify(odsAst.content.map(n => ({ ...n, comments: undefined, children: (n.children || []).map(r => ({ ...r, children: (r.children || []).map(c => ({ ...c, comments: undefined })) })) })));
    assert.ok(odsBody.includes('CellValue'), 'ODS: cell value preserved');
    assert.ok(!odsBody.includes('Cell note here'), 'ODS: comment body does not leak into the cell text');
    assert.strictEqual(collectComments(await OfficeParser.parseOffice(pkg(odsMt, odsContent), { fileType: 'ods', ignoreComments: true })).length, 0, 'ODS: ignoreComments suppresses the note');

    // ODP: slide-level comment as a direct child of draw:page.
    const odpMt = 'application/vnd.oasis.opendocument.presentation';
    const odpContent = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:presentation><draw:page draw:name="Slide1"><office:annotation><dc:creator>Bob</dc:creator><dc:date>2026-02-02</dc:date><text:p>Slide comment here</text:p></office:annotation><draw:frame><draw:text-box><text:p>Slide body</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`;
    const odpAst = await OfficeParser.parseOffice(pkg(odpMt, odpContent), { fileType: 'odp' });
    const odpComments = collectComments(odpAst);
    assert.strictEqual(odpComments.length, 1, `ODP: 1 page comment, got ${odpComments.length}`);
    assert.strictEqual((odpComments[0].metadata as any)?.author, 'Bob', 'ODP: page comment author');
    const odpBody = JSON.stringify(odpAst.content.map(n => ({ ...n, comments: undefined })));
    assert.ok(odpBody.includes('Slide body'), 'ODP: slide body text preserved');
    assert.ok(!odpBody.includes('Slide comment here'), 'ODP: comment body does not leak into the slide text');
    assert.strictEqual(collectComments(await OfficeParser.parseOffice(pkg(odpMt, odpContent), { fileType: 'odp', ignoreComments: true })).length, 0, 'ODP: ignoreComments suppresses the comment');

    // A cell carrying a comment and repeated across many columns must SHARE the comment node by
    // reference, not deep-copy it per column (a large comment x number-columns-repeated would otherwise
    // amplify into hundreds of MB / a RangeError). Assert sharing rather than a fragile memory bound.
    const note = 'N'.repeat(2000);
    const repeated = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell table:number-columns-repeated="20000" office:value-type="string"><office:annotation><dc:creator>A</dc:creator><text:p>${note}</text:p></office:annotation><text:p>V</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`;
    const repAst = await OfficeParser.parseOffice(pkg(odsMt, repeated), { fileType: 'ods' });
    const repCells: OfficeContentNode[] = [];
    const walkCells = (n: OfficeContentNode) => { if (n.type === 'cell') repCells.push(n); (n.children || []).forEach(walkCells); };
    repAst.content.forEach(walkCells);
    const commented = repCells.filter(c => c.comments && c.comments.length);
    assert.ok(commented.length >= 2, `ODS repeat: many cells carry the comment, got ${commented.length}`);
    assert.strictEqual(commented[0].comments![0], commented[1].comments![0], 'ODS repeat: repeated cells share the comment node by reference (no per-cell duplication)');

    // Same amplification class in an EMBEDDED table (the general parseTable path, used by ODT/ODP/ODG):
    // a content-bearing cell repeated across many columns must share its child nodes by reference, not
    // deep-copy the whole cell body per column.
    const odtMt = 'application/vnd.oasis.opendocument.text';
    const embedded = `<?xml version="1.0"?><office:document-content ${NS}><office:body><office:text><table:table table:name="T"><table:table-row><table:table-cell table:number-columns-repeated="20000"><text:p>${note}</text:p></table:table-cell></table:table-row></table:table></office:text></office:body></office:document-content>`;
    const embAst = await OfficeParser.parseOffice(pkg(odtMt, embedded), { fileType: 'odt' });
    const embCells: OfficeContentNode[] = [];
    const walkEmb = (n: OfficeContentNode) => { if (n.type === 'cell') embCells.push(n); (n.children || []).forEach(walkEmb); };
    embAst.content.forEach(walkEmb);
    const bodied = embCells.filter(c => c.children && c.children.length && c.children[0].text);
    assert.ok(bodied.length >= 2, `ODT embedded repeat: many cells carry the body, got ${bodied.length}`);
    assert.strictEqual(bodied[0].children![0], bodied[1].children![0], 'ODT embedded repeat: repeated cells share the child node by reference (no per-cell duplication)');
}

/**
 * Cross-format consistency guarantees a user relies on (from the consistency review): an option must
 * not silently no-op where a user would expect it to work, and the CLI and library must agree.
 */
async function testConsistencyBehaviors(): Promise<void> {
    const warnCodes = (fn: (onWarning: (i: any) => void) => Promise<any>) => (async () => {
        const codes: string[] = [];
        await fn((i) => codes.push(i.code));
        return codes;
    })();

    // A-1: `ocr: true` without `extractAttachments` warns in EVERY format, not just PDF.
    const ocrCodes = await warnCodes(onWarning =>
        OfficeParser.parseOffice(Buffer.from('# Hi\n\ntext'), { fileType: 'md', ocr: true, onWarning }));
    assert.ok(ocrCodes.includes('OCR_REQUIRES_ATTACHMENTS'), `ocr without extractAttachments warns for a non-PDF format; got ${ocrCodes.join(',')}`);

    // B-1: `ignoreNotes` removes Markdown footnotes (previously a silent no-op).
    const md = 'See this.[^1]\n\n[^1]: The footnote body.\n';
    const kept = collectAllNodes(await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' }));
    const dropped = collectAllNodes(await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md', ignoreNotes: true }));
    assert.ok(kept.some(n => n.type === 'note'), 'md footnote is a note node by default');
    assert.ok(!dropped.some(n => n.type === 'note'), 'ignoreNotes removes md footnote nodes');
    assert.ok(!JSON.stringify(dropped).includes('The footnote body'), 'ignoreNotes leaves no md footnote text behind');

    // B-1 (HTML): same for a data-footnotes section.
    const html = '<p>See<sup data-footnote-ref="1">1</sup></p><section data-footnotes><div data-footnote-id="1"><p>Body.</p></div></section>';
    const htmlDropped = collectAllNodes(await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html', ignoreNotes: true }));
    assert.ok(!htmlDropped.some(n => n.type === 'note'), 'ignoreNotes removes html footnote nodes');

    // A-3: parse-side csvDelimiter reaches CSV output in the library, matching the CLI.
    const tableMd = '| a | b |\n| - | - |\n| 1 | 2 |\n';
    const semi = String((await (await OfficeParser.parseOffice(Buffer.from(tableMd), { fileType: 'md', csvDelimiter: ';' })).to('csv')).value);
    assert.ok(/1;2/.test(semi), `csvDelimiter reaches to('csv'); got ${JSON.stringify(semi)}`);
    // csvConfig.columnDelimiter still wins over the inherited csvDelimiter.
    const pipe = String((await (await OfficeParser.parseOffice(Buffer.from(tableMd), { fileType: 'md', csvDelimiter: ';' })).to('csv', { csvConfig: { columnDelimiter: '|' } } as any)).value);
    assert.ok(/1\|2/.test(pipe), `csvConfig.columnDelimiter overrides csvDelimiter; got ${JSON.stringify(pipe)}`);

    // E-1: to('csv') on a document with no table/sheet warns instead of returning '' silently.
    const csvCodes: string[] = [];
    const emptyCsv = await (await OfficeParser.parseOffice(Buffer.from('# Just a heading\n\nprose only'), { fileType: 'md' })).to('csv', { onWarning: (i: any) => csvCodes.push(i.code) } as any);
    assert.strictEqual(String(emptyCsv.value), '', 'no-table CSV is empty');
    assert.ok(csvCodes.includes('CONTENT_NOT_REPRESENTABLE'), `empty CSV warns; got ${csvCodes.join(',')}`);
}

// A minimal valid 1x1 PNG (sniffs to 1x1, Tesseract-independent) for image-bearing synthetic ASTs.
const TINY_PNG_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';

/** Counts every DOCX part whose name matches, unzipped once. */
function docxParts(bytes: Uint8Array): Record<string, string> {
    const files = unzipSync(bytes);
    const out: Record<string, string> = {};
    for (const [name, data] of Object.entries(files)) {
        if (name.endsWith('.xml') || name.endsWith('.rels')) out[name] = strFromU8(data as Uint8Array);
    }
    return out;
}

/**
 * DOCX generation. Two tiers:
 *  1. Round-trip the richest markdown fixture and confirm structure + specific text survive.
 *  2. Drive a synthetic AST whose shapes exercise exactly the OOXML edge cases a plain fixture never
 *     reaches - a hyperlink and a list inside a footnote, an image inside a header, duplicate heading
 *     slugs, a comment with no date, merged table cells, an orphan note - and assert the package stays
 *     well-formed with correct PER-PART relationships (the class of bug a document-wide rels list hides).
 */
async function testDocxGeneration(): Promise<void> {
    // ── Tier 1: fixture round-trip ────────────────────────────────────────────
    const src = await OfficeParser.parseOffice(path.join(__dirname, 'files/exhaustive/markdown.md'));
    const { value } = await src.to('docx', { metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } });
    const bytes = value as Uint8Array;
    assert.ok(bytes instanceof Uint8Array && bytes.length > 0, 'DOCX: produced non-empty Uint8Array');
    assert.ok(bytes[0] === 0x50 && bytes[1] === 0x4B, 'DOCX: PK zip signature');

    let parts = docxParts(bytes);
    for (const [name, xml] of Object.entries(parts)) {
        assert.doesNotThrow(() => parseXmlString(xml), `DOCX: ${name} is well-formed XML`);
    }
    const doc = parts['word/document.xml'];
    assert.ok(parts['[Content_Types].xml'], 'DOCX: has [Content_Types].xml');
    assert.ok(parts['word/styles.xml'], 'DOCX: has styles.xml');
    assert.ok(/<w:pStyle w:val="Heading1"\/>/.test(doc), 'DOCX: emits Heading1 style');
    assert.ok(/<w:numPr>/.test(doc) && !!parts['word/numbering.xml'], 'DOCX: emits numbered list + numbering.xml');
    assert.ok(/<w:hyperlink r:id="/.test(doc), 'DOCX: emits an external hyperlink');
    assert.ok(/TargetMode="External"/.test(parts['word/_rels/document.xml.rels']), 'DOCX: external hyperlink relationship present');
    assert.ok(!!parts['word/footnotes.xml'] && /<w:footnoteRef\/>/.test(parts['word/footnotes.xml']), 'DOCX: footnotes carry the numbered marker run');

    const back = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType: 'docx' });
    const cnt = (a: any, t: string) => { let c = 0; const w = (n: any) => { if (n.type === t) c++; (n.children || []).forEach(w); }; a.content.forEach(w); return c; };
    for (const t of ['heading', 'table', 'list']) assert.ok(cnt(back, t) > 0, `DOCX: round-trip preserves ${t} nodes`);
    const textOf = (a: any) => { let s = ''; const w = (n: any) => { if (n.type === 'text' && n.text) s += n.text; (n.children || []).forEach(w); }; a.content.forEach(w); return s; };
    assert.ok(textOf(back).includes('Heading Level 1'), 'DOCX: round-trip preserves a specific heading text');

    // Deep self-parity: formatting and links survive, compared by character volume (robust to a
    // parser splitting or merging runs differently than the generator did).
    const boldChars = (a: any) => { let n = 0; const w = (x: any) => { if (x.type === 'text' && x.formatting?.bold && x.text) n += x.text.length; (x.children || []).forEach(w); (x.notes || []).forEach(w); }; a.content.forEach(w); return n; };
    const linkCount = (a: any) => { let n = 0; const w = (x: any) => { if (x.type === 'text' && (x.metadata as any)?.link) n++; (x.children || []).forEach(w); (x.notes || []).forEach(w); }; a.content.forEach(w); return n; };
    assert.ok(boldChars(src) > 0 && boldChars(back) >= boldChars(src) * 0.8,
        `DOCX: round-trip preserves bold text (${boldChars(back)}/${boldChars(src)} chars)`);
    assert.ok(linkCount(src) > 0 && linkCount(back) >= 1,
        `DOCX: round-trip preserves hyperlinks (${linkCount(back)}/${linkCount(src)})`);

    // ── Tier 2: synthetic hostile-shape AST ───────────────────────────────────
    const footnote = { type: 'note', metadata: { noteType: 'footnote', noteId: 'fn1' }, children: [
        { type: 'paragraph', children: [
            { type: 'text', text: 'note with ' },
            { type: 'text', text: 'a link', metadata: { link: 'https://note.example.com/x', linkType: 'external' } },
        ] },
        { type: 'list', metadata: { listId: 'FL', listType: 'unordered', indentation: 0 }, children: [{ type: 'text', text: 'nested list item' }] },
    ] };
    const comment = { type: 'comment', metadata: { author: 'Reviewer', date: '' }, children: [
        { type: 'paragraph', children: [{ type: 'text', text: 'a review comment' }] } ] };
    const synthetic: any = {
        type: 'docx', metadata: { title: 'Synthetic' },
        attachments: [
            { name: 'himg', mimeType: 'image/png', data: TINY_PNG_B64 },
            { name: 'bimg', mimeType: 'image/png', data: TINY_PNG_B64 },
        ],
        auxiliary: { headers: [ { type: 'header', children: [
            { type: 'paragraph', children: [{ type: 'image', metadata: { attachmentName: 'himg', altText: 'logo' } }] } ] } ] },
        content: [
            { type: 'heading', text: 'Intro', metadata: { level: 1 }, children: [{ type: 'text', text: 'Intro' }] },
            { type: 'heading', text: 'Intro', metadata: { level: 2 }, children: [{ type: 'text', text: 'Intro' }] },
            { type: 'paragraph', comments: [comment], children: [
                { type: 'text', text: 'See ', notes: [footnote] } ] },
            { type: 'table', children: [
                { type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2, style: 'header' }, children: [{ type: 'text', text: 'H' }] }] },
                { type: 'row', children: [
                    { type: 'cell', metadata: { rowSpan: 2 }, children: [{ type: 'text', text: 'R' }] },
                    { type: 'cell', children: [{ type: 'text', text: 'b' }] } ] },
                { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'c' }] }] },
            ] },
            { type: 'image', metadata: { attachmentName: 'bimg', altText: 'body image' } },
            { type: 'note', metadata: { noteType: 'footnote', noteId: 'orphan' }, children: [
                { type: 'paragraph', children: [{ type: 'text', text: 'orphan footnote body' }] } ] },
        ],
        getImages: () => [],
    };
    const sbytes = (await OfficeGenerator.generate(synthetic, 'docx', {})).value as Uint8Array;
    parts = docxParts(sbytes);
    for (const [name, xml] of Object.entries(parts)) {
        assert.doesNotThrow(() => parseXmlString(xml), `DOCX synthetic: ${name} well-formed (critical for note/header parts)`);
    }
    const sdoc = parts['word/document.xml'];

    // Per-part relationships: the footnote's hyperlink resolves against footnotes.xml.rels, NOT the document's.
    assert.ok(parts['word/footnotes.xml'], 'DOCX synthetic: footnotes.xml emitted');
    assert.ok(parts['word/_rels/footnotes.xml.rels'] && /note\.example\.com/.test(parts['word/_rels/footnotes.xml.rels']),
        'DOCX synthetic: footnote hyperlink lives in footnotes.xml.rels');
    assert.ok(!/note\.example\.com/.test(parts['word/_rels/document.xml.rels'] || ''),
        'DOCX synthetic: footnote hyperlink is NOT leaked into document.xml.rels');
    // A list inside the footnote still produces a numbering part.
    assert.ok(parts['word/numbering.xml'] && /<w:numPr>/.test(parts['word/footnotes.xml']),
        'DOCX synthetic: list inside a footnote yields numbering.xml');
    // Header image resolves against header1.xml.rels.
    assert.ok(parts['word/header1.xml'], 'DOCX synthetic: header1.xml emitted');
    assert.ok(parts['word/_rels/header1.xml.rels'] && /relationships\/image/.test(parts['word/_rels/header1.xml.rels']),
        'DOCX synthetic: header image lives in header1.xml.rels');

    // Merged cells.
    assert.ok(/<w:gridSpan w:val="2"\/>/.test(sdoc), 'DOCX synthetic: colSpan -> gridSpan');
    assert.ok(/<w:vMerge w:val="restart"\/>/.test(sdoc) && /<w:vMerge\/>/.test(sdoc), 'DOCX synthetic: rowSpan -> vMerge restart + continuation');

    // A cell merged BOTH across columns and down: the continuation row must be ONE gridSpan'd vMerge
    // cell, not one narrow vMerge per spanned column.
    const dmAst: any = { type: 'docx', metadata: {}, content: [{
        type: 'table', children: [
            { type: 'row', children: [
                { type: 'cell', metadata: { row: 0, col: 0, colSpan: 2, rowSpan: 2 }, children: [{ type: 'text', text: 'BIG' }] },
                { type: 'cell', metadata: { row: 0, col: 2 }, children: [{ type: 'text', text: 'C' }] } ] },
            { type: 'row', children: [
                { type: 'cell', metadata: { row: 1, col: 2 }, children: [{ type: 'text', text: 'y' }] } ] },
        ] }] };
    const dmDoc = docxParts((await OfficeGenerator.generate(dmAst, 'docx' as any, {})).value as Uint8Array)['word/document.xml'];
    assert.ok(/<w:gridSpan w:val="2"\/><w:vMerge\/>/.test(dmDoc), 'DOCX: doubly-merged cell emits one gridSpan+vMerge continuation');
    assert.strictEqual((dmDoc.match(/<w:vMerge\/>/g) || []).length, 1, 'DOCX: doubly-merged continuation is one cell, not one bare vMerge per column');

    // Bookmarks: names and ids are unique; the duplicate slug is disambiguated.
    const names = [...sdoc.matchAll(/<w:bookmarkStart w:id="\d+" w:name="([^"]+)"/g)].map(m => m[1]);
    assert.ok(names.includes('intro') && names.includes('intro_2'), 'DOCX synthetic: duplicate heading slug disambiguated (intro, intro_2)');
    assert.strictEqual(new Set(names).size, names.length, 'DOCX synthetic: bookmark names are unique');
    const bmIds = [...sdoc.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map(m => m[1]);
    assert.strictEqual(new Set(bmIds).size, bmIds.length, 'DOCX synthetic: bookmark ids are unique');

    // Drawing ids (header + body image) are unique.
    const drawIds = [...Object.values(parts).join('').matchAll(/<wp:docPr id="(\d+)"/g)].map(m => m[1]);
    assert.ok(drawIds.length >= 2 && new Set(drawIds).size === drawIds.length, 'DOCX synthetic: drawing ids are unique across parts');

    // Comment with an empty date must not emit an invalid w:date="".
    assert.ok(parts['word/comments.xml'] && !/w:date=""/.test(parts['word/comments.xml']), 'DOCX synthetic: no empty w:date attribute');

    // No custom properties -> root .rels must not reference the custom-properties part.
    assert.ok(!/custom-properties/.test(parts['_rels/.rels']), 'DOCX synthetic: root .rels omits absent custom.xml');

    // Orphan note is kept (two footnotes: the referenced one and the orphan).
    assert.strictEqual((parts['word/footnotes.xml'].match(/<w:footnote w:id="\d+"/g) || []).length, 2, 'DOCX synthetic: orphan note preserved as a footnote');

    // Re-parse: footnote body text and comment survive the round trip.
    const sback = await OfficeParser.parseOffice(Buffer.from(sbytes), { fileType: 'docx' });
    const sjson = JSON.stringify(sback);
    assert.ok(sjson.includes('a review comment'), 'DOCX synthetic: comment text round-trips');
    assert.ok(sjson.includes('note with') || sjson.includes('a link'), 'DOCX synthetic: footnote text round-trips');

    // ── Tier 3: docxConfig knobs actually take effect ─────────────────────────
    const land = (await OfficeGenerator.generate(synthetic, 'docx',
        { docxConfig: { format: 'Legal', landscape: true, margin: { top: 36, right: 18, bottom: 36, left: 18 } } } as any)).value as Uint8Array;
    const ldoc = docxParts(land)['word/document.xml'];
    assert.ok(/<w:pgSz w:w="20160"[^>]*w:orient="landscape"/.test(ldoc), 'DOCX config: Legal + landscape sets pgSz and orient');
    assert.ok(/<w:pgMar w:top="720" w:right="360" w:bottom="720" w:left="360"/.test(ldoc), 'DOCX config: margins (points) convert to twips');
}

/** Asserts every element/attribute namespace prefix used in an XML part is declared on its root. */
function assertPrefixesDeclared(xml: string, label: string): void {
    const rootMatch = /<([a-zA-Z0-9]+:[a-zA-Z0-9-]+)\b([^>]*)>/.exec(xml);
    assert.ok(rootMatch, `${label}: has a prefixed root element`);
    const declared = new Set(['xml']);
    for (const m of rootMatch![2].matchAll(/xmlns:([a-zA-Z0-9]+)=/g)) declared.add(m[1]);
    const rootPrefix = rootMatch![1].split(':')[0];
    declared.add(rootPrefix);
    const used = new Set<string>();
    for (const m of xml.matchAll(/<\/?([a-zA-Z0-9]+):/g)) used.add(m[1]);          // element prefixes
    for (const m of xml.matchAll(/\s([a-zA-Z0-9]+):[a-zA-Z0-9-]+=/g)) { if (m[1] !== 'xmlns') used.add(m[1]); } // attribute prefixes (not xmlns decls)
    for (const p of used) assert.ok(declared.has(p), `${label}: prefix "${p}" is declared on the root`);
}

/**
 * ODT generation. Mirrors testDocxGeneration: round-trip the richest fixture, then a synthetic
 * hostile-shape AST for the ODF constructs a plain fixture lacks (a link inside a footnote, an image
 * in a header, duplicate slugs, a comment without a date, merged cells, an orphan note, awkward
 * whitespace), asserting the package is well-formed with every prefix declared and re-parses cleanly.
 */
async function testOdtGeneration(): Promise<void> {
    // ── Tier 1: fixture round-trip ────────────────────────────────────────────
    const src = await OfficeParser.parseOffice(path.join(__dirname, 'files/exhaustive/markdown.md'));
    const bytes = (await src.to('odt', { metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } })).value as Uint8Array;
    assert.ok(bytes instanceof Uint8Array && bytes.length > 0, 'ODT: produced non-empty Uint8Array');
    assert.ok(bytes[0] === 0x50 && bytes[1] === 0x4B, 'ODT: PK zip signature');
    // mimetype must be the first local entry (name at offset 30) and STORED (method at offset 8).
    assert.strictEqual(strFromU8(bytes.slice(30, 38)), 'mimetype', 'ODT: first zip entry is mimetype');
    assert.strictEqual(bytes[8] | (bytes[9] << 8), 0, 'ODT: mimetype entry is STORED (uncompressed)');

    let files = unzipSync(bytes);
    const partsOf = (f: Record<string, Uint8Array>) => Object.fromEntries(Object.entries(f).map(([n, d]) => [n, strFromU8(d)]));
    let parts = partsOf(files);
    for (const [name, xml] of Object.entries(parts)) {
        if (!name.endsWith('.xml')) continue;
        assert.doesNotThrow(() => parseXmlString(xml), `ODT: ${name} is well-formed`);
        assertPrefixesDeclared(xml, `ODT: ${name}`);
    }
    // Manifest lists exactly the packaged files (both directions), with the right root media-type.
    const manifest = parts['META-INF/manifest.xml'];
    assert.ok(/manifest:full-path="\/"[^>]*media-type="application\/vnd.oasis.opendocument.text"/.test(manifest), 'ODT: manifest root media-type matches mimetype');
    const listed = new Set([...manifest.matchAll(/manifest:full-path="([^"]+)"/g)].map(m => m[1]).filter(p => p !== '/'));
    const packaged = new Set(Object.keys(files).filter(n => n !== 'mimetype' && n !== 'META-INF/manifest.xml'));
    assert.deepStrictEqual([...listed].sort(), [...packaged].sort(), 'ODT: manifest lists exactly the packaged files');

    const content = parts['content.xml'];
    assert.ok(/<text:h text:outline-level="1"/.test(content), 'ODT: emits an outline-level-1 heading');
    assert.ok(/<text:list /.test(content) && /<text:list-item>/.test(content), 'ODT: emits nested lists');
    assert.ok(/<text:list-level-style-bullet/.test(content) || /<text:list-level-style-number/.test(content), 'ODT: list style defines level styles');
    assert.ok(/<text:a xlink:type="simple" xlink:href="http/.test(content), 'ODT: emits an external hyperlink');
    assert.ok(/<text:note [^>]*text:note-class="footnote"><text:note-citation>/.test(content), 'ODT: footnote carries a citation');

    const back = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType: 'odt' });
    const cnt = (a: any, t: string) => { let c = 0; const w = (n: any) => { if (n.type === t) c++; (n.children || []).forEach(w); (n.notes || []).forEach(w); }; a.content.forEach(w); return c; };
    for (const t of ['heading', 'table', 'list']) assert.ok(cnt(back, t) > 0, `ODT: round-trip preserves ${t} nodes`);
    const textOf = (a: any) => { let s = ''; const w = (n: any) => { if (n.type === 'text' && n.text) s += n.text; (n.children || []).forEach(w); (n.notes || []).forEach(w); }; a.content.forEach(w); return s; };
    assert.ok(textOf(back).includes('Heading Level 1'), 'ODT: round-trip preserves a specific heading text');

    // Deep parity: bold volume + link count preserved (robust to run splitting).
    const boldChars = (a: any) => { let n = 0; const w = (x: any) => { if (x.type === 'text' && x.formatting?.bold && x.text) n += x.text.length; (x.children || []).forEach(w); (x.notes || []).forEach(w); }; a.content.forEach(w); return n; };
    const linkCount = (a: any) => { let n = 0; const w = (x: any) => { if (x.type === 'text' && (x.metadata as any)?.link) n++; (x.children || []).forEach(w); (x.notes || []).forEach(w); }; a.content.forEach(w); return n; };
    assert.ok(boldChars(src) > 0 && boldChars(back) >= boldChars(src) * 0.8, `ODT: round-trip preserves bold text (${boldChars(back)}/${boldChars(src)})`);
    assert.ok(linkCount(src) > 0 && linkCount(back) >= 1, `ODT: round-trip preserves hyperlinks (${linkCount(back)}/${linkCount(src)})`);

    // ── Tier 2: synthetic hostile-shape AST ───────────────────────────────────
    const footnote = { type: 'note', metadata: { noteType: 'footnote', noteId: 'fn1' }, children: [
        { type: 'paragraph', children: [
            { type: 'text', text: 'note with ' },
            { type: 'text', text: 'a link', metadata: { link: 'https://note.example.com/x', linkType: 'external' } },
        ] },
        { type: 'list', metadata: { listId: 'FL', listType: 'unordered', indentation: 0 }, children: [{ type: 'text', text: 'nested list item' }] },
    ] };
    const comment = { type: 'comment', metadata: { author: 'Reviewer', date: '' }, children: [
        { type: 'paragraph', children: [{ type: 'text', text: 'a review comment' }] } ] };
    const synthetic: any = {
        type: 'odt', metadata: { title: 'Synthetic' },
        attachments: [
            { name: 'himg', mimeType: 'image/png', data: TINY_PNG_B64 },
            { name: 'bimg', mimeType: 'image/png', data: TINY_PNG_B64 },
        ],
        auxiliary: { headers: [ { type: 'paragraph', children: [{ type: 'image', metadata: { attachmentName: 'himg', altText: 'logo' } }] } ] },
        content: [
            { type: 'heading', text: 'Intro', metadata: { level: 1 }, children: [{ type: 'text', text: 'Intro' }] },
            { type: 'heading', text: 'Intro', metadata: { level: 2 }, children: [{ type: 'text', text: 'Intro' }] },
            { type: 'paragraph', comments: [comment], children: [{ type: 'text', text: 'See ', notes: [footnote] }] },
            { type: 'paragraph', children: [{ type: 'text', text: 'a  b\tc', metadata: { link: '#intro', linkType: 'internal' } }] },
            { type: 'table', children: [
                { type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2, style: 'header' }, children: [{ type: 'text', text: 'H' }] }] },
                { type: 'row', children: [
                    { type: 'cell', metadata: { rowSpan: 2 }, children: [{ type: 'text', text: 'R' }] },
                    { type: 'cell', children: [{ type: 'text', text: 'b' }] } ] },
                { type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'c' }] }] },
            ] },
            { type: 'image', metadata: { attachmentName: 'bimg', altText: 'body image' } },
            { type: 'image', metadata: { attachmentName: 'bimg', altText: 'reused image' } },
            { type: 'note', metadata: { noteType: 'footnote', noteId: 'orphan' }, children: [
                { type: 'paragraph', children: [{ type: 'text', text: 'orphan footnote body' }] } ] },
        ],
        getImages: () => [],
    };
    const sbytes = (await OfficeGenerator.generate(synthetic, 'odt', {})).value as Uint8Array;
    files = unzipSync(sbytes);
    parts = partsOf(files);
    for (const [name, xml] of Object.entries(parts)) {
        if (!name.endsWith('.xml')) continue;
        assert.doesNotThrow(() => parseXmlString(xml), `ODT synthetic: ${name} well-formed`);
        assertPrefixesDeclared(xml, `ODT synthetic: ${name}`);
    }
    const sc = parts['content.xml'];
    const styles = parts['styles.xml'];

    // Footnote body carries its link inline (no separate part, no dangling ref) and a numbered citation.
    assert.ok(/note\.example\.com/.test(sc) && /<text:note-citation>/.test(sc), 'ODT synthetic: footnote link is inline with a citation');
    // Merged cells become covered-table-cell continuations.
    assert.ok(/table:number-columns-spanned="2"/.test(sc), 'ODT synthetic: colSpan -> number-columns-spanned');
    assert.ok(/table:number-rows-spanned="2"/.test(sc) && /<table:covered-table-cell\/>/.test(sc), 'ODT synthetic: rowSpan -> covered-table-cell');
    assert.ok(/<table:table-header-rows>/.test(sc), 'ODT synthetic: header row wrapped in table-header-rows');
    // Duplicate heading slugs -> unique bookmark names; internal link resolves to the first.
    const bmNames = [...sc.matchAll(/<text:bookmark text:name="([^"]+)"/g)].map(m => m[1]);
    assert.ok(bmNames.includes('intro') && bmNames.includes('intro_2'), 'ODT synthetic: duplicate slug disambiguated (intro, intro_2)');
    assert.strictEqual(new Set(bmNames).size, bmNames.length, 'ODT synthetic: bookmark names are unique');
    assert.ok(/xlink:href="#intro"/.test(sc), 'ODT synthetic: internal link resolves to the first-claimed name');
    // Two distinct attachments (header himg + body bimg); bimg is referenced twice but deduped to one
    // Pictures part, so exactly two media files, with unique draw:frame names for every emission.
    assert.strictEqual(Object.keys(files).filter(n => n.startsWith('Pictures/')).length, 2, 'ODT synthetic: distinct images packaged once each (reuse deduped)');
    const frameNames = [...(sc + styles).matchAll(/draw:name="([^"]+)"/g)].map(m => m[1]);
    assert.ok(frameNames.length >= 3 && new Set(frameNames).size === frameNames.length, 'ODT synthetic: draw:frame names are unique across reuse');
    // Comment with an empty date must not emit an invalid dc:date.
    assert.ok(!/<dc:date><\/dc:date>/.test(sc) && !/<dc:date\/>/.test(sc), 'ODT synthetic: no empty dc:date element');
    // Whitespace encoding.
    assert.ok(/<text:tab\/>/.test(sc) && /<text:s /.test(sc), 'ODT synthetic: tabs and space runs are encoded');
    // Header image lives in Pictures + manifest, header content in styles.xml.
    assert.ok(/<style:header>/.test(styles) && /Pictures\//.test(styles), 'ODT synthetic: header image referenced from styles.xml');

    // Re-parse: comment, footnote, header, and whitespace survive.
    const sback = await OfficeParser.parseOffice(Buffer.from(sbytes), { fileType: 'odt' });
    const sjson = JSON.stringify(sback);
    assert.ok(sjson.includes('a review comment'), 'ODT synthetic: comment text round-trips');
    assert.ok(sjson.includes('note with') || sjson.includes('a link'), 'ODT synthetic: footnote text round-trips');
    assert.ok((sback.auxiliary?.headers?.length ?? 0) > 0, 'ODT synthetic: header round-trips into auxiliary');
    const wsBack = textOf(sback);
    assert.ok(wsBack.includes('a  b\tc'), 'ODT synthetic: awkward whitespace ("a  b\\tc") round-trips exactly');

    // ── Tier 3: odtConfig knobs actually take effect ──────────────────────────
    const land = (await OfficeGenerator.generate(synthetic, 'odt',
        { odtConfig: { format: 'Legal', landscape: true, margin: { top: 36, right: 18, bottom: 36, left: 18 } } } as any)).value as Uint8Array;
    const lstyles = strFromU8(unzipSync(land)['styles.xml']);
    assert.ok(/fo:page-width="14in"/.test(lstyles) && /fo:page-height="8.5in"/.test(lstyles) && /style:print-orientation="landscape"/.test(lstyles), 'ODT config: Legal + landscape sets page geometry');
    assert.ok(/fo:margin-top="36pt"/.test(lstyles) && /fo:margin-left="18pt"/.test(lstyles), 'ODT config: margins (points) are written as fo:margin lengths');

    // Empty AST -> valid package with one empty text:p, no throw.
    const empty: any = { type: 'odt', metadata: {}, attachments: [], content: [], getImages: () => [] };
    const ebytes = (await OfficeGenerator.generate(empty, 'odt', {})).value as Uint8Array;
    const econtent = strFromU8(unzipSync(ebytes)['content.xml']);
    assert.doesNotThrow(() => parseXmlString(econtent), 'ODT empty: content.xml well-formed');
    assert.ok(/<office:text><text:p\/><\/office:text>/.test(econtent), 'ODT empty: emits one empty paragraph');
}

/**
 * Unit coverage for the shared package-generator helpers. `lengthToPt` in particular pins the
 * units contract: a bare number/string is pixels, so a font size must carry 'pt' or it shrinks to
 * 75% - the exact PDF-parser regression this locks down.
 */
async function testOfficeGenUtils(): Promise<void> {
    assert.strictEqual(lengthToPt('11pt'), 11, 'lengthToPt: pt honored');
    assert.strictEqual(lengthToPt('1in'), 72, 'lengthToPt: in -> pt');
    assert.strictEqual(Math.round(lengthToPt('96px')!), 72, 'lengthToPt: px -> pt');
    assert.strictEqual(lengthToPt('11'), 8.25, 'lengthToPt: a bare string is pixels (x0.75)');
    assert.strictEqual(lengthToPt(96), 72, 'lengthToPt: a number is pixels');
    assert.strictEqual(lengthToPt(undefined), null, 'lengthToPt: undefined -> null');

    assert.strictEqual(hexColor('#f00'), 'FF0000', 'hexColor: #rgb expands');
    assert.strictEqual(hexColor('112233'), '112233', 'hexColor: bare 6-hex');
    assert.strictEqual(hexColor('red'), null, 'hexColor: named -> null');
    assert.strictEqual(hexColor('red"/><x'), null, 'hexColor: hostile -> null');

    assert.strictEqual(toBookmarkNameRaw('a b!'), 'a_b_', 'toBookmarkNameRaw sanitizes to [A-Za-z0-9_]');
    assert.ok(/^[A-Za-z_]/.test(toBookmarkNameRaw('9x')), 'toBookmarkNameRaw: forces a leading letter/_');

    const sz = sniffImageSize(decodeBase64(TINY_PNG_B64));
    assert.ok(sz && sz.w === 1 && sz.h === 1, 'sniffImageSize: reads a 1x1 PNG');

    const a = resolveZipInstant(new Date('2024-01-01T00:00:00Z'));
    assert.strictEqual(a.iso, '2024-01-01T00:00:00Z', 'resolveZipInstant: whole-second ISO');
    assert.strictEqual(a.iso, resolveZipInstant('2024-01-01T00:00:00Z').iso, 'resolveZipInstant: Date and string agree');
    // The clamp targets fflate's LOCAL-time floor (fflate stamps zip mtimes from local getters and
    // rejects a local year < 1980), so the clamped instant's local year is 1980 in every timezone.
    assert.strictEqual(resolveZipInstant(new Date('1900-01-01')).mtime.getFullYear(), 1980, 'resolveZipInstant: clamps below the zip 1980 floor');
    // Cross-timezone reproducibility: the mtime's LOCAL fields (what fflate stamps) must equal the
    // instant's UTC calendar fields, so the same instant produces the same DOS timestamp on every
    // machine regardless of its timezone. This holds by construction here and is asserted TZ-independently.
    const rzi = resolveZipInstant(new Date('2024-03-15T09:30:45Z')).mtime;
    assert.ok(rzi.getFullYear() === 2024 && rzi.getMonth() === 2 && rzi.getDate() === 15 && rzi.getHours() === 9 && rzi.getMinutes() === 30 && rzi.getSeconds() === 45,
        'resolveZipInstant: mtime local fields mirror the UTC wall clock (timezone-independent zip bytes)');

    // Header-row inference from the signals parsers actually set (not the test-only `isHeader`).
    const cell = (text: string, style?: string, bold?: boolean) => ({ type: 'cell', metadata: style ? { style } : undefined, children: [{ type: 'text', text, formatting: bold ? { bold: true } : undefined }] });
    const row = (cells: any[], meta?: any) => ({ type: 'row', metadata: meta, children: cells });
    assert.ok(isHeaderRow(row([cell('H', 'header')]) as any, true), 'isHeaderRow: cell style "header" (PDF TH)');
    assert.ok(isHeaderRow(row([cell('H')], { style: 'Table Header' }) as any, false), 'isHeaderRow: row style contains "header"');
    assert.ok(isHeaderRow(row([cell('A', undefined, true), cell('B', undefined, true)]) as any, true), 'isHeaderRow: an all-bold first row');
    assert.ok(!isHeaderRow(row([cell('A', undefined, true), cell('B', undefined, true)]) as any, false), 'isHeaderRow: all-bold is NOT a header beyond row 0');
    assert.ok(!isHeaderRow(row([cell('a'), cell('b')]) as any, true), 'isHeaderRow: a plain first row is not a header');
}

/**
 * The native PDF engine (pdf-lib, Standard-14 fonts) must not throw on characters outside WinAnsi -
 * Greek, arrows, CJK, emoji are all common - and must warn instead of crashing the whole conversion.
 */
async function testNativePdfEngine(): Promise<void> {
    const md = '# Heading Ω → ✓\n\nGreek Ω, arrow →, CJK 日本語, emoji 😀, Café.\n\n- item →\n- plain item\n\n```\nconst x = "日本"; // 😀\n```\n';
    const src = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' });
    const warnings: string[] = [];
    const { value } = await src.to('pdf', { pdfConfig: { engine: 'native' }, onWarning: (i: any) => warnings.push(i.code) } as any);
    const bytes = value as Uint8Array;
    assert.ok(bytes instanceof Uint8Array && bytes.length > 100, 'native PDF: produced non-trivial bytes (no crash on Unicode)');
    assert.strictEqual(strFromU8(bytes.slice(0, 5)), '%PDF-', 'native PDF: has the %PDF- signature');
    assert.ok(warnings.includes('CONTENT_NOT_REPRESENTABLE'), 'native PDF: warns that non-WinAnsi characters were replaced');
    // A purely-Latin document draws cleanly with no such warning.
    const latin = await OfficeParser.parseOffice(Buffer.from('# Hello\n\nPlain ASCII text.\n'), { fileType: 'md' });
    const w2: string[] = [];
    const first = (await latin.to('pdf', { pdfConfig: { engine: 'native' }, onWarning: (i: any) => w2.push(i.code) } as any)).value as Uint8Array;
    assert.ok(!w2.includes('CONTENT_NOT_REPRESENTABLE'), 'native PDF: no spurious warning for Latin-only text');

    // Determinism: a date-less source renders byte-identically every time. PDFDocument.create is passed
    // updateMetadata:false and no /ID is set, so a stray new Date() or a pdf-lib bump would regress this.
    const second = (await (await OfficeParser.parseOffice(Buffer.from('# Hello\n\nPlain ASCII text.\n'), { fileType: 'md' }))
        .to('pdf', { pdfConfig: { engine: 'native' } } as any)).value as Uint8Array;
    assert.ok(Buffer.from(first).equals(Buffer.from(second)), 'native PDF: rendering a date-less source is deterministic (byte-identical)');
}

/**
 * DOCX templating (OfficeTemplate.render). Fixture `test/files/template/invoice.docx` is generated by
 * `scripts/generate-template-fixtures.mjs` and carries a run-split `{{name}}`, a bold `{{amount}}`, a
 * table-cell `{{item}}`, a header `{{company}}`, a multiline `{{note}}`, and an undefined `{{missing}}`.
 */
async function testTemplate(): Promise<void> {
    const tpl = path.join(__dirname, 'files', 'template', 'invoice.docx');
    const data = { name: 'Acme Corp', amount: '$1,250.00', date: '2026-10-01', item: 'Widget', note: 'Line one\nLine two', company: 'GLOBEX' };
    const walk = (n: OfficeContentNode, fn: (x: OfficeContentNode) => void) => { fn(n); (n.children || []).forEach(c => walk(c, fn)); };
    const astText = (ast: OfficeParserAST) => ast.content.map(p => p.text).join('\n');

    const bytes = await OfficeTemplate.render(tpl, { data });
    assert.ok(bytes instanceof Uint8Array, 'render returns a Uint8Array for a single data map');
    const ast = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType: 'docx' }) as OfficeParserAST;
    const text = astText(ast);

    // A placeholder split across runs (`{{` | `na` | `me}}`) is still substituted.
    assert.ok(text.includes('Dear Acme Corp, welcome.'), `run-split {{name}} substituted; got: ${text.slice(0, 60)}`);
    assert.ok(text.includes('as of 2026-10-01'), '{{date}} substituted');
    assert.ok(!text.includes('{{name}}') && !text.includes('{{amount}}'), 'filled placeholders leave no delimiters');

    // A value adopts the formatting (bold) of the run its placeholder occupied.
    let amountBold = false, itemFound = false;
    ast.content.forEach(p => walk(p, nd => {
        if (nd.type === 'text' && String(nd.text).includes('1,250') && (nd.formatting as any)?.bold) amountBold = true;
        if (String(nd.text) === 'Widget') itemFound = true;
    }));
    assert.ok(amountBold, 'bold formatting preserved on the substituted {{amount}}');
    assert.ok(itemFound, 'table-cell {{item}} substituted');

    // Header part is templated too.
    assert.ok((ast.auxiliary?.headers || []).some(h => String(h.text).includes('GLOBEX - Confidential')), 'header {{company}} substituted');

    // Multiline value: both lines present (a <w:br/> separates them in the docx).
    assert.ok(text.includes('Line one') && text.includes('Line two'), 'multiline value rendered');

    // onMissing: keep (default) leaves the tag; empty removes it; error rejects.
    assert.ok(text.includes('{{missing}}'), "onMissing 'keep' leaves an unknown placeholder");
    const emptied = astText(await OfficeParser.parseOffice(Buffer.from(await OfficeTemplate.render(tpl, { data, onMissing: 'empty' })), { fileType: 'docx' }) as OfficeParserAST);
    assert.ok(!emptied.includes('{{missing}}') && emptied.includes('Ref:'), "onMissing 'empty' blanks an unknown placeholder");
    await assert.rejects(() => OfficeTemplate.render(tpl, { data, onMissing: 'error' }) as Promise<any>, (e: any) => e?.officeIssue?.code === 'TEMPLATE_FIELD_MISSING', "onMissing 'error' rejects with TEMPLATE_FIELD_MISSING");

    // Custom delimiters.
    // (The fixture uses {{ }}, so custom delimiters simply must not match and leave the doc unchanged.)
    const customName = astText(await OfficeParser.parseOffice(Buffer.from(await OfficeTemplate.render(tpl, { data, delimiters: { start: '<<', end: '>>' } })), { fileType: 'docx' }) as OfficeParserAST);
    assert.ok(customName.includes('{{name}}'), 'a non-matching delimiter set leaves {{ }} placeholders intact');

    // Batch: an array of data maps yields one document each, correctly separated.
    const docs = await OfficeTemplate.render(tpl, { data: [ { ...data, name: 'First' }, { ...data, name: 'Second' } ] });
    assert.ok(Array.isArray(docs) && docs.length === 2, 'array data yields one document per entry');
    const first = astText(await OfficeParser.parseOffice(Buffer.from(docs[0]), { fileType: 'docx' }) as OfficeParserAST);
    assert.ok(first.includes('Dear First,') && !first.includes('Dear Second,'), 'batch documents get their own data');

    // Determinism: same inputs -> byte-identical output (pinned zip mtime).
    const again = await OfficeTemplate.render(tpl, { data });
    assert.ok(Buffer.from(bytes).equals(Buffer.from(again as Uint8Array)), 'rendering is deterministic');

    // Untouched parts are copied verbatim.
    assert.deepStrictEqual(docxParts(bytes)['[Content_Types].xml'], docxParts(fs.readFileSync(tpl))['[Content_Types].xml'], 'non-text parts are unchanged');

    // --- adversarial cases (each builds a tiny docx around a specific hazard) ---
    const buildDocx = (bodyXml: string): Uint8Array => {
        const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
        const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
        const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document>`;
        return zipSync({ '[Content_Types].xml': strToU8(ct), '_rels/.rels': strToU8(rels), 'word/document.xml': strToU8(doc) });
    };
    const renderDoc = async (bodyXml: string, cfg: any): Promise<string> => {
        const outBytes = await OfficeTemplate.render(Buffer.from(buildDocx(bodyXml)), cfg) as Uint8Array;
        const docXml = strFromU8(unzipSync(outBytes)['word/document.xml']);
        assert.doesNotThrow(() => parseXmlString(docXml), 'rendered document.xml is well-formed XML');
        return docXml;
    };

    // 1. A value's XML metacharacters are escaped (no markup injection, part stays well-formed).
    {
        const xml = await renderDoc('<w:p><w:r><w:t>{{v}}</w:t></w:r></w:p>', { data: { v: 'A & B <c> "x" </w:t>' } });
        assert.ok(xml.includes('A &amp; B &lt;c&gt;'), 'value XML metacharacters are escaped');
        assert.ok(!/<c>/.test(xml), 'value cannot inject a raw element');
    }
    // 2. Custom delimiters made of XML-special chars match the escaped source text.
    {
        const xml = await renderDoc('<w:p><w:r><w:t>&lt;&lt;name&gt;&gt;</w:t></w:r></w:p>', { data: { name: 'Acme' }, delimiters: { start: '<<', end: '>>' } });
        assert.ok(xml.includes('Acme') && !xml.includes('&lt;&lt;name'), 'custom << >> delimiters substitute');
    }
    // 3. A self-closing <w:t/> in a placeholder paragraph does not corrupt the output.
    {
        const xml = await renderDoc('<w:p><w:r><w:t/></w:r><w:r><w:t>{{v}}</w:t></w:r></w:p>', { data: { v: 'OK' } });
        assert.ok(xml.includes('OK'), 'self-closing <w:t/> handled, placeholder still substituted');
    }
    // 4. XML-illegal control characters in a value are stripped (part stays valid).
    {
        const xml = await renderDoc('<w:p><w:r><w:t>{{v}}</w:t></w:r></w:p>', { data: { v: 'a\x01b\x0Cc' } });
        assert.ok(xml.includes('abc'), 'XML-illegal control chars are stripped from the value');
        assert.ok(!/[\x00-\x08\x0B\x0C\x0E-\x1F]/.test(xml), 'no raw control characters remain in the part');
    }
    // 5. Text boxes: nested <w:p> in a run does not swallow later placeholders or join across boundary.
    {
        const body = '<w:p><w:r><w:t>Before {{a}} </w:t></w:r>'
            + '<w:r><mc:AlternateContent xmlns:mc="x"><w:txbxContent><w:p><w:r><w:t>{{b}}</w:t></w:r></w:p></w:txbxContent></mc:AlternateContent></w:r>'
            + '<w:r><w:t> after {{c}}</w:t></w:r></w:p>';
        const xml = await renderDoc(body, { data: { a: 'AA', b: 'BB', c: 'CC' } });
        assert.ok(xml.includes('AA') && xml.includes('BB') && xml.includes('CC'), 'placeholders around and inside a text box all substitute');
        assert.ok(!xml.includes('{{c}}'), 'placeholder after a text box is not skipped');
    }
    // 6. A placeholder split at a text-box boundary is NOT joined across it (no value teleporting).
    {
        const body = '<w:p><w:r><w:t>x {{na</w:t></w:r>'
            + '<w:r><w:txbxContent><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:txbxContent></w:r>'
            + '<w:r><w:t>me}} y</w:t></w:r></w:p>';
        const xml = await renderDoc(body, { data: { name: 'ZZ' } });
        assert.ok(!xml.includes('ZZ'), 'a placeholder is not matched across a paragraph (text-box) boundary');
        assert.ok(xml.includes('inner'), 'the text box content is preserved');
    }
    // 7. The timezone-crash regression: render must succeed west of UTC (fflate stamps mtimes in local
    //    time and rejects year < 1980). Run in a child process because Node caches TZ at startup.
    {
        const { execSync } = await import('child_process');
        const os = await import('os');
        const scriptPath = path.join(os.tmpdir(), `optz_${Date.now()}.mjs`);
        const src = path.join(__dirname, '..', 'src', 'OfficeTemplate.ts').replace(/\\/g, '/');
        fs.writeFileSync(scriptPath, `import { OfficeTemplate } from ${JSON.stringify(src)};\n`
            + `const doc = ${JSON.stringify(Buffer.from(buildDocx('<w:p><w:r><w:t>{{v}}</w:t></w:r></w:p>')).toString('base64'))};\n`
            + `const out = await OfficeTemplate.render(Buffer.from(doc, 'base64'), { data: { v: 'OK' } });\n`
            + `process.stdout.write(out.length > 0 ? 'RENDER_OK' : 'EMPTY');\n`);
        try {
            const out = execSync(`npx tsx ${JSON.stringify(scriptPath)}`, { cwd: path.join(__dirname, '..'), env: { ...process.env, TZ: 'America/Los_Angeles' }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
            assert.ok(out.includes('RENDER_OK'), 'render succeeds under TZ=America/Los_Angeles');
        } finally {
            try { fs.unlinkSync(scriptPath); } catch { /* best effort */ }
        }
    }
}

async function runTests(): Promise<void> {
    console.log('Starting exhaustive officeParser test suite...');
    let passed = 0;
    let failed = 0;

    const tests: Array<[string, () => Promise<void>]> = [
        ['Markdown', testMarkdown],
        ['HTML', testHtml],
        ['CSV', testCsv],
        ['RTF', testRtf],
        ['AttributeRoundtrip', testAttributeRoundtrip],
        ['BlobInput', testBlobInput],
        ['WordParagraphMark', testWordParagraphMarkFormatting],
        ['ExcelDates', testExcelDates],
        ['GeneratedOutput', testGeneratedOutput],
        ['ODG', testOdg],
        ['ODFComments', testOdfComments],
        ['ConsistencyBehaviors', testConsistencyBehaviors],
        ['DOCX', testDocxGeneration],
        ['ODT', testOdtGeneration],
        ['OfficeGenUtils', testOfficeGenUtils],
        ['NativePdfEngine', testNativePdfEngine],
        ['Template', testTemplate],
    ];

    for (const [name, fn] of tests) {
        try {
            await fn();
            passed++;
        } catch (err: any) {
            console.error(`\n✗ ${name} FAILED:`, err.message || err);
            if (err.stack) console.error(err.stack);
            failed++;
        }
    }

    console.log(`\n${'='.repeat(50)}`);
    console.log(`Results: ${passed} passed, ${failed} failed`);
    if (failed > 0) {
        process.exit(1);
    } else {
        console.log('All exhaustive tests passed! ✓');
    }
}

runTests().catch(err => {
    console.error('Unexpected error:', err);
    process.exit(1);
});
