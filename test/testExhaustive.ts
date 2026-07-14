/**
 * Exhaustive test suite for officeParser.
 * Covers every AST node type, metadata field, text formatting flag,
 * and round-trip output correctness for Markdown, HTML, CSV, and RTF formats.
 */

import { OfficeParser } from '../src/OfficeParser';
import { OfficeGenerator } from '../src/OfficeGenerator';
import * as assert from 'assert';
import * as path from 'path';
import type { OfficeContentNode, OfficeParserAST } from '../src/types';

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
    assertExists(codeNodes, n => (n.metadata as any)?.language === 'javascript' && n.text?.includes('tilde fence'), 'MD: tilde-fenced code block');

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
    assertExists(breaks, n => (n.metadata as any)?.breakType === 'textWrapping', 'HTML: textWrapping break');

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

    // ── Math ─────────────────────────────────────────────────────────────────
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'inline', 'HTML: inline math code node');
    assertExists(codeNodes, n => (n.metadata as any)?.math === 'block', 'HTML: block math code node');

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

    // ── Abbreviation ─────────────────────────────────────────────────────────
    assertExists(textNodes, n => (n.metadata as any)?.abbreviationTitle !== undefined, 'HTML: abbreviation text node');

    // ── Footnote ─────────────────────────────────────────────────────────────
    const noteNodes = nodes.filter(n => n.type === 'note');
    assert.ok(noteNodes.length >= 1, 'HTML: Has note nodes');
    assertExists(noteNodes, n => (n.metadata as any)?.noteType === 'footnote', 'HTML: footnote note node');

    // ── Roundtrip: generate to HTML ──────────────────────────────────────────
    const result = await OfficeGenerator.generate(ast, 'html');
    const htmlOutput = result.value as string;
    assert.ok(htmlOutput.includes('<h1'), 'HTML roundtrip: h1 tag');
    assert.ok(htmlOutput.includes('<ul') || htmlOutput.includes('<ol'), 'HTML roundtrip: list tag');
    assert.ok(htmlOutput.includes('<table'), 'HTML roundtrip: table tag');
    assert.ok(htmlOutput.includes('<ol') || htmlOutput.includes('<ul'), 'HTML roundtrip: list');

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

// ─── Entry point ─────────────────────────────────────────────────────────────

async function runTests(): Promise<void> {
    console.log('Starting exhaustive officeParser test suite...');
    let passed = 0;
    let failed = 0;

    const tests: Array<[string, () => Promise<void>]> = [
        ['Markdown', testMarkdown],
        ['HTML', testHtml],
        ['CSV', testCsv],
        ['RTF', testRtf],
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
