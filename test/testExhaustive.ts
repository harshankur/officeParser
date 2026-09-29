/**
 * Exhaustive test suite for officeParser.
 * Covers every AST node type, metadata field, text formatting flag,
 * and round-trip output correctness for Markdown, HTML, CSV, and RTF formats.
 */

import { OfficeParser } from '../src/OfficeParser';
import { OfficeGenerator } from '../src/OfficeGenerator';
import { OfficeTemplate } from '../src/OfficeTemplate';
import { OfficeConverter } from '../src/OfficeConverter';
import { zipSync, strToU8, unzipSync, strFromU8 } from 'fflate';
import * as assert from 'assert';
import * as path from 'path';
import * as fs from 'fs';
import type { ImageMetadata, OfficeContentNode, OfficeParserAST } from '../src/types';
import { parseXmlString } from '../src/utils/xmlUtils';
import { ocrTestHooks, performOcr, terminateOcr } from '../src/utils/ocrUtils';
import { imageFromPdf, imageToTextPdf, newDecodeBudget } from '../src/utils/textPdf';
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

    // `$$...$$` written inside a paragraph is display math, as KaTeX, MathJax, GitLab and Pandoc read it:
    // the paragraph is split around it, with no stray `$` left on either side. Where a block cannot go
    // (a heading, list item, table cell, quote, note, or inside emphasis) it is inline math in place.
    const mdMath = async (src: string) => (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).content;
    const shape = (n: OfficeContentNode): any => n.type === 'code' ? [`math:${(n.metadata as any)?.math}`, n.text] : [n.type, (n.children || []).map(c => c.type === 'code' ? `math:${(c.metadata as any)?.math}:${c.text}` : c.text).join('|')];
    assert.deepStrictEqual((await mdMath('Inline $$a+b$$ end')).map(shape), [['paragraph', 'Inline'], ['math:block', 'a+b'], ['paragraph', 'end']], 'MD: $$...$$ inside a paragraph is display math, the paragraph split around it');
    assert.deepStrictEqual((await mdMath('A $x$ and $$y$$ and $z$.')).map(shape), [['paragraph', 'A |math:inline:x| and'], ['math:block', 'y'], ['paragraph', 'and |math:inline:z|.']], 'MD: single-dollar inline math is unchanged beside display math');
    assert.deepStrictEqual((await mdMath('Two $$a$$ and $$b$$')).map(shape), [['paragraph', 'Two'], ['math:block', 'a'], ['paragraph', 'and'], ['math:block', 'b']], 'MD: several display equations in one paragraph');
    assert.deepStrictEqual((await mdMath('The sum $$a +\nb$$ is here.')).map(shape), [['paragraph', 'The sum'], ['math:block', 'a +\nb'], ['paragraph', 'is here.']], 'MD: display math may span the lines of a paragraph');
    assert.deepStrictEqual((await mdMath('Escaped $$a\\$b$$ and \\$$5$$')).map(shape), [['paragraph', 'Escaped'], ['math:block', 'a\\$b'], ['paragraph', 'and |$|math:inline:5|$']], 'MD: an escaped \\$ inside display math stays in the formula; an escaped $ before $$ is literal');
    for (const [src, expected, label] of [
        ['# Heading $$h$$', [['heading', 'Heading |math:inline:h']], 'heading'],
        ['- item $$i$$ more', [['list', 'item |math:inline:i| more']], 'list item'],
        ['**bold $$b$$**', [['paragraph', 'bold |math:inline:b']], 'emphasis'],
        ['Code `$$x$$` and $$ $$ and $$open', [['paragraph', 'Code |$$x$$| and |$$ $$| and $$open']], 'code spans, blank and unclosed $$ as literal text'],
    ] as const) {
        assert.deepStrictEqual((await mdMath(src)).map(shape), expected, `MD: $$...$$ in a ${label}`);
    }
    const mathAdmonition = (await mdMath(':::note\nNote $$n$$ body\n:::'))[0];
    assert.deepStrictEqual(mathAdmonition.children!.map(shape), [['paragraph', 'Note'], ['math:block', 'n'], ['paragraph', 'body']], 'MD: display math splits an admonition paragraph too');
    const displayOnce = (await (await OfficeParser.parseOffice(Buffer.from('Inline $$a+b$$ end'), { fileType: 'md' } as any)).to('md')).value as string;
    const displayTwice = (await (await OfficeParser.parseOffice(Buffer.from(displayOnce), { fileType: 'md' } as any)).to('md')).value as string;
    assert.ok(displayOnce.includes('$$\na+b\n$$') && !displayOnce.includes('Inline $') && displayTwice === displayOnce, 'MD: display math regenerates as a $$ block, a fixed point after one cycle');

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
    // A link out of the document is external, and a link to `#id` in it internal, as the other parsers read them
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'external' && (n.metadata as any)?.link?.startsWith('https://'), 'MD: external https link text node');
    assertExists(textNodes, n => (n.metadata as any)?.linkType === 'internal' && (n.metadata as any)?.link?.startsWith('#'), 'MD: anchor (#) link text node');
    assert.ok(!textNodes.some(n => (n.metadata as any)?.linkType === 'external' && (n.metadata as any)?.link?.startsWith('#')), 'MD: no anchor (#) link is external');
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

    // ── Source comments (`<!-- ... -->`): hidden notes, kept verbatim ──────────
    const sourceComments = nodes.filter(n => n.type === 'comment' && (n.metadata as any)?.sourceSyntax === 'html');
    const blockComment = ' Exhaustive: a source comment on its own lines,\n\nspanning a blank line, kept verbatim ';
    assert.ok(sourceComments.some(n => n.text === blockComment), 'MD: multi-line block comment kept verbatim (blank line included)');
    assert.ok(sourceComments.some(n => n.text === ' an inline comment '), 'MD: inline comment kept verbatim');
    assert.ok(!sourceComments.some(n => (n.text || '').includes('not a comment')), 'MD: a comment inside a code span stays code');
    assert.ok(nodes.some(n => n.type === 'text' && n.formatting?.font === 'monospace' && n.text === '<!-- not a comment -->'), 'MD: code span keeps the literal comment text');
    assert.ok(!nodes.some(n => n.type === 'text' && (n.text || '').includes('<!-- Exhaustive')), 'MD: a comment is never parsed as visible text');
    assert.ok(mdOutput.includes(`<!--${blockComment}-->`), 'MD roundtrip: block comment re-emitted byte-for-byte');
    assert.ok(mdOutput.includes('Inline <!-- an inline comment --> source comment'), 'MD roundtrip: inline comment re-emitted in its run');
    assert.ok(!mdOutput.includes('&lt;!--'), 'MD roundtrip: no comment escaped into visible text');

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

    // ── HTML comments: dropped by default, kept (minus conditional comments) when asked ──
    const isSourceCommentNode = (n: OfficeContentNode) => n.type === 'comment' && (n.metadata as any)?.sourceSyntax === 'html';
    assert.ok(!nodes.some(isSourceCommentNode), 'HTML: comments are dropped by default');
    const keptAst = await OfficeParser.parseOffice(filePath, { htmlParserConfig: { preserveComments: true } } as any);
    const kept = collectAllNodes(keptAst).filter(isSourceCommentNode);
    assert.ok(kept.some(n => n.text === ' An authored HTML comment, kept under preserveComments '), 'HTML: preserveComments keeps an authored comment verbatim');
    assert.ok(kept.some(n => n.text === ' an inline html comment '), 'HTML: preserveComments keeps an inline comment');
    assert.ok(!kept.some(n => /\[if |endif/.test(n.text || '')), 'HTML: conditional (Office/IE) comments are never kept');
    const keptMd = String((await OfficeGenerator.generate(keptAst, 'md')).value);
    assert.ok(keptMd.includes('<!-- An authored HTML comment, kept under preserveComments -->'), 'HTML->MD: preserved comment becomes a Markdown comment');

    // ── Entity decoding is the exact inverse of escaping (no double decode) ─────
    // Literal text `&quot;`, `&#39;`, `&lt;` is escaped as `&amp;quot;` etc.; decoding `&amp;` first
    // and then `&quot;` used to collapse it to `"`, silently changing the text.
    const entityAst = await OfficeParser.parseOffice(Buffer.from('<p>&amp;quot; &amp;#39; &amp;lt; &quot;</p>'), { fileType: 'html' } as any);
    const entityText = collectAllNodes(entityAst).filter(n => n.type === 'text').map(n => n.text).join('');
    assert.strictEqual(entityText, '&quot; &#39; &lt; "', 'HTML: entities decode in one pass (no double decode)');

    console.log('  HTML: All assertions passed ✓');
}

// ─── Source comments across every generator ──────────────────────────────────

async function testSourceComments(): Promise<void> {
    console.log('\n=== Running Source Comment Tests ===');
    const md = 'First paragraph.\n\n<!-- a hidden note -->\n\nHello <!-- inline "q" & <b> --> world.\n';
    const ast = await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' } as any);

    // Markdown -> Markdown is byte-for-byte.
    assert.strictEqual(String((await OfficeGenerator.generate(ast, 'md')).value), md.trimEnd(), 'Comments: md round trip is byte-for-byte');

    // The editor path: md -> sourceAttributes HTML -> HTML parse -> md, still byte-for-byte. The
    // comment text rides in an escaped attribute (quotes, ampersand and a tag included).
    const editorHtml = String((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { sourceAttributes: true, standalone: false } } as any)).value);
    assert.ok(editorHtml.includes('<span data-html-comment=" a hidden note "></span>'), 'Comments: sourceAttributes emits a data-html-comment span');
    assert.ok(!editorHtml.includes('<!--'), 'Comments: sourceAttributes emits no real comment (an editor would discard it)');
    const fromEditor = await OfficeParser.parseOffice(Buffer.from(editorHtml), { fileType: 'html' } as any);
    assert.strictEqual(String((await OfficeGenerator.generate(fromEditor, 'md')).value), md.trimEnd(), 'Comments: editor-shape round trip is byte-for-byte');

    // Default HTML carries real comments, which re-parse under preserveComments.
    const plainHtml = String((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: false } } as any)).value);
    assert.ok(plainHtml.includes('<!-- a hidden note -->'), 'Comments: default HTML emits a real comment');
    const reparsed = await OfficeParser.parseOffice(Buffer.from(plainHtml), { fileType: 'html', htmlParserConfig: { preserveComments: true } } as any);
    assert.strictEqual(String((await OfficeGenerator.generate(reparsed, 'md')).value), md.trimEnd(), 'Comments: default HTML round trip is byte-for-byte under preserveComments');

    // Every other format has no hidden-comment construct: the note must not appear at all.
    const leaks = (s: string) => s.includes('hidden note') || s.includes('inline "q"') || s.includes('inline &quot;q');
    for (const format of ['text', 'csv', 'rtf', 'chunks'] as const) {
        const value = (await OfficeGenerator.generate(ast, format as any)).value;
        assert.ok(!leaks(JSON.stringify(value)), `Comments: ${format} output omits the hidden note`);
    }
    for (const format of ['docx', 'odt', 'epub'] as const) {
        const bytes = (await OfficeGenerator.generate(ast, format as any)).value as Uint8Array;
        const text = Object.values(unzipSync(bytes)).map(b => strFromU8(b)).join('\n');
        assert.ok(text.includes('First paragraph'), `Comments: ${format} output has the body (positive control)`);
        assert.ok(!leaks(text), `Comments: ${format} package omits the hidden note`);
    }
    const pdfText = String((await (await OfficeParser.parseOffice(Buffer.from((await OfficeGenerator.generate(ast, 'pdf', { pdfConfig: { engine: 'native' } } as any)).value as Uint8Array), { fileType: 'pdf' } as any)).to('text')).value);
    assert.ok(pdfText.includes('First paragraph'), 'Comments: pdf has the body (positive control)');
    assert.ok(!leaks(pdfText), 'Comments: pdf omits the hidden note');

    // The AST handed to generate() is never mutated by the strip.
    assert.ok(collectAllNodes(ast).some(n => n.type === 'comment'), 'Comments: source AST still holds its comments after non-md generation');

    // A hidden note inside a heading, a list item or a footnote is not part of that node's text either,
    // so formats that read `node.text` (chunks, text) never show it.
    const nested = await OfficeParser.parseOffice(Buffer.from('# Head <!-- secret1 --> tail\n\n- item <!-- secret2 --> x\n\nRef[^1].\n\n[^1]: Note <!-- secret3 --> body.\n'), { fileType: 'md' } as any);
    assert.ok(collectAllNodes(nested).filter(n => n.type !== 'comment').every(n => !/secret/.test(n.text || '')), 'Comments: no node text carries a nested hidden note');
    for (const format of ['text', 'chunks', 'csv'] as const) {
        const value = JSON.stringify((await OfficeGenerator.generate(nested, format as any)).value);
        assert.ok(!/secret/.test(value), `Comments: ${format} omits notes nested in headings, items and footnotes`);
    }

    // LaTeX has a hidden-comment construct too: `% <!--body-->`, one `%` line per line of the body, which
    // the LaTeX parser restores, so md -> tex -> md keeps every comment and the typeset text never shows it.
    const tex = String((await OfficeGenerator.generate(ast, 'tex')).value);
    assert.ok(tex.includes('\n\n% <!-- a hidden note -->\n\n') && tex.includes('Hello % <!-- inline "q" & <b> -->\n world.'), 'Comments: tex writes source comments as % lines');
    const fromTex = await OfficeParser.parseOffice(Buffer.from(tex), { fileType: 'tex' } as any);
    assert.strictEqual(String((await OfficeGenerator.generate(fromTex, 'md')).value), 'First paragraph.\n\n<!-- a hidden note -->\n\nHello <!-- inline "q" & <b> -->world.',
        'Comments: md -> tex -> md keeps both comments (the space after an inline one moves before it, where TeX keeps it)');
    assert.strictEqual(String((await fromTex.to('text')).value).trim(), 'First paragraph.\nHello world.', 'Comments: tex text shows no note and keeps the word space');
    const multi = 'Before.\n\n<!-- line one\n\n  indented, trailing   \n-->\n\nAfter.';
    const multiTex = String((await OfficeGenerator.generate(await OfficeParser.parseOffice(Buffer.from(multi), { fileType: 'md' } as any), 'tex')).value);
    assert.ok(multiTex.includes('\n% <!-- line one\n%\n%   indented, trailing   \n% -->\n'), 'Comments: a multi-line comment is one % line per line, whitespace kept');
    assert.strictEqual(String((await (await OfficeParser.parseOffice(Buffer.from(multiTex), { fileType: 'tex' } as any)).to('md')).value), multi, 'Comments: multi-line comment survives tex byte-for-byte');

    // A comment line ending a heading, item or cell is followed by `{}`: callers trim a run, and the `%`
    // would otherwise swallow the closing brace, the next \item or the row's `\\`. Review comments too.
    const edges = await OfficeParser.parseOffice(Buffer.from('# Head <!-- h -->\n\n- item <!-- i -->\n- next\n\n| a | b <!-- c --> |\n|---|---|\n| d | e |\n'), { fileType: 'md' } as any);
    edges.content.push({ type: 'heading', text: 'Reviewed', metadata: { level: 2 }, children: [{ type: 'text', text: 'Reviewed', comments: [{ type: 'comment', text: 'check', metadata: {}, children: [] }] }] } as any);
    const edgesTex = String((await OfficeGenerator.generate(edges, 'tex')).value);
    assert.ok(edgesTex.includes('Head % <!-- h -->\n{}}') && edgesTex.includes('\\item item % <!-- i -->\n{}\n\\item next')
        && edgesTex.includes('% <!-- c -->\n{} \\\\') && edgesTex.includes('Reviewed% Comment: check\n{}}'), 'Comments: a trailing comment line is guarded by {}');
    const edgesBack = await OfficeParser.parseOffice(Buffer.from(edgesTex), { fileType: 'tex' } as any);
    const edgesTable = collectAllNodes(edgesBack).find(n => n.type === 'table')!;
    assert.deepStrictEqual(edgesTable.children!.map(r => r.children!.length), [2, 2], 'Comments: a comment ending a cell keeps the table intact');
    assert.strictEqual(collectAllNodes(edgesBack).filter(n => n.type === 'list').length, 2, 'Comments: a comment ending an item keeps the next item');

    // Comment edge cases in Markdown: `<!-->` and `<!--->` are complete empty comments (HTML, CommonMark),
    // a comment may open on one line of a paragraph and close on the next, and a comment inside a code
    // span stays code.
    const mdText = async (src: string) => String((await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).to('text')).value).trim();
    assert.strictEqual(await mdText('<!-->\n\ntext\n\nmore -->'), 'text\nmore -->', 'Comments: <!--> is an empty comment, not an opener');
    assert.strictEqual(await mdText('a <!---> b'), 'a  b', 'Comments: <!---> is an empty comment');
    assert.strictEqual(await mdText('Some text <!-- multi\nline --> more text'), 'Some text  more text', 'Comments: a comment spanning paragraph lines is hidden');
    const spanned = await OfficeParser.parseOffice(Buffer.from('Some text <!-- multi\nline --> more text'), { fileType: 'md' } as any);
    assert.strictEqual(String((await spanned.to('md')).value), 'Some text <!-- multi\nline --> more text', 'Comments: a comment spanning paragraph lines round-trips byte-for-byte');
    assert.strictEqual(await mdText('Literal `<!-- in code -->` span.'), 'Literal <!-- in code --> span.', 'Comments: a comment inside a code span stays code');

    // HTML: an inline comment keeps the words around it together, and only an EMPTY data-html-comment
    // element (the generator's shape) is a comment; one with content keeps its content.
    const inlineHtml = String((await (await OfficeParser.parseOffice(Buffer.from('word<!-- c -->joined'), { fileType: 'md' } as any)).to('html', { htmlConfig: { standalone: false } } as any)).value);
    assert.ok(inlineHtml.includes('<p>word<!-- c -->joined</p>'), 'Comments: no blank line after an inline comment in HTML');
    const withContent = await OfficeParser.parseOffice(Buffer.from('<div data-html-comment="x">visible?</div>'), { fileType: 'html' } as any);
    assert.strictEqual(String((await withContent.to('text')).value).trim(), 'visible?', 'Comments: a data-html-comment element with content keeps it');

    // Hand-written LaTeX: `% <!-- ... -->` is read as a source comment; a `<!--` never closed on the
    // following comment lines is an ordinary comment (discarded) and takes no text with it.
    const hand = await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\begin{document}\nOne.\n\n% <!-- kept -->\n\nTwo.\n% <!-- never closed\nThree.\n\\end{document}'), { fileType: 'tex' } as any);
    assert.deepStrictEqual(hand.content.map(n => [n.type, n.text]), [['paragraph', 'One.'], ['comment', ' kept '], ['paragraph', 'Two. Three.']], 'Comments: hand-written LaTeX % <!-- --> lines');

    console.log('  Source comments: All assertions passed ✓');
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

    await testRtfDestinations();
    await testRtfTables();
    await testRtfEncodings();

    console.log('  RTF: All assertions passed ✓');
}

/**
 * RTF text in the code page it is written in: the document's (`\ansicpg`, double-byte ones too), a
 * font's (`\fcharset`, for its group), and TextEdit's lines (a backslash ending each) and font table.
 */
async function testRtfEncodings(): Promise<void> {
    const texts = async (rtf: string) => (await OfficeParser.parseOffice(Buffer.from(rtf, 'latin1'), { fileType: 'rtf' } as any)).content.map(n => n.text);
    assert.deepStrictEqual(await texts(String.raw`{\rtf1\ansi\ansicpg932 \'93\'fa\'96\'7b\'8c\'ea\par}`), ['日本語'], 'RTF: \\ansicpg932 is Shift-JIS');
    assert.deepStrictEqual(await texts(String.raw`{\rtf1\ansi\ansicpg936 \'d6\'d0\'ce\'c4\par}`), ['中文'], 'RTF: \\ansicpg936 is GBK');
    assert.deepStrictEqual(await texts(String.raw`{\rtf1\ansi\ansicpg950 \'a4\'a4\'a4\'e5\par}`), ['中文'], 'RTF: \\ansicpg950 is Big5');
    assert.deepStrictEqual(await texts(String.raw`{\rtf1\ansi\ansicpg949 \'c7\'d1\'b1\'db\par}`), ['한글'], 'RTF: \\ansicpg949 is Korean');
    // A font's character set is its text's code page, for the font's group only.
    assert.deepStrictEqual(await texts(String.raw`{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Arial;}{\f1\fcharset204 Arial Cyr;}{\f2\fcharset128 Mincho;}}\f1 \'cf\'f0\'e8 {\f2 \'93\'fa\'96\'7b}\f0  \'e9\par}`), ['При 日本 é'], 'RTF: \\fcharset204 and \\fcharset128 fonts are read in their code pages, and the document\'s after');
    // TextEdit: a backslash ending a line ends the paragraph, and fonts are listed without groups.
    const cocoa = await OfficeParser.parseOffice(Buffer.from('{\\rtf1\\ansi\\ansicpg1252\\cocoartf2907\n{\\fonttbl\\f0\\froman\\fcharset0 Times-Bold;\\f1\\froman\\fcharset0 Times-Roman;}\n{\\colortbl;;\\red0\\green0\\blue233;}\n\\f0\\b First line\\\n\\f1\\b0\\cf2 Second line\\\r\nThird\\\n}', 'latin1'), { fileType: 'rtf' } as any);
    assert.deepStrictEqual(cocoa.content.map(n => [n.text, n.children?.[0]?.formatting?.font, n.children?.[0]?.formatting?.color]), [['First line', 'Times-Bold', undefined], ['Second line', 'Times-Roman', '#0000e9'], ['Third', 'Times-Roman', '#0000e9']], 'RTF: TextEdit lines are paragraphs, its fonts named, and ";;" in the colour table is two colours');

    // The writer: a tab is \tab, a line end or vertical tab \line, a form feed \page; no control
    // character is written raw (a tab stood in the text, and a line end, which readers ignore, joined words).
    const controls = await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [
        { type: 'paragraph', children: [{ type: 'text', text: 'a\tb c\nd e\r\nf g\u000bh i\u0001j' }, { type: 'text', text: 'bold\ttab', formatting: { bold: true } }] },
    ] } as any, 'rtf');
    const rtfControls = controls.value as string;
    assert.ok(!/[\t\v\f\x00-\x08\x0e-\x1f]/.test(rtfControls), `RTF: no control character is written raw (${JSON.stringify(rtfControls.slice(-160))})`);
    const controlsBack = await OfficeParser.parseOffice(Buffer.from(rtfControls, 'latin1'), { fileType: 'rtf' } as any);
    assert.deepStrictEqual(controlsBack.content.map(n => n.text), ['a\tb c\nd e\nf g\nh ijbold\ttab'], 'RTF: a tab and line ends read back');
}

/**
 * RTF tables as Word writes them: nested tables (`\itapN`, `\nestcell`, `\nesttableprops ... \nestrow`,
 * `\nonesttables`), merged cells (`\clvmgf`/`\clvmrg`, `\clmgf`/`\clmrg`), and the writer's side of both,
 * with colours in any CSS form.
 */
async function testRtfTables(): Promise<void> {
    const parse = async (rtf: string) => OfficeParser.parseOffice(Buffer.from(rtf, 'latin1'), { fileType: 'rtf' } as any);
    const shape = (n: OfficeContentNode): any => n.type === 'table' || n.type === 'row'
        ? [n.type, (n.children ?? []).map(shape)]
        : n.type === 'cell'
            ? ['cell', (n.metadata as any)?.col, (n.metadata as any)?.rowSpan ?? 1, (n.metadata as any)?.colSpan ?? 1, (n.children ?? []).map(shape)]
            : [n.type, n.text];

    // Word's nested table: the inner rows stay the inner table's (they were read as the outer table's).
    const nested = await parse(String.raw`{\rtf1\ansi\trowd\cellx3000\cellx6000\pard\intbl\itap2 One\par\pard\intbl\itap2 Three\nestcell{\nonesttables\par}\pard\intbl\itap2 Two\nestcell{\nonesttables\par}\pard\intbl\itap2 {\*\nesttableprops\trowd\clvmgf\cellx1300\cellx2600\nestrow}{\nonesttables\par}\pard\intbl\itap2 \nestcell{\nonesttables\par}Four\nestcell{\nonesttables\par}\pard\intbl\itap2 {\*\nesttableprops\trowd\clvmrg\cellx1300\cellx2600\nestrow}{\nonesttables\par}\trowd\cellx3000\cellx6000\pard\intbl \cell Beside\cell\pard\intbl {\trowd\cellx3000\cellx6000\row}\pard After\par}`);
    assert.deepStrictEqual(nested.content.map(shape), [
        ['table', [['row', [
            ['cell', 0, 1, 1, [['table', [
                ['row', [['cell', 0, 2, 1, [['paragraph', 'One'], ['paragraph', 'Three']]], ['cell', 1, 1, 1, [['paragraph', 'Two']]]]],
                ['row', [['cell', 1, 1, 1, [['paragraph', 'Four']]]]],
            ]]]],
            ['cell', 1, 1, 1, [['paragraph', 'Beside']]],
        ]]]],
        ['paragraph', 'After'],
    ], 'RTF: a nested table (\\itap2, \\nestcell, \\nesttableprops) is read nested, its vertical merge a rowSpan');

    // Merges across and down: the continuation cell's content joins the merged cell, later cells keep their columns.
    const merged = await parse(String.raw`{\rtf1\ansi\trowd\clvmgf\cellx2000\clmgf\cellx4000\clmrg\cellx6000\cellx8000\pard\intbl A\cell B\cell B2\cell C\cell\row\trowd\clvmrg\cellx2000\cellx4000\cellx6000\cellx8000\pard\intbl \cell D\cell E\cell F\cell\row\pard After\par}`);
    assert.deepStrictEqual(merged.content.map(shape)[0], ['table', [
        ['row', [['cell', 0, 2, 1, [['paragraph', 'A']]], ['cell', 1, 1, 2, [['paragraph', 'B'], ['paragraph', 'B2']]], ['cell', 3, 1, 1, [['paragraph', 'C']]]]],
        ['row', [['cell', 1, 1, 1, [['paragraph', 'D']]], ['cell', 2, 1, 1, [['paragraph', 'E']]], ['cell', 3, 1, 1, [['paragraph', 'F']]]]],
    ]], 'RTF: \\clvmrg lengthens the cell above (rowSpan) and \\clmrg widens the one before (colSpan), cells keeping their columns');
    // Word writes the row definition after the row's cells too (in {\trowd ... \row}): merges still apply.
    const trailing = await parse(String.raw`{\rtf1\ansi\pard\intbl A\cell B\cell\pard\intbl{\trowd\clvmgf\cellx2000\cellx4000\row}\pard\intbl \cell C\cell\pard\intbl{\trowd\clvmrg\cellx2000\cellx4000\row}\pard After\par}`);
    assert.deepStrictEqual(trailing.content.map(shape)[0], ['table', [
        ['row', [['cell', 0, 2, 1, [['paragraph', 'A']]], ['cell', 1, 1, 1, [['paragraph', 'B']]]]],
        ['row', [['cell', 1, 1, 1, [['paragraph', 'C']]]]],
    ]], 'RTF: a row definition after the cells applies its merges');

    // The writer: a nested table as Word writes one, read back nested; colours in any CSS form.
    const T = (text: string, extra: any = {}) => ({ type: 'text', text, ...extra }) as OfficeContentNode;
    const P = (...children: OfficeContentNode[]) => ({ type: 'paragraph', children }) as OfficeContentNode;
    const cell = (...children: OfficeContentNode[]) => ({ type: 'cell', children }) as OfficeContentNode;
    const inner = { type: 'table', children: [{ type: 'row', children: [cell(P(T('inner 1'))), cell(P(T('inner 2')))] }] } as OfficeContentNode;
    const written = await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [
        P(T('red', { formatting: { color: 'red' } }), T(' green', { formatting: { color: '#0f0', backgroundColor: 'navy' } }), T(' blue', { formatting: { color: 'rgb(0, 0, 255)', backgroundColor: 'transparent' } })),
        { type: 'table', children: [
            { type: 'row', children: [cell(P(T('outer A')), inner), cell(P(T('outer B')))] },
            { type: 'row', children: [cell(P(T('outer C'))), cell(P(T('outer D')))] },
        ] } as OfficeContentNode,
        P(T('After')),
    ] } as any, 'rtf');
    const rtf = written.value as string;
    assert.ok(!/NaN/.test(rtf) && rtf.includes('{\\colortbl;\\red255\\green0\\blue0;\\red0\\green255\\blue0;\\red0\\green0\\blue128;\\red0\\green0\\blue255;}'), `RTF: named, three-digit and rgb() colours are written as colours (${rtf.slice(0, 300)})`);
    assert.ok(rtf.includes('\\nestcell') && rtf.includes('\\nesttableprops') && rtf.includes('\\itap2') && rtf.includes('{\\nonesttables\\par}'), 'RTF: a nested table is written with \\itap2, \\nestcell and \\nesttableprops');
    const back = await parse(rtf);
    assert.deepStrictEqual(back.content.map(shape).slice(1), [
        ['table', [
            ['row', [['cell', 0, 1, 1, [['paragraph', 'outer A'], ['table', [['row', [['cell', 0, 1, 1, [['paragraph', 'inner 1']]], ['cell', 1, 1, 1, [['paragraph', 'inner 2']]]]]]]]], ['cell', 1, 1, 1, [['paragraph', 'outer B']]]]],
            ['row', [['cell', 0, 1, 1, [['paragraph', 'outer C']]], ['cell', 1, 1, 1, [['paragraph', 'outer D']]]]],
        ]],
        ['paragraph', 'After'],
    ], 'RTF: a nested table round-trips nested, the outer table whole');
    assert.deepStrictEqual((back.content[0].children ?? []).map(n => [n.text, n.formatting?.color, n.formatting?.backgroundColor]), [['red', '#ff0000', undefined], [' green', '#00ff00', '#000080'], [' blue', '#0000ff', undefined]], 'RTF: the colours read back');
}

/**
 * A destination read apart from the body (an annotation, a header or footer, a footnote, a shape's text
 * box) keeps the paragraph and table it sits in as they were, and takes nothing of the body. Word writes
 * `\pard\plain` inside each, which ended the list item, heading or table cell around it, and a table left
 * open was flushed into the destination.
 */
async function testRtfDestinations(): Promise<void> {
    const head = String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0 Calibri;}}`;
    const lists = String.raw`{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc23{\leveltext\'01舦 ?;}}\listid1}}{\*\listoverridetable{\listoverride\listid1\listoverridecount0\ls1}}`;
    const parse = async (body: string) => OfficeParser.parseOffice(Buffer.from(head + body + '}', 'latin1'), { fileType: 'rtf' } as any);
    const shape = (n: OfficeContentNode): any => n.type === 'table' || n.type === 'row' || n.type === 'cell'
        ? [n.type, (n.children ?? []).map(shape)]
        : [n.type + ((n.metadata as any)?.level ?? '') + ((n.metadata as any)?.listId ? `#${(n.metadata as any).itemIndex}` : ''), n.text];
    const table2x2 = [['table', [['row', [['cell', [['paragraph', 'A']]], ['cell', [['paragraph', 'B']]]]], ['row', [['cell', [['paragraph', 'C text']]], ['cell', [['paragraph', 'D']]]]]]]];
    const comments = (ast: OfficeParserAST) => collectAllNodes(ast).filter(n => n.type === 'comment').map(c => (c.children ?? []).map(shape));

    // A comment in a table cell, as Word writes it.
    const inCell = await parse(String.raw`\trowd\cellx3000\cellx6000\pard\plain\intbl A\cell B\cell\row\trowd\cellx3000\cellx6000\pard\plain\intbl {\*\atrfstart 0}C text{\*\atrfend 0}{\*\atnid JD}{\*\atnauthor John Doe}\chatn {\*\annotation{\*\atndate 1}\pard\plain \s16\ql {\chatn }{Comment text}}\cell D\cell\row\pard\plain After table\par`);
    assert.deepStrictEqual(inCell.content.map(shape), [...table2x2, ['paragraph', 'After table']], 'RTF: a comment in a table cell leaves the table in the body');
    assert.deepStrictEqual(comments(inCell), [[['paragraph', 'Comment text']]], 'RTF: the comment holds its own text only');

    // A comment on a list item or a heading.
    const onItem = await parse(lists + String.raw`\pard\ls1\ilvl0 Item one{\*\atnid JD}{\*\atnauthor J}\chatn {\*\annotation\pard\plain {Remark}}\par\pard\ls1\ilvl0 Item two\par`);
    assert.deepStrictEqual(onItem.content.map(shape), [['list#0', 'Item one'], ['list#1', 'Item two']], 'RTF: a comment on a list item keeps it an item');
    const onHeading = await parse(String.raw`\pard\s1 Title{\*\atnid JD}\chatn {\*\annotation\pard\plain {Remark}}\par\pard Body\par`);
    assert.deepStrictEqual(onHeading.content.map(shape), [['heading1', 'Title'], ['paragraph', 'Body']], 'RTF: a comment on a heading keeps it a heading');

    // A table ending a section, before the next section's header.
    const sections = await parse(String.raw`\sectd{\header \pard Head1\par}\pard Intro\par\trowd\cellx3000\cellx6000\pard\intbl A\cell B\cell\row\pard\par\sect\sectd{\header \pard Head2\par}\pard After\par`);
    assert.deepStrictEqual(sections.content.map(shape), [['paragraph', 'Intro'], ['table', [['row', [['cell', [['paragraph', 'A']]], ['cell', [['paragraph', 'B']]]]]]], ['paragraph', 'After']], 'RTF: a table ending a section stays in the body');
    assert.deepStrictEqual(sections.auxiliary?.headers?.map(shape), [['paragraph', 'Head1'], ['paragraph', 'Head2']], 'RTF: the headers hold their own paragraphs only');
    // A header's own table, and a footer written inside a table cell.
    const headerTable = await parse(String.raw`\trowd\cellx3000\pard\intbl A\cell\row{\header \trowd\cellx3000\pard\intbl H\cell\row\pard Hx\par}\pard After\par`);
    assert.deepStrictEqual(headerTable.content.map(shape), [['table', [['row', [['cell', [['paragraph', 'A']]]]]]], ['paragraph', 'After']], 'RTF: a header with a table leaves the body\'s table in the body');
    assert.deepStrictEqual(headerTable.auxiliary?.headers?.map(shape), [['table', [['row', [['cell', [['paragraph', 'H']]]]]]], ['paragraph', 'Hx']], 'RTF: a header keeps its own table');
    const footerInCell = await parse(String.raw`\trowd\cellx3000\cellx6000\pard\intbl A{\footer \pard Foot\par}\cell B\cell\row\pard After\par`);
    assert.deepStrictEqual(footerInCell.content.map(shape), [['table', [['row', [['cell', [['paragraph', 'A']]], ['cell', [['paragraph', 'B']]]]]]], ['paragraph', 'After']], 'RTF: a footer written in a cell leaves the cell as it was');

    // Footnotes in a table cell and in a list item (8.0.0 broke both the same way).
    const noteInCell = await parse(String.raw`\trowd\cellx3000\cellx6000\pard\intbl A{\super\chftn}{\footnote\pard\plain {\super\chftn} Note text\par}\cell B\cell\row\pard After\par`);
    assert.deepStrictEqual(noteInCell.content.map(shape), [['table', [['row', [['cell', [['paragraph', 'A']]], ['cell', [['paragraph', 'B']]]]]]], ['paragraph', 'After']], 'RTF: a footnote in a table cell leaves the table whole');
    const noteOnItem = await parse(lists + String.raw`\pard\ls1\ilvl0 Item one{\footnote\pard\plain Note text\par}\par\pard\ls1\ilvl0 Item two\par`);
    assert.deepStrictEqual(noteOnItem.content.map(shape), [['list#0', 'Item one'], ['list#1', 'Item two']], 'RTF: a footnote on a list item keeps it an item');
    assert.deepStrictEqual(collectAllNodes(noteOnItem).filter(n => n.type === 'note').map(n => n.text), ['Note text'], 'RTF: the footnote keeps its text');

    // A shape's text box follows the paragraph anchoring it, which stays whole.
    const shapeOnItem = await parse(lists + String.raw`\pard\ls1\ilvl0 Item one{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt \pard\plain Box words\par}}} more\par\pard\ls1\ilvl0 Item two\par`);
    assert.deepStrictEqual(shapeOnItem.content.map(shape), [['list#0', 'Item one more'], ['paragraph', 'Box words'], ['list#1', 'Item two']], 'RTF: a text box follows the list item anchoring it');
    const shapeInCell = await parse(String.raw`\trowd\cellx3000\pard\intbl A{\shp{\*\shpinst{\shptxt \pard\plain Box\par}}}\cell\row\pard After\par`);
    assert.deepStrictEqual(shapeInCell.content.map(shape), [['table', [['row', [['cell', [['paragraph', 'A'], ['paragraph', 'Box']]]]]]], ['paragraph', 'After']], 'RTF: a text box in a cell stays in the cell');

    // A comment inside a link is not the link; a paragraph after a table's last row without its own \par follows the table.
    const inLink = await parse(String.raw`\pard {\field{\*\fldinst HYPERLINK "https://x.test"}{\fldrslt Link{\*\atnid JD}\chatn{\*\annotation\pard Remark\par}}} tail\par`);
    const linked = collectAllNodes(inLink).filter(n => (n.metadata as any)?.link).map(n => n.text);
    assert.deepStrictEqual(linked, ['Link'], 'RTF: a comment inside a link is not linked');
    const tailAfterTable = await parse(String.raw`\trowd\cellx3000\pard\intbl A\cell\row\pard Tail`);
    assert.deepStrictEqual(tailAfterTable.content.map(shape), [['table', [['row', [['cell', [['paragraph', 'A']]]]]]], ['paragraph', 'Tail']], 'RTF: the last paragraph after a table follows it');

    // The writer keeps headers, footers and review comments (they were left out, with no message): an
    // RTF round trip reads them back, the comment with its author, initials and date.
    const running = await parse(String.raw`{\header \pard Running head\par}{\footer \pard Page foot\par}\pard Body text{\*\atnid JD}{\*\atnauthor John}\chatn{\*\annotation{\*\atndate 1742938438}\pard\plain Review remark\par} more\par`);
    const written = await OfficeGenerator.generate(running, 'rtf');
    assert.deepStrictEqual(written.messages, [], 'RTF: headers, footers and comments are written without a message');
    const rtfBack = await OfficeParser.parseOffice(Buffer.from(written.value as string, 'latin1'), { fileType: 'rtf' } as any);
    assert.deepStrictEqual([rtfBack.auxiliary?.headers?.map(n => n.text), rtfBack.auxiliary?.footers?.map(n => n.text)], [['Running head'], ['Page foot']], 'RTF: headers and footers are written');
    assert.deepStrictEqual(rtfBack.content.map(shape), [['paragraph', 'Body text more']], 'RTF: the body keeps only its own text');
    const back = collectAllNodes(rtfBack).filter(n => n.type === 'comment');
    assert.deepStrictEqual(back.map(c => [c.text, c.metadata]), [['Review remark', { author: 'John', initials: 'JD', date: '2026-03-04T05:06:00' }]], 'RTF: a comment is written with its author, initials and date');
    // A comment on a table cell and one on a table, written inside a table: annotations carry no \intbl
    // of the table they are in, and the table stays whole.
    const T = (text: string, extra: any = {}) => ({ type: 'text', text, ...extra }) as OfficeContentNode;
    const note = (text: string) => ({ type: 'comment', text, metadata: { author: 'A' }, children: [{ type: 'paragraph', text, children: [T(text)] }] }) as OfficeContentNode;
    const cellComments = await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [
        { type: 'table', comments: [note('On table')], children: [{ type: 'row', children: [{ type: 'cell', comments: [note('On cell')], children: [{ type: 'paragraph', children: [T('Cell', { comments: [note('On run')] })] }] }, { type: 'cell', children: [{ type: 'paragraph', children: [T('Other')] }] }] }] },
        { type: 'paragraph', children: [T('After')] },
    ] } as any, 'rtf');
    assert.ok(!/\\annotation[^}]*\\intbl/.test(cellComments.value as string), 'RTF: an annotation in a table carries no \\intbl');
    const cellBack = await OfficeParser.parseOffice(Buffer.from(cellComments.value as string, 'latin1'), { fileType: 'rtf' } as any);
    assert.deepStrictEqual(cellBack.content.map(shape).filter(b => b[0] === 'table' || b[1]), [['table', [['row', [['cell', [['paragraph', 'Cell']]], ['cell', [['paragraph', 'Other']]]]]]], ['paragraph', 'After']], 'RTF: comments in a table leave it whole');
    assert.deepStrictEqual(collectAllNodes(cellBack).filter(n => n.type === 'comment').map(c => c.text).sort(), ['On cell', 'On run', 'On table'], 'RTF: comments on a table, a cell and a run are all written');
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

    // A node of a type the AST does not define (a hand-built AST, or one from a newer version) is written
    // as its content by every generator: it made four of them fail and two leave it out. In a row it is
    // a cell, in a table a row, and among blocks its inline content is a paragraph.
    const unknownAst = { type: 'docx', metadata: {}, attachments: [], content: [
        { type: 'paragraph', children: [{ type: 'text', text: 'before ' }, { type: 'mark', children: [{ type: 'text', text: 'inline unknown', formatting: { bold: true } }] }] },
        { type: 'callout', children: [{ type: 'text', text: 'loose' }, { type: 'paragraph', children: [{ type: 'text', text: 'inside unknown' }] }] }, { type: 'widget', text: 'text only' },
        { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'c1' }] }, { type: 'fancyCell', children: [{ type: 'text', text: 'cell unknown' }] }] }, { type: 'tbody', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'text', text: 'row in tbody' }] }, { type: 'cell', children: [] }] }] }] },
    ] } as any;
    const unknownBefore = JSON.stringify(unknownAst);
    for (const format of ['md', 'text', 'html', 'tex', 'rtf', 'docx', 'odt', 'epub', 'chunks', 'csv'] as const) {
        const out = await OfficeGenerator.generate(unknownAst, format);
        const value: any = out.value;
        const written = typeof value === 'string' ? value : Array.isArray(value) ? JSON.stringify(value)
            : Object.entries(unzipSync(new Uint8Array(value))).filter(([name]) => /\.(xml|xhtml|html)$/.test(name)).map(([, data]) => strFromU8(data)).join('');
        const expected = format === 'csv' ? ['cell unknown', 'row in tbody'] : ['inline unknown', 'loose', 'inside unknown', 'text only', 'cell unknown', 'row in tbody'];
        assert.deepStrictEqual(expected.filter(t => !written.includes(t)), [], `Generated output: ${format} writes the content of unknown node types`);
    }
    assert.strictEqual(JSON.stringify(unknownAst), unknownBefore, 'Generated output: the AST is not changed');
    assert.strictEqual((await OfficeGenerator.generate(unknownAst, 'md')).value, 'before **inline unknown**\n\nloose\n\ninside unknown\n\ntext only\n\n| c1 | cell unknown |\n| --- | --- |\n| row in tbody |  |', 'Generated output: unknown node types in Markdown');

    // Math in a line of text stays in the line in plain text and chunks (a code node, it was taken for a
    // block and broke the line before it: "Let \nx^2 be positive"), and a code block ends its line (it
    // ran into the paragraph after it: "blockAfter").
    for (const [source, type] of [['Let $x^2$ be positive.\n\n```\nblock\n```\nAfter\n', 'md'], ['\\documentclass{article}\\begin{document}Let $x^2$ be positive.\n\n\\begin{verbatim}\nblock\n\\end{verbatim}\nAfter\n\\end{document}', 'tex']] as const) {
        const mathAst = await OfficeParser.parseOffice(Buffer.from(source), { fileType: type });
        assert.strictEqual((await mathAst.to('text')).value, 'Let x^2 be positive.\nblock\nAfter', `Generated output: ${type} inline math and a code block in plain text`);
        const mathChunks = (await mathAst.to('chunks')).value as any[];
        assert.ok(mathChunks.some(c => c.text === 'Let x^2 be positive.'), `Generated output: ${type} inline math in a chunk (${JSON.stringify(mathChunks.map(c => c.text))})`);
    }
    const preThenParagraph = await parseHtml('<pre>code one</pre><p>Para after</p>');
    assert.strictEqual((await preThenParagraph.to('text')).value, 'code one\nPara after', 'Generated output: a code block ends its line in plain text');
    // A note in a table's data row reaches a chunk under the row strategy: rows were rendered to measure
    // them and again to write them, and the first rendering took the note as written.
    for (const row of ['| h1 | h2 |\n|---|---|\n| alpha | 42[^n] |', '| h1 | h2[^n] |\n|---|---|\n| alpha | 42 |']) {
        const noteChunks = (await (await parseMd(`${row}\n\n[^n]: The footnote text.\n`)).to('chunks')).value as any[];
        assert.ok(noteChunks.some(c => c.text.includes('The footnote text.')), `Generated output: a table note reaches a chunk (${row.split('\n')[0]})`);
    }

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

/** PowerPoint comments, from the OOXML review before 8.1.0. */
async function testPptxComments(): Promise<void> {
    const { relsOf, rel, parse, texts } = ooxmlHelpers();
    // PowerPoint comments: a classic comment's plain text, a modern comment's paragraphs apart, and
    // replies with their date, initials and the comment they reply to.
    const pns = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"';
    const pptx = Buffer.from(zipSync({
        'ppt/presentation.xml': strToU8(`<?xml version="1.0"?><p:presentation ${pns}/>`),
        'ppt/slides/slide1.xml': strToU8(`<?xml version="1.0"?><p:sld ${pns}><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>S1</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`),
        'ppt/slides/_rels/slide1.xml.rels': strToU8(relsOf(rel('rId1', 'comments', '../comments/comment1.xml') + '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/modernComment_100_1.xml"/>')),
        'ppt/commentAuthors.xml': strToU8(`<?xml version="1.0"?><p:cmAuthorLst ${pns}><p:cmAuthor id="0" name="Old Author" initials="OA" lastIdx="2" clrIdx="0"/><p:cmAuthor id="1" name="Old Replier" initials="OR" lastIdx="1" clrIdx="1"/></p:cmAuthorLst>`),
        'ppt/comments/comment1.xml': strToU8(`<?xml version="1.0"?><p:cmLst ${pns}><p:cm authorId="0" dt="2020-01-01T00:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>CLASSIC LINE ONE\nCLASSIC LINE TWO</p:text></p:cm><p:cm authorId="1" dt="2020-01-02T00:00:00.000" idx="1"><p:pos x="10" y="10"/><p:text>CLASSIC REPLY</p:text><p:extLst><p:ext uri="{C676402C-5697-4E1C-873F-D02D1690AC5C}"><p15:threadingInfo timeZoneBias="0"><p15:parentCm authorId="0" idx="1"/></p15:threadingInfo></p:ext></p:extLst></p:cm></p:cmLst>`),
        'ppt/authors.xml': strToU8(`<?xml version="1.0"?><p188:authorLst ${pns}><p188:author id="{A1}" name="Modern Ann" initials="MA" userId="ann" providerId="None"/><p188:author id="{A2}" name="Replier Bob" initials="RB" userId="bob" providerId="None"/></p188:authorLst>`),
        'ppt/comments/modernComment_100_1.xml': strToU8(`<?xml version="1.0"?><p188:cmLst ${pns}><p188:cm id="{C1}" authorId="{A1}" created="2021-01-01T00:00:00.000"><p188:replyLst><p188:reply id="{R1}" authorId="{A2}" created="2021-01-02T00:00:00.000"><p188:txBody><a:bodyPr/><a:p><a:r><a:t>MODERN REPLY</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:p><a:r><a:t>MODERN ONE</a:t></a:r></a:p><a:p><a:r><a:t>MODERN TWO</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>`),
    }));
    const slideComments = (await parse(pptx, {}, 'pptx')).ast.content[0].comments!;
    assert.deepStrictEqual(slideComments.map(c => [c.text, texts(c.children!).map(t => t[1]), c.metadata]), [
        ['CLASSIC LINE ONE CLASSIC LINE TWO', ['CLASSIC LINE ONE', 'CLASSIC LINE TWO'], { commentId: '0-1', author: 'Old Author', initials: 'OA', date: '2020-01-01T00:00:00.000' }],
        ['CLASSIC REPLY', ['CLASSIC REPLY'], { commentId: '1-1', author: 'Old Replier', initials: 'OR', date: '2020-01-02T00:00:00.000', parentId: '0-1' }],
        ['MODERN ONE MODERN TWO', ['MODERN ONE', 'MODERN TWO'], { commentId: '{C1}', author: 'Modern Ann', initials: 'MA', date: '2021-01-01T00:00:00.000' }],
        ['MODERN REPLY', ['MODERN REPLY'], { commentId: '{R1}', author: 'Replier Bob', initials: 'RB', date: '2021-01-02T00:00:00.000', parentId: '{C1}' }],
    ], 'PPTX: classic and modern comments, their paragraphs, and replies with their date, initials and thread');
    console.log('  PPTX comments: All assertions passed ✓');
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
    // A delimiter of several characters carries over too (it wrote ',' with no word), and one that could
    // start a formula line or end a row is refused with INVALID_CONFIG_VALUE rather than silently.
    const doublePipe = String((await (await OfficeParser.parseOffice(Buffer.from(tableMd), { fileType: 'md', csvDelimiter: '||' })).to('csv')).value);
    assert.ok(/1\|\|2/.test(doublePipe), `a two-character csvDelimiter reaches to('csv'); got ${JSON.stringify(doublePipe)}`);
    const refusedCodes: string[] = [];
    const dash = String((await (await OfficeParser.parseOffice(Buffer.from(tableMd), { fileType: 'md', csvDelimiter: '-' })).to('csv', { onWarning: (issue: any) => refusedCodes.push(issue.code) } as any)).value);
    assert.ok(/1,2/.test(dash) && refusedCodes.includes('INVALID_CONFIG_VALUE'), `a csvDelimiter starting a formula is refused with a warning; got ${JSON.stringify(dash)} ${refusedCodes}`);

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

/**
 * DOCX notes and bookmarks, from the OOXML review before 8.1.0: a note referred to again is one note and
 * NOTEREF fields (read back as the same note), and every block a link can go to has a bookmark.
 */
async function testDocxNotesAndBookmarks(): Promise<void> {
    const { r, p, part, docx, parse } = ooxmlHelpers();
    // A note referred to again is one note; its later references are NOTEREF fields to a bookmark around
    // its first, read back as references to the same note, and saved again the same.
    const fn = '<w:r><w:footnoteReference w:id="1"/></w:r>';
    const en = '<w:r><w:endnoteReference w:id="3"/></w:r>';
    const shared = await parse(docx(`<w:p>${r('Body one')}${fn}</w:p><w:tbl><w:tr><w:tc><w:p>${r('Cell')}${fn}</w:p></w:tc></w:tr></w:tbl><w:p>${r('Body two')}${fn}${en}${en}</w:p>`, {
        'word/footnotes.xml': part('footnotes', `<w:footnote w:id="1">${p('FOOTNOTE')}</w:footnote>`),
        'word/endnotes.xml': part('endnotes', `<w:endnote w:id="3">${p('ENDNOTE')}</w:endnote>`),
    }));
    const saved = docxParts((await shared.ast.to('docx')).value as Uint8Array);
    for (const [name, xml] of Object.entries(saved)) assert.doesNotThrow(() => parseXmlString(xml), `DOCX: ${name} is well-formed with NOTEREF fields`);
    const savedDoc = saved['word/document.xml'];
    assert.deepStrictEqual([(savedDoc.match(/<w:footnoteReference /g) ?? []).length, (savedDoc.match(/<w:endnoteReference /g) ?? []).length, (saved['word/footnotes.xml'].match(/<w:footnote w:id="\d+">/g) ?? []).length, (saved['word/endnotes.xml'].match(/<w:endnote w:id="\d+">/g) ?? []).length], [1, 1, 1, 1], 'DOCX: a note referred to three times is one note and one reference');
    assert.deepStrictEqual([...savedDoc.matchAll(/<w:fldSimple w:instr=" NOTEREF (\w+) \\f \\h "><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:t>(\w+)<\/w:t><\/w:r><\/w:fldSimple>/g)].map(m => [m[1], m[2]]), [['_RefNote1', '1'], ['_RefNote1', '1'], ['_RefNote2', 'i']], 'DOCX: later references are NOTEREF fields showing the note\'s number');
    assert.ok(/<w:bookmarkStart w:id="\d+" w:name="_RefNote1"\/><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference /.test(savedDoc), 'DOCX: the NOTEREF bookmark is around the first reference');
    const reread = await parse(Buffer.from((await shared.ast.to('docx')).value as Uint8Array));
    const reread2 = collectAllNodes(reread.ast).flatMap(n => n.notes ?? []);
    assert.ok(reread2.length === 5 && new Set(reread2.filter(n => (n.metadata as any).noteType === 'footnote')).size === 1 && new Set(reread2).size === 2, 'DOCX: NOTEREF fields read back as references to the same note');
    assert.ok(!JSON.stringify(reread.ast.content).includes('anchorIds'), 'DOCX: a NOTEREF\'s bookmark is not read as an anchor of the document');
    assert.strictEqual(docxParts((await reread.ast.to('docx')).value as Uint8Array)['word/document.xml'], savedDoc, 'DOCX: a document with NOTEREF fields saves again the same');
    // Word's own cross-reference: a complex field; without `\f` the number is text, as Word shows it.
    const field = (instruction: string, result: string) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const wordRefs = await parse(docx(`<w:p>${r('One')}<w:bookmarkStart w:id="0" w:name="_Ref1"/>${fn}<w:bookmarkEnd w:id="0"/>${r(' two')}${field(' NOTEREF _Ref1 \\f \\h  \\* MERGEFORMAT ', '1')}${r(' see note ')}${field(' NOTEREF _Ref1 \\h ', '1')}</w:p>`, {
        'word/footnotes.xml': part('footnotes', `<w:footnote w:id="1">${p('FOOTNOTE')}</w:footnote>`),
    }));
    const wordNotes = wordRefs.ast.content[0].children!.map(n => [n.text, (n.notes ?? []).length]);
    assert.deepStrictEqual([wordNotes, wordRefs.ast.content[0].text, (wordRefs.ast.content[0].metadata as any)?.anchorIds], [[['One', 1], [' two', 1], [' see note ', 0], ['1', 0]], 'One two see note 1', undefined], 'DOCX: a formatted NOTEREF is a reference to its note; one showing the number as text stays text');
    assert.strictEqual(new Set(wordRefs.ast.content[0].children!.flatMap(n => n.notes ?? [])).size, 1, 'DOCX: a NOTEREF and the reference it names share one note');

    // Every block a link can go to has a bookmark (a table, a cell, a picture, an equation, a list item,
    // a note), and a cell still ends in a paragraph.
    const anchored = (id: string) => ({ anchorIds: [id] });
    const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const linked: any = { type: 'tex', metadata: {}, attachments: [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: png }], content: [
        { type: 'paragraph', children: ['tab:main', 'cell:one', 'fig:one', 'eq:one', 'item:one', 'note:one', 'code:inline'].map(id => ({ type: 'text', text: id + ' ', metadata: { link: '#' + id, linkType: 'internal' } })) },
        { type: 'table', metadata: anchored('tab:main'), children: [{ type: 'row', children: [{ type: 'cell', metadata: anchored('cell:one'), children: [{ type: 'image', metadata: { attachmentName: 'missing.png', anchorIds: ['cell:image'] } }] }] }] },
        { type: 'image', metadata: { attachmentName: 'p.png', ...anchored('fig:one') } },
        { type: 'code', text: 'a=b', metadata: { math: 'block', ...anchored('eq:one') } },
        { type: 'list', text: 'item', metadata: { listType: 'unordered', indentation: 0, ...anchored('item:one') }, children: [{ type: 'text', text: 'item' }] },
        { type: 'paragraph', children: [{ type: 'text', text: 'x', notes: [{ type: 'note', metadata: { noteType: 'footnote', ...anchored('note:one') }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'n' }] }] }] }, { type: 'code', text: 'y', metadata: anchored('code:inline') }] },
    ] };
    const linkedParts = docxParts((await OfficeGenerator.generate(linked, 'docx', { onWarning: () => {} } as any)).value as Uint8Array);
    const bookmarkNames = new Set(Object.values(linkedParts).flatMap(xml => [...xml.matchAll(/<w:bookmarkStart w:id="\d+" w:name="([^"]+)"\/>/g)].map(m => m[1])));
    const anchors = [...linkedParts['word/document.xml'].matchAll(/w:anchor="([^"]+)"/g)].map(m => m[1]);
    assert.ok(anchors.length === 7 && anchors.every(a => bookmarkNames.has(a)), `DOCX: every internal link has its bookmark (${anchors.filter(a => !bookmarkNames.has(a))})`);
    for (const [name, xml] of Object.entries(linkedParts)) assert.doesNotThrow(() => parseXmlString(xml), `DOCX: ${name} is well-formed with bookmarks of every block`);
    assert.ok(/<w:tc><w:tcPr>[^]*?<\/w:tcPr><w:bookmarkStart[^>]*w:name="cell_one"\/><w:bookmarkEnd w:id="\d+"\/><w:bookmarkStart[^>]*w:name="cell_image"\/><w:bookmarkEnd w:id="\d+"\/>(?:<w:p\/>|<w:p>[^]*?<\/w:p>)<\/w:tc>/.test(linkedParts['word/document.xml']), 'DOCX: a cell\'s bookmarks come first, and it ends in a paragraph');
    const linkedBack = await parse(Buffer.from((await OfficeGenerator.generate(linked, 'docx', { onWarning: () => {} } as any)).value as Uint8Array));
    assert.deepStrictEqual((linkedBack.ast.content.find(n => n.type === 'table')!.metadata as any)?.anchorIds, ['tab_main'], 'DOCX: a table\'s bookmark reads back as the table\'s');
    console.log('  DOCX notes and bookmarks: All assertions passed ✓');
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

    await testOdtInlineRuns();
    await testOdtNoteReferences();
    await testOdfCellsAndCharts();
    await testOdtBookmarks();
}

/**
 * Every node with anchors is a link target in ODT: a table (in its first cell), a picture, a code
 * block, a list item and a cell (in their paragraph), an inline equation (a bookmark's start and end
 * around it). Only paragraphs and headings wrote theirs.
 */
async function testOdtBookmarks(): Promise<void> {
    const T = (text: string, extra: any = {}) => ({ type: 'text', text, ...extra });
    const ast: any = { type: 'tex', metadata: {}, attachments: [], content: [
        { type: 'paragraph', children: [T('See '), T('the table', { metadata: { link: '#tab:main', linkType: 'internal' } }), T(' and '), { type: 'code', text: 'x^2', metadata: { math: 'inline', anchorIds: ['eq:inline'] } }] },
        { type: 'table', metadata: { anchorIds: ['tab:main'] }, children: [{ type: 'row', children: [{ type: 'cell', metadata: { anchorIds: ['cell:first'] }, children: [{ type: 'paragraph', children: [T('A')] }] }] }] },
        { type: 'image', metadata: { url: 'https://example.com/p.png', altText: 'pic', anchorIds: ['fig:a'] } },
        { type: 'code', text: 'let x = 1;', metadata: { anchorIds: ['lst:code'] } },
        { type: 'list', text: 'Item', metadata: { listType: 'ordered', listId: 'l', indentation: 0, itemIndex: 0, anchorIds: ['item:one'] }, children: [T('Item')] },
    ] };
    const content = strFromU8(unzipSync((await OfficeGenerator.generate(ast, 'odt')).value as Uint8Array)['content.xml']);
    assert.doesNotThrow(() => parseXmlString(content), 'ODT bookmarks: content.xml is well-formed');
    const marks = [...content.matchAll(/<text:bookmark(?:-start)? text:name="([^"]+)"/g)].map(m => m[1]).sort();
    assert.deepStrictEqual(marks, ['cell_first', 'eq_inline', 'fig_a', 'item_one', 'lst_code', 'tab_main'], 'ODT: every node with anchors has a bookmark');
    assert.ok(content.includes('<text:bookmark-end text:name="eq_inline"/>'), 'ODT: an inline node\'s bookmark has an end');
    assert.ok(/<table:table-cell[^>]*><text:p><text:bookmark text:name="tab_main"\/><text:bookmark text:name="cell_first"\/>A<\/text:p>/.test(content), `ODT: a table's and a cell's bookmarks are in the first cell's paragraph (${content.match(/<table:table-cell.*?<\/table:table-cell>/)?.[0]})`);
    assert.ok(content.includes('xlink:href="#tab_main"'), 'ODT: the link to the table points at its bookmark');
}

/**
 * ODF reading: merged cells keep the cells after them in their columns, a cell's paragraphs and a line
 * break are lines, and a chart LibreOffice writes (its namespaces past the first 500 bytes, its data
 * rows in `table:table-rows`) has its data.
 */
async function testOdfCellsAndCharts(): Promise<void> {
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"';
    const pkg = (kind: 'text' | 'spreadsheet', body: string, extra: Record<string, string> = {}) => Buffer.from(zipSync({
        mimetype: strToU8(`application/vnd.oasis.opendocument.${kind}`),
        'content.xml': strToU8(`<?xml version="1.0"?><office:document-content ${NS}><office:body><office:${kind}>${body}</office:${kind}></office:body></office:document-content>`),
        ...Object.fromEntries(Object.entries(extra).map(([k, v]) => [k, strToU8(v)])),
    }));
    const c = (t: string, attrs = '') => `<table:table-cell office:value-type="string"${attrs}><text:p>${t}</text:p></table:table-cell>`;
    const cellsOf = (rows: OfficeContentNode[]) => rows.filter(r => r.type === 'row').map(r => (r.children ?? []).map(x => [x.text, (x.metadata as any)?.col, (x.metadata as any)?.colSpan ?? 1, (x.metadata as any)?.rowSpan ?? 1]));
    const merged = `<table:table-row>${c('A1', ' table:number-columns-spanned="2" table:number-rows-spanned="2"')}<table:covered-table-cell/>${c('C1')}</table:table-row><table:table-row><table:covered-table-cell table:number-columns-repeated="2"/>${c('C2')}</table:table-row>`;
    const ods = await OfficeParser.parseOffice(pkg('spreadsheet', `<table:table table:name="S">${merged}<table:table-row><table:table-cell office:value-type="string"><text:p>Line one</text:p><text:p>Line two</text:p></table:table-cell></table:table-row></table:table>`), { fileType: 'ods' } as any);
    assert.deepStrictEqual(cellsOf(ods.content[0].children ?? []), [[['A1', 0, 2, 2], ['C1', 2, 1, 1]], [['C2', 2, 1, 1]], [['Line one\nLine two', 0, 1, 1]]], 'ODS: merged cells keep their spans, and the cells after them their columns');
    assert.ok(((await ods.to('text')).value as string).includes('Line one\nLine two'), 'ODS: a cell\'s paragraphs are lines (they ran together)');
    const odt = await OfficeParser.parseOffice(pkg('text', `<table:table table:name="T">${merged}</table:table><text:p>First<text:line-break/>Second</text:p>`), { fileType: 'odt' } as any);
    assert.deepStrictEqual(cellsOf(odt.content[0].children ?? []), [[['A1', 0, 2, 2], ['C1', 2, 1, 1]], [['C2', 2, 1, 1]]], 'ODT: a covered cell holds its column');
    // A line break is a break node, which every writer shows as one.
    assert.deepStrictEqual(odt.content[1].children?.map(n => [n.type, n.text ?? (n.metadata as any)?.breakType]), [['text', 'First'], ['break', 'textWrapping'], ['text', 'Second']], 'ODT: text:line-break is a break node');
    assert.ok(((await odt.to('md')).value as string).includes('First  \nSecond'), 'ODT: a line break is a Markdown hard break');
    assert.ok(((await odt.to('html')).value as string).includes('First<br>'), 'ODT: a line break is an HTML <br>');
    assert.ok(((await odt.to('rtf')).value as string).includes('First}\\line'), 'ODT: a line break is an RTF \\line');

    // A chart object as LibreOffice writes it.
    const declarations = Array.from({ length: 30 }, (_, i) => `xmlns:ns${i}="urn:example:namespace:number:${i}"`).join(' ');
    const chartXml = `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${declarations} ${NS}><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:title><text:p>My Chart</text:p></chart:title><table:table table:name="local-table"><table:table-header-rows><table:table-row><table:table-cell><text:p/></table:table-cell>${c('Sales')}${c('Costs')}</table:table-row></table:table-header-rows><table:table-rows><table:table-row>${c('Q1')}<table:table-cell office:value-type="float" office:value="10"><text:p>10</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row></table:table-rows></table:table></chart:chart></office:chart></office:body></office:document-content>`;
    assert.ok(chartXml.indexOf('opendocument:xmlns:chart') > 500, 'ODF chart fixture: the chart namespace is past byte 500');
    const withChart = await OfficeParser.parseOffice(pkg('text', '<text:p><draw:frame draw:name="c"><draw:object xlink:href="./Object 1"/></draw:frame></text:p>', { 'Object 1/content.xml': chartXml }), { fileType: 'odt', extractAttachments: true } as any);
    const chart = withChart.attachments.find(a => a.chartData)?.chartData;
    assert.deepStrictEqual(chart && { title: chart.title, labels: chart.labels, series: chart.dataSets.map(d => [d.name, d.values]) }, { title: 'My Chart', labels: ['Q1'], series: [['Sales', ['10']], ['Costs', ['7']]] }, 'ODF: a LibreOffice chart has its title, labels and values');
}

/**
 * A note cited again (`text:note-ref`) is that note, read back as the same node, never its number as
 * text; the writer refers to a note only within the part it wrote it in.
 */
async function testOdtNoteReferences(): Promise<void> {
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
    const odtOf = (body: string) => Buffer.from(zipSync({ mimetype: strToU8('application/vnd.oasis.opendocument.text'), 'content.xml': strToU8(`<?xml version="1.0"?><office:document-content ${NS}><office:body><office:text>${body}</office:text></office:body></office:document-content>`) }));
    const note = (id: string, body: string) => `<text:note text:id="${id}" text:note-class="footnote"><text:note-citation>1</text:note-citation><text:note-body><text:p>${body}</text:p></text:note-body></text:note>`;
    const ref = (id: string, format = 'text') => `<text:note-ref text:note-class="footnote" text:reference-format="${format}" text:ref-name="${id}">1</text:note-ref>`;
    const runs = async (body: string, config: any = {}) => {
        const ast = await OfficeParser.parseOffice(odtOf(body), { fileType: 'odt', ...config } as any);
        return { ast, runs: (ast.content[0]?.children ?? []).map(n => [n.text, n.notes?.map(x => x.text)]) };
    };
    const shared = await runs(`<text:p>First${note('ftn1', 'Shared')} again${ref('ftn1')} end</text:p>`);
    assert.deepStrictEqual(shared.runs, [['First', ['Shared']], [' again', ['Shared']], [' end', undefined]], 'ODT: a note cited again is that note, not its number');
    const [first, second] = (shared.ast.content[0].children ?? []).filter(n => n.notes);
    assert.ok(first.notes![0] === second.notes![0], 'ODT: both citations share the one note node');
    assert.deepStrictEqual((await runs(`<text:p>Before${ref('ftn1')} then${note('ftn1', 'Later')}</text:p>`)).runs, [['Before', ['Later']], [' then', ['Later']]], 'ODT: a citation before its note is that note too');
    assert.deepStrictEqual((await runs(`<text:p>A${note('ftn1', 'Shared')} see${ref('ftn1', 'page')}</text:p>`)).runs, [['A', ['Shared']], [' see', undefined], ['1', undefined]], 'ODT: a reference to the page a note is on stays the page');
    assert.deepStrictEqual((await runs(`<text:p>A${note('ftn1', 'Shared')} again${ref('ftn1')}</text:p>`, { ignoreNotes: true })).runs, [['A', undefined], [' again', undefined]], 'ODT: with notes left out, a note cited again shows nothing');
    // Hidden text and a script show nothing, even in a paragraph with nothing else.
    const hidden = await OfficeParser.parseOffice(odtOf('<text:p><text:hidden-text text:condition="true" text:string-value="SECRET">SECRET</text:hidden-text></text:p><text:p><text:script>var SCRIPT = 1;</text:script></text:p><text:p>Shown</text:p>'), { fileType: 'odt' } as any);
    assert.deepStrictEqual(hidden.content.map(n => n.text).filter(Boolean), ['Shown'], 'ODT: hidden text and a script are not read as text');

    // The writer: a note cited from the header as well as the body is written in full in styles.xml too,
    // where a reference to its id in content.xml would reach nothing.
    const shared2 = { type: 'note', text: 'Shared note', metadata: { noteType: 'footnote', noteId: '1' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'Shared note' }] }] };
    const written = unzipSync((await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [
        { type: 'paragraph', children: [{ type: 'text', text: 'First', notes: [shared2] }, { type: 'text', text: ' again', notes: [shared2] }] },
    ], auxiliary: { headers: [{ type: 'paragraph', children: [{ type: 'text', text: 'Head', notes: [shared2] }] }] } } as any, 'odt')).value as Uint8Array);
    const styles = strFromU8(written['styles.xml']), contentXml = strFromU8(written['content.xml']);
    const ids = (xml: string, tag: string) => [...xml.matchAll(new RegExp(`<text:${tag} [^>]*text:(?:id|ref-name)="([^"]+)"`, 'g'))].map(m => m[1]);
    assert.deepStrictEqual([ids(contentXml, 'note'), ids(contentXml, 'note-ref'), ids(styles, 'note'), ids(styles, 'note-ref')], [['ftn1'], ['ftn1'], ['ftn1_2'], []], 'ODT: a note is referred to only within the part it is written in');
}

/**
 * Inline runs among blocks (an ODS cell's runs, a note's, a comment's, an admonition's) are one
 * paragraph in ODT, not a paragraph each.
 */
async function testOdtInlineRuns(): Promise<void> {
    const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
    const ods = Buffer.from(zipSync({
        mimetype: strToU8('application/vnd.oasis.opendocument.spreadsheet'),
        'content.xml': strToU8(`<?xml version="1.0"?><office:document-content ${NS}><office:automatic-styles><style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p>Total: <text:span text:style-name="B">5</text:span> units</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>`),
    }));
    const sheet = await OfficeParser.parseOffice(ods, { fileType: 'ods' } as any);
    const cellXml = strFromU8(unzipSync((await sheet.to('odt')).value as Uint8Array)['content.xml']).match(/<table:table-cell[^>]*>(.*?)<\/table:table-cell>/)?.[1] ?? '';
    assert.strictEqual((cellXml.match(/<text:p[ >]/g) ?? []).length, 1, `ODT: an ODS cell's runs are one paragraph (${cellXml})`);
    const T = (text: string, extra: any = {}) => ({ type: 'text', text, ...extra });
    const runs = () => [T('Alpha '), T('bold', { formatting: { bold: true } }), T(' omega')];
    const odt = strFromU8(unzipSync((await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [
        { type: 'paragraph', children: [T('Noted', { notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: runs() }], comments: [{ type: 'comment', metadata: { author: 'A' }, children: runs() }] })] },
        { type: 'admonition', metadata: { admonitionType: 'note' }, children: runs() },
    ] } as any, 'odt')).value as Uint8Array)['content.xml']);
    const back = await OfficeParser.parseOffice(Buffer.from(zipSync({ mimetype: strToU8('application/vnd.oasis.opendocument.text'), 'content.xml': strToU8(odt) })), { fileType: 'odt' } as any);
    const notes = collectAllNodes(back).filter(n => n.type === 'note' || n.type === 'comment').map(n => (n.children ?? []).map(c => c.text));
    assert.deepStrictEqual(notes, [['Alpha bold omega'], ['Alpha bold omega']], 'ODT: a note\'s and a comment\'s runs are one paragraph each');
    assert.deepStrictEqual(back.content.map(n => n.text).slice(1), ['Note', 'Alpha bold omega'], 'ODT: an admonition\'s runs are one paragraph');
}

/**
 * The LaTeX that TeX interprets as code: verbatim/lstlisting bodies (literal) and `%` comments
 * (an unescaped `%` to the end of the line) removed.
 */
function liveLatex(tex: string): string {
    return tex
        .replace(/\\begin\{(verbatim|lstlisting)\}[\s\S]*?\\end\{\1\}/g, '')
        .replace(/(^|[^\\])((?:\\\\)*)%.*$/gm, '$1$2');
}

/** A carried image's PDF: printable ASCII lines, none ending in a space, whose cross-reference offsets land on their objects. */
function assertTextPdf(pdf: string, label: string): void {
    assert.ok(!/[^\n\x20-\x7e]/.test(pdf) && !/ \n| $/.test(pdf), `${label}: the carried PDF is printable ASCII with no trailing spaces`);
    const xref = /(\d+) 0 obj\n<< \/Type \/XRef [^]*?stream\n([^]*?)>\nendstream/.exec(pdf);
    const rows = xref ? xref[2].split('\n').slice(1) : [];
    assert.ok(rows.length > 0 && rows.every((row, i) => pdf.startsWith(`${i + 1} 0 obj`, parseInt(row.slice(2, 10), 16))), `${label}: every cross-reference offset lands on its object`);
    const start = /startxref\n(\d+)\n%%EOF$/.exec(pdf);
    assert.ok(start && pdf.startsWith(`${xref![1]} 0 obj`, Number(start[1])), `${label}: startxref points at the cross-reference stream`);
}

async function testLatexGeneration(): Promise<void> {
    // ── Tier 1: exhaustive Markdown fixture ───────────────────────────────────
    const src = await OfficeParser.parseOffice(path.join(__dirname, 'files/exhaustive/markdown.md'));
    const warnings: any[] = [];
    const tex = (await src.to('tex', { onWarning: (w: any) => warnings.push(w) })).value as string;
    assert.strictEqual(typeof tex, 'string', 'TEX: default output is a string');
    assert.ok(tex.startsWith('% Generated by officeParser.\n\\RequirePackage{iftex}\n'), 'TEX: article preamble');
    assert.ok(tex.includes('\\def\\officeparserdriver{}\\ifpdf\\else\\ifXeTeX\\else\\def\\officeparserdriver{dvipdfmx,}\\fi\\fi\n\\documentclass[\\officeparserdriver 11pt]{article}\n'), 'TEX: DVI output gets the dvipdfmx driver as a global class option');
    assert.ok(/\\iftutex\n {2}\\usepackage\{fontspec\}\n\\else\n {2}\\usepackage\[T1\]\{fontenc\}\n {2}\\usepackage\[utf8\]\{inputenc\}\n {2}\\usepackage\{lmodern\}\n\\fi/.test(tex), 'TEX: per-engine font setup (XeTeX/LuaTeX vs pdfTeX and the pTeX family)');
    assert.ok(!tex.includes('xeCJK') && !tex.includes('cmunrm'), 'TEX: no script fonts for a Latin document');
    assert.ok(tex.includes('\\title{Exhaustive Markdown Test}') && tex.includes('\\author{Test Author}'), 'TEX: \\title/\\author from metadata');
    assert.ok(tex.includes('pdftitle={Exhaustive Markdown Test}') && tex.includes('Description={Tests every markdown feature}'), 'TEX: hyperref PDF metadata');
    assert.ok(tex.includes('\\setcounter{secnumdepth}{-\\maxdimen}'), 'TEX: sections unnumbered by default');
    assert.ok(!tex.includes('\\maketitle'), 'TEX: no visible title block without renderMetadata');
    assert.strictEqual((tex.match(/\\begin\{document\}/g) || []).length, 1, 'TEX: one document environment');
    // Headings H1-H6, with the \paragraph/\subparagraph display fix and labels.
    assert.ok(tex.includes('\\section{Heading Level 1}\\label{h1-anchor}\\label{heading-level-1}'), 'TEX: H1 -> \\section with source and generated labels');
    for (const [cmd, n] of [['subsection', 2], ['subsubsection', 3], ['paragraph', 4], ['subparagraph', 5], ['subparagraph', 6]] as const) {
        assert.ok(tex.includes(`\\${cmd}{Heading Level ${n}}`), `TEX: H${n} -> \\${cmd}`);
    }
    assert.ok(tex.includes('\\renewcommand\\paragraph{\\@startsection{paragraph}'), 'TEX: \\paragraph made a display heading');
    // Inline formatting.
    for (const frag of ['\\textbf{bold text}', '\\textit{italic text}', '\\sout{strikethrough text}', '\\uline{underlined text}', '\\textsubscript{subscript}', '\\textsuperscript{superscript}', '\\texttt{monospace code}']) {
        assert.ok(tex.includes(frag), `TEX: inline ${frag}`);
    }
    // Links, citations, footnotes.
    assert.ok(tex.includes('\\href{https://example.com}{Visit Example}'), 'TEX: external link -> \\href');
    assert.ok(tex.includes('\\hyperref[h1-anchor]{Go to H1}'), 'TEX: internal link -> \\hyperref to a defined label');
    assert.ok(tex.includes('\\cite{smith2023}'), 'TEX: citation -> \\cite');
    assert.ok(tex.includes('\\footnote{This is the footnote definition text.}'), 'TEX: footnote at its reference point');
    assert.ok(!/WikiPage\}|\\hyperref\[WikiPage\]/.test(tex) && tex.includes('Wikilink: WikiPage'), 'TEX: unresolved wikilink stays plain text');
    // Lists, task items, description lists.
    assert.ok(/\\item Unordered item B\n\\begin\{itemize\}\n\\item Nested unordered item\n\\end\{itemize\}/.test(tex), 'TEX: nested itemize');
    assert.ok(/\\item Ordered item 2\n\\begin\{enumerate\}\n\\item Nested ordered item/.test(tex), 'TEX: nested enumerate');
    assert.ok(tex.includes('\\item[$\\boxtimes$] Completed task item') && tex.includes('\\item[$\\square$] Incomplete task item'), 'TEX: task items');
    assert.ok(tex.includes('\\item[{Term Alpha}] Description for term alpha') && tex.includes('\\item[] Description two for beta'), 'TEX: description list');
    // Code and math.
    assert.ok(/\\begin\{lstlisting\}\[language=\{Python\}\]\ndef hello\(\):/.test(tex), 'TEX: known language -> lstlisting');
    assert.ok(/\\begin\{verbatim\}\nconst x: number = 42;/.test(tex), 'TEX: other code -> verbatim');
    assert.ok(tex.includes('$E=mc^2$') && tex.includes('\\[\na^2 + b^2 = c^2\n\\]'), 'TEX: inline and display math');
    // Tables: column alignment, merged cells, repeated header.
    assert.ok(/\{\|>\{\\raggedright\\arraybackslash\}p\{[^}]+\}\|>\{\\centering\\arraybackslash\}p\{[^}]+\}\|>\{\\raggedleft\\arraybackslash\}p/.test(tex), 'TEX: column alignment from the separator row');
    assert.ok(/\\multicolumn\{2\}\{\|[^}]*\}p\{[^}]+\}\|\}\{Merged Header\}/.test(tex), 'TEX: colSpan -> \\multicolumn');
    assert.ok(tex.includes('\\multirow{2}{=}{Rowspan Cell}') && tex.includes('\\cline{2-3}'), 'TEX: rowSpan -> \\multirow with \\cline past it');
    assert.ok(/\\hline\n\\endhead\n/.test(tex), 'TEX: bold header row repeats via \\endhead');
    // Admonitions, quotes, alignment, rule, non-ASCII fallback.
    assert.ok(tex.includes('\\definecolor{hex0969DA}{HTML}{0969DA}') && tex.includes('\\textbf{\\textcolor{hex0969DA}{Note}}\\par'), 'TEX: admonition -> labelled quote');
    assert.ok(tex.includes('{\\raggedleft Right-aligned paragraph content.\\par}'), 'TEX: right-aligned paragraph');
    assert.ok(tex.includes('\\rule{0.5\\linewidth}{0.5pt}'), 'TEX: thematic break -> rule');
    assert.ok(tex.includes('\\newunicodechar{\u2764}{\\ding{170}}') && tex.includes('\\usepackage{pifont}'), 'TEX: symbol fallback for a character the fonts lack');
    assert.ok(tex.includes('-{}-{}-{}-{}- not actually'), 'TEX: hyphen runs keep every hyphen');
    // Only the remote image cannot be represented, and the citation has no bibliography to resolve it.
    assert.deepStrictEqual([...new Set(warnings.map(w => w.code))], ['CONTENT_NOT_REPRESENTABLE', 'CITATIONS_NOT_RESOLVED'], 'TEX: only warnings are the remote image and the unresolved citation');
    assert.strictEqual(warnings.find(w => w.code === 'CITATIONS_NOT_RESOLVED')?.message, `The LaTeX output cites a key ('smith2023') that has no entry in a bibliography it holds, so \\cite prints [?] for it until one is added: a \\bibliography{file} with a .bib file that has it, or a thebibliography list.`, 'TEX: CITATIONS_NOT_RESOLVED exact message');
    const live = liveLatex(tex).replace(/\\[{}]/g, '');
    assert.strictEqual((live.match(/\{/g) || []).length, (live.match(/\}/g) || []).length, 'TEX: braces balance outside verbatim/comments');
    assert.strictEqual((live.match(/\\begin\{/g) || []).length, (live.match(/\\end\{/g) || []).length, 'TEX: every \\begin has an \\end');

    // ── Tier 1b: DOCX fixture (endnotes, bundle, images) ──────────────────────
    const docx = await OfficeParser.parseOffice(path.join(__dirname, 'files/test.docx'), { extractAttachments: true });
    const docxWarnings: any[] = [];
    const docxTex = (await docx.to('tex', { onWarning: (w: any) => docxWarnings.push(w) })).value as string;
    assert.ok(docxTex.includes('\\usepackage{endnotes}') && /\\endnote\{Endnotes are typically/.test(docxTex) && docxTex.includes('\\theendnotes\n\n\\end{document}'), 'TEX docx: endnotes via the endnotes package');
    assert.ok(/\\footnote\{In paged media/.test(docxTex), 'TEX docx: footnote body without the note run size');
    // The image travels inside the .tex: a filecontents* block writes it as a PDF, which \includegraphics reads.
    const carriedImage = /\\includegraphics\[bb=0 0 810 810,width=\{\\ifdim [\d.]+pt>\\linewidth\\linewidth\\else [\d.]+pt\\fi\},height=\{[^}]+\},keepaspectratio,alt=\{[^}]*\}\]\{(image-[0-9a-f]{8}\.pdf)\}/.exec(docxTex);
    assert.ok(carriedImage, 'TEX docx: image bounded to the line and page, carried as a PDF of its stated size');
    const carriedBlock = new RegExp(`\\\\begin\\{filecontents\\*\\}\\{${carriedImage![1].replace('.', '\\.')}\\}\\n([^]*?)\\n\\\\end\\{filecontents\\*\\}`).exec(docxTex);
    assert.ok(carriedBlock && docxTex.indexOf(carriedBlock[0]) < docxTex.indexOf('\\begin{document}'), 'TEX docx: the image is carried in a filecontents* block in the preamble');
    assertTextPdf(carriedBlock![1], 'TEX docx');
    assert.ok(!docxWarnings.some(w => w.code === 'IMAGES_NOT_BUNDLED'), 'TEX docx: a carried image needs no file placed');
    const docxBack = await OfficeParser.parseOffice(Buffer.from(docxTex), { fileType: 'tex', extractAttachments: true } as any);
    const docxJpeg = docx.attachments.find(a => a.mimeType === 'image/jpeg')!.data;
    assert.deepStrictEqual(docxBack.attachments.map(a => [a.name, a.mimeType, a.data === docxJpeg]), [[carriedImage![1].replace('.pdf', '.jpg'), 'image/jpeg', true]], 'TEX docx: parsing the .tex back gives the very JPEG it carries');
    assert.ok(((await docxBack.to('tex')).value as string).includes(carriedBlock![0]), 'TEX docx: regenerating the parsed .tex carries the same image block, under the same name');
    // Without embedImages the .tex refers to the file, and the warning says where it goes.
    const refDocxWarnings: any[] = [];
    const referencedTex = (await docx.to('tex', { texConfig: { embedImages: false }, onWarning: (w: any) => refDocxWarnings.push(w) } as any)).value as string;
    assert.ok(/\]\{images\/image\.jpg\}/.test(referencedTex) && !referencedTex.includes('filecontents'), 'TEX docx: embedImages false refers to images/');
    assert.strictEqual(refDocxWarnings.find(w => w.code === 'IMAGES_NOT_BUNDLED')?.message, `The LaTeX output references 1 image file ('images/image.jpg') that is not part of the .tex source (texConfig.embedImages is off). Place it at that path, relative to the .tex (the bytes are in ast.attachments), or set texConfig.bundle: true to get a zip containing the .tex and its images.`, 'TEX docx: IMAGES_NOT_BUNDLED names the file (exact message)');
    const bundle = (await docx.to('tex', { texConfig: { bundle: true }, metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } } as any)).value as Uint8Array;
    assert.ok(bundle instanceof Uint8Array && bundle[0] === 0x50 && bundle[1] === 0x4B, 'TEX bundle: a zip');
    const bundleFiles = unzipSync(bundle);
    assert.deepStrictEqual(Object.keys(bundleFiles), ['main.tex', 'images/image.jpg'], 'TEX bundle: main.tex first, then the image');
    const pinnedTex = (await docx.to('tex', { texConfig: { embedImages: false }, metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') }, onWarning: () => {} } as any)).value as string;
    assert.strictEqual(strFromU8(bundleFiles['main.tex']), pinnedTex, 'TEX bundle: main.tex is the source the string mode returns without embedImages');
    const again = (await docx.to('tex', { texConfig: { bundle: true }, metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } } as any)).value as Uint8Array;
    assert.ok(Buffer.from(bundle).equals(Buffer.from(again)), 'TEX bundle: byte-identical across runs');

    // ── Tier 2: synthetic AST covering every construct ────────────────────────
    const TINY_GIF_B64 = 'R0lGODlhAQABAAAAACw=';
    const note = (text: string, noteType = 'footnote') => ({ type: 'note', metadata: { noteType }, children: [{ type: 'paragraph', children: [{ type: 'text', text }] }] });
    const cell = (text: string, metadata: any = {}, extra: any = {}) => ({ type: 'cell', metadata, children: [{ type: 'paragraph', children: [{ type: 'text', text, ...extra }] }] });
    const synthetic: any = {
        type: 'docx',
        metadata: { formatting: { size: '11pt' }, title: 'Synthetic', author: 'A. Author', subject: 'Subj', keywords: 'k1, k2', language: 'en-GB', created: new Date('2023-05-06T07:08:09Z'), modified: new Date('2024-01-01T00:00:00Z'), customProperties: { 'Dept Name!': 'Finance', Count: 3 } },
        attachments: [
            { name: '../../etc/passwd.png', mimeType: 'image/png', data: TINY_PNG_B64 },
            { name: 'anim.gif', mimeType: 'image/gif', data: TINY_GIF_B64 },
        ],
        auxiliary: {
            headers: [{ type: 'header', metadata: { type: 'default' }, children: [{ type: 'paragraph', metadata: { alignment: 'right' }, children: [{ type: 'text', text: 'Running head', notes: [note('dropped in header')] }] }] }],
            footers: [{ type: 'footer', metadata: { type: 'default' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'Page footer' }] }] }],
        },
        content: [
            { type: 'heading', metadata: { level: 1 }, children: [{ type: 'text', text: 'Intro', formatting: { bold: true, size: '20pt' } }] },
            { type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: 'Intro' }] },
            { type: 'paragraph', metadata: { anchorIds: ['target', 'unlinked'] }, comments: [{ type: 'comment', metadata: { author: 'Rev', date: '2024-02-02' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'line one\nline two' }] }] }], children: [
                { type: 'text', text: 'Big ', formatting: { size: '16pt' } },
                { type: 'text', text: 'high', formatting: { backgroundColor: '#ffff00' } },
                { type: 'text', text: 'lighted words', formatting: { backgroundColor: '#ffff00' } },
                { type: 'text', text: ' and ', notes: [note('a footnote'), note('an endnote', 'endnote')] },
                { type: 'text', text: 'self', metadata: { link: '#target', linkType: 'internal' } },
                { type: 'text', text: ' and ', },
                { type: 'text', text: 'nowhere', metadata: { link: '#missing', linkType: 'internal' } },
                { type: 'text', text: ' and ' },
                { type: 'text', text: 'bad', metadata: { link: 'javascript:alert(1)', linkType: 'external' } },
                { type: 'code', metadata: { math: 'inline' }, text: 'x^2' },
                { type: 'code', metadata: { math: 'inline' }, text: '\\input{/etc/passwd}' },
            ] },
            { type: 'list', metadata: { listId: 'L', listType: 'ordered', indentation: 0, itemIndex: 4 }, children: [{ type: 'text', text: 'five' }] },
            { type: 'list', metadata: { listId: 'L', listType: 'unordered', indentation: 2, itemIndex: 0 }, children: [{ type: 'text', text: 'deep' }] },
            { type: 'list', metadata: { listId: 'L', listType: 'ordered', indentation: 9, itemIndex: 0 }, children: [{ type: 'text', text: 'too deep' }] },
            { type: 'list', metadata: { listId: 'L', listType: 'unordered', indentation: 0, itemIndex: 1 }, children: [{ type: 'text', text: 'switched type' }] },
            { type: 'table', children: [
                { type: 'row', children: [cell('H1', { col: 0, style: 'header', backgroundColor: '#DDEEFF' }), cell('H2', { col: 1, style: 'header' })] },
                { type: 'row', children: [
                    { type: 'cell', metadata: { col: 0 }, children: [
                        { type: 'heading', metadata: { level: 2 }, children: [{ type: 'text', text: 'Cell heading' }] },
                        { type: 'table', children: [{ type: 'row', children: [cell('inner', {}, { notes: [note('inner note')] })] }] },
                    ] },
                    { type: 'cell', metadata: { col: 1 }, children: [{ type: 'code', text: 'a\\end{verbatim}\\input{x}', metadata: { language: 'python' } }] },
                ] },
            ] },
            { type: 'code', text: 'print(1)\n\tindented', metadata: { language: 'PYTHON' } },
            { type: 'code', text: 'x = 1', metadata: { language: 'brainfuck' } },
            { type: 'code', text: 'safe\n\\end{verbatim}\\input{/etc/passwd}', metadata: {} },
            { type: 'code', metadata: { math: 'block' }, text: 'a &= b \\\\ c &= d' },
            { type: 'code', metadata: { math: 'block' }, text: '\\begin{align}x&=1\\end{align}' },
            { type: 'image', metadata: { attachmentName: '../../etc/passwd.png', altText: 'logo', align: 'center' } },
            { type: 'image', metadata: { attachmentName: 'anim.gif', altText: 'animation' } },
            { type: 'image', metadata: { url: 'https://example.com/a.png', altText: 'remote' } },
            { type: 'admonition', metadata: { admonitionType: 'warning', title: 'Careful' }, children: [{ type: 'paragraph', children: [{ type: 'text', text: 'body' }] }] },
            { type: 'embed', metadata: { embedType: 'youtube', videoId: 'dQw4w9WgXcQ', label: 'Video' } },
            { type: 'break', metadata: { breakType: 'page' } },
            { type: 'paragraph', children: [{ type: 'text', text: 'check \u2713 and \u4E2D' }] },
            { type: 'sheet', metadata: { sheetName: 'Data_1' }, children: [
                { type: 'row', children: [cell('r0', { row: 0, col: 0 })] },
                { type: 'row', children: [cell('r2', { row: 2, col: 1 })] },
            ] },
        ],
        getImages: () => [],
    };
    const sWarn: any[] = [];
    const stex = (await OfficeGenerator.generate(synthetic, 'tex', { renderMetadata: true, onWarning: (w: any) => sWarn.push(w) } as any)).value as string;
    const sLive = liveLatex(stex);
    // Metadata and title block.
    assert.ok(stex.includes('\\maketitle') && stex.includes('\\date{2024-01-01}'), 'TEX synthetic: renderMetadata -> \\maketitle with the modified date');
    for (const key of ['pdfsubject={Subj}', 'pdfkeywords={k1, k2}', 'pdflang={en-GB}', 'pdfcreationdate={D:20230506070809Z}', 'pdfmoddate={D:20240101000000Z}', 'DeptName={Finance}', 'Count={3}']) {
        assert.ok(stex.includes(key), `TEX synthetic: hyperref metadata ${key}`);
    }
    // Headings: uniform bold and sizes dropped inside the heading; duplicate slug labelled once.
    assert.ok(stex.includes('\\section{Intro}\\label{intro}') && /\\subsection\{Intro\}\n/.test(stex), 'TEX synthetic: heading formatting implied, duplicate slug not relabelled');
    // Comments, anchors, sizes, highlight, notes.
    assert.ok(stex.includes('% Comment (Rev, 2024-02-02): line one\n% line two\n'), 'TEX synthetic: comment as % lines, one per line');
    assert.ok(stex.includes('\\phantomsection\\label{target}') && !stex.includes('\\label{unlinked}'), 'TEX synthetic: only linked anchors are labelled');
    assert.ok(stex.includes('{\\fontsize{16pt}{19.2pt}\\selectfont Big }'), 'TEX synthetic: a non-body size is written explicitly');
    assert.ok(stex.includes('\\colorbox{hexFFFF00}{\\strut highlighted} \\colorbox{hexFFFF00}{\\strut words}'), 'TEX synthetic: highlight merged across runs and boxed per word');
    assert.ok(stex.includes('\\footnote{a footnote}\\endnote{an endnote}'), 'TEX synthetic: footnote and endnote marks at the reference point');
    assert.ok(stex.includes('\\hyperref[target]{self}') && stex.includes(' and nowhere and bad'), 'TEX synthetic: resolvable internal link linked, unresolvable and unsafe links left as text');
    assert.ok(stex.includes('$x^2$'), 'TEX synthetic: safe inline math typeset');
    assert.ok(stex.includes('\\texttt{\\textbackslash{}input\\{/etc/passwd\\}}'), 'TEX synthetic: unsafe math written as literal text');
    const mathWarn = sWarn.find(w => w.code === 'MATH_WRITTEN_AS_TEXT');
    assert.strictEqual(mathWarn?.message, `A math expression used '\\input', which can read or write files, run programs, or redefine commands, so it was written to the LaTeX output as literal text instead of typeset math.`, 'TEX synthetic: MATH_WRITTEN_AS_TEXT exact message');
    // Lists: start counter, level jump, depth clamp, type switch at a level.
    assert.ok(/\\begin\{enumerate\}\n\\setcounter\{enumi\}\{4\}\n\\item five\n\\begin\{itemize\}\n\\item\[\]\n\\begin\{itemize\}\n\\item deep\n/.test(stex), 'TEX synthetic: enumerate starts at 5; a two-level jump hangs on an empty \\item[]');
    assert.ok(/\\item deep\n\\begin\{enumerate\}\n\\item too deep\n\\end\{enumerate\}/.test(stex), 'TEX synthetic: nesting past LaTeX\'s limit is clamped');
    assert.ok(/\\end\{enumerate\}\n\\begin\{itemize\}\n\\item switched type\n\\end\{itemize\}/.test(stex), 'TEX synthetic: a type change at a level closes and reopens the list');
    // Tables: header background, heading in a cell, nested tabular with deferred footnote, code in a cell.
    assert.ok(stex.includes('\\usepackage[table]{xcolor}') && stex.includes('\\cellcolor{hexDDEEFF}H1'), 'TEX synthetic: cell background via colortbl');
    assert.ok(stex.includes('\\phantomsection\\label{cell-heading}\\textbf{Cell heading}') || stex.includes('\\textbf{Cell heading}'), 'TEX synthetic: a heading in a cell becomes a bold line');
    assert.ok(/\\begin\{tabular\}[\s\S]*inner\\footnotemark\{\}[\s\S]*\\end\{tabular\}/.test(stex), 'TEX synthetic: nested table is a tabular with a footnote mark');
    assert.ok(stex.includes('\\end{tabular}\n\\addtocounter{footnote}{-1}\n\\stepcounter{footnote}\\footnotetext{inner note}'), 'TEX synthetic: the tabular flushes its deferred footnote text right after itself');
    assert.ok(/\{\\ttfamily a\\textbackslash\{\}end\\\{verbatim\\\}\\textbackslash\{\}input\\\{x\\\}\}/.test(stex), 'TEX synthetic: code in a cell is escaped typewriter text');
    // Code blocks.
    assert.ok(/\\begin\{lstlisting\}\[language=\{Python\}\]\nprint\(1\)\n {4}indented\n\\end\{lstlisting\}/.test(stex), 'TEX synthetic: language matched case-insensitively, tabs expanded');
    assert.ok(/\\begin\{verbatim\}\nx = 1\n\\end\{verbatim\}/.test(stex), 'TEX synthetic: unknown language -> verbatim');
    assert.ok(!/\\begin\{verbatim\}\nsafe/.test(stex) && stex.includes('{\\ttfamily safe\\hfil\\break{}\\textbackslash{}end\\{verbatim\\}\\textbackslash{}input\\{/etc/passwd\\}}'), 'TEX synthetic: code containing \\end{verbatim} is not put in verbatim');
    assert.ok(stex.includes('\\[\n\\begin{aligned}\na &= b \\\\ c &= d\n\\end{aligned}\n\\]'), 'TEX synthetic: top-level & and \\\\ wrapped in aligned');
    assert.ok(stex.includes('\\begin{align}x&=1\\end{align}') && !stex.includes('\\[\n\\begin{align}'), 'TEX synthetic: a whole display environment is written bare');
    // Images: sanitized bundle path, placeholder for an unincludable type, remote link. The tiny PNG is
    // inconsistent (its header says gray+alpha, its data is RGBA), so it cannot be carried and stays a file.
    assert.ok(stex.includes('{\\centering \\includegraphics[') && stex.includes(']{images/passwd.png}\\par}') && !stex.includes('filecontents'), 'TEX synthetic: attachment name reduced to a safe base file name; a PNG it cannot read is not carried');
    assert.ok(stex.includes('\\fbox{Image: animation}'), 'TEX synthetic: GIF drawn as a placeholder');
    assert.ok(stex.includes('\\href{https://example.com/a.png}{remote}'), 'TEX synthetic: remote image becomes a link');
    assert.ok(sWarn.some(w => w.code === 'CONTENT_NOT_REPRESENTABLE' && w.details?.feature === 'image/gif image'), 'TEX synthetic: GIF warns');
    const sNotBundled = sWarn.find(w => w.code === 'IMAGES_NOT_BUNDLED');
    assert.strictEqual(sNotBundled?.message.split(' Place')[0], `The LaTeX output references 1 image file ('images/passwd.png') that is not part of the .tex source (a .tex carries only PNG and JPEG images it can read, up to a limit per document).`, 'TEX synthetic: IMAGES_NOT_BUNDLED names only the file the .tex reads (the GIF is a labelled box)');
    // Images referenced by path, with no image data: a plain relative path stays an \includegraphics of
    // that path (the caller supplies the file), a web image is a link, and any other path is refused.
    const refAst = await OfficeParser.parseOffice(Buffer.from('<p><img src="pics/a.png"> <img src="https://example.com/b.png"> <img src="../up/c.png"> <img src="/etc/d.png"></p>'), { fileType: 'html' });
    const refWarn: any[] = [];
    const refTex = (await refAst.to('tex', { onWarning: (w: any) => refWarn.push(w) } as any)).value as string;
    assert.ok(refTex.includes('\\includegraphics{pics/a.png}'), 'TEX refs: a relative image path stays an image at that path, at its natural size');
    const spaced = (await (await OfficeParser.parseOffice(Buffer.from('<p><img src="my%20pics/fig%20one.png"></p>'), { fileType: 'html' })).to('tex')).value as string;
    assert.ok(spaced.includes('\\includegraphics{my pics/fig one.png}'), 'TEX refs: a percent-encoded path with spaces is decoded to the file name');
    assert.ok(refTex.includes('\\href{https://example.com/b.png}') && refTex.includes('\\href{../up/c.png}') && refTex.includes('\\href{/etc/d.png}')
        && !/\\includegraphics(\[[^\]]*\])?\{[^}]*(\.\.\/|\/etc)/.test(refTex), 'TEX refs: web images and paths leaving the folder are links, never read by TeX');
    assert.deepStrictEqual(refWarn.filter(w => w.code === 'CONTENT_NOT_REPRESENTABLE').map(w => w.details?.feature),
        ['remote image', 'image path that is absolute, leaves its folder or uses characters other than letters, digits, spaces and . _ - /'], 'TEX refs: "remote" only for a web image');
    assert.strictEqual(refWarn.find(w => w.code === 'IMAGES_NOT_BUNDLED')?.message,
        `The LaTeX output references 1 image file ('pics/a.png') by the path the source document gave, without the image data, which the source did not contain. Place it at that path, relative to the .tex, before compiling.`,
        'TEX refs: IMAGES_NOT_BUNDLED names the path-referenced image (exact message)');
    const refZipWarn: any[] = [];
    const refZip = unzipSync((await refAst.to('tex', { texConfig: { bundle: true }, onWarning: (w: any) => refZipWarn.push(w) } as any)).value as Uint8Array);
    // A bundle must compile as is, but cannot hold a file the source only named: the image is drawn when
    // the file is added, and a labelled box stands in otherwise.
    assert.ok(strFromU8(refZip['main.tex']).includes('\\IfFileExists{pics/a.png}{\\includegraphics{pics/a.png}}{\\fbox{Image: pics/a.png}}'), 'TEX refs: a bundle guards a path-referenced image with \\IfFileExists');
    const noExt = unzipSync((await (await OfficeParser.parseOffice(Buffer.from('![d](figures/diagram)'), { fileType: 'md' } as any)).to('tex', { texConfig: { bundle: true } } as any)).value as Uint8Array);
    assert.ok(strFromU8(noExt['main.tex']).includes('\\IfFileExists{figures/diagram}{\\includegraphics[alt={d}]{figures/diagram}}{\\IfFileExists{figures/diagram.pdf}{'), 'TEX refs: a path without an extension probes the graphics extensions');
    assert.ok(refZipWarn.some(w => w.code === 'IMAGES_NOT_BUNDLED' && w.message.includes(`('pics/a.png')`)), 'TEX refs: a bundle still names the image it has no data to package');
    // Admonition, embed, page break, sheet, header/footer, Unicode.
    assert.ok(stex.includes('\\textbf{\\textcolor{hex9A6700}{Careful}}\\par\nbody'), 'TEX synthetic: admonition title and colour');
    assert.ok(stex.includes('\\href{https://www.youtube.com/watch?v=dQw4w9WgXcQ}{Video}'), 'TEX synthetic: YouTube embed becomes a link');
    assert.ok(/\n\\clearpage\n/.test(stex), 'TEX synthetic: page break');
    assert.ok(stex.includes('\\section{Data\\_1}') && /r0 &  \\\\\n\\hline\n &  \\\\\n\\hline\n & r2 \\\\/.test(stex), 'TEX synthetic: sheet title and sparse rows/columns kept in place');
    assert.ok(stex.includes('\\fancyhead[R]{Running head}') && stex.includes('\\fancyfoot[L]{Page footer}') && !stex.includes('dropped in header'), 'TEX synthetic: running header/footer via fancyhdr, notes dropped there');
    assert.ok(stex.includes('\\newunicodechar{\u2713}{\\ensuremath{\\checkmark}}') && stex.includes('\\DeclareUnicodeCharacter{4E2D}{{\\ttfamily[U+4E2D]}}'), 'TEX synthetic: symbol fallback everywhere, pdfLaTeX-only marker for a script it lacks');
    assert.ok(/\\iftutex\n( {2}\\newunicodechar[^\n]*\n)+\\else\n( {2}%[^\n]*\n)*( {2}\\DeclareUnicodeCharacter[^\n]*\n)+\\fi/.test(stex) && stex.includes('\\DeclareUnicodeCharacter{2713}{\\ensuremath{\\checkmark}}'),
        'TEX synthetic: \\newunicodechar under XeTeX/LuaTeX, \\DeclareUnicodeCharacter (fallbacks and markers) under pdfTeX and the pTeX family');
    // XeLaTeX and LuaLaTeX get fonts for the scripts Latin Modern lacks, the main CJK font following the text's language.
    const scriptTex = async (md: string) => (await (await OfficeParser.parseOffice(Buffer.from(md), { fileType: 'md' } as any)).to('tex')).value as string;
    const [jaTex, koTex, zhTex, mixedTex, greekTex] = await Promise.all([
        scriptTex('日本語のテキストです。カタカナも。'), scriptTex('한국어 텍스트입니다 漢字'), scriptTex('中文文本，简体。'),
        scriptTex('An English paragraph that mentions 中文 once, with "quotes" -- and dashes.'), scriptTex('Ωμέγα and Привет'),
    ]);
    assert.ok(jaTex.includes('\\setCJKmainfont{HaranoAjiMincho-Regular.otf}') && jaTex.includes('\\setmainjfont{HaranoAjiMincho-Regular.otf}'), 'TEX scripts: Japanese text gets Harano Aji');
    assert.ok(koTex.includes('\\setCJKmainfont{UnBatang.ttf}') && koTex.includes('\\setmainjfont{UnBatang.ttf}') && !koTex.includes('AltFont'), 'TEX scripts: Korean text gets UnBatang, with no Hangul AltFont needed');
    assert.ok(zhTex.includes('\\setCJKmainfont{FandolSong-Regular.otf}') && zhTex.includes('{{HaranoAjiMincho-Regular.otf}}') && !zhTex.includes('UnBatang'), 'TEX scripts: Chinese text gets Fandol, with Harano Aji as the fallback');
    assert.ok(zhTex.includes('jacharrange={-2}') && !zhTex.includes('xeCJKDeclareCharClass'), 'TEX scripts: CJK-dominant text keeps CJK punctuation');
    assert.ok(mixedTex.includes('jacharrange={-2, -3, -9}') && mixedTex.includes('\\xeCJKDeclareCharClass{Default}'), 'TEX scripts: mostly Latin text keeps quotes and dashes in the Latin font');
    assert.ok(/\\IfFontExistsTF\{FandolSong-Regular\.otf\}\{\\IfFontExistsTF\{HaranoAjiMincho-Regular\.otf\}\{%\n {2}\\ifXeTeX\n {4}\\IfFileExists\{xeCJK\.sty\}/.test(zhTex) && zhTex.includes('\\IfFileExists{luatexja-fontspec.sty}'), 'TEX scripts: the CJK setup is guarded by its fonts and packages');
    assert.ok(greekTex.includes('\\IfFontExistsTF{cmunrm.otf}{%') && !greekTex.includes('xeCJK'), 'TEX scripts: Greek and Cyrillic get Computer Modern Unicode');
    // An image the source embedded as a data: URI carries its bytes: carried inside the .tex like an extracted image.
    const dataPng = fs.readFileSync(path.join(__dirname, '..', 'docs', 'favicon.png')).toString('base64');
    const dataAst = await OfficeParser.parseOffice(Buffer.from(`![logo](data:image/png;base64,${dataPng})`), { fileType: 'md' } as any);
    const dataWarnings: any[] = [];
    const dataTex = (await dataAst.to('tex', { onWarning: (w: any) => dataWarnings.push(w) } as any)).value as string;
    const dataCarried = /\\includegraphics\[bb=[^\]]*\]\{(image-[0-9a-f]{8}\.pdf)\}/.exec(dataTex);
    const dataBlock = dataCarried && new RegExp(`\\\\begin\\{filecontents\\*\\}\\{${dataCarried[1].replace('.', '\\.')}\\}\\n([^]*?)\\n\\\\end\\{filecontents\\*\\}`).exec(dataTex);
    assert.ok(dataBlock && !dataWarnings.some(w => w.code === 'CONTENT_NOT_REPRESENTABLE' || w.code === 'IMAGES_NOT_BUNDLED'), 'TEX: a data: URI image is carried inside the .tex, not reported as a bad path');
    assertTextPdf(dataBlock![1], 'TEX data: image');
    // Parsed back it is a PNG with the same pixels: carried again, it gives the very same block.
    const dataBack = await OfficeParser.parseOffice(Buffer.from(dataTex), { fileType: 'tex', extractAttachments: true } as any);
    assert.deepStrictEqual(dataBack.attachments.map(a => a.mimeType), ['image/png'], 'TEX: the carried data: image parses back as a PNG');
    assert.ok(((await dataBack.to('tex')).value as string).includes(dataBlock![0]), 'TEX: the parsed-back PNG has the same pixels (it is carried as the same block)');
    const dataRefWarnings: any[] = [];
    const dataRefTex = (await dataAst.to('tex', { texConfig: { embedImages: false }, onWarning: (w: any) => dataRefWarnings.push(w) } as any)).value as string;
    assert.ok(/\\includegraphics\[[^\]]*\]\{images\/image\.png\}/.test(dataRefTex) && dataRefWarnings.some(w => w.code === 'IMAGES_NOT_BUNDLED' && w.message.includes('data: URIs')), 'TEX: without embedImages, IMAGES_NOT_BUNDLED says where a data: image\'s bytes are');
    // A fragment carries its images in its body, where filecontents* is also allowed.
    const dataFragment = (await dataAst.to('tex', { texConfig: { standalone: false } } as any)).value as string;
    assert.ok(dataFragment.includes(dataBlock![0]) && !dataFragment.includes('\\begin{document}'), 'TEX: a fragment carries its images in the body');
    const dataZip = unzipSync((await dataAst.to('tex', { texConfig: { bundle: true }, onWarning: () => {} } as any)).value as Uint8Array);
    assert.strictEqual(Buffer.from(dataZip['images/image.png']).toString('base64'), dataPng, 'TEX: the bundle holds the data: image\'s bytes');
    const fragment = (await (await OfficeParser.parseOffice(Buffer.from('中文 ✓ Ωμέγα'), { fileType: 'md' } as any)).to('tex', { texConfig: { standalone: false } } as any)).value as string;
    assert.ok(fragment.includes('% \\usepackage{iftex}\n') && fragment.includes('% \\iftutex\n%   % Greek and Cyrillic') && /\n% +\\usepackage\{xeCJK\}%\n/.test(fragment) && !fragment.includes('chosen above'),
        'TEX scripts: a fragment lists iftex and the script font setup its including document needs');
    assert.ok(koTex.includes('\\setCJKfallbackfamilyfont{\\CJKrmdefault}{{FandolSong-Regular.otf},{HaranoAjiMincho-Regular.otf}}'), 'TEX scripts: Hanja in Korean text fall back to the Chinese and Japanese fonts');
    const jaKo = await scriptTex('日本語のテキストです。カタカナも。한국어');
    assert.ok(jaKo.includes('AltFont={{Range="1100-"11FF, Font=UnBatang.ttf}') && jaKo.includes('{{FandolSong-Regular.otf},{UnBatang.ttf}}'), 'TEX scripts: Hangul in Japanese text comes from UnBatang (luatexja AltFont, xeCJK fallback)');
    // Structure holds.
    const sBraces = sLive.replace(/\\[{}]/g, '');
    assert.strictEqual((sBraces.match(/\{/g) || []).length, (sBraces.match(/\}/g) || []).length, 'TEX synthetic: braces balance');
    assert.strictEqual((sLive.match(/\\begin\{/g) || []).length, (sLive.match(/\\end\{/g) || []).length, 'TEX synthetic: environments balance');

    // An image's alternative text is graphicx's `alt` key, inline as well as on its own line, and parses back.
    const altMd = await OfficeParser.parseOffice(Buffer.from('Text ![A cat, [photo]](cat.png) more\n\n![Block alt](dog.png)\n'), { fileType: 'md' } as any);
    const altTex = (await altMd.to('tex', { onWarning: () => {} } as any)).value as string;
    assert.ok(altTex.includes('Text \\includegraphics[alt={A cat, {[}photo{]}}]{cat.png} more') && altTex.includes('\\includegraphics[alt={Block alt}]{dog.png}') && !altTex.includes('% alt:'), 'TEX: alt text as the alt key');
    assert.deepStrictEqual(collectAllNodes(await OfficeParser.parseOffice(Buffer.from(altTex), { fileType: 'tex', onWarning: () => {} } as any)).filter(n => n.type === 'image').map(n => (n.metadata as any).altText), ['A cat, [photo]', 'Block alt'], 'TEX round trip: inline and block alt text');

    // Column alignment is what most of a column's own cells have, a merged cell not voting; a cell that
    // differs from its column (or is merged) is written with its own alignment, so every cell parses back as it was.
    const alignedHtml = await OfficeParser.parseOffice(Buffer.from('<table><tr><th colspan="2" style="text-align:center">Group</th><th style="text-align:center">Total</th></tr><tr><td>Name</td><td style="text-align:right">12.5</td><td style="text-align:right">100</td></tr><tr><td>Other</td><td style="text-align:right">3.0</td><td style="text-align:right">7</td></tr></table>'), { fileType: 'html' });
    const alignedTex = (await alignedHtml.to('tex')).value as string;
    const cellAligns = (a: any) => collectAllNodes(a).filter(n => n.type === 'cell').map(n => (n.metadata as any).align ?? 'left');
    assert.ok(/\\begin\{longtable\}\{\|>\{\\raggedright\\arraybackslash\}p\{[^}]+\}\|>\{\\raggedleft\\arraybackslash\}p\{[^}]+\}\|>\{\\raggedleft/.test(alignedTex) && /\\multicolumn\{1\}\{>\{\\centering\\arraybackslash\}p\{[^}]+\}\|\}\{Total\}/.test(alignedTex), 'TEX: columns take their cells\' majority alignment; a differing cell its own');
    assert.deepStrictEqual(cellAligns(await OfficeParser.parseOffice(Buffer.from(alignedTex), { fileType: 'tex' })), cellAligns(alignedHtml), 'TEX round trip: every cell keeps its alignment');
    const mergedHeader = await OfficeParser.parseOffice(Buffer.from('\\begin{document}\\begin{tabular}{|l|c|r|}\\multicolumn{2}{|c|}{Header} & R \\\\ a & b & c \\\\\\end{tabular}\\end{document}'), { fileType: 'tex' });
    const mergedBack = await OfficeParser.parseOffice(Buffer.from((await mergedHeader.to('tex')).value as string), { fileType: 'tex' });
    assert.deepStrictEqual(cellAligns(mergedBack), ['center', 'right', 'left', 'center', 'right'], 'TEX round trip: a centred header over two columns leaves the first column left-aligned');

    // Wide tables continue below themselves in bands of 16 columns; a cell spanning a band boundary
    // is clipped, and its remainder is an empty merged cell that keeps the rule above it.
    const wideRow = (r: number) => ({ type: 'row', children: Array.from({ length: 40 }, (_, c) => ({ type: 'cell', metadata: { row: r, col: c }, children: [{ type: 'text', text: `r${r}c${c}` }] })) });
    const spanRow = { type: 'row', children: [{ type: 'cell', metadata: { row: 1, col: 0, colSpan: 20 }, children: [{ type: 'text', text: 'wide span' }] }] };
    const wide: any = { type: 'xlsx', metadata: {}, attachments: [], content: [{ type: 'table', children: [wideRow(0), spanRow, wideRow(2)] }], getImages: () => [] };
    const wtex = (await OfficeGenerator.generate(wide, 'tex', {})).value as string;
    assert.strictEqual((wtex.match(/\\begin\{longtable\}/g) || []).length, 3, 'TEX wide: 40 columns -> three bands');
    assert.ok(wtex.includes('\\textit{(continued: columns 17--32)}') && wtex.includes('\\textit{(continued: columns 33--40)}'), 'TEX wide: later bands are labelled');
    assert.ok(/\\multicolumn\{16\}\{[^}]*\}p\{[^}]+\}\|\}\{wide span\}/.test(wtex) && /\\multicolumn\{4\}\{\|[^}]*\}p\{[^}]+\}\|\}\{\}/.test(wtex), 'TEX wide: a span over the boundary is clipped, its remainder an empty merged cell');
    assert.ok(wtex.includes('r0c16') && wtex.includes('r2c39'), 'TEX wide: every cell lands in its band');

    // ── Tier 3: texConfig knobs and generic options ───────────────────────────
    const report = (await OfficeGenerator.generate(synthetic, 'tex', { texConfig: { documentClass: 'report', numberSections: true, format: 'Legal', landscape: true, margin: { top: '1in', right: 18, bottom: '2cm', left: 36 } } } as any)).value as string;
    assert.ok(report.includes('\\documentclass[\\officeparserdriver 11pt]{report}') && report.includes('\\chapter{Intro}') && !report.includes('secnumdepth'), 'TEX config: report class, chapters, numbering');
    assert.ok(report.includes('\\usepackage[paperwidth=612pt,paperheight=1008pt,landscape,top=72pt,bottom=56.69pt,left=36pt,right=18pt]{geometry}'), 'TEX config: paper, orientation and margins');
    const frag = (await OfficeGenerator.generate(synthetic, 'tex', { texConfig: { standalone: false } } as any)).value as string;
    assert.ok(frag.startsWith('% LaTeX fragment generated by officeParser. The including document needs, in its preamble:\n% \\usepackage{iftex}\n% \\usepackage{amsmath,amssymb}\n') && !frag.includes('\\documentclass') && frag.includes('\n\\definecolor{hex9A6700}{HTML}{9A6700}'), 'TEX config: fragment lists packages and defines its colours in the body');
    const plain = (await OfficeGenerator.generate(synthetic, 'tex', { includeFormatting: false, ignoreInternalLinks: true, generateIds: false, includeImages: 'none', includeCharts: false } as any)).value as string;
    assert.ok(!plain.includes('\\colorbox') && !plain.includes('\\fontsize') && !plain.includes('\\label{') && !plain.includes('\\hyperref[') && !plain.includes('\\includegraphics'), 'TEX config: formatting, labels, links and images all switch off');
    const skipped = (await OfficeGenerator.generate(synthetic, 'tex', { onNode: (n: any) => (n.type === 'admonition' ? false : n.type === 'embed' ? '\\CUSTOM{}' : undefined) } as any)).value as string;
    assert.ok(!skipped.includes('Careful') && skipped.includes('\\CUSTOM{}'), 'TEX config: onNode can drop a node or replace its output');

    // Beamer: slides become frames with titles and notes; content taller than a frame continues.
    const pptx = await OfficeParser.parseOffice(path.join(__dirname, 'files/test.pptx'), { extractAttachments: true });
    const beamer = (await pptx.to('tex')).value as string;
    assert.ok(/^% Generated by officeParser\.\n[\s\S]*?\n\\documentclass\[\\officeparserdriver aspectratio=169,\d+pt(,xcolor=table)?\]\{beamer\}/.test(beamer), 'TEX beamer: auto picks beamer for slides');
    assert.ok(beamer.includes('\\begin{frame}[allowframebreaks]\n\\frametitle{\\protect\\textcolor{hex17365D}{Demonstration of DOCX support in~calibre}}') && beamer.includes('\\setbeamertemplate{frametitle continuation}{}'), 'TEX beamer: frame with title');
    assert.ok(/\\end\{frame\}\n\\note\{/.test(beamer), 'TEX beamer: speaker notes as \\note');
    assert.ok(!beamer.includes('\\section') && !beamer.includes('longtable') && !beamer.includes('\\usepackage[hyperfootnotes=false]{hyperref}'), 'TEX beamer: no sectioning, longtable or second hyperref');
    const asArticle = (await pptx.to('tex', { texConfig: { documentClass: 'article' } } as any)).value as string;
    assert.ok(asArticle.includes('\\textbf{Speaker notes}') && /\\clearpage/.test(asArticle), 'TEX: slides in an article are separated by page breaks, notes kept');

    // Empty AST: a valid, empty document.
    const empty: any = { type: 'docx', metadata: {}, attachments: [], content: [], getImages: () => [] };
    const etex = (await OfficeGenerator.generate(empty, 'tex', {})).value as string;
    assert.ok(etex.endsWith('\\begin{document}\n\n\n\n\\end{document}\n') && !etex.includes('\\title'), 'TEX empty: bare document, no title');

    // mathtools' environments (dcases, pmatrix*, ...) are math, and load mathtools: refused, they were
    // written as text.
    const mathtoolsTex = (await (await OfficeParser.parseOffice(Buffer.from('$$\n|x| = \\begin{dcases} x & x \\ge 0 \\\\ -x & x < 0 \\end{dcases}\n$$\n\nand $\\begin{pmatrix*}[r] 1 & -2 \\end{pmatrix*}$.\n'), { fileType: 'md' })).to('tex')).value as string;
    assert.ok(mathtoolsTex.includes('\\usepackage{mathtools}') && mathtoolsTex.includes('\\begin{dcases}') && mathtoolsTex.includes('\\begin{pmatrix*}[r]') && !mathtoolsTex.includes('\\begin{verbatim}'), 'TEX: mathtools environments are math, with mathtools loaded');
}

async function testLatexParsing(): Promise<void> {
    // ── Tier 1: hand-written exhaustive fixture ───────────────────────────────
    const warnings: any[] = [];
    const ast = await OfficeParser.parseOffice(path.join(__dirname, 'files/exhaustive/latex.tex'), { onWarning: (w: any) => warnings.push(w) } as any);
    const nodes = collectAllNodes(ast);
    const texts = (n: OfficeContentNode) => (n.children || []).map(c => c.text || '').join('');
    assert.strictEqual(ast.type, 'tex', 'TEX parse: AST type');
    // Metadata from \title/\author/\date/\hypersetup and the class.
    assert.strictEqual(ast.metadata.title, 'Exhaustive LaTeX Test', 'TEX parse: \\title (with \\thanks dropped, \\LaTeX expanded)');
    assert.strictEqual(ast.metadata.author, 'Ada Lovelace, Alan Turing', 'TEX parse: \\author split on \\and');
    assert.strictEqual(ast.metadata.subject, 'Parser coverage', 'TEX parse: pdfsubject');
    assert.strictEqual(ast.metadata.keywords, 'latex, parser', 'TEX parse: pdfkeywords');
    assert.strictEqual(ast.metadata.language, 'en-GB', 'TEX parse: pdflang');
    assert.deepStrictEqual(ast.metadata.customProperties, { Reviewer: 'Grace Hopper' }, 'TEX parse: pdfinfo custom entries');
    assert.ok(ast.metadata.description?.startsWith('This document exercises the officeParser LaTeX parser, version 8.1.'), 'TEX parse: abstract -> description, macros expanded');
    assert.strictEqual(ast.metadata.nativeProperties?.documentClass, 'article', 'TEX parse: document class');
    assert.strictEqual(ast.metadata.formatting?.size, '12pt', 'TEX parse: class size is the body size');
    // Preamble text never reaches the content; text after \end{document} is ignored.
    assert.ok(!JSON.stringify(ast.content).includes('Text after the document end'), 'TEX parse: stops at \\end{document}');
    // \maketitle typesets the title block where it stands: a Title heading, then Author and Date lines.
    assert.deepStrictEqual(ast.content.slice(0, 3).map(n => [n.type, (n.metadata as any).style, (n.metadata as any).alignment, n.text]),
        [['heading', 'Title', 'center', 'Exhaustive LaTeX Test'], ['paragraph', 'Author', 'center', 'Ada Lovelace, Alan Turing'], ['paragraph', 'Date', 'center', '2024-03-15']], 'TEX parse: \maketitle -> title block');
    assert.strictEqual(collectAllNodes({ content: [ast.content[0]] } as any).flatMap(n => n.notes || [])[0]?.text, 'With thanks.', 'TEX parse: \thanks -> a footnote on the title');
    // Headings, levels, labels.
    const headings = nodes.filter(n => n.type === 'heading' && (n.metadata as any).style !== 'Title');
    assert.deepStrictEqual(headings.map(h => [(h.metadata as any).level, h.text]).slice(0, 5),
        [[1, 'Introduction'], [2, 'Links and references'], [3, 'Deep heading'], [4, 'Paragraph heading'], [1, 'Lists']], 'TEX parse: sectioning levels');
    assert.deepStrictEqual((headings[0].metadata as any).anchorIds, ['sec:intro'], 'TEX parse: \\label after a heading names it');
    // Inline formatting, colours, sizes, ligatures, accents, symbols, macros.
    const intro = nodes.find(n => n.type === 'paragraph' && (n.text || '').startsWith('Plain paragraph'))!;
    const fmt = (t: string) => intro.children!.find(c => c.text === t)?.formatting;
    assert.deepStrictEqual(fmt('bold'), { bold: true }, 'TEX parse: \\textbf');
    assert.deepStrictEqual(fmt('emphasis'), { italic: true }, 'TEX parse: \\emph');
    assert.deepStrictEqual(fmt('strikeout'), { strikethrough: true }, 'TEX parse: \\sout');
    assert.deepStrictEqual(fmt('typewriter'), { font: 'monospace' }, 'TEX parse: \\texttt');
    assert.deepStrictEqual(fmt('bold group'), { bold: true }, 'TEX parse: {\\bfseries ...} declaration scoped to its group');
    assert.deepStrictEqual(fmt('brand'), { color: '#1F6FEB' }, 'TEX parse: \\definecolor');
    assert.deepStrictEqual(fmt('soft brand'), { color: '#8FB7F5' }, 'TEX parse: \\colorlet with an xcolor mix');
    assert.deepStrictEqual(fmt('highlighted'), { backgroundColor: '#FFFF00' }, 'TEX parse: \\hl');
    assert.deepStrictEqual(fmt('large'), { size: '14.4pt' }, 'TEX parse: \\large against the 12pt class');
    assert.deepStrictEqual(fmt('twenty'), { size: '20pt' }, 'TEX parse: \\fontsize');
    const lig = nodes.find(n => n.type === 'paragraph' && (n.text || '').startsWith('Ligatures'))!.text!;
    for (const frag of ['“double”', '‘single’', 'en–dash', 'em—dash', 'don’t', 'café', 'naïve', 'ça', 'š', 'í', 'å', 'ß', 'Ø', '§3', '©', '50%', '$5', '& more', '# tag', 'a_b', '\\', '…', 'non breaking', 'Hello, World!', 'Hi, there!', '\\raw{x}', 'hy­phen', 'thin space', 'forced\nline']) {
        assert.ok(lig.includes(frag), `TEX parse: text contains ${JSON.stringify(frag)}`);
    }
    // Links, references, citations, notes, comments.
    const links = nodes.filter(n => (n.metadata as any)?.link);
    assert.ok(links.some(l => (l.metadata as any).link === 'https://example.com/a_b?x=1&y=2' && l.text === 'the example site'), 'TEX parse: \\href with escaped URL characters');
    assert.ok(links.some(l => (l.metadata as any).link === '#sec:intro' && l.text === '1'), 'TEX parse: \\ref resolves to the section number');
    assert.ok(links.some(l => (l.metadata as any).link === '#sec:links' && l.text === 'Links and references'), 'TEX parse: \\nameref resolves to the heading title');
    assert.ok(links.some(l => (l.metadata as any).link === '#tab:main' && l.text === '1') && links.some(l => (l.metadata as any).link === '#fig:diagram' && l.text === '1'), 'TEX parse: table/figure refs resolve to their numbers');
    assert.deepStrictEqual(nodes.filter(n => (n.metadata as any)?.citationKey).map(n => (n.metadata as any).citationKey), ['knuth1984', 'lamport1994', 'mittelbach2004'], 'TEX parse: \\cite and \\citep keys');
    const notes = nodes.flatMap(n => n.notes || []);
    assert.deepStrictEqual(notes.map(n => (n.metadata as any).noteType), ['footnote', 'footnote', 'endnote'], 'TEX parse: \\thanks, \\footnote and \\endnote');
    assert.strictEqual(notes[1].text, 'The footnote body.', 'TEX parse: footnote body');
    const comment = nodes.flatMap(n => n.comments || [])[0];
    assert.deepStrictEqual([comment?.text, (comment?.metadata as any)?.author, (comment?.metadata as any)?.date], ['Please expand this section.\nIt needs more detail.', 'Reviewer', '2024-02-02'], 'TEX parse: comment lines read back as a comment');
    // Lists.
    const items = nodes.filter(n => n.type === 'list');
    assert.deepStrictEqual(items.slice(0, 6).map(i => [(i.metadata as any).listType, (i.metadata as any).indentation, i.text]),
        [['unordered', 0, 'First bullet'], ['unordered', 0, 'Second bullet with a nested list:'], ['ordered', 1, 'Nested one'], ['ordered', 1, 'Nested two'], ['unordered', 0, 'Open task'], ['unordered', 0, 'Done task']], 'TEX parse: nested lists');
    assert.deepStrictEqual([(items[4].metadata as any).isTask, (items[4].metadata as any).checked, (items[5].metadata as any).checked], [true, false, true], 'TEX parse: task items');
    assert.strictEqual((items.find(i => i.text === 'Fifth item')!.metadata as any).itemIndex, 4, 'TEX parse: \\setcounter{enumi} continues numbering');
    const dl = nodes.find(n => n.type === 'definitionList')!;
    assert.deepStrictEqual(dl.children!.map(c => [c.type, c.text]), [['definitionTerm', 'Term A'], ['definitionDescription', 'Description A.'], ['definitionTerm', 'Term [B]'], ['definitionDescription', 'Description B.']], 'TEX parse: description list');
    // Tables.
    const [booktabs, longtable] = nodes.filter(n => n.type === 'table');
    assert.deepStrictEqual((booktabs.metadata as any).anchorIds, ['tab:main'], 'TEX parse: float label names its table');
    const cellOf = (t: OfficeContentNode, r: number, c: number) => t.children![r].children!.find(x => (x.metadata as any).col === c)!;
    assert.deepStrictEqual(cellOf(booktabs, 0, 1).metadata, { row: 0, col: 1, align: 'center', style: 'header' }, 'TEX parse: header row above \\midrule, column alignment');
    assert.deepStrictEqual([(cellOf(booktabs, 1, 0).metadata as any).colSpan, (cellOf(booktabs, 2, 0).metadata as any).rowSpan], [2, 2], 'TEX parse: \\multicolumn / \\multirow');
    assert.strictEqual(booktabs.children![3].children!.length, 2, 'TEX parse: a \\multirow-covered cell is not a cell of its own');
    assert.strictEqual((cellOf(booktabs, 3, 2).metadata as any).backgroundColor, '#FFFF00', 'TEX parse: \\cellcolor');
    assert.ok(longtable.children![0].children!.every(c => (c.metadata as any).style === 'header') && longtable.children!.length === 3, 'TEX parse: longtable \\endhead rows are header rows');
    // Figures, code, math.
    const img = nodes.find(n => n.type === 'image')!;
    assert.deepStrictEqual(img.metadata, { width: '50%', url: 'figures/diagram', align: 'center', anchorIds: ['fig:diagram'] }, 'TEX parse: \\includegraphics in a centred figure');
    const code = nodes.filter(n => n.type === 'code');
    assert.deepStrictEqual(code[0].metadata, { language: 'python' }, 'TEX parse: lstlisting language');
    assert.strictEqual(code[1].text, 'raw \\text{is} kept & literal', 'TEX parse: verbatim body is literal');
    const math = code.filter(c => (c.metadata as any).math).map(c => [(c.metadata as any).math, c.text]);
    assert.deepStrictEqual(math.slice(0, 4), [['inline', 'E = mc^2'], ['inline', 'a^2 + b^2'], ['inline', 'x \\in \\mathbb{R}'], ['block', '\\int_0^1 x\\,dx = \\frac{1}{2}']], 'TEX parse: math, with user macros expanded');
    assert.ok(math.some(([m, t]) => m === 'block' && t!.startsWith('\\begin{align*}')), 'TEX parse: a multi-line environment keeps its environment');
    // Blocks, bibliography, header/footer, warnings.
    assert.ok(nodes.some(n => n.type === 'paragraph' && (n.metadata as any)?.alignment === 'center' && n.text === 'Centered text.'), 'TEX parse: center environment');
    assert.ok(nodes.some(n => n.type === 'paragraph' && (n.metadata as any)?.style === 'Quote' && texts(n).startsWith('Tip: Custom')), 'TEX parse: user environment expanded');
    assert.ok(nodes.some(n => n.type === 'break' && (n.metadata as any)?.breakType === 'thematic'), 'TEX parse: \\rule -> thematic break');
    const bib = items.filter(i => (i.metadata as any).anchorIds?.[0] === 'knuth1984');
    assert.ok(bib.length === 1 && headings.some(h => h.text === 'References'), 'TEX parse: thebibliography -> References list with citation anchors');
    assert.deepStrictEqual(ast.auxiliary?.headers?.[0].children!.map(p => [p.text, (p.metadata as any).alignment]), [['Running Title', 'left'], ['Draft', 'right']], 'TEX parse: fancyhdr header fields');
    assert.strictEqual(ast.auxiliary?.footers, undefined, 'TEX parse: a footer holding only \\thepage is no footer');
    assert.deepStrictEqual(warnings.map(w => w.code).sort(), ['LATEX_CONSTRUCT_NOT_INTERPRETED', 'LATEX_FILE_NOT_FOUND'], 'TEX parse: warnings');
    assert.strictEqual(warnings.find(w => w.code === 'LATEX_CONSTRUCT_NOT_INTERPRETED').message, `The LaTeX input uses 'tikzpicture environment', which the parser does not interpret. Text inside it was kept where there was any; drawing environments (such as TikZ pictures) were omitted.`, 'TEX parse: LATEX_CONSTRUCT_NOT_INTERPRETED exact message');
    assert.strictEqual(warnings.find(w => w.code === 'LATEX_FILE_NOT_FOUND').message, `The LaTeX input references files the parser could not read ('figures/diagram', 'chapters/missing'). A .tex file holds only the files it carries in filecontents blocks; parse the project as a .zip (for example an Overleaf download) to include the others. Images were kept as references to their path.`, 'TEX parse: LATEX_FILE_NOT_FOUND exact message');

    // ── Tier 2: project zip (includes and images) ─────────────────────────────
    const zipOf = (files: Record<string, string | Uint8Array>) => Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, typeof v === 'string' ? strToU8(v) : v]))));
    const project = zipOf({
        'paper/main.tex': '\\documentclass{article}\\graphicspath{{img/}}\\begin{document}\\input{sections/intro}\\includegraphics{logo}\\end{document}',
        'paper/sections/intro.tex': '\\section{Intro} Included text.',
        'paper/img/logo.png': decodeBase64(TINY_PNG_B64),
    });
    const pAst = await OfficeParser.parseOffice(project, { extractAttachments: true });
    assert.strictEqual(pAst.type, 'tex', 'TEX project: a zip with a LaTeX main file is detected as tex');
    assert.ok(collectAllNodes(pAst).some(n => n.type === 'heading' && n.text === 'Intro'), 'TEX project: \\input resolved relative to the main file');
    const pImg = collectAllNodes(pAst).find(n => n.type === 'image')!;
    assert.strictEqual((pImg.metadata as any).attachmentName, 'paper/img/logo.png', 'TEX project: image found through \\graphicspath');
    assert.deepStrictEqual(pAst.attachments.map(a => [a.name, a.mimeType]), [['paper/img/logo.png', 'image/png']], 'TEX project: image extracted as an attachment');
    // The shipped samples: test.tex (the visualizer's LaTeX sample) carries the image it shows, and
    // latex-project.zip holds the same document with the image as a file.
    const sample = await OfficeParser.parseOffice(path.join(__dirname, 'files/latex-project.zip'), { extractAttachments: true } as any);
    const loneWarnings: any[] = [];
    const lone = await OfficeParser.parseOffice(path.join(__dirname, 'files/test.tex'), { extractAttachments: true, onWarning: (w: any) => loneWarnings.push(w) } as any);
    assert.strictEqual(sample.type, 'tex', 'TEX project: latex-project.zip is LaTeX');
    assert.deepStrictEqual(sample.attachments.map(a => [a.name, a.mimeType]), [['images/image.jpg', 'image/jpeg']], 'TEX project: the sample image is attached');
    assert.deepStrictEqual(lone.attachments.map(a => [a.mimeType, a.data === sample.attachments[0].data]), [['image/jpeg', true]], 'TEX sample: test.tex carries the very image the project zip holds as a file');
    assert.ok(!loneWarnings.some(w => ['LATEX_FILE_NOT_FOUND', 'LATEX_CONSTRUCT_NOT_INTERPRETED'].includes(w.code)), `TEX sample: test.tex reads with no missing file or uninterpreted construct (${loneWarnings.map(w => w.code).join(', ')})`);
    // Identical but for the image's name.
    const withoutImages = (t: unknown) => String(t).replace(/\[Image: [^\]]*\]/g, '[Image]');
    assert.strictEqual(withoutImages((await sample.to('text')).value), withoutImages((await lone.to('text')).value), 'TEX project: the sample holds the same document as test.tex');

    // filecontents: compiling writes the file (keeping one already there unless told to overwrite), so
    // \input and \includegraphics read it, as TeX would; a name leaving the project writes nothing.
    const fcWarnings: any[] = [];
    const fc = await OfficeParser.parseOffice(Buffer.from([
        '\\documentclass{article}',
        '\\begin{filecontents*}{part.tex}', '\\section{Written}', 'From a file.   ', '\\end{filecontents*}',
        '\\begin{filecontents}{part.tex}', '\\section{Ignored}', '\\end{filecontents}',
        '\\begin{filecontents*}[overwrite]{note.tex}', 'First.', '\\end{filecontents*}',
        '\\begin{filecontents*}[overwrite]{note.tex}', 'Second.', '\\end{filecontents*}',
        '\\begin{filecontents*}{../escape.tex}', 'Outside.', '\\end{filecontents*}',
        '\\begin{document}', '\\input{part}', '\\input{note}', '\\input{../escape}', '\\end{document}',
    ].join('\n')), { fileType: 'tex', onWarning: (w: any) => fcWarnings.push(w) } as any);
    const fcText = (await fc.to('text')).value as string;
    assert.ok(/Written/.test(fcText) && /From a file\./.test(fcText) && /Second\./.test(fcText) && !/Ignored|First\.|Outside\./.test(fcText), `TEX filecontents: written files are read, an existing one kept unless overwritten (${JSON.stringify(fcText)})`);
    assert.deepStrictEqual(fcWarnings.map(w => w.code), ['LATEX_FILE_NOT_FOUND'], 'TEX filecontents: interpreted, and only the file outside the project is missing');
    // \RequirePackage (as the generator loads iftex) loads a package as \usepackage does.
    const required = await OfficeParser.parseOffice(Buffer.from('\\RequirePackage{iftex}\\documentclass{article}\\usepackage{graphicx}\\makeatletter\\@ifpackageloaded{iftex}{\\def\\x{yes}}{\\def\\x{no}}\\makeatother\\begin{document}\\x\\end{document}'), { fileType: 'tex' } as any);
    assert.deepStrictEqual([(required.metadata as any).nativeProperties.packages, ((await required.to('text')).value as string).trim()], [['iftex', 'graphicx'], 'yes'], 'TEX: \\RequirePackage loads a package as \\usepackage does');
    // A PDF figure that is more than a picture (a vector drawing, here text) stays a PDF.
    const vectorPdf = '%PDF-1.5\n1 0 obj\n<< /Type /Page /Contents 2 0 R >>\nendobj\n2 0 obj\n<< /Length 23 >>\nstream\nBT /F1 12 Tf (Hi) Tj ET\nendstream\nendobj\n%%EOF\n';
    const vector = await OfficeParser.parseOffice(Buffer.from(zipSync({ 'main.tex': strToU8('\\documentclass{article}\\begin{document}\\includegraphics{fig}\\end{document}'), 'fig.pdf': strToU8(vectorPdf) })), { extractAttachments: true } as any);
    assert.deepStrictEqual(vector.attachments.map(a => [a.name, a.mimeType]), [['fig.pdf', 'application/pdf']], 'TEX project: a vector PDF figure stays a PDF');

    // ── Tier 3: beamer ────────────────────────────────────────────────────────
    const deck = await OfficeParser.parseOffice(Buffer.from('\\documentclass{beamer}\\begin{document}\\frame{\\titlepage}\\begin{frame}{First}{Sub}\\begin{itemize}\\item<1-> One\\end{itemize}\\note{Say hi.}\\end{frame}\\begin{frame}[fragile]\\frametitle{Code}\\begin{block}{Idea}Body\\end{block}\\end{frame}\\end{document}'), { fileType: 'tex' });
    assert.deepStrictEqual(deck.content.map(s => [s.type, (s.metadata as any).slideNumber]), [['slide', 1], ['slide', 2]], 'TEX beamer: frames -> slides (the title frame is not a slide)');
    assert.deepStrictEqual(deck.content[0].children!.map(c => [c.type, c.text]), [['heading', 'First'], ['heading', 'Sub'], ['list', 'One']], 'TEX beamer: frame title, subtitle, overlay spec dropped');
    assert.strictEqual(deck.content[0].notes?.[0].text, 'Say hi.', 'TEX beamer: \\note -> slide notes');
    assert.deepStrictEqual(deck.content[1].children!.map(c => c.type), ['heading', 'admonition'], 'TEX beamer: block -> admonition');
    const titled = await OfficeParser.parseOffice(Buffer.from('\\documentclass{beamer}\\title{Deck}\\subtitle{Sub}\\author{Ann \\and Bob}\\date{\\today}\\begin{document}\\begin{frame}\\titlepage\\end{frame}\\end{document}'), { fileType: 'tex', texParserConfig: { today: 'May 1, 2024' } } as any);
    assert.deepStrictEqual(titled.content[0].children!.map(c => [(c.metadata as any).style, c.text]), [['Title', 'Deck'], ['Subtitle', 'Sub'], ['Author', 'Ann, Bob'], ['Date', 'May 1, 2024']],
        'TEX beamer: \\titlepage -> a title slide, its \\today date as texParserConfig.today sets it');
    const deckTex = (await titled.to('tex')).value as string;
    assert.ok(/\\title\{Deck\}\n\\subtitle\{Sub\}\n\\author\{Ann, Bob\}\n\\date\{May 1, 2024\}/.test(deckTex) && /\\begin\{frame\}\[allowframebreaks\]\n\\relax\n\\titlepage\n\\end\{frame\}/.test(deckTex), 'TEX beamer: a title slide regenerates as \\titlepage');

    // A \section between frames stays a \section (not a frame of its own), and a frame's subtitle a
    // \framesubtitle, so the deck parses back with the same slides.
    const sectioned = await OfficeParser.parseOffice(Buffer.from('\\documentclass{beamer}\\begin{document}\\section{Intro}\\subsection{Start}\\begin{frame}{First}{Sub}Body.\\end{frame}\\section{End}\\begin{frame}\\frametitle{Last}\\framesubtitle{Bye}\\end{frame}\\end{document}'), { fileType: 'tex' });
    const sectionedTex = (await sectioned.to('tex')).value as string;
    assert.ok(/\\section\{Intro\}[^\n]*\n\n\\subsection\{Start\}[^\n]*\n\n\\begin\{frame\}\[allowframebreaks\]\n\\frametitle\{First\}\n\\framesubtitle\{Sub\}/.test(sectionedTex) && (sectionedTex.match(/\\begin\{frame\}/g) || []).length === 2, 'TEX beamer: sections between frames and frame subtitles are written as such');
    const sectionedBack = await OfficeParser.parseOffice(Buffer.from(sectionedTex), { fileType: 'tex' });
    const outline = (a: any) => a.content.map((n: any) => n.type === 'slide' ? ['slide', n.children.filter((c: any) => c.type === 'heading').map((c: any) => [c.metadata.level, c.text])] : [n.type, n.metadata?.level, n.text]);
    assert.deepStrictEqual(outline(sectionedBack), outline(sectioned), 'TEX round trip: a sectioned deck keeps its slides, titles and subtitles');

    // A \label after \phantomsection names what follows it (the generator's anchor before a block).
    const anchored = await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\begin{document}First.\n\n\\phantomsection\\label{next}Second.\n\n\\section{S}\\label{sec}\\end{document}'), { fileType: 'tex' });
    assert.deepStrictEqual(anchored.content.map(n => [n.text, (n.metadata as any).anchorIds]), [['First.', undefined], ['Second.', ['next']], ['S', ['sec']]], 'TEX parse: \\phantomsection anchors the following block');

    // A resolved reference updates the text of everything that holds it, so text output and chunks
    // show the number, not the label.
    const refDoc = await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\begin{document}\\section{Sec}\\label{s} See \\ref{s} and \\nameref{s}.\n\\begin{itemize}\\item item \\ref{s}\\end{itemize}\n\\begin{tabular}{c}cell \\ref{s}\\\\\\end{tabular}\\end{document}'), { fileType: 'tex' } as any);
    const refTexts = collectAllNodes(refDoc).filter(n => n.type === 'paragraph' || n.type === 'list' || n.type === 'cell').map(n => n.text);
    assert.ok(refTexts.includes('See 1 and Sec.') && refTexts.includes('item 1') && refTexts.includes('cell 1') && !refTexts.some(t => /\bs\b/.test(t || '')), `TEX parse: reference text refreshed (${JSON.stringify(refTexts)})`);
    assert.ok(((await refDoc.to('chunks')).value as any[]).some(c => c.text === 'See 1 and Sec.'), 'TEX parse: chunks see the resolved reference');

    // `\IfFileExists{file}{then}{else}`: nothing is read from disk, so the document as written (then) is read.
    const ifExists = await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\begin{document}\\IfFileExists{fig.png}{\\includegraphics{fig.png}}{\\fbox{none}}\\end{document}'), { fileType: 'tex', onWarning: () => {} } as any);
    assert.deepStrictEqual(collectAllNodes(ifExists).filter(n => n.type === 'image').map(n => (n.metadata as any).url), ['fig.png'], 'TEX parse: \\IfFileExists reads its then-branch');

    // ── Tier 3b: engines, languages, packages and classes ─────────────────────
    const texOf = async (src: string | Buffer, extra: any = {}) => {
        const w: any[] = [];
        const a = await OfficeParser.parseOffice(Buffer.isBuffer(src) ? src : Buffer.from(src), { fileType: 'tex', onWarning: (x: any) => w.push(x), ...extra } as any);
        return { ast: a, warnings: w, paras: collectAllNodes(a).filter(n => n.type === 'paragraph' || n.type === 'heading' || n.type === 'list').map(n => n.text) };
    };
    const runsOf = (n: OfficeContentNode | undefined) => (n?.children ?? []).map(c => [c.text, c.formatting ?? {}]);

    // Conditionals are decided as pdfLaTeX would: only the branch taken is read.
    const cond = await texOf(String.raw`\documentclass{article}
\newif\ifdraft \drafttrue
\begin{document}
A\ifXeTeX XE\else\ifLuaTeX LUA\else PDF\fi\fi{} B\ifPDFTeX{} pdf\else{} other\fi{}
C\ifdraft{} draft\else{} final\fi{} \draftfalse D\ifdraft{} draft\else{} final\fi{}
E\ifdefined\directlua{} lua\else{} nolua\fi{} F\ifx\XeTeXversion\undefined{} nonxe\else{} xe\fi{} G\unless\ifXeTeX{} unless\fi{}
H\iffalse{} hidden \ifnum1>2 x\fi{} hidden\else{} shown\fi{} I\ifnum\pdfoutput>0{} pdfout\else{} dvi\fi{} J\ifnum 1>2 {} big\else{} small\fi
\end{document}`);
    assert.strictEqual(cond.paras[0], 'APDF B pdf C draft D final E nolua F nonxe G unless H shown I pdfout J small', `TEX parse: engine, \\newif, \\ifdefined, \\ifx, \\unless, \\ifnum conditionals (${cond.paras[0]})`);
    assert.strictEqual(cond.warnings.length, 0, 'TEX parse: decided conditionals are not reported');
    const undecided = await texOf(String.raw`\begin{document}\ifnum\value{page}>1 X\else Y\fi\end{document}`);
    assert.ok(undecided.paras[0]?.includes('X') && undecided.paras[0]?.includes('Y') && undecided.warnings.some(w => JSON.stringify(w).includes('\\\\ifnum')), 'TEX parse: a test that cannot be decided reads both branches and is reported');
    const pkgTests = await texOf(String.raw`\documentclass{article}\usepackage{hyperref,ifthen,etoolbox}
\newboolean{long}\setboolean{long}{true}\newtoggle{t}\toggletrue{t}
\begin{document}\makeatletter\@ifpackageloaded{hyperref}{HYP}{NOHYP} \@ifundefined{directlua}{NOLUA}{LUA}\makeatother{}
\ifthenelse{\boolean{long}}{LONG}{SHORT} \ifthenelse{\NOT\boolean{long}}{NOT}{YES} \iftoggle{t}{TOG}{NOTOG} \ifthenelse{\equal{a}{b}}{EQ}{NE}\end{document}`);
    assert.strictEqual(pkgTests.paras[0], 'HYP NOLUA LONG YES TOG NE', `TEX parse: \\@ifpackageloaded, \\@ifundefined, ifthen booleans, etoolbox toggles (${pkgTests.paras[0]})`);

    // Languages: the switch commands keep only their text; the main language becomes metadata.language.
    const langs = await texOf(String.raw`\documentclass[ngerman]{article}
\usepackage[english,main=french]{babel}
\babeltags{de = ngerman}
\begin{document}
Hello \foreignlanguage{german}{Hallo} \textde{Welt} \textfrench{Bonjour} \begin{otherlanguage}{spanish}Hola\end{otherlanguage} \begin{german}Tag\end{german}.
\end{document}`);
    assert.deepStrictEqual([langs.paras, langs.ast.metadata.language, langs.warnings.length], [['Hello Hallo Welt Bonjour Hola Tag.'], 'fr', 0], 'TEX parse: \\foreignlanguage, \\text<lang>, language environments, babel main=');
    assert.strictEqual((await texOf(String.raw`\documentclass[british]{article}\usepackage{babel}\begin{document}x\end{document}`)).ast.metadata.language, 'en-GB', 'TEX parse: babel takes the language from the class options');
    assert.strictEqual((await texOf(String.raw`\usepackage{polyglossia}\setdefaultlanguage[variant=american]{english}\begin{document}x\end{document}`)).ast.metadata.language, 'en-US', 'TEX parse: polyglossia \\setdefaultlanguage (with a variant)');
    assert.strictEqual((await texOf(String.raw`\usepackage[german]{babel}\usepackage{hyperref}\hypersetup{pdflang=en-US}\begin{document}x\end{document}`)).ast.metadata.language, 'en-US', 'TEX parse: pdflang wins over babel');

    // LaTeX 2.09 font switches (each starts from the normal font), and \documentstyle.
    const oldFonts = await texOf(String.raw`\documentstyle[12pt]{article}\begin{document}{\bf bold {\it italic} bold} {\tt mono} {\sl slanted} {\sf sans}\end{document}`);
    assert.deepStrictEqual(runsOf(collectAllNodes(oldFonts.ast).find(n => n.type === 'paragraph')), [['bold ', { bold: true }], ['italic', { italic: true }], [' bold', { bold: true }], [' ', {}], ['mono', { font: 'monospace' }], [' ', {}], ['slanted', { italic: true }], [' ', {}], ['sans', { font: 'sans-serif' }]], 'TEX parse: \\bf \\it \\tt \\sl \\sf');
    assert.deepStrictEqual([oldFonts.ast.metadata.nativeProperties?.documentClass, oldFonts.ast.metadata.formatting?.size], ['article', '12pt'], 'TEX parse: \\documentstyle names the class and its options');

    // xparse commands and environments (the LaTeX kernel's \NewDocumentCommand), and expl3 code skipped.
    const xp = await texOf(String.raw`\documentclass{article}
\NewDocumentCommand{\greet}{s O{World} m}{\IfBooleanTF{#1}{Hi}{Hello}, #2 and #3!}
\NewDocumentCommand\opt{o m}{\IfNoValueTF{#1}{[#2]}{(#1:#2)}}
\NewDocumentCommand\dl{d() m}{\IfValueT{#1}{<#1>}#2}
\NewDocumentEnvironment{boxed}{O{Note} m}{\textbf{#1 #2:} }{ (end #1)}
\NewDocumentEnvironment{wrap}{m +b}{[#1|#2]}{}
\ExplSyntaxOn
\cs_new:Npn \my_fn:n #1 { \tl_upper_case:n {#1} }
\ExplSyntaxOff
\begin{document}
\greet{you} \greet*[Earth]{me} \opt{a} \opt[b]{c} \dl(x){y} \dl{z} \begin{boxed}{Title}Body\end{boxed} \begin{wrap}{W}inner\end{wrap}
\end{document}`);
    assert.deepStrictEqual([xp.paras[0], xp.warnings.length], ['Hello, World and you! Hi, Earth and me! [a] (b:c) <x>y z Note Title: Body (end Note) [W|inner]', 0], 'TEX parse: \\NewDocumentCommand/\\NewDocumentEnvironment signatures, \\IfBooleanTF/\\IfNoValueTF, expl3 skipped');
    assert.ok((await texOf(String.raw`\NewDocumentCommand\e{e{^_}}{x}\begin{document}\e\end{document}`)).warnings.some(w => JSON.stringify(w).includes('argument specification')), 'TEX parse: an unsupported xparse signature is reported');

    // Theorems: numbered (within sections, shared counters), styles, notes, proofs, and \ref to them.
    const thm = await texOf(String.raw`\documentclass{article}\usepackage{amsthm}
\newtheorem{theorem}{Theorem}[section]
\newtheorem{lemma}[theorem]{Lemma}
\theoremstyle{definition}\newtheorem{definition}{Definition}\newtheorem*{remark*}{Remark}
\theoremstyle{remark}\newtheorem{note}{Note}
\begin{document}
\section{One}
\begin{theorem}[Fermat]\label{thm:f}No solutions.\end{theorem}
\begin{lemma}\label{lem:a}A lemma.\end{lemma}
\begin{definition}A word is \emph{nice}.\label{def:n}\end{definition}
\begin{remark*}Unnumbered.\end{remark*}
\begin{note}Remark style.\end{note}
\begin{proof}Trivial.\end{proof}
\begin{proof}[Proof of Theorem~\ref{thm:f}]\begin{itemize}\item one\end{itemize}\end{proof}
\section{Two}
\begin{theorem}Again.\end{theorem}
See \ref{thm:f}, \ref{lem:a}, \ref{def:n}.
\end{document}`);
    const thmParas = collectAllNodes(thm.ast).filter(n => n.type === 'paragraph');
    assert.deepStrictEqual(thmParas.map(p => p.text), ['Theorem 1.1 (Fermat). No solutions.', 'Lemma 1.2. A lemma.', 'Definition 1. A word is nice.', 'Remark. Unnumbered.', 'Note 1. Remark style.', 'Proof. Trivial. □', 'Proof of Theorem\u00A01.1.', '□', 'Theorem 2.1. Again.', 'See 1.1, 1.2, 1.'], 'TEX parse: theorem numbering, notes, proofs, references');
    assert.deepStrictEqual(runsOf(thmParas[0]), [['Theorem 1.1', { bold: true }], [' (Fermat)', {}], ['.', { bold: true }], [' ', {}], ['No solutions.', { italic: true }]], 'TEX parse: plain-style theorem head bold, body italic');
    assert.deepStrictEqual([runsOf(thmParas[2])[0], runsOf(thmParas[4])[0], runsOf(thmParas[5])[0]], [['Definition 1.', { bold: true }], ['Note 1.', { italic: true }], ['Proof.', { italic: true }]], 'TEX parse: definition, remark and proof heads');
    assert.deepStrictEqual((thmParas[0].metadata as any).anchorIds, ['thm:f'], 'TEX parse: a theorem label anchors the theorem');
    const llncs = await texOf(String.raw`\documentclass{llncs}\begin{document}\begin{theorem}T\end{theorem}\begin{lemma}L\end{lemma}\begin{theorem}T2\end{theorem}\keywords{First \and Second}\end{document}`);
    assert.deepStrictEqual([llncs.paras, llncs.ast.metadata.keywords], [['Theorem 1. T', 'Lemma 1. L', 'Theorem 2. T2', 'Keywords: First, Second'], 'First, Second'], 'TEX parse: theorem environments a class provides, and \\keywords');
    const unseen = await texOf(String.raw`\documentclass{exam}\usepackage{mythms}\begin{document}\begin{theorem}T\end{theorem}\begin{solution}S\end{solution}\end{document}`);
    assert.deepStrictEqual(unseen.paras, ['Theorem. T', 'S'], 'TEX parse: a theorem defined out of sight is headed but unnumbered; a class\'s own solution environment is not a theorem');
    const llncsAbs = await texOf(String.raw`\documentclass{llncs}\begin{document}\begin{abstract}The abstract.\keywords{A \and B}\end{abstract}\end{document}`);
    assert.deepStrictEqual([llncsAbs.ast.metadata.description, llncsAbs.ast.metadata.keywords, llncsAbs.paras], ['The abstract.', 'A, B', ['The abstract.', 'Keywords: A, B']], 'TEX parse: keywords inside the abstract are not part of the description');
    const mailInAuthor = await texOf(String.raw`\documentclass{article}\title{T}\author{Ann\thanks{x} \email{ann@x.org}}\begin{document}\maketitle\end{document}`);
    assert.deepStrictEqual([mailInAuthor.ast.metadata.author, (mailInAuthor.ast.metadata.nativeProperties as any).emails], ['Ann', ['ann@x.org']], 'TEX parse: an \\email inside \\author is recorded once, and is not part of the name');
    const cjkEnv = await texOf(String.raw`\documentclass{article}\usepackage{CJKutf8}\begin{document}\begin{CJK*}{UTF8}{gbsn}中文\end{CJK*}\end{document}`);
    assert.deepStrictEqual([cjkEnv.paras, cjkEnv.warnings.length], [['中文'], 0], 'TEX parse: the CJK package\'s environment keeps only its text');
    const thmDeck = await texOf(String.raw`\documentclass{beamer}\begin{document}\begin{frame}{F}\begin{theorem}[Name]Body\end{theorem}\end{frame}\end{document}`);
    assert.deepStrictEqual((thmDeck.ast.content[0].children![1].metadata as any), { admonitionType: 'note', title: 'Theorem (Name)' }, 'TEX beamer: a theorem is a titled block');

    // References resolve to numbers with or without internal links, including enumerate items.
    const refSrc = String.raw`\begin{document}\section{Intro}\label{sec:i}
\begin{enumerate}\item a\label{it:a} \item b \begin{enumerate}\item c\label{it:c}\end{enumerate}\end{enumerate}
\begin{figure}\includegraphics{x.png}\caption{Cap}\label{fig:x}\end{figure}
\begin{equation}x\label{eq:x}\end{equation}
See \textbf{\ref{sec:i}} \ref{it:a} \ref{it:c} \ref{fig:x} \eqref{eq:x}.\end{document}`;
    for (const ignoreInternalLinks of [false, true]) {
        const r = await texOf(refSrc, { ignoreInternalLinks });
        assert.strictEqual(r.paras[r.paras.length - 1], 'See 1 1 2a 1 (1).', `TEX parse: references resolve (ignoreInternalLinks: ${ignoreInternalLinks})`);
        const see = collectAllNodes(r.ast).filter(n => n.type === 'paragraph').pop()!;
        assert.ok(see.children!.some(c => c.text === '1' && c.formatting?.bold), `TEX parse: a reference keeps its run formatting (ignoreInternalLinks: ${ignoreInternalLinks})`);
    }

    // Class commands: amsart author information, IEEEtran author blocks and keywords, KOMA-Script, epigraphs.
    const ams = await texOf(String.raw`\documentclass{amsart}\title{T}\author{A. Author}\address{Dept. of Math}\email{a@b.org}\urladdr{https://x.org}\subjclass[2020]{Primary 35K05}\keywords{heat, equation}\begin{document}\maketitle Text.\end{document}`);
    const amsNative = ams.ast.metadata.nativeProperties as any;
    assert.deepStrictEqual([amsNative.addresses, amsNative.emails, amsNative.urls, amsNative.subjectClassification, ams.ast.metadata.keywords, ams.warnings.length],
        [['Dept. of Math'], ['a@b.org'], ['https://x.org'], { codes: 'Primary 35K05', scheme: 'MSC2020' }, 'heat, equation', 0], 'TEX parse: amsart \\address, \\email, \\urladdr, \\subjclass, \\keywords');
    const ieee = await texOf(String.raw`\documentclass[conference]{IEEEtran}\IEEEoverridecommandlockouts\title{P}
\author{\IEEEauthorblockN{Alice}\IEEEauthorblockA{MIT\\ alice@mit.edu}\and\IEEEauthorblockN{Bob}\IEEEauthorblockA{CMU}}
\begin{document}\maketitle\begin{IEEEkeywords}graphs, networks\end{IEEEkeywords}\IEEEPARstart{T}{his} paper.\IEEEpeerreviewmaketitle\end{document}`);
    assert.deepStrictEqual([ieee.ast.metadata.author, ieee.ast.metadata.keywords, ieee.paras.slice(1), ieee.warnings.length],
        ['Alice, Bob', 'graphs, networks', ['Alice, MIT alice@mit.edu, Bob, CMU', 'Index Terms: graphs, networks', 'This paper.'], 0], 'TEX parse: IEEEtran author blocks, IEEEkeywords, \\IEEEPARstart');
    const koma = await texOf(String.raw`\documentclass{scrartcl}\begin{document}\minisec{Small}Text.\epigraph{To be.}{\textit{Hamlet}}\dictum[Author]{Wise.}\end{document}`);
    const komaParas = collectAllNodes(koma.ast).filter(n => n.type === 'paragraph');
    assert.deepStrictEqual(komaParas.map(p => [p.text, (p.metadata as any).style, (p.metadata as any).alignment]),
        [['Small', undefined, undefined], ['Text.', undefined, undefined], ['To be.', 'Quote', undefined], ['Hamlet', 'Quote', 'right'], ['Wise.', 'Quote', undefined], ['(Author)', 'Quote', 'right']], 'TEX parse: \\minisec, \\epigraph, \\dictum');
    assert.deepStrictEqual(runsOf(komaParas[0]), [['Small', { bold: true }]], 'TEX parse: \\minisec is bold');

    // Legacy input encodings: a declared inputenc encoding, an undeclared 8-bit file, a TeXShop magic line.
    const latin1 = await texOf(Buffer.from(String.raw`\documentclass{article}\usepackage[latin1]{inputenc}\begin{document}Café Größe\end{document}`, 'latin1'));
    assert.deepStrictEqual(latin1.paras, ['Café Größe'], 'TEX parse: inputenc latin1');
    assert.deepStrictEqual((await texOf(Buffer.from(String.raw`\begin{document}Déjà vu\end{document}`, 'latin1'))).paras, ['Déjà vu'], 'TEX parse: an undeclared 8-bit file reads as Windows-1252');
    assert.deepStrictEqual((await texOf(Buffer.from(String.raw`\usepackage[latin1]{inputenc}\begin{document}Café\end{document}`, 'utf8'))).paras, ['Café'], 'TEX parse: a file converted to UTF-8 that still declares latin1 reads as UTF-8');
    const koi = Buffer.from([...String.raw`\usepackage[koi8-r]{inputenc}\begin{document}`].map(c => c.charCodeAt(0)).concat([0xf0, 0xd2, 0xc9, 0xd7, 0xc5, 0xd4], [...'\\end{document}'].map(c => c.charCodeAt(0))));
    assert.deepStrictEqual((await texOf(koi)).paras, ['Привет'], 'TEX parse: inputenc koi8-r');
    const magic = Buffer.concat([Buffer.from('% !TEX encoding = ISO-8859-15\n\\begin{document}'), Buffer.from([0xa4]), Buffer.from('\\end{document}')]);
    assert.deepStrictEqual((await texOf(magic)).paras, ['€'], 'TEX parse: % !TEX encoding line');
    const damaged = Buffer.concat([Buffer.from('\\begin{document}Größe über ', 'utf8'), Buffer.from([0xff]), Buffer.from(' end\\end{document}')]);
    assert.deepStrictEqual((await texOf(damaged)).paras, ['Größe über � end'], 'TEX parse: mostly-UTF-8 text with a damaged byte stays UTF-8');
    const includedZip = zipSync({ 'main.tex': strToU8('\\documentclass{article}\\usepackage[latin1]{inputenc}\\begin{document}\\input{ch}\\end{document}'), 'ch.tex': Buffer.from('Chapitre é', 'latin1') });
    assert.deepStrictEqual((await texOf(Buffer.from(includedZip))).paras, ['Chapitre é'], 'TEX parse: an included file takes the main file\'s encoding');

    // Plain TeX and ConTeXt are other formats: their text is read, with a warning for ConTeXt.
    const plain = await texOf(String.raw`\magnification=\magstep1 \hsize=6.5truein \parskip 6pt plus 1pt
\beginsection Introduction

Some text \vskip 12pt plus 2pt more {\it text}.
\bye
ignored`);
    assert.deepStrictEqual([plain.paras, plain.warnings.length], [['Introduction', 'Some text more text.'], 0], 'TEX parse: plain TeX \\beginsection, glue and register assignments, \\bye');
    const context = await texOf(String.raw`\setupbodyfont[11pt]
\starttext
\startsection[title={Introduction},reference=sec:intro]
Hello {\bf world}.
\startitemize[n]
\item One
\item Two
\stopitemize
\starttyping
code here
\stoptyping
\stopsection
\stoptext`);
    assert.deepStrictEqual(context.ast.content.map(n => [n.type, n.text, (n.metadata as any)?.listType]), [['heading', 'Introduction', undefined], ['paragraph', 'Hello world.', undefined], ['list', 'One', 'ordered'], ['list', 'Two', 'ordered'], ['code', 'code here', undefined]], 'TEX parse: ConTeXt sections, itemize and typing');
    assert.ok(context.warnings.some(w => JSON.stringify(w).includes('ConTeXt')), 'TEX parse: a ConTeXt document is reported as not LaTeX');

    // Review follow-ups: a package's own conditional keeps its \else and \fi; nested theorems keep their
    // heads and numbers; a control word ends with its macro body; verbatim arguments print as written.
    assert.deepStrictEqual((await texOf(String.raw`\ifPDFTeX \ifdraft X \else Y \fi Z\fi W`)).paras, ['X Y ZW'], 'TEX parse: an unknown \\ifdraft inside a taken branch reads both of its own branches');
    assert.deepStrictEqual((await texOf(String.raw`\ifXeTeX \ifdraft A\fi B \else C \fi D`)).paras, ['C D'], 'TEX parse: an unknown \\ifdraft inside a skipped branch is skipped with it');
    assert.deepStrictEqual((await texOf(String.raw`\ifXeTeX \ifstrequal{a}{a}{S}{T} X\else Y\fi`)).paras, ['Y'], 'TEX parse: etoolbox tests are not taken for conditionals');
    assert.deepStrictEqual((await texOf(String.raw`\newtheorem{thm}{T}\begin{document}\begin{thm}\label{o}\begin{thm}\label{i}x\end{thm}\end{thm}\ref{o},\ref{i}\end{document}`)).paras, ['T 1.', 'T 2. x', '1,2'], 'TEX parse: nested theorems keep their own heads and numbers');
    assert.deepStrictEqual(runsOf(collectAllNodes((await texOf(String.raw`\newcommand{\foo}{\textbf}\foo baz`)).ast).find(n => n.type === 'paragraph')), [['b', { bold: true }], ['az', {}]], 'TEX parse: a control word ends where its macro body ends');
    assert.deepStrictEqual((await texOf(String.raw`\NewDocumentCommand\code{v}{\texttt{#1}}\code|\foo{x} a_b|`)).paras, ['\\foo{x} a_b'], 'TEX parse: an xparse v argument prints as written');
    const ifxIdioms = await texOf(String.raw`\ifx\pdfoutput\undefined DVI\else PDF\fi{} \ifx\undefined\directlua NOLUA\else LUA\fi`);
    assert.deepStrictEqual([ifxIdioms.paras, ifxIdioms.warnings.length], [['PDF NOLUA'], 0], 'TEX parse: \\ifx\\pdfoutput\\undefined and the reversed \\ifx\\undefined\\cs');
    assert.deepStrictEqual((await texOf(Buffer.from('% \\usepackage[koi8-r]{inputenc}\n\\begin{document}caf\u00e9\\end{document}', 'latin1'))).paras, ['café'], 'TEX parse: a commented-out inputenc declares nothing');
    assert.deepStrictEqual((await texOf(Buffer.from('\\usepackage[utf8]{inputenc}\\begin{document}caf\u00e9\\end{document}', 'latin1'))).paras, ['café'], 'TEX parse: a declared utf8 the bytes contradict is judged by the bytes');
    assert.deepStrictEqual((await texOf(Buffer.concat([Buffer.from('% !TEX encoding = IsoLatin2\n\\begin{document}'), Buffer.from([0xb1]), Buffer.from('\\end{document}')]))).paras, ['ą'], 'TEX parse: TeXShop encoding names');

    // A URL argument's `%` is a character (hyperref reads it so), not a comment that takes the closing
    // brace and the rest of the document into the link.
    const urlPct = await texOf(String.raw`\documentclass{article}\begin{document}
See \url{https://example.com/a%20b} for details, \href{https://example.com/c%20d}{the site} and \nolinkurl{x%y}.

Second paragraph.
\end{document}`);
    assert.deepStrictEqual([urlPct.paras, collectAllNodes(urlPct.ast).filter(n => (n.metadata as any)?.link).map(n => (n.metadata as any).link), urlPct.warnings.length],
        [['See https://example.com/a%20b for details, the site and x%y.', 'Second paragraph.'], ['https://example.com/a%20b', 'https://example.com/c%20d'], 0], 'TEX parse: % in \\url, \\href and \\nolinkurl is part of the URL');
    // The generator escapes every % it writes in a URL (a bare one is a comment inside a heading or a
    // footnote), and those links parse back.
    const zurich = await OfficeParser.parseOffice(Buffer.from('# See [Zürich](https://de.wikipedia.org/wiki/Zürich) here\n\nA note[^1] and [home](https://example.edu/~alice/).\n\n[^1]: See [Zürich](https://de.wikipedia.org/wiki/Zürich).\n'), { fileType: 'md' } as any);
    const zurichTex = (await zurich.to('tex')).value as string;
    assert.ok(zurichTex.includes('\\href{https://de.wikipedia.org/wiki/Z\\%C3\\%BCrich}{Zürich} here}') && zurichTex.includes('\\href{https://example.edu/\\%7Ealice/}{home}')
        && !/[^\\]%[0-9A-F]{2}/.test(zurichTex.slice(zurichTex.indexOf('\\begin{document}'))), 'TEX: URL encodings are written \\%XX');
    const zurichBack = await OfficeParser.parseOffice(Buffer.from(zurichTex), { fileType: 'tex' } as any);
    assert.deepStrictEqual(collectAllNodes(zurichBack).filter(n => (n.metadata as any)?.link).map(n => [n.text, (n.metadata as any).link]),
        [['Zürich', 'https://de.wikipedia.org/wiki/Z%C3%BCrich'], ['Zürich', 'https://de.wikipedia.org/wiki/Z%C3%BCrich'], ['home', 'https://example.edu/%7Ealice/']], 'TEX round trip: links in a heading and a footnote parse back');
    // What is left open at the end of the input took in the rest: reported, not passed over.
    const openGroup = await texOf(String.raw`\begin{document}Start \textbf{bold never closed.

Second.\end{document}`);
    assert.ok(openGroup.warnings.some(w => w.code === 'LATEX_CONSTRUCT_NOT_INTERPRETED' && w.message.includes('{ without a matching } (read to the end of the input)')), 'TEX parse: an unclosed argument is reported');
    for (const [open, what] of [[String.raw`\begin{itemize}\item a`, '\\begin{itemize} without its \\end{itemize}'], [String.raw`\begin{tabular}{l} a`, '\\begin{tabular} without its \\end{tabular}'], ['$x', 'math opened with $ without its $'], [String.raw`\iffalse x`, '\\iffalse without its \\fi']]) {
        assert.ok((await texOf(open)).warnings.some(w => w.message.includes(what)), `TEX parse: ${what} is reported`);
    }
    // TeX drops a comment in a formula: kept, it was typeset (`\%`), and an unbalanced brace in one made
    // the generator write the whole formula as verbatim text.
    const mathComments = await texOf(String.raw`\begin{document}\begin{equation} x = 1 % a comment
\end{equation}\begin{align}
a &= b \\ % first { unbalanced
c &= d
\end{align}\ensuremath{y % z
}\end{document}`);
    const mathTexts = collectAllNodes(mathComments.ast).filter(n => (n.metadata as any)?.math).map(n => n.text);
    assert.deepStrictEqual(mathTexts, ['\\begin{equation} x = 1 \n\\end{equation}', '\\begin{align}\na &= b \\\\ \nc &= d\n\\end{align}', 'y'], 'TEX parse: comments in display-math environments and \\ensuremath are dropped');
    const mathCommentsTexWarnings: any[] = [];
    const mathCommentsTex = (await mathComments.ast.to('tex', { onWarning: (w: any) => mathCommentsTexWarnings.push(w) } as any)).value as string;
    assert.ok(mathCommentsTex.includes('\\begin{align}\na &= b \\\\ \nc &= d\n\\end{align}') && !mathCommentsTex.includes('comment') && !mathCommentsTexWarnings.length, 'TEX: the formulas are written as math, with no comment typeset');
    // \input reads a file as TeX does: its last line end is a space (never a paragraph break), and the
    // primitive form takes a name without braces.
    const inputZip = (main: string) => Buffer.from(zipSync({ 'main.tex': strToU8(`\\documentclass{article}\\begin{document}\n${main}\n\\end{document}`), 'acc.tex': strToU8('95.3\n'), 'chapter1.tex': strToU8('Chapter text.') }));
    assert.deepStrictEqual((await texOf(inputZip('The accuracy is \\input{acc} percent on the test set.'))).paras, ['The accuracy is 95.3 percent on the test set.'], 'TEX parse: \\input in a sentence keeps the paragraph');
    assert.deepStrictEqual((await texOf(inputZip('The accuracy is \\input{acc}\npercent.'))).paras, ['The accuracy is 95.3 percent.'], 'TEX parse: \\input at a line end keeps the paragraph');
    assert.deepStrictEqual((await texOf(inputZip('Before \\input chapter1 after.\n\n\\input{acc}\n\nAfter.'))).paras, ['Before Chapter text. after.', '95.3', 'After.'], 'TEX parse: \\input without braces; blank lines around an \\input still break paragraphs');
    // The main file of a project is not a part of another document.
    for (const [layout, expected] of [
        [{ 'paper-main/main.tex': '\\documentclass{article}\\begin{document}MAIN BODY \\input{chapter1}\\end{document}', 'paper-main/chapter1.tex': '\\documentclass[main]{subfiles}\\begin{document}Chapter one.\\end{document}' }, 'MAIN BODY Chapter one.'],
        [{ 'paper.tex': '\\documentclass{article}\\begin{document}MAIN BODY \\subfile{intro}\\end{document}', 'intro.tex': '\\documentclass[paper]{subfiles}\\begin{document}Intro.\\end{document}' }, 'MAIN BODY Intro.'],
        [{ 'a.tex': '\\documentclass{article}\\begin{document}Part A.\\end{document}', 'z.tex': '% \\documentclass{article}\n\\documentclass{book}\\begin{document}Whole \\include{a}\\end{document}' }, 'Whole Part A.'],
        [{ 'main.tex': '\\input{preamble}\\begin{document}Body.\\end{document}', 'preamble.tex': '\\documentclass{article}' }, 'Body.'],
    ] as const) {
        const project = await texOf(Buffer.from(zipSync(Object.fromEntries(Object.entries(layout).map(([k, v]) => [k, strToU8(v)])))));
        assert.deepStrictEqual(project.paras, [expected], `TEX project: the main file of ${Object.keys(layout).join(', ')}`);
    }
    // Scans of the raw source read comments as TeX does.
    assert.deepStrictEqual((await texOf('\\documentclass{article}\\begin{document}\n% \\chapter{Old draft}\n\\section{Intro}\\end{document}')).ast.content.map(n => (n.metadata as any).level), [1], 'TEX parse: a commented-out \\chapter does not demote the sections');
    const commentedClass = await texOf('%\\documentclass{beamer}\n\\documentclass{article}\\begin{document}\\begin{theorem}T\\end{theorem}\\end{document}');
    assert.deepStrictEqual([commentedClass.ast.metadata.nativeProperties?.documentClass, commentedClass.ast.content[0].type], ['article', 'paragraph'], 'TEX parse: a commented-out \\documentclass is not the class');
    const commentedTable = await texOf('\\begin{document}\\begin{tabular}{ll}\na & b \\\\\n% \\begin{tabular}{lll} was the old layout\nc & \\verb|d&e| \\\\\n\\end{tabular}\n\n\\section{After}Text.\\end{document}');
    assert.deepStrictEqual([commentedTable.ast.content.map(n => n.type), commentedTable.ast.content[0].children!.map(r => r.children!.map(c => c.text))], [['table', 'heading', 'paragraph'], [['a', 'b'], ['c', 'd&e']]], 'TEX parse: a commented-out \\begin inside a tabular, and \\verb in a cell');

    // A row's `\\` right before `\end{tabular}` (no space between) ends the row, not the environment's end:
    // read as one `\\end`, the table took the caption and the text after it into a row of its own.
    const rowEnd = await texOf(String.raw`\begin{document}\begin{table}\begin{tabular}{ll}a & b\\c & x\\\end{tabular}\caption{Cap}\end{table}After text.\end{document}`);
    assert.deepStrictEqual([rowEnd.ast.content.map(n => n.type), rowEnd.ast.content[0].children!.length, rowEnd.ast.content.slice(1).map(n => n.text)], [['table', 'paragraph', 'paragraph'], 2, ['Cap', 'After text.']], 'TEX parse: \\\\ directly before \\end{tabular}');

    // amsmath's \DeclareMathOperator and mathtools' paired delimiters are expanded, as \newcommand is.
    const declared = await texOf(String.raw`\documentclass{article}\usepackage{amsmath,mathtools,bm}
\DeclareMathOperator*{\argmax}{arg\,max}\DeclareMathOperator{\Tr}{Tr}
\DeclarePairedDelimiter\abs{\lvert}{\rvert}
\DeclarePairedDelimiterX\innerp[2]{\langle}{\rangle}{#1,\delimsize\vert #2}
\begin{document}$\argmax_\theta \Tr A \coloneqq \abs{y} + \abs*{\frac{a}{b}} + \abs[\big]{z} + \innerp*{u}{v}$ and $\bm{x}$\end{document}`);
    assert.deepStrictEqual([collectAllNodes(declared.ast).filter(n => (n.metadata as any)?.math).map(n => n.text), declared.warnings.length],
        [['\\operatorname*{arg\\,max}_\\theta \\operatorname{Tr} A \\coloneqq \\lvert y \\rvert + \\left\\lvert \\frac{a}{b} \\right\\rvert + \\bigl\\lvert z \\bigr\\rvert + \\left\\langle u,\\middle\\vert v \\right\\rangle', '\\bm{x}'], 0],
        'TEX parse: \\DeclareMathOperator, \\DeclarePairedDelimiter(X), sized and starred');
    // The generator loads the packages the formulas' commands come from, writes KaTeX's and MathJax's own
    // macros as the LaTeX they stand for, and gives a command nothing defines its name as its meaning.
    const declaredWarnings: any[] = [];
    const declaredTex = (await declared.ast.to('tex', { onWarning: (w: any) => declaredWarnings.push(w) } as any)).value as string;
    assert.ok(/\\usepackage\{amsmath,amssymb\}\n\\usepackage\{mathtools\}\n\\usepackage\{bm\}\n/.test(declaredTex) && !declaredWarnings.length, 'TEX: \\coloneqq loads mathtools and \\bm loads bm (last)');
    const dialectWarnings: any[] = [];
    const dialect = await OfficeParser.parseOffice(Buffer.from('Reals $x \\in \\R^n$, $\\lang a, b \\rang$, $\\Rightarrow \\Reals$ and $\\zork{x}$.'), { fileType: 'md' } as any);
    const dialectTex = (await dialect.to('tex', { onWarning: (w: any) => dialectWarnings.push(w) } as any)).value as string;
    assert.ok(dialectTex.includes('$x \\in \\mathbb{R}^n$, $\\langle a, b \\rangle$, $\\Rightarrow \\mathbb{R}$ and $\\zork{x}$')
        && dialectTex.includes('\\AtBeginDocument{%\n  \\providecommand{\\zork}{\\texttt{\\textbackslash zork}}%\n}'), 'TEX: KaTeX macros written as LaTeX; an unknown command prints its name');
    assert.deepStrictEqual(dialectWarnings.map(w => [w.code, w.details?.feature]), [['CONTENT_NOT_REPRESENTABLE', 'math commands no package the output loads defines (\\zork), each printed as its name']], 'TEX: the unknown math command is reported');
    const dialectAgain = (await (await OfficeParser.parseOffice(Buffer.from(dialectTex), { fileType: 'tex' } as any)).to('tex', { onWarning: () => {} } as any)).value as string;
    assert.strictEqual(dialectAgain, dialectTex, 'TEX round trip: the math definitions and rewritten macros are a fixed point');

    // \bibliography and \printbibliography print the cited entries of the .bib database (the project's
    // .bbl when it ships one, as compiling reads it); a key with no entry, or a database the parser
    // cannot read, prints the key, and the file is reported.
    const refsBib = '% refs\n@string{aw = "Addison-Wesley"}\n@book{knuth, author = {Knuth, Donald E.}, title = {The {\\TeX}book}, publisher = aw, year = 1984}\n'
        + '@article{lamport, author = {Leslie Lamport and Frank Mittelbach and Michel Goossens}, title = {On Things}, journal = {J. Things}, year = {1994}}\n@comment{@book{fake, title={no}}}\n@misc(extra, title = "Extra", howpublished = "Online", year = 2001)\n';
    const bibMain = (tail: string) => `\\documentclass{article}\\begin{document}See \\cite{knuth,lamport} and \\cite{missing}.\\nocite{extra}\n${tail}\\end{document}`;
    // (A heading's generated id is not compared: the generator labels every heading.)
    const bibParas = (a: any) => collectAllNodes(a).filter(n => n.type === 'heading' || n.type === 'list').map(n => [n.type, n.text, n.type === 'list' ? (n.metadata as any).anchorIds?.[0] : undefined]);
    const bibProject = await texOf(Buffer.from(zipSync({ 'main.tex': strToU8(bibMain('\\bibliographystyle{plain}\\bibliography{refs}')), 'refs.bib': strToU8(refsBib) })));
    const bibExpected = [['heading', 'References', undefined], ['list', 'Donald E. Knuth. The TeXbook. Addison-Wesley, 1984.', 'knuth'],
        ['list', 'Leslie Lamport, Frank Mittelbach and Michel Goossens. On Things. J. Things, 1994.', 'lamport'], ['list', 'missing', 'missing'], ['list', 'Extra. Online, 2001.', 'extra']];
    assert.deepStrictEqual([bibParas(bibProject.ast), bibProject.warnings.length], [bibExpected, 0], 'TEX parse: \\bibliography prints the cited entries of the .bib');
    const bibLatex = await texOf(Buffer.from(zipSync({ 'main.tex': strToU8(bibMain('\\printbibliography[title={Works Cited}]').replace('\\begin{document}', '\\addbibresource{refs.bib}\\begin{document}')), 'refs.bib': strToU8(refsBib) })));
    assert.deepStrictEqual(bibParas(bibLatex.ast), [['heading', 'Works Cited', undefined], ...bibExpected.slice(1)], 'TEX parse: biblatex \\printbibliography with a title');
    const bibBbl = await texOf(Buffer.from(zipSync({ 'main.tex': strToU8(bibMain('\\bibliography{refs}')), 'main.bbl': strToU8('\\begin{thebibliography}{1}\n\\bibitem{knuth} From the bbl.\n\\end{thebibliography}\n') })));
    assert.deepStrictEqual(bibParas(bibBbl.ast), [['heading', 'References', undefined], ['list', 'From the bbl.', 'knuth']], 'TEX parse: a project\'s .bbl is its bibliography');
    const bibLone = await texOf(bibMain('\\bibliography{refs}'));
    assert.deepStrictEqual([bibParas(bibLone.ast).map(p => p[1]), bibLone.warnings.map(w => w.code)], [['References', 'knuth', 'lamport', 'missing', 'extra'], ['LATEX_FILE_NOT_FOUND']], 'TEX parse: without its .bib, a bibliography lists the keys and the file is reported');
    // The generator writes a bibliography as thebibliography, so the citations resolve; a citation with no
    // entry is reported.
    const bibWarnings: any[] = [];
    const bibTex = (await bibProject.ast.to('tex', { onWarning: (w: any) => bibWarnings.push(w) } as any)).value as string;
    assert.ok(/\\begin\{thebibliography\}\{9\}\n\\bibitem\{knuth\} Donald E\. Knuth\. \\textit\{The TeXbook\}\. Addison-Wesley, 1984\.\n\\bibitem\{lamport\}/.test(bibTex) && !bibTex.includes('\\section{References}') && !bibWarnings.length, 'TEX: a bibliography is written as thebibliography under its own heading');
    assert.deepStrictEqual(bibParas(await OfficeParser.parseOffice(Buffer.from(bibTex), { fileType: 'tex' } as any)), bibExpected, 'TEX round trip: the bibliography parses back the same');
    const worksCited = (await bibLatex.ast.to('tex')).value as string;
    assert.ok(worksCited.includes('\\renewcommand{\\refname}{Works Cited}\n\\begin{thebibliography}'), 'TEX: a bibliography under another heading renames \\refname');
    // beamer's and llncs' \institute: a line of the title block.
    const institute = await texOf(String.raw`\documentclass{beamer}\title{T}\author{A \and B}\institute{Uni \and Lab}\begin{document}\begin{frame}\titlepage\end{frame}\end{document}`);
    assert.deepStrictEqual([institute.ast.content[0].children!.map(c => [(c.metadata as any).style, c.text]), (institute.ast.metadata.nativeProperties as any).institute],
        [[['Title', 'T'], ['Author', 'A, B'], ['Institute', 'Uni, Lab']], 'Uni, Lab'], 'TEX parse: \\institute in the title block');
    assert.ok(/\\author\{A, B\}\n\\institute\{Uni, Lab\}\n/.test((await institute.ast.to('tex')).value as string), 'TEX beamer: the institute line regenerates as \\institute');

    // \autoref, \cref and \Cref name what they refer to; \MakeUppercase and \lowercase set their text's
    // case; a symbol command's character is never part of an input ligature; the math environment is inline.
    const namedRefs = await texOf(String.raw`\documentclass{article}\usepackage{hyperref,cleveref}\newtheorem{lemma}{Lemma}\begin{document}
\section{Intro}\label{sec:i}\begin{figure}\includegraphics{a.png}\caption{Fig}\label{fig:a}\end{figure}\begin{equation}x\label{eq:x}\end{equation}\begin{lemma}\label{lem:l}L.\end{lemma}
See \autoref{sec:i}, \autoref{fig:a}, \cref{fig:a,eq:x}, \Cref{sec:i}, \cref{lem:l} and \ref{fig:a}.
\MakeUppercase{shout} \lowercase{QUIET}. \textasciigrave{}x\textasciigrave{} \textquotesingle{}q\textquotesingle{} and \begin{math}a+b\end{math} inline.\end{document}`, { onWarning: () => {} });
    const namedParas = namedRefs.paras.slice(-1);
    assert.deepStrictEqual(namedParas, ['See section\u00A01, Figure\u00A01, fig.\u00A01 and eq.\u00A0(1), Section\u00A01, lemma\u00A01 and 1. SHOUT quiet. `x` \'q\' and a+b inline.'], `TEX parse: named references, text case, symbols, inline math environment (${JSON.stringify(namedParas)})`);
    // A caption regenerates as \captionof (capt-of), so it parses back as a caption; a Word caption,
    // numbered in its text already, stays a paragraph.
    const captioned = await texOf(String.raw`\begin{document}\begin{table}\caption{Results}\begin{tabular}{l}a\\\end{tabular}\end{table}\begin{figure}\includegraphics{x.png}\caption{A picture}\end{figure}\end{document}`, { onWarning: () => {} });
    const captionTex = (await captioned.ast.to('tex', { onWarning: () => {} } as any)).value as string;
    assert.ok(captionTex.includes('\\usepackage{capt-of}') && captionTex.includes('\\captionof{table}{Results}') && captionTex.includes('\\captionof{figure}{A picture}'), 'TEX: captions as \\captionof, typed by the table or picture beside them');
    const captionsBack = collectAllNodes(await OfficeParser.parseOffice(Buffer.from(captionTex), { fileType: 'tex', onWarning: () => {} } as any)).filter(n => (n.metadata as any)?.style === 'Caption').map(n => n.text);
    assert.deepStrictEqual(captionsBack, ['Results', 'A picture'], 'TEX round trip: captions stay captions');
    const wordCaption: any = { type: 'docx', metadata: {}, attachments: [], content: [{ type: 'paragraph', metadata: { style: 'Caption' }, children: [{ type: 'text', text: 'Figure 1: Sales' }] }] };
    assert.ok(!((await OfficeGenerator.generate(wordCaption, 'tex', {})).value as string).includes('captionof'), 'TEX: a caption numbered in its text stays a paragraph');

    // \today prints the date of the parse (as LaTeX prints the date of the compile), in the document's
    // language, or the text texParserConfig.today sets; it is never dropped.
    const dateIn = (tag: string) => new Intl.DateTimeFormat(tag, { year: 'numeric', month: 'long', day: 'numeric' }).format(new Date());
    const todayDoc = await texOf(String.raw`\documentclass{article}\title{T}\date{\today}\begin{document}\maketitle Written on \today.\end{document}`);
    assert.deepStrictEqual([todayDoc.paras.slice(1), todayDoc.ast.metadata.nativeProperties?.date], [[dateIn('en-US'), `Written on ${dateIn('en-US')}.`], dateIn('en-US')], 'TEX parse: \\today is the date of the parse, in the title block and the text');
    assert.deepStrictEqual((await texOf(String.raw`\usepackage[ngerman]{babel}\begin{document}Berlin, \today\end{document}`)).paras, [`Berlin, ${dateIn('de')}`], 'TEX parse: \\today in the document\'s language');
    const fixedToday = await texOf(String.raw`\begin{document}Written on \today, \DTMtoday.\end{document}`, { texParserConfig: { today: '[DATE]' } });
    assert.deepStrictEqual(fixedToday.paras, ['Written on [DATE], [DATE].'], 'TEX parse: texParserConfig.today replaces \\today (and datetime2\'s \\DTMtoday)');
    const badToday = await texOf(String.raw`\begin{document}\today\end{document}`, { texParserConfig: { today: 42 } });
    assert.ok(badToday.paras[0] === dateIn('en-US') && badToday.warnings.some(w => w.code === 'INVALID_CONFIG_VALUE' && w.message.includes('texParserConfig.today')), 'TEX parse: a texParserConfig.today that is not a string is reported and ignored');

    // A figure holding only a drawing (omitted) keeps its caption, which carries the figure's number for \ref.
    for (const ignoreInternalLinks of [false, true]) {
        const drawing = await texOf(String.raw`\begin{document}\begin{figure}\begin{tikzpicture}\draw (0,0)--(1,1);\end{tikzpicture}\caption{A drawing}\label{fig:d}\end{figure}
\begin{figure}\label{fig:e}\begin{tikzpicture}\end{tikzpicture}\caption{Label first}\end{figure}
\begin{figure}\includegraphics{x.png}\caption{Photo}\label{fig:p}\end{figure}See \ref{fig:d}, \ref{fig:e}, \ref{fig:p}.\end{document}`, { ignoreInternalLinks });
        assert.strictEqual(drawing.paras[drawing.paras.length - 1], 'See 1, 2, 3.', `TEX parse: a drawing-only figure's label reads its number (ignoreInternalLinks: ${ignoreInternalLinks})`);
    }

    // Names from Object.prototype are not table entries.
    const proto = await texOf(String.raw`\begin{document}\constructor \toString \color{constructor}x \texthasOwnProperty{z}\end{document}`);
    assert.ok(!JSON.stringify(proto.ast.content).includes('function') && !JSON.stringify(proto.ast.content).includes('native code'), 'TEX parse: \\constructor and \\toString are unknown commands, not table lookups');

    // ── Tier 4: round trip with the generator reaches a fixed point ───────────
    const docx = await OfficeParser.parseOffice(path.join(__dirname, 'files/test.docx'), { extractAttachments: true });
    const pin = { metadataOverrides: { modified: new Date('2024-01-01T00:00:00Z') } } as any;
    const t1 = (await docx.to('tex', pin)).value as string;
    const t2 = (await (await OfficeParser.parseOffice(Buffer.from(t1), { fileType: 'tex' })).to('tex', pin)).value as string;
    const t3 = (await (await OfficeParser.parseOffice(Buffer.from(t2), { fileType: 'tex' })).to('tex', pin)).value as string;
    assert.strictEqual(t3, t2, 'TEX round trip: generate -> parse -> generate is stable after one cycle');
    assert.ok(/\\title\{\\protect\\textcolor\{hex17365D\}\{Demonstration of DOCX support in calibre\}\}\n\\author\{\}\n\\date\{\}/.test(t1) && /\\label\{demonstration-of-docx-support-in-calibre\}\\maketitle\n/.test(t1),
        'TEX round trip: a Word Title paragraph is typeset with \\maketitle, printing only what the document shows');
    const fixtureTex = (await ast.to('tex')).value as string;
    const f2 = (await (await OfficeParser.parseOffice(Buffer.from(fixtureTex), { fileType: 'tex' })).to('tex')).value as string;
    const f3 = (await (await OfficeParser.parseOffice(Buffer.from(f2), { fileType: 'tex' })).to('tex')).value as string;
    assert.strictEqual(f3, f2, 'TEX round trip: the hand-written fixture is stable after one cycle too');
    const back = await OfficeParser.parseOffice(Buffer.from(t1), { fileType: 'tex' });
    const count = (a: any, t: string) => collectAllNodes(a).filter(n => n.type === t).length;
    for (const t of ['heading', 'table', 'list', 'image']) assert.strictEqual(count(back, t), count(docx, t), `TEX round trip: ${t} count preserved`);
    assert.deepStrictEqual([back.metadata.title, back.metadata.author], [docx.metadata.title, docx.metadata.author], 'TEX round trip: pdftitle/pdfauthor keep the metadata, whatever the title block prints');

    // ── Carried images and PDF pictures ────────────────────────────────────────
    // Two different pictures whose carried PDFs share the short content hash are both carried, apart.
    const COLLIDING_A = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgMHMDAAC2AH3s+igHAAAAAElFTkSuQmCC';
    const COLLIDING_B = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGNgWNcGAAHmATW5wUTxAAAAAElFTkSuQmCC';
    const pair: any = {
        type: 'docx', metadata: {},
        attachments: [{ name: 'dot.png', mimeType: 'image/png', data: COLLIDING_A }, { name: 'dot.png', mimeType: 'image/png', data: COLLIDING_B }].map((a, i) => ({ ...a, name: `${i}/dot.png`, type: 'image' })),
        content: [0, 1].map(i => ({ type: 'paragraph', children: [{ type: 'image', metadata: { attachmentName: `${i}/dot.png` } }] })),
    };
    const pairTex = (await OfficeGenerator.generate(pair, 'tex', { onWarning: () => {} } as any)).value as string;
    const pairBlocks = [...pairTex.matchAll(/\\begin\{filecontents\*\}\{([^}]+)\}/g)].map(m => m[1]);
    const pairIncludes = [...pairTex.matchAll(/\\includegraphics\[[^\]]*\]\{([^}]+)\}/g)].map(m => m[1]);
    assert.ok(pairBlocks.length === 2 && pairBlocks[0] !== pairBlocks[1] && pairIncludes.join() === pairBlocks.join(), `TEX: two pictures sharing a hash are carried apart (${pairBlocks}; ${pairIncludes})`);
    // Decoding draws from the document's budget; an image PDF can read as it is does not.
    const rgbaPng = fs.readFileSync(path.join(__dirname, '..', 'docs', 'favicon.png'));
    assert.ok(imageToTextPdf(rgbaPng, 0.75, newDecodeBudget()) && !imageToTextPdf(rgbaPng, 0.75, { pixels: 10 }), 'TEX: a decoded image draws from the budget and is refused past it');
    assert.ok(imageToTextPdf(Buffer.from(COLLIDING_A, 'base64'), 0.75, { pixels: 0 }), 'TEX: a PNG passed through as it is costs no budget');
    // A PDF picture takes a free name beside an attachment already called what it would be.
    const pictureA = imageToTextPdf(Buffer.from(COLLIDING_A, 'base64'), 0.75, newDecodeBudget())!.pdf;
    const named = await OfficeParser.parseOffice(Buffer.from(zipSync({
        'main.tex': strToU8('\\documentclass{article}\\begin{document}\\includegraphics{fig.png}\\includegraphics{fig.pdf}\\end{document}'),
        'fig.png': new Uint8Array(rgbaPng), 'fig.pdf': strToU8(pictureA),
    })), { extractAttachments: true } as any);
    assert.deepStrictEqual(named.attachments.map(a => [a.name, a.mimeType]), [['fig.png', 'image/png'], ['fig-2.png', 'image/png']], 'TEX project: a PDF picture named like an existing attachment gets a free name');
    // A picture drawn turned (a rotated placement, or a rotated page) stays the PDF it is.
    const upright = imageToTextPdf(Buffer.from(COLLIDING_A, 'base64'), 0.75, newDecodeBudget())!.pdf;
    const turned = upright.replace(/q ([\d.]+) 0 0 ([\d.]+) 0 0 cm/, (_m, w, h) => `q 0 ${w} -${h} 0 ${h} 0 cm`);
    const rotatedPage = upright.replace('/Type /Page ', '/Type /Page /Rotate 90 ');
    assert.ok(turned !== upright && rotatedPage !== upright, 'TEX: the rotated fixtures differ from the upright one');
    assert.deepStrictEqual([upright, turned, rotatedPage].map(pdf => imageFromPdf(new TextEncoder().encode(pdf), newDecodeBudget())?.mimeType ?? 'kept as PDF'),
        ['image/png', 'kept as PDF', 'kept as PDF'], 'TEX: only an upright, unrotated single-image page is taken as its picture');

    // The sectioning a document uses is that of the files it includes: a thesis whose chapters are in
    // files of their own read its sections as top-level headings, beside the chapters.
    for (const cls of ['report', 'book']) {
        const thesis = zipSync({ 'main.tex': strToU8(`\\documentclass{${cls}}\\begin{document}\n\\include{ch1}\n\\end{document}\n`), 'ch1.tex': strToU8('\\chapter{Intro}\nText.\n\\section{Background}\nMore.\n') });
        const levels = collectAllNodes(await OfficeParser.parseOffice(Buffer.from(thesis), { fileType: 'zip' } as any)).filter(n => n.type === 'heading').map(n => (n.metadata as any)?.level);
        assert.deepStrictEqual(levels, [1, 2], `TEX: a ${cls}'s chapters in an included file are the top level`);
    }
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
    // Any CSS colour the AST holds (an HTML or Markdown document's own): a named one was dropped.
    assert.deepStrictEqual(['red', 'Navy', 'rgb(255, 128, 0)', 'rgba(0 0 255 / 50%)', 'hsl(120, 100%, 25%)', '#0f08', '#11223380', 'transparent', 'currentColor', 'inherit', 'rgb(1,2)', 'x'.repeat(100)].map(c => hexColor(c)),
        ['FF0000', '000080', 'FF8000', '0000FF', '008000', '00FF00', '112233', null, null, null, null, null], 'hexColor: named, rgb(), hsl() and alpha forms resolve; others are null');
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

/**
 * Configuration consistency: every option means the same on every surface, a mistake in a config is
 * reported rather than silently ignored, and a value an option does not accept falls back visibly.
 */
async function testConfigConsistency(): Promise<void> {
    const md = await OfficeParser.parseOffice(Buffer.from('# Title\n\nText.'), { fileType: 'md' } as any);
    const run = async (destination: string, config: any) => {
        const seen: any[] = [];
        const result = await md.to(destination as any, { ...config, onWarning: (w: any) => seen.push(w) });
        return { result, seen, codes: seen.map(w => w.code), messages: result.messages.map((m: any) => m.message) };
    };

    // A value an option does not accept: reported with what it accepts, and the default used.
    const tex = await run('tex', { texConfig: { documentClass: 'reprot', format: 'A44', margin: { top: '1 inch', left: '2cm' } } });
    assert.deepStrictEqual(tex.seen.filter(w => w.code === 'INVALID_CONFIG_VALUE').map(w => w.message), [
        'Invalid texConfig.documentClass: "reprot". Expected one of auto, article, report, book, beamer; the default ("auto") is used instead.',
        'Invalid texConfig.format: "A44". Expected one of letter, legal, tabloid, ledger, a0, a1, a2, a3, a4, a5, a6; the default ("A4") is used instead.',
        'Invalid texConfig.margin.top: "1 inch". Expected a number of points or a length such as \'1in\', \'2cm\', \'36pt\'; the default (72) is used instead.',
    ], 'Config: invalid texConfig values are reported with what the option accepts');
    const texOut = String(tex.result.value);
    assert.ok(texOut.includes('\\documentclass[\\officeparserdriver 11pt]{article}') && texOut.includes('top=72pt') && texOut.includes('left=56.69pt'), 'Config: the defaults replace invalid values; valid ones are kept');
    assert.ok(tex.messages.some((m: string) => m.startsWith('Invalid texConfig.documentClass')), 'Config: configuration warnings are among the result\'s messages');
    for (const [destination, config, option] of [
        ['md', { mdConfig: { dialect: 'githb' } }, 'mdConfig.dialect'],
        ['md', { mdConfig: { dialect: { extends: 'gitlb' } } }, 'mdConfig.dialect.extends'],
        ['html', { htmlConfig: { standalone: { styles: 'minimal' } } }, 'htmlConfig.standalone.styles'],
        ['chunks', { chunksConfig: { strategy: 'by-heading' } }, 'chunksConfig.strategy'],
        ['chunks', { chunksConfig: { splitBy: 'section' } }, 'chunksConfig.splitBy'],
        ['docx', { docxConfig: { format: 'B5', margin: { bottom: 'wide' } } }, 'docxConfig.format'],
        ['odt', { odtConfig: { format: 'A4 ' } }, 'odtConfig.format'],
    ] as const) {
        const r = await run(destination, config);
        assert.ok(r.seen.some(w => w.code === 'INVALID_CONFIG_VALUE' && w.message.startsWith(`Invalid ${option}:`)), `Config: ${option} is validated`);
    }
    const valid = await run('tex', { texConfig: { documentClass: 'report', format: 'letter', margin: { top: '1in', bottom: 36, left: '2.5cm', right: '' } }, mdConfig: { dialect: { extends: 'github' } } });
    assert.deepStrictEqual(valid.codes, [], 'Config: valid values (any case of a paper format, unit strings, numbers) raise nothing');
    const callerConfig = { texConfig: { margin: { top: 'huge' } }, mdConfig: { dialect: { extends: 'nope', math: 'none' } } };
    await run('tex', callerConfig);
    assert.deepStrictEqual(callerConfig, { texConfig: { margin: { top: 'huge' } }, mdConfig: { dialect: { extends: 'nope', math: 'none' } } }, 'Config: validation never edits the caller\'s own config');

    // A key a generator does not recognize is reported, as on the parser side.
    const unknown = await run('tex', { texConfig: { bundel: true }, extractAttachments: true });
    const unrecognized = unknown.seen.find(w => w.code === 'UNRECOGNIZED_CONFIG_OPTION');
    assert.ok(unrecognized && unrecognized.message.includes("'texConfig.bundel'") && unrecognized.message.includes("'extractAttachments'"), `Config: unrecognized generator keys are reported (${unrecognized?.message})`);
    assert.deepStrictEqual((await run('html', { htmlConfig: { containerWidth: 900 }, chunksConfig: { similarityThreshold: 0.5, embeddingFunction: async () => [] }, metadataOverrides: { title: 'x', custom: { a: 1 } } })).codes, [], 'Config: every documented key is recognized (all chunking strategies, metadata overrides)');

    // convert(): a parser or generator option at the top level is not read; the warning says where it belongs.
    const docx = path.join(__dirname, 'files/test.docx');
    const convertWarnings: any[] = [];
    const converted = await OfficeConverter.convert(docx, 'tex', { texConfig: { bundle: true }, ignoreNotes: true, ignoreInternalLinks: true, onWarning: (w: any) => convertWarnings.push(w) } as any);
    const misplaced = convertWarnings.find(w => w.code === 'UNRECOGNIZED_CONFIG_OPTION');
    assert.ok(typeof converted.value === 'string' && misplaced
        && misplaced.message.includes("'texConfig' (use 'generatorConfig.texConfig' instead)")
        && misplaced.message.includes("'ignoreNotes' (use 'parseConfig.ignoreNotes' instead)")
        && misplaced.message.includes("'ignoreInternalLinks' (it goes under parseConfig or generatorConfig)"), `Config: convert() names where a misplaced option belongs (${misplaced?.message})`);
    assert.ok(converted.messages.some(m => m.code === 'UNRECOGNIZED_CONFIG_OPTION'), 'Config: the misplaced-option warning is among convert()\'s messages');
    const bundled = await OfficeConverter.convert(docx, 'tex', { generatorConfig: { texConfig: { bundle: true } } });
    assert.ok(bundled.value instanceof Uint8Array, 'Config: generatorConfig.texConfig.bundle gives the zip, as the README shows');
    const counted = await OfficeConverter.convert(docx, 'tex');
    const keys = counted.messages.map(m => `${m.code}:${m.message}`);
    assert.strictEqual(new Set(keys).size, keys.length, `Config: convert() lists each warning once (${keys.length} messages)`);

    // A decompression limit reached while identifying an archive is that limit's error.
    const projectZip = path.join(__dirname, 'files/latex-project.zip');
    for (const [label, input, limits, code] of [
        ['a .zip by name', projectZip, { maxZipEntries: 1 }, 'ZIP entry count exceeds limit (1)'],
        ['an unnamed buffer', fs.readFileSync(projectZip), { maxUncompressedBytes: 100 }, 'ZIP uncompressed size limit exceeded (100 bytes)'],
    ] as const) {
        let message = '';
        await OfficeParser.parseOffice(input as any, { decompressionLimits: limits, onWarning: () => {} } as any).catch((e: any) => { message = e.message; });
        assert.ok(message.includes(code), `Config: a limit reached identifying ${label} is reported as the limit (${message})`);
    }

    // fileType accepts the names the matching extensions do.
    const texSource = Buffer.from('\\documentclass{article}\\begin{document}Hi\\end{document}');
    for (const alias of ['tex', 'latex', 'ltx']) assert.strictEqual((await OfficeParser.parseOffice(texSource, { fileType: alias } as any)).type, 'tex', `Config: fileType '${alias}'`);
    assert.strictEqual((await OfficeParser.parseOffice(fs.readFileSync(projectZip), { fileType: 'zip' } as any)).type, 'tex', "Config: fileType 'zip' parses what the archive holds");
    assert.strictEqual((await OfficeParser.parseOffice(fs.readFileSync(path.join(__dirname, 'files/test.odt')), { fileType: 'ott' } as any)).type, 'odt', "Config: fileType 'ott' is ODT");

    // The document language reaches every output with a place for it.
    const german = await OfficeParser.parseOffice(Buffer.from('\\documentclass{article}\\usepackage[ngerman]{babel}\\begin{document}Hallo\\end{document}'), { fileType: 'tex' } as any);
    assert.strictEqual(german.metadata.language, 'de', 'Language: from babel');
    assert.ok(String((await german.to('html')).value).includes('<html lang="de">'), 'Language: HTML lang');
    const epub = unzipSync((await german.to('epub')).value as Uint8Array);
    const opf = Object.entries(epub).find(([n]) => n.endsWith('.opf'))![1];
    assert.ok(strFromU8(opf).includes('<dc:language>de</dc:language>'), 'Language: EPUB dc:language');
    assert.strictEqual((await OfficeParser.parseOffice(Buffer.from((await german.to('epub')).value as Uint8Array), { fileType: 'epub' } as any)).metadata.language, 'de', 'Language: the EPUB parser puts it in metadata.language');
    const overridden = String((await german.to('html', { metadataOverrides: { language: 'fr-CA' } } as any)).value);
    assert.ok(overridden.includes('<html lang="fr-CA">'), 'Language: metadataOverrides.language wins');
    assert.ok(String((await md.to('html')).value).includes('<html lang="en">'), 'Language: English when the source states none');
    console.log('  Config consistency: All assertions passed ✓');
}

/**
 * A controller whose signal aborts the moment OCR starts listening to it: the abort lands while the
 * image's OCR job is queued or its worker is starting, whatever the timing (a timer could lose the
 * race to a fast parse). Aborts that land after a worker was handed the job use `ocrTestHooks`.
 */
function abortWhenOcrListens(): { signal: AbortSignal; listened: () => boolean } {
    const controller = new AbortController();
    const signal = controller.signal;
    const listen = signal.addEventListener.bind(signal);
    let listened = false;
    signal.addEventListener = ((type: string, listener: any, options?: any) => {
        listen(type, listener, options);
        if (type === 'abort' && !listened) {
            listened = true;
            queueMicrotask(() => controller.abort());
        }
    }) as typeof signal.addEventListener;
    return { signal, listened: () => listened };
}

/**
 * Cancellation: once a parse's signal has fired it rejects with AbortError and never resolves, in
 * every format that runs OCR, even when the abort lands while an image is being recognized (which
 * was taken for a failed recognition, so the parse resolved without the text). An OCR timeout is
 * still a failed recognition: an OCR_FAILED warning, and the parse goes on.
 */
async function testCancellation(): Promise<void> {
    const OCR_FORMATS = ['docx', 'pptx', 'xlsx', 'odt', 'odp', 'rtf', 'pdf', 'tex'];
    const file = (ext: string) => path.join(__dirname, `files/test.${ext}`);
    try {
        for (const ext of OCR_FORMATS) {
            const probe = abortWhenOcrListens();
            let outcome = 'resolved';
            await OfficeParser.parseOffice(file(ext), { abortSignal: probe.signal, ocr: true, extractAttachments: true, onWarning: () => {} } as any)
                .catch((e: any) => { outcome = e.name; });
            assert.ok(probe.listened(), `Cancellation ${ext}: OCR started, so the abort landed during OCR`);
            assert.strictEqual(outcome, 'AbortError', `Cancellation ${ext}: an abort during OCR rejects with AbortError`);
        }
        // ocrConfig.abortSignal is the same cancellation, delivered to OCR.
        const ocrOnly = abortWhenOcrListens();
        let ocrOnlyOutcome = 'resolved';
        await OfficeParser.parseOffice(file('docx'), { ocr: true, extractAttachments: true, ocrConfig: { abortSignal: ocrOnly.signal }, onWarning: () => {} } as any)
            .catch((e: any) => { ocrOnlyOutcome = e.name; });
        assert.strictEqual(ocrOnlyOutcome, 'AbortError', 'Cancellation: an abort through ocrConfig.abortSignal during OCR rejects with AbortError');
        // A timer abort, wherever it lands, never leaves a parse resolved after it fired.
        for (const ext of OCR_FORMATS) {
            const controller = new AbortController();
            setTimeout(() => controller.abort(), 5);
            let resolvedAfterAbort = false;
            await OfficeParser.parseOffice(file(ext), { abortSignal: controller.signal, ocr: true, extractAttachments: true, onWarning: () => {} } as any)
                .then(() => { resolvedAfterAbort = controller.signal.aborted; }, () => {});
            assert.ok(!resolvedAfterAbort, `Cancellation ${ext}: a parse never resolves after its signal fired`);
        }
        // An OCR timeout is a failed recognition, not a cancellation.
        const codes: string[] = [];
        const timedOut = await OfficeParser.parseOffice(file('docx'), { ocr: true, extractAttachments: true, ocrConfig: { timeout: { recognition: 1 } }, onWarning: (w: any) => codes.push(w.code) } as any);
        assert.ok(codes.includes('OCR_FAILED') && timedOut.attachments.length > 0 && !timedOut.attachments.some(a => a.ocrText),
            `Cancellation: an OCR timeout is OCR_FAILED and the parse resolves (${codes.join(', ')})`);

        // The worker has been handed the job, but Tesseract has not yet sent it: terminating the
        // worker now would leave Tesseract's send() rejecting with nothing to catch it, which ends a
        // Node process. An abort at exactly that moment must reject the parse and nothing else.
        const unhandled: unknown[] = [];
        const onUnhandled = (reason: unknown) => { unhandled.push(reason); };
        process.on('unhandledRejection', onUnhandled);
        const timersBefore = process.getActiveResourcesInfo().filter(r => r === 'Timeout').length;
        try {
            for (const [label, abortWhen] of [
                ['before Tesseract sends the job', (abort: () => void) => abort()],
                ['while the job is recognized', (abort: () => void) => { setTimeout(abort, 0); }],
            ] as const) {
                const controller = new AbortController();
                let hooked = false;
                ocrTestHooks.afterRecognizeCall = () => { if (!hooked) { hooked = true; abortWhen(() => controller.abort()); } };
                let outcome = 'resolved';
                await OfficeParser.parseOffice(file('docx'), { abortSignal: controller.signal, ocr: true, extractAttachments: true, onWarning: () => {} } as any)
                    .catch((e: any) => { outcome = e.name; });
                ocrTestHooks.afterRecognizeCall = undefined;
                assert.ok(hooked && outcome === 'AbortError', `Cancellation: an abort ${label} rejects with AbortError (${hooked}, ${outcome})`);
            }
            // terminateOcr() while an image is being recognized fails that recognition (OCR_FAILED, the
            // parse goes on) without ending the process.
            const terminateCodes: string[] = [];
            let terminated = false;
            ocrTestHooks.afterRecognizeCall = () => { if (!terminated) { terminated = true; void terminateOcr(); } };
            const survived = await OfficeParser.parseOffice(file('docx'), { ocr: true, extractAttachments: true, onWarning: (w: any) => terminateCodes.push(w.code) } as any);
            ocrTestHooks.afterRecognizeCall = undefined;
            assert.ok(terminated && survived.content.length > 0 && terminateCodes.includes('OCR_FAILED'), `Cancellation: terminateOcr() during recognition fails that image's OCR, not the parse (${terminateCodes.join(', ')})`);
            await new Promise(resolve => setTimeout(resolve, 50));
            assert.deepStrictEqual(unhandled, [], 'Cancellation: no rejection escapes when a worker is stopped mid-job');
        } finally {
            ocrTestHooks.afterRecognizeCall = undefined;
            process.off('unhandledRejection', onUnhandled);
        }
        // A cancelled recognition leaves no timer running: its timeout is cleared, so the process can exit.
        await terminateOcr();
        await new Promise(resolve => setTimeout(resolve, 50));
        const timersAfter = process.getActiveResourcesInfo().filter(r => r === 'Timeout').length;
        assert.ok(timersAfter <= timersBefore, `Cancellation: no timer outlives a cancelled recognition (${timersBefore} before, ${timersAfter} after)`);

        // A worker re-initializing for another language (the pool full of idle workers of the first)
        // is let finish before it is terminated; one that never finishes, as a language download that
        // hangs, holds terminateOcr() only until the load timeout, never for ever.
        ocrTestHooks.createWorker = async () => ({
            recognize: async () => ({ data: { text: 'x' } }),
            reinitialize: () => new Promise(() => { }),
            terminate: async () => { },
        });
        try {
            await Promise.all(Array.from({ length: 4 }, () => performOcr(Buffer.from('x'), { language: 'eng' })));
            const hung = performOcr(Buffer.from('x'), { language: 'fra', timeout: { workerLoad: 200 } }).catch(() => 'rejected');
            await new Promise(resolve => setTimeout(resolve, 20));
            const ended = await Promise.race([terminateOcr().then(() => 'ended'), new Promise(resolve => setTimeout(() => resolve('hung'), 3000))]);
            assert.strictEqual(ended, 'ended', 'Cancellation: terminateOcr() during a re-initialization that never finishes ends by the load timeout');
            assert.strictEqual(await hung, 'rejected', 'Cancellation: the job waiting on that re-initialization fails');
        } finally {
            ocrTestHooks.createWorker = undefined;
        }

        // A recognition outlasting the idle period is not cut short: the pool is idle only when no job is left.
        const idleCodes: string[] = [];
        const idle = await OfficeParser.parseOffice(file('docx'), { ocr: true, extractAttachments: true, ocrConfig: { timeout: { autoTerminate: 1 } }, onWarning: (w: any) => idleCodes.push(w.code) } as any);
        assert.ok(idle.attachments.some(a => a.ocrText) && !idleCodes.includes('OCR_FAILED'), `Cancellation: a recognition longer than autoTerminate completes (${idleCodes.join(', ')})`);

        // The OCR-only signal cancels the parse too, even with nothing left to recognize.
        const preAborted = new AbortController();
        preAborted.abort();
        let ocrSignalOutcome = 'resolved';
        await OfficeParser.parseOffice(file('docx'), { ocrConfig: { abortSignal: preAborted.signal }, onWarning: () => {} } as any).catch((e: any) => { ocrSignalOutcome = e.name; });
        assert.strictEqual(ocrSignalOutcome, 'AbortError', 'Cancellation: a fired ocrConfig.abortSignal rejects the parse');
    } finally {
        ocrTestHooks.afterRecognizeCall = undefined;
        ocrTestHooks.createWorker = undefined;
        await terminateOcr();
    }
    console.log('  Cancellation: All assertions passed ✓');
}

/**
 * Markdown and HTML round trips keep what the author wrote. Literal text that looks like a character
 * reference survives repeated saves (it lost one level of escaping each time), including through
 * HTML the way an editor saves (md -> HTML -> editor -> HTML -> md); references decode everywhere a
 * renderer decodes them (HTML text and attributes, Markdown alt text and titles); a one-line code
 * block stays a block; and blocks are separated by exactly one blank line.
 */
async function testMarkdownRoundTrips(): Promise<void> {
    const md = async (src: string) => ((await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).to('md', { generateIds: false } as any)).value as string).trim();
    const viaHtml = async (src: string) => {
        const html = (await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).to('html', { htmlConfig: { sourceAttributes: true } } as any)).value as string;
        const back = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' } as any);
        return ((await back.to('md', { generateIds: false, mdConfig: { dialect: 'extended' } } as any)).value as string).replace(/^---\n[^]*?\n---\n\n/, '').trim();
    };
    const stable = async (save: (s: string) => Promise<string>, src: string, expected: string, label: string) => {
        let text = src;
        for (let pass = 1; pass <= 3; pass++) {
            text = await save(text);
            assert.strictEqual(text, expected, `MD round trip: ${label} (save ${pass})`);
        }
    };

    // Literal reference-like text keeps its meaning; a bare & is left alone.
    for (const save of [md, viaHtml]) {
        const how = save === md ? 'md' : 'md -> html -> md';
        await stable(save, 'Text &amp;quot; literal.', 'Text &amp;quot; literal.', `literal &quot; (${how})`);
        await stable(save, 'Tom &amp;amp; Jerry', 'Tom &amp;amp; Jerry', `literal &amp; (${how})`);
        await stable(save, 'Tom & Jerry, a && b', 'Tom & Jerry, a && b', `a bare & (${how})`);
        await stable(save, 'x &#39; &#x27; &copy; &foo; &constructor;', "x ' ' © &amp;foo; &amp;constructor;", `references decoded, literal names kept (${how})`);
        await stable(save, '![a &amp;quot; b](x.png "say &quot;hi&quot;")', '![a &amp;quot; b](x.png "say &quot;hi&quot;")', `alt text and title (${how})`);
        await stable(save, '```\nls -la\n```', '```\nls -la\n```', `a one-line code block stays a block (${how})`);
    }
    await stable(md, 'code `&quot;` and $a &= b$ stay', 'code `&quot;` and $a &= b$ stay', 'code spans and math are not escaped');
    // A link target is decoded as a renderer decodes it, and a literal reference in it is escaped.
    await stable(md, '[l](http://x.com/?a=1&b=2)', '[l](http://x.com/?a=1&b=2)', 'a query string');
    await stable(md, '[l](http://x.com/?q=&amp;copy;)', '[l](http://x.com/?q=&amp;copy;)', 'a literal reference in a link target');
    const fromHtml = await OfficeParser.parseOffice(Buffer.from('<p><a href="http://x.com/?q=&amp;copy;&amp;b=2">l</a></p>'), { fileType: 'html' } as any);
    assert.strictEqual((await fromHtml.to('md')).value, '[l](http://x.com/?q=&amp;copy;&b=2)', 'MD: a URL holding a literal &copy; keeps it for a renderer');
    // A title escaping its quotes (as other tools write it) is read, and written back as a reference.
    const titled = await OfficeParser.parseOffice(Buffer.from('[l](http://x.com "say \\"hi\\"") ![a &amp; b](i.png)'), { fileType: 'md' } as any);
    const [link, , image] = titled.content[0].children!;
    assert.deepStrictEqual([(link.metadata as any).link, (link.metadata as any).title, (image.metadata as any).altText], ['http://x.com', 'say "hi"', 'a & b'], 'MD: a title may escape its quotes; alt text is decoded');
    await stable(md, '[l](http://x.com "say \\"hi\\"")', '[l](http://x.com "say &quot;hi&quot;")', 'a title with quotes');
    // A fence is read as CommonMark reads it: any info string (its first word is the language),
    // up to three spaces of indentation, a longer closing fence; and inside a list item, quote or
    // admonition it stays a code block (in an admonition, inside it).
    for (const [src, lang, code] of [
        ['```c++\nint x;\n```', 'c++', 'int x;'], ['```objective-c\nid x;\n```', 'objective-c', 'id x;'],
        ['```js title="a.js"\nlet x;\n```', 'js', 'let x;'], ['``` js\nlet y;\n```', 'js', 'let y;'],
        ['   ```\n   indented\n   ```', '', 'indented'], ['````\na\n```\nb\n`````', '', 'a\n```\nb'],
    ] as const) {
        const fenced = (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).content[0];
        assert.deepStrictEqual([fenced.type, (fenced.metadata as any)?.language, fenced.text], ['code', lang, code], `MD: fence ${JSON.stringify(src)}`);
    }
    await stable(md, '- item\n\n  ```js\n  x\n  ```\n- next', '- item\n\n```js\nx\n```\n\n- next', 'a fenced block in a list item stays a code block');
    await stable(md, '> quote\n>\n> ```js\n> code\n> ```\n>\n> after', '> quote\n\n```js\ncode\n```\n\n> after', 'a fenced block in a quote stays a code block');
    await stable(md, '> [!NOTE]\n> text\n>\n> ```sh\n> ls\n> ```', '> [!NOTE]\n> text\n>\n> ```sh\n> ls\n> ```', 'a fenced block in an admonition stays inside it');
    // A reference name the object prototype has is not a character.
    const proto = await OfficeParser.parseOffice(Buffer.from('a &constructor; b'), { fileType: 'md' } as any);
    assert.strictEqual(proto.content[0].children!.map(c => c.text).join(''), 'a &constructor; b', 'MD: &constructor; is literal text');

    // One blank line around every block, and before the abbreviation definitions.
    await stable(md, 'Energy is $$E=mc^2$$ here.', 'Energy is\n\n$$\nE=mc^2\n$$\n\nhere.', 'display math split out of a paragraph');
    await stable(md, 'Text $$a+b$$ more\nHeading\n===', 'Text\n\n$$\na+b\n$$\n\nmore\n\n# Heading', 'display math in the text before a setext heading');
    for (const [label, src] of [
        ['math', 'P.\n\n$$\nx\n$$\n\nQ.'], ['fenced code', 'P.\n\n```js\nx\n```\n\nQ.'], ['table', 'P.\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\nQ.'],
        ['abbreviations', 'R&D.\n\n*[R&D]: Research'],
    ] as const) {
        await stable(md, src, src, `one blank line around ${label}`);
    }
    // Code in a pipe-table cell, where a fence cannot go, stays an inline span.
    const cellCode = await OfficeParser.parseOffice(Buffer.from('<table><tr><th>a</th></tr><tr><td><pre><code>x</code></pre></td></tr></table>'), { fileType: 'html' } as any);
    assert.ok(/\| `x` \|/.test((await cellCode.to('md')).value as string), 'MD: code in a pipe-table cell stays an inline span');
    const pre = (await (await OfficeParser.parseOffice(Buffer.from('```\nls\n```'), { fileType: 'md' } as any)).to('html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(/<pre><code>ls<\/code><\/pre>/.test(pre), `HTML: a one-line code block is a <pre> (${pre})`);

    // Text reads back as itself: its Markdown characters are escaped where they would be markup (a
    // block marker only where a line starts), and an underscore inside a word is left as it is.
    for (const save of [md, viaHtml]) {
        const how = save === md ? 'md' : 'md -> html -> md';
        for (const src of [
            'Price: \\$5-\\$10 and 2\\*3\\*4 = 24, snake_case_name, -8 and +7',
            'a \\[bracket\\], \\[link\\](x), !\\[img\\](y), \\[^1\\], \\[@cite\\] and \\[\\[wiki\\]\\]',
            'C:\\path\\\\(x), \\`tick\\`, \\~tilde\\~, \\=\\=mark\\=\\=, \\{#id} and trailing\\\\',
            '\\# not a heading', '\\- not a list', '1\\. not ordered', '\\> not a quote', '\\---', '\\: not a definition',
        ]) await stable(save, src, src, `literal Markdown characters ${JSON.stringify(src)} (${how})`);
        await stable(save, 'a  \n\\\nb', 'a  \n\\\nb', `consecutive hard breaks stay in the paragraph (${how})`);
    }
    const words = await OfficeParser.parseOffice(Buffer.from('snake_case_name and a_b_c'), { fileType: 'md' } as any);
    assert.deepStrictEqual(words.content[0].children!.map(c => [c.text, !!c.formatting?.italic]), [['snake_case_name and a_b_c', false]], 'MD: an underscore inside a word is not emphasis');

    // Emphasis a renderer reads as emphasis: whitespace stays outside the delimiters (`**Note: **` is
    // not bold in CommonMark), bold and italic together read back as both, and `_` emphasis inside a
    // word or next to other emphasis takes the `*` form.
    const T = (text: string, formatting?: object) => ({ type: 'text', text, ...(formatting && { formatting }) });
    const BR = { type: 'break', metadata: { breakType: 'carriageReturn' } };
    const gen = async (children: any[], dialect: any = 'extended') => ((await OfficeGenerator.generate({ type: 'md', metadata: {}, content: [{ type: 'paragraph', children }], attachments: [] } as any, 'md', { mdConfig: { dialect } } as any)).value as string).trim();
    const underscore = { extends: 'extended', emphasisMarker: 'underscore' };
    for (const [children, dialect, expected, label] of [
        [[T('Note: ', { bold: true }), T('body')], 'extended', '**Note:** body', 'whitespace outside emphasis'],
        [[T('a '), T('both', { bold: true, italic: true }), T(' b')], 'extended', 'a ***both*** b', 'bold and italic'],
        [[T('x', { bold: true }), T('y', { italic: true })], underscore, '**x**_y_', 'underscore emphasis never touches another'],
        [[T('un'), T('believ', { italic: true }), T('able')], underscore, 'un*believ*able', 'emphasis inside a word'],
        [[T('a '), T('b', { italic: true }), T(' c')], underscore, 'a _b_ c', 'underscore emphasis where it fits'],
        [[T('a'), BR, BR, T('b')], 'extended', 'a  \n\\\nb', 'a hard break starting a line is a backslash'],
    ] as const) {
        assert.strictEqual(await gen(children as any, dialect), expected, `MD: ${label}`);
    }
    await stable(md, 'a ***both*** b and ___also___', 'a ***both*** b and ***also***', 'bold and italic together read back as both');
    // Code spans keep their spaces and backticks, as CommonMark reads a padded span.
    await stable(md, '`` `x` `` and `  a  `', '`` `x` `` and `  a  `', 'code spans keep their spaces and backticks');
    // Titles may hold parentheses; alt text and link text may hold brackets and backslashes.
    await stable(md, '[see](http://h/x "Figure 1 (a)") ![arr\\[0\\] C:\\path\\\\(y)](a.png "t\\\\* (1)")', '[see](http://h/x "Figure 1 (a)") ![arr\\[0\\] C:\\path\\\\(y)](a.png "t\\\\* (1)")', 'titles, alt text');
    const nested = await OfficeParser.parseOffice(Buffer.from('[a [b [c]] d](e) and [' + 'w'.repeat(5000) + '](f)'), { fileType: 'md' } as any);
    const nestedLinks = nested.content[0].children!.filter(c => (c.metadata as any)?.link);
    assert.deepStrictEqual([nestedLinks.filter(c => (c.metadata as any).link === 'e').map(c => c.text).join(''), nestedLinks.filter(c => (c.metadata as any).link === 'f').map(c => c.text).join('').length], ['a [b [c]] d', 5000], 'MD: link text may hold nested brackets, and be long');
    await stable(md, '[a \\[b\\] c](e)', '[a \\[b\\] c](e)', 'link text holding brackets');
    // A pipe in a table cell, in text or code, stays in its cell.
    await stable(md, '| a | b |\n| --- | --- |\n| x\\|y | `p\\|q` |', '| a | b |\n| --- | --- |\n| x\\|y | `p\\|q` |', 'a pipe in a table cell');
    // A fence indented under a list item (four spaces, as editors write it) is a fence.
    await stable(md, '- item\n\n    ```js\n    x\n    ```\n\n- next', '- item\n\n```js\nx\n```\n\n- next', 'a fence indented under a list item');
    // A math line that is itself `$$` (or an indented one) reads back as written.
    const mathMd = (await OfficeGenerator.generate({ type: 'md', metadata: {}, content: [{ type: 'code', text: 'a\n$$\n $$\n\n\nb', metadata: { math: 'block' } }], attachments: [] } as any, 'md')).value as string;
    const mathBack = await OfficeParser.parseOffice(Buffer.from(mathMd), { fileType: 'md' } as any);
    assert.strictEqual(mathBack.content[0].text, 'a\n$$\n $$\n\n\nb', `MD: a math line of $$ reads back as written (${JSON.stringify(mathMd)})`);
    // An embed's label is escaped (it cannot write a tag) and a directive label reads back as written.
    const embedAst = { type: 'md', metadata: {}, content: [{ type: 'embed', text: 'https://www.youtube.com/watch?v=abc', metadata: { embedType: 'youtube', videoId: 'abc', label: 'A <b>bold</b> & [odd] label' } }], attachments: [] };
    for (const embeds of ['directive', 'link', 'thumbnail'] as const) {
        const out = (await OfficeGenerator.generate(embedAst as any, 'md', { mdConfig: { dialect: { extends: 'extended', embeds } } } as any)).value as string;
        assert.ok(!out.includes('<b>'), `MD: an embed label cannot write a tag (${embeds}: ${out})`);
        if (embeds === 'directive') {
            const back = await OfficeParser.parseOffice(Buffer.from(out), { fileType: 'md' } as any);
            assert.strictEqual((back.content[0].metadata as any)?.label, 'A <b>bold</b> & [odd] label', `MD: a directive label reads back as written (${out})`);
        }
    }

    // MDX components are stripped and their content kept, except in code, where they are code.
    const mdx = await OfficeParser.parseOffice(Buffer.from('<Callout>\nSee `<Br/>` and <Badge />! <Open>unclosed\n</Callout>\n\n```jsx\n<Button>click</Button>\n```'), { fileType: 'md' } as any);
    assert.deepStrictEqual(mdx.content.map(n => (n.type === 'code' ? n.text : n.children!.map(c => c.text).join(''))), ['See <Br/> and ! <Open>unclosed', '<Button>click</Button>'], 'MD: MDX components are stripped, not in code');
    // A heading's id is the `{#id}` ending it.
    const headingIds = await OfficeParser.parseOffice(Buffer.from('## a {#x}\n\n## b{#y} {#z}\n\n## c {#p {#q}'), { fileType: 'md' } as any);
    assert.deepStrictEqual(headingIds.content.map(h => [h.text, (h.metadata as any).anchorIds]), [['a', ['x']], ['b{#y}', ['z']], ['c', ['p {#q']]], 'MD: heading ids');

    // HTML whitespace as a browser shows it: a space between two inline elements is kept (two code
    // spans stay two), spaces at a line's edges and around <br> are not shown, a lone &nbsp; spacer
    // is layout, and &nbsp; between words is content.
    const ws = await OfficeParser.parseOffice(Buffer.from('<p> <code>a</code> <code>b</code><br>\n c&nbsp;&nbsp;d </p><p>&nbsp;</p><p><b>x </b> <i>y</i></p>'), { fileType: 'html' } as any);
    assert.deepStrictEqual(ws.content.map(b => b.children!.map(c => (c.type === 'text' ? c.text : c.type))), [['a', ' ', 'b', 'break', 'c\u00a0\u00a0d'], [], ['x ', 'y']], 'HTML: whitespace as a browser shows it');
    // A code block inside a paragraph is written beside it: HTML cannot nest a <pre> in a <p>.
    const split = (await (await OfficeParser.parseOffice(Buffer.from('<p id="x">a <pre>code</pre> b</p>'), { fileType: 'html' } as any)).to('html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(!/<p[^>]*>(?:(?!<\/p>)[\s\S])*<pre/.test(split) && /<p id="x">a ?<\/p>/.test(split) && (split.match(/id="x"/g) || []).length === 1, `HTML: a code block is not written inside a paragraph (${split})`);

    // A table written as HTML in Markdown (merged cells, or the commonmark dialect) keeps what its cells
    // hold, and reads back to the same table: formatting, a link, a line break, code, math, literal
    // Markdown characters, a footnote and a linked picture. Text around it in its block is kept.
    const cellText = (text: string, extra: object = {}) => ({ type: 'text', text, ...extra });
    const htmlTableAst = { type: 'md', metadata: {}, attachments: [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: fs.readFileSync(path.join(__dirname, '..', 'docs', 'favicon.png')).toString('base64') }], content: [{ type: 'table', children: [
        { type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2 }, children: [{ type: 'paragraph', children: [cellText('Head')] }] }] },
        { type: 'row', children: [
            { type: 'cell', children: [{ type: 'paragraph', children: [cellText('a', { formatting: { bold: true } }), cellText(' link', { metadata: { link: 'https://x.y', linkType: 'external' } }), { type: 'break', metadata: { breakType: 'carriageReturn' } }, cellText('code', { formatting: { font: 'monospace' } }), { type: 'code', text: 'x^2', metadata: { math: 'inline' } }, cellText(' 2*3 & <b>')] }] },
            { type: 'cell', children: [{ type: 'paragraph', children: [cellText('b', { notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [cellText('note text')] }] }), { type: 'image', metadata: { attachmentName: 'p.png', altText: 'pic', link: 'https://l.k', linkType: 'external' } }] }] },
        ] },
    ] }] } as any;
    const htmlTableMd = (await OfficeGenerator.generate(htmlTableAst, 'md')).value as string;
    const htmlTableBack = await OfficeParser.parseOffice(Buffer.from(htmlTableMd), { fileType: 'md', extractAttachments: true } as any);
    const backCells = htmlTableBack.content[0].children!.map(r => r.children!.map(c => c.children![0].children!.map(n => (n.type === 'text' ? [n.text, n.formatting?.bold || n.formatting?.font || (n.metadata as any)?.link || n.notes?.[0]?.children?.[0]?.text || ''] : [n.type, (n.metadata as any)?.math || (n.metadata as any)?.link || '']))));
    assert.deepStrictEqual(backCells, [[[['Head', '']]], [[['a', true], [' link', 'https://x.y'], ['break', ''], ['code', 'monospace'], ['code', 'inline'], [' 2*3 & <b>', '']], [['b', 'note text'], ['image', 'https://l.k']]]], `MD: an HTML table's cells read back as written (${htmlTableMd.slice(0, 400)})`);
    assert.strictEqual((await OfficeGenerator.generate(htmlTableBack, 'md')).value, htmlTableMd, 'MD: an HTML table is stable over saves');
    const aroundTable = await OfficeParser.parseOffice(Buffer.from('Intro\n<table><tr><td>x</td></tr></table>'), { fileType: 'md' } as any);
    assert.deepStrictEqual(aroundTable.content.map(n => n.type), ['paragraph', 'table'], 'MD: text beside an HTML table in its block is kept');

    // HTML's optional end tags, as a browser reads them: an unclosed paragraph, list item, row or cell
    // ends at the next one instead of holding it.
    const omitted = await OfficeParser.parseOffice(Buffer.from('<p>a<p>b<ul><li>c<li>d</ul><table><tr><td>1<td>2<tr><td>3</table>'), { fileType: 'html' } as any);
    assert.deepStrictEqual(omitted.content.map(n => (n.type === 'table' ? n.children!.map(r => r.children!.length) : [n.type, n.children!.map(c => c.text).join('')])), [['paragraph', 'a'], ['paragraph', 'b'], ['list', 'c'], ['list', 'd'], [2, 1]], 'HTML: omitted end tags');

    // Only spaces, tabs and line endings are whitespace to Markdown: a no-break or em space at the edge
    // of a paragraph, heading, list item, quote or cell is text, and one before a list marker makes the
    // line a paragraph. (JavaScript's trim() and \s take them too, and each save lost them.)
    const exact = async (src: string) => (await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).to('md', { generateIds: false } as any)).value as string;
    for (const src of ['\u00a0a\u2003', '# \u00a0Title\u00a0', '- \u2003item', '1. \u00a0one', '> \u00a0quoted', '| \u00a0a | b |\n| --- | --- |\n| c\u00a0 | d |', '\u00a0- not a list']) {
        await stable(exact, src, src, `Unicode spaces kept in ${JSON.stringify(src)}`);
    }
    assert.strictEqual((await OfficeParser.parseOffice(Buffer.from('\u00a0- x'), { fileType: 'md' } as any)).content[0].type, 'paragraph', 'MD: a no-break space before a list marker makes a paragraph');
    // A backslash escaping a line's last backslash is not a hard break; a table may have no body rows;
    // empty anchors starting a paragraph or a list item are its anchor ids.
    const escapedBreak = await OfficeParser.parseOffice(Buffer.from('a\\\\\nb'), { fileType: 'md' } as any);
    assert.strictEqual(escapedBreak.content[0].children!.map(c => (c.type === 'break' ? '<br>' : c.text)).join(''), 'a\\ b', 'MD: an escaped trailing backslash is text, not a hard break');
    await stable(md, '| a | b |\n| --- | --- |', '| a | b |\n| --- | --- |', 'a table with no body rows');
    const anchored = await OfficeParser.parseOffice(Buffer.from('<a id="x"></a>Text\n\n- <a id="y"></a>item'), { fileType: 'md' } as any);
    assert.deepStrictEqual(anchored.content.map(n => [n.type, (n.metadata as any)?.anchorIds, n.children!.map(c => c.text).join('')]), [['paragraph', ['x'], 'Text'], ['list', ['y'], 'item']], 'MD: leading anchors are anchor ids');
    await stable(md, '<a id="x"></a>Text\n\n- <a id="y"></a>item', '<a id="x"></a>Text\n\n- <a id="y"></a>item', 'leading anchors');

    // Blocks in a container stand apart as top-level blocks do, and an endnote standing among them is
    // written with the definitions at the end: a heading after a picture on a page stays a heading, and
    // a paragraph after a list item stays a paragraph. An image's anchors go on its line, as a
    // paragraph's do; runs differing only in what Markdown does not write are one span; a row missing a
    // cell is filled with an empty one.
    const text = (t: string, extra: object = {}) => ({ type: 'text', text: t, ...extra });
    const writeMd = async (content: any[]) => (await OfficeGenerator.generate({ type: 'pdf', metadata: {}, attachments: [], content } as any, 'md')).value as string;
    assert.strictEqual(await writeMd([{ type: 'page', children: [
        { type: 'image', metadata: { url: 'a.png', altText: 'pic' } }, { type: 'note', metadata: { noteType: 'endnote', noteId: '9' }, children: [{ type: 'paragraph', children: [text('an endnote')] }] },
        { type: 'heading', metadata: { level: 1 }, children: [text('Title')] }, { type: 'list', metadata: { listType: 'ordered', listId: 'l', indentation: 0, itemIndex: 0 }, children: [text('one')] },
        { type: 'paragraph', children: [text('7. after')] },
    ] }]), '---\n\n![pic](a.png)\n\n# Title {#title}\n\n1. one\n\n7\\. after\n\n[^9]: an endnote', 'MD: blocks on a page stand apart');
    assert.strictEqual(await writeMd([{ type: 'paragraph', children: [text('x '), { type: 'image', metadata: { url: 'a.png', altText: 'pic', anchorIds: ['fig'] } }, text(' y')] }]), 'x <a id="fig"></a>![pic](a.png) y', 'MD: an image\'s anchors go on its line');
    assert.strictEqual(await writeMd([{ type: 'paragraph', children: [
        text('see the', { formatting: { font: 'A' } }), text(' ', { formatting: { font: 'A', color: '#0000ff' }, metadata: { link: 'http://x.y', linkType: 'external' } }),
        text('page', { formatting: { font: 'A' }, metadata: { link: 'http://x.y', linkType: 'external' } }),
    ] }]), 'see the[ page](http://x.y)', 'MD: runs differing only in an unwritten colour are one link');
    assert.strictEqual(await writeMd([{ type: 'table', children: [
        { type: 'row', children: [{ type: 'cell', children: [text('a')] }, { type: 'cell', children: [text('b')] }] }, { type: 'row', children: [{ type: 'cell', children: [text('c')] }] },
    ] }]), '| a | b |\n| --- | --- |\n| c |  |', 'MD: a missing cell is written empty');
    // So a document from any format reads back from its Markdown as the same Markdown.
    for (const file of ['test.docx', 'test.pptx', 'test.xlsx', 'test.odt', 'test.odp', 'test.ods', 'test.rtf', 'test.html', 'test.tex', 'test.pdf']) {
        const first = (await (await OfficeParser.parseOffice(path.join(__dirname, 'files', file), { extractAttachments: true } as any)).to('md')).value as string;
        const second = (await (await OfficeParser.parseOffice(Buffer.from(first), { fileType: 'md', extractAttachments: true } as any)).to('md')).value as string;
        let at = 0;
        while (at < first.length && first[at] === second[at]) at++;
        assert.ok(first === second, `MD: ${file} saves the same Markdown twice (first difference at ${at}: ${JSON.stringify(first.slice(at - 40, at + 40))} became ${JSON.stringify(second.slice(at - 40, at + 40))})`);
    }

    // Emphasis is read as CommonMark reads it: a delimiter run followed by whitespace opens nothing
    // (`5 * 3 * 2` holds no emphasis), and an escaped delimiter before the closing run is text.
    // (A closer after whitespace still closes: earlier versions wrote `**Note: **body`.)
    for (const [src, expected] of [
        ['5 * 3 * 2', '5 \\* 3 \\* 2'], ['a ** b ** c', 'a \\*\\* b \\*\\* c'], ['a ~~ b ~~ c', 'a \\~\\~ b \\~\\~ c'],
        ['if a == b or a === c', 'if a \\=\\= b or a \\=\\=\\= c'], ['x* y*', 'x\\* y\\*'], ['**a\\*** b', '**a\\*** b'], ['~~a\\~~~ b', '~~a\\~~~ b'],
    ] as const) await stable(md, src, expected, `emphasis ${JSON.stringify(src)}`);
    const legacyBold = await OfficeParser.parseOffice(Buffer.from('**Note: **body'), { fileType: 'md' } as any);
    assert.deepStrictEqual(legacyBold.content[0].children!.map(c => [c.text, !!c.formatting?.bold]), [['Note: ', true], ['body', false]], 'MD: a closer after whitespace still closes');
    for (const [label, formatting] of [['bold', { bold: true }], ['italic', { italic: true }], ['strikethrough', { strikethrough: true }]] as const) {
        const written = await gen([T(label === 'strikethrough' ? 'a~' : 'a*', formatting), T(' after')]);
        const back = await OfficeParser.parseOffice(Buffer.from(written), { fileType: 'md' } as any);
        assert.strictEqual(back.content[0].children!.map(c => c.text).join(''), label === 'strikethrough' ? 'a~ after' : 'a* after', `MD: ${label} text ending in its delimiter reads back (${written})`);
    }

    // Blocks as CommonMark reads them: a thematic break may have spaces and is never a list item; a
    // setext underline ends its heading's block; an item interrupts a paragraph only if it has content
    // and, when ordered, starts at 1; a marker alone is an empty item; four columns of indentation make
    // code, even of a line starting `#`; `>` needs no space after it.
    for (const [src, types] of [
        ['* * *', ['break']], ['- - -', ['break']], ['_ _ _', ['break']], ['text\n***', ['paragraph', 'break']], ['text\n-\nmore', ['heading', 'paragraph']],
        ['text\n2. not a list', ['paragraph']], ['text\n1. a list', ['paragraph', 'list']], ['-\n- b', ['list', 'list']], ['    # comment\n    ls -la', ['code']],
        ['>quote', ['paragraph']], ['#', ['heading']],
    ] as const) {
        const blocks = await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any);
        assert.deepStrictEqual(blocks.content.map(n => n.type), types, `MD: blocks of ${JSON.stringify(src)}`);
        await stable(md, await md(src), await md(src), `blocks of ${JSON.stringify(src)} over saves`);
    }
    // A quote or an admonition holds blocks: a list, heading or table in it keeps its structure.
    for (const [src, expected] of [
        ['> - a\n> - b', '- a\n- b'], ['> # Heading\n> text', '# Heading\n\n> text'], ['> para\n>\n> 1. one\n> 2. two', '> para\n\n1. one\n2. two'],
        ['> [!NOTE]\n> - a\n> - b', '> [!NOTE]\n> - a\n> - b'], ['> a  \n> b', '> a  \nb'],
    ] as const) await stable(md, src, expected, `quote ${JSON.stringify(src)}`);

    // A footnote may hold several paragraphs, a list, and references to other notes.
    for (const src of [
        'Body[^1] continues.\n\n[^1]: First para.\n\n    Second para.', 'Body[^1]\n\n[^1]: intro\n\n    - li',
        'x[^1]\n\n[^1]: body with [^2] inside\n\n[^2]: inner', 'x[^1]\n\n[^1]: body with [^2] inside\n\n    more\n\n[^2]: inner',
    ]) await stable(md, src, src, `footnote ${JSON.stringify(src)}`);

    // Inline: a code span may run across lines, a link target may hold parenthesis pairs, a link may
    // have no text, and a backslash ending a paragraph is text.
    await stable(md, 'x `a\nb` y', 'x `a b` y', 'a code span across lines');
    const wiki = await OfficeParser.parseOffice(Buffer.from('[w](https://en.wikipedia.org/wiki/Foo_(bar)) end'), { fileType: 'md' } as any);
    assert.deepStrictEqual(wiki.content[0].children!.map(c => [c.text, (c.metadata as any)?.link]), [['w', 'https://en.wikipedia.org/wiki/Foo_(bar)'], [' end', undefined]], 'MD: a link target holding parentheses');
    await stable(md, '[](http://x) after', '[](http://x) after', 'a link with no text');
    await stable(md, 'path C:\\dir\\', 'path C:\\dir\\\\', 'a backslash ending a paragraph');

    // Writing: an empty heading is its hashes (with its id); line breaks ending a paragraph, and a
    // paragraph of breaks, are not written; an empty paragraph's bookmark goes to the next block; code
    // and math in a list item or cell are inline; block markers starting a list item or definition
    // are escaped; a list starts at the left margin and nests one level at a time.
    const writeDoc = async (content: any[]) => (await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content } as any, 'md', { generateIds: false } as any)).value as string;
    const item = (children: any[], indentation = 0, itemIndex = 0) => ({ type: 'list', metadata: { listType: 'unordered', listId: 'l', indentation, itemIndex }, children });
    for (const [label, content, expected] of [
        ['an empty heading', [{ type: 'heading', metadata: { level: 1 }, children: [] }, { type: 'paragraph', children: [T('after')] }], '#\n\nafter'],
        ['an empty heading with an id', [{ type: 'heading', metadata: { level: 2, anchorIds: ['h'] }, children: [] }, { type: 'paragraph', children: [T('after')] }], '## {#h}\n\nafter'],
        ['breaks ending a paragraph', [{ type: 'paragraph', children: [T('a'), BR, BR] }, { type: 'paragraph', children: [BR] }, { type: 'paragraph', children: [T('next')] }], 'a\n\nnext'],
        ['a bookmark on an empty paragraph', [{ type: 'paragraph', metadata: { anchorIds: ['bm'] }, children: [] }, { type: 'paragraph', children: [T('text')] }], '<a id="bm"></a>text'],
        ['a rule inside a paragraph', [{ type: 'paragraph', children: [T('a'), { type: 'break', metadata: { breakType: 'thematic' } }, T('b')] }], 'a\n\n---\n\nb'],
        ['code in a list item', [item([T('see'), { type: 'code', text: 'a\nb' }])], '- see<br>`a`<br>`b`'],
        ['math in a list item', [item([T('see '), { type: 'code', text: 'x^2\n+1', metadata: { math: 'block' } }])], '- see<br>$x^2 +1$'],
        ['code in a cell', [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [T('h')] }] }, { type: 'row', children: [{ type: 'cell', children: [{ type: 'code', text: 'a\nb' }] }] }] }], '| h |\n| --- |\n| `a`<br>`b` |'],
        ['a list item starting with #', [item([T('# not a heading')])], '- \\# not a heading'],
        ['a definition term starting with -', [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [T('- term')] }, { type: 'definitionDescription', children: [T('def')] }] }], '\\- term\n: def'],
        ['a list starting deeper', [item([T('a')], 2), item([T('b')], 4, 1), item([T('c')], 0, 2)], '- a\n    - b\n- c'],
        ['an empty list item', [item([]), item([T('b')], 0, 1)], '-\n- b'],
    ] as const) {
        const written = await writeDoc(content as any);
        assert.strictEqual(written, expected, `MD: ${label}`);
        await stable(md, written, written, `${label} over saves`);
    }
    // A table written as HTML keeps a blank line inside its code: the Markdown HTML block ends at one.
    const htmlTableCode = await writeDoc([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2 }, children: [{ type: 'code', text: 'a\n\nb' }] }] }, { type: 'row', children: [{ type: 'cell', children: [T('a')] }, { type: 'cell', children: [T('b')] }] }] }]);
    assert.ok(!/\n[ \t]*\n/.test(htmlTableCode), `MD: no blank line inside an HTML table (${htmlTableCode})`);
    assert.strictEqual((await OfficeParser.parseOffice(Buffer.from(htmlTableCode), { fileType: 'md' } as any)).content[0].children![0].children![0].children![0].text, 'a\n\nb', 'MD: an HTML table cell\'s code keeps its blank line');
    // Front matter: a tag in a value is written `\u003c`, and quoted values decode their escapes.
    const frontAst = { type: 'md', metadata: { title: 'say "hi" <b>x</b>', author: 'A\\B', customProperties: { tags: ['a', '<script>', 'c,d'] } }, attachments: [], content: [{ type: 'paragraph', children: [T('t')] }] } as any;
    const frontMd = (await OfficeGenerator.generate(frontAst, 'md')).value as string;
    assert.ok(!frontMd.includes('<'), `MD: no tag in front matter (${frontMd})`);
    const frontBack = await OfficeParser.parseOffice(Buffer.from(frontMd), { fileType: 'md' } as any);
    assert.deepStrictEqual([frontBack.metadata.title, frontBack.metadata.author, (frontBack.metadata.customProperties as any)?.tags], ['say "hi" <b>x</b>', 'A\\B', ['a', '<script>', 'c,d']], 'MD: front matter reads back as written');

    // HTML: a rule in the body is a block, beside text too; a new item closes the one before it and a
    // paragraph open in it; a picture beside text in the body is part of its paragraph, one alone (or
    // with its caption) stands apart.
    for (const [src, shape] of [
        ['<p>a</p><hr><p>b</p>', 'paragraph|break|paragraph'], ['a<hr>b', 'paragraph|break|paragraph'],
        ['<dl><dt><p>a<dd>b<dt>c<dd>d</dl>', 'definitionList(definitionTerm,definitionDescription,definitionTerm,definitionDescription)'],
        ['<ul><li><p>a<li>b</ul>', 'list|list'], ['text <img src="a.png"> more', 'paragraph'], ['<div><img src="x.png"><div class="caption">cap</div></div>', 'image|paragraph'],
        ['<div class="image-container"><img src="x.png"><div class="caption">x.png</div></div>', 'image'],
        ['<div class="image-container"><img src="x.png"><div class="caption">image1.tmp</div></div>', 'image'],
        ['<div class="image-container"><img src="x.png"><div class="caption"><b>The team</b></div></div>', 'image|paragraph'],
    ] as const) {
        const parsed = await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'html' } as any);
        const shapeOf = (n: any): string => n.type === 'definitionList' ? `definitionList(${n.children.map((c: any) => c.type).join(',')})` : n.type;
        assert.strictEqual(parsed.content.map(shapeOf).join('|'), shape, `HTML: ${src}`);
    }
    const splitHtml = (await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [{ type: 'paragraph', children: [T('a'), { type: 'break', metadata: { breakType: 'page' } }, T('b')] }] } as any, 'html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(/<p>a<\/p><hr class="page-break">/.test(splitHtml) && !/<p>[^<]*<hr/.test(splitHtml), `HTML: a page break is not written inside a paragraph (${splitHtml})`);

    // HTML round trips add nothing: the caption HtmlGenerator writes under a picture (its file name) is a
    // label, and a picture in a paragraph is written inline; a <figcaption> or other caption stays a
    // block of its own.
    const htmlOnce = async (src: string) => (await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'html' } as any)).to('html', { htmlConfig: { standalone: false } } as any)).value as string;
    for (const src of ['<div><div class="image-container"><img src="a.png" alt=""><div class="caption">a.png</div></div></div>', 'text <img src="a.png" alt="A"> more',
        '<p>before</p><figure><img src="a.png" alt="A"><figcaption>cap</figcaption></figure><p>after</p>']) {
        const first = await htmlOnce(src);
        assert.strictEqual(await htmlOnce(first), first, `HTML: ${src} saves the same HTML twice (${first})`);
    }
    const figure = await OfficeParser.parseOffice(Buffer.from('<figure><img src="a.png" alt="A"><figcaption>cap</figcaption></figure>'), { fileType: 'html' } as any);
    assert.deepStrictEqual(figure.content.map(n => n.type), ['image', 'paragraph'], 'HTML: a figure caption is a block beside its picture');
    for (const file of ['test.html', 'test.epub']) {
        let ast: any = await OfficeParser.parseOffice(path.join(__dirname, 'files', file), {} as any);
        const counts: number[] = [];
        for (let save = 0; save < 3; save++) {
            counts.push(ast.content.length);
            ast = await OfficeParser.parseOffice(Buffer.from((await OfficeGenerator.generate(ast, 'html', { htmlConfig: { standalone: false } } as any)).value as string), { fileType: 'html' } as any);
        }
        assert.ok(counts[1] === counts[0] && counts[2] === counts[0], `HTML: ${file} keeps its blocks over HTML saves (${counts.join(', ')})`);
    }

    // Markdown: code and lists in a definition, rules in a list item or cell, a bookmark on an empty
    // paragraph before a heading, quote or code block, a rule in an admonition, a note-carrying run,
    // and a note cited by a note defined after it all save the same Markdown twice.
    const dlCode = await OfficeParser.parseOffice(Buffer.from('<dl><dt>t</dt><dd><pre>a\nb</pre></dd></dl>'), { fileType: 'html' } as any);
    assert.strictEqual((await dlCode.to('md', { generateIds: false } as any)).value, 't\n: `a`<br>`b`', 'MD: code in a definition is inline');
    for (const [label, content] of [
        ['a list in a definition', [{ type: 'definitionList', children: [{ type: 'definitionTerm', children: [T('t')] }, { type: 'definitionDescription', children: [item([T('x')]), item([T('y')], 0, 1)] }] }]],
        ['a rule in a list item', [item([T('a'), { type: 'break', metadata: { breakType: 'thematic' } }, T('b')])]],
        ['a bookmark before a heading', [{ type: 'paragraph', metadata: { anchorIds: ['a'] }, children: [] }, { type: 'heading', metadata: { level: 2 }, children: [T('x')] }, { type: 'paragraph', children: [T('more')] }]],
        ['a bookmark before a closing quote', [{ type: 'paragraph', metadata: { anchorIds: ['a'] }, children: [] }, { type: 'paragraph', metadata: { style: 'Quote' }, children: [T('q')] }]],
        ['a bookmark before a closing code block', [{ type: 'paragraph', metadata: { anchorIds: ['a'] }, children: [] }, { type: 'code', text: 'c' }]],
        ['a rule in an admonition', [{ type: 'admonition', metadata: { admonitionType: 'note' }, children: [{ type: 'paragraph', children: [T('a'), { type: 'break', metadata: { breakType: 'thematic' } }, T('b')] }] }]],
        ['bold text ending in its delimiter, with a note', [{ type: 'paragraph', children: [T('a*', { formatting: { bold: true }, notes: [{ type: 'note', metadata: { noteType: 'footnote', noteId: '1' }, children: [T('n')] }] })] }]],
    ] as const) {
        const written = await writeDoc(content as any);
        await stable(md, written, written, `${label} (${written})`);
    }
    assert.strictEqual(await md('[^b]: bbb\n\n[^a]: see [^b]\n\nbody'), 'body\n\n[^b]: bbb\n\n[^a]: see [^b]', 'MD: a note cited by a note defined after it is written once');
    // A quote followed by a rule is a quote and a rule (not a heading); a line of spaces is a blank line.
    for (const [src, types] of [['> quote\n---', ['paragraph', 'break']], ['a\n  \nb', ['paragraph', 'paragraph']]] as const) {
        assert.deepStrictEqual((await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).content.map(n => n.type), types, `MD: blocks of ${JSON.stringify(src)}`);
    }
    // An unknown node's link reaches its text at block level too.
    assert.strictEqual((await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: [{ type: 'mystery', children: [T('inner')], metadata: { link: 'http://l.k', linkType: 'external' } }] } as any, 'md')).value, '[inner](http://l.k)', 'MD: an unknown node keeps its link');

    // HTML captions: the caption HtmlGenerator writes (a file name) is a label, any other caption in an
    // image container is text; a caption holding blocks keeps them as blocks.
    const htmlAst = (src: string) => OfficeParser.parseOffice(Buffer.from(src), { fileType: 'html' } as any);
    const plain = (nodes: any[]): string => nodes.map((n: any) => n.text ?? plain(n.children ?? [])).join('|');
    const realCaption = await htmlAst('<div class="image-container"><img src="photo.png"><div class="caption">The team in 2024</div></div>');
    assert.deepStrictEqual([realCaption.content.map(n => n.type), plain(realCaption.content)], [['image', 'paragraph'], '|The team in 2024'], 'HTML: a caption of text in an image container is kept');
    const blockCaption = await htmlAst('<figure><img src="a.png"><figcaption><p>line one</p><p>line two</p></figcaption></figure>');
    assert.deepStrictEqual(blockCaption.content.map(n => n.type), ['image', 'paragraph', 'paragraph'], 'HTML: a caption of paragraphs is those paragraphs');
    // A picture in a paragraph touches the text around it, and its id is on the <img>; the second and
    // later ids of a node, written as anchors before it, are read back as its ids.
    for (const src of ['<p>see <img src="a.png" alt="A">.</p>', '<p>text <img id="pic1" src="a.png" alt="A"> more</p>', '<a id="x1" name="x1"></a><a id="x2" name="x2"></a><p>para</p>']) {
        const first = await htmlOnce(src);
        assert.strictEqual(await htmlOnce(first), first, `HTML: ${src} saves the same HTML twice (${first})`);
    }
    assert.ok((await htmlOnce('<p>see <img src="a.png" alt="A">.</p>')).includes('alt="A">.</p>'), 'HTML: nothing is written between a picture and the text after it');
    const idsOf = (nodes: any[]): any[] => nodes.flatMap((n: any) => [...(n.metadata?.anchorIds ? [[n.type, n.metadata.anchorIds]] : []), ...idsOf(n.children ?? [])]);
    for (const [src, ids] of [
        ['<p>text <img id="pic1" src="a.png" alt="A"> more</p>', [['image', ['pic1']]]],
        ['<a id="x1" name="x1"></a><a id="x2" name="x2"></a><p>para</p>', [['paragraph', ['x1', 'x2']]]],
        ['<p id="p1">see <a id="in"></a>here <a name="b1"></a><img id="pic2" src="b.png"></p>', [['paragraph', ['p1', 'in']], ['image', ['pic2', 'b1']]]],
        ['text <a id="d" name="d"></a><h2>H</h2>', [['heading', ['d']]]],
        ['<table><tr><td id="c1">a</td></tr></table><hr id="r"><dl id="l"><dt id="t">t</dt><dd>d</dd></dl><blockquote id="q">q</blockquote>', [['cell', ['c1']], ['break', ['r']], ['definitionList', ['l']], ['definitionTerm', ['t']], ['paragraph', ['q']]]],
        ['<div><a id="only"></a></div>', [['paragraph', ['only']]]],
    ] as const) {
        assert.deepStrictEqual(idsOf((await htmlAst(src)).content), ids, `HTML: the ids of ${src}`);
    }
    // A sheet written as HTML reads back as its data: the row numbers and column letters HtmlGenerator
    // draws around it are not cells of it.
    const sheetHtml = (await OfficeGenerator.generate({ type: 'xlsx', metadata: {}, attachments: [], content: [{ type: 'sheet', metadata: { sheetName: 'S' }, children: [{ type: 'row', children: [{ type: 'cell', children: [T('a')] }, { type: 'cell', children: [T('b')] }] }, { type: 'row', children: [{ type: 'cell', children: [T('c')] }, { type: 'cell', children: [T('d')] }] }] }] } as any, 'html', { htmlConfig: { standalone: false } } as any)).value as string;
    const sheetBack = (await htmlAst(sheetHtml)).content.find(n => n.type === 'table')!;
    assert.deepStrictEqual(sheetBack.children!.map(row => row.children!.map(cell => plain(cell.children ?? []))), [['a', 'b'], ['c', 'd']], 'HTML: a sheet reads back without its row numbers and column letters');
    // A sheet with merged cells, or under a dialect without pipe tables, is written to Markdown as an HTML
    // table (it was written as nothing), and content among a table's rows (a CSV comment line) is a
    // row of one cell where it stands (it was an empty row). RTF writes code blocks, equations (as their
    // LaTeX), definition terms and descriptions, and a CSV comment line, which it dropped or ran together.
    const mergedSheet = { type: 'xlsx', metadata: {}, attachments: [], content: [{ type: 'sheet', metadata: { sheetName: 'S' }, children: [
        { type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0, colSpan: 2 }, children: [T('merged')] }] },
        { type: 'comment', text: '# a comment line' },
        { type: 'row', children: [{ type: 'cell', metadata: { row: 2, col: 0 }, children: [T('c')] }, { type: 'cell', metadata: { row: 2, col: 1 }, children: [T('d')] }] }] }] } as any;
    for (const dialect of ['extended', 'commonmark'] as const) {
        const written = (await OfficeGenerator.generate(mergedSheet, 'md', { mdConfig: { dialect } } as any)).value as string;
        assert.ok(['merged', '# a comment line', '>c<', '>d<'].every(part => written.includes(part)), `MD: a sheet with merged cells keeps its content (${dialect}: ${written})`);
    }
    const commentSheet = { ...mergedSheet, content: [{ ...mergedSheet.content[0], children: [mergedSheet.content[0].children[2], mergedSheet.content[0].children[1]] }] };
    assert.ok(((await OfficeGenerator.generate(commentSheet, 'md')).value as string).includes('| # a comment line |'), 'MD: a CSV comment line is a row of its table');
    const rtfBack = async (content: any[]) => {
        const rtf = (await OfficeGenerator.generate({ type: 'md', metadata: {}, attachments: [], content } as any, 'rtf', { onWarning: () => {} } as any)).value as string;
        return (await OfficeParser.parseOffice(Buffer.from(rtf), { fileType: 'rtf' } as any)).content.map(n => plain(n.children ?? [])).filter(Boolean);
    };
    assert.deepStrictEqual(await rtfBack([
        { type: 'code', text: 'let x = 1;\nx++;', metadata: { language: 'js' } }, { type: 'code', text: 'x^2', metadata: { math: 'block' } },
        { type: 'paragraph', children: [T('math '), { type: 'code', text: 'y_1', metadata: { math: 'inline' } }, T(' inline')] },
        { type: 'definitionList', children: [{ type: 'definitionTerm', children: [T('Term')] }, { type: 'definitionDescription', children: [T('Description')] }] },
        { type: 'comment', text: '# a comment line' }, { type: 'comment', text: 'hidden', metadata: { sourceSyntax: 'html' } },
    ]), ['let x = 1;\nx++;', 'x^2', 'math |y_1| inline', 'Term', 'Description', '# a comment line'], 'RTF: code, equations, definitions and a comment line are written');
    // Content a writer dropped: an embed naming no type (built by hand) or a YouTube one by its URL alone
    // lost its URL in HTML and Markdown; a table of one row reached no chunk; a picture by URL was left
    // out of RTF; the native PDF engine drew no picture that sits in a paragraph.
    const handBuilt = (content: any[]) => ({ type: 'docx', metadata: {}, attachments: [], content } as any);
    const embeds = handBuilt([{ type: 'embed', metadata: { url: 'https://example.com/page' } }, { type: 'embed', metadata: { embedType: 'youtube', url: 'https://youtu.be/abcdefghijk' } }]);
    for (const format of ['html', 'md'] as const) {
        const written = (await OfficeGenerator.generate(embeds, format, { htmlConfig: { standalone: false } } as any)).value as string;
        assert.ok(written.includes('https://example.com/page') && written.includes('abcdefghijk'), `${format}: embeds keep their URL (${written})`);
    }
    const oneRow = (await OfficeGenerator.generate(handBuilt([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [T('only row')] }] }] }]), 'chunks')).value as any[];
    assert.ok(oneRow.some(chunk => chunk.text.includes('only row')), 'chunks: a table of one row is chunked');
    const urlPicture = (await OfficeGenerator.generate(handBuilt([{ type: 'paragraph', children: [{ type: 'image', metadata: { url: 'https://example.com/p.png', altText: 'the picture' } }] }]), 'rtf', { onWarning: () => {} } as any)).value as string;
    assert.ok(urlPicture.includes('HYPERLINK "https://example.com/p.png"') && urlPicture.includes('the picture'), `RTF: a picture by URL is a link on its alt text (${urlPicture.slice(-200)})`);
    for (const file of ['test.docx', 'test.md']) {
        const ast = await OfficeParser.parseOffice(path.join(__dirname, 'files', file), { extractAttachments: true } as any);
        const pdf = (await OfficeGenerator.generate(ast, 'pdf', { pdfConfig: { engine: 'native' }, onWarning: () => {} } as any)).value as Buffer;
        const back = await OfficeParser.parseOffice(Buffer.from(pdf), { extractAttachments: true } as any);
        assert.ok(back.attachments.length > 0, `PDF (native): the picture in a paragraph of ${file} is drawn`);
    }

    // Findings of the pre-release review. Each saves the same output from the first save on.
    const mdOf = (src: string) => OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md', onWarning: () => {} } as any);
    const htmlOf = (src: string) => OfficeParser.parseOffice(Buffer.from(src), { fileType: 'html', onWarning: () => {} } as any);
    const plainOf = (nodes: any[]): string => nodes.map((n: any) => n.text ?? plainOf(n.children ?? [])).join('|');
    const idsIn = (nodes: any[]): any[] => nodes.flatMap((n: any) => [...(n.metadata?.anchorIds ? [[n.type, n.metadata.anchorIds]] : []), ...idsIn(n.children ?? [])]);
    const settles = async (format: 'md' | 'html', ast: any, label: string) => {
        const outs: string[] = [];
        for (let save = 0; save < 3; save++) {
            outs.push((await OfficeGenerator.generate(ast, format, { generateIds: false, htmlConfig: { standalone: false }, onWarning: () => {} } as any)).value as string);
            ast = await OfficeParser.parseOffice(Buffer.from(outs[save]), { fileType: format, onWarning: () => {} } as any);
        }
        assert.ok(outs[0] === outs[1] && outs[1] === outs[2], `${format}: ${label} saves the same output (${JSON.stringify(outs)})`);
        return { out: outs[0], ast };
    };
    const docOf = (content: any[], type = 'docx') => ({ type, metadata: {}, attachments: [], content } as any);
    // A document starting with a rule (or a page, slide or sheet boundary) is not front matter: the
    // text to the next rule was lost.
    const pages = await settles('md', docOf([{ type: 'page', children: [{ type: 'paragraph', children: [T('Page one text')] }] }, { type: 'page', children: [{ type: 'paragraph', children: [T('Page two text')] }] }]), 'pages');
    assert.ok(plainOf(pages.ast.content).includes('Page one text'), 'MD: a leading rule is not front matter');
    assert.strictEqual((await mdOf('---\ntitle: T\n---\n\nbody')).metadata.title, 'T', 'MD: front matter is still read');
    // A line break in an admonition's or a footnote's own text is written (the words ran together).
    for (const [src, ft, expected] of [['<div class="admonition warning">line one<br>line two</div>', 'html', '> [!WARNING]\n> line one  \n> line two'], ['Text[^1] more.\n\n[^1]: line one<br>line two', 'md', 'Text[^1] more.\n\n[^1]: line one  \n    line two']] as const) {
        const { out } = await settles('md', await OfficeParser.parseOffice(Buffer.from(src), { fileType: ft } as any), src);
        assert.strictEqual(out, expected, `MD: line breaks of ${src}`);
    }
    // An HTML table in a block-level wrapper on its line is a table; captions holding a line break or
    // inline math stay their own paragraph.
    for (const src of ['<center><table><tr><td>A1</td></tr></table></center>', '<figure><table><tr><td>A1</td></tr></table></figure>']) {
        assert.ok((await mdOf(src)).content.some(n => n.type === 'table'), `MD: a table in ${src}`);
    }
    for (const src of ['<figure><img src="a.png" alt="cat"><figcaption>a<br>b</figcaption></figure>', '<figure><img src="a.png" alt="cat"><figcaption>Area <span class="math math-inline" data-math="inline">x^2</span></figcaption></figure>']) {
        assert.deepStrictEqual((await htmlOf(src)).content.map(n => n.type), ['image', 'paragraph'], `HTML: the caption of ${src} is a paragraph`);
    }
    // Headings in any script keep their ids, and links to them their targets.
    const scripts = await settles('md', await mdOf('## Überblick\n\n## Введение\n\n[a](#überblick) [b](#введение)'), 'non-ASCII ids');
    assert.ok(scripts.out.includes('(#überblick)') && scripts.out.includes('(#введение)'), `MD: non-ASCII link targets (${scripts.out})`);
    const scriptHtml = (await OfficeGenerator.generate(await mdOf('## Введение\n\n[b](#введение)'), 'html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(scriptHtml.includes('id="введение"') && scriptHtml.includes('href="#введение"'), `HTML: a non-ASCII heading id (${scriptHtml})`);
    const scriptTex = (await OfficeGenerator.generate(await mdOf('## Введение\n\n## Обзор\n\n[b](#введение)'), 'tex', { onWarning: () => {} } as any)).value as string;
    const labels = [...scriptTex.matchAll(/\\label\{([^}]*)\}/g)].map(m => m[1]);
    assert.ok(labels.length === 2 && labels[0] !== labels[1] && scriptTex.includes(`\\hyperref[${labels[0]}]`), `LaTeX: distinct labels for non-ASCII headings (${labels})`);
    // A literal reference inside emphasis keeps its level of escaping.
    await stable(md, '**&amp;quot;** *&amp;amp;* [&amp;amp;x](u)', '**&amp;quot;** *&amp;amp;* [&amp;amp;x](u)', 'references inside emphasis and link text');
    // A picture among the blocks with two ids, bookmarks in a footnote and an admonition, and the ids in
    // a table written as HTML are kept, and saved the same way each time.
    await settles('html', await htmlOf('<p>Intro</p><a id="fig1"></a><img id="pic" src="a.png" alt="A"><p>End</p>'), 'a picture with two ids');
    const noteBookmark = await settles('md', await mdOf('Text[^1] and more.\n\n[^1]: <a id="fn-target"></a>Note body\n\nSee [the note](#fn-target).'), 'a bookmark in a footnote');
    assert.ok(noteBookmark.out.includes('[^1]: <a id="fn-target"></a>Note body'), `MD: a footnote's bookmark (${noteBookmark.out})`);
    await settles('html', await mdOf('Text[^1] and more.\n\n[^1]: <a id="fn-target"></a>Note body'), 'a bookmark in a footnote');
    assert.deepStrictEqual(idsIn((await htmlOf('<div class="admonition note">text <a id="x"></a>more</div>')).content), [['admonition', ['x']]], 'HTML: a bookmark in an admonition\'s text');
    const htmlTableIds = (await OfficeGenerator.generate(docOf([{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', metadata: { colSpan: 2, anchorIds: ['c1'] }, children: [{ type: 'paragraph', metadata: { anchorIds: ['p1'] }, children: [T('merged')] }] }] }, { type: 'row', children: [{ type: 'cell', children: [{ type: 'image', metadata: { url: 'a.png', altText: 'A', anchorIds: ['im'] } }] }, { type: 'cell', children: [T('b')] }] }] }]), 'md', { generateIds: false } as any)).value as string;
    assert.deepStrictEqual(idsIn((await mdOf(htmlTableIds)).content), [['cell', ['c1']], ['paragraph', ['p1']], ['image', ['im']]], `MD: ids in a table written as HTML (${htmlTableIds})`);
    // An admonition, definition list or embed in a list item or cell is written as that one line holds it.
    for (const src of ['<table><tr><th>H</th></tr><tr><td><div class="admonition note"><p>careful</p></div></td></tr></table>', '<ul><li>Item<dl><dt>T</dt><dd>D</dd></dl></li></ul>', '<ul><li>Item <div data-youtube-video="abcdefghijk"></div></li></ul>']) {
        await settles('md', await htmlOf(src), src);
    }
    // RTF: a picture among the blocks is a paragraph of its own; plain text: a definition's term and
    // description are lines of their own.
    const blockPicture = (await OfficeGenerator.generate(docOf([{ type: 'image', metadata: { url: 'https://example.com/b.png' } }, { type: 'paragraph', children: [T('next')] }]), 'rtf', { onWarning: () => {} } as any)).value as string;
    assert.deepStrictEqual((await OfficeParser.parseOffice(Buffer.from(blockPicture), { fileType: 'rtf' } as any)).content.map(n => plainOf(n.children ?? [])), ['https://example.com/b.png', 'next'], 'RTF: a picture among the blocks is its own paragraph');
    assert.strictEqual((await (await mdOf('Term Alpha\n: Description\n\nAfter')).to('text')).value, 'Term Alpha\nDescription\nAfter', 'Text: definition lines');
    // Content indented to a list item's content column is the item's (a table, a paragraph), not code;
    // Markdown in an HTML table's cells between blank lines is read into the cell.
    assert.deepStrictEqual((await mdOf('1. Step one\n\n    <table><tr><td>A</td></tr></table>\n\n2. Step two')).content.map(n => n.type), ['list', 'table', 'list'], 'MD: a table under a list item');
    assert.deepStrictEqual((await mdOf('- item\n\n    more of the item\n\n      code under it')).content.map(n => n.type), ['list', 'paragraph', 'code'], 'MD: a paragraph and code under a list item');
    const mdCell = await mdOf('<table>\n<tr>\n<td>\n\n**bold** cell\n\n</td>\n</tr>\n</table>\n\nAfter');
    // (A paragraph of the cell's, as Markdown between blank lines is.)
    assert.deepStrictEqual([mdCell.content.map(n => n.type), mdCell.content[0].children![0].children![0].children!.map(n => n.type), mdCell.content[0].children![0].children![0].children![0].children!.map(n => [n.text, n.formatting?.bold ?? false])], [['table', 'paragraph'], ['paragraph'], [['bold', true], [' cell', false]]], 'MD: Markdown in an HTML table cell');
    // Preformatted text keeps the tokens a highlighter wraps and its line breaks.
    for (const [src, code] of [['<pre>line a<br>line b</pre>', 'line a\nline b'], ['<pre><code><span class="k">let</span> x = <span class="n">1</span>;</code></pre>', 'let x = 1;']] as const) {
        assert.strictEqual((await htmlOf(src)).content[0].text, code, `HTML: ${src}`);
    }
    // A sheet's own id is kept beside the id its tab links to, written once.
    const sheetIds = (await OfficeGenerator.generate(docOf([{ type: 'sheet', metadata: { sheetName: 'S', anchorIds: ['sh'] }, children: [{ type: 'row', children: [{ type: 'cell', metadata: { row: 0, col: 0 }, children: [T('a')] }] }] }], 'xlsx'), 'html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(sheetIds.includes('<a id="sh" name="sh"></a><div id="sheet-0"') && (sheetIds.match(/sheet-0"/g) || []).length === 2, `HTML: a sheet's ids (${sheetIds.slice(0, 200)})`);
    // The ids HtmlGenerator writes on equations and on the wrapper of a picture among the blocks read back.
    const idNodes = [
        { type: 'paragraph', children: [T('p '), { type: 'code', text: 'x', metadata: { math: 'inline', anchorIds: ['mi'] } }] },
        { type: 'code', text: 'y^2', metadata: { math: 'block', anchorIds: ['mb'] } },
        { type: 'image', metadata: { url: 'https://x.test/b.png', altText: 'B', anchorIds: ['ib'] } },
    ];
    const idsHtml = (await OfficeGenerator.generate({ type: 'docx', metadata: {}, attachments: [], content: idNodes } as any, 'html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.deepStrictEqual(idsOf((await htmlAst(idsHtml)).content), [['code', ['mi']], ['code', ['mb']], ['image', ['ib']]], `HTML: equation and picture ids read back (${idsHtml})`);
    // A heading's generated id comes from its text, whether the node or only its runs carry it, and a
    // heading whose text slugifies to nothing gets none rather than id="".
    assert.ok((await htmlOnce('<h2>My Title</h2>')).includes('<h2 id="my-title">'), 'HTML: a heading read from HTML gets its generated id');
    assert.ok(!(await htmlOnce('<h2>标题</h2>')).includes('id=""'), 'HTML: no empty id');

    // Markdown blocks as CommonMark reads them: an HTML table a blank line runs through is one table, an
    // indented code block runs on across blank lines, and a `=`/`-` line after a quote is not a heading's
    // underline; `<table>` in a code span is text.
    const mdAst = (src: string) => OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any);
    const splitTable = await mdAst('<table>\n<tr><td>a</td></tr>\n   \n<tr><td>b</td></tr>\n</table>\n\nafter');
    assert.deepStrictEqual([splitTable.content.map(n => n.type), splitTable.content[0].children!.length], [['table', 'paragraph'], 2], 'MD: an HTML table with a blank line in it is one table');
    for (const [src, code] of [['    a\n    \n    b', 'a\n\nb'], ['    a\n\n\n    b', 'a\n\n\nb'], ['para\n\n    c1\n\n    c2', 'c1\n\nc2']] as const) {
        const blocks = (await mdAst(src)).content.filter(n => n.type === 'code');
        assert.deepStrictEqual(blocks.map(n => n.text), [code], `MD: indented code ${JSON.stringify(src)}`);
    }
    for (const [src, types] of [['> quote\n===', ['paragraph']], ['   > quote\n---', ['paragraph', 'break']], ['> quote\nlazy\n-', ['paragraph']], ['use `<table>` here', ['paragraph']]] as const) {
        const parsed = await mdAst(src);
        assert.deepStrictEqual(parsed.content.map(n => n.type), types, `MD: blocks of ${JSON.stringify(src)}`);
    }
    assert.strictEqual(plain((await mdAst('use `<table>` here')).content), 'use |<table>| here', 'MD: a tag in a code span is its text');
    // Anchors: an empty `<a id>` in a line is the id of the picture after it, else of the line's block,
    // one naming no id is text, and a bookmark before a code block, rule, admonition or definition list
    // is that block's.
    for (const [src, ids] of [
        ['text <a id="m"></a> more', [['paragraph', ['m']]]],
        ['text <a id="pic"></a>![A](a.png) more', [['image', ['pic']]]],
        ['a <a href="x"></a> b', []],
        ['| h |\n| --- |\n| <a id="c"></a>cell |', [['cell', ['c']]]],
        ['<a id="a"></a>\n\n> [!NOTE]\n> n\n\nafter', [['admonition', ['a']]]],
        ['<a id="a"></a>\n\n```\nc\n```', [['code', ['a']]]],
        ['<a id="a"></a>\n\n---\n\nafter', [['break', ['a']]]],
        ['<a id="a"></a>\n\nTerm\n: <a id="d"></a>Desc', [['definitionList', ['a']], ['definitionDescription', ['d']]]],
        ['<a id="t"></a>Term\n: Desc\n<a id="t2"></a>Term 2\n: Desc 2', [['definitionTerm', ['t']], ['definitionTerm', ['t2']]]],
    ] as const) {
        const parsed = await mdAst(src);
        assert.deepStrictEqual(idsOf(parsed.content), ids, `MD: the ids of ${JSON.stringify(src)}`);
        const written = ((await parsed.to('md', { generateIds: false } as any)).value as string).trim();
        await stable(md, written, written, `the ids of ${JSON.stringify(src)} over saves (${written})`);
        assert.deepStrictEqual(idsOf((await mdAst(written)).content), ids, `MD: the ids of ${JSON.stringify(src)} read back`);
    }
    assert.strictEqual(plain((await mdAst('a <a href="x"></a> b')).content), 'a <a href="x"></a> b', 'MD: an empty anchor naming no id is text');
    // One line (an item, a cell, a definition) holds no heading or quote: their text is written. A line
    // break among blocks, before a rule, or an id that slugifies to nothing writes nothing.
    const cr = { type: 'break', metadata: { breakType: 'carriageReturn' } };
    for (const [label, content, expected] of [
        ['a heading in a definition', (await htmlAst('<dl><dt>t</dt><dd><h1>h</h1></dd></dl>')).content, 't\n: h'],
        ['a quote in a list item', (await htmlAst('<ul><li>a<blockquote>q</blockquote></li></ul>')).content, '- a<br>q'],
        ['a line break among blocks', [{ type: 'paragraph', children: [T('a')] }, cr, { type: 'paragraph', children: [T('b')] }], 'a\n\nb'],
        ['a line break before a rule', [{ type: 'paragraph', children: [T('a'), cr, { type: 'break', metadata: { breakType: 'thematic' } }, T('b')] }], 'a\n\n---\n\nb'],
        ['an id that slugifies to nothing', [{ type: 'paragraph', metadata: { anchorIds: ['!!!'] }, children: [T('x')] }], 'x'],
        // Ids of what a line holds are its container's, written where a reader gives them.
        ['ids in a list item', [{ type: 'list', metadata: { listType: 'unordered', listId: 'l', indentation: 0, itemIndex: 0 }, children: [T('item '), { type: 'code', text: 'c', metadata: { anchorIds: ['ic'] } }] }], '- <a id="ic"></a>item<br>`c`'],
        ['ids in a cell', [{ type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [T('p1')] }, { type: 'paragraph', metadata: { anchorIds: ['p2'] }, children: [T('p2')] }] }] }] }], '| <a id="p2"></a>p1<br>p2 |\n| --- |'],
        ['ids of inline math', [{ type: 'paragraph', children: [T('see '), { type: 'code', text: 'x', metadata: { math: 'inline', anchorIds: ['m'] } }] }], '<a id="m"></a>see $x$'],
        ['spaces before a block in a paragraph', [{ type: 'paragraph', children: [T('a '), { type: 'code', text: 'blk' }, T(' b')] }], 'a\n\n```\nblk\n```\n\nb'],
    ] as const) {
        const written = await writeDoc(content as any);
        assert.strictEqual(written, expected, `MD: ${label}`);
        await stable(md, written, written, `${label} over saves`);
    }

    // HTML decodes every numeric reference and HTML 4's names, in text and in attribute values.
    const html = await OfficeParser.parseOffice(Buffer.from('<p>it&rsquo;s &copy; &#8217; &#x2019; &eacute; &amp;quot; <a href="http://x.com/?a=1&amp;b=2" title="t &amp; u">l</a> <img src="i.png" alt="Tom &amp; Jerry" title="q &quot;x&quot;"></p>'), { fileType: 'html' } as any);
    // (The space between the link and the image is a node of its own, as in HTML.)
    const [htmlText, htmlLink, htmlImage] = html.content[0].children!.filter(c => c.type !== 'text' || c.text !== ' ');
    assert.deepStrictEqual(
        [htmlText.text, (htmlLink.metadata as any).link, (htmlLink.metadata as any).title, (htmlImage.metadata as any).altText, (htmlImage.metadata as any).title],
        ['it’s © ’ ’ é &quot; ', 'http://x.com/?a=1&b=2', 't & u', 'Tom & Jerry', 'q "x"'],
        'HTML: character references decode in text, href, title and alt');
    console.log('  Markdown round trips: All assertions passed ✓');
}

/**
 * Findings of the second pre-release review: Markdown's HTML tables, escapes, lists, link definitions,
 * autolinks, tabs and front matter; HTML's numeric references, preformatted text, heading ids and
 * whitespace; nested notes in every writer; plain text's blocks in list items; grids of far cells.
 */
async function testSecondReview(): Promise<void> {
    const parse = (src: string, fileType: string, extra: object = {}) => OfficeParser.parseOffice(Buffer.from(src), { fileType, onWarning: () => {}, ...extra } as any);
    const write = async (ast: any, format: string, extra: object = {}) => (await OfficeGenerator.generate(ast, format as any, { generateIds: false, htmlConfig: { standalone: false }, onWarning: () => {}, ...extra } as any)).value as string;
    const plain = (nodes: any[]): string => nodes.map((n: any) => n.text ?? plain(n.children ?? [])).join('|');
    const settles = async (format: string, ast: any, label: string, extra: object = {}) => {
        const outs: string[] = [];
        for (let save = 0; save < 3; save++) {
            outs.push(await write(ast, format, extra));
            ast = await parse(outs[save], format, extra);
        }
        assert.ok(outs[0] === outs[1] && outs[1] === outs[2], `${format}: ${label} saves the same output (${JSON.stringify(outs)})`);
        return { out: outs[0], ast };
    };
    const T = (text: string, extra: object = {}) => ({ type: 'text', text, ...extra });
    const doc = (content: any[], type = 'docx') => ({ type, metadata: {}, attachments: [], content } as any);

    // Markdown in an HTML table: a cell's preformatted text or picture keeps its blank lines and leaves
    // no placeholder behind; a cell's two paragraphs are two lines of it; inline code naming a tag is
    // not a table.
    for (const src of ['<table><tr><td><pre>\n\nline1\n\n</pre></td></tr></table>\n', '<table><tr><td><img alt="\n\nfoo bar\n\n" src="a.png"></td></tr></table>\n']) {
        const ast = await parse(src, 'md');
        assert.ok(!/[\uE000\uE001]/.test(JSON.stringify(ast.content)), `MD: no placeholder is left of ${JSON.stringify(src)}`);
    }
    const twoParagraphs = await settles('md', await parse('<table>\n<tr>\n<td>\n\nFirst paragraph.\n\nSecond paragraph.\n\n</td>\n</tr>\n</table>\n', 'md'), 'a cell of two paragraphs');
    assert.strictEqual(twoParagraphs.out, '| First paragraph.<br>Second paragraph. |\n| --- |', 'MD: a cell of two paragraphs');
    const codeTag = await parse('Use `<table>` for data.\n<p>Next</p>\n', 'md');
    assert.ok(!codeTag.content.some(n => n.type === 'table') && plain(codeTag.content).includes('<table>'), 'MD: inline code naming <table> is not a table');

    // Escapes: a bold run ending in a backslash, and a code span holding an escaped pipe in a cell.
    const backslash = await settles('md', doc([{ type: 'paragraph', children: [T('Save to '), T('C:\\ ', { formatting: { bold: true } }), T('and continue.')] }]), 'a bold run ending in a backslash');
    assert.strictEqual(backslash.out, 'Save to **C:\\\\** and continue.', 'MD: a bold run ending in a backslash');
    const pipe = await settles('md', await parse('<table><tr><th>Pattern</th><th>Meaning</th></tr><tr><td><code>a\\|b</code></td><td>alternation</td></tr></table>', 'html'), 'an escaped pipe in a cell');
    assert.ok(JSON.stringify(pipe.ast.content).includes('"a\\\\|b"'), `MD: the escaped pipe reads back (${pipe.out})`);
    // A row of dashes after the delimiter row is a row.
    const dashes = await parse('| Name | Value |\n| --- | --- |\n| - | - |\n| x | 1 |\n', 'md');
    assert.deepStrictEqual(dashes.content[0].children!.map(r => plain(r.children!)), ['Name|Value', '-|-', 'x|1'], 'MD: a row of dashes is a row');

    // Lists: a change of marker kind starts a new list, as HTML's two lists stay two.
    assert.strictEqual(await write(await parse('- a\n- b\n1. one\n2. two\n', 'md'), 'md'), '- a\n- b\n1. one\n2. two', 'MD: bullets then numbers');
    const twoLists = await settles('md', await parse('<ul><li>a</li><li>b</li></ul><ol><li>one</li><li>two</li></ol>', 'html'), 'two lists');
    assert.strictEqual(new Set(twoLists.ast.content.map((n: any) => n.metadata?.listId)).size, 2, 'MD: two lists keep two list ids');

    // Link definitions: brackets in link text are escaped; an escaped `]` ends no label; a definition
    // does not interrupt a paragraph or an item.
    const brackets = await settles('md', doc([{ type: 'paragraph', children: [T('Step [2]: configure', { metadata: { link: 'https://e.com/step2', linkType: 'external' } })] }]), 'brackets in link text');
    assert.strictEqual(brackets.out, '[Step \\[2\\]: configure](https://e.com/step2)', 'MD: brackets in link text');
    for (const src of ['[a\\]: b\n', '- item\n[x]: y\n', 'para\n[x]: y\n']) {
        assert.ok(plain((await parse(src, 'md')).content).includes(': '), `MD: ${JSON.stringify(src)} is not a definition`);
    }

    // Autolinks: references decoded in text and target; an address is a mail link.
    const autolink = await settles('md', await parse('<https://e.com/?a=1&amp;b=2>\n', 'md'), 'an autolink with a reference');
    assert.strictEqual(autolink.out, '[https://e.com/?a=1&b=2](https://e.com/?a=1&b=2)', 'MD: an autolink with a reference');
    const mail = await parse('mail <me@x.com> now', 'md');
    assert.deepStrictEqual(mail.content[0].children!.map(n => [n.text, (n.metadata as any)?.link]), [['mail ', undefined], ['me@x.com', 'mailto:me@x.com'], [' now', undefined]], 'MD: an address in angle brackets is a mail link');
    await settles('md', mail, 'a mail link');

    // A tab after a list marker, and a tab-indented paragraph under the item, are the item's.
    assert.deepStrictEqual((await parse('-\ttab item\n\n\tcontinued\n', 'md')).content.map(n => n.type), ['list', 'paragraph'], 'MD: a tab-indented paragraph under an item is not code');
    // Indented code ends at its last line of code.
    assert.deepStrictEqual((await parse('para\n\n    code\n', 'md')).content.map(n => n.text ?? n.type), ['paragraph', 'code'], 'MD: indented code');
    assert.strictEqual((await parse('    a\n\n    b\n\n\n', 'md')).content[0].text, 'a\n\nb', 'MD: indented code keeps its inner blank lines only');
    // Front matter is YAML: a `---` block of prose is a rule, a heading and text.
    const notYaml = await parse('---\nfoo\n---\n\nbar\n', 'md');
    assert.ok(plain(notYaml.content).includes('foo') && plain(notYaml.content).includes('bar') && !notYaml.metadata.customProperties, 'MD: a --- block of prose is not front matter');
    assert.strictEqual((await parse('---\ntitle: "T"\ntags: [a, b]\nlist:\n  - x\n# comment\n---\n\nbody\n', 'md')).metadata.title, 'T', 'MD: YAML front matter is read');
    // A CSV comment row survives a save to Markdown.
    await settles('md', await parse('# note\na,b\n1,2\n', 'csv'), 'a CSV comment row');

    // Raw HTML in Markdown is read as a renderer shows it, where its tags were escaped into view on
    // the first save: an HTML block by the HTML parser, an inline element as its formatting.
    for (const [src, expected] of [
        ['<p align="center">\n  <img src="logo.png" width="200" alt="Logo">\n</p>\n', '<div style="text-align: center">![Logo](logo.png){width=200}</div>'],
        ['<h1 align="center">Title</h1>\n', '<div style="text-align: center">\n\n# Title\n\n</div>'],
        ['<details>\n<summary>More</summary>\n\nHidden text\n\n</details>\n', 'More\n\nHidden text'],
        ['Intro line\n<p>Para</p>\n', 'Intro line\n\nPara'],
        ['Press <kbd>Ctrl</kbd>+<kbd>C</kbd>, <b>bold</b>, <i>it</i>, <code>x &lt; y</code>, <sup class="n">2</sup>.', 'Press `Ctrl`+`C`, **bold**, *it*, `x < y`, <sup>2</sup>.'],
        ['See <a href="https://e.com">the site</a> and <img src="a.png" alt="A"> here.', 'See [the site](https://e.com) and ![A](a.png) here.'],
        ['Unknown <foo>tag</foo>, unclosed <b>bold, a lone </i>.', 'Unknown &lt;foo>tag&lt;/foo>, unclosed &lt;b>bold, a lone &lt;/i>.'],
        ['Code `<b>x</b>` stays code.', 'Code `<b>x</b>` stays code.'],
    ] as const) {
        const { out } = await settles('md', await parse(src, 'md'), `raw HTML ${JSON.stringify(src)}`);
        assert.strictEqual(out, expected, `MD: raw HTML ${JSON.stringify(src)}`);
    }
    assert.strictEqual(JSON.stringify((await parse('<abbr title="HyperText">HTML</abbr>', 'md')).content[0].children![0].metadata), '{"abbreviationTitle":"HyperText"}', 'MD: <abbr title> is an abbreviation');
    // HTML: numeric references in 0x80-0x9F are Windows-1252's characters, as browsers read them.
    assert.strictEqual(plain((await parse('<p>It&#146;s &#147;q&#148; &#150; x</p>', 'html')).content), 'It’s “q” – x', 'HTML: Windows-1252 references');
    // Preformatted text is all of the block's text, less the line break after <pre>.
    for (const [src, text] of [['<pre><code>line1</code>\n<code>line2</code></pre>', 'line1\nline2'], ['<pre><span>$ </span><code>npm i</code></pre>', '$ npm i'], ['<pre>\nx</pre>', 'x'], ['<pre><code>x\n</code></pre>', 'x\n']] as const) {
        assert.strictEqual((await parse(src, 'html')).content[0].text, text, `HTML: ${src}`);
    }
    // A heading's id is GitHub's, so a link written for the Markdown heading reaches it.
    assert.ok((await write(await parse('## Version 2.0\n\n[v](#version-20)', 'md'), 'html', { generateIds: true })).includes('<h2 id="version-20">'), 'HTML: a heading id matches GitHub\'s');
    // Whitespace kept as written: whitespace between blocks is no paragraph, and the space before a
    // footnote's back-link is not the note's.
    const kept = await settles('html', await parse('Text[^1] and more.\n\n[^1]: Note body', 'md'), 'a footnote, whitespace kept', { preserveXmlWhitespace: true });
    assert.deepStrictEqual([kept.ast.content.length, plain(kept.ast.content[0].children!.flatMap((c: any) => c.notes ?? []))], [1, 'Note body'], 'HTML: whitespace kept adds no paragraph and no space to a note');
    // A figure caption holding a block keeps both.
    await settles('html', await parse('<figure><img src="a.png" alt="A"><figcaption>Text <code>x</code> and <pre>block</pre> after</figcaption></figure>', 'html'), 'a caption holding a block');

    // A note in a note's text is written by every writer (LaTeX as a mark with its text after the note).
    const nested = await parse('Text[^1].\n\n[^1]: First, see[^2].\n\n[^2]: Second NESTEDMARK.\n', 'md');
    for (const format of ['tex', 'html', 'md', 'text', 'rtf']) {
        assert.ok((await write(nested, format)).includes('NESTEDMARK'), `${format}: a note in a note's text is written`);
    }
    assert.ok(/\\footnote\{First, see\\footnotemark\{\}\.\}\\addtocounter\{footnote\}\{-1\}\\stepcounter\{footnote\}\\footnotetext\{Second NESTEDMARK\.\}/.test(await write(nested, 'tex')), 'TEX: a nested note is a mark, its text after the note');

    // Plain text: a block in a list item starts a line of its own.
    assert.strictEqual(await write(await parse('<ul><li>item<dl><dt>T</dt><dd>D</dd></dl></li></ul>', 'html'), 'text'), '- item\nT\nD', 'Text: a definition list in an item');

    // A cell two million rows from the others lays out close, in no time.
    const far = doc([{ type: 'sheet', metadata: { sheetName: 'S' }, children: [{ type: 'row', children: [{ type: 'cell', metadata: { row: 1999999, col: 0 }, children: [T('far')] }] }, { type: 'row', children: Array.from({ length: 40 }, (_, i) => ({ type: 'cell', children: [T('c' + i)] })) }] }], 'xlsx');
    const started = Date.now();
    const farHtml = await write(far, 'html');
    assert.ok(Date.now() - started < 3000 && farHtml.length < 100_000 && farHtml.includes('far') && farHtml.includes('c39'), `HTML: a far cell (${Date.now() - started}ms, ${farHtml.length} chars)`);
    console.log('  Second review: All assertions passed ✓');
}

/**
 * An image that is a link (a badge, a clickable picture) keeps its target through every format that has
 * links: Markdown `[![alt](src)](target "title")`, HTML `<a href><img></a>`, and linked pictures in Word
 * (a hyperlink around the picture, or a link on the picture itself), PowerPoint, ODF, RTF and LaTeX.
 */
async function testImageLinks(): Promise<void> {
    const findImages = (nodes: OfficeContentNode[]): OfficeContentNode[] => nodes.flatMap(n => [...(n.type === 'image' ? [n] : []), ...findImages(n.children || [])]);
    const links = (nodes: OfficeContentNode[]) => findImages(nodes).map(n => { const m = n.metadata as any; return [m?.link, m?.linkType, m?.linkTitle]; });
    const md = async (src: string) => ((await (await OfficeParser.parseOffice(Buffer.from(src), { fileType: 'md' } as any)).to('md', { generateIds: false } as any)).value as string).trim();

    // Markdown: the badge is an image carrying the link and its title, written back as it was, also through HTML.
    const badge = '[![Build](https://img.example/b.svg "badge")](https://ci.example/run "CI") and [![Logo](logo.png)](#intro)';
    const parsed = await OfficeParser.parseOffice(Buffer.from(badge), { fileType: 'md' } as any);
    assert.deepStrictEqual(links(parsed.content), [['https://ci.example/run', 'external', 'CI'], ['#intro', 'internal', undefined]], 'Image links: a Markdown badge carries its link');
    assert.strictEqual(await md(badge), badge, 'Image links: a Markdown badge is written back as it was');
    const html = (await parsed.to('html', { htmlConfig: { standalone: false } } as any)).value as string;
    assert.ok(html.includes('<a href="https://ci.example/run" title="CI" target="_blank"><img src="https://img.example/b.svg" alt="Build" title="badge"></a>'), `Image links: HTML wraps the picture in its link (${html})`);
    const fromHtml = await OfficeParser.parseOffice(Buffer.from(html), { fileType: 'html' } as any);
    assert.deepStrictEqual(links(fromHtml.content), [['https://ci.example/run', 'external', 'CI'], ['#intro', 'internal', undefined]], 'Image links: HTML reads a linked picture back');
    const ignored = (await parsed.to('md', { generateIds: false, ignoreInternalLinks: true } as any)).value as string;
    assert.ok(ignored.includes('![Logo](logo.png)') && !ignored.includes('](#intro)'), `Image links: ignoreInternalLinks drops an internal image link (${ignored})`);

    // The office formats write a linked picture and read it back, external and internal targets alike.
    const png = fs.readFileSync(path.join(__dirname, '..', 'docs', 'favicon.png'));
    const office: OfficeParserAST = { type: 'md', metadata: {}, attachments: [{ type: 'image', name: 'p.png', mimeType: 'image/png', extension: 'png', data: png.toString('base64') }], content: [
        { type: 'paragraph', children: [{ type: 'text', text: 'See ' }, { type: 'image', metadata: { attachmentName: 'p.png', altText: 'pic', link: 'https://example.com/x?a=1&b=2', linkType: 'external' } as ImageMetadata }] },
        { type: 'heading', metadata: { level: 1, anchorIds: ['target'] } as any, children: [{ type: 'text', text: 'Target' }] },
        { type: 'paragraph', children: [{ type: 'image', metadata: { attachmentName: 'p.png', altText: 'internal', link: '#target', linkType: 'internal' } as ImageMetadata }] },
    ] } as any;
    for (const format of ['docx', 'odt', 'rtf', 'tex'] as const) {
        const out = (await OfficeGenerator.generate(office, format, {} as any)).value;
        const back = await OfficeParser.parseOffice(Buffer.from(out as any), { fileType: format, extractAttachments: true, onWarning: () => { } } as any);
        assert.deepStrictEqual(links(back.content).map(l => l.slice(0, 2)), [['https://example.com/x?a=1&b=2', 'external'], ['#target', 'internal']], `Image links: ${format} writes and reads back a linked picture`);
    }

    // Links a generator does not write, from other producers: a Word picture's own link (on docPr), a
    // PowerPoint picture's click link, and an ODF frame inside draw:a (which used to lose the picture too).
    const docx = Buffer.from(zipSync({
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="png" ContentType="image/png"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        'word/document.xml': strToU8('<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><w:body><w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p" descr="pic"><a:hlinkClick r:id="rId2"/></wp:docPr><a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>'),
        'word/_rels/document.xml.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/p.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docx" TargetMode="External"/></Relationships>'),
        'word/media/p.png': new Uint8Array(png),
    }));
    const pptx = Buffer.from(zipSync({
        'ppt/presentation.xml': strToU8('<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'),
        'ppt/slides/slide1.xml': strToU8('<?xml version="1.0"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:pic><p:nvPicPr><p:cNvPr id="2" name="p" descr="pic"><a:hlinkClick r:id="rId2"/></p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic></p:spTree></p:cSld></p:sld>'),
        'ppt/slides/_rels/slide1.xml.rels': strToU8('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/p.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/pptx" TargetMode="External"/></Relationships>'),
        'ppt/media/p.png': new Uint8Array(png),
    }));
    const odp = Buffer.from(zipSync({
        mimetype: strToU8('application/vnd.oasis.opendocument.presentation'),
        'content.xml': strToU8('<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page draw:name="p1"><draw:a xlink:type="simple" xlink:href="https://example.com/odp"><draw:frame><draw:image xlink:href="Pictures/p.png"/><svg:title>pic</svg:title></draw:frame></draw:a></draw:page></office:presentation></office:body></office:document-content>'),
        'Pictures/p.png': new Uint8Array(png),
    }));
    const odt = Buffer.from(zipSync({
        mimetype: strToU8('application/vnd.oasis.opendocument.text'),
        'content.xml': strToU8('<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text><text:p>See <draw:a xlink:type="simple" xlink:href="https://example.com/odt"><draw:frame text:anchor-type="as-char"><draw:image xlink:href="Pictures/p.png"/></draw:frame></draw:a></text:p></office:text></office:body></office:document-content>'),
        'Pictures/p.png': new Uint8Array(png),
    }));
    for (const [format, bytes, target] of [['docx', docx, 'https://example.com/docx'], ['pptx', pptx, 'https://example.com/pptx'], ['odp', odp, 'https://example.com/odp'], ['odt', odt, 'https://example.com/odt']] as const) {
        const ast = await OfficeParser.parseOffice(bytes, { fileType: format, extractAttachments: true, onWarning: () => { } } as any);
        assert.deepStrictEqual(links(ast.content).map(l => l.slice(0, 2)), [[target, 'external']], `Image links: a linked picture in ${format} carries its link`);
    }
    console.log('  Image links: All assertions passed ✓');
}

/** Builders and readers of small OOXML packages, for the OOXML review tests below. */
function ooxmlHelpers() {
    const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"';
    const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
    const p = (text: string) => `<w:p>${r(text)}</w:p>`;
    const relsOf = (rels: string) => `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
    const rel = (id: string, type: string, target: string) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
    const part = (root: string, inner: string) => `<?xml version="1.0"?><w:${root} ${W}>${inner}</w:${root}>`;
    const docx = (body: string, parts: Record<string, string | Uint8Array> = {}) => Buffer.from(zipSync({
        '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'),
        'word/document.xml': strToU8(part('document', `<w:body>${body}</w:body>`)),
        ...Object.fromEntries(Object.entries(parts).map(([name, value]) => [name, typeof value === 'string' ? strToU8(value) : value])),
    }));
    const parse = async (bytes: Buffer | Uint8Array, extra: object = {}, fileType = 'docx') => {
        const issues: any[] = [];
        const ast = await OfficeParser.parseOffice(Buffer.from(bytes), { fileType, onWarning: (issue: any) => issues.push(issue), ...extra } as any);
        return { ast, issues };
    };
    const texts = (nodes: OfficeContentNode[]) => nodes.map(n => [n.type, n.text]);
    return { W, r, p, relsOf, rel, part, docx, parse, texts };
}

/**
 * DOCX reading, from the OOXML review before 8.1.0: chunks that cannot be read, chunks in notes, comments
 * and headers (each part's ids its own relationships), text boxes and SmartArt read as blocks of their own,
 * and alternate content read from the branch the reader understands.
 */
async function testDocxReadingReview(): Promise<void> {
    const { W, r, p, relsOf, rel, part, docx, parse, texts } = ooxmlHelpers();
    // A chunk that cannot be read is not read, with a warning, and the document is; the document's
    // budgets still hold for its chunks.
    const chunked = (name: string, content: string | Uint8Array) => docx(p('MAIN-TEXT') + '<w:altChunk r:id="c1"/>' + p('MAIN-END'), { 'word/_rels/document.xml.rels': relsOf(rel('c1', 'aFChunk', name)), [`word/${name}`]: content });
    for (const [label, name, content, code] of [
        ['a DOCX chunk that is not a ZIP', 'chunk.docx', 'not a zip', 'ZIP_NO_ENTRIES_FOUND'],
        ['a DOCX chunk without a document part', 'chunk.docx', zipSync({ 'foo.txt': strToU8('x') }), 'REQUIRED_PART_MISSING'],
        ['an RTF chunk nested 300 deep', 'chunk.rtf', '{\\rtf1 ' + '{'.repeat(300) + 'x' + '}'.repeat(300) + '}', 'MAX_NESTING_DEPTH_EXCEEDED'],
    ] as const) {
        const { ast, issues } = await parse(chunked(name, content));
        assert.deepStrictEqual(ast.content.map(n => n.text), ['MAIN-TEXT', 'MAIN-END'], `DOCX: ${label} leaves the document read`);
        assert.ok(issues.length === 1 && issues[0].code === 'ALT_CHUNK_NOT_READ' && issues[0].message.includes(code), `DOCX: ${label} is one warning naming ${code} (${JSON.stringify(issues.map(i => [i.type, i.code, i.message]))})`);
    }
    let overBudget: any;
    try { await parse(chunked('chunk.txt', 'a\n'.repeat(2000)), { decompressionLimits: { maxXmlElements: 1000 } }); } catch (e) { overBudget = e; }
    assert.strictEqual(overBudget?.officeIssue?.code, 'XML_ELEMENT_LIMIT_EXCEEDED', 'DOCX: a chunk past the element budget still fails the document');

    // Chunks in notes, comments and headers are read, and each part's ids name its own relationships
    // (its links and pictures too); an embedded object is not inflated; an HTML chunk keeps its encoding.
    const parts = await parse(docx(`<w:p>${r('Body')}<w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:commentReference w:id="4"/></w:r></w:p><w:altChunk r:id="c1"/><w:altChunk r:id="c2"/>`, {
        'word/_rels/document.xml.rels': relsOf(rel('rId1', 'hyperlink', 'https://document.example/') + rel('c1', 'aFChunk', 'utf16.htm') + rel('c2', 'aFChunk', 'cp1252.htm')),
        'word/utf16.htm': Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('<p>UTF-16 café</p>', 'utf16le')]),
        'word/cp1252.htm': Buffer.from('<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"></head><body><p>Windows-1252 caf\xe9</p></body></html>', 'latin1'),
        'word/footnotes.xml': part('footnotes', `<w:footnote w:id="1"><w:p>${r('FN ')}<w:hyperlink r:id="rId1">${r('link')}</w:hyperlink></w:p><w:altChunk r:id="c1"/></w:footnote>`),
        'word/_rels/footnotes.xml.rels': relsOf(rel('rId1', 'hyperlink', 'https://footnote.example/') + rel('c1', 'aFChunk', 'footnote.txt')),
        'word/footnote.txt': 'CHUNK IN FOOTNOTE',
        'word/comments.xml': part('comments', `<w:comment w:id="4" w:author="A"><w:p>${r('CM')}</w:p><w:altChunk r:id="c1"/></w:comment>`),
        'word/_rels/comments.xml.rels': relsOf(rel('c1', 'aFChunk', 'comment.txt')),
        'word/comment.txt': 'CHUNK IN COMMENT',
        'word/header1.xml': part('hdr', `<w:p><w:r><w:drawing><wp:inline><wp:docPr id="1" name="p" descr="logo"/><a:graphic><a:graphicData><a:blip r:embed="rId1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p><w:altChunk r:id="c1"/>`),
        'word/_rels/header1.xml.rels': relsOf(rel('rId1', 'image', 'media/logo.png') + rel('c1', 'aFChunk', 'header.txt')),
        'word/header.txt': 'CHUNK IN HEADER',
        'word/media/logo.png': new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    }), { extractAttachments: true });
    const bodyRefs = parts.ast.content[0].children!;
    const footnote = bodyRefs.flatMap(n => n.notes ?? [])[0];
    const comment = bodyRefs.flatMap(n => n.comments ?? [])[0];
    const plain = (node: OfficeContentNode): string => node.text ?? (node.children ?? []).map(plain).join('');
    assert.deepStrictEqual(parts.ast.content.slice(1).map(plain), ['UTF-16 café', 'Windows-1252 café'], 'DOCX: HTML chunks are read in the encoding their byte order mark or <meta> names');
    assert.deepStrictEqual(texts(footnote.children!), [['paragraph', 'FN link'], ['paragraph', 'CHUNK IN FOOTNOTE']], 'DOCX: a chunk in a footnote is read');
    assert.strictEqual((footnote.children![0].children![1].metadata as any)?.link, 'https://footnote.example/', 'DOCX: a footnote\'s link is its part\'s relationship, not the document\'s');
    assert.deepStrictEqual(texts(comment.children!), [['paragraph', 'CM'], ['paragraph', 'CHUNK IN COMMENT']], 'DOCX: a chunk in a comment is read');
    const header = parts.ast.auxiliary!.headers!;
    assert.deepStrictEqual([(header[0].children![0].metadata as any)?.attachmentName, header[1].text], ['logo.png', 'CHUNK IN HEADER'], 'DOCX: a header\'s picture and chunk are read through its own relationships');
    assert.deepStrictEqual(parts.issues, [], 'DOCX: chunks in every part read without a warning');
    const embeddedObject = zipSync({ 'word/document.xml': strToU8('<w:document/>'), 'big.bin': new Uint8Array(3_000_000) }, { level: 0 });
    const embedded = await parse(docx(p('MAIN'), { 'word/embeddings/Microsoft_Word_Document.docx': embeddedObject }), { decompressionLimits: { maxUncompressedBytes: 2_000_000 } });
    assert.deepStrictEqual(embedded.ast.content.map(n => n.text), ['MAIN'], 'DOCX: an embedded object no chunk names is not inflated against the limit');

    // A text box's blocks follow the paragraph drawing it: paragraphs apart, its list and table kept.
    // Alternate content is read from the branch this reader understands: a text box's Choice (the
    // Fallback here says so if taken), an emoji's Fallback (its Choice is `w16se`).
    const box = (inner: string) => `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:docPr id="2" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${p('FALLBACK')}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
    const item = (text: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${r(text)}</w:p>`;
    const boxed = await parse(docx(`<w:p>${r('Intro text.')}${box(p('Box title') + p('Box line two') + item('bullet one') + item('bullet two') + `<w:tbl><w:tr><w:tc>${p('Cell')}</w:tc></w:tr></w:tbl>`)}</w:p>${p('After')}`, {
        'word/numbering.xml': part('numbering', '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'),
    }), { extractAttachments: true });
    assert.deepStrictEqual(texts(boxed.ast.content), [['paragraph', 'Intro text.'], ['paragraph', 'Box title'], ['paragraph', 'Box line two'], ['list', 'bullet one'], ['list', 'bullet two'], ['table', undefined], ['paragraph', 'After']], 'DOCX: a text box\'s paragraphs, list and table follow the paragraph drawing it');
    assert.ok(!collectAllNodes(boxed.ast).some(n => n.type === 'image'), 'DOCX: a text box\'s drawing is no picture');
    for (const format of ['text', 'md', 'html', 'rtf', 'tex'] as const) {
        const out = (await boxed.ast.to(format)).value as string;
        assert.ok(out.includes('Box title') && out.includes('bullet two') && !/text\.Box|titleBox|twobullet|onebullet/.test(out), `${format}: a text box's paragraphs keep their word boundaries`);
    }
    for (const format of ['docx', 'odt'] as const) {
        const back = await OfficeParser.parseOffice(Buffer.from((await boxed.ast.to(format)).value as Uint8Array), { fileType: format } as any);
        assert.deepStrictEqual(back.content.map(n => n.text).filter(Boolean), ['Intro text.', 'Box title', 'Box line two', 'bullet one', 'bullet two', 'After'], `${format}: a text box's paragraphs are written apart`);
    }
    const emoji = await parse(docx(`<w:p>${r('I love ')}<w:r><mc:AlternateContent xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"><mc:Choice Requires="w16se"><w16se:symEx w16se:font="Segoe UI Emoji" w16se:char="1F600"/></mc:Choice><mc:Fallback><w:t>😀</w:t></mc:Fallback></mc:AlternateContent></w:r>${r(' emoji')}</w:p>`));
    assert.strictEqual(emoji.ast.content[0].text, 'I love 😀 emoji', 'DOCX: an emoji in alternate content is read from its Fallback');

    // SmartArt is a list after the paragraph drawing it: an item a line in every writer, and none
    // before the text after it.
    const point = (id: string, text: string) => `<dgm:pt modelId="${id}"><dgm:t><a:p><a:r><a:t>${text}</a:t></a:r></a:p></dgm:t></dgm:pt>`;
    const smartData = `<?xml version="1.0"?><dgm:dataModel ${W}><dgm:ptLst><dgm:pt modelId="0" type="doc"/>${point('1', 'Plan')}${point('2', 'Build')}${point('3', 'Ship')}</dgm:ptLst><dgm:cxnLst>${['1', '2', '3'].map((id, i) => `<dgm:cxn modelId="c${id}" srcId="0" destId="${id}" srcOrd="${i}"/>`).join('')}</dgm:cxnLst></dgm:dataModel>`;
    const smart = await parse(docx(`<w:p>${r('See:')}<w:r><w:drawing><wp:inline><wp:docPr id="3" name="Diagram"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>${p('After')}`, {
        'word/_rels/document.xml.rels': relsOf(rel('rId9', 'diagramData', 'diagrams/data1.xml')),
        'word/diagrams/data1.xml': smartData,
    }), { extractAttachments: true });
    assert.deepStrictEqual(texts(smart.ast.content), [['paragraph', 'See:'], ['list', 'Plan'], ['list', 'Build'], ['list', 'Ship'], ['paragraph', 'After']], 'DOCX: SmartArt is a list after its paragraph');
    assert.ok(!collectAllNodes(smart.ast).some(n => n.type === 'image'), 'DOCX: a SmartArt drawing is no picture');
    for (const format of ['docx', 'odt', 'rtf', 'md', 'html'] as const) {
        const back = await OfficeParser.parseOffice(Buffer.from((await smart.ast.to(format)).value as any), { fileType: format } as any);
        const lines = ((await back.to('text')).value as string).split('\n').map(line => line.replace(/^\W+/, '')).filter(Boolean);
        assert.deepStrictEqual(lines.slice(-4), ['Plan', 'Build', 'Ship', 'After'], `${format}: SmartArt items are lines of their own`);
    }
    // A line break the author typed (w:br, w:cr) and a tab are content, kept whatever includeBreakNodes
    // says (left out, "Line one<br/>Line two" read "Line oneLine two", and "A<tab>B" read "AB"); a page
    // break is layout, kept with includeBreakNodes.
    const breaks = docx('<w:p><w:r><w:t>Line one</w:t><w:br/><w:t>Line two</w:t></w:r></w:p><w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:cr/><w:t>C</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    const byDefault = (await parse(breaks)).ast;
    assert.strictEqual((await byDefault.to('text')).value, 'Line one\nLine two\nA\tB\nC', 'DOCX: line breaks and tabs are kept by default');
    assert.ok(!collectAllNodes(byDefault).some(n => n.type === 'break' && (n.metadata as any)?.breakType === 'page'), 'DOCX: a page break is not a node by default');
    assert.ok(collectAllNodes((await parse(breaks, { includeBreakNodes: true })).ast).some(n => n.type === 'break' && (n.metadata as any)?.breakType === 'page'), 'DOCX: a page break is a node with includeBreakNodes');
    assert.strictEqual((await (await parse(Buffer.from((await byDefault.to('docx')).value as Uint8Array))).ast.to('text')).value, 'Line one\nLine two\nA\tB\nC', 'DOCX: line breaks and tabs round-trip through DOCX');
    console.log('  DOCX reading review: All assertions passed ✓');
}

// ─── HTML and EPUB: reading as a browser does, writing what a reader accepts ─────

async function testHtmlBrowserReading(): Promise<void> {
    console.log('\n=== Running HTML/EPUB reading and writing tests ===');
    const parse = (src: string | Buffer, fileType = 'html', extra: object = {}) => OfficeParser.parseOffice(Buffer.isBuffer(src) ? src : Buffer.from(src), { fileType, onWarning: () => { }, ...extra } as any);
    const write = async (ast: any, format = 'html', extra: object = {}) => (await OfficeGenerator.generate(ast, format as any, { generateIds: false, htmlConfig: { standalone: false }, onWarning: () => { }, ...extra } as any)).value;
    const T = (text: string, extra: object = {}) => ({ type: 'text', text, ...extra });
    const doc = (content: any[], metadata: object = {}) => ({ type: 'docx', metadata, attachments: [], content } as any);
    const idsOf = (html: string) => [...html.matchAll(/\sid="([^"]*)"/g)].map(m => m[1]);
    const chapterOf = async (ast: any, extra: object = {}) => strFromU8(unzipSync((await write(ast, 'epub', extra)) as Uint8Array)['OEBPS/chapter1.xhtml']);
    // Characters XML does not allow at all, which no XHTML chapter may hold.
    const XML_INVALID = /[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/;

    // A paragraph holding a block: its wrapper (a quote's <blockquote>) is kept, and its ids go on its
    // first part, whatever that part is.
    const quoted = await parse('<blockquote><pre>code</pre></blockquote>');
    const quotedHtml = await write(quoted) as string;
    assert.ok(quotedHtml.includes('<blockquote><pre><code>code</code></pre></blockquote>'), `HTML: a quote holding only a code block keeps its <blockquote> (${quotedHtml})`);
    assert.strictEqual((await parse(quotedHtml)).content[0].metadata && ((await parse(quotedHtml)).content[0].metadata as any).style, 'Quote', 'HTML: the quoted code block reads back quoted');
    const anchored = await write(doc([
        { type: 'paragraph', metadata: { anchorIds: ['eq1'] }, children: [{ type: 'code', text: 'E=mc^2', metadata: { math: 'block' } }] },
        { type: 'paragraph', metadata: { anchorIds: ['bm1'] }, children: [{ type: 'break', metadata: { breakType: 'page' } }] },
        { type: 'paragraph', metadata: { anchorIds: ['p3'] }, children: [{ type: 'code', text: 'x' }, T('after')] },
    ])) as string;
    for (const id of ['eq1', 'bm1', 'p3']) assert.strictEqual(idsOf(anchored).filter(x => x === id).length, 1, `HTML: a paragraph starting with a block keeps its id ${id} (${anchored})`);
    assert.ok(anchored.indexOf('id="eq1"') < anchored.indexOf('math-block'), 'HTML: the id stands before the block it starts with');
    // Every block a paragraph holds is written outside its <p> (a table, an admonition, a video), and a
    // chapter holding one is well-formed.
    const holding = doc([
        { type: 'paragraph', children: [T('before'), { type: 'table', children: [{ type: 'row', children: [{ type: 'cell', children: [{ type: 'paragraph', children: [T('in a cell')] }] }] }] }, T('after')] },
        { type: 'paragraph', children: [T('x'), { type: 'admonition', metadata: { admonitionType: 'note' }, children: [{ type: 'paragraph', children: [T('inside')] }] }] },
        { type: 'paragraph', children: [T('Watch: '), { type: 'embed', metadata: { embedType: 'youtube', videoId: 'abc' } }, T(' now')] },
    ]);
    const holdingHtml = await write(holding) as string;
    const paragraphsHoldNoBlock = [...holdingHtml.matchAll(/<p[ >]/g)].every(m => !/<(?:div|table|ul|ol|p[ >])/.test(holdingHtml.slice(m.index! + 2, holdingHtml.indexOf('</p>', m.index!))));
    assert.ok(paragraphsHoldNoBlock && holdingHtml.includes('in a cell') && holdingHtml.includes('data-youtube-video="abc"'), `HTML: no block inside a <p> (${holdingHtml})`);
    parseXmlString(await chapterOf(holding));
    // The EPUB writer pairs each </p> with the <p> it closes (a paragraph written whole by onNode).
    const replaced = await chapterOf(doc([{ type: 'paragraph', children: [T('x')] }]), { onNode: (n: any) => n.type === 'paragraph' ? '<p>a<table><tr><td><p>cell</p></td></tr></table>b</p>' : undefined });
    parseXmlString(replaced);
    assert.ok(replaced.includes('<div>a<table>') && replaced.includes('b</div>'), `EPUB: a paragraph holding a table is promoted to a div (${replaced})`);

    // A note cited twice: each reference has an id of its own, and the note's back-link returns to the first.
    // (In the Markdown, the note `1-2` holds the id a second reference to note 1 would take, which passes it over.)
    for (const [label, src, type, second] of [
        ['HTML', '<p>a<sup data-footnote-ref="1">1</sup> b<sup data-footnote-ref="1">1</sup></p><section data-footnotes><div data-footnote-id="1">Note one</div></section>', 'html', 'footnote-ref-1-2'],
        ['Markdown', 'A[^1] and C[^1-2] and B[^1].\n\n[^1]: Note.\n\n[^1-2]: Other.\n', 'md', 'footnote-ref-1-3'],
    ] as const) {
        const ast = await parse(src, type);
        const html = await write(ast) as string;
        const ids = idsOf(html);
        assert.strictEqual(new Set(ids).size, ids.length, `HTML: ${label} footnote cited twice gives no duplicate id (${ids})`);
        assert.ok(ids.includes('footnote-ref-1') && ids.includes(second), `HTML: ${label} a later reference has an id of its own (${ids})`);
        assert.ok(html.includes('<a href="#footnote-ref-1">↩</a>'), 'HTML: the back-link returns to the first reference');
        const chapter = await chapterOf(ast);
        const chapterIds = idsOf(chapter);
        assert.strictEqual(new Set(chapterIds).size, chapterIds.length, `EPUB: ${label} footnote cited twice gives no duplicate id`);
        const back = await parse(html);
        assert.strictEqual(collectAllNodes(back).filter(n => n.type === 'note' && !(n.metadata as any)?.unreferenced).length >= 2, true, `HTML: both references read back (${label})`);
    }

    // Characters XML does not allow never reach an EPUB, whatever the source (a decoded `&#1;`, a raw
    // control character in Markdown or a CSV cell, metadata).
    for (const [src, type] of [['<p>a&#1;b&#11;c&#xFFFE;d</p>', 'html'], ['a\x01b\x0Bc\n', 'md'], ['a\x01b,c\x0Bd\n', 'csv']] as const) {
        const ast = await parse(src, type);
        const chapter = await chapterOf(ast);
        assert.ok(!XML_INVALID.test(chapter), `EPUB: no character XML forbids from ${JSON.stringify(src)}`);
        assert.ok(/a.?b/.test(chapter), 'EPUB: the text around it is kept');
    }
    const metaEpub = unzipSync(await write(doc([{ type: 'paragraph', children: [T('x')] }], { title: 'Ti\x01tle', author: 'A\x0Bb', description: 'd\uFFFE', keywords: 'k\x02' }), 'epub') as Uint8Array);
    for (const [name, bytes] of Object.entries(metaEpub)) if (/\.(xhtml|opf|xml)$/.test(name)) assert.ok(!XML_INVALID.test(strFromU8(bytes)), `EPUB: ${name} holds no character XML forbids`);
    assert.ok(strFromU8(metaEpub['OEBPS/content.opf']).includes('<dc:title>Title</dc:title>'), 'EPUB: the title is kept without the control character');
    // HTML reads a number past the last code point as U+FFFD, as it reads NUL and a surrogate.
    assert.strictEqual(collectAllNodes(await parse('<p>&#0;&#xD800;&#1114112;&#x80;</p>')).find(n => n.type === 'text')?.text, '\uFFFD\uFFFD\uFFFD\u20AC', 'HTML: out-of-range numeric references read as U+FFFD');

    // ── The tree, as a browser builds it ──
    const shape = (nodes: OfficeContentNode[]): string => nodes.map(n => n.type === 'text'
        ? JSON.stringify(n.text) + (n.formatting?.bold ? 'B' : '') + (n.formatting?.italic ? 'I' : '')
        : `${n.type}(${shape(n.children ?? [])})`).join(' ');
    const read = async (src: string | Buffer, extra: object = {}) => (await parse(src, 'html', extra)).content;
    const isSourceCommentNode = (n: OfficeContentNode | undefined) => n?.type === 'comment' && (n.metadata as any)?.sourceSyntax === 'html';
    const plainOf = (nodes: OfficeContentNode[]) => collectAllNodes({ content: nodes } as any).map(n => n.text ?? '').join(' ');
    // Markup declarations and prefixed names: Word's <o:p> and conditional markers, a DOCTYPE or XML
    // declaration before a fragment, an EPUB's <m:math>, are markup, never text.
    assert.strictEqual(shape(await read('<p>First item<o:p></o:p></p><p><![if !supportLists]>·<![endif]>List text</p>')), 'paragraph("First item") paragraph("·List text")', 'HTML: Word markup is read as markup');
    assert.strictEqual(shape(await read('<!DOCTYPE html><?xml version="1.0"?><p>x</p>')), 'paragraph("x")', 'HTML: declarations are read as nothing');
    const prefixedMath = await read('<p>x <m:math><m:mi>y</m:mi></m:math> z</p>');
    assert.ok(collectAllNodes({ content: prefixedMath } as any).some(n => n.type === 'code' && n.text === 'y' && (n.metadata as any)?.math === 'inline'), 'HTML: prefixed MathML is math');
    assert.strictEqual(shape(await read('<p>a < b <i>c</i> and 5<6</p>')), 'paragraph("a < b " "c"I " and 5<6")', 'HTML: a < that starts no tag is text, and the tags after it are tags');
    assert.strictEqual(shape(await read('<p>a<!-->b<!--->c<!-- x --!>d</p>')), 'paragraph("abcd")', 'HTML: comments end where HTML ends them');
    assert.strictEqual(shape(await read('<p>a<textarea><b>x</b> &amp;</textarea></p><title>T</title><noscript><p>Enable JS</p></noscript><template><p>inert</p></template><p>b</p>')),
        'paragraph("a" "<b>x</b> &") paragraph("b")', 'HTML: textarea holds text, and title, noscript and template are not shown');
    // A cell's implied end never reaches a table outside its own; a cell written in a table is given a row.
    const nested = await read('<table><tr><td>outer<table><td>inner</td></table>after</td></tr></table>');
    assert.strictEqual(shape(nested), 'table(row(cell("outer" table(row(cell("inner"))) "after")))', `HTML: a nested table's cell stays in it (${shape(nested)})`);
    assert.strictEqual(shape(await read('<b>bold<table><tr><td>x</b>y</td></tr></table>z')), 'paragraph("bold"B) table(row(cell("xy"B))) paragraph("z"B)', 'HTML: </b> in a cell does not end a <b> outside its table');
    // Formatting a paragraph's implied end closed goes on in the next paragraph, until its end tag.
    assert.strictEqual(shape(await read('<p><b>bold<p>still bold</b> after</p>')), 'paragraph("bold"B) paragraph("still bold"B " after")', 'HTML: formatting is carried across an implied paragraph end');
    assert.strictEqual(shape(await read('<b>1<i>2</b>3</i>4')), 'paragraph("1"B "2"BI "3"I "4")', 'HTML: misnested formatting');
    assert.strictEqual(shape(await read('<h1>A<h2>B</h2>')), 'heading("A") heading("B")', 'HTML: a heading ends an open heading');
    assert.strictEqual(shape(await read('<span>a<div>b</span>c</div>d')), 'paragraph("a") paragraph("bc") paragraph("d")', 'HTML: an end tag does not end a block opened inside it');
    assert.strictEqual(shape(await read('<html><head><title>T</title></head><p>before</p><body><p>in</p></body></html><p>after</p>')), 'paragraph("before") paragraph("in") paragraph("after")', 'HTML: content outside <body> is the body\'s');
    // Whitespace next to any inline element is a space (MathML, a custom element, Word's <o:p>), and
    // an element never shown between words leaves the space.
    const spaced = await read('<p><em>word</em> <math><mi>x</mi></math> <my-el>c</my-el> <o:p>o</o:p> <a href="#x">x</a> <script>1</script> <b>y</b></p>');
    assert.strictEqual(collectAllNodes({ content: spaced } as any).filter(n => n.type === 'text' || n.type === 'code').map(n => n.text).join(''), 'word x c o x y', 'HTML: whitespace next to inline elements is kept');
    // Linear time and bounded work are checked in test/security (htmlReadingTests).

    // ── Tables: a caption, and what stands between rows, go before the table ──
    assert.strictEqual(shape(await read('<table><caption>Table 1: Sales</caption><thead><tr><th>A<th>B<tbody><tr><td>1<td>2</table>')),
        'paragraph("Table 1: Sales") table(row(cell("A") cell("B")) row(cell("1") cell("2")))', 'HTML: a caption is a paragraph before its table');
    assert.strictEqual(shape(await read('<table><tr><td>a</td></tr>stray text<p>para</p></table>')), 'paragraph("stray text") paragraph("para") table(row(cell("a")))', 'HTML: stray content stands before the table');
    // Comments (preserveComments): between rows, before the table; between list items, at the end of the item before.
    const commented = { htmlParserConfig: { preserveComments: true } };
    const tableComments = await read('<table><!-- a --><tr><th>A</th></tr><!-- b --><tr><td>1</td></tr></table>', commented);
    assert.deepStrictEqual(tableComments.map(n => n.type), ['comment', 'comment', 'table'], 'HTML: comments among rows stand before the table');
    assert.ok((String((await OfficeGenerator.generate({ type: 'html', metadata: {}, attachments: [], content: tableComments } as any, 'md')).value)).includes('| A |\n| --- |\n| 1 |'), 'MD: the table keeps its header row');
    const listComments = await read('<ul><!-- first --><li>a</li><!-- c --><li>b</li></ul>', commented);
    assert.deepStrictEqual(listComments.map(n => n.type), ['comment', 'list', 'list'], 'HTML: a comment between items does not part the list');
    assert.ok(isSourceCommentNode(listComments[1].children![1]), 'HTML: the comment ends the item before it');

    // ── Footnotes in each writer's markup ──
    const notesOf = (nodes: OfficeContentNode[]) => collectAllNodes({ content: nodes } as any).flatMap(n => (n.notes ?? []).map(note => `${(note.metadata as any)?.noteId}:${note.text}`));
    for (const [writer, src, expected] of [
        ['GitHub', '<p>Claim<sup><a href="#user-content-fn-1" id="user-content-fnref-1" data-footnote-ref>1</a></sup>.</p><section data-footnotes class="footnotes"><h2 class="sr-only">Footnotes</h2><ol><li id="user-content-fn-1"><p>Source text here. <a href="#user-content-fnref-1" data-footnote-backref>↩</a></p></li></ol></section>', '1:Source text here.'],
        ['Pandoc', '<p>Text<a href="#fn1" class="footnote-ref" id="fnref1" role="doc-noteref"><sup>1</sup></a>.</p><section class="footnotes" role="doc-endnotes"><hr /><ol><li id="fn1"><p>The note.<a href="#fnref1" class="footnote-back" role="doc-backlink">↩︎</a></p></li></ol></section>', '1:The note.'],
        ['markdown-it', '<p>Here<sup class="footnote-ref"><a href="#fn1" id="fnref1">[1]</a></sup>.</p><section class="footnotes"><ol class="footnotes-list"><li id="fn1" class="footnote-item"><p>Note body <a href="#fnref1" class="footnote-backref">↩︎</a></p></li></ol></section>', '1:Note body'],
        ['EPUB 3', '<p>See<a epub:type="noteref" href="#n1">1</a>.</p><aside epub:type="footnote" id="n1"><p>Aside note.</p></aside>', 'n1:Aside note.'],
    ] as const) {
        const nodes = await read(src);
        assert.deepStrictEqual(notesOf(nodes), [expected], `HTML: a ${writer} footnote is a note (${shape(nodes)})`);
        // (Its text is in the note alone, not also in the body; its back-link is in neither.)
        assert.ok(!JSON.stringify(nodes).includes('↩') && !shape(nodes).includes(expected.split(':')[1]), `HTML: the ${writer} note's text is in its note alone (${shape(nodes)})`);
    }
    assert.strictEqual(shape(await read('<section data-footnotes><p>Plain content</p></section>')), 'paragraph("Plain content")', 'HTML: a footnotes section with no note in it is content');
    assert.strictEqual(shape(await read('<p>See<a epub:type="noteref" href="#n1">1</a></p><aside epub:type="footnote" id="n2"><p>Uncited.</p></aside>')), 'paragraph("See" "1") paragraph("Uncited.")', 'HTML: an uncited aside stays content, a link to no note a link');

    // ── The document: its title, language, encoding and base ──
    const titled = await parse('<html lang="fr"><head><title>A &amp;  B\n</title></head><body><p>x</p></body></html>');
    assert.strictEqual(titled.metadata.title, 'A & B', 'HTML: the title is decoded and its whitespace collapsed');
    assert.strictEqual(titled.metadata.language, 'fr', 'HTML: <html lang> is the language');
    assert.ok((await write(titled, 'html', { htmlConfig: { standalone: true } }) as string).includes('<html lang="fr">'), 'HTML: the language is written back');
    assert.ok(strFromU8(unzipSync(await write(titled, 'epub') as Uint8Array)['OEBPS/content.opf']).includes('<dc:language>fr</dc:language>'), 'EPUB: the language is the book\'s');
    assert.strictEqual((await parse('<body><title>T</title><p>x</p></body>')).metadata.title, 'T', 'HTML: a title outside the head is the title, not text');
    const windows1252 = Buffer.concat([Buffer.from('<meta http-equiv="Content-Type" content="text/html; charset=windows-1252"><p>caf'), Buffer.from([0xE9, 0x20, 0x92, 0x71, 0x92]), Buffer.from('</p>')]);
    assert.strictEqual(collectAllNodes(await parse(windows1252)).find(n => n.type === 'text')?.text, 'café \u2019q\u2019', 'HTML: a page in its declared encoding');
    assert.strictEqual(collectAllNodes(await parse(Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from('<p>h\u00E9llo</p>', 'utf16le')]))).find(n => n.type === 'text')?.text, 'h\u00E9llo', 'HTML: a UTF-16 page with its byte order mark');
    assert.strictEqual(collectAllNodes(await parse('<meta charset="windows-1252"><p>\u00E9</p>')).find(n => n.type === 'text')?.text, '\u00E9', 'HTML: UTF-8 bytes stay UTF-8 whatever the page declares');
    const based = await read('<head><base href="https://example.com/dir/"></head><a href="page.html">p</a> <a href="#frag">f</a> <img src="i.png"> <img srcset="small.png 1x, big.png 2x">');
    assert.deepStrictEqual(collectAllNodes({ content: based } as any).map(n => (n.metadata as any)?.link ?? (n.metadata as any)?.url).filter(Boolean),
        ['https://example.com/dir/page.html', '#frag', 'https://example.com/dir/i.png', 'https://example.com/dir/small.png'], 'HTML: <base href> resolves links and pictures; srcset gives a picture');

    // ── EPUB: every chapter the spine lists is read, or reported ──
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');
    const book = (manifest: string, spine: string, files: Record<string, string | Uint8Array>) => Buffer.from(zipSync({
        mimetype: strToU8('application/epub+zip'),
        'META-INF/container.xml': strToU8('<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>'),
        'OEBPS/content.opf': strToU8(`<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>T</dc:title></metadata><manifest>${manifest}</manifest><spine>${spine}</spine></package>`),
        ...Object.fromEntries(Object.entries(files).map(([k, v]) => [k, typeof v === 'string' ? strToU8(v) : v])),
    }));
    const chapter = (body: string) => `<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>c</title></head><body>${body}</body></html>`;
    const epubWarnings: string[] = [];
    const epub = await parse(book(
        '<item id="c1" href="chapter%201.xhtml" media-type="application/xhtml+xml"/><item id="c2" href="ch2.xml" media-type="application/xhtml+xml"/><item id="c3" href="gone.xhtml" media-type="application/xhtml+xml"/><item id="i1" href="image_1.png" media-type="image/png"/><item id="i2" href="my%20pic.png" media-type="image/png"/>',
        '<itemref idref="c1"/><itemref idref="c2"/><itemref idref="c3"/>',
        {
            'OEBPS/chapter 1.xhtml': chapter(`<p>One</p><p><img src="image_1.png" alt="book"/><img src="data:image/png;base64,${png.toString('base64')}" alt="inline"/><img src="my%20pic.png" alt="spaced"/></p>`),
            'OEBPS/ch2.xml': chapter('<p>Two</p>'), 'OEBPS/image_1.png': png, 'OEBPS/my pic.png': png,
        }), 'epub', { extractAttachments: true, onWarning: (w: any) => epubWarnings.push(w.code) });
    assert.strictEqual(plainOf(epub.content).includes('One') && plainOf(epub.content).includes('Two'), true, 'EPUB: a percent-encoded href and a .xml chapter are read');
    assert.deepStrictEqual(epubWarnings, ['CONTENT_PART_NOT_READ'], 'EPUB: a chapter missing from the archive is reported');
    const pictures = collectAllNodes(epub).filter(n => n.type === 'image').map(n => `${(n.metadata as any).altText}=${(n.metadata as any).attachmentName}`);
    assert.deepStrictEqual(pictures, ['book=image_1.png', 'inline=image_1-2.png', 'spaced=my pic.png'], 'EPUB: each picture keeps the attachment it shows');
    const drm = book('<item id="c1" href="ch1.xhtml" media-type="application/xhtml+xml"/>', '<itemref idref="c1"/>', {
        'OEBPS/ch1.xhtml': new Uint8Array([0x9F, 0x12, 0x80, 0x33, 0xFF]),
        'META-INF/encryption.xml': '<?xml version="1.0"?><encryption xmlns="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:enc="http://www.w3.org/2001/04/xmlenc#"><enc:EncryptedData><enc:EncryptionMethod Algorithm="http://www.w3.org/2001/04/xmlenc#aes128-cbc"/><enc:CipherData><enc:CipherReference URI="OEBPS/ch1.xhtml"/></enc:CipherData></enc:EncryptedData></encryption>',
    });
    await assert.rejects(parse(drm, 'epub'), /encrypted \(DRM/, 'EPUB: a book whose chapters are encrypted is refused with a clear error');

    // ── Writing: an ordered task list, generated heading ids, ids an EPUB accepts ──
    const tasks = await parse('1. [ ] one\n2. [x] two\n\n# Intro\n\ntext\n\n# Intro\n\n# 2024 Results\n\n[second](#intro-1) [year](#2024-results)\n', 'md');
    const tasksHtml = await write(tasks, 'html', { generateIds: true }) as string;
    assert.ok(/<ol data-type="taskList">[\s\S]*<\/ol>/.test(tasksHtml) && !tasksHtml.includes('<ul'), `HTML: an ordered task list is an <ol> closed as one (${tasksHtml})`);
    const tasksBack = collectAllNodes(await parse(tasksHtml)).filter(n => n.type === 'list').map(n => `${(n.metadata as any).listType}:${(n.metadata as any).isTask}`);
    assert.deepStrictEqual(tasksBack, ['ordered:true', 'ordered:true'], 'HTML: an ordered task list reads back ordered');
    assert.deepStrictEqual(idsOf(tasksHtml).filter(id => id.startsWith('intro')), ['intro', 'intro-1'], 'HTML: a second heading of the same text is numbered, as GitHub numbers it');
    const taskChapter = await chapterOf(tasks, { generateIds: true });
    parseXmlString(taskChapter);
    assert.ok(taskChapter.includes('id="_2024-results"') && taskChapter.includes('href="#_2024-results"') && taskChapter.includes('href="#intro-1"'), 'EPUB: an id starting with a digit is made a name, and the link to it with it');
    const anchoredChapter = await chapterOf(doc([{ type: 'paragraph', metadata: { anchorIds: ['a', 'b c', '9'] }, children: [T('x'), T('to 9', { metadata: { link: '#9', linkType: 'internal' } })] }]));
    const anchorIds = idsOf(anchoredChapter);
    assert.ok(anchorIds.every(id => /^[A-Za-z_][\w.-]*$/.test(id)) && new Set(anchorIds).size === anchorIds.length && !anchoredChapter.includes(' name="'), `EPUB: every id is an XML name, and anchors carry no name attribute (${anchorIds})`);
    assert.ok(anchoredChapter.includes('href="#_9"'), 'EPUB: a link follows its renamed id');

    // ── MHT: a soft line break padded with spaces or tabs is still one (RFC 2045) ──
    const { readMht, mhtPartText } = await import('../src/utils/mhtUtils');
    const mht = 'MIME-Version: 1.0\r\nContent-Type: text/html; charset="utf-8"\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n<p>Soft =  \r\nbreak with=\t\r\ntabs, a plain=\r\nbreak, =3D, and =\nLF</p>';
    assert.strictEqual(mhtPartText(readMht(Buffer.from(mht, 'latin1'))[0]), '<p>Soft break withtabs, a plainbreak, =, and LF</p>', 'MHT: padded soft line breaks join their lines');

    console.log('  HTML/EPUB reading and writing: All assertions passed ✓');
}

async function runTests(): Promise<void> {
    console.log('Starting exhaustive officeParser test suite...');
    let passed = 0;
    let failed = 0;

    const tests: Array<[string, () => Promise<void>]> = [
        ['Markdown', testMarkdown],
        ['Markdown round trips', testMarkdownRoundTrips],
        ['Second review', testSecondReview],
        ['Image links', testImageLinks],
        ['DOCX reading review', testDocxReadingReview],
        ['HTML', testHtml],
        ['HTML browser reading', testHtmlBrowserReading],
        ['SourceComments', testSourceComments],
        ['CSV', testCsv],
        ['RTF', testRtf],
        ['AttributeRoundtrip', testAttributeRoundtrip],
        ['BlobInput', testBlobInput],
        ['WordParagraphMark', testWordParagraphMarkFormatting],
        ['GeneratedOutput', testGeneratedOutput],
        ['ODG', testOdg],
        ['ODFComments', testOdfComments],
        ['PPTX comments', testPptxComments],
        ['ConsistencyBehaviors', testConsistencyBehaviors],
        ['DOCX', testDocxGeneration],
        ['DOCX notes and bookmarks', testDocxNotesAndBookmarks],
        ['ODT', testOdtGeneration],
        ['LaTeX', testLatexGeneration],
        ['LaTeX parsing', testLatexParsing],
        ['Config consistency', testConfigConsistency],
        ['OfficeGenUtils', testOfficeGenUtils],
        ['NativePdfEngine', testNativePdfEngine],
        ['Template', testTemplate],
        ['Cancellation', testCancellation],
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
