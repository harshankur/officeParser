# Changelog

All notable changes to `officeParser` are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [7.6.0] - 2026-08-14
### Added
- **Attribute-driven HTML interop for wikilinks, citations, math and mermaid.** `HtmlParser` now accepts the on-the-wire shapes structured editors emit, with no change to default output: `a[data-wikilink]` (page in `data-target`, display text from the anchor body or `data-alias`), a `span.citation` carrying `data-key`, math whose `data-math` holds the raw LaTeX with the mode in the class token (the delimited `data-math="inline|block"` shape is still read exactly as before; any other `data-math` value, previously read as inline math with the attribute ignored, is now taken as the LaTeX source), and `div[data-mermaid]`/`div.mermaid`/`pre.mermaid` mapped to a `mermaid` code node (previously the div flattened to plain text).
- **`HtmlGeneratorConfig.sourceAttributes`** (default `false`): carry each rich node's raw source in a `data-*` attribute with undelimited text content, so an attribute-driven consumer can rehydrate the node - `data-wikilink`/`data-target`/`data-alias` alongside the existing wikilink attributes, a `<span class="citation" data-key>` instead of `<cite>`, the LaTeX undelimited in `data-math`, and a `<div class="mermaid" data-mermaid>` instead of `<pre><code>`. Off by default the output is byte-identical to before, and the widened parser reads every shape it emits, so output stays self-round-trippable. Every sink is entity-escaped, and PDF/EPUB generation force the flag off.
- **`HtmlParserConfig.preserveIframes`** (default `false`): keep non-YouTube `<iframe>` embeds, which are otherwise dropped, as `embed` nodes (`embedType: 'iframe'`). `true` preserves any iframe; an array is a hostname allowlist (an entry matches the host exactly or as a `.`-suffix). The src is still scheme-checked on generation, so a `javascript:`/`data:` src never survives. The same flag governs a raw `<iframe>` block in Markdown input. `EmbedMetadata` widens accordingly: `embedType` is `'youtube' | 'iframe'`, `videoId` is optional, and a `height` field is added. **Type note for strict TypeScript consumers:** because `videoId` is now optional and `embedType` is a wider union, code that read `EmbedMetadata.videoId` as a non-optional `string`, or that switched exhaustively on `embedType`, may need a small type adjustment on upgrade.
- **Multi-line footnote definitions.** A `[^id]: ...` definition may continue across lines indented one level, 4 spaces or a tab (Pandoc/GFM); the generator indents a wrapped definition's continuation lines to match, so a multi-line body round-trips instead of emitting a definition a bare newline would cut short. Single-line definitions are unchanged.
- **`parseOffice` and `OfficeConverter.convert` accept a web `Blob`/`File`** (or any `BlobLike` - an object with an `arrayBuffer()` method), so browser callers no longer convert to a Buffer first. A filename, when present, drives extension-based type detection; a nameless blob resolves through magic-byte sniffing.
- **`MdGeneratorConfig.fallbackToHtml.inlineFormatting`** (default `false`, opt-in even when `fallbackToHtml` is `true`): round-trips inline text color, highlight and font size through `.md` as a `<span style="...">` run that the Markdown parser reads back. These have no Markdown syntax and were otherwise silently lost when `.md` is the storage format. Default output is unchanged; emitted CSS values are sanitized.
- **`<mark>` and `data-color` are now parsed** into highlight/background formatting on import (previously a highlight survived only when the producer also wrote an inline `background-color`). **Behaviour change:** a document containing a bare `<mark>` now round-trips as a highlighted `<span>` (default yellow) rather than losing the highlight.
- **Highlights are now generated as `<mark>`** (`<mark data-color="…" style="background-color: …">`) rather than a `<span style="background-color: …">`. Editors whose highlight extension matches only the `mark` element (e.g. Tiptap's Highlight) now rehydrate a highlighted run that previously came back as plain text; the widened parser reads the emitted shape back, and colour/font-size still ride their own span so a run carrying both rehydrates both. **Behaviour change** to the default highlight markup.
- **The public config and metadata types are exported from the package root** - `HtmlGeneratorConfig`, `HtmlParserConfig`, `MdGeneratorConfig`, `MarkdownDialectConfig`, `StandaloneConfig`, `MetadataOverrides`, `FallbackToHtmlConfig`, the per-node `*Metadata` interfaces, and the other generator sub-configs - so `import type { HtmlGeneratorConfig } from 'officeparser'` resolves (previously only the browser `.d.ts` carried them).

### Fixed
- **A paragraph mark's run properties bled onto every run.** A DOCX paragraph mark's run properties (`<w:pPr><w:rPr>`) were folded into the base formatting applied to all of the paragraph's text runs, so bold/italic/underline/color/size/font set on the paragraph mark appeared on text that never specified it. Per OOXML ISO 29500 §17.3.1.29 those properties format only the paragraph mark glyph; runs now inherit only from the style chain (paragraph and character styles) and their own run properties. **Behaviour change:** formatting on affected runs differs from prior releases for any document that carried paragraph-mark formatting. (#109)
- **Chunking dropped HTML- and Markdown-origin paragraph text.** `ChunkingGenerator` read a node's own `.text` with no fallback to its children, so paragraphs built as `{ children: [...] }` (which every HTML and `.md` document produces) chunked to empty and were filtered out. It now collects text recursively, matching the text generator, restoring near-parity with `.to('text')`.
- **Empty metadata corrupted the Markdown save/reload cycle.** A document with no metadata fields (a bare editor fragment) emitted an empty `---\n---` frontmatter block, which reparsed as a `## ---` setext heading and accumulated on every save. The generator now omits the fence when no field is present, and the parser recognises an empty frontmatter block.
- **The footnote-definition HTML nested `<p>` inside `<p>`.** It was emitted as a `<p data-footnote-id>` wrapping block content, which every DOM parser splits (leaving the wrapper empty). It is now a `<div data-footnote-id>`, which is valid and matches attribute-driven editors. **Behaviour change** to the default footnote-definition markup, which was broken by construction before.
- **`<pre><code>` blocks did not decode HTML entities.** Escaped characters in a code block (`&lt;`, `&gt;`, `&amp;` - e.g. `a &lt; b`, or a mermaid `--&gt;` arrow) were carried into the code node's text verbatim, so they surfaced still-escaped in text/Markdown/chunk output and double-escaped a little more on every HTML round trip. They are now decoded to their literal characters, matching inline code and ordinary text. **Behaviour change** for HTML sources whose code blocks contain escaped entities.
- **`rtfConfig` was the one generator sub-config not deep-merged** in config resolution. `RtfGeneratorConfig` is currently a reserved, empty interface, so this had no field-level effect yet; the merge is corrected now so future fields behave like every other sub-config.
- **Footnotes made a Markdown document grow a `### Notes` heading on every save.** The Markdown generator prefixed the collected `[^id]:` definitions with a literal `---\n\n### Notes\n\n`, which reparsed as a real heading plus a thematic break, so each save re-emitted the parsed heading and a fresh one, unbounded. The definitions are now emitted on their own (GitHub/Pandoc render the footnotes section automatically), making a save/reload cycle byte-stable. The `[^id]` reference marker also stays at its citation point (`Body[^1].`) rather than migrating to the end of the paragraph, aligning the Markdown and HTML generators. **Behaviour change** to the Markdown footnote-section markup. **Migration note:** a `.md` file saved by officeParser &le; 7.5.x carries the old `### Notes` section; on its first re-save that line is parsed as a genuine heading and kept once (a single `### Notes` heading, bounded - it does not accumulate, and settles by the second cycle). The parser deliberately does not swallow it, so a document's own genuine `### Notes` heading is never removed. Documents authored at &ge; 7.6.0 are unaffected.
- **Footnote bodies never reached RAG chunks.** `ChunkingGenerator` collected a node's children but not its `notes`, and emitted a content node's chunk without recursing into them, so a footnote's text was absent from the chunk index although present in `.to('text')`. Note bodies are now folded into the referencing node's chunk text.
- **Quoted frontmatter scalars lost their type on a Markdown round trip.** A quoted YAML scalar (`version: "123"`, `flag: "true"`) was unquoted and then coerced to a number/boolean, which the generator re-emitted unquoted, destroying the string permanently after one save. A quoted scalar now stays a string; only bare, unquoted scalars coerce (matching YAML semantics). **Behaviour change:** a quoted numeric or boolean now retains its declared string type and re-emits quoted, rather than re-emitting unquoted as a number/boolean.
- **Orphan footnote definitions were silently dropped.** A `[^x]: …` definition with no matching `[^x]` reference was discarded on any conversion, so deleting a reference but keeping its definition lost the definition on the next save. Unreferenced definitions are now preserved on both sides - recovered from a `.md` document and from a `section[data-footnotes]` on the HTML side - and emitted in the footnotes section on output (a `[^x]:` line in Markdown, a `div[data-footnote-id]` inside `section[data-footnotes]` in HTML) with no citation marker and no dangling back-link, so an unreferenced definition survives a full md -> HTML -> md round trip. **Behaviour change:** the definition is now kept rather than lost.
- **Thematic breaks (`---` / `<hr>`) were lost on save.** A Markdown `---` and an HTML `<hr>` parsed to a page break, which the Markdown generator emits as a bare newline, so a horizontal rule silently vanished on the first save and re-collapsed on the next. A thematic break is now a distinct `breakType: 'thematic'` (an additive metadata value) that emits `---` in Markdown and `<hr>` in HTML; an office page break (`<hr class="page-break">`) stays a page break. **Behaviour change:** a `---` / `<hr>` now survives a round trip instead of collapsing to whitespace.
- **A footnote referenced in a table cell was defined twice.** The HTML generator re-processes table cells (header-row detection and sparse-column rows) after they were already processed, which pushed a cell's footnote into the collected definitions twice - a duplicated `id`, definition, and back-links. Collected notes now de-duplicate by identity, so a table-cell footnote is defined exactly once.
- **Repeated references to one footnote id were duplicated.** Two `[^1]` references to a single definition renumbered to `[^1]`/`[^2]` on save with the body copied. Repeated references now share one note, so they stay `[^1]` with a single definition; office notes that merely share a numeric id (a footnote and an endnote both numbered `1`) remain distinct.
- **Documentation**: `generate(ast, 'chunks')` returns an `OfficeChunk[]` array, not a JSON string (the code comment and README said otherwise); consumers serialize to JSON/JSONL themselves.

## [7.5.1] - 2026-07-28
### Added
- **`error.officeIssue`**, carrying the same structured issue warnings use, so a failed parse is identified by a stable `code` rather than by matching message text. The `OfficeError` type is exported for typed access.
- **`NO_WORKSHEETS_FOUND` and `NO_SLIDES_FOUND` warnings**, so a chartsheet-only workbook or a slide-less presentation says why it is empty instead of returning silently.

### Fixed
- **The browser bundles broke webpack builds.** One dynamic import escaped the build's ignore annotations, because the rule that added them treated an interpolated template literal as a fixed string. Webpack does not skip a specifier it cannot resolve: it builds a context module over the whole directory, so consumers ended up bundling all of `dist/`, Node-only files included, and their build failed on `child_process`. The bundles also still named `child_process` and `url` directly, from host-inspection paths that only ever run under Node; both now resolve to browser stubs. A webpack 5 build of the ESM and slim ESM bundles is now warning-free. (#108)
- **Valid documents were rejected as `zip` files.** Magic-byte detection reports a generic `zip` for archives that place too much data before their format declaration, which streaming ZIP writers and part-heavy decks both do. When detection is inconclusive the archive is now opened and the format read from `[Content_Types].xml` or the `mimetype` entry. Affected every ZIP-backed format, in Node and the browser, and also made a correct `fileType` hint report a false `BUFFER_TYPE_MISMATCH`. (#82)
- **Corrupt input parsed as an empty document instead of failing.** Since 7.3.0 an unreadable, truncated, or mislabeled archive resolved to an empty AST with no warnings, indistinguishable from a genuinely empty file. These now reject with `ZIP_NO_ENTRIES_FOUND`, `ZIP_TRUNCATED`, or `REQUIRED_PART_MISSING` (a readable archive missing the part its format requires). Truncation inside an entry could also leave the promise pending forever. (#107)
- **DOCX headers, footers and comments were never extracted.** Those parts were left out of the archive extraction, so the code that parses them could never run and `ast.auxiliary` was always empty. `ignoreHeadersAndFooters` and `ignoreComments` now gate extraction as well, and part names match past `header9.xml`.
- **ODF documents could parse as empty or as the wrong content.** An ODS/ODP with no `mimetype` entry was walked as a text document, and a missing root `content.xml` could promote an embedded `Object N/content.xml` to the document body.
- **Errors were reported twice, with a doubled `[OfficeParser]: ` prefix and their code flattened to `FILE_CORRUPTED`.** Archive extraction errors also ignored `outputErrorToConsole` and `onWarning`, always writing to the console.
- **Reusing one config object across calls leaked state between documents.** A sufficiently complete config was used directly rather than copied, so each run wrote to the caller's object: later parses appended their warnings to earlier, already-returned `ast.warnings`, and generation rewrote an invalid `htmlConfig.containerWidth` to `'auto'` on the caller's own config, silencing the warning on every later run. Config resolution now always returns an object it owns, on both the parser and generator side. `abortSignal` and callbacks keep their identity.

## [7.5.0] - 2026-07-26
### Added
- **Equations are parsed from every format that can carry them.** DOCX and PPTX equations (OMML, `<m:oMath>`) and native MathML (`<math>`, used by ODF embedded objects, HTML pages and EPUB3 spine items) are now converted to LaTeX and emitted as `code` nodes carrying `CodeMetadata.math`, the same shape Markdown's `$...$` already produced. Fractions, sub/superscripts, radicals, delimiters, n-ary operators, functions, accents, matrices and math alphabets (`ℝ` and friends) are all carried through, and a document's own `<annotation encoding="application/x-tex">` is preferred over anything reconstructed from the presentation markup. (#97)
- **`includeBreakNodes` now works for ODF**, not just DOCX. ODF attaches breaks to the paragraph *style* (`fo:break-before`/`fo:break-after`) rather than writing an inline element, so nothing was ever found to emit; those now produce `break` nodes around the paragraph, and `<text:soft-page-break/>` maps onto the same `lastRenderedPage` type DOCX uses. (#104)

### Fixed
- **Equations were silently corrupted rather than merely missing.** DOCX and HTML/EPUB equations fell through to generic text concatenation, so `<m:num>1</m:num><m:den>2</m:den>` and `<mfrac><mn>1</mn><mn>2</mn></mfrac>` both came out as `12` - a plausible-looking number with no indication anything was lost. PPTX equations were dropped outright. ODF equations survived but in an ad hoc notation (`(1)/(2)`) that no generator could render as maths. (#97)
- Table subtrees were rendered twice by the Markdown generator, making conversion time grow quadratically with nesting depth and duplicating any footnote inside a table cell. (#105)
- Paragraph-level text properties (bold, font size, colour, font family) in ODF documents were parsed into the style table but never applied to the paragraph's text, so they were absent from every generated output. (#104)
- ODF styles that explicitly turn formatting *off* (`fo:font-weight="normal"`, `style:text-underline-style="none"`, and the italic/strikethrough equivalents) were not recorded at all, so a span could not override an inherited paragraph style. LibreOffice writes exactly this whenever part of a bold paragraph is un-bolded. Such styles now appear in `metadata.styleMap` with the flag set to `false`, and a span carrying one clears the inherited value.
- **Formatting a heading or table header row already implies is no longer emitted twice.** A heading whose every run is bold produced `# **Heading**` in Markdown, and in RTF and HTML an inner font size that *overrode* the heading's own and visibly shrank it. Uniform bold/size inside a heading or header row is now dropped in favour of the element's own styling; partial emphasis (`# Normal **Bold** Normal`) is untouched, and colour, font and every other property still come through.
- Text inside PowerPoint grouped shapes (`<p:grpSp>`), including nested groups, was silently dropped. (#106)
- A crafted ODF table row carrying a large `table:number-rows-repeated` with no cells bypassed the per-document cell budget entirely and could exhaust memory.

## [7.4.0] - 2026-07-19
### Added
- **Markdown output dialect** (`MdGeneratorConfig.dialect`): target a real-world flavor (`'github'`, `'gitlab'`, `'obsidian'`, `'pandoc'`, strict `'commonmark'`, or the default `'extended'`), or pass a `MarkdownDialectConfig` object to control admonitions, footnotes, citations, wikilinks, math, tables, list markers and emphasis per feature. The default is byte-identical to prior output.
- **`MdGeneratorConfig.fallbackToHtml` also accepts a `FallbackToHtmlConfig` object**, splitting the single flag into independently-controllable parts (text formatting, alignment, anchors, tables, embeds, cell line breaks). The boolean form is unchanged.
- **`GeneratorConfig.metadataOverrides`**: set the metadata written into generated output (title, author, description, subject, keywords, lastModifiedBy, created, modified, language, plus arbitrary `custom` pairs) without mutating the parsed AST, merged per field, applied across HTML, EPUB, Markdown, RTF and the text/CSV metadata header. Setting `modified` also makes EPUB output byte-stable across runs. Formats with a fixed metadata vocabulary (EPUB, RTF) warn on unrepresentable `custom` keys via `onWarning` rather than dropping them silently.
- **`subject` and `keywords` now reach generated output** (HTML `<meta>`, Markdown frontmatter, EPUB `dc:subject`, RTF `\info`); previously they were parsed but never written.
- **Opt-in HTML attribute pass-through** (`htmlParserConfig.preserveAttributes`, default off): preserves generic attributes on `BaseContentNode.htmlAttributes` that previously vanished on an HTML round trip. Event handlers, `srcdoc`, `style` and `id` are never carried.
- **`TextGeneratorConfig.renderNotes`** (default `true`): controls whether the collected footnote/endnote section is appended to `.to('text')` output.
- **Footnotes degrade gracefully when disabled** (`dialect: { footnotes: false }`, or under `commonmark`): note content is inlined as a parenthetical at the reference point instead of being dropped.
- **`OfficeParserConfig.mdParserConfig`**: per-input-format parser config sub-object mirroring the generator side. Reserved for future use.
- **`AdmonitionMetadata.sourceSyntax`** (`'github' | 'gitlab'`): records which concrete syntax produced an admonition, for round-trip-aware tooling.
- **Expanded CommonMark / Markdown Extra parsing**: backslash escapes, reference-style links/images, underscore emphasis, multi-backtick inline code, setext headings, `<url>` autolinks, HTML entity and numeric-character-reference decoding, hard vs. soft line breaks, `~~~`-fenced and 4-space-indented code blocks, and list-item continuation lines.

### Fixed
- **Several Markdown constructs emitted document text without escaping** (math, wikilinks, citation keys, footnote ids, abbreviation definitions, image attribute lists, admonition types/titles), so a hostile document's markup could survive a parse-then-generate cycle intact. Each now escapes or allowlists its content.
- **A crafted ODF spreadsheet could exhaust memory.** `table:number-columns-repeated` / `table:number-rows-repeated` were expanded into AST nodes with no bound, so a few hundred bytes could materialize millions of cells. A per-document cap (`decompressionLimits.maxTableCells`, default 1000000) now clamps and warns instead of expanding without limit.
- **`abortSignal` could not interrupt work already in progress.** It was checked once per parser and honoured by only one generator; parsers now consult it through their content loops and all generators check during traversal.
- **`styleMap` class lists and attribute names were emitted unescaped/unvalidated** on the spreadsheet row and sheet paths, allowing attribute injection through a crafted mapping.
- **Config resolution could pollute `Object.prototype`** when a generator config was supplied via `JSON.parse` (a `__proto__` key reached the prototype through the recursive merge). Guarded.
- **`sanitizeCssValue` could be bypassed with CSS backslash escapes** (`u\rl(...)` reassembled into a live `url()` after the safety check); the escape strip now runs before the check.
- **RTF hyperlinks had no scheme allowlist**, so `javascript:`, `file://` and UNC targets were emitted verbatim; they are now restricted to the schemes HTML and Markdown already allow. **Behaviour change:** intranet `file://`/UNC links lose their clickable target (the link text is kept).
- **The HTML parser's nesting-depth guard never fired** (it tripped at 1000, but the stack overflowed around 800). Lowered to 256 so a deeply nested document raises the typed error instead of a `RangeError`.
- **`.to('text')`/`.to('md')` stripped genuine leading/trailing document whitespace** by trimming the whole output; only the generator's own separator artifact is stripped now (issue #102).
- **EPUB generation was not reproducible**: `dcterms:modified` and every zip entry's mtime defaulted to the current time, so the same AST produced different bytes each run. Both now derive from one resolved instant, falling back to the document's own modification date.
- **`.to('text')` dropped chart data series and CSV comments** (nodes carrying content in `text` with no children) and **concatenated adjacent table cells** with no separator (`ITEMNEEDED`). Both fixed.
- **The plain-text `renderMetadata` header could be forged**: a newline in a document's title rendered as extra `Key: value` fields. Line breaks are now folded to spaces.
- **Inline styles were matched by substring**, so `color:` matched inside `background-color:` and `max-width: 100%` was read as an explicit width, while spaceless and vendor-prefixed forms were missed. Replaced with a real declaration parser.
- **`csvSafeCell` tested its formula-injection trigger against the untrimmed value**, so a leading-space formula (`" =1+1"`) slipped through unprefixed.
- **Markdown parser correctness**: nested-list numbering leaking across siblings, an indentation-unit mismatch that doubled nesting depth on re-parse, `> >` blockquotes stripping only one `>` level, `)`-style ordered lists (`1) item`) not recognized, `----- text` misread as a horizontal rule, and short table separator rows (`|-|-|`) rejected.
- **Out-of-range numeric character references crashed the parser**: `&#999999999;` (Markdown `&#N;`/`&#xH;`, and XML entities via `xmlUtils.ts` used by e.g. `ExcelParser`) threw a `RangeError`; such references are now left as literal text.

### Changed
- **`styleMap`'s `output.tag` now takes effect in HTML output.** It was silently ignored (shadowed in every branch), while Markdown and RTF honoured it and the README documented it as working. A hostile tag is now rejected against an element allowlist. This also activates the built-in default style mappings for HTML, so a `Heading N`/`Quote`/`Title`-styled paragraph maps to the corresponding element instead of remaining a `<p>`.
- **Node.js 18/20 caveat.** Some hardening (notably the parser nesting-depth threshold) is calibrated against current Node; on 18/20 exact thresholds may differ, though failure stays graceful. Both remain supported until the next major version.
- **Documentation fix**: `TextGeneratorConfig.preserveLayout` defaults to `true`, not `false` as previously documented.

## [7.3.0] - 2026-07-08
### Added
- **EPUB Support (Parser & Generator)**: `epub` is now a first-class `SupportedFileType`/`UniversalGeneratorFormat`.
  - `EpubParser` unzips the archive, resolves the spine's reading order from `content.opf`, and parses each XHTML document through the existing `HtmlParser`, so EPUB content shares the same AST shape (and the same Markdown-dialect fidelity below) as every other format. Dublin Core metadata (`dc:title`, `dc:creator`, `dc:description`, `dc:subject`, `dc:date`, `dc:publisher`, `dc:language`, `dc:identifier`) maps into `ast.metadata`/`ast.metadata.nativeProperties`; cover art is exposed via `metadata.customProperties.coverImageName`.
  - `EpubGenerator` renders the AST through `HtmlGenerator` and packages a minimal, valid EPUB 3 (`mimetype`, `META-INF/container.xml`, an OPF manifest, a nav document, one XHTML chapter). Images are packaged as real zip entries (`OEBPS/images/...`) declared in the OPF manifest — not `data:` URIs, which most EPUB reading systems do not render. Generated XHTML is sanitized for strict-XML validity (stray `&`, unvalued boolean attributes, void-element self-closing) and paragraphs containing nested block content (e.g. an image's wrapping `<div>`) are promoted to `<div>` so the markup is valid against the HTML5 content model as well as being well-formed XML — a `<p>` containing a `<div>` is silently auto-corrected by browsers but causes strict-XML EPUB readers to drop the block instead of erroring.
  - Requires `extractAttachments: true` on the parse step to embed images when converting to/from EPUB; `OfficeConverter.convert()` sets this automatically.
- **GFM Task Lists**: `- [x] Done` / `- [ ] Todo` now round-trip through `ListMetadata.isTask`/`.checked` across Markdown and HTML (`<ul data-type="taskList"><li data-checked="true">`).
- **Admonitions / Alerts**: New `admonition` AST node type. Parses both GitHub (`> [!NOTE]`) and GLFM (`:::note ... :::`) syntax; always generates the GitHub blockquote form. HTML round-trips via `<div class="admonition admonition-note" data-type="note">`.
- **HTML Round-Trip Fidelity**:
  - Image size/alignment (`ImageMetadata.width`/`.align`) now read from `data-width`/`data-align`/inline `style="width:..."` on `<img>` (previously write-only).
  - Table alignment (`TableMetadata.align`) now read from `data-align` on `<table>` (previously write-only).
  - **Merged cells**: `HtmlParser`'s `<td>`/`<th>` handling now reads `colspan`/`rowspan` into `CellMetadata` — previously every merged cell silently collapsed to 1×1 on an HTML save→reload cycle.
  - **YouTube embeds**: New `embed` AST node type (`EmbedMetadata`). Round-trips `<div data-youtube-video="ID" data-width="..." data-align="...">` through HTML; Markdown falls back to the raw HTML block or a plain link.
- **Frontmatter Arrays**: Markdown frontmatter values written as a flow array (`tags: [a, b]`) or a JSON array (`tags: ["a","b"]`) now parse into real arrays in `customProperties`/`nativeProperties` instead of a literal string, with no new YAML dependency.
- **Footnotes**: Real `[^id]` inline references and `[^id]: definition` blocks now parse into `note` nodes keyed by id; the generator emits the same syntax instead of a `> **Footnote:**` blockquote.
- **Definition Lists & Abbreviations** (Markdown Extra): `Term\n: Definition` blocks parse into new `definitionList`/`definitionTerm`/`definitionDescription` node types; `*[HTML]: Hypertext Markup Language`-style abbreviation definitions populate `TextMetadata.abbreviationTitle`.
- **Attribute Lists** (Pandoc-style): `{width=50% .centered}` immediately after an image or table folds into `ImageMetadata.width`/`.align` and `TableMetadata.align`.
- **Citations**: `[@citekey]` inline citation syntax populates `TextMetadata.citationKey`.
- **Wikilinks**: Obsidian-style `[[Page]]` / `[[Page|Alias]]` populates `TextMetadata.wikilink` plus `.link`/`.linkType`.
- **MDX Import Stripping**: `<Component prop="x">...</Component>` and self-closing JSX tags are stripped on Markdown import (parse-only; officeParser never authors JSX back into Markdown).
- **Math Tokenisation**: Inline `$E=mc^2$` and block `$$...$$` LaTeX now tokenise into `TextMetadata.math` (`'inline' | 'block'`) instead of passing through as literal text.
- **Granular HTML Envelope Control (`standalone`)**: `HtmlGeneratorConfig.standalone` now accepts a `StandaloneConfig` object in addition to its existing `boolean`. The boolean conflated three unrelated decisions (document shell, CSS delivery, script/injection emission) into one flag and emitted a *global, unscoped* stylesheet whenever `standalone: false` was combined with a fragment embedded in a host page. The object splits these into independently-controllable fields — `document`, `metaTags`, `styles` (`'full' | 'scoped' | 'none'`), `scripts`, `headInjections`, `bodyInjections` — each defaulting to its "on" (standalone) value when omitted, so `{ document: false }` alone yields a fully-styled fragment with just the `<html>` shell removed. New `styles: 'scoped'` mode wraps the built-in stylesheet in a CSS `@scope` block so it cannot leak onto a host page's own elements (requires Chrome 118+, Safari 17.4+, or Firefox 128+). `bodyInjections` (unlike `headInjections`) now applies even to a bare content fragment, fixing an asymmetry where `injections.bodyStart`/`bodyEnd` were silently dropped outside standalone mode. `EpubGenerator` (which renders through `HtmlGenerator` with `standalone: false`) now gets a genuinely style-less, script-less fragment for free, simplifying its own XHTML sanitization.

### Changed
- `HtmlGenerator`'s footnotes section now emits `data-footnotes=""` (an explicit empty value) instead of a bare `data-footnotes` attribute, so the markup is valid XHTML as well as HTML.
- **Behavior change — `standalone: false`**: previously emitted an HTML fragment containing a global, unscoped `<style>` block (leaking onto any host page it was embedded in). It now emits a genuinely bare fragment with no `<style>`/`<script>` at all, matching the new "every envelope part off" semantics. **The old output is not lost — it moved from `false` to `{ document: false }`:** because an object's omitted fields each default to their "on" value, `standalone: { document: false }` keeps the full (global, unscoped) stylesheet and the spreadsheet script while dropping only the document shell — reproducing the previous `standalone: false` output byte-for-byte in the common case (the sole difference being that `injections.bodyStart`/`bodyEnd` now apply to the fragment instead of being silently dropped). Callers that want the old styled fragment should pass `{ document: false }`; those that want a leak-free styled fragment can pass `{ document: false, styles: 'scoped' }`.

### Fixed
- **Centralized Output Sanitization**: Added `src/utils/sanitize.ts` as the single source of truth for escaping AST-derived (untrusted document) text in every generator's output context, closing several injection gaps: HTML/XHTML text and attributes (`escapeHtml`/`escapeXml`), inline `<style>` CSS values (`sanitizeCssValue` — strips `url()`/`expression()`/`@import`/`javascript:` and CSS-breakout characters), `href`/`src` URLs (`sanitizeUrl`/`sanitizeImageUrl` — reject script-executing schemes, allow only `http(s)`/`mailto`/`tel`/relative/fragment, plus `data:image/*` for images), inline `<script>` JSON payloads (`serializeForInlineScript` — escapes `<`/`>` and U+2028/U+2029 so a document-supplied chart label can't close the script tag early), CSV cells (`csvSafeCell` — guards against formula/DDE injection per CWE-1236), RTF control words (`escapeRtf`), and Markdown text/URLs (`markdownEscapeText`/`sanitizeMarkdownUrl`). `CsvGenerator`, `EpubGenerator`, `HtmlGenerator`, `MarkdownGenerator`, `PdfGenerator`, and `RtfGenerator` all now delegate to these helpers instead of ad hoc per-generator escaping. Covered by a new `test/security/testSanitization.ts` regression suite (`npm run test:security`).
- **Zip Bomb Protection**: `extractFiles` (`src/utils/zipUtils.ts`) now decompresses via `fflate`'s streaming `Unzip`/`UnzipInflate` and caps `decompressionLimits.maxUncompressedBytes` against the *actual* inflated byte count as it streams in, instead of the ZIP header's declared (and attacker-controlled) `originalSize` — a crafted archive can understate that field and still inflate to gigabytes under the old declared-size check.
- **Denial-of-Service Hardening**:
  - `HtmlParser`'s tree builder no longer re-scans/re-lowercases the whole remaining document for every tag or `<script>`/`<style>` close tag (was `O(n²)` on documents with many tags); `parseNode` recursion is now capped at depth 1000, throwing the new `OfficeErrorType.MAX_NESTING_DEPTH_EXCEEDED` instead of overflowing the call stack on a maliciously deep element tree.
  - `MarkdownParser`'s MDX-unwrap fixed-point loop is capped at 100 passes, bounding the cost of a pathologically deep `<A><A>...</A></A>` input.
- **SSRF Hardening (PDF generation)**: `PdfGenerator`'s Puppeteer page now intercepts every network request and aborts anything that isn't an inline `data:`/`blob:` URI or the configured `htmlConfig.chartJsSrc` host — previously, rendering a document containing an external image or stylesheet URL would let Puppeteer fetch it from the server, which could reach internal services or a cloud metadata endpoint (`169.254.169.254`). A warning is emitted when a resource is blocked.
- **PDF Parsing Hardening**: `PdfParser` now passes `isEvalSupported: false` to `pdf.js`, preventing its font/CMap fast-path from compiling attacker-controlled PDF content via `new Function`.
- **Markdown Round-Trip**: Standalone bookmark-anchor blocks (`<a id="x"></a>` on their own line, emitted by `MarkdownGenerator` just before a heading/paragraph) and table cells using the `<div style="text-align: X">` alignment fallback are now correctly folded back into `anchorIds` / cell alignment on re-parse, instead of surviving as escaped literal text on a save→reload cycle.

## [7.2.3] - 2026-06-28
### Added
- **Slim Browser Bundles**: Introduced `officeparser.browser.slim.mjs` and `officeparser.browser.slim.iife.js` bundles along with types `officeparser.browser.slim.d.ts`. In the slim bundles, `tesseract.js` is stubbed out entirely and default CDN URLs for PDF workers and Chart.js are removed, making the library fully compliant with strict environments like Chrome/Edge Manifest V3 extensions where remotely hosted code is prohibited.
- **MathML Formula Support (ODF)**: Added parsing and extraction for MathML formulas in OpenOffice/LibreOffice documents (`.odt`, `.odp`, `.ods`), handling them at both the block level and inline level.

### Changed
- **Dependency Upgrades**:
  - Upgraded `pdfjs-dist` from `5.6.205` to `6.1.200` for optimized rendering performance, modernized Node.js compatibility, and security CVE mitigations.
  - Upgraded `fflate` from `^0.8.2` to `^0.8.3` to resolve Zip64 over-read bugs and improve large archive parsing stability.

## [7.2.2] - 2026-06-26
### Added
- **Configurable Decompression Limits**: Introduced a unified `decompressionLimits` configuration object to `OfficeParserConfig` to customize extraction checks during ZIP decompression (preventing excessive resource consumption). Configurable parameters:
    - `maxUncompressedBytes` (default: 512 MB)
    - `maxZipEntries` (default: 10,000 entries)
- **Centralized ZIP Exception Mapping**: Added new standardized error enums (`ZIP_ENTRY_COUNT_LIMIT_EXCEEDED`, `ZIP_ENTRY_INVALID_SIZE`, `ZIP_SIZE_LIMIT_EXCEEDED`, `EMBEDDING_TIMEOUT`) to `OfficeErrorType` and mapped all extraction limit rejections to these typed errors.

### Fixed
- **HTML Generator Attribute Escaping**: Sanitized and escaped HTML element attributes (such as the `src` attribute of generated image elements) inside the HTML generator to ensure structural integrity and correct document formatting.

## [7.2.1] - 2026-06-07
### Added
- **CLI Overhaul**: Rewrote CLI option parsing to support nested options, bare flags, and space-separated values; fixed positional-argument swallowing for bare CLI options.
- **CLI Test Suite**: Added a dedicated CLI test suite (`test/cli/testCli.ts`) and browser integration tests (`test/testIntegration.js`).
- **`FORMAT_UNSUPPORTED` Error Type**: Added new `FORMAT_UNSUPPORTED` generator error to `OfficeErrorType` for cleaner format-mismatch signalling.
- **Binary Build Script**: Added `scripts/build-binaries.js` and `scripts/browser-shims.js` for standalone binary and browser bundle builds.

### Changed
- **Default `preserveLayout`**: Changed the default value of `preserveLayout` to `true`.

### Fixed
- **Note Preservation in All Generators**: All generators now correctly collect and render footnotes/endnotes at the end of the document; slide notes are rendered inline. `MarkdownGenerator` additionally fixes note loss during text-node merging in `optimizeNodes`.
- **PDF Worker Path Resolution**: Resolved dynamic module load errors and PDF worker path resolution in ESM/bundled contexts.
- **Comment Sanitisation in Source Code**: Removed the word `fetch` from inline code comments in `OfficeParser.ts`, `PdfParser.ts`, and `types.ts` to prevent automated scanners from falsely flagging the repository as one that directly accesses the internet.

## [7.2.0] - 2026-06-04
### Added
- **Parser Enhancements**:
    - **Comments Extraction (DOCX, XLSX, PPTX)**: Parser now extracts inline comments/annotations from Word, Excel, and PowerPoint documents. Comments are attached to their target node via `node.comments` and use the new `CommentMetadata` type (carrying `author`, `initials`, `date`, `commentId`). Controlled by the new `ignoreComments` config flag.
    - **Headers & Footers Extraction (DOCX)**: Word document headers and footers are now parsed into the new `ast.auxiliary.headers` / `ast.auxiliary.footers` arrays (of type `OfficeAuxiliaryContent`). Controlled by the new `ignoreHeadersAndFooters` flag.
    - **Slide Masters Extraction (PPTX)**: PowerPoint slide masters are now extracted into `ast.auxiliary.slideMasters` as `slideMaster` nodes with `SlideMetadata`. Controlled by the new `ignoreSlideMasters` flag.
    - **Cell Background Color (DOCX/XLSX)**: `CellMetadata.backgroundColor` now populated from `<w:shd>` fills in DOCX and equivalent elements in XLSX.
- **HTML Generator Enhancements**:
    - **Config Additions**: `containerWidth`, `customCss`, and `injections` (`headStart`, `headEnd`, `bodyStart`, `bodyEnd`) added to `HtmlGeneratorConfig`.
- **AST & Metadata Extensions**:
    - **`OfficeAuxiliaryContent` Interface**: New root-level `auxiliary` property on `OfficeParserAST` for out-of-band layout/template elements.
    - **`OfficeMetadata` Extensions**: `keywords` and `nativeProperties` fields added — `nativeProperties` exposes all raw format-specific metadata (e.g. all `<meta>` tags in HTML, `app.xml` properties in DOCX, XMP dicts in PDF).
    - **`NoteMetadata.slideNumber`**: Slide notes (`note` nodes from PPTX) now carry `metadata.slideNumber`.
- **Types Improvements**:
    - **`TextAlignment`**: Extracted as a standalone type to replace inline string unions across multiple formatting interfaces.
    - **`ConversionResult<D>`**: Removed the universal type fallback, forcing the generic interface to strictly map to the destination type requested.
    - **Metadata Typing**: Added `CommentMetadata`, `HeaderFooterMetadata`, and `TableMetadata` to strongly type newly supported document structures.
    - **`OfficeContentNodeType`**: Expanded to explicitly include `'header'`, `'footer'`, and `'slideMaster'`.
    - **`BaseContentNode`**: Extracted common node properties into a shared interface to reduce duplication.
    - **Configuration Deep-Merging**: `resolveGeneratorConfig` now recursively deep-merges nested configuration objects (like `injections`) instead of shallow-overwriting them.
    - **Error Types**: Added `INVALID_CONTAINER_WIDTH` to `OfficeWarningType`.

### Changed
- **Parser Enhancements**:
    - **Notes Placement (RTF, DOCX, ODT, ODP)**: Notes (footnotes, endnotes, slide speaker notes) are now structurally attached via `node.notes[]` to their closest preceding sibling node, rather than being appended to the flat `content` array. The `putNotesAtLast` flag is **deprecated** (notes are no longer re-ordered; use `node.notes` for access).
    - **Slide Notes (PPTX, ODP)**: Slide notes are now attached to their parent `slide` node via `slideNode.notes[]` instead of being inserted as top-level `note` nodes in `content`.
- **Types Improvements**:
    - **`OfficeContentNode` is now a Discriminated Union Type**: Previously an interface with a generic `metadata?: ContentMetadata`, it is now a union type (`BaseContentNode & (| { type: 'slide'; metadata?: SlideMetadata } | ...)`) providing precise, compile-time type narrowing per `node.type`.

### Deprecated
- **`putNotesAtLast`**: Notes are now structurally attached to specific nodes via `node.notes`. This flag no longer has an effect. It will be removed in a future major version.

### Fixed
- **RTF Notes Inline Placement**: Footnotes and endnotes in RTF documents are now correctly attached inline to their preceding text node (via `node.notes`), resolving incorrect end-of-document appending regardless of `putNotesAtLast`.
- **Generator Sub-Config Merging**: Fixed shallow-merge bug where providing partial `htmlConfig` (e.g., only `standalone`) would discard previously set defaults for other keys in nested objects like `injections`.

## [7.1.0] - 2026-05-25
### Added
- **Cancellation Support (AbortSignal)**: Enabled passing an `abortSignal` in `OfficeParserConfig` and `OcrConfig` to gracefully interrupt document loading, parsing loops, and worker execution.
- **Consolidated OCR Timeouts**: Grouped OCR-specific timeouts under a unified `timeout` object (`workerLoad`, `recognition`, `autoTerminate` in `OcrTimeoutConfig`) for reliable limit enforcement.
- **Visualizer Upgrades**: Added a fullscreen preview modal, dynamic scroll forwarding via `ResizeObserver`, and integrated Puppeteer-driven layout and scroll verification tests.
- **ESLint Enforcements**: Added rules to restrict catch blocks from passing unhandled `AbortError` to `getWrappedError`, and ban direct error string literals in `new Error()` and `new DOMException()`.

### Fixed
- **XLSX Entity Decoding**: Corrected matching of `inlineStr` cells with XML attributes and resolved decimal, hex, and named XML entities during spreadsheet parsing.
- **Worker/Thread Cleanup**: Terminated and evicted stalled or timed-out OCR workers to prevent memory leaks and dangling background threads.
- **ESM CSP Compliance**: Replaced standard dynamic module loading via `new Function()` with direct dynamic `import()` to comply with strict Content Security Policies.

## [7.0.3] - 2026-05-15
### Added
- **Native Uint8Array Support**: Added `Uint8Array` as a first-class input format for `parseOffice` and `convert`, improving browser-side binary data handling.
- **Visualizer Refactor**: Introduced a schema-driven configuration engine and a dual-pass RTF previewer (`AST -> RTF -> AST -> HTML`) for high-fidelity verification.

### Changed
- **Visualizer UI/UX**: Standardized navbar interactivity, optimized responsive breakpoints (1200px), and unified global layout symmetry.
- **Parser Core**: Refined `ArrayBuffer` logic and improved `fs`/`path` shimming for better compatibility with modern bundlers.
- **Telemetry**: Integrated `onWarning` accumulation into the `OfficeParserAST` to preserve parser-phase issues throughout the generation pipeline.
- **Generator API**: Enforced a strict return contract (`string | false | void`) for `onNode` callbacks to ensure deterministic AST transformations.

### Fixed
- **RTF Generator Fidelity**: Restored manual indentation for lists; implemented `\cellx` table layouts and `\pict` binary image embedding.
- **Visualizer Layout**: Resolved `ReferenceError` regressions and cross-zoom layout drift on high-DPI displays.

## [7.0.0] - 2026-05-12
### Added
- **OfficeConverter**: A high-level, streamlined API (`convert`) for one-step document transformations with automatic parser/generator configuration sync.
- **OfficeGenerator**: A comprehensive conversion engine for document ASTs, enabling high-fidelity output in `Markdown`, `HTML`, `CSV`, `RTF`, and `Text`.
- **RAG Chunking Suite**: Native, metadata-aware document splitting optimized for Vector Databases.
    - Supports `fixed-size` (recursive), `document-structure`, and `semantic` strategies.
    - Features robust sentence boundary detection (abbreviations, Japanese punctuation) and deterministic HTML output.
- **Parser Extensions**: Added native support for parsing `CSV`, `HTML`, and `Markdown` files into the unified Office AST.
- **StyleMapper Engine**: A semantic translation layer for preserving document styles across formats.
    - Supports a robust DSL with quoted attributes, commas, and regex-based (`~=`) matching.
    - Introduced **Structured Style Mappings** for type-safe, object-based configuration.
- **Conversion Results API**: Unified `ConversionResult` and `ConversionMessage` interfaces for consistent, structured feedback across all tasks.
- **Standardized Error System**: Introduced `OfficeErrorType` and `OfficeWarningType` enums for predictable and typed error/warning handling.
- **Link Filtering**: Added granular controls `ignoreInternalLinks` to prune noisy document navigation and bookmarks from the AST.

### Changed
- **Unified Office AST**: Redesigned the core document representation to support complex tables, nested lists, and format-specific metadata across all parsers.
- **Performance Optimizations**:
    - **RTF Parser**: Rewritten string accumulation logic to resolve $O(n^2)$ bottlenecks in large documents.
    - **OpenOffice Parser**: Improved XML pre-parsing and style caching, yielding significant speedups (up to 23x for ODP).
    - **Excel Parser**: Replaced global regex matching with `matchAll` iteration to significantly reduce memory overhead and prevent execution stalls on large, sparse spreadsheets (Fixed #91).
- **Browser Build**: Optimized the bundling process to suppress dynamic import warnings in browser environments by injecting ignore comments into dynamic imports.
- **Configuration Engine**: Migrated to a strictly-typed architecture using `DeepRequired` to ensure robust defaults and eliminate runtime configuration errors.
- **CLI Enhancements**: Expanded CLI capabilities with `--format`, `--output`, `--verbose` (for stack traces), and specialized flags for XML serialization.
- **CSV API**: Standardized single-sheet exports to return plain strings for better ergonomics.

### Fixed
- **DOCX Table Fidelity**: Implemented support for vertical cell merging (`w:vMerge`) and horizontal spanning (`w:gridSpan`) in Word documents.
- **Document Anchors**: Added preservation of bookmarks and anchor IDs during Word document parsing.
- **Error Reporting**: Standardized reporting for OCR and chart data extraction failures.
- **Excel Coordinate Indexing**: Resolved a bug where self-closing XML tags caused incorrect row/column metadata indexing and added support for multi-letter column coordinates (e.g., AA, XFD).

## [6.1.1] - 2026-04-28
### Added
- **Break Nodes (DOCX)**: Comprehensive support for `w:br`, `w:cr`, and `w:lastRenderedPageBreak` nodes in Word documents.
- **Indentation Metadata (DOCX)**: Extraction of `<w:ind>` properties for precise paragraph layout analysis.
- **Field Extraction (PPTX)**: Support for `<a:fld>` elements, ensuring slide numbers and other dynamic fields are captured.

### Fixed
- **Soft Break Handling**: Standardized splitting of list items on soft breaks (`Shift+Enter`) across PPTX and ODP, treating interruptions as independent paragraph nodes.
- **List Indexing (ODP)**: Re-engineered stateful index tracking for nested lists in ODP to ensure sequential continuity.
- **Excel Multi-line Parsing**: Resolved failures in XLSX parsing for cells containing complex multi-line content.
- **RTF Encoding**: Implemented robust byte-buffering and character decoding to resolve smart quote and double-quote dropouts.
- **XLSX Fidelity**: Fixed case-sensitivity issues in regex for `inlineStr` cell types.
- **Security & Stability**: Upgraded `@xmldom/xmldom` to `0.9.10` to address upstream vulnerabilities.

### Changed
- **PPTX Engine**: Migrated to an iterative child-processing model for paragraphs to guarantee correct content ordering and support for all inline elements.
- **Documentation**: Updated OpenGraph metadata and project specs for better social sharing and developer clarity.

## [6.1.0] - 2026-04-14
### Added
- **OCR Scheduler**: Intelligent worker pool that optimizes Tesseract lifecycle across parallel requests.
- **Custom Properties**: Support for extracting document metadata across OOXML, ODF, and PDF formats.
- **Sponsorship**: Integrated `funding.json` manifest and GitHub Sponsors support.
- **Governance**: Added `.editorconfig`, `.gitattributes`, and `SUPPORT.md`.

### Changed
- **Core Engine**: Replaced legacy zip extraction with `fflate` for significant performance gains and robust browser/edge compatibility.
- **Module System**: Full native ESM support with `Node16` resolution and verified browser bundles (Vite/Angular compatible).
- **Format Refinements**: Hierarchical PDF coordinate alignment and ODT/RTF list parsing stability.

## [6.0.0] - 2025-12-29
### Added
- **Major Overhaul**: Transitioned from simple text extraction to a rich **Abstract Syntax Tree (AST)** output.
- **Structured Output**: Access hierarchical document structure (paragraphs, headings, tables, lists, etc.).
- **Rich Metadata**: Extracted document properties (author, title, creation date).
- **Enhanced Formatting**: Support for bold, italic, colors, fonts, alignment, etc.
- **Attachment Handling**: Extract images, charts, and embedded files as Base64.
- **OCR Integration**: Optional OCR for images using Tesseract.js.
- **RTF Support**: Added full support for Rich Text Format files.
- **TypeScript**: Full TypeScript support with detailed interfaces and improved type definitions.

### Changed
- **Simplified API**: Transitioned to the unified `parseOffice` for all parsing needs (returns a Promise).

## [5.1.1] - 2024-11-12
### Added
- Added `ArrayBuffer` as a type of file input. 
- Introduced browser bundle generation, exposing the `officeParser` namespace for direct browser usage.

## [5.0.0] - 2024-10-21
### Added
- Replaced `decompress` with `yauzl` for zip extraction. 
- Migrated to in-memory extraction (no longer writing to disk).
- Removed config flags related to extracted files and added flags for CLI execution.

## [4.2.0] - 2024-10-15
### Added
- Fixed race conditions when deleting temp files during parallel execution.
- Resolved errors occurring when multiple executions were made without waiting for the previous one to finish.
- Upgraded project dependencies.

## [4.1.2] - 2024-10-13
### Fixed
- Fixed text parsing from XLSX files containing no shared strings file or using `inlineStr` based strings.

## [4.1.1] - 2024-05-06
### Changed
- Replaced `pdf-parse` with a native `pdf.js` implementation for more robust PDF analysis.
- Added `pdfjs-dist` build as a local library.

## [4.0.5] - 2023-11-25
### Fixed
- Improved error catching during file parsing, specifically post-decompression.
- Fixed parallel parsing issues caused by timestamp-only file naming.

## [4.0.0] - 2023-10-24
### Added
- **Revamped Content Parsing**: Resolved content ordering issues (e.g., table positioning in Word files).
- Added `config` object as an argument for `parseOffice` to set delimiters and other configurations.
- Added initial support for parsing PDF files using the `pdf-parse` library.
- Removed support for individual file parsing functions in favor of a unified approach.

## [3.3.0] - 2023-04-26
### Added
- Added support for file buffers as an argument for `filepath` in `parseOffice` and `parseOfficeAsync`.

## [3.2.0] - 2023-04-07
### Added
- Added comprehensive typings to methods for enhanced TypeScript support.

## [3.1.4] - 2022-12-28
### Added
- Added Command Line Interface (CLI) functionality to use `officeParser` directly from the terminal.

## [3.0.0] - 2022-12-10
### Added
- Resolved memory leak issues and bugs related to Open Document (ODF) parsing.
- Improved global error handling.

## [2.3.0] - 2021-11-21
### Added
- Implemented Promise-based wrappers for existing callback functions.

## [2.2.2] - 2020-06-01
### Added
- Added error handling and configurable `console.log` methods.
- Maintained full backward compatibility.

## [2.1.1] - 2019-06-17
### Added
- Added configuration to change the location for decompressing office files (useful for restricted write access environments).

## [2.0.3] - 2019-04-30
### Fixed
- Fixed case-sensitivity bug for file extensions; capital lettered extensions are now supported.

## [2.0.0] - 2019-04-23
### Added
- Added support for Open Office files (`*.odt`, `*.odp`, `*.ods`) through `parseOffice`.
- Created the dedicated `parseOpenOffice` method.
- Added feature to automatically delete the generated dist folder after function callback.

## [1.3.0] - 2019-04-22
### Added
- Introduced the `parseOffice` method to unify parsing across different extensions.
- Added file extension validations.
- Resolved errors for Excel files lacking drawing elements.

## [1.2.0] - 2019-04-19
### Added
- Added support for `*.xlsx` (Excel) files.

## [1.1.2] - 2019-04-18
### Added
- **Initial Release**: Added support for `*.pptx` and `*.docx` files.
