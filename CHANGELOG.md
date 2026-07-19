# Changelog

All notable changes to `officeParser` are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [7.4.0] - 2026-07-16
### Added
- **Markdown output dialect**: `MdGeneratorConfig.dialect` targets a specific real-world Markdown
  flavor - `'github'`, `'gitlab'`, `'obsidian'`, `'pandoc'`, strict `'commonmark'`, or officeParser's
  own kitchen-sink `'extended'` (the default, byte-identical to this library's prior output). Each
  preset picks the right native syntax per feature: admonitions (`> [!NOTE]` on GitHub/Obsidian,
  `:::note` on GitLab, `::: {.note}` on Pandoc, plain blockquote elsewhere), definition lists,
  footnotes, citations, wikilinks, math delimiters, attribute lists, strikethrough, and forces HTML
  `<table>` under `commonmark` (which has no native table syntax). Pass an object instead of a
  preset name for granular control over any of the above, or to override bullet/ordered-list
  markers and emphasis style (`*`/`_`) independently of dialect - see `MarkdownDialectConfig`.
- **`MdGeneratorConfig.fallbackToHtml` now also accepts a `FallbackToHtmlConfig` object**, splitting
  the single flag into independently-controllable parts: text formatting (underline/sub/sup),
  alignment, anchors, tables (nested/merged-cell), embeds, and cell line breaks. The boolean form is
  unchanged and still applies uniformly to every part.
- **Footnotes now degrade gracefully when disabled** (`dialect: { footnotes: false }`, or under the
  `commonmark` preset): instead of the `[^id]` syntax, note content is inlined as a parenthetical
  right at the reference point, so it's never silently dropped.
- **Per-input-format parser configuration** (`OfficeParserConfig.mdParserConfig`), mirroring the
  generator side's existing per-destination-format config sub-objects. Currently reserved for
  future use.
- **Admonition source-syntax provenance**: `AdmonitionMetadata.sourceSyntax` (`'github' | 'gitlab'`)
  now records which concrete syntax produced the node, always populated by the parser - useful for
  round-trip-aware tooling that wants to preserve a document's original dialect on re-save.
- **`GeneratorConfig.metadataOverrides`** sets the metadata embedded in generated output without
  mutating the parsed AST. Named fields mirror `OfficeMetadata` (`title`, `author`, `description`,
  `subject`, `keywords`, `lastModifiedBy`, `created`, `modified`, `language`) plus a `custom` bucket
  for caller-defined key/value pairs; `custom` is kept separate from the named fields so a typo like
  `titel` is a compile error rather than a silently-created custom entry. Merged **per field** over
  `ast.metadata`, so overriding one field leaves the rest of the parsed metadata intact, and applied
  uniformly across every generator that embeds metadata (HTML `<meta>`, EPUB OPF, Markdown
  frontmatter, RTF `\info`, and the `renderMetadata` header in text/CSV).
  - **`modified` is what makes generation reproducible.** EPUB embeds it as both the required
    `dcterms:modified` property and each zip entry's mtime, so pinning it yields byte-identical
    archives across runs. When unset it falls back to the source document's own
    `ast.metadata.modified` before the current time, so most documents are now reproducible with no
    configuration at all.
  - EPUB's OPF and RTF's `\info` have fixed metadata vocabularies with no place for caller-defined
    keys; those generators report unrepresentable `custom` entries through `onWarning`
    (`OfficeWarningType.METADATA_NOT_REPRESENTABLE`) rather than dropping them silently.

### Fixed
- **Markdown generation could carry an unescaped HTML tag out of a hostile document.** Text nodes
  were escaped, but seven other constructs interpolated their content directly: inline and block
  math, wikilink page/alias, citation keys, footnote ids, abbreviation definitions, image
  attribute lists, and admonition types/titles. A payload correctly neutralized in a paragraph
  survived a full parse to generate cycle verbatim in any of them. Each sink now escapes or
  allowlists according to what round-trips for that specific value: entity-encoding where the
  value returns to a text node (which the parser entity-decodes), and a character allowlist where
  it returns to `metadata.*` or a `code` node (which it does not). The `[^id]` fix lives in
  `BaseGenerator.getFootnoteKey`, so it covers HTML output too. Also routes CR-only code text to
  the fenced-block branch, where a renderer normalizing `\r` would otherwise kill an inline span
  and expose the remainder as raw Markdown.
- **`styleMap` class lists and attribute names were emitted without escaping or validation.** The
  spreadsheet row and sheet paths rebuilt their `class` attribute from the raw mapping array
  rather than the escaped value every other node type uses, and both `styleMap` attribute loops
  escaped the value while interpolating the *name* unchecked - which no amount of value escaping
  can compensate for, since a name containing a quote closes the attribute and opens another. The
  attribute-name check is now a shared `isSafeHtmlAttributeName` rather than a literal repeated at
  each site.
- **Config resolution could be used to poison `Object.prototype`.** The recursive generator-config
  merge walked caller input with an unguarded `for...in`, so a config that arrived via
  `JSON.parse` (an object literal cannot express this; JSON can) reached `Object.prototype` and
  corrupted every object in the process. The two `Object.assign` sites had a narrower form of the
  same problem: they cannot pollute globally, but the `__proto__` setter replaced the returned
  config's prototype, so it silently inherited attacker-chosen properties.
- **`sanitizeCssValue` could be bypassed with CSS backslash escapes.** Control characters and CSS
  comments were stripped before the `url()`/`expression()` denylist test, but backslash escapes
  were stripped after it, so `u\rl(http://evil)` was validated as safe and then reassembled into a
  live `url()`. Inert in practice only because every call site pairs the value with a property
  where `url()` does nothing - it would have become a live SSRF the moment anyone used
  `background-image`, `cursor` or `mask`. The regression test covering this path was itself
  vacuous: it asserted against an empty string, because the payload it used caused no `style`
  attribute to be emitted at all.
- **RTF hyperlinks had no scheme allowlist.** `escapeRtf` neutralizes the field metacharacters but
  says nothing about the target, so RTF was the only generator that would emit
  `HYPERLINK "javascript:..."`, `file:///...` or a UNC `\\host\share` - the latter two being
  phishing and NTLM-credential-leak vectors in Word rather than rendering quirks. **This changes
  behaviour for non-hostile input:** an RTF document generated from content with intranet
  `file://` or UNC links loses those links. The link *text* is preserved (rejection degrades to
  bare text, matching how HTML and Markdown already degrade), so no content disappears - only the
  clickable target.
- **The HTML parser's nesting-depth guard could never fire.** It tripped at depth > 1000, but the
  call stack overflows at roughly 800 (and nondeterministically), so the typed
  `MAX_NESTING_DEPTH_EXCEEDED` error was never produced. Failure was already graceful, so this was
  a dead guard rather than a denial of service. Lowered to 256, chosen to hold across engines
  rather than tuned to one measurement; real documents nest orders of magnitude shallower.
- **`csvSafeCell` tested its formula trigger against the untrimmed value** while the numeric
  exemption tested the trimmed one, so `" =1+1"` was emitted unprefixed. Most spreadsheet
  applications treat a leading-space cell as text, so this is defence in depth rather than a
  demonstrated bypass.
- **`.to('text')`/`.to('md')` silently stripped genuine document whitespace.** Both generators
  unconditionally called `.trim()` on the entire accumulated output before returning, which not
  only cleaned up the generator's own trailing-newline join artifacts (its original purpose) but
  also destroyed real content: an intentionally-indented opening line, or trailing spaces on the
  document's last line. This meant migrating from the deprecated synchronous `ast.toText()` (which
  never trimmed) to its documented replacement `ast.to('text')` could silently change output.
  Fixed by stripping only a run of the exact block-separator sequence at either end - the one
  demonstrable generator artifact - and nothing else. This also fixed a related, previously-latent
  leading-newline artifact: a document consisting of only a table (`.to('text')` with
  `preserveLayout`, or `.to('md')` under a dialect that forces tables to HTML, e.g. `commonmark`)
  would leak a spurious leading blank line, since the table renderer's own separator convention had
  nothing to separate from.
- **EPUB generation was not reproducible.** Generating the same AST twice produced archives that
  differed byte-for-byte, from two independent wall-clock sources: `dcterms:modified` in the OPF,
  and the mtime `fflate` stamps onto every zip entry when none is supplied. The second source is
  easy to miss, because DOS zip timestamps have two-second granularity - back-to-back generation
  looks stable and only a longer gap reveals it. Both are now derived from a single resolved
  instant (`metadataOverrides.modified`, else `ast.metadata.modified`, else now), and a date outside zip's
  representable 1980-2099 window clamps instead of throwing. This also un-masks real diffs: while
  every `gen.*.to.epub.epub` baseline churned on each run, genuine EPUB content changes were
  indistinguishable from timestamp noise and got reverted along with it.
- **The `renderMetadata` header in plain-text output could be forged.** Metadata values were
  interpolated into the `Key: value` block verbatim, so a line break in a document's title rendered
  as additional fields - a title of `Real\nAuthor: Attacker` produced an `Author:` line the document
  never had, plus a second separator rule. No code execution, but consumers parsing that block would
  believe it. Line breaks are now folded to spaces, matching how `CsvGenerator` already guarded its
  `#` comment block, and a malformed `created` date is omitted rather than rendered as the literal
  `Invalid Date`.
- **`.to('text')` silently dropped chart data and CSV comments.** `TextGenerator` rendered only a
  node's children, so any node carrying its content in `text` with no children vanished entirely:
  `chart` nodes (ODP/ODS/PPTX/XLSX), whose whole data series lives in `text`, and CSV `comment`
  nodes. The deprecated `ast.toText()` kept both, making this a regression for anyone following the
  documented migration. Fixed generically - the fallback now prefers rendered children but falls
  back to the node's own `text`, matching what every parser's own `toTextSync` already does, so it
  covers any future node of the same shape rather than just the two known today.
- **`.to('text')` concatenated adjacent table cells with no separator.** `ITEM` + `NEEDED` rendered
  as `ITEMNEEDED`, destroying the cell boundary, on every path that doesn't draw an aligned grid
  (`preserveLayout: false`, and `sheet`/`row`/`cell`, which never route through the table renderer).
  This was masked for XLSX/ODS only because their cell values happen to carry a trailing
  non-breaking space *in the source data* - the delimiter came from the document, not the generator
  - so formats without it (MD, HTML) collided outright. Cells are now tab-separated, but only when
  the cell doesn't already end in a line break, since DOCX/ODT/RTF wrap cell content in a paragraph
  that emits its own newline.

### Changed
- **`styleMap`'s `output.tag` now takes effect in HTML output.** It was written and then shadowed
  in every branch of `HtmlGenerator`, so it was silently ignored - while `MarkdownGenerator` and
  `RtfGenerator` both honoured it and the README documented it as working. A mapping of
  `{ tag: 'blockquote' }` produced a plain `<p>`. Restoring it required validation in the same
  change, because that shadowing was the only thing making a hostile tag inert: a tag name is
  interpolated into both the opening and closing tag, where no escaping applies. Only a known-safe
  set of block, heading and inline elements is accepted; anything else falls back to the node's
  default tag and emits an `INVALID_STYLE_MAP_TAG` warning. This also makes the built-in default
  style mappings live for HTML, so a paragraph carrying a `Heading N`/`Quote`/`Title` style name
  now maps to the corresponding element instead of remaining a `<p>`.
- **Node.js 18 and 20 caveat.** Some of the hardening above is calibrated against current Node.
  The clearest case is the parser's nesting-depth threshold, which is derived from stack-overflow
  behaviour measured on a modern runtime; on Node 18/20 the exact boundary may differ, though
  failure remains graceful in every case. Node 18 and 20 stay in the supported set for now and
  will be dropped at the next major version, along with the other deprecations already batched
  for it.
- **`TextGeneratorConfig.renderNotes`** (default `true`) controls whether the collected
  footnote/endnote section is appended to `.to('text')` output. Set it to `false`, alongside
  `includeImages: false`, for the leaner rendering the deprecated `ast.toText()` produced.
- **Documentation fix**: `TextGeneratorConfig.preserveLayout` was documented as defaulting to
  `false`; the actual default is and always was `true`. Only the doc comment was wrong.

## [7.3.1] - 2026-07-14
### Fixed
- Nested-list item numbering could leak across siblings (a second top-level item's nested children could continue a previous sibling's child numbering instead of restarting at 0).
- Parser/generator indentation-unit mismatch caused this generator's own output to double its nesting depth on re-parse; tab-indented items also failed to nest.
- Nested blockquotes (`> >`) only had one level of `>` stripped, leaking a literal `>` into text.
- `)`-style ordered lists (`1) item`) were not recognized, only the `.`-marker form.
- `----- trailing text` was misclassified as a horizontal rule.
- Table separator rows with short cells (`|-|-|`) were rejected.
- Numeric character/entity references decoded via `String.fromCodePoint` (`&#NNN;`/`&#xHHH;` in
  Markdown, and XML entities in `xmlUtils.ts` used by e.g. `ExcelParser`) would throw a `RangeError`
  and crash the parser on an out-of-range code point (`&#999999999;`); such references are now left
  as literal text instead.

### Added
- Backslash escapes for ASCII punctuation (`\*`, `` \` ``, `\[`, etc.).
- Reference-style links/images (`[text][ref]`, `[text][]`, `[text]`) resolved against `[ref]: url "title"` definitions.
- Underscore emphasis (`_em_`, `__strong__`) alongside `*`/`**`.
- Multi-backtick inline code spans (`` ``code with a ` backtick`` ``).
- Setext headings (`Heading\n===` / `Heading\n---`) alongside ATX.
- `<url>` autolinks.
- HTML entity and numeric character reference decoding (`&amp;`, `&#39;`, `&#x2764;`, etc.).
- Hard vs. soft line breaks (a trailing 2+ spaces or backslash now produces a real hard break).
- `~~~`-fenced and 4-space-indented code blocks.
- List-item continuation lines (a single indented line merges into the previous item).

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
