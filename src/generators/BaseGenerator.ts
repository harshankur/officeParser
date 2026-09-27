import { OfficeAttachment, OfficeIssue, ConversionResult, FullGeneratorConfig, GeneratorConfig, ImageMode, OfficeContentNode, OfficeMetadata, OfficeParserAST, OfficeWarningType, StructuredStyleMapping, UniversalGeneratorFormat } from '../types.js';
import { resolveGeneratorConfig } from '../utils/configUtils.js';
import { checkAbortSignal, getWarningMessage } from '../utils/errorUtils.js';
import { resolveImageMode } from '../utils/officeGenUtils.js';
import { StyleMapper } from '../utils/styleMapper.js';
import { isSourceComment } from '../utils/commentUtils.js';

/**
 * Base class for all document generators.
 * Provides common traversal logic and configuration handling.
 */
export abstract class BaseGenerator<D extends UniversalGeneratorFormat = UniversalGeneratorFormat> {
    protected config: FullGeneratorConfig;
    protected ast: OfficeParserAST;
    protected messages: OfficeIssue[] = [];
    protected styleMapper: StyleMapper;
    protected collectedNotes: OfficeContentNode[] = [];
    private readonly collectedNoteSet = new Set<OfficeContentNode>();
    private noteReferenceCounts: Map<OfficeContentNode, number> | undefined;

    /**
     * How many references the AST makes to `note` (a parser shares one note node among all the
     * references to it). A writer that writes a note at its reference writes it in full once and
     * refers back to it from the others, so a note referred to thousands of times is not written
     * thousands of times.
     */
    protected noteReferences(note: OfficeContentNode): number {
        if (!this.noteReferenceCounts) {
            const counts = new Map<OfficeContentNode, number>();
            const seen = new Set<OfficeContentNode>();
            const stack: OfficeContentNode[] = [...(this.ast?.content ?? [])];
            while (stack.length) {
                const node = stack.pop()!;
                if (!node || typeof node !== 'object') continue;
                for (const n of node.notes ?? []) counts.set(n, (counts.get(n) ?? 0) + 1);
                if (seen.has(node)) continue;
                seen.add(node);
                for (const list of [node.children, node.notes, node.comments]) if (list) for (const child of list) stack.push(child);
            }
            this.noteReferenceCounts = counts;
        }
        return this.noteReferenceCounts.get(note) ?? 0;
    }
    /** Lazily-built `name -> attachment` index for {@link getAttachment}. */
    private attachmentIndex?: Map<string, OfficeAttachment>;

    constructor(protected destination: D, ast: OfficeParserAST, config?: GeneratorConfig<D> | FullGeneratorConfig) {
        // Problems with the configuration itself are among the result's messages too.
        this.config = resolveGeneratorConfig(destination, ast.config, config, issue => this.messages.push(issue));
        this.ast = ast;
        this.styleMapper = new StyleMapper(this.config.styleMap, this.config.ignoreDefaultStyleMap, this.config);
    }

    /**
     * Resolves an attachment by name via a lazily-built index, replacing a linear
     * `ast.attachments.find(...)` per image/chart node. First-match-wins keeps the exact semantics of
     * `find`. Every generator resolves image/chart bytes this way, so a media-heavy document (for
     * example a fully inlined self-contained export) stays O(nodes + attachments) rather than
     * O(nodes x attachments).
     */
    protected getAttachment(name: string | undefined): OfficeAttachment | undefined {
        if (!name) return undefined;
        if (!this.attachmentIndex) {
            this.attachmentIndex = new Map();
            for (const a of this.ast.attachments || []) {
                if (a.name && !this.attachmentIndex.has(a.name)) this.attachmentIndex.set(a.name, a);
            }
        }
        return this.attachmentIndex.get(name);
    }

    /**
     * Resolves `config.includeImages` (a boolean, a CLI-provided `'true'`/`'false'` string, or an
     * {@link ImageMode}) to a single mode. Every generator's image handling must route through this
     * rather than a truthy check, since a mode string like `'none'` is truthy.
     */
    protected imageMode(): ImageMode {
        return resolveImageMode(this.config.includeImages);
    }

    /**
     * The document metadata a generator should write out: `ast.metadata` with
     * `config.metadataOverrides` applied on top, per field.
     *
     * Every generator must read metadata through here rather than touching `this.ast.metadata`
     * directly, so an override reaches all of them uniformly instead of one format at a time.
     *
     * Merged rather than replaced, so overriding one field doesn't blank the rest, and computed
     * fresh rather than cached on the AST: overrides are an output concern, and mutating
     * `ast.metadata` would leak one generation's settings into the next use of the same AST.
     * `custom` merges into `customProperties` so callers see one bucket regardless of origin.
     */
    protected get effectiveMetadata(): OfficeMetadata {
        const base = this.ast.metadata || {};
        const overrides = this.config.metadataOverrides;
        if (!overrides) return base;

        const { custom, language, ...named } = overrides;
        const merged: OfficeMetadata = { ...base };
        // Assign only the fields actually supplied - spreading `named` wholesale would write
        // `undefined` over parsed values for every field the caller left out.
        for (const [key, value] of Object.entries(named)) {
            if (value !== undefined) (merged as any)[key] = value;
        }
        // `language` lives in two places generators read from: the top-level `OfficeMetadata.language`
        // (what PdfParser sets and the DOCX/ODT generators read) and `nativeProperties.language` (what
        // EpubGenerator reads). Write the override to both so it reaches every generator.
        if (language !== undefined) {
            merged.language = language;
            merged.nativeProperties = { ...(base.nativeProperties || {}), language };
        }
        if (custom && Object.keys(custom).length > 0) {
            merged.customProperties = { ...(base.customProperties || {}), ...custom };
        }
        return merged;
    }

    /**
     * Reports caller-supplied `metadataOverrides.custom` entries that the destination format has
     * no way to represent (EPUB's OPF and RTF's `\info` both have fixed vocabularies).
     *
     * Warning rather than dropping silently: a caller who sets metadata and never sees it in the
     * output otherwise has no way to find out. Only `custom` keys are reported - the named fields
     * map onto something in every format that carries metadata at all.
     */
    protected warnUnrepresentableCustomMetadata(format: string): void {
        const custom = this.config.metadataOverrides?.custom;
        if (!custom) return;
        const keys = Object.keys(custom);
        if (keys.length === 0) return;
        this.warn(OfficeWarningType.METADATA_NOT_REPRESENTABLE, { keys, format });
    }

    /**
     * Retrieves the semantic mapping for a node, respecting the includeFormatting flag.
     * Per design requirements: Style mapping is bypassed if formatting is disabled.
     */
    protected getSemanticMapping(node: OfficeContentNode) {
        if (this.config.includeFormatting === false) {
            return undefined;
        }
        return this.styleMapper.getMapping(node);
    }

    /**
     * Entry point for generation.
     */
    abstract generate(): Promise<ConversionResult<D>>;

    /**
     * Centralized logic for handling the onNode callback.
     * Evaluates the callback and returns a result that tells the generator how to proceed.
     * 
     * @returns 
     * - `string`: Use this as the node's output, skip default processing.
     * - `false`: Skip this node and its subtree.
     * - `void`: Proceed with default processing.
     */
    /**
     * When set, memoizes each node's onNode verdict, so a generator that visits a node more than once
     * in a single pass (e.g. TextGenerator's spatial-layout collect followed by a flow fallback) still
     * asks the hook exactly once per node, per the documented contract. Left null by default.
     */
    protected onNodeMemo: WeakMap<OfficeContentNode, string | false | undefined> | null = null;

    protected async handleOnNode(node: OfficeContentNode): Promise<string | false | void> {
        if (this.onNodeMemo?.has(node)) return this.onNodeMemo.get(node);
        const result = await this.config.onNode(node);
        const verdict: string | false | undefined = result === false ? false : (typeof result === 'string' ? result : undefined);
        this.onNodeMemo?.set(node, verdict);
        return verdict;
    }

    /**
     * Recursively processes nodes and builds output.
     * 
     * @param node - The current node being processed
     * @param processor - A function that takes a node and its children's output and returns the node's output string.
     * @returns The generated string for this node and its subtree.
     */
    protected async processNodeRecursive(
        node: OfficeContentNode,
        processor: (node: OfficeContentNode, childrenOutput: string) => string | Promise<string>
    ): Promise<string> {
        // Every text-based generator funnels its whole traversal through here, so one check
        // makes `abortSignal` effective for all of them. Previously the signal was read once
        // before generation began and never again, which meant it could decline to start work
        // but could not stop work already underway - not much use against a document large
        // enough to be worth aborting. The check is a property read on an optional signal, so
        // the per-node cost is negligible.
        checkAbortSignal(this.config.abortSignal);
        const override = await this.handleOnNode(node);

        if (override === false) return '';
        if (typeof override === 'string') return override;

        let childrenOutput = '';
        if (node.children) {
            // The output of the child before, for what goes between it and the next (see childSeparator).
            let previous = '';
            for (const child of node.children) {
                const piece = await this.processNodeRecursive(child, processor);
                if (!piece) continue;
                childrenOutput += this.childSeparator(previous, child, piece) + piece;
                previous = piece;
            }
        }

        if (node.notes && node.notes.length > 0) {
            if (node.type !== 'slide') {
                // Each note once, however many references share it: listed per reference, one note a
                // small document refers to thousands of times was written out that many times.
                for (const note of node.notes) {
                    if (this.collectedNoteSet.has(note)) continue;
                    this.collectedNoteSet.add(note);
                    this.collectedNotes.push(note);
                }
            }
        }

        let result = await processor(node, childrenOutput);

        if (node.type === 'slide' && node.notes && node.notes.length > 0) {
            for (const note of node.notes) {
                result += await this.processNodeRecursive(note, processor);
            }
        }

        return result;
    }

    /**
     * What goes between `previous`, the output of a node's child (empty for none), and `piece`, the
     * output of the child after it, `child`. Given that output alone, not all the node's output so
     * far: reading the end of a string built up piece by piece copies all of it each time.
     */
    protected childSeparator(_previous: string, _child: OfficeContentNode, _piece: string): string {
        return '';
    }

    /**
     * Helper to generate a unique ID (slug) from text. Letters and digits of every script are kept, as
     * GitHub's heading ids keep them (`#überblick`, `#введение`): only punctuation and symbols go.
     */
    protected slugify(text: string): string {
        return text
            .toLowerCase()
            .replace(/[^\p{L}\p{M}\p{N}\s_-]/gu, '')
            .replace(/[\s_-]+/g, '-')
            .replace(/^-+|-+$/g, '');
    }

    private noteFootnoteKeys = new Map<OfficeContentNode, string>();
    private usedFootnoteKeys = new Set<string>();
    private footnoteKeyCounter = 0;

    /**
     * Assigns a stable, unique reference key to a footnote/endnote node, reused for both
     * its inline reference marker and its collected definition. Source ids aren't reliably
     * unique across a document - DOCX/ODT number footnotes and endnotes in separate
     * sequences, so both can carry noteId "1" - so a source id is only reused when it
     * hasn't already been claimed; otherwise a sequential counter guarantees uniqueness.
     */
    protected getFootnoteKey(note: OfficeContentNode): string {
        const cached = this.noteFootnoteKeys.get(note);
        if (cached) return cached;

        const preferred = (note.metadata as any)?.noteId;
        // The id is document-supplied and lands inside `[^...]` in Markdown, where the parser's
        // own recognizer accepts everything but `]` - so a note id carrying markup round-tripped
        // into the output verbatim. Accept it only when it looks like a label; otherwise fall
        // through to the sequential counter, which is always safe. Real ids are numeric (DOCX),
        // `ftn1`-shaped (ODT) or short slugs (`fn1`), and a Markdown label like `[^my note]`
        // still qualifies, so the fallback fires only for genuinely hostile input.
        const isLabelLike = typeof preferred === 'string' && /^[A-Za-z0-9_.:-][A-Za-z0-9 _.:-]*$/.test(preferred);
        let key: string;
        if (preferred && isLabelLike && !this.usedFootnoteKeys.has(preferred)) {
            key = preferred;
        } else {
            do {
                this.footnoteKeyCounter++;
                key = String(this.footnoteKeyCounter);
            } while (this.usedFootnoteKeys.has(key));
        }

        this.usedFootnoteKeys.add(key);
        this.noteFootnoteKeys.set(note, key);
        return key;
    }

    /**
     * True when every content-bearing text descendant satisfies `test` - i.e. the property is
     * uniform across the whole node and therefore says nothing the node type does not already say.
     *
     * Used to decide whether a heading's or header row's inherited formatting can be dropped. The
     * distinction matters: an ODF heading whose paragraph style is bold and 14pt yields a heading
     * where *every* run is bold and 14pt, and re-emitting that gives `# **Heading**` in Markdown
     * and, worse in RTF/HTML, an inner font-size that overrides the heading's own and visibly
     * shrinks it. But `# Normal **Bold** Normal` is an author contrasting one word against the
     * rest, and dropping that would discard real meaning. Only the uniform case is safe.
     *
     * Returns false when there is no text to judge, so an empty or image-only node never triggers
     * suppression.
     */
    protected hasUniformFormatting(
        node: OfficeContentNode,
        test: (formatting: OfficeContentNode['formatting']) => boolean
    ): boolean {
        let sawText = false;
        const walk = (n: OfficeContentNode): boolean => {
            if (n.type === 'text') {
                // Whitespace-only runs carry no visible formatting either way, so they neither
                // count as evidence nor veto - otherwise a stray unformatted space between two
                // bold runs would defeat the check on almost every real heading.
                if (!(n.text || '').trim()) return true;
                sawText = true;
                return test(n.formatting);
            }
            return (n.children ?? []).every(walk);
        };
        const uniform = (node.children ?? []).every(walk);
        return sawText && uniform;
    }

    /**
     * Recursively extracts plain text from a node and its children. A source comment (the author's
     * hidden note, `<!-- ... -->`) is not text, so it contributes nothing.
     */
    protected getNodeText(node: OfficeContentNode): string {
        if (isSourceComment(node)) return '';
        if (node.text) return node.text;
        if (node.children) {
            return node.children.map(c => this.getNodeText(c)).join('');
        }
        return '';
    }

    /**
     * Reports a warning to the user and collects it for the final result.
     */
    protected warn(type: OfficeWarningType, info?: any, node?: OfficeContentNode): void {
        const message = getWarningMessage(type, info);
        const issue: OfficeIssue = {
            type: 'warning',
            code: type,
            message,
            node,
            details: info
        };
        this.messages.push(issue);
        this.config.onWarning(issue);
    }
}
