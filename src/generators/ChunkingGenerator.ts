import {
    DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG,
    DEFAULT_FIXED_SIZE_CHUNKING_CONFIG,
    DEFAULT_SEMANTIC_CHUNKING_CONFIG,
} from '../defaults.js';
import {
    ChunkingConfig,
    ConversionResult,
    DeepRequired,
    DocumentStructureChunkingConfig,
    FixedSizeChunkingConfig,
    GeneratorConfig,
    OfficeChunk,
    OfficeContentNode,
    OfficeErrorType,
    OfficeParserAST,
    OfficeWarningType,
    PageMetadata,
    SemanticChunkingConfig,
    SheetMetadata,
    SlideMetadata
} from '../types.js';
import { getOfficeError, checkAbortSignal } from '../utils/errorUtils.js';
import { BaseGenerator } from './BaseGenerator.js';

/**
 * Generates a list of OfficeChunk objects from an AST for use in RAG pipelines.
 * Supports three strategies: 'fixed-size', 'document-structure', and 'semantic'.
 */
export class ChunkingGenerator extends BaseGenerator<'chunks'> {
    /** The resolved chunking config (with defaults applied). */
    private chunkConfig: DeepRequired<ChunkingConfig>;
    /** Whether the user provided an explicit sentence boundary regex. */
    private isCustomRegex: boolean;

    constructor(ast: OfficeParserAST, config?: GeneratorConfig<'chunks'>) {
        super('chunks', ast, config);
        this.chunkConfig = this.resolveChunkingConfig(this.config.chunksConfig);
        // Track if the user explicitly provided a regex (vs using the library default)
        this.isCustomRegex = !!(config as any)?.chunksConfig?.sentenceBoundaryRegex;
    }

    /**
     * Merges the user's chunking config with the appropriate defaults for the chosen strategy.
     */
    private resolveChunkingConfig(userChunkConfig: ChunkingConfig): DeepRequired<ChunkingConfig> {
        const strategy = userChunkConfig?.strategy ?? 'document-structure';
        switch (strategy) {
            case 'fixed-size':
                return { ...DEFAULT_FIXED_SIZE_CHUNKING_CONFIG, ...userChunkConfig } as DeepRequired<FixedSizeChunkingConfig>;
            case 'semantic':
                return { ...DEFAULT_SEMANTIC_CHUNKING_CONFIG, ...userChunkConfig } as DeepRequired<SemanticChunkingConfig>;
            case 'document-structure':
                return { ...DEFAULT_DOCUMENT_STRUCTURE_CHUNKING_CONFIG, ...userChunkConfig } as DeepRequired<DocumentStructureChunkingConfig>;
        }
    }

    /**
     * Main entry point. Routes to the correct strategy implementation.
     * Note: ConversionResult.value is a JSON string of OfficeChunk[] for the 'chunks' destination.
     */
    async generate(): Promise<ConversionResult<'chunks'>> {
        checkAbortSignal(this.config.abortSignal);
        let chunks: OfficeChunk[];

        switch (this.chunkConfig.strategy) {
            case 'fixed-size':
                chunks = await this.generateFixedSize(this.chunkConfig as DeepRequired<FixedSizeChunkingConfig>);
                break;
            case 'semantic':
                chunks = await this.generateSemantic(this.chunkConfig as DeepRequired<SemanticChunkingConfig>);
                break;
            case 'document-structure':
                chunks = await this.generateDocumentStructure(this.chunkConfig as DeepRequired<DocumentStructureChunkingConfig>);
                break;
        }

        return {
            value: chunks as any,
            messages: this.messages,
        };
    }

    // ─── Strategy 1: Fixed-Size ────────────────────────────────────────────────

    /**
     * Splits the full document text into fixed-size chunks with optional overlap.
     * Attempts to split on natural separators before hard-cutting.
     */
    private async generateFixedSize(config: DeepRequired<FixedSizeChunkingConfig>): Promise<OfficeChunk[]> {
        const { chunkSize, chunkOverlap, separators, lengthFunction: measure } = config;

        // Build a flat text with positional map from top-level nodes
        const { text: fullText, nodeMap } = await this.buildFlatTextWithPositions();
        const rawChunks = this.splitTextRecursively(fullText, chunkSize, chunkOverlap, separators, measure);

        return rawChunks.map(({ text, start, end }) => {
            const chunk: OfficeChunk = { text, metadata: { sourceType: this.ast.type } };
            if (config.includeMetadata ?? true) {
                this.enrichMetadataFromPosition(chunk, nodeMap, start);
            }
            if (config.addStartIndex) {
                chunk.startIndex = start;
                chunk.endIndex = end;
            }
            return chunk;
        });
    }

    /**
     * Recursively tries separators to split text into chunks of at most `chunkSize`,
     * with `chunkOverlap` characters of overlap between consecutive chunks.
     */
    private splitTextRecursively(
        text: string,
        chunkSize: number,
        chunkOverlap: number,
        separators: string[],
        measure: (t: string) => number
    ): { text: string; start: number; end: number }[] {
        if (measure(text) <= chunkSize) {
            return text.trim() ? [{ text, start: 0, end: text.length }] : [];
        }

        let chosenSep: string | undefined;
        let nextSeparators: string[] = [];
        let parts: string[] = [];
        let actualSep = '';

        // 1. Try splitting into sentences first
        const sentences = this.splitIntoSentences(text);
        if (sentences.length > 1) {
            parts = sentences;
            actualSep = ' ';
            chosenSep = 'SENTENCE_SPLIT'; // Internal marker to avoid hard-cut
            nextSeparators = separators;
        } else {
            // 2. Fallback to the character separators provided in config
            for (let i = 0; i < separators.length; i++) {
                const sep = separators[i];
                parts = sep ? text.split(sep) : [...text];
                if (parts.length > 1) {
                    chosenSep = sep;
                    actualSep = sep;
                    nextSeparators = separators.slice(i + 1);
                    break;
                }
            }
        }

        if (chosenSep === undefined) {
            // Last resort: hard cut
            const results: { text: string; start: number; end: number }[] = [];
            let offset = 0;
            while (offset < text.length) {
                const slice = text.slice(offset, offset + chunkSize);
                results.push({ text: slice, start: offset, end: offset + slice.length });
                offset += Math.max(1, chunkSize - chunkOverlap);
            }
            return results;
        }

        const results: { text: string; start: number; end: number }[] = [];
        let currentChunk = '';
        let currentStart = 0;
        let absoluteOffset = 0;

        for (let i = 0; i < parts.length; i++) {
            const part = parts[i];
            const candidate = currentChunk ? currentChunk + actualSep + part : part;

            if (measure(candidate) <= chunkSize) {
                currentChunk = candidate;
            } else {
                if (currentChunk.trim()) {
                    // If currentChunk is still too big (should only happen if it's a single part), recurse!
                    if (measure(currentChunk) > chunkSize) {
                        const subResults = this.splitTextRecursively(currentChunk, chunkSize, chunkOverlap, nextSeparators, measure);
                        for (const r of subResults) {
                            results.push({ text: r.text, start: currentStart + r.start, end: currentStart + r.end });
                        }
                    } else {
                        results.push({ text: currentChunk, start: currentStart, end: currentStart + currentChunk.length });
                    }
                }

                // Start next chunk
                if (chunkOverlap > 0 && currentChunk.length > chunkOverlap && measure(currentChunk) <= chunkSize) {
                    const overlapText = currentChunk.slice(-chunkOverlap);
                    currentStart = absoluteOffset - chunkOverlap;
                    currentChunk = overlapText + actualSep + part;
                } else {
                    currentStart = absoluteOffset;
                    currentChunk = part;
                }
            }
            absoluteOffset += part.length + actualSep.length;
        }

        if (currentChunk.trim()) {
            if (measure(currentChunk) > chunkSize) {
                const subResults = this.splitTextRecursively(currentChunk, chunkSize, chunkOverlap, nextSeparators, measure);
                for (const r of subResults) {
                    results.push({ text: r.text, start: currentStart + r.start, end: currentStart + r.end });
                }
            } else {
                results.push({ text: currentChunk, start: currentStart, end: currentStart + currentChunk.length });
            }
        }

        return results;
    }

    // ─── Strategy 2: Document Structure ───────────────────────────────────────

    /**
     * Walks the AST and splits at the designated structural boundaries (slide, page, heading, paragraph).
     */
    private async generateDocumentStructure(config: DeepRequired<DocumentStructureChunkingConfig>): Promise<OfficeChunk[]> {
        const { splitBy, maxChunkSize, lengthFunction: measure } = config;

        const chunks: OfficeChunk[] = [];
        const contextStack: { heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string } = {};

        // Walk top-level nodes and decide where to cut
        for (const node of this.ast.content) {
            await this.processNodeForStructure(node, config, splitBy, maxChunkSize, measure, chunks, contextStack);
        }

        return this.finalizeChunks(chunks, config);
    }

    private async processNodeForStructure(
        node: OfficeContentNode,
        config: DeepRequired<DocumentStructureChunkingConfig>,
        splitBy: string,
        maxChunkSize: number,
        measure: (t: string) => number,
        chunks: OfficeChunk[],
        contextStack: { heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string }
    ): Promise<void> {
        checkAbortSignal(this.config.abortSignal);
        // Check for node override or skip
        const override = await this.handleOnNode(node);
        if (override === false) return;

        // Update context from structural container nodes
        if (node.type === 'slide') {
            const meta = node.metadata as SlideMetadata;
            contextStack.slideNumber = meta?.slideNumber;
        } else if (node.type === 'page') {
            const meta = node.metadata as PageMetadata;
            contextStack.pageNumber = meta?.pageNumber;
        } else if (node.type === 'sheet') {
            const meta = node.metadata as SheetMetadata;
            contextStack.sheetName = meta?.sheetName;
        } else if (node.type === 'heading') {
            contextStack.heading = node.text;
        }

        const isStructuralBoundary = this.isStructuralBoundary(node, splitBy);
        const isForcedSplit = splitBy === 'slide' && node.type === 'slide'
            || splitBy === 'page' && node.type === 'page'
            || splitBy === 'sheet' && node.type === 'sheet';

        if (isForcedSplit) {
            // Process children of the container as individual chunks within the boundary
            if (node.children || node.notes) {
                const innerChunks: OfficeChunk[] = [];
                if (node.children) {
                    for (const child of node.children) {
                        await this.processNodeForStructure(child, config, 'paragraph', maxChunkSize, measure, innerChunks, contextStack);
                    }
                }
                if (node.notes) {
                    for (const note of node.notes) {
                        await this.processNodeForStructure(note, config, 'paragraph', maxChunkSize, measure, innerChunks, contextStack);
                    }
                }
                for (const ic of innerChunks) {
                    ic.metadata.slideNumber = contextStack.slideNumber;
                    ic.metadata.pageNumber = contextStack.pageNumber;
                    ic.metadata.sheetName = contextStack.sheetName;
                    chunks.push(ic);
                }
            }
            return;
        }

        if (node.type === 'table') {
            await this.processTableNode(node, config, maxChunkSize, measure, chunks, contextStack);
            return;
        }

        const isContentNode = node.type === 'paragraph' || node.type === 'heading' || node.type === 'list' || node.type === 'code' || node.type === 'cell' || (node.text && (!node.children || node.children.length === 0));

        if (isStructuralBoundary || isContentNode) {
            const text = typeof override === 'string' ? override : (node.text ?? '');
            const isWhitespaceOnly = !text.trim() && !text.includes('\u00A0');

            if (isWhitespaceOnly && text.length > 0) {
                // Log skipped empty nodes if debugging
                // Only warn for non-cell nodes to reduce spreadsheet noise
                if (node.type !== 'cell') {
                    this.warn(OfficeWarningType.WHITESPACE_NODE_SKIPPED, node.type, node);
                }
                return;
            }

            if (node.type === 'heading') contextStack.heading = text;

            const chunk: OfficeChunk = {
                text,
                metadata: {
                    sourceType: this.ast.type,
                    closestHeading: contextStack.heading,
                    slideNumber: contextStack.slideNumber,
                    pageNumber: contextStack.pageNumber,
                    sheetName: contextStack.sheetName,
                },
            };

            // If this chunk is too big, further split it
            if (measure(text) > maxChunkSize) {
                const subChunks = this.splitTextRecursively(text, maxChunkSize, 0, ['\n\n', '\n', ' ', ''], measure);
                for (const sub of subChunks) {
                    chunks.push({ ...chunk, text: sub.text });
                }
            } else {
                chunks.push(chunk);
            }
            return;
        }

        // Recurse into children and notes for container nodes
        if (node.children) {
            for (const child of node.children) {
                await this.processNodeForStructure(child, config, splitBy, maxChunkSize, measure, chunks, contextStack);
            }
        }
        if (node.notes) {
            for (const note of node.notes) {
                await this.processNodeForStructure(note, config, splitBy, maxChunkSize, measure, chunks, contextStack);
            }
        }
    }

    private isStructuralBoundary(node: OfficeContentNode, splitBy: string): boolean {
        if (splitBy === 'heading') return node.type === 'heading';
        if (splitBy === 'paragraph') return node.type === 'paragraph' || node.type === 'heading' || node.type === 'cell';
        return false;
    }

    /**
     * Handles table chunking with the configured tableSplitStrategy.
     * 'row': keeps header row attached to every chunk.
     * 'flatten': converts table to text and splits normally.
     */
    private async processTableNode(
        node: OfficeContentNode,
        config: DeepRequired<DocumentStructureChunkingConfig>,
        maxChunkSize: number,
        measure: (t: string) => number,
        chunks: OfficeChunk[],
        contextStack: { heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string }
    ): Promise<void> {
        const strategy = config.tableSplitStrategy;

        if (strategy === 'flatten' || !node.children || node.children.length === 0) {
            // Flatten: treat as plain text
            const text = node.text ?? '';
            if (!text.trim()) return;
            chunks.push({
                text,
                metadata: {
                    sourceType: this.ast.type,
                    closestHeading: contextStack.heading,
                    slideNumber: contextStack.slideNumber,
                    pageNumber: contextStack.pageNumber,
                    sheetName: contextStack.sheetName,
                },
            });
            return;
        }

        // 'row' strategy: extract header row(s) and chunk remaining rows
        const rows = node.children; // Each child is a 'row' node
        const headerRows: OfficeContentNode[] = [];
        const dataRows: OfficeContentNode[] = [];

        // Heuristic: first row is the header
        if (rows.length > 0) {
            const firstRow = rows[0];
            const override = await this.handleOnNode(firstRow);
            if (override !== false) {
                // If overridden, we use the string as header, but we still treat it as a header
                headerRows.push(firstRow); // We still add the node to get metadata later if needed, but renderRowsAsText will handle it?
                // Actually renderRowsAsText needs to be updated too.
            }

            for (let i = 1; i < rows.length; i++) {
                dataRows.push(rows[i]);
            }
        }

        const headerText = await this.renderRowsAsText(headerRows);
        const baseMetadata = {
            sourceType: this.ast.type,
            closestHeading: contextStack.heading,
            slideNumber: contextStack.slideNumber,
            pageNumber: contextStack.pageNumber,
            sheetName: contextStack.sheetName,
            isTableChunk: true,
        };

        // Group data rows into chunks
        let currentRows: OfficeContentNode[] = [];
        let currentSize = headerText ? measure(headerText) : 0;

        const flushCurrentRows = async () => {
            if (currentRows.length === 0) return;
            const rowText = await this.renderRowsAsText(currentRows);
            const chunkText = headerText ? `${headerText}\n${rowText}` : rowText;
            if (chunkText.trim()) {
                if (measure(chunkText) > maxChunkSize) {
                    // Fallback to recursive splitting for oversized table chunks
                    const subChunks = this.splitTextRecursively(chunkText, maxChunkSize, 0, ['\n', ' ', ''], measure);
                    for (const sub of subChunks) {
                        chunks.push({ text: sub.text, metadata: { ...baseMetadata } });
                    }
                } else {
                    chunks.push({ text: chunkText, metadata: { ...baseMetadata } });
                }
            }
            currentRows = [];
            currentSize = headerText ? measure(headerText) : 0;
        };

        for (const row of dataRows) {
            const override = await this.handleOnNode(row);
            if (override === false) continue;

            const rowText = typeof override === 'string' ? override : await this.renderRowsAsText([row]);
            const rowSize = measure(rowText);
            if (currentSize + rowSize > maxChunkSize && currentRows.length > 0) {
                await flushCurrentRows();
            }
            if (typeof override === 'string') {
                // If row was overridden, we flush current, then add the override as a chunk.
                await flushCurrentRows();
                const chunkText = headerText ? `${headerText}\n${override}` : override;
                chunks.push({ text: chunkText, metadata: { ...baseMetadata } });
                continue;
            }
            currentRows.push(row);
            currentSize += rowSize;
        }
        await flushCurrentRows();
    }

    /**
     * Renders a list of row nodes as a pipe-separated text string.
     */
    private async renderRowsAsText(rows: OfficeContentNode[]): Promise<string> {
        const renderedRows: string[] = [];
        for (const row of rows) {
            const override = await this.handleOnNode(row);
            if (override === false) continue;
            if (typeof override === 'string') {
                renderedRows.push(override);
                continue;
            }

            if (!row.children) {
                renderedRows.push(row.text ?? '');
                continue;
            }

            const getCellText = (cell: OfficeContentNode): string => {
                if (cell.text) return cell.text;
                if (!cell.children || cell.children.length === 0) return '';
                return cell.children.map(c => getCellText(c)).join(' ');
            };

            const cells = row.children.map(cell => getCellText(cell).replace(/\n/g, ' ').trim());
            renderedRows.push(`| ${cells.join(' | ')} |`);
        }
        return renderedRows.join('\n');
    }

    // ─── Strategy 3: Semantic ─────────────────────────────────────────────────

    /**
     * Splits document into semantically coherent chunks using cosine similarity
     * between sentence embeddings. A new chunk begins when similarity drops
     * below `similarityThreshold`.
     */
    private async generateSemantic(config: DeepRequired<SemanticChunkingConfig>): Promise<OfficeChunk[]> {
        const { similarityThreshold: threshold, maxChunkSize, bufferSize, lengthFunction: measure, embeddingBatchSize: batchSize } = config;

        if (typeof config.embeddingFunction !== 'function') {
            throw getOfficeError(OfficeErrorType.MISSING_EMBEDDING_FUNCTION, this.config as any);
        }

        // Extract all leaf-level text sentences from the AST
        const sentences = await this.extractSentences();
        if (sentences.length === 0) return [];

        // Embed all sentences in batches to avoid rate limiting
        const embeddings = await this.batchEmbeddings(sentences, config.embeddingFunction, batchSize, config.timeout);

        // Calculate cosine similarity between adjacent sentence windows
        const chunks: OfficeChunk[] = [];
        let currentSentences: typeof sentences = [];
        let currentSize = 0;

        for (let i = 0; i < sentences.length; i++) {
            const sentenceText = sentences[i].text;
            currentSentences.push(sentences[i]);
            currentSize += measure(sentenceText);

            // Check if we should split here
            const isLast = i === sentences.length - 1;
            const exceedsMax = currentSize > maxChunkSize;

            let shouldSplit = isLast || exceedsMax;

            if (!shouldSplit && i < sentences.length - 1) {
                // Compare the current window with the next window
                const currentWindowEnd = Math.min(i + bufferSize, sentences.length - 1);
                const nextWindowStart = i + 1;
                const nextWindowEnd = Math.min(i + 1 + bufferSize, sentences.length - 1);

                const currentEmbedding = this.averageEmbeddings(embeddings.slice(Math.max(0, i - bufferSize + 1), currentWindowEnd + 1));
                const nextEmbedding = this.averageEmbeddings(embeddings.slice(nextWindowStart, nextWindowEnd + 1));
                const similarity = this.cosineSimilarity(currentEmbedding, nextEmbedding);

                if (similarity < threshold) {
                    shouldSplit = true;
                }
            }

            if (shouldSplit && currentSentences.length > 0) {
                const chunkText = currentSentences.map(s => s.text).join(' ');
                const firstSentence = currentSentences[0];
                const baseMetadata = {
                    sourceType: this.ast.type,
                    closestHeading: firstSentence.closestHeading,
                    slideNumber: firstSentence.slideNumber,
                    pageNumber: firstSentence.pageNumber,
                    sheetName: firstSentence.sheetName,
                };

                if (measure(chunkText) > maxChunkSize) {
                    // Fallback to recursive splitting for oversized semantic chunks
                    const subChunks = this.splitTextRecursively(chunkText, maxChunkSize, 0, ['\n', ' ', ''], measure);
                    for (const sub of subChunks) {
                        chunks.push({
                            text: sub.text,
                            metadata: { ...baseMetadata }
                        });
                    }
                } else {
                    chunks.push({
                        text: chunkText,
                        metadata: baseMetadata
                    });
                }
                currentSentences = [];
                currentSize = 0;
            }
        }

        return this.finalizeChunks(chunks, config);
    }

    /**
     * Extracts all text sentences from the AST with their contextual metadata.
     */
    private async extractSentences(): Promise<{
        text: string;
        closestHeading?: string;
        slideNumber?: number;
        pageNumber?: number;
        sheetName?: string;
    }[]> {
        const results: {
            text: string;
            closestHeading?: string;
            slideNumber?: number;
            pageNumber?: number;
            sheetName?: string;
        }[] = [];
        let currentHeading: string | undefined;
        let currentSlide: number | undefined;
        let currentPage: number | undefined;
        let currentSheet: string | undefined;

        const walk = async (node: OfficeContentNode) => {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) return;

            if (node.type === 'heading') currentHeading = node.text;
            if (node.type === 'slide') currentSlide = (node.metadata as SlideMetadata)?.slideNumber;
            if (node.type === 'page') currentPage = (node.metadata as PageMetadata)?.pageNumber;
            if (node.type === 'sheet') currentSheet = (node.metadata as SheetMetadata)?.sheetName;

            const isContentNode = node.type === 'paragraph' || node.type === 'heading' || node.type === 'list' || node.type === 'cell' || (node.text && (!node.children || node.children.length === 0));

            if (isContentNode) {
                const text = (typeof override === 'string' ? override : (node.text ?? '')).trim();
                if (!text) return;
                // Split paragraph text into individual sentences for finer-grained similarity
                const sentences = this.splitIntoSentences(text);
                for (const sentence of sentences) {
                    if (sentence.trim()) {
                        results.push({
                            text: sentence.trim(),
                            closestHeading: currentHeading,
                            slideNumber: currentSlide,
                            pageNumber: currentPage,
                            sheetName: currentSheet,
                        });
                    }
                }
                return; // don't recurse into children; we already have the text
            }

            if (node.children) {
                for (const child of node.children) await walk(child);
            }
            if (node.notes) {
                for (const note of node.notes) await walk(note);
            }
        };

        for (const node of this.ast.content) await walk(node);
        return results;
    }

    // ─── Shared Utilities ─────────────────────────────────────────────────────

    /**
     * Builds a flat text string from the entire document and a map of
     * character offsets to AST node metadata for position-based metadata lookups.
     */
    private async buildFlatTextWithPositions(): Promise<{
        text: string;
        nodeMap: { offset: number; heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string }[];
    }> {
        const parts: string[] = [];
        const nodeMap: { offset: number; heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string }[] = [];
        let offset = 0;
        let currentHeading: string | undefined;
        let currentSlide: number | undefined;
        let currentPage: number | undefined;
        let currentSheet: string | undefined;

        const walk = async (node: OfficeContentNode) => {
            checkAbortSignal(this.config.abortSignal);
            const override = await this.handleOnNode(node);
            if (override === false) return;

            if (node.type === 'heading') currentHeading = node.text;
            if (node.type === 'slide') currentSlide = (node.metadata as SlideMetadata)?.slideNumber;
            if (node.type === 'page') currentPage = (node.metadata as PageMetadata)?.pageNumber;
            if (node.type === 'sheet') currentSheet = (node.metadata as SheetMetadata)?.sheetName;

            const isContentNode = node.type === 'paragraph' || node.type === 'heading' || node.type === 'list' || node.type === 'code' || node.type === 'cell' || (node.text && (!node.children || node.children.length === 0));

            if (isContentNode) {
                const nodeText = typeof override === 'string' ? override : (node.text || '');
                const txt = nodeText + '\n';
                nodeMap.push({ offset, heading: currentHeading, slideNumber: currentSlide, pageNumber: currentPage, sheetName: currentSheet });
                parts.push(txt);
                offset += txt.length;
                return;
            }
            if (node.children) {
                for (const child of node.children) await walk(child);
            }
            if (node.notes) {
                for (const note of node.notes) await walk(note);
            }
        };

        for (const node of this.ast.content) await walk(node);
        return { text: parts.join(''), nodeMap };
    }

    /**
     * Finds the closest AST metadata for a given character position.
     */
    private enrichMetadataFromPosition(
        chunk: OfficeChunk,
        nodeMap: { offset: number; heading?: string; slideNumber?: number; pageNumber?: number; sheetName?: string }[],
        charOffset: number
    ): void {
        let best = nodeMap[0];
        for (const entry of nodeMap) {
            if (entry.offset <= charOffset) best = entry;
            else break;
        }
        if (best) {
            chunk.metadata.closestHeading = best.heading;
            chunk.metadata.slideNumber = best.slideNumber;
            chunk.metadata.pageNumber = best.pageNumber;
            chunk.metadata.sheetName = best.sheetName;
        }
    }

    /**
     * Applies final post-processing: strips whitespace, sets sourceType.
     */
    private finalizeChunks(chunks: OfficeChunk[], config: { stripWhitespace?: boolean; includeMetadata?: boolean }): OfficeChunk[] {
        if (chunks.length === 0) {
            this.warn(OfficeWarningType.EMPTY_CHUNK_GENERATED, this.chunkConfig.strategy);
        }

        return chunks
            .map(chunk => {
                const text = config.stripWhitespace !== false ? chunk.text.trim() : chunk.text;
                const result: OfficeChunk = { text, metadata: { sourceType: this.ast.type } };
                if (config.includeMetadata !== false) {
                    result.metadata = chunk.metadata;
                }
                return result;
            })
            .filter(chunk => chunk.text.length > 0);
    }

    // ─── Embedding Math Utilities ──────────────────────────────────────────────

    /**
     * Helper to process embeddings in sequential batches to avoid API rate limits and memory issues.
     */
    private async batchEmbeddings(
        sentences: { text: string }[],
        embedFn: (text: string) => Promise<number[]>,
        batchSize: number = 50,
        timeoutMs?: number
    ): Promise<number[][]> {
        const results: number[][] = [];
        for (let i = 0; i < sentences.length; i += batchSize) {
            checkAbortSignal(this.config.abortSignal);
            const batch = sentences.slice(i, i + batchSize);
            const batchPromises = batch.map(s => {
                const call = embedFn(s.text);
                if (timeoutMs !== undefined && timeoutMs > 0) {
                    let timerId: any;
                    const timeoutPromise = new Promise<never>((_, reject) => {
                        timerId = setTimeout(() => {
                            reject(new Error(`Embedding call timed out after ${timeoutMs}ms`));
                        }, timeoutMs);
                    });
                    return Promise.race([call, timeoutPromise]).finally(() => {
                        clearTimeout(timerId);
                    });
                }
                return call;
            });
            const batchResults = await Promise.all(batchPromises);
            results.push(...batchResults);
        }
        return results;
    }

    private cosineSimilarity(a: number[], b: number[]): number {
        if (a.length !== b.length || a.length === 0) return 0;
        let dot = 0, normA = 0, normB = 0;
        for (let i = 0; i < a.length; i++) {
            dot += a[i] * b[i];
            normA += a[i] * a[i];
            normB += b[i] * b[i];
        }
        const denom = Math.sqrt(normA) * Math.sqrt(normB);
        return denom === 0 ? 0 : dot / denom;
    }

    private averageEmbeddings(embeddings: number[][]): number[] {
        if (embeddings.length === 0) return [];
        const len = embeddings[0].length;
        const avg = new Array(len).fill(0);
        for (const emb of embeddings) {
            for (let i = 0; i < len; i++) avg[i] += emb[i];
        }
        return avg.map(v => v / embeddings.length);
    }

    /**
     * Robustly splits text into sentences, respecting abbreviations and non-Western punctuation.
     */
    private splitIntoSentences(text: string): string[] {
        if (this.isCustomRegex) {
            const userRegex = this.chunkConfig.sentenceBoundaryRegex;
            const regex = typeof userRegex === 'string' ? new RegExp(userRegex, 'g') : userRegex;
            // Split while keeping the separator if possible, or just split
            return text.split(regex).map(s => s.trim()).filter(Boolean);
        }

        const abbreviations = this.chunkConfig.abbreviations;
        const sentences: string[] = [];
        let start = 0;

        // Japanese full stop: 。 Exclamation: ！ Question: ？
        // Western: . ! ?
        const markRegex = /[.!?。！？]/g;
        let match;

        while ((match = markRegex.exec(text)) !== null) {
            const mark = match[0];
            const pos = match.index;

            const nextChar = text[pos + 1];
            const isAtEnd = pos === text.length - 1;
            const isFollowedByWhitespace = !nextChar || /\s/.test(nextChar);
            const isJapaneseMark = /[。！？]/.test(mark);

            if (isFollowedByWhitespace || isJapaneseMark) {
                // Check for abbreviations (only for period)
                if (mark === '.') {
                    const prevSpace = text.lastIndexOf(' ', pos - 1);
                    let lastWord = text.substring(prevSpace + 1, pos);

                    // Strip punctuation like quotes, parentheses, brackets
                    lastWord = lastWord.replace(/^[^\w]+|[^\w]+$/g, '');

                    if (abbreviations.includes(lastWord)) continue;
                }

                sentences.push(text.substring(start, pos + 1).trim());
                start = pos + 1;
            }
        }

        if (start < text.length) {
            const remaining = text.substring(start).trim();
            if (remaining) sentences.push(remaining);
        }

        return sentences.length > 0 ? sentences : [text];
    }
}
