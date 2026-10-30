import * as assert from 'assert';
import { OfficeParser } from '../../src/OfficeParser';
import type { FixedSizeChunkingConfig, OfficeChunk } from '../../src/types';

export async function testUnicodeChunks(): Promise<void> {
    const split = async (text: string, config: Omit<FixedSizeChunkingConfig, 'strategy'>): Promise<OfficeChunk[]> => {
        const ast = await OfficeParser.parseOffice(Buffer.from(text), { fileType: 'md' });
        return (await ast.to('chunks', { chunksConfig: { strategy: 'fixed-size', addStartIndex: true, ...config } })).value;
    };
    const wellFormed = (chunks: OfficeChunk[]) => {
        for (const chunk of chunks) assert.ok(!/[\uD800-\uDFFF]/u.test(chunk.text), `isolated surrogate: ${JSON.stringify(chunk.text)}`);
    };
    for (const size of [0.5, 1, 1.5, 2, 2.5, 3, 4]) {
        const text = 'a😀b𐐀c';
        const chunks = await split(text, { chunkSize: size, chunkOverlap: 0, separators: [] });
        wellFormed(chunks);
        assert.strictEqual(chunks.map(c => c.text).join('').trimEnd(), text, `size ${size}: preserve every code point once`);
        for (const c of chunks) assert.strictEqual((text + '\n').slice(c.startIndex, c.endIndex), c.text, 'UTF-16 source offsets');
    }
    for (const overlap of [1, 1.5, 2]) {
        const chunks = await split('aa😀 bb', { chunkSize: 5, chunkOverlap: overlap, separators: [' '] });
        wellFormed(chunks);
        assert.ok(chunks.some(c => c.text.startsWith('😀')), 'overlap keeps the whole emoji');
        for (const c of chunks) assert.strictEqual('aa😀 bb\n'.slice(c.startIndex, c.endIndex), c.text, 'overlap source offsets');
    }
    for (const overlap of [1, 2, 5]) {
        const chunks = await split('a😀b𐐀c', { chunkSize: 3, chunkOverlap: overlap, separators: [] });
        wellFormed(chunks);
        assert.ok(chunks.length > 0 && chunks.length <= 8, 'overlap makes progress');
        for (const c of chunks) assert.strictEqual('a😀b𐐀c\n'.slice(c.startIndex, c.endIndex), c.text);
    }
    const ascii = await split('abcdefgh', { chunkSize: 3, chunkOverlap: 1, separators: [] });
    assert.deepStrictEqual(ascii.map(c => c.text), ['abc', 'cde', 'efg', 'gh\n', '\n'], 'ASCII hard-cut behavior stays unchanged');
}
