/**
 * SmartArt: a DrawingML diagram (`dgm:`), whose text lives in its own data part (`diagrams/data1.xml`),
 * named from the frame showing it by `dgm:relIds r:dm`. Every SmartArt graphic's text was dropped, in
 * presentations and documents alike.
 */

import { ListMetadata, OfficeContentNode, OfficeParserConfig } from '../types.js';
import { getAttribute, getChildElements, getOutermostElements, parseXmlString } from './xmlUtils.js';

/** One item of a diagram's text: a paragraph of one of its points, at the point's depth in the diagram. */
export interface DiagramItem {
    text: string;
    depth: number;
}

/** The point types that are the diagram's content (its nodes and assistants); the rest are connectors and layout. */
const CONTENT_POINTS = new Set(['node', 'asst']);

/** The text of a DrawingML text body's paragraph: its runs' and fields' text, in order. */
const paragraphText = (paragraph: Element): string => getOutermostElements(paragraph, 'a:t').map(t => t.textContent || '').join('');

/**
 * The text of a diagram's data part, in the diagram's order: each content point's paragraphs, nested as
 * its parent-of connections nest the points (a point under the document point is at depth 0), siblings
 * in their connection order. A point no connection reaches is read after the rest, at depth 0, not
 * dropped. Linear in the part: each point is visited once, however the connections loop or repeat.
 */
export const readDiagram = (dataXml: string, config: OfficeParserConfig): DiagramItem[] => {
    const doc = parseXmlString(dataXml, { config });
    const model = doc.documentElement;
    if (!model) return [];
    const ptLst = getChildElements(model, 'dgm:ptLst')[0];
    const cxnLst = getChildElements(model, 'dgm:cxnLst')[0];

    const paragraphsOf = new Map<string, string[]>();
    const typeOf = new Map<string, string>();
    const order: string[] = [];
    for (const point of ptLst ? getChildElements(ptLst, 'dgm:pt') : []) {
        const id = getAttribute(point, 'modelId');
        if (id === undefined || typeOf.has(id)) continue;
        const type = getAttribute(point, 'type') ?? 'node';
        typeOf.set(id, type);
        order.push(id);
        const body = getChildElements(point, 'dgm:t')[0];
        if (body) paragraphsOf.set(id, getChildElements(body, 'a:p').map(paragraphText));
    }

    // Each point's children by the parent-of connections (a connection's type defaults to parOf),
    // in their source order.
    const childrenOf = new Map<string, { id: string; ord: number }[]>();
    for (const connection of cxnLst ? getChildElements(cxnLst, 'dgm:cxn') : []) {
        if ((getAttribute(connection, 'type') ?? 'parOf') !== 'parOf') continue;
        const source = getAttribute(connection, 'srcId');
        const destination = getAttribute(connection, 'destId');
        if (source === undefined || destination === undefined || !typeOf.has(source) || !typeOf.has(destination)) continue;
        const ord = Number(getAttribute(connection, 'srcOrd') ?? 0);
        let list = childrenOf.get(source);
        if (!list) childrenOf.set(source, list = []);
        list.push({ id: destination, ord: Number.isFinite(ord) ? ord : 0 });
    }
    for (const list of childrenOf.values()) list.sort((a, b) => a.ord - b.ord);

    const items: DiagramItem[] = [];
    const visited = new Set<string>();
    const emit = (id: string, depth: number) => {
        for (const text of paragraphsOf.get(id) ?? []) if (text.trim()) items.push({ text, depth });
    };
    // Depth first from each root, without recursion: a document point (whose children are at depth 0),
    // then any content point no walk reached.
    const walk = (root: string, rootDepth: number) => {
        const pending: { id: string; depth: number }[] = [{ id: root, depth: rootDepth }];
        while (pending.length) {
            const { id, depth } = pending.pop()!;
            if (visited.has(id)) continue;
            visited.add(id);
            if (CONTENT_POINTS.has(typeOf.get(id)!)) emit(id, Math.max(0, depth));
            const children = childrenOf.get(id) ?? [];
            for (let i = children.length - 1; i >= 0; i--) pending.push({ id: children[i].id, depth: depth + 1 });
        }
    };
    for (const id of order) if (typeOf.get(id) === 'doc') walk(id, -1);
    for (const id of order) if (!visited.has(id) && CONTENT_POINTS.has(typeOf.get(id)!)) walk(id, 0);
    return items;
};

/** A diagram's items as a bulleted list, `listId` naming it. */
export const diagramList = (items: DiagramItem[], listId: string): OfficeContentNode[] =>
    items.map((item, index) => ({
        type: 'list',
        text: item.text,
        children: [{ type: 'text', text: item.text }],
        metadata: { listType: 'unordered', indentation: Math.min(8, item.depth), listId, itemIndex: index } as ListMetadata,
    }));
