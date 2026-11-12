/**
 * Embedded objects (OLE): an Excel sheet, an equation or another program's document a word processor
 * holds with a picture of it.
 *
 * @module oleUtils
 */

import { lookupTable } from './lookupUtils.js';

/** The longest ProgID kept as written, for a picture's alt text or a warning. */
const MAX_PROG_ID = 128;

/** The kind of object each ProgID names, by its first two parts, lowercased (`Excel.Sheet.12` is `excel.sheet`). */
const OBJECT_KINDS: Readonly<Record<string, string>> = lookupTable({
    'excel.sheet': 'Excel worksheet',
    'excel.sheetmacroenabled': 'Excel worksheet',
    'excel.sheetbinarymacroenabled': 'Excel worksheet',
    'excel.chart': 'Excel chart',
    'word.document': 'Word document',
    'word.documentmacroenabled': 'Word document',
    'powerpoint.show': 'PowerPoint presentation',
    'powerpoint.showmacroenabled': 'PowerPoint presentation',
    'powerpoint.slide': 'PowerPoint slide',
    'visio.drawing': 'Visio drawing',
    'msgraph.chart': 'Chart',
    'acroexch.document': 'PDF document',
    'package': 'Embedded file',
    'paint.picture': 'Picture',
    'pbrush': 'Picture',
});

/** A ProgID as written, cut to 128 characters; undefined for none. */
export function progIdOf(progId: string | undefined): string | undefined {
    const trimmed = (progId ?? '').trim();
    return trimmed ? trimmed.slice(0, MAX_PROG_ID) : undefined;
}

/**
 * The kind of an embedded object (OLE) as a reader would name it, from the ProgID its file gives
 * (DOCX `o:OLEObject/@ProgID` or `w:objectEmbed/@w:progId`, RTF `\objclass`): "Excel worksheet" for
 * Excel.Sheet.12, "Equation" for Equation.3 and MathType's Equation.DSMT4. A ProgID not listed is
 * returned as written (cut to 128 characters); none, undefined. Used as the alt text of the object's picture.
 */
export function embeddedObjectKind(progId: string | undefined): string | undefined {
    const written = progIdOf(progId);
    if (!written) return undefined;
    const parts = written.split('.');
    if (parts[0].toLowerCase() === 'equation') return 'Equation';
    return OBJECT_KINDS[parts.slice(0, 2).join('.').toLowerCase()] ?? OBJECT_KINDS[parts[0].toLowerCase()] ?? written;
}

/** How a warning names an embedded object: "the embedded Excel.Sheet.12 object", or "an embedded object" with no ProgID. */
export function embeddedObjectPart(progId: string | undefined): string {
    const written = progIdOf(progId);
    return written ? `the embedded ${written} object` : 'an embedded object';
}
