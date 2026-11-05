/**
 * The marks notes are shown by at their references, for the writers that print one (HTML, EPUB, the
 * native PDF engine, slides, Markdown's HTML table cells).
 *
 * A note's key (`NoteMetadata.noteId`, DOCX's `w:id`, ODT's `ftn1`) is internal: it names the note in
 * ids and links and is never what a reader sees. What is shown is the mark the source records
 * (`NoteMetadata.mark`), or else the note's ordinal within its kind, footnotes and endnotes counted
 * apart: footnotes `1, 2, 3`, endnotes `i, ii, iii`, as word processors number them by default.
 */
import type { NoteMetadata, OfficeContentNode } from '../types.js';
import { formatListNumber } from './listMarkers.js';

/** The longest mark kept as recorded; a longer one is no reference mark, and the note is numbered instead. */
const MAX_MARK_LENGTH = 64;

/** The mark a note's source records for it, when it records one (trimmed, and of a reference's size). */
export function recordedNoteMark(note: OfficeContentNode): string | undefined {
    const mark = (note.metadata as NoteMetadata | undefined)?.mark;
    if (typeof mark !== 'string') return undefined;
    const trimmed = mark.replace(/\s+/g, ' ').trim();
    return trimmed && trimmed.length <= MAX_MARK_LENGTH ? trimmed : undefined;
}

/** The custom mark an author chose for a note (`NoteMetadata.customMark`), for writers that number notes themselves. */
export function customNoteMark(note: OfficeContentNode): string | undefined {
    return (note.metadata as NoteMetadata | undefined)?.customMark === true ? recordedNoteMark(note) : undefined;
}

/**
 * Gives each note its mark, once, in the order the writer meets the notes: the recorded mark, or the
 * next ordinal of its kind (a note shown by a recorded mark takes no number, as a custom mark takes
 * none in a word processor).
 */
export class NoteMarks {
    private readonly marks = new Map<OfficeContentNode, string>();
    private footnotes = 0;
    private endnotes = 0;

    /** The note's mark (the same one at every reference to it). */
    markOf(note: OfficeContentNode): string {
        const known = this.marks.get(note);
        if (known !== undefined) return known;
        const endnote = (note.metadata as NoteMetadata | undefined)?.noteType === 'endnote';
        const mark = recordedNoteMark(note) ?? (endnote ? formatListNumber(++this.endnotes, 'lowerRoman') : String(++this.footnotes));
        this.marks.set(note, mark);
        return mark;
    }
}
