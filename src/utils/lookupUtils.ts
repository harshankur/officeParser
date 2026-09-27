/**
 * `entries` as a lookup table with no prototype, for a table read with keys a document or an AST
 * supplies. A plain object answers `constructor`, `toString` or `__proto__` with what it inherits:
 * a function went into the AST as a colour, a package whose `mimetype` read `constructor` failed as a
 * type, and a table written through such a key (`table[key].field = value`) wrote onto every object
 * in the process. Without a prototype, those are absent keys like any other.
 */
export function lookupTable<T extends object>(entries: T): T {
    return Object.assign(Object.create(null), entries);
}

/**
 * A plain object holding each of `record`'s entries as its own property: how a map built without a
 * prototype (keyed by the document's own names) is handed out, as the objects of the AST are plain.
 * A key such as `__proto__` stays an ordinary property, where assigning it to a plain object set the
 * object's prototype (and dropped the entry).
 */
export function plainRecord<T>(record: Record<string, T>): Record<string, T> {
    return Object.fromEntries(Object.entries(record));
}

/** `record[key] = value`, as an own property whatever `key` is (see plainRecord). */
export function setOwn<T>(record: Record<string, T>, key: string, value: T): void {
    Object.defineProperty(record, key, { value, enumerable: true, writable: true, configurable: true });
}
