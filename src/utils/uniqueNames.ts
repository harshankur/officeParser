/**
 * Names kept unique by a numbered suffix. The numbering resumes, for each base, where it left off:
 * tried from 2 again each time, the k-th use of one base tested k names, so 60,000 bookmarks named `a`
 * in a 7 KB DOCX took two minutes to write, and pictures sharing a name took time in the square of
 * their number.
 */
export class UniqueNames {
    private readonly used = new Set<string>();
    private readonly next = new Map<string, number>();

    /** `key` folds names that count as the same (a case-insensitive file system's, say). */
    constructor(private readonly key: (name: string) => string = name => name) { }

    has(name: string): boolean {
        return this.used.has(this.key(name));
    }

    add(name: string): void {
        this.used.add(this.key(name));
    }

    /** `base` when it is free, else the first free `numbered(n)` for n = 2, 3, ...; either way now taken. */
    claim(base: string, numbered: (n: number) => string): string {
        if (!this.has(base)) { this.add(base); return base; }
        const baseKey = this.key(base);
        let n = this.next.get(baseKey) ?? 2;
        let name = numbered(n);
        while (this.has(name)) name = numbered(++n);
        this.next.set(baseKey, n + 1);
        this.add(name);
        return name;
    }
}
