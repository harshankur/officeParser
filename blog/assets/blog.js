/**
 * What the blog's pages do once they are loaded: keep the reader's theme, filter the list by topic
 * and by the words searched for, bring a post's interactive figures to life, place a revision's
 * note by its button, and show or put away a revised post's changes. The pages themselves are built (scripts/build-blog.js) and
 * are complete without this script: every post is in its page, and the list shows every post.
 * Views are counted by the viewcounter, whose consent prompt the layout loads.
 */
(() => {
    'use strict';

    // ── Theme: set before the page is first painted (this script is in the head) ──
    const root = document.documentElement;
    const storedTheme = () => { try { return localStorage.getItem('theme'); } catch { return null; } };
    const saved = storedTheme() || 'auto';
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    root.setAttribute('data-theme', saved === 'auto' ? (prefersDark ? 'dark' : 'light') : saved);

    // ── A revised post's changes: shown unless the reader has put them away, which is kept for every post ──
    const CHANGES_KEY = 'blog-changes';
    const changesStored = () => { try { return localStorage.getItem(CHANGES_KEY); } catch { return null; } };
    root.classList.toggle('changes-hidden', changesStored() === 'hidden');

    const wireThemeToggle = () => {
        const icon = document.getElementById('theme-icon');
        const paint = () => { icon.textContent = root.getAttribute('data-theme') === 'dark' ? '☀️' : '🌙'; };
        paint();
        document.getElementById('theme-toggle').addEventListener('click', () => {
            const next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
            root.setAttribute('data-theme', next);
            try { localStorage.setItem('theme', next); } catch { /* the choice lasts for this page */ }
            paint();
        });
    };

    /**
     * Text as the search compares it: lower-case and without accents, so "cafe" finds "café".
     * `at[i]` is where character `i` of the folded text stands in `text`, to show the place a word was
     * found in as it is written.
     */
    const fold = text => {
        const at = [];
        let folded = '';
        let offset = 0;
        for (const character of text) {
            const plain = character.normalize('NFD').replace(/\p{M}/gu, '').toLowerCase();
            for (let i = 0; i < plain.length; i++) at.push(offset);
            folded += plain;
            offset += character.length;
        }
        return { folded, at };
    };

    /**
     * The list's filter: the topic chosen and the words searched for, together. A card is shown when
     * it is of the topic and every word is in its article: in the title, summary and topic the card
     * shows, or in the article's text, which is fetched (search.json, built with the pages) when the
     * reader first uses the search. Until it arrives, or where it cannot be fetched (the page opened
     * from a file), the search looks through what the cards show. The topic and the words are kept
     * in the address, so a search can be linked to and survives a reload.
     */
    const wireListFilter = () => {
        const tabs = document.getElementById('filter-tabs');
        if (!tabs) return;
        const input = document.getElementById('post-search-input');
        const status = document.getElementById('search-status');
        const none = document.getElementById('no-results');
        const buttons = [...tabs.querySelectorAll('.filter-btn')];
        const cards = [...document.querySelectorAll('.article-card')].map(card => ({
            card,
            category: card.dataset.category,
            match: card.querySelector('.card-match'),
            shown: fold(card.querySelector('.card-top').textContent).folded,
            body: null,
        }));
        const bySlug = new Map(cards.map(entry => [entry.card.dataset.slug, entry]));
        let topic = 'all';

        /** The place in an article's text where `term` stands, with a few words around it. */
        const showPlace = (entry, term) => {
            const { text, folded, at } = entry.body;
            const hit = folded.indexOf(term);
            const start = at[hit];
            const end = at[hit + term.length] ?? text.length;
            let from = Math.max(0, start - 60);
            let to = Math.min(text.length, end + 100);
            // From a word's start to a word's end, where there is one within reach.
            const firstSpace = text.indexOf(' ', from);
            if (from > 0 && firstSpace !== -1 && firstSpace < start) from = firstSpace + 1;
            const lastSpace = text.lastIndexOf(' ', to);
            if (to < text.length && lastSpace > end) to = lastSpace;
            const found = document.createElement('mark');
            found.textContent = text.slice(start, end);
            entry.match.replaceChildren((from > 0 ? '… ' : '') + text.slice(from, start), found, text.slice(end, to) + (to < text.length ? ' …' : ''));
        };

        const apply = () => {
            const terms = fold(input.value).folded.split(/\s+/).filter(Boolean);
            let visible = 0;
            for (const entry of cards) {
                // The words the card does not show have to be in the article's text.
                const unseen = terms.filter(term => !entry.shown.includes(term));
                const found = unseen.every(term => entry.body?.folded.includes(term));
                const show = found && (topic === 'all' || entry.category === topic);
                entry.card.hidden = !show;
                if (show) visible++;
                entry.match.hidden = !show || unseen.length === 0;
                if (!entry.match.hidden) showPlace(entry, unseen[0]);
            }
            none.hidden = visible > 0;
            const filtering = terms.length > 0 || topic !== 'all';
            status.hidden = !filtering;
            status.textContent = filtering ? `${visible} of ${cards.length} articles` : '';

            const address = new URL(location.href);
            const keep = (name, value) => { if (value) address.searchParams.set(name, value); else address.searchParams.delete(name); };
            keep('q', input.value.trim());
            keep('topic', topic === 'all' ? '' : topic);
            try { history.replaceState(null, '', address); } catch { /* a page opened from a file keeps its address */ }
        };

        const chooseTopic = chosen => {
            topic = chosen.dataset.category;
            for (const button of buttons) {
                button.classList.toggle('active', button === chosen);
                button.setAttribute('aria-pressed', String(button === chosen));
            }
        };

        let index = null;
        const loadIndex = () => {
            index ??= fetch('./search.json')
                .then(response => (response.ok ? response.json() : Promise.reject(new Error(String(response.status)))))
                .then(({ posts }) => {
                    for (const post of posts) {
                        const entry = bySlug.get(post.slug);
                        if (entry) entry.body = { text: post.text, ...fold(post.text) };
                    }
                    apply();
                })
                .catch(() => { /* the search still looks through what the cards show */ });
        };

        tabs.addEventListener('click', event => {
            const chosen = event.target.closest('.filter-btn');
            if (!chosen) return;
            chooseTopic(chosen);
            apply();
        });
        input.addEventListener('focus', loadIndex);
        input.addEventListener('input', () => { loadIndex(); apply(); });
        input.addEventListener('keydown', event => {
            if (event.key !== 'Escape' || !input.value) return;
            input.value = '';
            apply();
        });
        document.getElementById('clear-search').addEventListener('click', () => {
            input.value = '';
            apply();
            input.focus();
        });

        // The search works by this script, so the script is what shows it.
        document.getElementById('post-search').hidden = false;

        // A search or a topic the address carries is the one shown.
        const asked = new URLSearchParams(location.search);
        const askedTopic = buttons.find(button => button.dataset.category === asked.get('topic'));
        if (askedTopic) chooseTopic(askedTopic);
        if (asked.get('q')) {
            input.value = asked.get('q');
            loadIndex();
        }
        if (askedTopic || asked.get('q')) apply();
    };


    // ── Interactive figures ──
    // A post may hold a figure a reader can play with (`<figure class="demo ..." data-demo="...">`). The
    // figure is complete as written (a list, a table, a paragraph: what a feed reader and a page without
    // this script show); the script adds the controls and the motion.

    /** An element with its class, attributes and children (strings are text). */
    const el = (tag, attrs = {}, ...children) => {
        const node = document.createElement(tag);
        for (const [key, value] of Object.entries(attrs)) {
            if (value === undefined || value === null || value === false) continue;
            if (key === 'class') node.className = value;
            else if (key === 'text') node.textContent = value;
            else if (key.startsWith('on')) node.addEventListener(key.slice(2), value);
            else node.setAttribute(key, value === true ? '' : value);
        }
        for (const child of children.flat()) if (child !== null && child !== undefined) node.append(child);
        return node;
    };
    const reduceMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    /** A row of buttons of which one is chosen; `pick(value)` is called with the chosen one's value. */
    const segmented = (label, options, chosen, pick) => {
        const group = el('div', { class: 'demo-segmented', role: 'radiogroup', 'aria-label': label });
        const buttons = options.map(([value, text]) => el('button', {
            type: 'button', role: 'radio', 'aria-checked': String(value === chosen),
            onclick: () => { for (const b of buttons) b.setAttribute('aria-checked', String(b === buttons[options.findIndex(o => o[0] === value)])); pick(value); },
        }, text));
        group.append(...buttons);
        return group;
    };

    /** Moves children to new places with a glide (they are measured before and after: FLIP). */
    const glide = (container, change) => {
        const items = [...container.children];
        const before = new Map(items.map(item => [item, item.getBoundingClientRect()]));
        change();
        if (reduceMotion()) return;
        for (const item of items) {
            const a = before.get(item), b = item.getBoundingClientRect();
            const dx = a.left - b.left, dy = a.top - b.top;
            if (!dx && !dy) continue;
            item.animate([{ transform: `translate(${dx}px, ${dy}px)` }, { transform: 'none' }], { duration: 520, easing: 'cubic-bezier(0.2, 0.8, 0.2, 1)' });
        }
    };

    const DEMOS = {
        /**
         * A word's code points in the order typed, and its glyphs in the order drawn: each `li` of the list
         * is a code point, `data-drawn` its glyph's place in drawing order, `data-mark` one that takes no
         * room. The views: as typed, as drawn, and what copying gives without and with /ActualText.
         */
        shaping(figure) {
            const list = figure.querySelector('ol');
            if (!list) return;
            const items = [...list.children];
            const typed = items.map(li => li.querySelector('.demo-glyph')?.textContent ?? '').join('');
            const drawnOrder = [...items].sort((a, b) => Number(a.dataset.drawn) - Number(b.dataset.drawn));
            const drawn = drawnOrder.map(li => li.querySelector('.demo-glyph')?.textContent ?? '').join('');
            const result = el('div', { class: 'demo-result', 'aria-live': 'polite' });
            const VIEWS = {
                typed: { order: items, marks: false, say: 'The text, as typed and stored: one code point after another, in the order they are read.', copy: null },
                drawn: { order: drawnOrder, marks: true, say: 'The glyphs, as the shaper places them, left to right: the vowel sign first, the dot above the letter taking no room.', copy: null },
                glyphs: { order: drawnOrder, marks: true, say: 'Copied by mapping each glyph back to its characters, in the order drawn:', copy: drawn },
                actual: { order: drawnOrder, marks: true, say: 'Copied from the /ActualText the span carries:', copy: typed },
            };
            const show = view => {
                const v = VIEWS[view];
                glide(list, () => {
                    for (const li of v.order) list.append(li);
                    for (const li of items) li.classList.toggle('demo-glyph-mark', v.marks && li.dataset.mark !== undefined);
                });
                result.replaceChildren(...[el('span', {}, v.say), v.copy ? el('span', { class: 'demo-copied', lang: 'hi' }, v.copy) : null].filter(Boolean));
                figure.dataset.view = view;
            };
            figure.insertBefore(segmented('View', [['typed', 'As typed'], ['drawn', 'As drawn'], ['glyphs', 'Copied from glyphs'], ['actual', 'Copied from /ActualText']], 'typed', show), list);
            list.after(result);
            show('typed');
        },

        /**
         * Text going round a frame, as the native PDF engine sets it: drag the frame, choose how text goes
         * round it, and the lines are laid out again: on the side with more room, below the frame where it
         * leaves too little (less than 36 points or a third of the line), or not round it at all.
         */
        wrap(figure) {
            const WORDS = ('A frame stands where its document puts it, measured from the paragraph it is anchored to, '
                + 'and the lines of that paragraph go round it. Each line looks at the frames beside it before its first word '
                + 'is placed: the room on the left of a frame and the room on its right are measured, the text takes the side '
                + 'with more, and where that is less than a third of the line, the line moves below the frame. A table, a '
                + 'picture or a chart does not go round a frame at all: it starts below it. Drag the frame, change how text '
                + 'goes round it, and watch the lines find their places again.').split(' ');
            // A Letter page's column (468 points) on a wide screen; a narrower one on a phone, where the
            // text would be too small to read at that width.
            const PAGE = { width: 468, lineHeight: 15.5, size: 11 };
            const state = { x: 300, y: 18, width: 150, height: 80, wrap: 'beside' };
            const fit = () => {
                const width = page.clientWidth >= 560 ? 468 : 300;
                if (width === PAGE.width) return;
                state.x *= width / PAGE.width;
                state.width = width >= 468 ? 150 : 110;
                PAGE.width = width;
                clamp();
            };
            const page = el('div', { class: 'demo-page', 'aria-label': 'A column of text with a floating frame' });
            const lines = el('div', { class: 'demo-page-lines', 'aria-hidden': 'true' });
            const frame = el('div', { class: 'demo-frame', tabindex: '0', role: 'slider', 'aria-label': 'Frame position (arrow keys move it)' }, el('span', {}, 'Frame'));
            page.append(lines, frame);
            const readout = el('pre', { class: 'demo-readout' });
            const canvas = document.createElement('canvas').getContext('2d');
            const layout = () => {
                fit();
                const scale = page.clientWidth / PAGE.width;
                const font = `${PAGE.size * scale}px ${getComputedStyle(page).fontFamily}`;
                canvas.font = font;
                const width = w => canvas.measureText(w).width / scale;
                const space = width(' ');
                const gap = 6;
                const box = { left: state.x - gap, right: state.x + state.width + gap, top: state.y, bottom: state.y + state.height };
                const out = [];
                let y = 0, i = 0;
                while (i < WORDS.length && y < 2000) {
                    let from = 0, to = PAGE.width;
                    const lh = PAGE.lineHeight;
                    if (state.wrap !== 'none' && box.top < y + lh && box.bottom > y && box.right > 0 && box.left < PAGE.width) {
                        if (state.wrap === 'aboveBelow') { y = box.bottom; continue; }
                        if (box.left - from >= to - box.right) to = box.left; else from = box.right;
                        if (to - from < Math.min(36, PAGE.width / 3)) { y = box.bottom; continue; }
                    }
                    let used = 0;
                    const words = [];
                    while (i < WORDS.length) {
                        const w = width(WORDS[i]);
                        if (words.length && used + space + w > to - from) break;
                        used += (words.length ? space : 0) + w;
                        words.push(WORDS[i++]);
                    }
                    out.push({ x: from, y, text: words.join(' ') });
                    y += lh;
                }
                page.style.height = `${Math.max(y, state.y + state.height) * scale + 8}px`;
                lines.replaceChildren(...out.map(line => el('div', { class: 'demo-page-line', style: `left:${line.x * scale}px;top:${line.y * scale}px;font:${font}` }, line.text)));
                frame.style.left = `${state.x * scale}px`;
                frame.style.top = `${state.y * scale}px`;
                frame.style.width = `${state.width * scale}px`;
                frame.style.height = `${state.height * scale}px`;
                frame.dataset.wrap = state.wrap;
                readout.textContent = `float: { horizontalFrom: 'column', verticalFrom: 'paragraph',\n         x: ${Math.round(state.x)}, y: ${Math.round(state.y)}, width: ${state.width}, height: ${state.height}, wrap: '${state.wrap}' }\n// a column ${PAGE.width} points wide`;
            };
            // Dragged with a pointer, or moved with the arrow keys (shift for bigger steps).
            const clamp = () => {
                state.x = Math.max(0, Math.min(PAGE.width - state.width, state.x));
                state.y = Math.max(0, Math.min(220, state.y));
            };
            frame.addEventListener('pointerdown', down => {
                down.preventDefault();
                frame.setPointerCapture(down.pointerId);
                const scale = page.clientWidth / PAGE.width;
                const start = { x: state.x, y: state.y, px: down.clientX, py: down.clientY };
                const move = event => { state.x = start.x + (event.clientX - start.px) / scale; state.y = start.y + (event.clientY - start.py) / scale; clamp(); layout(); };
                const up = () => { frame.removeEventListener('pointermove', move); frame.removeEventListener('pointerup', up); frame.removeEventListener('pointercancel', up); };
                frame.addEventListener('pointermove', move);
                frame.addEventListener('pointerup', up);
                frame.addEventListener('pointercancel', up);
            });
            frame.addEventListener('keydown', event => {
                const step = event.shiftKey ? 24 : 6;
                const moves = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] };
                if (!moves[event.key]) return;
                event.preventDefault();
                state.x += moves[event.key][0]; state.y += moves[event.key][1];
                clamp(); layout();
            });
            const controls = el('div', { class: 'demo-controls' },
                segmented('How text goes round the frame', [['beside', 'Beside'], ['aboveBelow', 'Above and below'], ['none', 'Not at all']], state.wrap, value => { state.wrap = value; layout(); }));
            figure.querySelector('.demo-fallback')?.remove();
            figure.prepend(controls, page, readout);
            figure.classList.add('demo-live');
            layout();
            new ResizeObserver(() => layout()).observe(page);
        },

        /**
         * What a spreadsheet stores and what it shows: the table (each row a cell, its stored value, its format
         * and the text shown) is drawn as a sheet, which shows either, and its hidden rows or not.
         */
        cells(figure) {
            const table = figure.querySelector('table');
            if (!table) return;
            const rows = [...table.querySelectorAll('tbody tr')].map(tr => {
                const cells = [...tr.children].map(td => td.textContent.trim());
                return { label: cells[0], stored: cells[1], format: cells[2], shown: cells[3], hidden: tr.dataset.hidden !== undefined };
            });
            const state = { view: 'shown', hidden: false };
            const sheet = el('div', { class: 'demo-sheet', role: 'table', 'aria-label': 'The sheet' });
            const draw = () => {
                glide(sheet, () => {
                    sheet.replaceChildren(
                        el('div', { class: 'demo-sheet-row demo-sheet-head', role: 'row' }, el('span', { role: 'columnheader' }, ''), el('span', { role: 'columnheader' }, 'A'), el('span', { role: 'columnheader' }, 'B'), el('span', { class: 'demo-sheet-format', role: 'columnheader' }, 'Number format')),
                        ...rows.flatMap((row, i) => (row.hidden && !state.hidden) ? [el('div', { class: 'demo-sheet-gap', role: 'presentation', title: `Row ${i + 1} is hidden` })] : [el('div', { class: `demo-sheet-row${row.hidden ? ' demo-sheet-hidden' : ''}`, role: 'row' },
                            el('span', { class: 'demo-sheet-num', role: 'rowheader' }, String(i + 1)),
                            el('span', { role: 'cell' }, row.label),
                            el('span', { class: 'demo-sheet-value', role: 'cell', 'data-view': state.view }, state.view === 'shown' ? (row.shown || row.stored) : row.stored),
                            el('span', { class: 'demo-sheet-format', role: 'cell' }, row.format))]));
                });
            };
            const controls = el('div', { class: 'demo-controls' },
                segmented('Show', [['shown', 'What the sheet shows'], ['stored', 'What the file stores']], state.view, value => { state.view = value; draw(); }),
                el('label', { class: 'demo-check' }, el('input', { type: 'checkbox', onchange: event => { state.hidden = event.target.checked; draw(); } }), 'Show hidden rows'));
            figure.insertBefore(controls, table);
            figure.insertBefore(sheet, table);
            figure.classList.add('demo-live');
            draw();
        },

        /**
         * Where a PDF reader puts spaces: the fragments of a line (`data-runs`: each its text and the gap
         * before it, in hundredths of the font size) and a threshold to move. A space goes in where the gap
         * is at least the threshold times the font size (pdfParserConfig.spaceToleranceFactor).
         */
        spaces(figure) {
            let runs;
            try { runs = JSON.parse(figure.dataset.runs || '[]'); } catch { return; }
            if (!runs.length) return;
            const state = { factor: 0.25 };
            const strip = el('div', { class: 'demo-runs', 'aria-hidden': 'true' });
            const output = el('div', { class: 'demo-result', 'aria-live': 'polite' });
            const value = el('output', { class: 'demo-factor' });
            const slider = el('input', { type: 'range', min: '0', max: '0.5', step: '0.005', value: String(state.factor), 'aria-label': 'spaceToleranceFactor',
                oninput: event => { state.factor = Number(event.target.value); draw(); } });
            const draw = () => {
                value.textContent = String(Number(state.factor.toFixed(3))) + (Math.abs(state.factor - 0.25) < 0.001 ? ' (the default)' : '');
                let text = '';
                strip.replaceChildren(...runs.flatMap(([word, gap], i) => {
                    const spaced = i > 0 && gap / 100 >= state.factor;
                    if (spaced) text += ' ';
                    text += word;
                    return [i > 0 ? el('span', { class: `demo-gap${spaced ? ' demo-gap-space' : ''}${gap < 0 ? ' demo-gap-kern' : ''}`, style: `width:${Math.max(2, Math.abs(gap) * 0.6)}px`, title: `${(gap / 100).toFixed(2)} em` }) : null, el('span', { class: 'demo-run' }, word)];
                }).filter(Boolean));
                output.replaceChildren(el('span', {}, 'Read as: '), el('span', { class: 'demo-copied' }, text));
            };
            figure.querySelector('.demo-fallback')?.remove();
            figure.prepend(el('div', { class: 'demo-controls' }, el('label', { class: 'demo-slider' }, el('code', {}, 'spaceToleranceFactor'), slider, value)), strip, output);
            figure.classList.add('demo-live');
            draw();
        },

        /**
         * The same document cut two ways: by a window of so many characters over its text (`data-text`), and
         * by officeParser's chunker (`data-chunks`: its real output for each size, each chunk its text, its
         * closest heading and whether it is a table's).
         */
        chunks(figure) {
            let text, chunks;
            try { text = JSON.parse(figure.dataset.text || '""'); chunks = JSON.parse(figure.dataset.chunks || '{}'); } catch { return; }
            const sizes = Object.keys(chunks).map(Number).sort((a, b) => a - b);
            if (!text || !sizes.length) return;
            const state = { mode: 'window', size: sizes[1] ?? sizes[0] };
            const list = el('div', { class: 'demo-chunks', 'aria-live': 'polite' });
            const value = el('output', { class: 'demo-factor' });
            const slider = el('input', { type: 'range', min: '0', max: String(sizes.length - 1), step: '1', value: String(sizes.indexOf(state.size)), 'aria-label': 'maxChunkSize',
                oninput: event => { state.size = sizes[Number(event.target.value)]; draw(); } });
            const draw = () => {
                value.textContent = `${state.size} characters`;
                const cards = state.mode === 'window'
                    ? Array.from({ length: Math.ceil(text.length / state.size) }, (_, i) => ({ text: text.slice(i * state.size, (i + 1) * state.size), heading: null, table: false }))
                    : chunks[state.size].map(([t, heading, table]) => ({ text: t, heading, table }));
                glide(list, () => list.replaceChildren(...cards.map((card, i) => el('div', { class: `demo-chunk${card.table ? ' demo-chunk-table' : ''}` },
                    el('span', { class: 'demo-chunk-label' }, `Chunk ${i + 1}${card.heading ? ` · under “${card.heading}”` : ''}`),
                    el('pre', {}, card.text)))));
            };
            figure.querySelector('.demo-fallback')?.remove();
            figure.prepend(el('div', { class: 'demo-controls' },
                segmented('Cut by', [['window', 'A fixed window'], ['structure', "officeParser's chunker"]], state.mode, mode => { state.mode = mode; draw(); }),
                el('label', { class: 'demo-slider' }, el('code', {}, 'maxChunkSize'), slider, value)), list);
            figure.classList.add('demo-live');
            draw();
        },
    };

    const wireDemos = () => {
        for (const figure of document.querySelectorAll('figure.demo[data-demo]')) {
            try { DEMOS[figure.dataset.demo]?.(figure); } catch { /* the figure stays as written */ }
        }
    };

    /**
     * A revision's note, by its button: opened by a click (the page's own popover, which works without
     * this script) or by resting the pointer on the button, and placed beside the button where the page
     * alone would centre it. A note opened by the pointer closes when the pointer leaves the button and
     * the note, and a click keeps it open.
     */
    const wireRevisionNotes = () => {
        const GUTTER = 16;
        for (const button of document.querySelectorAll('.revision-mark[popovertarget]')) {
            const note = document.getElementById(button.getAttribute('popovertarget'));
            if (!note || typeof note.showPopover !== 'function') continue;
            let byPointer = false;
            let closing = 0;
            const place = () => {
                const at = button.getBoundingClientRect();
                const width = note.offsetWidth;
                const height = note.offsetHeight;
                const left = Math.min(Math.max(GUTTER, at.left + at.width / 2 - width / 2), window.innerWidth - GUTTER - width);
                const below = at.bottom + 8;
                const top = below + height <= window.innerHeight - GUTTER || at.top - 8 - height < GUTTER ? below : at.top - 8 - height;
                Object.assign(note.style, { position: 'fixed', inset: 'auto', margin: '0', left: `${Math.max(GUTTER, left)}px`, top: `${top}px` });
            };
            note.addEventListener('toggle', event => {
                if (event.newState === 'open') place();
                else byPointer = false;
            });
            const stay = () => clearTimeout(closing);
            const leave = () => {
                if (!byPointer) return;
                closing = setTimeout(() => { if (byPointer && note.matches(':popover-open')) note.hidePopover(); }, 200);
            };
            button.addEventListener('pointerenter', event => {
                if (event.pointerType !== 'mouse') return;
                stay();
                if (!note.matches(':popover-open')) { byPointer = true; note.showPopover(); }
            });
            button.addEventListener('click', event => {
                // A note the pointer opened is kept open by a click, not closed by it.
                if (byPointer && note.matches(':popover-open')) { event.preventDefault(); byPointer = false; }
            });
            for (const part of [button, note]) {
                part.addEventListener('pointerleave', leave);
                part.addEventListener('pointerenter', stay);
            }
            window.addEventListener('scroll', () => { if (note.matches(':popover-open')) place(); }, { passive: true });
            window.addEventListener('resize', () => { if (note.matches(':popover-open')) place(); });
        }
    };

    /**
     * The switch in a revised post's byline. Off, the post reads as it is now; the choice is kept for
     * every post. A link to a change in the list of revisions shows the changes on this page first
     * (without changing the choice), so it never lands where nothing is to be seen.
     */
    const wireRevisionToggle = () => {
        const toggle = document.getElementById('revision-toggle');
        if (!toggle) return;
        const show = (shown, keep) => {
            root.classList.toggle('changes-hidden', !shown);
            toggle.setAttribute('aria-checked', String(shown));
            if (!shown) for (const open of document.querySelectorAll('.revision-note:popover-open')) open.hidePopover();
            if (keep) { try { localStorage.setItem(CHANGES_KEY, shown ? 'shown' : 'hidden'); } catch { /* the choice lasts for this page */ } }
        };
        show(!root.classList.contains('changes-hidden'), false);
        toggle.hidden = false;
        toggle.addEventListener('click', () => show(toggle.getAttribute('aria-checked') !== 'true', true));
        for (const link of document.querySelectorAll('.article-revisions a[href^="#change-"]')) {
            link.addEventListener('click', () => { if (root.classList.contains('changes-hidden')) show(true, false); });
        }
    };

    const start = () => {
        wireThemeToggle();
        wireRevisionToggle();
        wireListFilter();
        wireDemos();
        wireRevisionNotes();
    };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
    else start();
})();
