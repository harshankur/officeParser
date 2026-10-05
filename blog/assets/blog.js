/**
 * What the blog's pages do once they are loaded: keep the reader's theme, filter the list by topic
 * and by the words searched for, and count a view. The pages themselves are built
 * (scripts/build-blog.js) and are complete without this script: every post is in its page, and the
 * list shows every post.
 */
(() => {
    'use strict';

    // ── Theme: set before the page is first painted (this script is in the head) ──
    const root = document.documentElement;
    const storedTheme = () => { try { return localStorage.getItem('theme'); } catch { return null; } };
    const saved = storedTheme() || 'auto';
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    root.setAttribute('data-theme', saved === 'auto' ? (prefersDark ? 'dark' : 'light') : saved);

    /** One view of the page, counted by the self-hosted view counter on the published site only. */
    const registerPageView = () => {
        if (location.hostname !== 'officeparser.harshankur.com') return;
        let sessionId;
        try {
            sessionId = sessionStorage.getItem('vc-session') || crypto.randomUUID();
            sessionStorage.setItem('vc-session', sessionId);
        } catch { sessionId = crypto.randomUUID(); }
        fetch('https://views.harshankur.com/registerView?' + new URLSearchParams({
            appId: 'officeparser',
            deviceSize: innerWidth < 768 ? 'small' : innerWidth < 1200 ? 'medium' : 'large',
            page: location.pathname,
            title: document.title.slice(0, 200),
            referrer: document.referrer,
            sessionId,
        }), { keepalive: true }).catch(() => { });
    };

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

    const start = () => {
        wireThemeToggle();
        wireListFilter();
        registerPageView();
    };

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
    else start();
})();
