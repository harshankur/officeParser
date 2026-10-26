/**
 * The repository's documents (the readme, the changelog, how to contribute, the security policy, where
 * to get support, the code of conduct and the licence), read on the site: a link to one of them on
 * GitHub opens it here, in a reader with a tab for each document and its contents beside it, instead of
 * sending the visitor away. The site's copies are in repo/ (see scripts/sync-site-documents.js).
 *
 * A link opens the reader where it is a plain click on a link to one of the documents on GitHub
 * (github.com/harshankur/officeParser/blob/master/<document>, with its section, if any). A click with a
 * modifier key, or a middle click, opens GitHub as the link says. Each document has its own address
 * (#doc/changelog, #doc/readme/installation), so a document can be linked to and the back button closes it.
 */
(() => {
    const REPO = 'https://github.com/harshankur/officeParser';
    const BLOB = `${REPO}/blob/master/`;
    const RAW = 'https://raw.githubusercontent.com/harshankur/officeParser/master/';
    const DOCUMENTS = [
        { slug: 'readme', file: 'README.md', title: 'Readme', depth: 3 },
        { slug: 'changelog', file: 'CHANGELOG.md', title: 'Changelog', depth: 2 },
        { slug: 'contributing', file: 'CONTRIBUTING.md', title: 'Contributing', depth: 3 },
        { slug: 'security', file: 'SECURITY.md', title: 'Security', depth: 3 },
        { slug: 'support', file: 'SUPPORT.md', title: 'Support', depth: 3 },
        { slug: 'code-of-conduct', file: 'CODE_OF_CONDUCT.md', title: 'Code of Conduct', depth: 3 },
        { slug: 'license', file: 'LICENSE', served: 'LICENSE.txt', title: 'MIT License', plain: true },
    ];
    const byFile = new Map(DOCUMENTS.map(d => [d.file.toLowerCase(), d]));
    const bySlug = new Map(DOCUMENTS.map(d => [d.slug, d]));
    const texts = new Map();

    const ICONS = {
        close: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M18 6 6 18M6 6l12 12"/></svg>',
        github: '<svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true"><path d="M12 .5a11.5 11.5 0 0 0-3.64 22.41c.58.1.79-.25.79-.56v-2c-3.2.7-3.88-1.37-3.88-1.37-.53-1.33-1.28-1.69-1.28-1.69-1.05-.71.08-.7.08-.7 1.16.08 1.77 1.19 1.77 1.19 1.03 1.77 2.7 1.26 3.36.96.1-.75.4-1.26.73-1.55-2.55-.29-5.24-1.28-5.24-5.69 0-1.26.45-2.29 1.19-3.1-.12-.29-.52-1.46.11-3.05 0 0 .97-.31 3.17 1.18a11 11 0 0 1 5.77 0c2.2-1.49 3.17-1.18 3.17-1.18.63 1.59.23 2.76.11 3.05.74.81 1.19 1.84 1.19 3.1 0 4.42-2.7 5.4-5.26 5.68.41.36.78 1.06.78 2.14v3.17c0 .31.21.67.8.56A11.5 11.5 0 0 0 12 .5Z"/></svg>',
        list: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>',
        top: '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m18 15-6-6-6 6"/></svg>',
    };
    const ALERTS = {
        NOTE: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>',
        TIP: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M9 18h6M10 22h4M12 2a7 7 0 0 0-4 12.74V17h8v-2.26A7 7 0 0 0 12 2z"/></svg>',
        IMPORTANT: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10zM12 8v4M12 16h.01"/></svg>',
        WARNING: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3ZM12 9v4M12 17h.01"/></svg>',
        CAUTION: '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg>',
    };

    const escapeHtml = text => text.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

    // ── The reader ──────────────────────────────────────────────────────────────────────────────
    const dialog = document.createElement('dialog');
    dialog.className = 'repo-doc';
    dialog.setAttribute('aria-labelledby', 'repo-doc-title');
    dialog.innerHTML = `
        <div class="repo-doc-panel">
            <div class="repo-doc-header">
                <div class="repo-doc-heading">
                    <span class="repo-doc-kicker">officeParser</span>
                    <h2 class="repo-doc-title" id="repo-doc-title"></h2>
                </div>
                <div class="repo-doc-actions">
                    <button type="button" class="repo-doc-icon-btn repo-doc-toc-toggle" aria-expanded="false" aria-controls="repo-doc-toc" title="Contents">${ICONS.list}<span>Contents</span></button>
                    <a class="repo-doc-icon-btn repo-doc-github" target="_blank" rel="noopener" title="View on GitHub">${ICONS.github}<span>GitHub</span></a>
                    <button type="button" class="repo-doc-icon-btn repo-doc-close" title="Close (Esc)" aria-label="Close">${ICONS.close}</button>
                </div>
            </div>
            <nav class="repo-doc-tabs" aria-label="Documents">
                ${DOCUMENTS.map(d => `<a class="repo-doc-tab" href="#doc/${d.slug}" data-slug="${d.slug}">${escapeHtml(d.title)}</a>`).join('')}
            </nav>
            <div class="repo-doc-body">
                <aside class="repo-doc-toc" id="repo-doc-toc" aria-label="Contents"></aside>
                <div class="repo-doc-scroll" tabindex="-1">
                    <article class="doc-body repo-doc-article"></article>
                    <button type="button" class="repo-doc-top" title="Back to the top" aria-label="Back to the top">${ICONS.top}</button>
                </div>
            </div>
        </div>`;
    document.body.appendChild(dialog);
    const $ = selector => dialog.querySelector(selector);
    const titleEl = $('.repo-doc-title'), article = $('.repo-doc-article'), toc = $('.repo-doc-toc');
    const scroller = $('.repo-doc-scroll'), github = $('.repo-doc-github'), tocToggle = $('.repo-doc-toc-toggle'), topButton = $('.repo-doc-top');

    let current = null, opener = null, pushed = false, markReading = null, frame = 0;

    /** GitHub's id for a heading: lower case, punctuation left out, each space a hyphen; repeats numbered. */
    const slugger = () => {
        const seen = new Map();
        return text => {
            const base = text.trim().toLowerCase().replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-');
            const n = seen.get(base) ?? 0;
            seen.set(base, n + 1);
            return n ? `${base}-${n}` : base;
        };
    };

    /** A document's text, fetched once. */
    async function textOf(doc) {
        if (!texts.has(doc.slug)) {
            texts.set(doc.slug, fetch(`repo/${doc.served || doc.file}`).then(r => {
                if (!r.ok) throw new Error(`${r.status} ${r.statusText}`);
                return r.text();
            }).catch(error => { texts.delete(doc.slug); throw error; }));
        }
        return texts.get(doc.slug);
    }

    /** Where a link in a document goes: another document here, a place in this one, or GitHub. */
    function retarget(root) {
        for (const a of root.querySelectorAll('a[href]')) {
            const href = a.getAttribute('href');
            if (href.startsWith('#')) { a.dataset.anchor = href.slice(1); continue; }
            if (/^[a-z][a-z0-9+.-]*:/i.test(href) || href.startsWith('//')) { a.target = '_blank'; a.rel = 'noopener'; continue; }
            const [file, anchor] = href.replace(/^\.\//, '').split('#');
            const doc = byFile.get(file.toLowerCase());
            if (doc) { a.href = BLOB + doc.file + (anchor ? `#${anchor}` : ''); continue; }
            a.href = BLOB + file + (anchor ? `#${anchor}` : '');
            a.target = '_blank';
            a.rel = 'noopener';
        }
        for (const img of root.querySelectorAll('img[src]')) {
            const src = img.getAttribute('src');
            if (!/^[a-z][a-z0-9+.-]*:/i.test(src) && !src.startsWith('//')) img.src = RAW + src.replace(/^\.\//, '');
            img.loading = 'lazy';
        }
    }

    /** GitHub's alerts (`> [!NOTE]`), as the site's spec pages show them. */
    function alerts(root) {
        for (const quote of root.querySelectorAll('blockquote')) {
            const first = quote.querySelector('p');
            const match = first && /^\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*/.exec(first.innerHTML);
            if (!match) continue;
            first.innerHTML = first.innerHTML.slice(match[0].length).replace(/^<br>\s*/, '');
            if (!first.innerHTML.trim()) first.remove();
            const callout = document.createElement('div');
            callout.className = `callout callout-${match[1].toLowerCase()}`;
            callout.innerHTML = `<div class="callout-header"><span class="callout-icon">${ALERTS[match[1]]}</span>${match[1]}</div><div class="callout-content"></div>`;
            callout.lastChild.append(...quote.childNodes);
            quote.replaceWith(callout);
        }
    }

    /** The document's headings, each with its id, and the contents beside it. */
    function contents(doc) {
        const slug = slugger();
        const headings = [...article.querySelectorAll('h1, h2, h3, h4, h5, h6')];
        for (const h of headings) {
            h.id = slug(h.textContent);
            const link = document.createElement('a');
            link.className = 'repo-doc-permalink';
            link.href = `#doc/${doc.slug}/${h.id}`;
            link.setAttribute('aria-label', `Link to ${h.textContent}`);
            link.textContent = '#';
            h.appendChild(link);
        }
        const listed = headings.filter(h => h.tagName === 'H2' || (doc.depth >= 3 && h.tagName === 'H3'));
        toc.innerHTML = listed.length
            ? `<p class="repo-doc-toc-title">On this page</p><ol>${listed.map(h => `<li class="repo-doc-toc-${h.tagName.toLowerCase()}"><a href="#doc/${doc.slug}/${h.id}" data-anchor="${h.id}">${escapeHtml(h.textContent.replace(/#$/, '').trim())}</a></li>`).join('')}</ol>`
            : '';
        dialog.classList.toggle('repo-doc-has-toc', listed.length > 1);
        // The section being read (the last listed heading above the reading line), marked in the contents.
        markReading = null;
        if (listed.length > 1) {
            const links = [...toc.querySelectorAll('a')];
            let shown = -1;
            markReading = () => {
                const line = scroller.getBoundingClientRect().top + 96;
                let lo = 0, hi = listed.length - 1, at = -1;
                while (lo <= hi) { const mid = (lo + hi) >> 1; if (listed[mid].getBoundingClientRect().top <= line) { at = mid; lo = mid + 1; } else hi = mid - 1; }
                if (at === shown) return;
                links[shown]?.classList.remove('active');
                shown = at;
                const active = links[at];
                if (!active) return;
                active.classList.add('active');
                // Kept in view in the contents, without moving anything else.
                const top = active.offsetTop - toc.offsetTop, bottom = top + active.offsetHeight;
                if (top < toc.scrollTop + 40) toc.scrollTop = Math.max(0, top - 40);
                else if (bottom > toc.scrollTop + toc.clientHeight - 40) toc.scrollTop = bottom - toc.clientHeight + 40;
            };
        }
    }

    /** Scrolls to a heading of the document shown (its id as GitHub gives it), or to the top. */
    function scrollToAnchor(anchor, smooth) {
        const target = anchor ? article.querySelector(`[id="${CSS.escape(decodeURIComponent(anchor))}"]`) : null;
        if (target) target.scrollIntoView({ behavior: smooth ? 'smooth' : 'auto', block: 'start' });
        else scroller.scrollTo({ top: 0, behavior: smooth ? 'smooth' : 'auto' });
    }

    async function show(doc, anchor) {
        for (const tab of dialog.querySelectorAll('.repo-doc-tab')) {
            const on = tab.dataset.slug === doc.slug;
            tab.classList.toggle('active', on);
            if (on) { tab.setAttribute('aria-current', 'page'); tab.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
            else tab.removeAttribute('aria-current');
        }
        titleEl.textContent = doc.title;
        github.href = BLOB + doc.file + (anchor ? `#${anchor}` : '');
        dialog.classList.remove('repo-doc-toc-open');
        tocToggle.setAttribute('aria-expanded', 'false');
        if (current === doc) { scrollToAnchor(anchor, true); return; }
        current = doc;
        article.innerHTML = '<div class="repo-doc-loading" role="status"><span></span><span></span><span></span></div>';
        toc.innerHTML = '';
        try {
            const text = await textOf(doc);
            if (current !== doc) return;
            article.classList.toggle('repo-doc-plain', !!doc.plain);
            article.innerHTML = doc.plain || !window.marked
                ? `<pre class="repo-doc-text">${escapeHtml(text)}</pre>`
                : window.marked.parse(text, { gfm: true });
            alerts(article);
            retarget(article);
            if (window.hljs) for (const code of article.querySelectorAll('pre code')) { try { window.hljs.highlightElement(code); } catch { /* left plain */ } }
            for (const table of article.querySelectorAll('table')) {
                const wrap = document.createElement('div');
                wrap.className = 'repo-doc-table';
                table.replaceWith(wrap);
                wrap.appendChild(table);
            }
            contents(doc);
            scrollToAnchor(anchor, false);
            markReading?.();
        } catch (error) {
            if (current !== doc) return;
            current = null;
            article.innerHTML = `<div class="repo-doc-error"><p>This document could not be loaded (${escapeHtml(String(error.message || error))}).</p><p><a href="${BLOB + doc.file}" target="_blank" rel="noopener">Read it on GitHub</a></p></div>`;
        }
    }

    function open(doc, anchor, viaHash) {
        if (!dialog.open) {
            opener = document.activeElement;
            dialog.showModal();
            document.documentElement.classList.add('repo-doc-locked');
        }
        const hash = `#doc/${doc.slug}${anchor ? `/${anchor}` : ''}`;
        if (!viaHash && location.hash !== hash) {
            // The first document opened is a step in the history (the back button closes the reader); moving
            // between documents and sections takes its place.
            if (pushed) history.replaceState(null, '', hash); else history.pushState(null, '', hash);
            pushed = true;
        }
        show(doc, anchor);
    }

    function close() {
        if (!dialog.open) return;
        dialog.close();
    }

    dialog.addEventListener('close', () => {
        document.documentElement.classList.remove('repo-doc-locked');
        if (location.hash.startsWith('#doc/')) {
            if (pushed) history.back();
            else history.replaceState(null, '', location.pathname + location.search);
        }
        pushed = false;
        if (opener && document.contains(opener)) opener.focus({ preventScroll: true });
    });

    /** The document and section a link to GitHub names, or none. */
    function linked(a) {
        let url;
        try { url = new URL(a.href, location.href); } catch { return null; }
        if (url.origin + url.pathname.replace(/\/$/, '') === REPO) return null;
        if (!(url.origin + url.pathname).startsWith(BLOB)) return null;
        const doc = byFile.get(decodeURIComponent(url.pathname.slice(new URL(BLOB).pathname.length)).toLowerCase());
        return doc ? { doc, anchor: url.hash.slice(1) || '' } : null;
    }

    document.addEventListener('click', event => {
        if (event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey || event.shiftKey || event.altKey) return;
        const a = event.target.closest?.('a[href]');
        if (!a) return;
        // In the reader: a place in the document, the contents, or another document.
        if (dialog.contains(a)) {
            if (a.dataset.anchor !== undefined && current) {
                event.preventDefault();
                open(current, a.dataset.anchor);
                return;
            }
            if (a.classList.contains('repo-doc-tab') || a.classList.contains('repo-doc-permalink')) {
                event.preventDefault();
                const [slug, anchor] = a.getAttribute('href').slice('#doc/'.length).split('/');
                const doc = bySlug.get(slug);
                if (doc) open(doc, anchor || '');
                return;
            }
        }
        const target = linked(a);
        if (!target) return;
        event.preventDefault();
        open(target.doc, target.anchor);
    });

    // A click on the backdrop (outside the panel) closes the reader.
    dialog.addEventListener('click', event => { if (event.target === dialog) close(); });
    $('.repo-doc-close').addEventListener('click', close);
    tocToggle.addEventListener('click', () => {
        const openToc = !dialog.classList.contains('repo-doc-toc-open');
        dialog.classList.toggle('repo-doc-toc-open', openToc);
        tocToggle.setAttribute('aria-expanded', String(openToc));
    });
    topButton.addEventListener('click', () => scroller.scrollTo({ top: 0, behavior: 'smooth' }));
    scroller.addEventListener('scroll', () => {
        topButton.classList.toggle('visible', scroller.scrollTop > 600);
        if (markReading && !frame) frame = setTimeout(() => { frame = 0; markReading?.(); }, 60);
    }, { passive: true });

    /** The reader follows the address: #doc/<document>[/<section>] opens it, any other closes it. */
    function route() {
        const match = /^#doc\/([^/]+)(?:\/(.*))?$/.exec(location.hash);
        const doc = match && bySlug.get(match[1]);
        if (doc) open(doc, match[2] || '', true);
        else if (dialog.open) { pushed = false; dialog.close(); }
    }
    window.addEventListener('hashchange', route);
    route();
})();
