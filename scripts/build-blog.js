/**
 * Blog Builder
 *
 * The blog is written as content and built into pages:
 *
 *   blog/posts.json           the blog's title and address, its topics, and every post's title, topic,
 *                             month and summary, in the order shown
 *   blog/posts/<slug>.html    a post's body: content only, no page shell, no styles, no scripts
 *                             (blog/README.md says what a post is written with)
 *   blog/templates/*.html     the page every post is set in (layout), the list page and a post's page
 *   blog/assets/              the stylesheet and the script the pages share
 *
 * and this writes the finished pages to docs/blog/, which is what is published: the list
 * (index.html), a page for each post (<slug>.html) with its own title, description and address in
 * its head, the shared assets, and what is worked out from the posts for those who do not read the
 * pages: the feed (feed.xml) and the text the list's search looks through (search.json). Beside
 * docs/blog/, at the site's root, it writes what search engines read of the whole site: the sitemap
 * (docs/sitemap.xml) and docs/robots.txt, which names it. Nothing in docs/blog/, and neither of those
 * two, is edited by hand: a change of content is
 * made in blog/posts, a change of look in blog/templates and blog/assets, and both reach every page
 * by building again. The same content always builds the same pages, byte for byte, so the test that
 * guards the blog (test/testBlog.js) can tell a page that was not rebuilt.
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SOURCE_DIR = path.join(ROOT, 'blog');
const OUTPUT_DIR = path.join(ROOT, 'docs', 'blog');
/** The site's root, which the blog is a folder of: where the site's sitemap and robots.txt stand. */
const SITE_DIR = path.join(ROOT, 'docs');

/** Words read in a minute, for a post's reading time. */
const WORDS_PER_MINUTE = 220;
/** A post's name in its address and its file: lower-case words joined by hyphens. */
const SLUG = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;
/** A link from one post to another, which names the post and not its address: `href="post:<slug>"`. */
const POST_LINK = /href="post:([^"]*)"/g;
/**
 * The elements a post's body is written with. A tag of any other name is an error: on a blog about
 * document formats it is nearly always sample markup someone forgot to escape (`<w:p>` for
 * `&lt;w:p&gt;`), which a browser takes for an element and shows nothing of. Each of these has its
 * look in blog/assets/blog.css: an element added here is given one there.
 */
const POST_ELEMENTS = new Set([
    'p', 'h2', 'h3', 'ul', 'ol', 'li', 'blockquote', 'pre', 'code', 'hr', 'br', 'div', 'span',
    'a', 'strong', 'em', 'sub', 'sup', 'img', 'figure', 'figcaption',
    'table', 'thead', 'tbody', 'tr', 'th', 'td',
]);
/**
 * The months, as a date is written for readers. They are not taken from the runtime (`Intl`), whose
 * names for them differ between versions: the same content builds the same pages on every one.
 */
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
/** The characters of the named references a post's text is likely to hold, by their code points. */
const ENTITIES = { lt: 0x3C, gt: 0x3E, amp: 0x26, quot: 0x22, apos: 0x27, nbsp: 0xA0, rarr: 0x2192, larr: 0x2190, hellip: 0x2026, ndash: 0x2013, times: 0xD7, middot: 0xB7 };
/** What a table's parts are, said on each: the stylesheet lays them out as blocks on a phone, which takes it away. */
const TABLE_ROLES = { table: 'table', thead: 'rowgroup', tbody: 'rowgroup', tr: 'row' };

// ---------------------------------------------------------------------------
// Templates
// ---------------------------------------------------------------------------

const escapeHtml = value => String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');

/**
 * A template as a tree: text, `{{name}}` (a value, escaped), `{{{name}}}` (a value, as it is),
 * `{{#name}}...{{/name}}` (its inside once for a value, once for each of a list's, not at all for
 * none) and `{{^name}}...{{/name}}` (its inside when there is none). A section's tag standing alone
 * on its line takes the line with it, so the output has no empty lines where the tags were.
 */
function parseTemplate(source, name) {
    const text = source.replace(/^[ \t]*(\{\{[#^/][^}]*\}\})[ \t]*\r?\n/gm, '$1');
    const tag = /\{\{\{\s*([\w-]+)\s*\}\}\}|\{\{\s*([#^/])?\s*([\w-]+)\s*\}\}/g;
    const root = { children: [] };
    const open = [root];
    let last = 0;
    for (let match = tag.exec(text); match; match = tag.exec(text)) {
        const parent = open[open.length - 1];
        if (match.index > last) parent.children.push({ type: 'text', value: text.slice(last, match.index) });
        last = tag.lastIndex;
        if (match[1]) parent.children.push({ type: 'raw', name: match[1] });
        else if (!match[2]) parent.children.push({ type: 'value', name: match[3] });
        else if (match[2] === '/') {
            if (open.length === 1 || parent.name !== match[3]) throw new Error(`${name}: {{/${match[3]}}} closes no section`);
            open.pop();
        } else {
            const section = { type: 'section', name: match[3], inverted: match[2] === '^', children: [] };
            parent.children.push(section);
            open.push(section);
        }
    }
    if (open.length > 1) throw new Error(`${name}: {{#${open[open.length - 1].name}}} is never closed`);
    if (last < text.length) root.children.push({ type: 'text', value: text.slice(last) });
    return { name, children: root.children };
}

/**
 * A template filled from `view`. A name is looked up in the section it is in, then in those around
 * it. A name nothing gives a value for is an error, not an empty string: a page with a hole in it
 * is not written.
 */
function renderTemplate(template, view) {
    const lookup = (scopes, key) => {
        for (let i = scopes.length - 1; i >= 0; i--) {
            if (Object.prototype.hasOwnProperty.call(scopes[i], key)) return scopes[i][key];
        }
        throw new Error(`${template.name}: nothing gives {{${key}}} a value`);
    };
    const render = (nodes, scopes) => nodes.map(node => {
        if (node.type === 'text') return node.value;
        const value = lookup(scopes, node.name);
        if (node.type === 'raw') return String(value);
        if (node.type === 'value') return escapeHtml(value);
        const none = value === null || value === undefined || value === false || (Array.isArray(value) && value.length === 0);
        if (node.inverted) return none ? render(node.children, scopes) : '';
        if (none) return '';
        if (Array.isArray(value)) return value.map(item => render(node.children, [...scopes, item])).join('');
        return render(node.children, typeof value === 'object' ? [...scopes, value] : scopes);
    }).join('');
    return render(template.children, [view]);
}

// ---------------------------------------------------------------------------
// Content
// ---------------------------------------------------------------------------

/**
 * A day a post names (`2026-10-04`), as its parts and as an instant (the start of that day, UTC). A
 * value that is not a day of the calendar (`2026-10`, `2026-02-30`) is an error naming where it is.
 */
function dayOf(value, where) {
    const parts = /^(\d{4})-(\d{2})-(\d{2})$/.exec(typeof value === 'string' ? value : '');
    const [year, month, day] = parts ? parts.slice(1).map(Number) : [];
    const instant = new Date(Date.UTC(year, month - 1, day));
    if (!parts || instant.getUTCFullYear() !== year || instant.getUTCMonth() !== month - 1 || instant.getUTCDate() !== day) {
        throw new Error(`${where}: "${value}" is not a date (YYYY-MM-DD)`);
    }
    return { year, month, day, instant };
}

/** A day as readers write it: `4 October 2026`, or `4 Oct 2026` on a card. */
const longDate = ({ year, month, day }) => `${day} ${MONTHS[month - 1]} ${year}`;
const shortDate = ({ year, month, day }) => `${day} ${MONTHS[month - 1].slice(0, 3)} ${year}`;

/** The minutes a post's body takes to read. */
function readingMinutes(body) {
    const text = body.replace(/<[^>]*>/g, ' ').replace(/&[a-z#0-9]+;/gi, ' ');
    return Math.max(1, Math.ceil(text.split(/\s+/).filter(Boolean).length / WORDS_PER_MINUTE));
}

/** The address of a post's page, from another page of the blog. */
const pageOf = slug => `./${slug}.html`;

/** A post's body as the text a reader sees, on one line: what the list's search looks through. */
function plainText(body) {
    return body.replace(/<[^>]*>/g, ' ')
        .replace(/&(?:#x([0-9a-f]+)|#(\d+)|([a-z]+));/gi, (reference, hex, decimal, named) => {
            const codePoint = hex ? parseInt(hex, 16) : decimal ? Number(decimal) : ENTITIES[named];
            return codePoint === undefined || codePoint > 0x10FFFF ? reference : String.fromCodePoint(codePoint);
        })
        .replace(/\s+/g, ' ').trim();
}

/**
 * A post's body as a reader outside the blog's pages is given it (a feed reader): every address in
 * it whole, since `./other-post.html`, `figure.png` and `#part` mean nothing away from the page.
 */
function standaloneBody(body, post, blogUrl) {
    return body.replace(/<(?:a|img)\b[^>]*>/g, tag => tag.replace(/\s(href|src)="([^"]*)"/g, (attribute, name, target) => {
        if (/^(?:https?:|mailto:|data:)/.test(target)) return attribute;
        return ` ${name}="${target.startsWith('#') ? post.address + target : blogUrl + target.replace(/^\.\//, '')}"`;
    }));
}

/** Text for a double-quoted attribute, from a piece of a post's body (which is HTML already: its `&` stay). */
const attributeText = html => html.replace(/<[^>]*>/g, '').replace(/\s+/g, ' ').trim().replace(/"/g, '&quot;');

/**
 * A post's tables as a page shows them. A post writes a plain table with its header row in `<thead>`.
 * On a phone the stylesheet lays a table out as one card a row, each cell under its column's name, so
 * every cell after a row's first is given that name (`data-label`), and every part is given its role,
 * which laying it out as a block takes away. A table this cannot label (no header row, merged cells, a
 * row with more cells than the header, a table in a table) is an error naming it.
 */
function decorateTables(body, name) {
    return body.replace(/<table\b([^>]*)>([\s\S]*?)<\/table>/g, (table, attributes, inner) => {
        if (/<table\b/.test(inner)) throw new Error(`${name}: a table inside a table is not laid out`);
        if (/<t[dh]\b[^>]*\s(?:col|row)span\s*=/.test(inner)) throw new Error(`${name}: a table with merged cells is not laid out`);
        const head = /<thead\b[^>]*>([\s\S]*?)<\/thead>/.exec(inner);
        if (!head) throw new Error(`${name}: a table needs its header row in <thead> (its cells name the columns on a phone)`);
        const labels = [...head[1].matchAll(/<th\b[^>]*>([\s\S]*?)<\/th>/g)].map(cell => attributeText(cell[1]));
        const withRole = (tag, role, rest = '') => `<${tag} role="${role}"${rest}`;
        // Each body row's cells, numbered: the first is the card's title, the others take their column's name.
        const rows = inner.replace(/(<tbody\b[^>]*>)([\s\S]*?)(<\/tbody>)/g, (_, open, bodyRows, close) => open + bodyRows.replace(/<tr\b[^>]*>[\s\S]*?<\/tr>/g, row => {
            let column = 0;
            return row.replace(/<(td|th)\b/g, (_, tag) => {
                const index = column++;
                if (index >= labels.length) throw new Error(`${name}: a table row has more cells than its header row (${labels.length})`);
                return withRole(tag, tag === 'th' ? 'rowheader' : 'cell', index > 0 ? ` data-label="${labels[index]}"` : '');
            });
        }) + close);
        const parts = rows
            .replace(/(<thead\b[^>]*>)([\s\S]*?)(<\/thead>)/, (_, open, headRows, close) => open + headRows.replace(/<th\b/g, withRole('th', 'columnheader')) + close)
            .replace(/<(thead|tbody|tr)\b/g, (_, tag) => withRole(tag, TABLE_ROLES[tag]));
        return `${withRole('table', TABLE_ROLES.table)}${attributes}>${parts}</table>`;
    });
}

/**
 * A post's body as its page holds it: its links to other posts given their addresses, and its tables
 * what a page shows (see decorateTables). What a page could not show as written is an error naming
 * it, so a mistake fails the build and is not found by a reader: a tag that is not one a post is
 * written with, a class the stylesheet does not have, a role or a column name written by hand (the
 * build writes them), a link to a post there is not, and a picture or a file that is not in
 * blog/assets.
 *
 * @param {string} body
 * @param {{ name: string, slugs: Set<string>, assets: Set<string>, classes: Set<string> }} blog
 */
function preparePostBody(body, { name, slugs, assets, classes }) {
    for (const [, tag] of body.matchAll(/<\/?([A-Za-z][^\s>/]*)/g)) {
        if (!POST_ELEMENTS.has(tag.toLowerCase())) {
            throw new Error(`${name}: <${tag}> is not an element a post is written with. Sample markup is written escaped: &lt;${tag}&gt;`);
        }
    }
    if (/\s(?:role|data-label)\s*=/.test(body.replace(/<code\b[\s\S]*?<\/code>/g, ''))) {
        throw new Error(`${name}: role and data-label are written by the build. A post writes a plain table`);
    }
    for (const [, list] of body.matchAll(/\sclass="([^"]*)"/g)) {
        for (const used of list.split(/\s+/).filter(Boolean)) {
            if (!classes.has(used)) throw new Error(`${name}: the class "${used}" is not in blog/assets/blog.css`);
        }
    }
    // A link to another post names the post. Its address is decided here, in one place.
    const linked = body.replace(POST_LINK, (_, target) => {
        if (!slugs.has(target)) throw new Error(`${name}: links to the post "${target}", which there is not`);
        return `href="${pageOf(target)}"`;
    });
    // Anything else a post points at is on the web, in the page (`#part`), or a file of blog/assets.
    for (const [, attribute, target] of linked.matchAll(/<(?:a|img)\b[^>]*?\s(href|src)="([^"]*)"/g)) {
        if (/^(?:https?:|mailto:|#)/.test(target) || (attribute === 'src' && target.startsWith('data:'))) continue;
        const file = target.replace(/^\.\//, '');
        if (attribute === 'href' && slugs.has(file.replace(/\.html$/, '')) && file.endsWith('.html')) continue;
        if (!assets.has(file)) throw new Error(`${name}: ${attribute}="${target}" is not a file in blog/assets`);
    }
    return decorateTables(linked, name);
}

/**
 * The blog as its pages show it: posts.json, with each post's body, reading time, topic's name and
 * dates worked out. What cannot be built (a post with no body, a link to a post there is not, a date
 * that is not one) is an error naming it.
 */
function readBlog(sourceDir = SOURCE_DIR) {
    const { blog, categories, posts } = JSON.parse(fs.readFileSync(path.join(sourceDir, 'posts.json'), 'utf8'));
    const labelOf = new Map(categories.map(category => [category.id, category.label]));
    const slugs = new Set(posts.map(post => post.slug));
    if (slugs.size !== posts.length) throw new Error('posts.json: a slug is used twice');
    const assets = new Set(fs.readdirSync(path.join(sourceDir, 'assets')));
    const classes = new Set([...fs.readFileSync(path.join(sourceDir, 'assets', 'blog.css'), 'utf8').matchAll(/\.([A-Za-z_][\w-]*)/g)].map(match => match[1]));

    const read = posts.map(post => {
        if (!SLUG.test(post.slug)) throw new Error(`posts.json: "${post.slug}" is not a slug (lower-case words joined by hyphens)`);
        if (!labelOf.has(post.category)) throw new Error(`posts.json: ${post.slug} has the topic "${post.category}", which is not listed`);
        const bodyPath = path.join(sourceDir, 'posts', `${post.slug}.html`);
        if (!fs.existsSync(bodyPath)) throw new Error(`blog/posts/${post.slug}.html: the post is listed and has no body`);
        const body = preparePostBody(fs.readFileSync(bodyPath, 'utf8').trimEnd(), { name: `blog/posts/${post.slug}.html`, slugs, assets, classes });
        // The day a post went out, and the day it was last changed when it has been since.
        const published = dayOf(post.published, `posts.json: ${post.slug}: published`);
        const updated = post.updated === undefined ? null : dayOf(post.updated, `posts.json: ${post.slug}: updated`);
        if (updated && updated.instant < published.instant) throw new Error(`posts.json: ${post.slug} was updated (${post.updated}) before it was published (${post.published})`);
        return {
            ...post,
            body,
            text: plainText(body),
            link: pageOf(post.slug),
            address: `${blog.url}${post.slug}.html`,
            categoryLabel: labelOf.get(post.category),
            dateLong: longDate(published),
            dateShort: shortDate(published),
            publishedAt: published.instant,
            updated: updated ? post.updated : null,
            updatedLong: updated ? longDate(updated) : null,
            lastChanged: updated ? post.updated : post.published,
            minutes: readingMinutes(body),
        };
    });
    return { blog, categories, posts: read };
}

// ---------------------------------------------------------------------------
// Pages
// ---------------------------------------------------------------------------

/** The day the blog last changed: the latest any post was published or updated (`YYYY-MM-DD` sorts as dates do). */
const lastChanged = posts => posts.reduce((latest, post) => (post.lastChanged > latest ? post.lastChanged : latest), '');

/**
 * The feed (RSS 2.0): every post, in the list's order, with its whole body, so a feed reader shows
 * the post and not a teaser. Its dates are the posts' own, never the time of the build: the same
 * content builds the same feed.
 */
function renderFeed(blog, posts) {
    const item = post => `    <item>
      <title>${escapeHtml(post.title)}</title>
      <link>${escapeHtml(post.address)}</link>
      <guid isPermaLink="true">${escapeHtml(post.address)}</guid>
      <pubDate>${post.publishedAt.toUTCString()}</pubDate>
      <dc:creator>${escapeHtml(blog.author)}</dc:creator>
      <category>${escapeHtml(post.categoryLabel)}</category>
      <description>${escapeHtml(post.description)}</description>
      <content:encoded><![CDATA[${standaloneBody(post.body, post, blog.url).replace(/]]>/g, ']]]]><![CDATA[>')}]]></content:encoded>
    </item>
`;
    const latest = lastChanged(posts);
    return `<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom" xmlns:content="http://purl.org/rss/1.0/modules/content/" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <channel>
    <title>${escapeHtml(blog.title)}</title>
    <link>${escapeHtml(blog.url)}</link>
    <description>${escapeHtml(blog.description)}</description>
    <language>en</language>
    <lastBuildDate>${dayOf(latest, 'posts.json').instant.toUTCString()}</lastBuildDate>
    <atom:link href="${escapeHtml(blog.url)}feed.xml" rel="self" type="application/rss+xml"/>
${posts.map(item).join('')}  </channel>
</rss>
`;
}

/** The address of the site the blog is a folder of: `https://example.com/` for a blog at `https://example.com/blog/`. */
const siteOf = blog => new URL('../', blog.url).href;

/**
 * The site's sitemap: every page a reader can land on. That is the home page, the blog's list and
 * every post. It stands at the site's root, because a sitemap may only name pages in its own folder
 * or below (sitemaps.org, "Sitemap file location"): in the blog's folder it could not name the home
 * page. A page of another site cannot be in it (the README and the changelog are on GitHub), and the
 * files under docs/specs/ are parts the home page loads, not pages. A post and the list have the day
 * they last changed; the home page has none, since nothing here knows it, and a wrong one is worse
 * than none.
 */
function renderSitemap(blog, posts) {
    const url = (address, day) => `  <url>\n    <loc>${escapeHtml(address)}</loc>\n${day ? `    <lastmod>${day}</lastmod>\n` : ''}  </url>\n`;
    return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${url(siteOf(blog))}${url(blog.url, lastChanged(posts))}${posts.map(post => url(post.address, post.lastChanged)).join('')}</urlset>
`;
}

/** The site's robots.txt: everything may be read, and the sitemap is where search engines are told it is. */
const renderRobots = blog => `User-agent: *\nAllow: /\n\nSitemap: ${siteOf(blog)}sitemap.xml\n`;

/**
 * The files written at the site's root, by their names in docs/: the sitemap and robots.txt.
 *
 * @returns {Map<string, string>}
 */
function renderSiteRoot(sourceDir = SOURCE_DIR) {
    const { blog, posts } = readBlog(sourceDir);
    return new Map([['sitemap.xml', renderSitemap(blog, posts)], ['robots.txt', renderRobots(blog)]]);
}

/**
 * Every file of the published blog, by its name in docs/blog/: the list, each post's page, the feed,
 * the text the search looks through, and the shared assets.
 *
 * @returns {Map<string, string | Buffer>}
 */
function renderSite(sourceDir = SOURCE_DIR) {
    const template = name => parseTemplate(fs.readFileSync(path.join(sourceDir, 'templates', name), 'utf8'), `blog/templates/${name}`);
    const layout = template('layout.html');
    const list = template('list.html');
    const postPage = template('post.html');
    const { blog, categories, posts } = readBlog(sourceDir);
    const files = new Map();

    files.set('index.html', renderTemplate(layout, {
        pageTitle: blog.title,
        shareTitle: blog.title,
        shareType: 'website',
        description: blog.description,
        address: blog.url,
        isPost: false,
        published: null,
        updated: null,
        main: renderTemplate(list, { categories, posts }).trimEnd(),
    }));

    posts.forEach((post, index) => {
        files.set(`${post.slug}.html`, renderTemplate(layout, {
            pageTitle: `${post.title} | officeParser Blog`,
            shareTitle: post.title,
            shareType: 'article',
            description: post.description,
            address: post.address,
            isPost: true,
            published: post.published,
            updated: post.updated,
            main: renderTemplate(postPage, { ...post, previous: posts[index - 1] ?? null, next: posts[index + 1] ?? null }).trimEnd(),
        }));
    });

    files.set('feed.xml', renderFeed(blog, posts));
    // The list's search looks through each post's text, which the list's page does not hold.
    files.set('search.json', `${JSON.stringify({ posts: posts.map(post => ({ slug: post.slug, text: post.text })) })}\n`);

    const assets = path.join(sourceDir, 'assets');
    for (const entry of fs.readdirSync(assets, { withFileTypes: true }).sort((a, b) => (a.name < b.name ? -1 : 1))) {
        // A page and its assets are published side by side, so an asset is a file, and not one named as a built file is.
        if (!entry.isFile()) throw new Error(`blog/assets/${entry.name}: blog/assets holds files only`);
        if (files.has(entry.name)) throw new Error(`blog/assets/${entry.name}: a file the build writes has this name`);
        files.set(entry.name, fs.readFileSync(path.join(assets, entry.name)));
    }
    return files;
}

/**
 * Writes the published blog, and removes from it what the content no longer has; and writes the
 * site's sitemap and robots.txt at the site's root, where nothing else is touched.
 */
function build(sourceDir = SOURCE_DIR, outputDir = OUTPUT_DIR, siteDir = SITE_DIR) {
    const files = renderSite(sourceDir);
    const rootFiles = renderSiteRoot(sourceDir);
    fs.mkdirSync(outputDir, { recursive: true });
    for (const entry of fs.readdirSync(outputDir)) {
        if (!files.has(entry)) fs.rmSync(path.join(outputDir, entry), { recursive: true, force: true });
    }
    for (const [name, content] of files) fs.writeFileSync(path.join(outputDir, name), content);
    for (const [name, content] of rootFiles) fs.writeFileSync(path.join(siteDir, name), content);
    return files;
}

module.exports = { SOURCE_DIR, OUTPUT_DIR, SITE_DIR, parseTemplate, renderTemplate, preparePostBody, plainText, standaloneBody, readBlog, renderSite, renderSiteRoot, build };

if (require.main === module) {
    try {
        const files = build();
        const pages = [...files.keys()].filter(name => name.endsWith('.html'));
        console.log(`Blog built → docs/blog/ (${pages.length - 1} posts, ${files.size} files), docs/sitemap.xml, docs/robots.txt`);
    } catch (error) {
        console.error(`Blog build failed: ${error.message}`);
        process.exit(1);
    }
}
