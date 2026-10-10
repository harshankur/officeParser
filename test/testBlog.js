/**
 * Blog Validation
 *
 * The blog is written as content (blog/) and built into pages (docs/blog/, by
 * scripts/build-blog.js). A post is an entry in `blog/posts.json` and a body in
 * `blog/posts/<slug>.html`, content only: the templates and the shared stylesheet decide how every
 * post looks, so a change of look touches no post.
 *
 * That only holds while a post stays content, and while what is published is what the content
 * builds. These checks fail when a post brings its own page shell, styles or scripts; when the
 * content cannot be built (a listed post with no body, a link to a post there is not, sample markup
 * left unescaped, a table the build cannot label, a picture that is not in blog/assets); and when a
 * page in docs/blog/ is not the one the content builds (edited by hand, or not rebuilt after a
 * change). The feed and the site's sitemap, which the build writes too, are read as XML, as a feed
 * reader and a search engine read them. They read files only: no browser is needed.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { DOMParser } = require('@xmldom/xmldom');
const { SOURCE_DIR, OUTPUT_DIR, SITE_DIR, parseTemplate, renderTemplate, preparePostBody, plainText, standaloneBody, renderSite, renderSiteRoot } = require('../scripts/build-blog.js');

let passed = 0;
let failed = 0;
const check = (name, ok, detail = '') => {
    if (ok) { passed++; return; }
    failed++;
    console.log(`  ✗ FAIL: ${name}${detail ? ` (${detail})` : ''}`);
};
const throws = run => { try { run(); return false; } catch { return true; } };

/** Elements a post's body has no business holding: they belong to the page, not to a post. */
const SHELL_ELEMENTS = ['html', 'head', 'body', 'script', 'style', 'link', 'meta', 'title', 'header', 'footer', 'nav', 'main', 'iframe'];
const SLUG = /^[a-z0-9]+(?:-[a-z0-9]+)*$/;

// ── The templates' own rules, against fixed inputs ──
const template = source => parseTemplate(source, 'test');
check('template: a value is escaped, a raw value is not', renderTemplate(template('<p title="{{t}}">{{{b}}}</p>'), { t: 'a "b" <c>', b: '<i>x</i>' }) === '<p title="a &quot;b&quot; &lt;c&gt;"><i>x</i></p>');
check('template: a section repeats for a list and reads its items', renderTemplate(template('{{#items}}[{{name}}:{{outer}}]{{/items}}'), { outer: 'o', items: [{ name: 'a' }, { name: 'b' }] }) === '[a:o][b:o]');
check('template: a section is skipped for none, and its inverse is shown', renderTemplate(template('{{#next}}N{{/next}}{{^next}}none{{/next}}'), { next: null }) === 'none');
check('template: a value written in is not read as a template', renderTemplate(template('{{{body}}}'), { body: '{{name}}' }) === '{{name}}');
check('template: a section tag alone on its line leaves no empty line', renderTemplate(template('a\n  {{#on}}\nb\n  {{/on}}\nc\n'), { on: true }) === 'a\nb\nc\n');
check('template: a name with no value is an error', throws(() => renderTemplate(template('{{missing}}'), {})));
check('template: an unclosed section is an error', throws(() => template('{{#open}}x')));

// ── What the build makes of a post's body, and what it refuses, against fixed inputs ──
const sample = { name: 'sample', slugs: new Set(['other-post']), assets: new Set(['blog.css', 'figure.png']), classes: new Set(['callout', 'callout-title']) };
const prepared = body => preparePostBody(body, sample);
const refused = body => { try { prepared(body); return ''; } catch (error) { return error.message; } };
check('body: a plain table is given its parts\' roles, and each cell after a row\'s first its column\'s name',
    prepared('<table><thead><tr><th>Name</th><th>What <code>it</code> "is"</th></tr></thead><tbody><tr><td>a</td><td class="callout">b</td></tr></tbody></table>')
    === '<table role="table"><thead role="rowgroup"><tr role="row"><th role="columnheader">Name</th><th role="columnheader">What <code>it</code> "is"</th></tr></thead>'
    + '<tbody role="rowgroup"><tr role="row"><td role="cell">a</td><td role="cell" data-label="What it &quot;is&quot;" class="callout">b</td></tr></tbody></table>');
check('body: a row header in the body is one', prepared('<table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><th>a</th><td>b</td></tr></tbody></table>').includes('<th role="rowheader">a</th><td role="cell" data-label="B">b</td>'));
check('body: text that looks like a table part is left alone', prepared('<p><code>&lt;td colspan="2"&gt;</code> and rowspan="N"</p>') === '<p><code>&lt;td colspan="2"&gt;</code> and rowspan="N"</p>');
check('body: a link to a post is given its address', prepared('<a href="post:other-post">x</a>') === '<a href="./other-post.html">x</a>');
check('body: a picture in blog/assets, a web address and a place in the page are kept', prepared('<img src="figure.png" alt="f"><a href="https://example.com/">w</a><a href="#part">p</a>') === '<img src="figure.png" alt="f"><a href="https://example.com/">w</a><a href="#part">p</a>');
for (const [what, body, said] of [
    ['a table with no header row', '<table><tbody><tr><td>a</td></tr></tbody></table>', /header row in <thead>/],
    ['a table with merged cells', '<table><thead><tr><th>A</th><th>B</th></tr></thead><tbody><tr><td colspan="2">a</td></tr></tbody></table>', /merged cells/],
    ['a row with more cells than the header', '<table><thead><tr><th>A</th></tr></thead><tbody><tr><td>a</td><td>b</td></tr></tbody></table>', /more cells than its header row/],
    ['a table in a table', '<table><thead><tr><th>A</th></tr></thead><tbody><tr><td><table><thead><tr><th>B</th></tr></thead></table></td></tr></tbody></table>', /table inside a table/],
    ['sample markup left unescaped', '<pre><code><w:p>text</w:p></code></pre>', /<w:p> is not an element.*&lt;w:p&gt;/],
    ['a page element', '<section><p>x</p></section>', /<section> is not an element/],
    ['a role written by hand', '<table role="table"><thead><tr><th>A</th></tr></thead></table>', /written by the build/],
    ['a column name written by hand', '<table><thead><tr><th>A</th></tr></thead><tbody><tr><td data-label="A">a</td></tr></tbody></table>', /written by the build/],
    ['a class the stylesheet does not have', '<div class="calout">x</div>', /"calout" is not in blog\/assets\/blog.css/],
    ['a link to a post there is not', '<a href="post:no-such-post">x</a>', /post "no-such-post", which there is not/],
    ['a picture that is not in blog/assets', '<img src="missing.png" alt="m">', /src="missing.png" is not a file in blog\/assets/],
    ['a link to a file that is not in blog/assets', '<a href="notes.pdf">n</a>', /href="notes.pdf" is not a file in blog\/assets/],
]) check(`body: ${what} is refused, and said`, said.test(refused(body)), refused(body) || 'it was built');

check('body: its text is what a reader reads, sample markup as it is shown', plainText('<p>A run is\n   <code>&lt;w:r&gt;</code> &amp; more&nbsp;&rarr; &#xFB01;&#65;</p>\n<pre><code>x &unknown; y</code></pre>')
    === 'A run is <w:r> & more \u2192 \uFB01A x &unknown; y');
check('body: away from its page, every address in it is whole',
    standaloneBody('<a href="./other-post.html">o</a> <a href="#part">p</a> <img src="figure.png" alt="f"> <a href="https://example.com/">w</a> <code>&lt;img src="x.png"&gt;</code>', { address: 'https://blog.example/this.html' }, 'https://blog.example/')
    === '<a href="https://blog.example/other-post.html">o</a> <a href="https://blog.example/this.html#part">p</a> <img src="https://blog.example/figure.png" alt="f"> <a href="https://example.com/">w</a> <code>&lt;img src="x.png"&gt;</code>');

// ── posts.json ──
const blog = JSON.parse(fs.readFileSync(path.join(SOURCE_DIR, 'posts.json'), 'utf8'));
const categoryIds = (blog.categories || []).map(category => category.id);
check('posts.json names the blog', ['url', 'title', 'author', 'description'].every(field => typeof blog.blog?.[field] === 'string' && blog.blog[field]) && /^https:\/\/.+\/$/.test(blog.blog?.url || ''));
check('posts.json lists topics and posts', Array.isArray(blog.categories) && blog.categories.length > 0 && Array.isArray(blog.posts) && blog.posts.length > 0);
check('every topic has an id and a label, and no id twice', blog.categories.every(c => typeof c.id === 'string' && c.id && typeof c.label === 'string' && c.label)
    && new Set(categoryIds).size === categoryIds.length);

const slugs = blog.posts.map(post => post.slug);
check('no slug twice', new Set(slugs).size === slugs.length, slugs.join(', '));
check('no post is named as the list is', !slugs.includes('index'));
for (const post of blog.posts) {
    const label = `post ${post.slug}`;
    check(`${label}: its slug is lower-case words joined by hyphens`, typeof post.slug === 'string' && SLUG.test(post.slug));
    for (const field of ['title', 'shortTitle', 'description', 'deck']) {
        check(`${label}: has a ${field}`, typeof post[field] === 'string' && post[field].trim().length > 0);
    }
    check(`${label}: its topic is one of the listed`, categoryIds.includes(post.category), String(post.category));
    // A day of the calendar: the 30th of February is the 2nd of March to `Date`, and no date here.
    const isDay = value => /^\d{4}-\d{2}-\d{2}$/.test(value || '') && new Date(`${value}T00:00:00Z`).toISOString().slice(0, 10) === value;
    check(`${label}: the day it was published is YYYY-MM-DD`, isDay(post.published), String(post.published));
    check(`${label}: the day it was updated, when it has one, is YYYY-MM-DD and not before it was published`,
        post.updated === undefined || (isDay(post.updated) && post.updated >= post.published), String(post.updated));
    check(`${label}: has a body file`, fs.existsSync(path.join(SOURCE_DIR, 'posts', `${post.slug}.html`)));
}
check('every topic has a post', categoryIds.every(id => blog.posts.some(post => post.category === id)));

// ── A post's body is content only ──
const bodyFiles = fs.readdirSync(path.join(SOURCE_DIR, 'posts')).filter(name => name.endsWith('.html'));
check('every body file is a listed post', bodyFiles.every(name => slugs.includes(name.replace(/\.html$/, ''))),
    bodyFiles.filter(name => !slugs.includes(name.replace(/\.html$/, ''))).join(', '));
for (const name of bodyFiles) {
    const body = fs.readFileSync(path.join(SOURCE_DIR, 'posts', name), 'utf8');
    const label = `blog/posts/${name}`;
    const shell = SHELL_ELEMENTS.filter(tag => new RegExp(`<${tag}[\\s>/]`, 'i').test(body));
    check(`${label}: holds no page shell, style or script`, shell.length === 0, shell.map(tag => `<${tag}>`).join(', '));
    check(`${label}: carries no style of its own`, !/\sstyle\s*=/i.test(body));
    check(`${label}: carries no class of the page's layout`, !/class="[^"]*\b(?:article-wrapper|article-meta-header|byline-row|article-nav-footer|blog-header|blog-footer)\b/.test(body));
    check(`${label}: carries no event handler`, !/\son[a-z]+\s*=/i.test(body));
    check(`${label}: has no top-level heading (the page writes the title)`, !/<h1[\s>]/i.test(body));
    // A link to another post names the post (`post:<slug>`), never an address: the build decides addresses.
    const linked = [...body.matchAll(/href="post:([^"]*)"/g)].map(match => match[1]);
    check(`${label}: links to posts there are`, linked.every(slug => slugs.includes(slug)), linked.filter(slug => !slugs.includes(slug)).join(', '));
    check(`${label}: links to no page of a post by its address`, !/href="(?:\.\/)?[^":/#?]+\.html"/.test(body) && !/href="[^"]*\?post=/.test(body));
}

// ── What is published is what the content builds ──
let files = new Map();
let rootFiles = new Map();
let buildError = '';
try { files = renderSite(); rootFiles = renderSiteRoot(); } catch (error) { buildError = error.message; }
check('the content builds', buildError === '', buildError);
// What is published is compared with what the content builds, so only for content that builds: the
// one failure above says why it does not.
if (!buildError) {
    const published = fs.existsSync(OUTPUT_DIR) ? fs.readdirSync(OUTPUT_DIR).sort() : [];
    check('docs/blog holds the built files and no others', JSON.stringify(published) === JSON.stringify([...files.keys()].sort()),
        `published: ${published.join(', ')}`);
    for (const [name, content] of files) {
        const file = path.join(OUTPUT_DIR, name);
        const current = fs.existsSync(file) && Buffer.compare(fs.readFileSync(file), Buffer.from(content)) === 0;
        check(`docs/blog/${name} is the one the content builds`, current, 'run `npm run build:blog`');
    }
    for (const post of blog.posts) {
        const page = String(files.get(`${post.slug}.html`) ?? '');
        const label = `page ${post.slug}.html`;
        check(`${label}: has the post's own title, description and address`, page.includes(`<link rel="canonical" href="${blog.blog.url}${post.slug}.html">`)
            && page.includes('<meta property="og:type" content="article">') && /<title>[^<]+ \| officeParser Blog<\/title>/.test(page));
        check(`${label}: names no post by anything but its page`, !page.includes('href="post:'));
    }
    // ── What the build works out of the posts for those who do not read the pages ──
    /** A file of the build as XML, or the reason it is not XML. */
    const xmlOf = content => {
        try {
            return new DOMParser({ onError: (level, message) => { throw new Error(`${level}: ${message}`); } }).parseFromString(String(content ?? ''), 'text/xml');
        } catch (error) { return String(error.message).slice(0, 200); }
    };
    const textsOf = (parent, tag) => [...parent.getElementsByTagName(tag)].map(element => element.textContent);
    const addresses = blog.posts.map(post => `${blog.blog.url}${post.slug}.html`);

    const feed = xmlOf(files.get('feed.xml'));
    check('the feed is XML', typeof feed !== 'string', String(feed));
    if (typeof feed !== 'string') {
        const items = [...feed.getElementsByTagName('item')];
        check('the feed has every post, in order, by its address', JSON.stringify(items.map(item => textsOf(item, 'link')[0])) === JSON.stringify(addresses));
        check('the feed names itself and the blog', feed.getElementsByTagName('atom:link')[0]?.getAttribute('href') === `${blog.blog.url}feed.xml`
            && textsOf(feed.getElementsByTagName('channel')[0], 'link')[0] === blog.blog.url);
        blog.posts.forEach((post, index) => {
            const item = items[index];
            const body = item ? textsOf(item, 'content:encoded')[0] : '';
            check(`feed item ${post.slug}: has the post's title, summary and day`, !!item && textsOf(item, 'title')[0] === post.title
                && textsOf(item, 'description')[0] === post.description && textsOf(item, 'pubDate')[0] === new Date(`${post.published}T00:00:00Z`).toUTCString());
            // Away from its page, an address in a post has to be whole.
            const partial = [...body.matchAll(/<(?:a|img)\b[^>]*\s(?:href|src)="([^"]*)"/g)].map(match => match[1]).filter(target => !/^(?:https?:|mailto:|data:)/.test(target));
            check(`feed item ${post.slug}: holds the post's body, every address in it whole`, body.includes('<p>') && partial.length === 0, partial.join(', '));
        });
    }

    // The sitemap and robots.txt are the site's, and stand at its root: a sitemap may only name pages
    // in its own folder or below, so one in the blog's folder could not name the home page.
    const site = new URL('../', blog.blog.url).href;
    for (const [name, content] of rootFiles) {
        const file = path.join(SITE_DIR, name);
        check(`docs/${name} is the one the content builds`, fs.existsSync(file) && fs.readFileSync(file, 'utf8') === content, 'run `npm run build:blog`');
    }
    check('no sitemap is left in the blog\'s folder', !fs.existsSync(path.join(OUTPUT_DIR, 'sitemap.xml')));
    const sitemap = xmlOf(rootFiles.get('sitemap.xml'));
    check('the sitemap is XML', typeof sitemap !== 'string', String(sitemap));
    if (typeof sitemap !== 'string') {
        const listed = textsOf(sitemap, 'loc');
        check('the sitemap lists the home page, the blog and every post', JSON.stringify(listed) === JSON.stringify([site, blog.blog.url, ...addresses]));
        // The sitemap's own rule: only pages of its site, in its folder or below, each once.
        check('the sitemap names only pages of its site, each once', listed.every(address => address.startsWith(site)) && new Set(listed).size === listed.length,
            listed.filter(address => !address.startsWith(site)).join(', '));
        // A page is a whole document a reader can land on. The parts the home page loads (docs/specs/)
        // and the sample documents (docs/test/) are files of the site and not pages of it.
        const notPages = listed.filter(address => /\/(?:specs|test|dist)\//.test(address.slice(site.length - 1)));
        check('the sitemap names no part or sample file as a page', notPages.length === 0, notPages.join(', '));
        const days = Object.fromEntries([...sitemap.getElementsByTagName('url')].map(url => [textsOf(url, 'loc')[0], textsOf(url, 'lastmod')[0] ?? null]));
        check('the sitemap gives a post and the blog the day they last changed, and the home page none',
            days[site] === null && blog.posts.every(post => days[`${blog.blog.url}${post.slug}.html`] === (post.updated ?? post.published))
            && days[blog.blog.url] === blog.posts.map(post => post.updated ?? post.published).sort().pop());
    }
    check('robots.txt lets everything be read and names the sitemap', String(rootFiles.get('robots.txt')) === `User-agent: *\nAllow: /\n\nSitemap: ${site}sitemap.xml\n`);

    let searched = [];
    try { searched = JSON.parse(String(files.get('search.json') ?? '')).posts; } catch { /* reported below */ }
    check('the search text has every post', JSON.stringify(searched.map(post => post.slug)) === JSON.stringify(slugs));
    // Sample markup is searched for as it is read (`<w:p>`), so a reference is the character it stands for.
    check('the search text is what a reader reads: no references, one line', searched.every(post => post.text.length > 200 && !/&(?:lt|gt|amp|quot);|\n/.test(post.text)),
        searched.filter(post => /&(?:lt|gt|amp|quot);|\n/.test(post.text)).map(post => post.slug).join(', '));

    for (const post of blog.posts) {
        const page = String(files.get(`${post.slug}.html`) ?? '');
        check(`page ${post.slug}.html: says the day it was published, and links the feed`, new RegExp(`Published <time datetime="${post.published}">\\d{1,2} [A-Z][a-z]+ \\d{4}</time>`).test(page)
            && page.includes(`<meta property="article:published_time" content="${post.published}">`) && page.includes('<link rel="alternate" type="application/rss+xml"'));
    }

    const listPage = String(files.get('index.html') ?? '');
    check('the list has the search field, hidden until the script shows it, and a place for each card\'s match', /<div class="post-search" id="post-search" role="search" hidden>/.test(listPage)
        && (listPage.match(/<p class="card-match" hidden><\/p>/g) || []).length === blog.posts.length
        && JSON.stringify([...listPage.matchAll(/data-slug="([^"]+)"/g)].map(match => match[1])) === JSON.stringify(slugs));
    check('the list links every post, in order', JSON.stringify([...listPage.matchAll(/class="article-card"[^>]*href="\.\/([^"]+)\.html"/g)].map(match => match[1])) === JSON.stringify(slugs));
}

// ── A post's days, on a copy of the content with them changed ──
const copy = fs.mkdtempSync(path.join(os.tmpdir(), 'officeparser-blog-'));
try {
    fs.cpSync(SOURCE_DIR, copy, { recursive: true });
    /** The copy built with its second post's days set, or the reason it does not build. */
    const builtWith = days => {
        const changed = JSON.parse(fs.readFileSync(path.join(SOURCE_DIR, 'posts.json'), 'utf8'));
        Object.assign(changed.posts[1], days);
        fs.writeFileSync(path.join(copy, 'posts.json'), JSON.stringify(changed));
        try { return new Map([...renderSite(copy), ...renderSiteRoot(copy)]); } catch (error) { return error.message; }
    };
    const second = blog.posts[1];
    const later = builtWith({ published: '2026-09-03', updated: '2027-01-20' });
    const page = typeof later === 'string' ? later : String(later.get(`${second.slug}.html`));
    check('days: a post says the day it was published and the day it was updated', page.includes('Published <time datetime="2026-09-03">3 September 2026</time>, updated <time datetime="2027-01-20">20 January 2027</time> ·')
        && page.includes('<meta property="article:modified_time" content="2027-01-20">'), page.slice(0, 200));
    check('days: its card has the day it was published, short', typeof later !== 'string' && String(later.get('index.html')).includes('<time datetime="2026-09-03">3 Sep 2026</time>'));
    check('days: a post never updated says no such day', !String(files.get(`${second.slug}.html`) ?? '').includes('updated <time') && !String(files.get(`${second.slug}.html`) ?? '').includes('article:modified_time'));
    check('days: the sitemap and the feed take the day a post last changed', typeof later !== 'string'
        && String(later.get('sitemap.xml')).includes(`<loc>${blog.blog.url}${second.slug}.html</loc>\n    <lastmod>2027-01-20</lastmod>`)
        && String(later.get('sitemap.xml')).includes(`<loc>${blog.blog.url}</loc>\n    <lastmod>2027-01-20</lastmod>`)
        && String(later.get('feed.xml')).includes('<lastBuildDate>Wed, 20 Jan 2027 00:00:00 GMT</lastBuildDate>')
        && String(later.get('feed.xml')).includes('<pubDate>Thu, 03 Sep 2026 00:00:00 GMT</pubDate>'));
    for (const [what, days, said] of [
        ['a month for a day', { published: '2026-10' }, /"2026-10" is not a date/],
        ['a day the calendar does not have', { published: '2026-02-30' }, /"2026-02-30" is not a date/],
        ['an update before the post', { updated: '2026-01-01' }, /was updated \(2026-01-01\) before it was published/],
    ]) check(`days: ${what} is refused, and said`, typeof builtWith(days) === 'string' && said.test(builtWith(days)), String(builtWith(days)).slice(0, 120));
} finally {
    fs.rmSync(copy, { recursive: true, force: true });
}

console.log(failed === 0 ? `✓ Blog: ${passed} passed, 0 failed` : `✗ Blog: ${passed} passed, ${failed} failed`);
process.exit(failed === 0 ? 0 : 1);
