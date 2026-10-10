# The officeParser blog

The blog is written as content here and built into pages in `docs/blog/`, which is what is published.
Nothing in `docs/blog/` is edited by hand.

| Path | What it holds |
|---|---|
| `posts.json` | The blog's title, author and address, the topics, and every post's title, topic, date and summary, in the order the list shows them |
| `posts/<slug>.html` | A post's body: content only |
| `templates/` | The page every post is set in (`layout.html`), the list page (`list.html`) and a post's page (`post.html`) |
| `assets/` | The stylesheet, the script, and any picture a post shows |

## Adding a post

1. Add an entry to `posts` in `posts.json`, where it should stand in the list:

   ```json
   {
       "slug": "how-xlsx-stores-a-date",
       "title": "How a Spreadsheet Stores a Date",
       "shortTitle": "How XLSX Stores a Date",
       "category": "deep-dive",
       "published": "2026-11-18",
       "description": "One or two sentences for the list, search results and link previews.",
       "deck": "The standfirst under the headline on the post's own page."
   }
   ```

   The slug is the post's address (`/blog/<slug>.html`) and its file name: lower-case words joined by
   hyphens. `category` is the `id` of a topic in `categories`. A new topic is a new entry there.
   `published` is the day the post goes out. The list, the feed and the previous and next links follow
   the order of `posts`, so a new post goes first. A post changed after it went out also gets
   `"updated": "2026-12-02"`: the page says so, and the feed and the sitemap take that day.

2. Write the body in `posts/<slug>.html`.

3. Build and check:

   ```bash
   npm run build:blog
   ```

   ```bash
   npm run test:blog
   ```

4. Commit `blog/` and `docs/blog/` together. The pre-commit hook builds and stages `docs/blog/` too.

Reading time, the date as readers see it, the previous and next links, the page's title, description
and share tags are all worked out by the build. So are these files, which nobody writes:

| File | What it is for |
|---|---|
| `feed.xml` | The RSS feed: every post with its whole body, for feed readers. Every page links it |
| `../sitemap.xml` | The site's sitemap, at the site's root (`docs/sitemap.xml`): the home page, the list and every post, for search engines. `docs/robots.txt`, which the build writes too, names it |
| `search.json` | Each post's text, which the list page's search looks through |

The sitemap stands at the site's root and not in `docs/blog/` because a sitemap may only name pages in
its own folder or below, so one in the blog's folder could not name the home page. It names pages of
this site only: the README and the changelog are on GitHub, and the files in `docs/specs/` are parts
the home page loads, not pages.

The list page's search and topic filter need nothing from a post: a new post is found by its title,
summary, topic and text as soon as it is built.

## What a body is written with

A body starts at its first paragraph. The page writes the headline, the deck and the byline, so a body
has no `<h1>`, and no page shell, styles, scripts or inline `style`.

- **Text:** `<p>`, `<h2>`, `<h3>`, `<ul>`, `<ol>`, `<li>`, `<blockquote>`, `<hr>`, `<strong>`, `<em>`,
  `<code>`, `<a>`.
- **Code:** `<pre><code>…</code></pre>`. Sample markup is written escaped: `&lt;w:p&gt;`, not `<w:p>`.
  The build refuses a tag it does not know, because a browser would show nothing of it.
- **A callout:**

  ```html
  <div class="callout">
      <div class="callout-title">The short version</div>
      <p>What the reader should take away.</p>
  </div>
  ```

- **A table:** plain, with its header row in `<thead>`. On a phone each row becomes a card with every
  cell under its column's name, so a row's first cell should be what the row is about. The build adds
  the labels and the accessibility roles. Merged cells and tables inside tables are refused.

  ```html
  <table>
      <thead>
          <tr><th>Format</th><th>Stored as</th></tr>
      </thead>
      <tbody>
          <tr><td>XLSX</td><td>A day count</td></tr>
      </tbody>
  </table>
  ```

- **A link to another post:** by its slug, `<a href="post:how-pptx-stores-a-slide">`. The build writes
  the address and refuses a slug there is no post for.
- **A picture:** put the file in `assets/` and name it, `<img src="date-epoch.png" alt="…">`, alone or
  in a `<figure>` with a `<figcaption>`. The build refuses a picture that is not there.

A class a post uses has to be in `assets/blog.css`. The look of every post is changed there and in
`templates/`, never in a post.

## Writing

The blog is written by Harsh Ankur in the first person. Every claim is checked against the code, and
every sample that is called real output is run. A limit of the library is said plainly.
