---
title: Exhaustive Markdown Test
author: Test Author
description: Tests every markdown feature
tags: [tag1, tag2]
version: 1
---

# Heading Level 1 {#h1-anchor}

## Heading Level 2

### Heading Level 3

#### Heading Level 4

##### Heading Level 5

###### Heading Level 6

Plain paragraph with **bold text**, *italic text*, ~~strikethrough text~~, <u>underlined text</u>, <sub>subscript</sub>, <sup>superscript</sup>.

Inline `monospace code` in paragraph.

External link: [Visit Example](https://example.com)

Internal anchor link: [Go to H1](#h1-anchor)

Wikilink: [[WikiPage]] and [[WikiPage|Alias Text]]

Citation: [@smith2023]

Footnote reference[^fn1].

[^fn1]: This is the footnote definition text.

*[ABBR]: Abbreviation Full Title

The word ABBR appears here and gets an abbreviation title.

> Regular blockquote paragraph without admonition type.

> [!NOTE]
> This is a note admonition.

> [!TIP]
> This is a tip admonition.

> [!IMPORTANT]
> This is an important admonition.

> [!WARNING]
> This is a warning admonition.

> [!CAUTION]
> This is a caution admonition.

:::danger
This is a GLFM danger admonition mapped to caution.
:::

- Unordered item A
- Unordered item B
  - Nested unordered item
- Unordered item C

1. Ordered item 1
2. Ordered item 2
   1. Nested ordered item
3. Ordered item 3

- [x] Completed task item
- [ ] Incomplete task item

Term Alpha
: Description for term alpha

Term Beta
: Description one for beta
: Description two for beta

```typescript
const x: number = 42;
console.log(x);
```

```python
def hello():
    print("world")
```

Inline math: $E=mc^2$

Block math:

$$
a^2 + b^2 = c^2
$$

| Left Aligned | Center Aligned | Right Aligned |
| :--- | :---: | ---: |
| Row1-L | Row1-C | Row1-R |
| Row2-L | Row2-C | Row2-R |

<table data-align="center">
  <tr><th colspan="2">Merged Header</th><th>Normal</th></tr>
  <tr><td rowspan="2">Rowspan Cell</td><td>Cell B2</td><td>Cell C2</td></tr>
  <tr><td>Cell B3</td><td>Cell C3</td></tr>
</table>

![Logo Image](https://example.com/logo.png){width=50px align=center}

<div data-youtube-video="dQw4w9WgXcQ" data-width="500"></div>

<div align="right">
Right-aligned paragraph content.
</div>

> > Two-level nested blockquote line one.
> > Two-level nested blockquote line two.

> > > Three-level nested blockquote.

1) Paren-marker ordered item one
2) Paren-marker ordered item two

| C1 | C2 |
|-|-|
| v1 | v2 |

~~~javascript
console.log('tilde fence');
~~~

----- not actually a horizontal rule, just a paragraph with leading hyphens

- Sibling parent Alpha
  - Alpha child one
  - Alpha child two
- Sibling parent Beta
  - Beta child one

Escaped literal characters: \*not bold\*, \_not italic\_, \`not code\`, \[not a link\].

See [the reference link][refone] and a [shortcut reference] and an image ![ref image][refimg]. An [unresolved reference][nowhere] and a [bare bracket] stay literal.

[refone]: https://example.com/reference "Reference Title"
[shortcut reference]: https://example.com/shortcut
[refimg]: https://example.com/ref-image.png

Underscore emphasis: this is _underscore italic_ and this is __underscore bold__ text.

Multi-backtick code span: ``code with a ` backtick inside``.

Entity references: Fish & Chips, Q&A, &amp; &lt; &gt; &#39; &#x2764;, and out-of-bounds &#999999999; and &#x999999999;.

Hard break line one  
Hard break line two follows a two-space trailing marker.

Setext Heading One
===================

Setext Heading Two
-------------------

Autolink: visit <https://example.com/autolink> for details.

- Continuation parent item
  continuation text merged into the parent item

    indented code block line one
    indented code block line two

---
