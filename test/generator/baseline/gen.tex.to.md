---
title: "DOCX Demo"
author: "Kovid Goyal"
created: 2013-06-05T07:56:00.000Z
modified: 2025-11-25T20:13:45.000Z
description: "Demonstration of DOCX support in calibre"
keywords: "calibre, docs, ebook, conversion"
TestString: "Hello from custom props"
TestNumber: "42"
TestBool: "true"
---

<a name="toc581531977"></a>
<div style="text-align: center">

# Demonstration of DOCX support in calibre {#demonstration-of-docx-support-in-calibre}

</div>

This document demonstrates the ability of the calibre DOCX Input plugin to convert the various typographic features in a Microsoft Word (2007 and newer) document. Convert this document to a modern ebook format, such as AZW3 for Kindles or EPUB for other ebook readers, to see it in action.

There is support for images, tables, lists, footnotes, endnotes, links, dropcaps and various types of text and paragraph level formatting.

To see the DOCX conversion in action, simply add this file to calibre using the **“Add Books” **button and then click “**Convert”. **Set the output format in the top right corner of the conversion dialog to EPUB or AZW3 and click **“OK”**.

<a name="toc2054249818"></a>
# Text Formatting {#text-formatting}

<a name="toc2137712100"></a>
## Inline formatting {#inline-formatting}

Here, we demonstrate various types of inline text formatting and the use of embedded fonts.

Here is some **bold, ***italic, ****bold-italic, ***<u>underlined </u>and ~~struck out ~~text. Then, we have a super<sup>script</sup> and a sub<sub>script</sub>. Now we see some red, green and blue text. Some text with a ==yellow highlight==. Some text in a box. Some text in ==inverse video==.

A paragraph with styled text: *subtle emphasis *followed by **strong text **and ***intense emphasis***. This paragraph uses document wide styles for styling rather than inline text properties as demonstrated in the previous paragraph — calibre can handle both with equal ease.

<a name="toc1074133965"></a>
## Fun with fonts {#fun-with-fonts}

This document has embedded the Ubuntu font family. The body text is in the Ubuntu typeface, here is `some text in the Ubuntu Mono typeface, notice how every letter has the same width, even i and m`. Every embedded font will automatically be embedded in the output ebook during conversion.

<a name="paragraph-level-formatting"></a><a name="toc2022725662"></a>
## Paragraph level formatting {#paragraph-level-formatting}

<div style="text-align: right">==You can do crazy things with paragraphs, if the urge strikes you. For instance this paragraph is right aligned and has a right border. It has also been given a light gray background.==</div>

For the lovers of poetry amongst you, paragraphs with hanging indents, like this often come in handy. You can use hanging indents to ensure that a line of poetry retains its individual identity as a line even when the screen is too narrow to display it as a single line. Not only does this paragraph have a hanging indent, it is also has an extra top margin, setting it apart from the preceding paragraph.

<a name="toc28114276"></a>
# Tables {#tables}


| ITEM | NEEDED |
| --- | --- |
| Books | 1 |
| Pens | 3 |
| Pencils | 2 |
| Highlighter | 2 colors |
| Scissors | 1 pair |

Tables in Word can vary from the extremely simple to the extremely complex. calibre tries to do its best when converting tables. While you may run into trouble with the occasional table, the vast majority of common cases should be converted very well, as demonstrated in this section. Note that for optimum results, when creating tables in Word, you should set their widths using percentages, rather than absolute units. To the left of this paragraph is a floating two column table with a nice green border and header row.

Now let’s look at a fancier table—one with alternating row colors and partial borders. This table is stretched out to take 100% of the available width.


| City or Town | <div style="text-align: center">Point A</div> | <div style="text-align: center">Point B</div> | <div style="text-align: center">Point C</div> | <div style="text-align: center">Point D</div> | <div style="text-align: center">Point E</div> |
| --- | --- | --- | --- | --- | --- |
| Point A | <div style="text-align: center">—</div> |  |  |  |  |
| Point B | <div style="text-align: center">87</div> | <div style="text-align: center">—</div> |  |  |  |
| Point C | <div style="text-align: center">64</div> | <div style="text-align: center">56</div> | <div style="text-align: center">—</div> |  |  |
| Point D | <div style="text-align: center">37</div> | <div style="text-align: center">32</div> | <div style="text-align: center">91</div> | <div style="text-align: center">—</div> |  |
| Point E | <div style="text-align: center">93</div> | <div style="text-align: center">35</div> | <div style="text-align: center">54</div> | <div style="text-align: center">43</div> | <div style="text-align: center">—</div> |

Next, we see a table with special formatting in various locations. Notice how the formatting for the header row and sub header rows is preserved.


| College | New students | Graduating students | Change |
| --- | --- | --- | --- |
|  | *Undergraduate* |  |  |
| Cedar University | 110 | 103 | +7 |
| Oak Institute | 202 | 210 | -8 |
|  | *Graduate* |  |  |
| Cedar University | 24 | 20 | +4 |
| Elm College | 43 | 53 | -10 |
| Total | 998 | 908 | 90 |

*Source:*Fictitious data, for illustration purposes only

Next, we have something a little more complex, a nested table, i.e. a table inside another table. Additionally, the inner table has some of its cells merged. The table is displayed horizontally centered.


<table>
  <tr>
    <td><table>
  <tr>
    <td rowspan="2"><p>One</p><p>Three</p></td>
    <td><p>Two</p></td>
  </tr>
  <tr>
    <td><p>Four</p></td>
  </tr>
</table>
</td>
    <td><p>To the left is a table inside a table, with some cells merged.</p></td>
  </tr>
</table>

We end with a fancy calendar, note how much of the original formatting is preserved. Note that this table will only display correctly on relatively wide screens. In general, very wide tables or tables whose cells have fixed width requirements don’t fare well in ebooks.


<table>
  <tr>
    <td colspan="13"><p>December 2007</p></td>
    <td></td>
  </tr>
  <tr>
    <td><p>Sun</p></td>
    <td></td>
    <td><p>Mon</p></td>
    <td></td>
    <td><p>Tue</p></td>
    <td></td>
    <td><p>Wed</p></td>
    <td></td>
    <td><p>Thu</p></td>
    <td></td>
    <td><p>Fri</p></td>
    <td></td>
    <td colspan="2"><p>Sat</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"><p>1</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td><p>2</p></td>
    <td></td>
    <td><p>3</p></td>
    <td></td>
    <td><p>4</p></td>
    <td></td>
    <td><p>5</p></td>
    <td></td>
    <td><p>6</p></td>
    <td></td>
    <td><p>7</p></td>
    <td></td>
    <td colspan="2"><p>8</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td><p>9</p></td>
    <td></td>
    <td><p>10</p></td>
    <td></td>
    <td><p>11</p></td>
    <td></td>
    <td><p>12</p></td>
    <td></td>
    <td><p>13</p></td>
    <td></td>
    <td><p>14</p></td>
    <td></td>
    <td colspan="2"><p>15</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td><p>16</p></td>
    <td></td>
    <td><p>17</p></td>
    <td></td>
    <td><p>18</p></td>
    <td></td>
    <td><p>19</p></td>
    <td></td>
    <td><p>20</p></td>
    <td></td>
    <td><p>21</p></td>
    <td></td>
    <td colspan="2"><p>22</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td><p>23</p></td>
    <td></td>
    <td><p>24</p></td>
    <td></td>
    <td><p>25</p></td>
    <td></td>
    <td><p>26</p></td>
    <td></td>
    <td><p>27</p></td>
    <td></td>
    <td><p>28</p></td>
    <td></td>
    <td colspan="2"><p>29</p></td>
  </tr>
  <tr>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td><p>30</p></td>
    <td></td>
    <td><p>31</p></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
</table>

<a name="toc54889875"></a>
# Structural Elements {#structural-elements}

Miscellaneous structural elements you can add to your document, like footnotes, endnotes, dropcaps and the like.

<a name="toc201580556"></a>
## Footnotes & Endnotes {#footnotes-endnotes}

Footnotes[^1] and endnotes[^2] are automatically recognized and both are converted to endnotes, with backlinks for maximum ease of use in ebook devices.

<a name="toc1977424358"></a>
## Dropcaps {#dropcaps}

D

rop caps are used to emphasize the leading paragraph at the start of a section. In Word it is possible to specify how many lines of text a drop-cap should use.

<a name="toc1233048813"></a>
## Links {#links}

Two kinds of links are possible, those that refer to an external website and those that refer to locations inside the document itself. Both are supported by calibre. For example, here is a link pointing to the [<u>calibre download page</u>](http://calibre-ebook.com/download). Then we have a link that points back to the section on [<u>paragraph level formatting</u>](#paragraph-level-formatting) in this document.

<a name="toc64145348"></a>
## Table of Contents {#table-of-contents}

You can see the Table of Contents created by calibre by clicking the Table of Contents button in whatever viewer you are using to view the converted ebook.

[<u>**Demonstration of DOCX support in calibre1**</u>](#toc581531977)

[<u>**Text Formatting1**</u>](#toc2054249818)

[<u>*Inline formatting2*</u>](#toc2137712100)

[<u>*Fun with fonts2*</u>](#toc1074133965)

[<u>*Paragraph level formatting2*</u>](#toc2022725662)

[<u>**Tables2**</u>](#toc28114276)

[<u>**Structural Elements4**</u>](#toc54889875)

[<u>*Footnotes & Endnotes5*</u>](#toc201580556)

[<u>*Dropcaps5*</u>](#toc1977424358)

[<u>*Links5*</u>](#toc1233048813)

[<u>*Table of Contents5*</u>](#toc64145348)

[<u>**Images6**</u>](#toc484565143)

[<u>**Lists7**</u>](#toc1359965655)

[<u>*Bulleted List8*</u>](#toc1958162433)

[<u>*Numbered List8*</u>](#toc415190676)

[<u>*Multi-level Lists8*</u>](#toc1093260318)

[<u>*Continued Lists8*</u>](#toc1471533984)

<a name="toc484565143"></a>
# Images {#images}

Centered images like this are useful for large pictures that should be a focus of attention.

![image](images/image.jpg){width=810pt}

There is no analogous technology in ebooks, so the conversion will usually end up placing the image either centered or floating close to the point in the text where it was *inserted*, not necessarily where it appears on the page in Word.

<a name="toc1359965655"></a>
# Lists {#lists}

All types of lists are supported by the conversion, with the exception of lists that use fancy bullets, these get converted to regular bullets.

<a name="toc1958162433"></a>
## Bulleted List {#bulleted-list}

- One
- Two

<a name="toc415190676"></a>
## Numbered List {#numbered-list}

1. One, with a very long line to demonstrate that the hanging indent for the list is working correctly
2. Two

<a name="toc1093260318"></a>
## Multi-level Lists {#multi-level-lists}

1. One
    1. Two
        1. Three
        2. Four with a very long line to demonstrate that the hanging indent for the list is working correctly.
        3. Five
2. Six

A Multi-level list with bullets:

- One
    - Two
        - This bullet uses an image as the bullet item
            - Four
- Five

<a name="toc1471533984"></a>
## Continued Lists {#continued-lists}

1. One
2. Two

An interruption in our regularly scheduled listing, for this essential and very relevant public service announcement.

3. We now resume our normal programming
4. Four

[^1]: In paged media, footnotes are usually displayed at the bottom of the text. However, in ebooks, a better paradigm is to make them clickable endnotes that the user can browse at her pleasure. This conversion is handled automatically by calibre.

[^2]: Endnotes are typically used for longer notes, they remain endnotes when converted into ebook form, except that they have an additional backlink to make it easy to return to the current position after reading the note.