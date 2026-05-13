---
created: 2025-11-25T16:36:03.000Z
modified: 2026-01-01T16:08:06.000Z
TestString: "Hello from custom props"
TestNumber: 42
TestBool: true
---


---

# Demonstration of DOCX support in calibre {#demonstration-of-docx-support-in-calibre}

This document demonstrates the ability of the calibre DOCX Input plugin to convert the various typographic features in a Microsoft Word (2007 and newer) document. Convert this document to a modern ebook format, such as AZW3 for Kindles or EPUB for other ebook readers, to see it in action.

There is support for images, tables, lists, footnotes, endnotes, links, dropcaps and various types of text and paragraph level formatting.

To see the DOCX conversion in action, simply add this file to calibre using the **“Add Books” **button and then click “**Convert”. ** Set the output format in the top right corner of the conversion dialog to EPUB or AZW3 and click **“OK”**.



Slide Note 1.

<div style="text-align: right">1</div>


---

# **Text Formatting** {#text-formatting}

**Inline formatting**

Here, we demonstrate various types of inline text formatting and the use of embedded fonts.

Here is some **bold, ***italic, ****bold-italic, ***<u>underlined </u>and ~~struck out ~~ text. Then, we have a superscript and a subscript. Now we see some red, green and blue text. Some text with a yellow highlight. Some text in a box. Some text in inverse video.

A paragraph with styled text: *subtle emphasis  *followed by **strong text **and ***intense emphasis***. This paragraph uses document wide styles for styling rather than inline text properties as demonstrated in the previous paragraph — calibre can handle both with equal ease.

**Fun with fonts**

This document has embedded the Ubuntu font family. The body text is in the Ubuntu typeface, here is some text in the Ubuntu Mono typeface, notice how every letter has the same width, even i and m. Every embedded font will automatically be embedded in the output ebook during conversion. 

**Paragraph level formatting**

<div style="text-align: right">You can do crazy things with paragraphs, if the urge strikes you. For instance this paragraph is right aligned and has a right border. It has also been given a light gray background.</div>

For the lovers of poetry amongst you, paragraphs with hanging indents, like this often come in handy. You can use hanging indents to ensure that a line of poetry retains its individual identity as a line even when the screen is  too narrow to display it as a single line. Not only does this paragraph have a hanging indent, it is also has an extra top margin, setting it apart from the preceding paragraph.



Slide Note 2. This is a **bold** text, but no *italic*. Let's strike ~~that~~ out.

<div style="text-align: right">2</div>


---

# **Tables** {#tables}


| **ITEM ** | **NEEDED ** |
|  ---  |  ---  |
| Books | 1 |
| Pens | 3 |
| Pencils | 2 |
| Highlighter | 2 colors |
| Scissors | 1 pair |

 Tables in Word can vary from the extremely simple to the extremely complex. calibre tries to do its best when converting tables. While you may run into trouble with the occasional table, the vast majority of common cases should be converted very well, as demonstrated in this section. Note that for optimum results, when creating tables in Word, you should set their widths using percentages, rather than absolute units.  To the left of this paragraph is a floating two column table with a nice green border and header row. 

 Now let’s look at a fancier table—one with alternating row colors and partial borders. This table is stretched out to take 100% of the available width.


| City or Town | <div style="text-align: center">Point A </div> | <div style="text-align: center">Point B </div> | <div style="text-align: center">Point C </div> | <div style="text-align: center">Point D </div> | <div style="text-align: center">Point E </div> |
|  ---  |  ---  |  ---  |  ---  |  ---  |  ---  |
| Point A | <div style="text-align: center">— </div> |  |  |  |  |
| Point B | <div style="text-align: center">87 </div> | <div style="text-align: center">— </div> |  |  |  |
| Point C | <div style="text-align: center">64 </div> | <div style="text-align: center">56 </div> | <div style="text-align: center">— </div> |  |  |
| Point D | <div style="text-align: center">37 </div> | <div style="text-align: center">32 </div> | <div style="text-align: center">91 </div> | <div style="text-align: center">— </div> |  |
| Point E | <div style="text-align: center">93 </div> | <div style="text-align: center">35 </div> | <div style="text-align: center">54 </div> | <div style="text-align: center">43 </div> | <div style="text-align: center">— </div> |



And then about tables

<div style="text-align: right">3</div>


---

Next, we see a table with special formatting in various locations. Notice how the formatting for the header row and sub header rows is preserved.


| **College ** | **New students ** | **Graduating students ** | **Change ** |
|  ---  |  ---  |  ---  |  ---  |
|  | *Undergraduate* |  |  |
| Cedar University | 110 | 103 | +7 |
| Oak Institute | 202 | 210 | -8 |
|  | *Graduate* |  |  |
| Cedar University | 24 | 20 | +4 |
| Elm College | 43 | 53 | -10 |
| Total | 998 | 908 | 90 |

*Source:* Fictitious data, for illustration purposes only 

Next, we have something a little more complex, a nested table, i.e. a table inside another table. Additionally, the inner table has some of its cells merged. The table is displayed horizontally centered. 


<table>
  <tr>
    <td rowspan="2"><p>One </p><p>Three </p></td>
    <td><p>Two </p></td>
  </tr>
  <tr>
    <td><p>Four </p></td>
  </tr>
</table>


|  | To the left is a table inside a table, with some cells merged. |
|  ---  |  ---  |




---

 We end with a fancy calendar, note how much of the original formatting is preserved. Note that this table will only display correctly on relatively wide screens. In general, very wide tables or tables whose cells have fixed width requirements don’t fare well in ebooks.


<table>
  <tr>
    <td colspan="13"><p>December 2007 </p></td>
    <td></td>
  </tr>
  <tr>
    <td><p>Sun </p></td>
    <td><p> </p></td>
    <td><p>Mon </p></td>
    <td><p> </p></td>
    <td><p>Tue </p></td>
    <td><p> </p></td>
    <td><p>Wed </p></td>
    <td><p> </p></td>
    <td><p>Thu </p></td>
    <td><p> </p></td>
    <td><p>Fri </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>Sat </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>1 </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
  <tr>
    <td><p>2 </p></td>
    <td><p> </p></td>
    <td><p>3 </p></td>
    <td><p> </p></td>
    <td><p>4 </p></td>
    <td><p> </p></td>
    <td><p>5 </p></td>
    <td><p> </p></td>
    <td><p>6 </p></td>
    <td><p> </p></td>
    <td><p>7 </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>8 </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
  <tr>
    <td><p>9 </p></td>
    <td><p> </p></td>
    <td><p>10 </p></td>
    <td><p> </p></td>
    <td><p>11 </p></td>
    <td><p> </p></td>
    <td><p>12 </p></td>
    <td><p> </p></td>
    <td><p>13 </p></td>
    <td><p> </p></td>
    <td><p>14 </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>15 </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
  <tr>
    <td><p>16 </p></td>
    <td><p> </p></td>
    <td><p>17 </p></td>
    <td><p> </p></td>
    <td><p>18 </p></td>
    <td><p> </p></td>
    <td><p>19 </p></td>
    <td><p> </p></td>
    <td><p>20 </p></td>
    <td><p> </p></td>
    <td><p>21 </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>22 </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
  <tr>
    <td><p>23 </p></td>
    <td><p> </p></td>
    <td><p>24 </p></td>
    <td><p> </p></td>
    <td><p>25 </p></td>
    <td><p> </p></td>
    <td><p>26 </p></td>
    <td><p> </p></td>
    <td><p>27 </p></td>
    <td><p> </p></td>
    <td><p>28 </p></td>
    <td><p> </p></td>
    <td colspan="2"><p>29 </p></td>
  </tr>
  <tr>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
  <tr>
    <td><p>30 </p></td>
    <td><p> </p></td>
    <td><p>31 </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td><p> </p></td>
    <td colspan="2"><p> </p></td>
  </tr>
</table>



Now calendars

1. Sun
2. Mon
7. Sat

<div style="text-align: right">5</div>


---

# **Structural Elements** {#structural-elements}

Miscellaneous structural elements you can add to your document, like footnotes, endnotes, dropcaps and the like. 

**Footnotes & Endnotes**

Footnotes1 and endnotesi are automatically recognized and both are converted to endnotes, with backlinks for maximum ease of use in ebook devices.

**Dropcaps**

D

rop caps are used to emphasize the leading paragraph at the start of a section. In Word it is possible to specify how many lines of text a drop-cap should use. 

**Links**

Two kinds of links are possible, those that refer to an external website and those that refer to locations inside the document itself. Both are supported by calibre. For example, here is a link pointing to the [<u>calibre download page</u>](http://calibre-ebook.com/download). Then we have a link that points back to the section on [<u>paragraph level formatting</u>](#slide2) in this document.




---

# **Images** {#images}

Centered images like this are useful for large pictures that should be a focus of attention. 

<a id="picture-8"></a>
![165 Inspirational Quotes To Keep You Motivated In Life](data:image/png;base64,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)There is no analogous technology in ebooks, so theconversion will usually end up placingthe image either centered or floating close to the point in the text where itwas *inserted*, notnecessarily where it appears on the page in Word.




---

<div style="text-align: center">**Lists**</div>

All types of lists are supported by the conversion, with the exception of lists that use fancy bullets, these get converted to regular bullets.

**Bulleted List**

- One
- Two
**Numbered List**

1. One, with a very long line to demonstrate that the hanging indent for the list is working correctly
2. Two
**Multi-level Lists**

1. One
    1. Two
        1. Three
        2. Four with a very long line to demonstrate that the hanging indent for the list is working correctly.
        3. Five
2. Six
A Multi-level list with bullets:

- One
    - Two
        - This bullet uses an image as the bullet item
            - Four
- Five
**Continued Lists**

1. One
2. Two
An interruption in our regularly scheduled listing, for this essential and very relevant public service announcement.

3. We now resume our normal programming
4. Four



---

# Charts {#charts}