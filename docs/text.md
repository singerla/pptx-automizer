---
title: Modify Text
description: Set, replace and style text — plain, tagged, MultiText paragraphs, or converted from HTML.
---

You can select and import generic shapes from any loaded template. It is possible to update the containing text in several ways:

```ts
import { ModifyTextHelper, XmlElement } from 'pptx-automizer';

pres.addSlide('SlideWithImages.pptx', 1, (slide) => {
  // You can directly modify the child nodes of <p:sp>
  slide.addElement('shapes', 2, 'Arrow', (element: XmlElement) => {
    element.getElementsByTagName('a:t').item(0).textContent = 'Custom content';
  });

  // You might prefer a built-in function to set text:
  slide.addElement('shapes', 2, 'Arrow', [
    ModifyTextHelper.setText('This is my text'),
  ]);
});
```

## Replace tagged text

`pptx-automizer` also provides a powerful helper to replace tagged text. You can use e.g. `{{myTag}}` on your slide and apply a modifier to insert dynamic text. Font style can be inherited from template or updated by the modifier.

```ts
import { modify } from 'pptx-automizer';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement(
    // This is the name of the target element on slide #1 of
    // 'TextReplace.pptx
    'replaceText',
    // This will look for a string `{{replace}}` inside the text
    // contents of 'replaceText' shape
    modify.replaceText([
      {
        replace: 'replace',
        by: {
          text: 'Apples',
        },
      },
    ]),
  );
});
```

## MultiText: styled paragraphs and lists

You can use `modify.setMultiText` to replace all text contents of an existing textfield by styled paragraphs, bulleted lists and text runs:

```ts
import { modify } from 'pptx-automizer';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement(
    'setText',
    modify.setMultiText([
      {
        paragraph: {
          bullet: true,
          level: 0,
          marginLeft: 41338,
          indent: -87325,
          alignment: 'l',
        },
        textRuns: [
          {
            text: 'Bullet point level 0',
            style: {
              isItalics: true,
              color: {
                type: 'srgbClr',
                value: 'CCCCCC',
              },
            },
          },
        ],
      },
    ]),
  );
});
```

Within a text run, `\n` and `\v` (U+000B, the character PowerPoint itself uses for
a soft line break created with Shift+Enter) are converted into an `<a:br/>` line
break inside the same paragraph. Add another entry to the array if you need a real
paragraph with its own bullet, level and alignment.

`paragraph.lineSpacing`, `spaceBefore` and `spaceAfter` take points as a plain
number, or `{ percent: 100 }` for spacing relative to the line height (100 =
one line) — the relative form scales with the paragraph's font size.

## Convert HTML to text contents

It is also possible to directly convert an HTML page into pptx text contents. HTML code will be flattened and converted into a MultiText array.

```ts
import { modify } from 'pptx-automizer';

const html =
  '<html><body>' +
  '<h2 style="text-align: center">Quarterly report</h2>' +
  '<p>Plain text with <strong>bold</strong>, <em>italics</em> and ' +
  '<span style="color: #ff0000; font-size: 12pt">styling</span>.</p>' +
  '<ul>' +
  '<li>bullet level 0' +
  '<ul><li>bullet level 1</li></ul>' +
  '</li>' +
  '</ul>' +
  '<ol><li>numbered</li><li>list</li></ol>' +
  '<p><a href="https://example.com">external link</a> and ' +
  '<a href="3">a link to slide 3</a></p>' +
  '</body></html>';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modifyElement('setText', modify.htmlToMultiText(html));
});
```

### What HTML is supported

PPTX text is strictly flat: a text body is a list of paragraphs, each a list of
text runs, with no nesting anywhere. HTML hierarchy is therefore *projected*
onto that — nested inline tags become one run with accumulated character
properties, nested lists become paragraphs with a 0-based level, and a block
inside a block (`<li><p>…</p></li>`) yields a single paragraph, with the
innermost block winning.

| | Supported |
|---|---|
| Paragraphs | `<p>`, `<div>`, `<h1>`–`<h6>`, `<blockquote>`, `<pre>`, `<section>` & friends |
| Lists | `<ul>`, `<ol>`, `<li>`, nested to 9 levels. `<ol>` renders as real automatic numbering (1. / a. / i. per level) |
| Inline | `<strong>`/`<b>`, `<em>`/`<i>`, `<u>`/`<ins>`, `<s>`/`<strike>`/`<del>`, `<sub>`, `<sup>`, `<code>`/`<kbd>`/`<samp>` (monospace), `<mark>`, `<br>`, `<span>`, `<a>`, `<font>` |
| Links | `<a href="https://…">` external, `<a href="3">` to slide 3 |
| CSS (on any element) | `font-size` (`px` converted at 96dpi, or `pt`), `color`, `background-color` (highlight), `font-weight`, `font-style`, `text-decoration`, `font-family`, `text-align` |

Both list-nesting styles work and produce identical output — properly nested
(`<li>text<ul>…</ul></li>`) and the sibling form CKEditor emits
(`<ul><li/><ul>…</ul></ul>`).

Good to know:

- The input has to be wrapped in `<html><body>…</body></html>`, and is parsed as
  XML: quote your attributes and close your container tags, as WYSIWYG editors
  do. `&nbsp;`, `&amp;` and an unclosed `<br>` are fine.
- Colors can be written in any CSS notation (`#f00`, `red`, `rgb(255,0,0)`) and
  are normalized to the 6-digit hex OOXML requires.
- Relative font sizes (`em`, `%`) are ignored rather than guessed, leaving the
  size inherited from the template.
- Whitespace collapses the way a browser collapses it; `&nbsp;` survives.
- Vertical spacing mirrors the browser's default stylesheet: `<p>`, headings
  and the outer edges of lists get one collapsed gap of one line height
  (`spaceBefore: { percent: 100 }`, so it scales with each paragraph's font
  size), while items of the same list sit tight. A trailing `<br>` before a
  closing block tag is dropped, exactly as a browser renders it; a deliberate
  `<br><br>` keeps its one empty line.
- Alignment is only written when the HTML asks for it — otherwise the target
  shape's layout keeps deciding.
- Font size and color of the target shape's existing text are used as the
  fallback style, so generated text keeps the template's look.
- `<table>` markup has no equivalent in a single text shape: the cell text is
  kept, but flattened into one paragraph. Use `modify.setTableData` for
  [tables](./tables.md).
- `<a href="4">` is a slide **number in the finished output deck**, counting the
  root template's existing slides — not the index of your `addSlide()` calls.
  The target slide has to exist, or the relationship dangles and PowerPoint
  shows the text underlined but unlinked, without a warning.
- A `color` on an `<a>` element is written to the run, but PowerPoint paints
  hyperlink text in the theme's `<a:hlink>` color regardless. To restyle links,
  change that theme color; per-link colors are not achievable in PPTX.

## Text helpers (MultiText/HTML)

Generate complex text (multiple runs, links, bullets) either from a structured value or directly from HTML.

```ts
// From structured paragraphs
slide.modifyElement('TextBox', [
  ModifyTextHelper.setMultiText([
    {
      paragraph: { bullet: false },
      textRuns: [
        { text: 'Hello ', style: { isBold: true } },
        { text: 'World' },
      ],
    },
  ]),
]);

// From HTML - note the required <html><body> wrapper
const html =
  '<html><body><p><b>Bold</b> and ' +
  '<a href="https://example.com">link</a></p></body></html>';
slide.modifyElement('TextBox', [ModifyTextHelper.htmlToMultiText(html)]);
```

`HtmlToMultiTextHelper` and `MultiTextHelper` also support hyperlinks: an
external target (`<a href="https://...">`) or a slide number for an internal
link (`<a href="3">`). See [What HTML is supported](#what-html-is-supported)
above for the full tag and CSS coverage, and the tests:

- [Replace text by MultiText objects](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text.test.ts)
- [Replace text by HTML](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-html.test.ts)
- [HTML conversion rules, unit level](https://github.com/singerla/pptx-automizer/blob/main/__tests__/html-to-multitext-converter.test.ts)

## Find out more

- [Replace and style by tags](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-tagged-text.test.ts)
- [Modify text elements using getAllTextElementIds](https://github.com/singerla/pptx-automizer/blob/main/__tests__/get-all-text-element-ids.test.ts)
- [Replace text by multitext objects](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text.test.ts)
- [Soft line breaks inside a text run](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-linebreaks.test.ts)
- [Replace text by HTML](https://github.com/singerla/pptx-automizer/blob/main/__tests__/replace-multi-text-html.test.ts)
