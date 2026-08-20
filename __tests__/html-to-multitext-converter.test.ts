import { HtmlToMultiTextHelper } from '../src/helper/html-to-multitext-helper';
import { MultiTextParagraph } from '../src/interfaces/imulti-text';

const convert = (body: string): MultiTextParagraph[] =>
  new HtmlToMultiTextHelper().run(`<html><body>${body}</body></html>`);

/** Compact view of the flat structure a converted tree must project onto. */
const outline = (paragraphs: MultiTextParagraph[]) =>
  paragraphs.map((paragraph) => ({
    level: paragraph.paragraph.level,
    bullet: paragraph.paragraph.bullet,
    text: paragraph.textRuns.map((run) => run.text ?? '<br>').join(''),
  }));

describe('list projection', () => {
  // PPTX `lvl` is 0-based; a top-level bullet used to come out as level 1 and
  // therefore rendered one indent step too deep
  test('top-level bullets are level 0', () => {
    expect(outline(convert('<ul><li>top</li></ul>'))).toEqual([
      { level: 0, bullet: true, text: 'top' },
    ]);
  });

  test('a plain paragraph is level 0 without a bullet', () => {
    expect(outline(convert('<p>plain</p>'))).toEqual([
      { level: 0, bullet: false, text: 'plain' },
    ]);
  });

  // The two list-nesting shapes must not diverge: CKEditor emits the invalid
  // sibling form, hand-written HTML the properly nested one
  test('sibling-nested and properly nested lists converge', () => {
    const sibling = convert(
      '<ul><li>a</li><ul><li>b</li><ul><li>c</li></ul></ul><li>d</li></ul>',
    );
    const nested = convert(
      '<ul><li>a<ul><li>b<ul><li>c</li></ul></li></ul></li><li>d</li></ul>',
    );

    expect(outline(sibling)).toEqual([
      { level: 0, bullet: true, text: 'a' },
      { level: 1, bullet: true, text: 'b' },
      { level: 2, bullet: true, text: 'c' },
      { level: 0, bullet: true, text: 'd' },
    ]);
    expect(outline(nested)).toEqual(outline(sibling));
  });

  test('ordered lists carry automatic numbering, varying per level', () => {
    const paragraphs = convert('<ol><li>one</li><ol><li>two</li></ol></ol>');

    expect(paragraphs[0].paragraph.bulletType).toBe('number');
    expect(paragraphs[0].paragraph.autoNumberType).toBe('arabicPeriod');
    expect(paragraphs[1].paragraph.autoNumberType).toBe('alphaLcPeriod');
  });

  test('unordered lists stay glyph bullets', () => {
    expect(convert('<ul><li>a</li></ul>')[0].paragraph.bulletType).toBeUndefined();
  });

  test('level saturates at the deepest level PPTX can express', () => {
    const deep = '<ul>'.repeat(12) + '<li>deep</li>' + '</ul>'.repeat(12);

    expect(convert(deep)[0].paragraph.level).toBe(8);
  });
});

describe('block projection', () => {
  // A block inside a block must not nest: the innermost one becomes the
  // paragraph, the ancestors only contribute properties
  test('<li><p>text</p></li> yields one bulleted paragraph', () => {
    expect(outline(convert('<ul><li><p>text</p></li></ul>'))).toEqual([
      { level: 0, bullet: true, text: 'text' },
    ]);
  });

  test('blockquote inside a list item keeps the bullet', () => {
    expect(
      outline(convert('<ul><li><blockquote>quoted</blockquote></li></ul>')),
    ).toEqual([{ level: 0, bullet: true, text: 'quoted' }]);
  });

  // This text used to be dropped silently
  test('bare text in body and div is kept, as separate paragraphs', () => {
    expect(outline(convert('loose<div>in div</div>'))).toEqual([
      { level: 0, bullet: false, text: 'loose' },
      { level: 0, bullet: false, text: 'in div' },
    ]);
  });

  test('text in an unknown element is kept', () => {
    expect(outline(convert('<section><custom>kept</custom></section>'))).toEqual(
      [{ level: 0, bullet: false, text: 'kept' }],
    );
  });

  test('an empty paragraph is a deliberate blank line', () => {
    expect(outline(convert('<p>a</p><p></p><p>b</p>'))).toEqual([
      { level: 0, bullet: false, text: 'a' },
      { level: 0, bullet: false, text: '' },
      { level: 0, bullet: false, text: 'b' },
    ]);
  });

  test('a wrapper div around a list does not add a blank paragraph', () => {
    expect(outline(convert('<div><ul><li>a</li></ul></div>'))).toEqual([
      { level: 0, bullet: true, text: 'a' },
    ]);
  });

  test('headings become styled paragraphs', () => {
    const [h1, h3] = convert('<h1>Title</h1><h3>Sub</h3>');

    expect(h1.textRuns[0].style).toMatchObject({ size: 2400, isBold: true });
    expect(h3.textRuns[0].style).toMatchObject({ size: 1800, isBold: true });
    expect(h1.paragraph.bullet).toBe(false);
  });

  test('text-align maps onto the paragraph', () => {
    expect(
      convert('<p style="text-align: center">c</p>')[0].paragraph.alignment,
    ).toBe('ctr');
    expect(
      convert('<p style="text-align: justify">j</p>')[0].paragraph.alignment,
    ).toBe('just');
    // Unset, so the shape's layout keeps deciding - HTML inherits alignment
    // from its container too
    expect(convert('<p>default</p>')[0].paragraph.alignment).toBeUndefined();
  });
});

describe('inline styles', () => {
  test('nested inline tags accumulate into one run', () => {
    const runs = convert('<p><strong><em>both</em></strong></p>')[0].textRuns;

    expect(runs).toHaveLength(1);
    expect(runs[0].style).toMatchObject({ isBold: true, isItalics: true });
  });

  test('tag coverage: u, s, del, sub, sup, code, mark', () => {
    const styleOf = (html: string) => convert(`<p>${html}</p>`)[0].textRuns[0].style;

    expect(styleOf('<u>x</u>')).toMatchObject({ isUnderlined: true });
    expect(styleOf('<ins>x</ins>')).toMatchObject({ isUnderlined: true });
    expect(styleOf('<s>x</s>')).toMatchObject({ isStrike: true });
    expect(styleOf('<strike>x</strike>')).toMatchObject({ isStrike: true });
    expect(styleOf('<del>x</del>')).toMatchObject({ isStrike: true });
    expect(styleOf('<sub>x</sub>')).toMatchObject({ isSubscript: true });
    expect(styleOf('<sup>x</sup>')).toMatchObject({ isSuperscript: true });
    expect(styleOf('<code>x</code>')).toMatchObject({ fontFamily: 'Consolas' });
    expect(styleOf('<mark>x</mark>')).toMatchObject({
      highlight: { type: 'srgbClr', value: 'FFFF00' },
    });
  });

  test('<br> becomes a break run, not text', () => {
    const runs = convert('<p>line1<br />line2</p>')[0].textRuns;

    expect(runs).toEqual([
      { text: 'line1', style: {} },
      { break: true, style: {} },
      { text: 'line2', style: {} },
    ]);
  });

  test('hyperlinks: numeric href is an internal slide link', () => {
    expect(convert('<p><a href="3">to slide</a></p>')[0].textRuns[0].style).toMatchObject(
      { hyperlink: { target: 3, isInternal: true } },
    );
    expect(
      convert('<p><a href="https://example.com">ext</a></p>')[0].textRuns[0].style,
    ).toMatchObject({
      hyperlink: { target: 'https://example.com', isInternal: false },
    });
  });
});

describe('CSS mapping', () => {
  const styleOf = (css: string) =>
    convert(`<p><span style="${css}">x</span></p>`)[0].textRuns[0].style;

  // A leading '#' is not valid in <a:srgbClr val="...">
  test('colors normalize to 6-digit hex without #', () => {
    expect(styleOf('color: #ff0000').color).toEqual({
      type: 'srgbClr',
      value: 'FF0000',
    });
    expect(styleOf('color: #f00').color.value).toBe('FF0000');
    expect(styleOf('color: rgb(0, 128, 255)').color.value).toBe('0080FF');
    expect(styleOf('color: rgba(0, 128, 255, 0.5)').color.value).toBe('0080FF');
    expect(styleOf('color: navy').color.value).toBe('000080');
  });

  test('an unusable color is skipped instead of written invalid', () => {
    expect(styleOf('color: transparent').color).toBeUndefined();
  });

  // 1px is 0.75pt at 96dpi - px used to be treated as pt
  test('font-size converts px to 1/100 pt, and accepts pt', () => {
    expect(styleOf('font-size: 12px').size).toBe(900);
    expect(styleOf('font-size: 16px').size).toBe(1200);
    expect(styleOf('font-size: 14pt').size).toBe(1400);
    expect(styleOf('font-size: 13.5pt').size).toBe(1350);
  });

  test('relative font sizes stay inherited rather than guessed', () => {
    expect(styleOf('font-size: 1.2em').size).toBeUndefined();
    expect(styleOf('font-size: 120%').size).toBeUndefined();
  });

  test('font-weight, font-style, text-decoration, font-family', () => {
    expect(styleOf('font-weight: bold').isBold).toBe(true);
    expect(styleOf('font-weight: 700').isBold).toBe(true);
    expect(styleOf('font-weight: 400').isBold).toBe(false);
    expect(styleOf('font-style: italic').isItalics).toBe(true);
    expect(styleOf('text-decoration: underline').isUnderlined).toBe(true);
    expect(styleOf('text-decoration: line-through').isStrike).toBe(true);
    expect(styleOf("font-family: 'Segoe UI', sans-serif").fontFamily).toBe(
      'Segoe UI',
    );
    expect(styleOf('font-family: Arial, sans-serif').fontFamily).toBe('Arial');
  });

  test('background-color becomes a highlight', () => {
    expect(styleOf('background-color: yellow').highlight).toEqual({
      type: 'srgbClr',
      value: 'FFFF00',
    });
  });

  test('CSS on an element overrides the tag it sits on', () => {
    expect(
      convert('<p><strong style="font-weight: normal">x</strong></p>')[0]
        .textRuns[0].style.isBold,
    ).toBe(false);
  });

  test('inline CSS wins over block CSS', () => {
    const runs = convert(
      '<p style="color: red"><span style="color: blue">x</span></p>',
    )[0].textRuns;

    expect(runs[0].style.color.value).toBe('0000FF');
  });
});

describe('whitespace', () => {
  test('runs of whitespace collapse to a single space', () => {
    expect(convert('<p>spaced   \n  out</p>')[0].textRuns[0].text).toBe(
      'spaced out',
    );
  });

  test('leading and trailing whitespace is dropped per paragraph', () => {
    expect(convert('<p>\n  padded  \n</p>')[0].textRuns[0].text).toBe('padded');
  });

  test('whitespace between block tags does not become a paragraph', () => {
    expect(convert('<p>a</p>\n  \n<p>b</p>')).toHaveLength(2);
  });

  // &nbsp; is what editors emit for deliberate spacing - it is not
  // collapsible whitespace, so it must survive both collapsing and trimming
  test('a non-breaking space survives collapsing and trimming', () => {
    expect(convert('<p>a&nbsp;&nbsp;b</p>')[0].textRuns[0].text).toBe(
      'a\u00A0\u00A0b',
    );
    expect(convert('<p>&nbsp;</p>')[0].textRuns[0].text).toBe('\u00A0');
  });

  test('no literal newline ends up inside a text run', () => {
    convert('<p>line\nbreak</p>')[0].textRuns.forEach((run) => {
      expect(run.text ?? '').not.toMatch(/[\r\n]/);
    });
  });
});

// The browser's default stylesheet gives <p>, headings and lists a vertical
// margin (`1em 0`, adjacent margins collapsed); without projecting that onto
// spaceBefore, every paragraph sits flush against the previous one
describe('vertical spacing', () => {
  const spaceBefore = (paragraphs: MultiTextParagraph[]) =>
    paragraphs.map((paragraph) => paragraph.paragraph.spaceBefore);

  test('block paragraphs get one collapsed gap, the first none', () => {
    expect(spaceBefore(convert('<p>a</p><p>b</p><h1>c</h1>'))).toEqual([
      undefined,
      { percent: 100 },
      { percent: 100 },
    ]);
  });

  test('list edges get the gap, list items inside sit tight', () => {
    const paragraphs = convert(
      '<p>a</p><ul><li>b</li><ul><li>c</li></ul><li>d</li></ul><p>e</p>',
    );

    expect(spaceBefore(paragraphs)).toEqual([
      undefined, // a
      { percent: 100 }, // b: list top margin
      undefined, // c: nested list, no margin
      undefined, // d
      { percent: 100 }, // e: list bottom margin
    ]);
  });

  test('two adjacent lists keep their outer margins apart', () => {
    expect(
      spaceBefore(convert('<ul><li>a</li></ul><ol><li>b</li></ol>')),
    ).toEqual([undefined, { percent: 100 }]);
  });

  test('divs carry no margin of their own', () => {
    expect(spaceBefore(convert('<div>a</div><div>b</div>'))).toEqual([
      undefined,
      undefined,
    ]);
  });
});

// A <br> right before a closing block tag is invisible in HTML - editors
// append one after almost every link. It must not become an extra empty line.
describe('trailing line breaks', () => {
  test('a trailing <br/> and the space before it are dropped', () => {
    expect(outline(convert('<p><a href="https://example.com">x</a> <br/></p>')))
      .toEqual([{ level: 0, bullet: false, text: 'x' }]);
    expect(outline(convert('<ul><li>item <br/></li></ul>'))).toEqual([
      { level: 0, bullet: true, text: 'item' },
    ]);
  });

  test('a deliberate double <br/> keeps its one empty line', () => {
    expect(outline(convert('<p>a<br/><br/></p>'))).toEqual([
      { level: 0, bullet: false, text: 'a<br>' },
    ]);
  });

  test('<p><br/></p> stays a blank line', () => {
    expect(outline(convert('<p><br/></p>'))).toEqual([
      { level: 0, bullet: false, text: '' },
    ]);
  });

  test('a <br/> between runs is untouched', () => {
    expect(outline(convert('<p>a<br/>b</p>'))).toEqual([
      { level: 0, bullet: false, text: 'a<br>b' },
    ]);
  });
});

test('missing body tag yields no paragraphs instead of throwing', () => {
  expect(new HtmlToMultiTextHelper().run('<p>no body</p>')).toEqual([]);
});
