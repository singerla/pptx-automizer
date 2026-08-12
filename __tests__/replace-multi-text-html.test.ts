import Automizer, { modify } from '../src/index';
import JSZip from 'jszip';
import fs from 'fs';

/**
 * Read back the txBody of a named shape from a written .pptx, so the
 * assertions below are about the XML that PowerPoint will actually open.
 */
const readTxBody = async (file: string, shapeName: string): Promise<string> => {
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${file}`),
  );
  const xml = await archive.file('ppt/slides/slide2.xml').async('string');
  const shape = (xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []).find((candidate) =>
    candidate.includes(`name="${shapeName}"`),
  );

  expect(shape).toBeDefined();

  return shape.match(/<p:txBody>[\s\S]*?<\/p:txBody>/)[0];
};

/** All <a:p> elements of a txBody, in document order. */
const paragraphsOf = (txBody: string): string[] =>
  txBody.match(/<a:p>[\s\S]*?<\/a:p>/g) || [];

const buildFromHtml = async (
  html: string,
  outputFile: string,
): Promise<string> => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`TextReplace.pptx`)
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(outputFile);

  return readTxBody(outputFile, 'setText');
};

test('create presentation, replace multi text from HTML string.', async () => {
  // The list nesting here is the invalid sibling form CKEditor produces
  const html =
    '<html><body><p>First Line 14pt</p>\n' +
    '<p><span style="font-size: 12px;">2nd line 12pt <strong>bold</strong> <em>italics</em></span></p>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 1 level 1</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 1 level 2</span></li>\n' +
    '</ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 1</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 2</span></li>\n' +
    '<ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 3</span></li>\n' +
    '</ul>\n' +
    '<li><span style="font-size: 14px;">bullet 2 level 2</span></li>\n' +
    '<li><span style="font-size: 14px;"><ins>bullet</ins> <em>mixed</em> <strong><em>formatting</em></strong></span></li>\n' +
    '</ul>\n' +
    '</ul>\n' +
    '<p><span style="font-size: 14px;"><strong><em>Text </em></strong>after bullet list</span></p></body></html>\n';

  const txBody = await buildFromHtml(html, `modify-multi-text-html.test.pptx`);
  const paragraphs = paragraphsOf(txBody);

  // 2 plain paragraphs, 7 bullets, 1 trailing paragraph
  expect(paragraphs).toHaveLength(10);

  // Plain paragraphs: level 0, bullet explicitly suppressed
  expect(paragraphs[0]).toContain('<a:buNone/>');
  expect(paragraphs[0]).toContain('<a:t>First Line 14pt</a:t>');

  // 12px is 9pt, not 12pt
  expect(paragraphs[1]).toContain('sz="900"');

  // Bullet levels are 0-based, and every level gets its own indent
  const levels = paragraphs
    .slice(2, 8)
    .map((paragraph) => paragraph.match(/lvl="(\d)"/)[1]);
  expect(levels).toEqual(['0', '1', '0', '1', '2', '1']);

  expect(paragraphs[2]).toContain('marL="228600"');
  expect(paragraphs[3]).toContain('marL="457200"');
  expect(paragraphs[6]).toContain('marL="685800"');

  // Bullets carry the typeface that holds the glyph
  expect(paragraphs[2]).toContain('<a:buFont typeface="Arial"/><a:buChar char="•"/>');

  // Accumulated inline styling: <strong><em> is one run with both properties
  const lastBullet = paragraphs[8];
  expect(lastBullet).toMatch(/<a:rPr[^>]*u="sng"[^>]*\/><a:t>bullet<\/a:t>/);
  expect(lastBullet).toMatch(/<a:rPr[^>]*b="1"[^>]*i="1"[^>]*\/><a:t>formatting<\/a:t>/);

  // The trailing paragraph is not a bullet again
  expect(paragraphs[9]).toContain('<a:buNone/>');
  expect(paragraphs[9]).toContain('<a:t>after bullet list</a:t>');
});

test('htmlToMultiText writes schema-valid <a:pPr> child order', async () => {
  // buChar/buNone after the spacing elements - the reverse is the child-order
  // violation PowerPoint offers to "repair"
  const html =
    '<html><body><ul><li>spaced bullet</li></ul>' +
    '<p>plain</p></body></html>';

  const txBody = await buildFromHtml(
    html,
    `modify-multi-text-html-order.test.pptx`,
  );

  paragraphsOf(txBody).forEach((paragraph) => {
    const pPr = paragraph.match(/<a:pPr[^>]*>([\s\S]*?)<\/a:pPr>/);
    if (!pPr) {
      return;
    }

    const children = Array.from(pPr[1].matchAll(/<(a:[a-zA-Z]+)/g)).map(
      (match) => match[1],
    );
    const order = [
      'a:lnSpc',
      'a:spcBef',
      'a:spcAft',
      'a:buFont',
      'a:buNone',
      'a:buAutoNum',
      'a:buChar',
    ];
    const ranks = children
      .filter((child) => order.includes(child))
      .map((child) => order.indexOf(child));

    expect(ranks).toEqual([...ranks].sort((a, b) => a - b));
  });
});

test('htmlToMultiText renders ordered lists as automatic numbering', async () => {
  const html =
    '<html><body><ol><li>first</li><li>second</li>' +
    '<ol><li>nested</li></ol></ol></body></html>';

  const txBody = await buildFromHtml(
    html,
    `modify-multi-text-html-ordered.test.pptx`,
  );
  const paragraphs = paragraphsOf(txBody);

  // Numbers, not '•' glyphs
  expect(txBody).not.toContain('<a:buChar');
  expect(paragraphs[0]).toContain('<a:buAutoNum type="arabicPeriod"/>');
  expect(paragraphs[1]).toContain('<a:buAutoNum type="arabicPeriod"/>');
  // PowerPoint varies the scheme per level
  expect(paragraphs[2]).toContain('<a:buAutoNum type="alphaLcPeriod"/>');
});

test('htmlToMultiText maps <br>, colors, sub/sup and highlights', async () => {
  const html =
    '<html><body>' +
    '<p>before<br />after</p>' +
    '<p><span style="color: rgb(0, 128, 255)">rgb</span>' +
    '<span style="color: navy">named</span></p>' +
    '<p>H<sub>2</sub>O and x<sup>2</sup></p>' +
    '<p><mark>marked</mark></p>' +
    '</body></html>';

  const txBody = await buildFromHtml(
    html,
    `modify-multi-text-html-inline.test.pptx`,
  );

  // A break is an element between runs, never text inside <a:t>
  expect(txBody).toMatch(
    /<a:t>before<\/a:t><\/a:r><a:br>[\s\S]*?<\/a:br><a:r>[\s\S]*?<a:t>after<\/a:t>/,
  );

  // Colors must be 6-digit hex without '#' - anything else is invalid OOXML
  expect(txBody).toContain('<a:srgbClr val="0080FF"/>');
  expect(txBody).toContain('<a:srgbClr val="000080"/>');
  Array.from(txBody.matchAll(/<a:srgbClr val="([^"]*)"/g)).forEach((match) => {
    expect(match[1]).toMatch(/^[0-9A-Fa-f]{6}$/);
  });

  expect(txBody).toMatch(/baseline="-25000"[^>]*\/><a:t>2<\/a:t>/);
  expect(txBody).toMatch(/baseline="30000"[^>]*\/><a:t>2<\/a:t>/);

  expect(txBody).toContain('<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>');
});

test('htmlToMultiText keeps hyperlinks, with the styling in schema order', async () => {
  // A numeric href is a slide *number in the output deck*, so the deck needs
  // enough slides for slide4.xml to exist - a hyperlink whose relationship
  // target is not in the archive is silently dead in PowerPoint.
  const html =
    '<html><body><p>' +
    '<a href="https://example.com" style="color: #ff0000">external</a> ' +
    '<a href="4">to slide 4</a>' +
    '</p></body></html>';

  const outputFile = `modify-multi-text-html-hyperlinks.test.pptx`;
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const result = await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`TextReplace.pptx`)
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .addSlide('TextReplace.pptx', 1)
    .addSlide('TextReplace.pptx', 1)
    .write(outputFile);

  expect(result.slides).toBe(4);

  const txBody = await readTxBody(outputFile, 'setText');

  // <a:solidFill> must precede <a:hlinkClick> inside <a:rPr>. Note that
  // PowerPoint still paints hyperlink text in the theme's hlink color - the
  // run fill is written, but does not win for a linked run.
  expect(txBody).toMatch(
    /<a:rPr[^>]*><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><a:hlinkClick/,
  );

  // An internal link needs the slide-jump action, or PowerPoint treats the
  // run as plain underlined text
  expect(txBody).toContain('action="ppaction://hlinksldjump"');

  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${outputFile}`),
  );
  const rels = await archive
    .file('ppt/slides/_rels/slide2.xml.rels')
    .async('string');

  expect(rels).toContain('Target="https://example.com"');
  expect(rels).toContain('TargetMode="External"');
  expect(rels).toContain('Target="../slides/slide4.xml"');

  // Every non-external relationship target must resolve to a part that is
  // actually in the archive; a dangling target is the corruption class that
  // makes a link quietly stop working
  const targets = Array.from(
    rels.matchAll(/Target="([^"]+)"(?![^>]*TargetMode="External")/g),
  ).map((match) => match[1]);

  targets.forEach((target) => {
    const resolved = target.replace(/^\.\.\//, 'ppt/');
    expect(archive.file(resolved)).not.toBeNull();
  });
});

test('htmlToMultiText does not inherit bold from the template placeholder', async () => {
  // The template's first run is italic at 20pt; size may carry over as the
  // shape's look, but bold/italic of that one run must not style everything
  const txBody = await buildFromHtml(
    '<html><body><p>plain text</p></body></html>',
    `modify-multi-text-html-default-style.test.pptx`,
  );

  const run = txBody.match(/<a:r>[\s\S]*?<\/a:r>/)[0];

  expect(run).toContain('<a:t>plain text</a:t>');
  expect(run).not.toContain('i="1"');
  expect(run).not.toContain('b="1"');
});
