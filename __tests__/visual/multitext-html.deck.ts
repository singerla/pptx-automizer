/**
 * Golden deck `multitext-html` (Tier 3): htmlToMultiText — nested (CKEditor
 * sibling-form) bullet lists, ordered lists, and inline styling. Pins current
 * behavior; the HTML→text feature track updates these baselines intentionally.
 * Mirrors replace-multi-text-html, but judged by rendered pixels.
 */
import Automizer, { modify } from '../../src/index';
import { expectDeckMatchesBaselines } from './helpers/golden-deck';

test('golden deck: multitext-html', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../pptx-templates`,
    outputDir: `${__dirname}/../pptx-output`,
  });

  // The invalid sibling-form nesting CKEditor produces
  const nestedLists =
    '<html><body><p>First Line 14pt</p>' +
    '<p><span style="font-size: 12px;">2nd line 12pt <strong>bold</strong> <em>italics</em></span></p>' +
    '<ul>' +
    '<li><span style="font-size: 14px;">bullet 1 level 1</span></li>' +
    '<ul><li><span style="font-size: 14px;">bullet 1 level 2</span></li></ul>' +
    '<li><span style="font-size: 14px;">bullet 2 level 1</span></li>' +
    '<ul>' +
    '<li><span style="font-size: 14px;">bullet 2 level 2</span></li>' +
    '<ul><li><span style="font-size: 14px;">bullet 2 level 3</span></li></ul>' +
    '<li><span style="font-size: 14px;"><ins>bullet</ins> <em>mixed</em> <strong><em>formatting</em></strong></span></li>' +
    '</ul>' +
    '</ul>' +
    '<p><span style="font-size: 14px;"><strong><em>Text </em></strong>after bullet list</span></p></body></html>';

  const orderedLists =
    '<html><body><ol><li>first</li><li>second</li>' +
    '<ol><li>nested</li></ol></ol></body></html>';

  const inlineStyles =
    '<html><body>' +
    '<p>before<br />after</p>' +
    '<p><span style="color: rgb(0, 128, 255)">rgb</span> ' +
    '<span style="color: navy">named</span></p>' +
    '<p>H<sub>2</sub>O and x<sup>2</sup></p>' +
    '<p><mark>marked</mark></p>' +
    '</body></html>';

  const outputFile = 'visual-multitext-html.pptx';
  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`TextReplace.pptx`)
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(nestedLists));
    })
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(orderedLists));
    })
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(inlineStyles));
    })
    .write(outputFile);

  await expectDeckMatchesBaselines('multitext-html', outputFile, 4);
});
