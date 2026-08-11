import Automizer, { modify } from './src/index';
import JSZip from 'jszip';
import fs from 'fs';

const html =
  '<html><body>' +
  '<h2 style="text-align: center">Report</h2>' +
  '<p><span style="font-size: 12px; color: #ff0000">red 9pt</span> and <strong>bold</strong></p>' +
  '<ul><li>bullet 1</li><ul><li><em>nested</em> <s>gone</s></li></ul></ul>' +
  '<ol><li>first</li><li>second<br />after break</li></ol>' +
  '<p><a href="https://example.com" style="background-color: yellow">link</a></p>' +
  '</body></html>';

(async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/__tests__/pptx-templates`,
    outputDir: `${__dirname}/__tests__/pptx-output`,
  });

  await automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`TextReplace.pptx`)
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', modify.htmlToMultiText(html));
    })
    .write(`probe-html.pptx`);

  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/__tests__/pptx-output/probe-html.pptx`),
  );
  const xml = await archive.file('ppt/slides/slide2.xml').async('string');
  const body = xml.match(/<p:txBody>[\s\S]*?<\/p:txBody>/);
  console.log((body ? body[0] : xml).replace(/></g, '>\n<'));
})();
