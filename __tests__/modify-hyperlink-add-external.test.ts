import Automizer from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { modify } from '../src/index';

test('Add a new external hyperlink', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  const outputFile = `modify-hyperlink-add-external.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);
  const newUrl = 'https://new-hyperlink.example.com';

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      // Find a text shape and add a hyperlink to it
      // The EmptySlide template has a text field named "Textfeld 3" instead of "Title"
      slide.modifyElement('Textfeld 3', (element, relation) => {
        // Using the addHyperlink function directly - notice the lowercase "a"
        modify.addHyperlink(newUrl)(element, relation);
      });
    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(2);

  // Read the generated PPTX file
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);

  // Check relationships file for slide 2
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();

  const slideRelsXml = await slideRelsFile!.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideRelsXml, 'application/xml');

  // Look for hyperlink relationships with the new URL
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  let foundHyperlink = false;
  let hyperlinkId = '';

  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    const target = relationship.getAttribute('Target');
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' &&
        target === newUrl) {
      foundHyperlink = true;
      hyperlinkId = relationship.getAttribute('Id') || '';
      break;
    }
  }

  // Verify that the hyperlink was added
  expect(foundHyperlink).toBe(true);
  expect(hyperlinkId).not.toBe('');

  // Now check if the slide XML contains the hyperlink reference
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();

  const slideXml = await slideFile!.async('text');

  // Verify that the hyperlink ID is referenced in the slide content
  expect(slideXml.includes(`r:id="${hyperlinkId}"`)).toBe(true);
});
