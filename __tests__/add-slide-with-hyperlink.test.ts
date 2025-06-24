import Automizer from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

test('Add a full slide with hyperlinks on', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `copy-slide-with-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  const result = await pres
    .addSlide('link', 1)
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

  // Look for hyperlink relationships
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  let hasExternalHyperlink = false;
  let hasInternalHyperlink = false;
  let externalHyperlinkId = '';
  let internalHyperlinkId = '';

  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
      const targetMode = relationship.getAttribute('TargetMode');
      const id = relationship.getAttribute('Id') || '';

      if (targetMode === 'External') {
        hasExternalHyperlink = true;
        externalHyperlinkId = id;
      } else if (!targetMode) {
        hasInternalHyperlink = true;
        internalHyperlinkId = id;
      }
    }
  }

  // Verify that hyperlink relationships exist
  expect(hasExternalHyperlink).toBe(true);
  expect(externalHyperlinkId).not.toBe('');

  // Check the slide XML content
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();

  const slideXml = await slideFile!.async('text');

  // Verify hyperlink references in slide content
  expect(slideXml.includes(`r:id="${externalHyperlinkId}"`)).toBe(true);

  if (hasInternalHyperlink) {
    expect(slideXml.includes(`r:id="${internalHyperlinkId}"`)).toBe(true);
  }
});
