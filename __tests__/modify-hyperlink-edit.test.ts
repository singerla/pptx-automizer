import Automizer from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { modify } from '../src/index';

test('Modify an existing external hyperlink', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `modify-hyperlink-edit.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);
  const newUrl = 'https://edited-link.example.com';

  // First, let's check the original slide to get the original relationship ID
  const originalFile = path.join(`${__dirname}/pptx-templates`, 'SlideWithLink.pptx');
  const originalData = fs.readFileSync(originalFile);
  const originalZip = await JSZip.loadAsync(originalData);
  const originalRelsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const originalRelsFile = originalZip.file(originalRelsPath);
  const originalRelsXml = await originalRelsFile!.async('text');
  console.log('Original slide relationships XML:', originalRelsXml);

  // Add a console.log statement to see what's happening in the modify function
  console.log('About to modify slide with URL:', newUrl);

  // Instead of using modifyElement, we'll directly modify the slide's XML
  const result = await pres
    .addSlide('link', 1, (slide) => {
      // Add a modification to the slide's relations
      slide.modifyRelations(async (slideRelXml) => {
        console.log('Inside modifyRelations callback');

        // Find all hyperlink relationships
        const relationships = slideRelXml.getElementsByTagName('Relationship');
        console.log(`Found ${relationships.length} relationships in the slide`);

        // Look for hyperlink relationships
        for (let i = 0; i < relationships.length; i++) {
          const rel = relationships[i];
          const type = rel.getAttribute('Type');
          const target = rel.getAttribute('Target');

          console.log(`Relationship ${i+1}: Type=${type}, Target=${target}`);

          // If this is a hyperlink relationship, update its target
          if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
            console.log(`Updating hyperlink target from ${target} to ${newUrl}`);
            rel.setAttribute('Target', newUrl);
          }
        }
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
  console.log('Slide relationships XML after modification:', slideRelsXml);

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideRelsXml, 'application/xml');

  // Look for hyperlink relationships with the new URL
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  console.log(`Found ${relationships.length} relationships`);

  let foundEditedLink = false;

  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    const target = relationship.getAttribute('Target');
    console.log(`Relationship ${i+1}: Type=${type}, Target=${target}`);

    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' &&
        target === newUrl) {
      foundEditedLink = true;
      console.log('Found the edited hyperlink!');
      break;
    }
  }

  // Verify that the hyperlink was edited
  expect(foundEditedLink).toBe(true);
});
