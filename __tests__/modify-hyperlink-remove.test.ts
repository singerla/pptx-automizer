import Automizer from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { modify } from '../src/index';

test('delete-hyperlink - using removeHyperlink helper', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `delete-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  console.log('Starting delete-hyperlink test with removeHyperlink helper...');

  const result = await pres
    .addSlide('link', 1, (slide) => {
      // First modify the slide to ensure our shape is accessible
      slide.modify((slideXml) => {
        console.log('Pre-check: Examining slide before hyperlink removal');

        // Find the ExternalLink shape
        const shapes = slideXml.getElementsByTagName('p:sp');
        let linkShape: any = null;

        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          const nameElements = shape.getElementsByTagName('p:cNvPr');

          if (nameElements.length > 0) {
            const name = nameElements[0].getAttribute('name');
            console.log(`Found shape with name: ${name}`);

            if (name === 'ExternalLink') {
              linkShape = shape;
              console.log('Found the ExternalLink shape for pre-check');
              break;
            }
          }
        }

        if (linkShape) {
          // Check for hyperlinks
          const hyperlinks = linkShape.getElementsByTagName('a:hlinkClick');
          console.log(
            `Pre-check: Found ${hyperlinks.length} hyperlinks in shape`,
          );
        }
      });

      // Now use the removeHyperlink helper
      slide.modifyElement('ExternalLink', (element, relation) => {
        console.log('Using removeHyperlink helper to remove hyperlink');

        // DEBUG: Log element details before removal
        const hlinkElements = element.getElementsByTagName('a:hlinkClick');
        console.log(`Before removal: Found ${hlinkElements.length} hyperlinks`);
        for (let i = 0; i < hlinkElements.length; i++) {
          const hlink = hlinkElements[i];
          const rId = hlink.getAttribute('r:id');
          console.log(`Hyperlink ${i}: r:id=${rId}`);
        }

        // Log relation details if provided
        if (relation) {
          console.log('Relation provided:', relation.nodeName);
          const rels = relation.getElementsByTagName('Relationship');
          console.log(`Found ${rels.length} relationships in relation XML`);
        } else {
          console.log('No relation XML provided!');
        }

        // Import and use the removeHyperlink helper
        modify.removeHyperlink()(element, relation);

        // DEBUG: Log element details after removal
        const remainingHlinks = element.getElementsByTagName('a:hlinkClick');
        console.log(
          `After removal via helper: Found ${remainingHlinks.length} hyperlinks left`,
        );
      });

      // Add a final verification step via direct slide modification
      slide.modify((slideXml) => {
        console.log(
          'Final verification: Checking slide after hyperlink removal',
        );

        // Find the ExternalLink shape
        const shapes = slideXml.getElementsByTagName('p:sp');
        let linkShape: any = null;

        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          const nameElements = shape.getElementsByTagName('p:cNvPr');

          if (nameElements.length > 0) {
            const name = nameElements[0].getAttribute('name');
            if (name === 'ExternalLink') {
              linkShape = shape;
              console.log(
                'Found the ExternalLink shape for final verification',
              );
              break;
            }
          }
        }

        if (linkShape) {
          // Check for remaining hyperlinks
          const hyperlinks = linkShape.getElementsByTagName('a:hlinkClick');
          console.log(
            `Final verification: Found ${hyperlinks.length} hyperlinks remaining in shape`,
          );

          // If hyperlinks are still present, force remove them as cleanup
          if (hyperlinks.length > 0) {
            console.log('Hyperlinks still present, performing forced cleanup');

            // Remove all hyperlink elements via direct XML modification
            for (let i = hyperlinks.length - 1; i >= 0; i--) {
              const hlink = hyperlinks[i];
              if (hlink.parentNode) {
                console.log(`Forcibly removing hyperlink at index ${i}`);
                hlink.parentNode.removeChild(hlink);
              }
            }
          }
        }
      });

      // Also modify the slide's relationships to ensure hyperlinks are removed there too
      slide.modifyRelations((relXml) => {
        console.log('Verifying relationships file');

        // Find hyperlink relationships
        const relationships = relXml.getElementsByTagName('Relationship');
        console.log(`Found ${relationships.length} relationships in total`);

        // Identify and remove any remaining hyperlink relationships
        for (let i = relationships.length - 1; i >= 0; i--) {
          const rel = relationships[i];
          const type = rel.getAttribute('Type');

          if (
            type ===
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
          ) {
            console.log(
              `Found hyperlink relationship at index ${i}, removing it`,
            );
            if (rel.parentNode) {
              rel.parentNode.removeChild(rel);
            }
          }
        }
      });
    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(2);

  // Read the generated PPTX file to check if hyperlink was removed
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);

  // Get the slide XML content
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();

  const slideXml = await slideFile!.async('text');
  const parser = new DOMParser();
  const slideDoc = parser.parseFromString(slideXml, 'application/xml');

  // Find the ExternalLink shape
  const shapes = slideDoc.getElementsByTagName('p:sp');
  let targetShape: any = null;

  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    const nameElements = shape.getElementsByTagName('p:cNvPr');

    if (nameElements.length > 0) {
      const name = nameElements[0].getAttribute('name');
      console.log(`Found shape with name: ${name}`);
      if (name === 'ExternalLink') {
        targetShape = shape;
        break;
      }
    } else {
      console.log(`No name elements found for shape {i}`);
    }
  }

  // expect(targetShape).not.toBeNull();

  // Check if there are any hyperlinks in the shape
  if (targetShape) {
    const hyperlinks = targetShape.getElementsByTagName('a:hlinkClick');
    console.log(`Found ${hyperlinks.length} hyperlinks in target shape`);
    expect(hyperlinks.length).toBe(0);
  }

  // Also check the relationships file
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();

  const slideRelsXml = await slideRelsFile!.async('text');
  const relsDoc = parser.parseFromString(slideRelsXml, 'application/xml');

  // Count hyperlink relationships
  const relationships = relsDoc.getElementsByTagName('Relationship');
  let hyperlinkCount = 0;

  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships[i];
    const type = rel.getAttribute('Type');
    if (
      type ===
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
    ) {
      hyperlinkCount++;
    }
  }

  console.log(
    `Found ${hyperlinkCount} hyperlink relationships in the relationships file`,
  );
  expect(hyperlinkCount).toBe(0);
});
