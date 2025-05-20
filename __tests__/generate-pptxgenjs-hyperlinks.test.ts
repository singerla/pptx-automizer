import { Automizer } from '../src';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

describe('Hyperlink works with generate', () => {
  const outputDir = path.join(__dirname, 'pptx-output');
  const outputFile = 'generate-pptxgenjs-hyperlinks.test.pptx';
  const outputPath = path.join(outputDir, outputFile);

  // Clean up before tests if the file exists
  beforeAll(() => {
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
  });

  it('should correctly generate external and internal hyperlinks using slide.generate', async () => {
    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'pptx-templates'),
      outputDir: outputDir,
    });

    // Load root and template files
    automizer.loadRoot('EmptyTemplate.pptx');
    automizer.load('EmptySlide.pptx', 'template');

    const externalUrls = [
      'https://google.com/1',
      'https://google.com/2',
      'https://google.com/3',
    ];
    const targetInternalSlides = [2, 3, 4]; // slide numbers in the final presentation

    // Create slide 1 with hyperlinks
    automizer.addSlide('template', 1, async (slide) => {
      for (let i = 0; i < externalUrls.length; i++) {
        slide.generate(async (pptxGenJSSlide) => {
          pptxGenJSSlide.addText(`External Link ${i + 1}`, {
            hyperlink: { url: externalUrls[i] },
            x: 1,
            y: 1 + i * 0.5,
            w: 2.5, h: 0.5,
            fontFace: 'Kanit',
            fontSize: 12,
          });
        });
      }

      for (let i = 0; i < targetInternalSlides.length; i++) {
        slide.generate(async (pptxGenJSSlide) => {
          pptxGenJSSlide.addText(`Go to Slide ${targetInternalSlides[i]}`, {
            hyperlink: { slide: targetInternalSlides[i] },
            x: 4,
            y: 1 + i * 0.5,
            w: 2.5, h: 0.5,
            fontFace: 'Kanit',
            fontSize: 12,
          });
        });
      }
    });

    for (let i = 0; i < targetInternalSlides.length; i++) {
      // addSlide will create slides 2, 3, 4
      automizer.addSlide('template', 1, async (slide) => {
        slide.generate(async (pptxGenJSSlide) => {
          pptxGenJSSlide.addText(`This is Slide ${targetInternalSlides[i]}`, {
            x: 1, y: 1, w: 5, h: 0.5,
            fontFace: 'Arial',
            fontSize: 18,
          });
        });
      });
    }

    // Write the presentation
    const summary = await automizer.write(outputFile);
    expect(summary.slides).toBe(1 + targetInternalSlides.length); // 1 (links) + 3 (targets) = 4 slides

    // Verify the generated PPTX
    const fileData = fs.readFileSync(outputPath);
    const zip = await JSZip.loadAsync(fileData);
    const parser = new DOMParser();

    // Slide 1 (index 0) is where links were generated.
    const slide1RelsPath = 'ppt/slides/_rels/slide1.xml.rels';
    const slide1RelsFile = zip.file(slide1RelsPath);
    expect(slide1RelsFile).not.toBeNull();
    const slide1RelsXml = await slide1RelsFile?.async('text');
    const slide1RelsDoc = parser.parseFromString(slide1RelsXml, 'application/xml');
    const slide1Relationships = slide1RelsDoc.getElementsByTagName('Relationship');

    const foundExternalRels = new Map<string, string>();
    const foundInternalRels = new Map<number, string>();

    for (let i = 0; i < slide1Relationships.length; i++) {
      const rel = slide1Relationships[i];
      const relType = rel.getAttribute('Type');
      const relTarget = rel.getAttribute('Target');
      const relId = rel.getAttribute('Id');

      if (relType === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
        expect(rel.getAttribute('TargetMode')).toBe('External');
        if (relTarget && relId) foundExternalRels.set(relTarget, relId);
      } else if (relType === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide') {
        const slideNumMatch = relTarget?.match(/slide(\d+)\.xml$/);
        if (slideNumMatch && relId) {
          foundInternalRels.set(parseInt(slideNumMatch[1]), relId);
        }
      }
    }

    // Check external links in rels
    expect(foundExternalRels.size).toBe(externalUrls.length);
    for (const url of externalUrls) {
      expect(foundExternalRels.has(url)).toBe(true);
    }

    // Check internal links in rels
    expect(foundInternalRels.size).toBe(targetInternalSlides.length);
    for (const slideNum of targetInternalSlides) {
      expect(foundInternalRels.has(slideNum)).toBe(true);
    }

    // Now check slide1.xml content
    const slide1Path = 'ppt/slides/slide1.xml';
    const slide1File = zip.file(slide1Path);
    expect(slide1File).not.toBeNull();
    const slide1Xml = await slide1File?.async('text');
    const slide1Doc = parser.parseFromString(slide1Xml, 'application/xml');

    const shapes = slide1Doc.getElementsByTagName('p:sp');
    const textToExpectedRid = new Map<string, string>();

    // Populate map with expected text content and their corresponding rIds from the rels file
    externalUrls.forEach((url, index) => {
      const text = `External Link ${index + 1}`;
      const rid = foundExternalRels.get(url);
      if (rid) textToExpectedRid.set(text, rid);
    });
    targetInternalSlides.forEach((slideNum, index) => {
      const text = `Go to Slide ${slideNum}`;
      const rid = foundInternalRels.get(slideNum);
      if (rid) textToExpectedRid.set(text, rid);
    });

    let foundLinksInShapes = 0;
    for (let i = 0; i < shapes.length; i++) {
      const shape = shapes[i];
      const textElements = shape.getElementsByTagName('a:t');
      if (textElements.length > 0) {
        const shapeText = Array.from(textElements).map(t => t.textContent).join('');
        const expectedRid = textToExpectedRid.get(shapeText);

        if (expectedRid) {
          foundLinksInShapes++;
          // Check shape-level hyperlink
          const cNvPr = shape.getElementsByTagName('p:cNvPr')[0];
          if (cNvPr) {
            const shapeHlink = cNvPr.getElementsByTagName('a:hlinkClick')[0];
            expect(shapeHlink).not.toBeNull();
            expect(shapeHlink?.getAttribute('r:id')).toBe(expectedRid);
          }

          // Check text-run-level hyperlink
          const rPrs = shape.getElementsByTagName('a:rPr');
          let textRunHlinkFound = false;
          for (let j = 0; j < rPrs.length; j++) {
            const textHlink = rPrs[j].getElementsByTagName('a:hlinkClick')[0];
            if (textHlink) {
              expect(textHlink.getAttribute('r:id')).toBe(expectedRid);
              textRunHlinkFound = true;
              break;
            }
          }
          expect(textRunHlinkFound).toBe(true);
        }
      }
    }
    expect(foundLinksInShapes).toBe(externalUrls.length + targetInternalSlides.length);
  });
});
