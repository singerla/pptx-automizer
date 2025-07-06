import { Automizer } from '../src';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

describe('addTable with hyperlinks', () => {
  const outputDir = path.join(__dirname, 'pptx-output');
  const outputFile = 'addTable-hyperlinks.test.pptx';
  const outputPath = path.join(outputDir, outputFile);

  beforeAll(() => {
    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }
  });

  it('preserves different hyperlink URLs in table cells', async () => {
    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'pptx-templates'),
      outputDir: outputDir,
    });

    automizer.loadRoot('EmptyTemplate.pptx');
    automizer.load('EmptySlide.pptx', 'template');

    automizer.addSlide('template', 1, (slide) => {
      slide.generate((pptxGenJSSlide) => {
        const tableData = [
          [
            {
              text: "Google",
              options: { hyperlink: { url: "https://google.com" } },
            },
            {
              text: "DuckDuckGo", 
              options: { hyperlink: { url: "https://duckduckgo.com" } },
            },
            { text: "No Link" },
          ],
        ];

        pptxGenJSSlide.addTable(tableData, {
          w: 9,
          rowH: 2,
          align: "left",
          fontFace: "Arial",
        });
      });
    });

    const summary = await automizer.write(outputFile);
    expect(summary.slides).toBe(1);

    const fileData = fs.readFileSync(outputPath);
    const zip = await JSZip.loadAsync(fileData);
    const parser = new DOMParser();

    const slide1RelsPath = 'ppt/slides/_rels/slide1.xml.rels';
    const slide1RelsFile = zip.file(slide1RelsPath);
    expect(slide1RelsFile).not.toBeNull();
    
    const slide1RelsXml = await slide1RelsFile!.async('text');
    const slide1RelsDoc = parser.parseFromString(slide1RelsXml, 'application/xml');
    const relationships = slide1RelsDoc.getElementsByTagName('Relationship');

    const hyperlinkRels = Array.from(relationships)
      .filter(rel => rel.getAttribute('Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink')
      .map(rel => ({
        id: rel.getAttribute('Id'),
        target: rel.getAttribute('Target'),
        targetMode: rel.getAttribute('TargetMode')
      }));

    expect(hyperlinkRels.length).toBe(2);
    
    const targets = hyperlinkRels.map(rel => rel.target);
    expect(targets).toContain('https://google.com');
    expect(targets).toContain('https://duckduckgo.com');

    hyperlinkRels.forEach(rel => {
      expect(rel.targetMode).toBe('External');
    });

    const slide1Path = 'ppt/slides/slide1.xml';
    const slide1File = zip.file(slide1Path);
    expect(slide1File).not.toBeNull();
    
    const slide1Xml = await slide1File!.async('text');
    const slide1Doc = parser.parseFromString(slide1Xml, 'application/xml');
    
    const hlinkClicks = slide1Doc.getElementsByTagName('a:hlinkClick');
    expect(hlinkClicks.length).toBe(2);
    
    const rIds = Array.from(hlinkClicks).map(hlink => hlink.getAttribute('r:id'));
    expect(new Set(rIds).size).toBe(rIds.length);
    
    const relationshipIds = hyperlinkRels.map(rel => rel.id);
    rIds.forEach(rId => {
      expect(relationshipIds).toContain(rId);
    });
  });
}); 