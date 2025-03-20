import { Automizer } from '../src';
import fs from 'fs';
import path from 'path';

describe('Hyperlink works with generate', () => {
  it('Hyperlink works with generate', async () => {
    const automizer = new Automizer({
      templateDir: path.join(__dirname, 'pptx-templates'),
      outputDir: path.join(__dirname, 'pptx-output')
    });

    // Load root and template files
    await automizer.loadRoot('EmptyTemplate.pptx');
    await automizer.load('EmptySlide.pptx', 'template');

    // Create slide 1 with hyperlinks
    await automizer.addSlide('template', 1, async (slide) => {
      // Test external links
      const externalUrls = [
        'https://google.com/1',
        'https://google.com/2',
        'https://google.com/3',
      ];

      for (let i = 0; i < externalUrls.length; i++) {
        slide.generate(async pptxGenJSSlide => {
          // External link
          pptxGenJSSlide.addText(`External Link ${i + 1}`, {
            hyperlink: { url: externalUrls[i] },
            x: 1,
            y: 1 + i,
            fontFace: 'Kanit',
          });
        });
      }

      // Test internal links
      const targetSlides = [2, 3, 4];
      for (let i = 0; i < targetSlides.length; i++) {
        slide.generate(async pptxGenJSSlide => {
          // Internal link
          pptxGenJSSlide.addText(`Go to Slide ${targetSlides[i]}`, {
            hyperlink: { slide: targetSlides[i] },
            x: 3,
            y: 1 + i,
            fontFace: 'Kanit',
          });
        });
      }
    });
    
    // Create target slides that we'll link to
    for (let i = 2; i <= 4; i++) {
      await automizer.addSlide('template', 1, async (slide) => {
        slide.generate(async pptxGenJSSlide => {
          pptxGenJSSlide.addText(`This is Slide ${i}`, {
            x: 1,
            y: 1,
            fontFace: 'Kanit',
          });
        });
      });
    }

    // Write the presentation
    await automizer.write('hyperlink-test.pptx');
  });
}); 