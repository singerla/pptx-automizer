import Automizer from '../src/automizer';
import { ChartData, modify } from '../src';

test('insert a table with pptxgenjs on a template slide', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  pres.addSlide('empty', 1, (slide) => {
    // Use pptxgenjs to add a table from scratch:
    slide.generate((pptxGenJSSlide) => {
      const rowsTest = [
        [
          { text: "Top Lft", options: { hyperlink: { url: 'https://duckduckgo.com' } }, },
          { text: "Top Ctr",  },
          { text: "Top Rgt",  },
        ],
        [
          { text: "Bot Lft", },
          { text: "Bot Ctr",  },
          { text: "Bot Rgt",  },
        ],
      ];

      pptxGenJSSlide.addTable(rowsTest, { w: 9, rowH: 2, align: "left", fontFace: "Arial" });
    }, 'custom object name');
  });

  const result = await pres.write(`generate-pptxgenjs-table.test.pptx`);

  expect(result.slides).toBe(2);
});
