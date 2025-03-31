import Automizer from './index';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir,
  });

  let pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'emptySlide');

  const object = {
    position: {
      x: 500000,
      y: 500000,
      cx: 500000,
      cy: 500000,
    },
  };

  pres.addSlide('emptySlide', 1, async (slide) => {
    slide.generate(async (pptxGenJSSlide) => {
      pptxGenJSSlide.addText('hello world1', {
        hyperlink: {
          url: 'https://duckduckgo.com',
        },
        x: object.position.x,
        y: object.position.y,
        w: object.position.cx,
        h: object.position.cy,
        fontFace: 'Kanit',
      });
    });
    slide.generate(async (pptxGenJSSlide) => {
      pptxGenJSSlide.addText('hello world2', {
        hyperlink: {
          url: 'https://duckduckgo.com',
        },
        x: object.position.x,
        y: object.position.y,
        w: object.position.cx,
        h: object.position.cy,
        fontFace: 'Kanit',
      });
    });
    slide.generate(async (pptxGenJSSlide) => {
      pptxGenJSSlide.addText('hello world3', {
        hyperlink: {
          url: 'https://duckduckgo.com',
        },
        x: object.position.x,
        y: object.position.y,
        w: object.position.cx,
        h: object.position.cy,
        fontFace: 'Kanit',
      });
    });
    slide.generate(async (pptxGenJSSlide) => {
      pptxGenJSSlide.addText('hello world5', {
        hyperlink: {
          url: 'https://duckduckgo.com',
        },
        x: object.position.x,
        y: object.position.y,
        w: object.position.cx,
        h: object.position.cy,
        fontFace: 'Kanit',
      });
    });

    slide.generate(async (pptxGenJSSlide) => {
      pptxGenJSSlide.addText('hello world6', {
        hyperlink: {
          url: 'https://duckduckgo.com',
        },
        x: object.position.x,
        y: object.position.y,
        w: object.position.cx,
        h: object.position.cy,
        fontFace: 'Kanit',
      });
    });
  });

  pres.write(`testHyperlinks.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
