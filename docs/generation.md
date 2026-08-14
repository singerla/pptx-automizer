---
title: Generate shapes with PptxGenJS
description: Create text, charts, images, tables and hyperlinked shapes from scratch via the PptxGenJS wrapper.
---

This library wraps around [PptxGenJS](https://github.com/gitbrent/PptxGenJS) to generate shapes from scratch. It is possible to use the `PptxGenJS` wrapper to generate shapes on a slide.

Here's an example of how to use `pptxGenJS` to add a text shape to a slide:

```ts
pres.addSlide('empty', 1, (slide) => {
  // Use PptxGenJS to add text from scratch:
  slide.generate((pptxGenJSSlide) => {
    pptxGenJSSlide.addText('Test 1', {
      x: 1,
      y: 1,
      h: 5,
      w: 10,
      color: '363636',
    });
  }, 'custom object name');
});
```

You can as well create charts with `pptxGenJS`:

```ts
const dataChartAreaLine = [
  {
    name: 'Actual Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1500, 4600, 5156],
  },
  {
    name: 'Projected Sales',
    labels: ['Jan', 'Feb', 'Mar'],
    values: [1000, 2600, 3456],
  },
];

pres.addSlide('empty', 1, (slide) => {
  // Use PptxGenJS to add generated content from scratch:
  slide.generate((pSlide, pptxGenJs) => {
    pSlide.addChart(pptxGenJs.ChartType.line, dataChartAreaLine, {
      x: 1,
      y: 1,
      w: 8,
      h: 4,
    });
  });
});
```

Note: inside `generate`, coordinates and sizes are in **inches** (PptxGenJS
convention) — unlike `modify.setPosition`, which uses EMU/DXA values (see the
[units reference](./helpers.md#units-reference)).

You can use the following functions to generate shapes with `pptxGenJS`:

- addChart
- addImage
- addShape
- addTable
- addText

## Create a new hyperlinked text shape

It is also possible to create a new hyperlink from scratch with the `pptxGenJS` wrapper. This is useful if you want to add hyperlinks to shapes that are not part of the template. (To manage hyperlinks on *existing* template shapes, see [Hyperlink Management](./hyperlinks.md).)

```ts
// Generate a new text shape pointing to an external site
slide.generate((pptxGenJSSlide) => {
  pptxGenJSSlide.addText(`External Link`, {
    hyperlink: { url: 'https://github.com' },
    x: 1,
    y: 1,
    w: 2.5,
    h: 0.5,
    fontSize: 12,
  });
});

// Or generate an internal hyperlink
slide.generate((pptxGenJSSlide) => {
  pptxGenJSSlide.addText(`Go to slide 3`, {
    hyperlink: { slide: 3 },
    x: 1,
    y: 1,
    w: 2.5,
    h: 0.5,
    fontSize: 12,
  });
});
```

See a complete example on [how to add a chart from scratch](https://github.com/singerla/pptx-automizer/blob/main/__tests__/generate-pptxgenjs-charts.test.ts).
