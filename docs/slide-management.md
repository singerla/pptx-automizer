---
title: Slide Management
description: Remove elements, sort output slides, loop through all slides of a presentation, and read slide numbers.
---

## Remove elements from a slide

You can as well remove elements from slides.

```ts
// Remove existing charts, images or shapes from added slide.
pres
  .addSlide('charts', 2, (slide) => {
    slide.removeElement('ColumnChart');
  })
  .addSlide('images', 2, (slide) => {
    slide.removeElement('imageJPG');
    slide.removeElement('Textfeld 5');
    slide.addElement('images', 2, 'imageJPG');
  });
```

## Loop through the slides of a presentation

If you would like to modify elements in a single .pptx file, it is important to know that `pptx-automizer` is not able to directly "jump" to a shape to modify it — see [the template/root model](./concepts.md#the-templateroot-model) for how it works internally.

In case you need to apply modifications to the root template, you need to load it as a normal template:

```ts
import Automizer, {
  CmToDxa,
  ISlide,
  ModifyColorHelper,
  ModifyShapeHelper,
  ModifyTextHelper,
} from 'pptx-automizer';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `path/to/pptx-templates`,
    outputDir: `path/to/pptx-output`,
    // this is required to start with no slides:
    removeExistingSlides: true,
  });

  let pres = automizer
    .loadRoot(`SlideWithShapes.pptx`)
    // We load it twice to make it available for modifying slides.
    // Defining a "name" as the second parameter makes it a little easier
    .load(`SlideWithShapes.pptx`, 'myTemplate');

  // This is brand new: get useful information about loaded templates:
  const myTemplates = await pres.getInfo();
  const mySlides = myTemplates.slidesByTemplate(`myTemplate`);

  // Feel free to create some functions to pre-define all modifications
  // you need to apply to your slides.
  type CallbackBySlideNumber = {
    slideNumber: number;
    callback: (slide: ISlide) => void;
  };
  const callbacks: CallbackBySlideNumber[] = [
    {
      slideNumber: 2,
      callback: (slide: ISlide) => {
        slide.modifyElement('Cloud', [
          ModifyTextHelper.setText('My content'),
          ModifyShapeHelper.setPosition({
            h: CmToDxa(5),
          }),
          ModifyColorHelper.solidFill({
            type: 'srgbClr',
            value: 'cccccc',
          }),
        ]);
      },
    },
  ];
  const getCallbacks = (slideNumber: number) => {
    return callbacks.find((callback) => callback.slideNumber === slideNumber)
      ?.callback;
  };

  // We can loop all slides and apply the callbacks if defined
  mySlides.forEach((mySlide) => {
    pres.addSlide('myTemplate', mySlide.number, getCallbacks(mySlide.number));
  });

  // This will result in an output presentation containing all slides of "SlideWithShapes.pptx"
  pres.write(`myOutputPresentation.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
```

## Quickly get all slide numbers of a template

When calling `pres.getInfo()`, it will gather information about all elements on all slides of all templates. In case you just want to loop through all slides of a certain template, you can use this shortcut:

```ts
const slideNumbers = await pres
  .getTemplate('myTemplate.pptx')
  .getAllSlideNumbers();

for (const slideNumber of slideNumbers) {
  // do the thing
}
```

## Sort output slides

There are three ways to arrange slides in an output presentation.

1. By default, all slides will be appended to the existing slides in your root template. The order of `addSlide` calls will define slide sorting in the output presentation.

2. You can alternatively remove all existing slides by setting the `removeExistingSlides` flag to true. The first slide added with `addSlide` will be first slide in the output presentation. If you want to insert slides from root template, you need to load it a second time.

```ts
import Automizer from 'pptx-automizer';

const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // truncate root presentation and start with zero slides
  removeExistingSlides: true,
});

let pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // We load this twice to make it available for sorting slides
  .load(`RootTemplate.pptx`, 'root')
  .load(`SlideWithShapes.pptx`, 'shapes')
  .load(`SlideWithGraph.pptx`, 'graph');

pres
  .addSlide('root', 1) // First slide will be taken from root
  .addSlide('graph', 1)
  .addSlide('shapes', 1)
  .addSlide('root', 3) // Third slide from root will be appended
  .addSlide('root', 2); // Second and third slide will switch position

pres.write(`mySortedPresentation.pptx`).then((summary) => {
  console.log(summary);
});
```

3. Use `sortSlides`-callback
   You can pass an array of numbers and create a callback and apply it to `presentation.xml`.
   This will also work without adding slides.

Slides will be appended to the existing slides by slide number (starting from 1). You may find irritating results in case you skip a slide number.

```ts
import { ModifyPresentationHelper } from 'pptx-automizer';

//
// You may truncate root template or you may not
// ...

// It is possible to skip adding slides, try sorting an unmodified presentation
pres
  .addSlide('charts', 1)
  .addSlide('charts', 2)
  .addSlide('images', 1)
  .addSlide('images', 2);

const order = [3, 2, 4, 1];
pres.modify(ModifyPresentationHelper.sortSlides(order));
```
