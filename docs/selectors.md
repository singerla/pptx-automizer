---
title: Selecting slides and shapes
description: Address slides by number or creationId, and shapes by name, nameIdx or creationId.
---

`pptx-automizer` needs a selector to find the required shape on a template slide. While an imported .pptx file is identified by filename or custom label, there are different ways to address its slides and shapes.

## Select slide by number and shape by name

If your .pptx-templates are more or less static and you do not expect them to evolve a lot, it's ok to use the slide number and the shape name to find the proper source of automation.

```ts
// This will take slide #2 from 'SlideWithGraph.pptx' and expect it
// to contain a shape called 'ColumnChart':
pres.addSlide('SlideWithGraph.pptx', 2, (slide) => {
  // `slide` is slide #2 of 'SlideWithGraph.pptx'
  slide.modifyElement('ColumnChart', [
    /* ... */
  ]);
});

// This example will take slide #1 from 'RootTemplate.pptx' and place
// 'ColumnChart' from slide #2 of 'SlideWithGraph.pptx' on it.
pres.addSlide('RootTemplate.pptx', 1, (slide) => {
  // `slide` is slide #1 of 'RootTemplate.pptx'
  slide.addElement('SlideWithGraph.pptx', 2, 'ColumnChart', [
    /* ... */
  ]);
});

// In case you have two or more shapes on your template slide with the same name,
// you can address the target by nameIdx:
pres.addSlide('SlideWithGraph.pptx', 1, (slide) => {
  // Starting from 0, this will find the second shape named 'ColumnChart'
  // on the slide.
  slide.modifyElement(
    {
      name: 'ColumnChart',
      nameIdx: 1,
    },
    [
      /* ... */
    ],
  );
});
```

> You can display and manage shape names directly in PowerPoint by opening the "Selection"-pane for your current slide. Hit `ALT+F10` and PowerPoint will give you a (nested) list including all (grouped) shapes. You can edit a shape name by double-click or by hitting `F2` after selecting a shape from the list. [See MS-docs for more info.](https://support.microsoft.com/en-us/office/use-the-selection-pane-to-manage-objects-in-documents-a6b2fd3e-d769-46c1-9b9c-b94e04a72550)

But be aware: Whenever your template slides are rearranged or a template shape is renamed, you need to update your code as well.

Please also make sure that each shape to add or modify has a unique name on its slide. Otherwise, only the last matching shape will be taken as target.

## Select slides by creationId

Additionally, each slide and shape is stored together with a (more or less) unique `creationId`. In XML, it looks like this:

```xml
<p:cNvPr name="MyPicture" id="64">
    <a:extLst>
        <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
            <a16:creationId id="{0980FF19-E7E7-493C-8D3E-15B2100EA940}" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"/>
        </a:ext>
    </a:extLst>
</p:cNvPr>
```

This is where `name` and `creationId` are coupled together for each shape.

While our shape could now be identified by both, `MyPicture` or by `{0980FF19-E7E7-493C-8D3E-15B2100EA940}`, `creationIds` for slides consist of an integer value, e.g. `501735012` below:

```xml
<p:extLst>
   <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
      <p14:creationId val="501735012" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>
   </p:ext>
</p:extLst>
```

You can add a simple code snippet to get a list of the available `creationIds` of your loaded templates:

```ts
const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithShapes.pptx`, 'shapes');

const creationIds = await pres.setCreationIds();

// This is going to print the slide creationId and a list of all
// shapes from slide #1 in `SlideWithShapes.pptx` (aka `shapes`).
console.log(
  creationIds
    .find((template) => template.name === 'shapes')
    .slides.find((slide) => slide.number === 1),
);
// Find the corresponding slide-creationId and -number on top of this list.
```

If your templates are not final and if you expect to have new slides and shapes added in the future, it is worth the effort and use `creationId` in general:

```ts
const automizer = new Automizer({
  templateDir: `${__dirname}/pptx-templates`,
  outputDir: `${__dirname}/pptx-output`,
  // turn this to true and use creationIds for both, slides and shapes
  useCreationIds: true,
});
```

Regarding shapes, it is also possible to use a `creationId` and the shape name as a fallback. These are the different types of a `FindElementSelector`:

```ts
import { FindElementSelector } from 'pptx-automizer';

// This is default when set up with `useCreationIds: true`:
const myShapeSelectorCreationId: FindElementSelector =
  '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}';

// pptx-generator will try to find the shape even if one of the given keys
// won't match any shape on the target slide:
const myShapeSelectorFallback: FindElementSelector = {
  creationId: '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}',
  name: 'Drum',
};

// Use this only if `useCreationIds: false`:
const myShapeSelectorName: FindElementSelector = 'Drum';

// Whenever `useCreationIds` was set to true, you need to replace slide numbers
// by `creationId`, too:
await pres.addSlide('shapes', 4167997312, (slide) => {
  // slide is now #1 of `SlideWithShapes.pptx`
  slide.addElement('shapes', 273148976, {
    creationId: '{E43D12C3-AD5A-4317-BC00-FDED287C0BE8}',
    name: 'Drum',
  });
  // 'Drum' is from #2 of `SlideWithShapes.pptx`, see __tests__ dir for an
  // example.
});
```

If you decide to use the `creationId` method, you are safe to add, remove and rearrange slides in your templates. It is also no problem to update shape names, and you also don't need to pay attention to unique shape names per slide.

> Please note: PowerPoint is going to update a shape's `creationId` only in case the shape was copied & pasted on a slide with an already existing identical shape `creationId`. If you were copying a slide, each shape `creationId` will be copied, too. As a result, you have unique shape ids, but different slide `creationIds`. If you are now going to paste a shape an such a slide, a new creationId will be given to the pasted shape. As a result, slide ids are unique throughout a presentation, but shape ids are unique only on one slide.

## Find and Modify Shapes

There are basically two ways to access a target shape on a slide:

- `slide.modifyElement(...)` requires an existing shape on the current slide,
- `slide.addElement(...)` adds a shape from another slide to the current slide.

Modifications can be applied to both in the same way:

```ts
import { modify, CmToDxa } from 'pptx-automizer';

pres.addSlide('shapes', 2, (slide) => {
  // This will only work if there is a shape called 'Drum'
  // on slide #2 of the template labelled 'shapes'.
  slide.modifyElement('Drum', [
    // You can use some of the builtin modifiers to edit a shape's xml:
    modify.setPosition({
      // set position from the left to 5 cm
      x: CmToDxa(5),
      // or use a number in DXA unit
      h: 5000000,
      w: 5000000,
    }),
    // Log your target xml into the console:
    modify.dump,
  ]);
});

pres.addSlide('shapes', 1, (slide) => {
  // This will import the 'Drum' shape from
  // slide #2 of the template labelled 'shapes'.
  slide.addElement('shapes', 2, 'Drum', [
    // add modifiers as seen in the example above
  ]);
});
```

## Inspect a slide from inside the callback

Besides `pres.getInfo()` (see the [Basic Example](./getting-started.md#basic-example)),
an added slide can describe itself at `write()` time. Prefer inspecting over
guessing shape names — an unresolvable selector rejects `write()` (or is
skipped with `continueOnError: true`), so listing what is actually there beats
retrying name variations:

```ts
pres.addSlide('myTemplate.pptx', 1, async (slide) => {
  // All shapes on this slide, with name, id, position and type:
  const elements = await slide.getAllElements();
  // Only certain tags, e.g. text-bearing shapes:
  const shapes = await slide.getAllElements(['sp']);
  // The slide's dimensions:
  const dimensions = await slide.getDimensions();
});
```

## Find all text elements on a slide

When processing an added slide, you might want to apply a modifier to any existing text element. Call `slide.getAllTextElementIds()` for this:

```ts
import Automizer, { modify } from 'pptx-automizer';

pres.addSlide('myTemplate.pptx', 1, async (slide) => {
  const elements = await slide.getAllTextElementIds();
  elements.forEach((element) => {
    // element has a text body:
    slide.modifyElement(element, [modify.setText('my text')]);
    // ... or use the tag replace function:
    slide.modifyElement(element, [
      modify.replaceText([
        {
          replace: 'TAG',
          by: {
            text: 'my tag text',
          },
        },
      ]),
    ]);
  });
});
```
