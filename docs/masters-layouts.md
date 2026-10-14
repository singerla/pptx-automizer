---
title: Slide Masters and Layouts
description: Import slide masters with their layouts, modify master shapes, and switch layouts on output slides.
---

`pptx-automizer` supports importing slide masters and their associated slide layouts into the output presentation, modifying the masters and layouts that are already in the root template in place, and removing a master together with its layouts. Individual layouts cannot be imported on their own — a layout always comes with its master. A modification on a slideMaster appears on all related slideLayouts.

Please be aware that importing slide layouts containing complex content, such as charts and images, is currently not supported. For instance, if a slide layout includes an icon that is not present on the slide master, this icon will break when the slide master is auto-imported into an output presentation. To avoid this issue, ensure that all images and charts are placed exclusively on a slide master and not on a slide layout.

## Import and modify slide Masters

You can import, modify and use one or more slideMasters and the related slideLayouts.

To specify the target index of the required slide master to import, you need to count slideMasters in your _template_ presentation.
To specify another slideLayout for an added output slide, you need to count slideLayouts in your _output_ presentation.

To add and modify shapes on a slide master, please take a look at [Find and Modify Shapes](./selectors.md#find-and-modify-shapes) — masters use the same selectors and modifiers as slides.

```ts
// Import another slide master and all its slide layouts.
// Index 1 means, you want to import the first of all masters:
pres.addMaster('SlidesWithAdditionalMaster.pptx', 1, (master) => {
  // Modify a certain shape on the slide master:
  master.modifyElement(
    `MasterRectangle`,
    ModifyTextHelper.setText('my text on master'),
  );
  // Add a shape from an imported templated to the current slideMaster.
  master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
});
```

Any imported slideMaster will be appended to the existing ones in the root template. If you have already e.g. one master with five layouts, and you import a new master coming with seven slide layouts, the first new layout will be #6.

```ts
// Import a slideMaster and its slideLayouts:
pres.addMaster('SlidesWithAdditionalMaster.pptx', 1);

// Add a slide and switch to another layout:
pres.addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
  // use another master, e.g. the imported one from 'SlidesWithAdditionalMaster.pptx'
  // You need to pass the index of the desired layout after all
  // related layouts of all imported masters have been added to rootTemplate.
  slide.useSlideLayout(12);
});

// It is also possible to use the original slideLayout of any added slide:
pres.addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
  // To use the original master from 'SlidesWithAdditionalMaster.pptx',
  // we can skip the argument:
  slide.useSlideLayout();
  // This will also auto-import the original slideMaster, if not done already,
  // and look for the created index of the source slideLayout.
});
```

## Modify masters and layouts of the root template in place

Text, logos or other shapes that live on a master of the root template can
be edited without importing a copy: `pres.modifyMaster(n, callback)` applies
the same `modifyElement` / `addElement` / `removeElement` vocabulary as
`addMaster` to the root's own master. The master keeps its number, no part is
added, and every slide using it keeps using it. `pres.modifyLayout(n,
callback)` does the same for a single slide layout.

```ts
pres
  .loadRoot('SlidesWithAdditionalMaster.pptx')
  .load('SlideWithShapes.pptx')
  // the first master of the root template, as loaded:
  .modifyMaster(1, (master) => {
    master.modifyElement('MasterRectangle', ModifyTextHelper.setText('Customer name'));
    master.addElement('SlideWithShapes.pptx', 1, 'Cloud 1');
    master.removeElement('Rechteck 6');
  })
  // the first layout of the root template, as loaded:
  .modifyLayout(1, (layout) => {
    layout.modifyElement('Titel 1', ModifyTextHelper.setText('Title placeholder'));
  });
```

- `n` is the 1-based number of `ppt/slideMasters/slideMaster<n>.xml` /
  `ppt/slideLayouts/slideLayout<n>.xml` in the root template as loaded — the
  order of PowerPoint's Slide Master view. Masters imported with `addMaster`
  are modified through their own callback instead.
- Deferred like everything else: the modifications run on `write()` before
  the slides are written, in call order with `addMaster` imports. Several
  `modifyMaster` calls for the same number share one queue.
- A number that does not exist in the root template rejects `write()`.
- An image replaced on a master (`ModifyImageHelper.setRelationTarget`) is
  copied into the archive like on a slide.

## Remove a slide master

`pres.removeMaster(n)` removes master `n` of the root template with its
layouts, its theme (unless another part — usually the presentation itself,
for the first theme — still refers to it), their relationships and their
entries in `ppt/presentation.xml` and `[Content_Types].xml`:

```ts
pres
  .loadRoot('SlidesWithAdditionalMaster.pptx')
  .load('SlideWithShapes.pptx')
  .addSlide('SlideWithShapes.pptx', 1, (slide) => {
    // this slide used a layout of master 1: switch it to master 2 first
    slide.useSlideLayout(13);
  })
  .removeMaster(1);
```

- Runs on `write()` after all slides have been written, so layout switches
  are in effect. If a slide of the output still uses one of the master's
  layouts, `write()` rejects and names the slide and layout; point the
  slide at another layout with `useSlideLayout()` first, or drop the root
  template's own slides with `removeExistingSlides: true`.
- The remaining masters and layouts keep their file numbers; the master and
  layout ids in `presentation.xml` are renumbered as usual (starting at
  2147483648, the smallest id PowerPoint accepts).
- Media used only by the removed master and layouts stays in the archive
  unless `cleanup: true` is set.

Please notice: If your root template and your imported slides have an equal structure of slideMasters and slideLayouts, it won't be necessary to add slideMasters manually.

If you have trouble with messed up slideMasters, and if you don't worry about the impact on performance, you can try and set `autoImportSlideMasters: true` to always import all required files:

```ts
import Automizer from 'pptx-automizer';

const automizer = new Automizer({
  // ...

  // Always use the original slideMaster and slideLayout of any
  // imported slide:
  autoImportSlideMasters: true,
  // ...
});
```
