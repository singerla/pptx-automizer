---
title: Slide Masters and Layouts
description: Import slide masters with their layouts, modify master shapes, and switch layouts on output slides.
---

`pptx-automizer` supports importing slide masters and their associated slide layouts into the output presentation. It is important to note that you cannot add, modify, or remove individual slide layouts directly. However, you have the flexibility to modify the underlying slide master, which can serve as a workaround for certain changes — each modification on a slideMaster will appear on all related slideLayouts.

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

## Select a layout by name

Instead of counting layouts in the output, you can address a layout by its
name — the name PowerPoint shows in the layout gallery and in "Rename Layout"
— and, if needed, by the name of its master and by the template it came
from. Pass an object to `useSlideLayout`:

```ts
pres.addSlide('SlideWithShapes.pptx', 1, (slide) => {
  // A layout of the root template, by name. Nothing is imported.
  slide.useSlideLayout({ layout: 'Nur Titel', template: 'root' });
});

pres
  .addMaster('SlidesWithAdditionalMaster.pptx', 1)
  .addMaster('SlidesWithAdditionalMaster.pptx', 2)
  .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
    // Both imported masters (and the root master) have a layout named
    // 'Titel und Inhalt': name the master to pick one.
    slide.useSlideLayout({ master: 'Orange Design', layout: 'Titel und Inhalt' });
  })
  .addSlide('SlidesWithAdditionalMaster.pptx', 3, (slide) => {
    // Only layouts imported from a given template:
    slide.useSlideLayout({
      layout: 'Nur Titel',
      template: 'SlidesWithAdditionalMaster.pptx',
    });
  });
```

How the selector (`SlideLayoutSelector`) is resolved:

- `layout` is compared exactly (including line breaks) with the layout's
  name in the output presentation.
- `master` is the master's name as PowerPoint shows it in the layout
  gallery: the name of the master's theme, or the master's own name if it
  has one. Renaming a master in PowerPoint's Slide Master view renames its
  theme.
- `template` restricts the lookup: `'root'` means the layouts of the root
  template as loaded; the alias of a loaded template means the layouts that
  were imported from it (with `addMaster` or by auto-import). Without
  `template`, every master in the output is searched — root and imported.
- The selector is resolved against the output at write time, after all
  masters have been imported; it never imports a master itself. Import the
  masters you need first (or use the no-argument form, which imports the
  source slide's own master).
- No matching layout, or more than one, rejects `write()` with a
  `SlideLayoutNotFoundError` whose message lists the available (or the
  ambiguous) layouts as `"<master> / <layout>"`. With
  `continueOnError: true` the error is logged and the slide keeps its
  current layout.

A plain string still works as before — `slide.useSlideLayout('Nur Titel')`
imports the source slide's master and looks the name up among the imported
layouts — and now falls back to the root template's own layouts when the
name is not found there. If the name is ambiguous, the first match wins and
a warning names the candidates; use the object form to pick one.

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
