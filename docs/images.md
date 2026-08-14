---
title: Modify Images
description: Import external media, swap image sources, auto-crop to cover, and apply duotone overlays.
---

`pptx-automizer` can extract images from loaded .pptx template files and add to your output presentation. You can use shape modifiers (e.g. for size and position) on images, too. Additionally, it is possible to load external media files directly and update relation `Target` of an existing image. This works on both, existing or added images.

```ts
const automizer = new Automizer({
  // ...
  // Specify a directory to import external media files from:
  mediaDir: `path/to/media`,
});

const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  // load one or more files from mediaDir
  .loadMedia([`feather.png`, `test.png`] /* or use a custom dir */)
  // and/or use a custom dir
  .loadMedia(`icon.png`, 'path/to/icons')
  .load(`SlideWithImages.pptx`, 'images');

pres.addSlide('images', 2, (slide) => {
  slide.modifyElement('imagePNG', [
    // Override the original media source of element 'imagePNG'
    // by an imported file:
    ModifyImageHelper.setRelationTarget('feather.png'),

    // You might need to update size
    ModifyShapeHelper.setPosition({
      w: CmToDxa(5),
      h: CmToDxa(3),
    }),
  ]);
});
```

Media can also come from memory instead of `mediaDir`:
`automizer.loadMediaBuffer('chart.png', myImageBuffer)` registers a `Buffer`
under a filename, usable with `setRelationTarget` like any loaded file.

Note: media loading requires the root template to be loaded first. An image that is imported more than once (e.g. a logo on several slides) is stored only once in the output file — identical media files are always shared.

This will also auto-crop the image to the new width and height,
based on the container aspect ratio, derived from the original image
and using the new image width and height based on the files loaded into
the presentation media folder using .loadMedia():

```ts
pres.addSlide('images', 2, (slide) => {
  slide.modifyElement('Image Placeholder', [
    ModifyImageHelper.setRelationTargetCover('feather.png', pres),
  ]);
});
```

## Image helpers

Swap images and apply duotone overlays.

```ts
// Point an existing image to a different media file loaded via .loadMedia()
slide.modifyElement('Image Placeholder', [
  ModifyImageHelper.setRelationTarget('feather.png'),
]);

// Auto-crop to cover based on new image aspect ratio (uses template media)
slide.modifyElement('Image Placeholder', [
  ModifyImageHelper.setRelationTargetCover('feather.png', pres),
]);

// Duotone overlay on images/icons
slide.modifyElement('Icon', [
  ModifyImageHelper.setDuotoneFill({
    color: { type: 'srgbClr', value: '00AAFF' },
    tint: 70000, // 0-100000
    satMod: 90000, // 0-100000
  }),
]);
```

## Find out more

- [Add external image](https://github.com/singerla/pptx-automizer/blob/main/__tests__/add-external-image.test.ts)
- [Modify duotone color overlay for images](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-image-duotone.test.ts)
- [Swap image source on a slide master](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-master-external-image.test.ts)
