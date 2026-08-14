---
title: Helper Utilities
description: Shape, cleanup, unit-conversion, generic and advanced XML helpers exposed by pptx-automizer.
---

This section documents handy helpers exposed by `pptx-automizer` that provide extra capabilities. You can import helpers directly from the package:

```ts
import {
  ModifyShapeHelper,
  ModifyTableHelper,
  ModifyCleanupHelper,
  ModifyTextHelper,
  ModifyImageHelper,
} from 'pptx-automizer';
```

Some helper families are documented with their feature area:

- [Table helpers](./tables.md#table-helpers) (`ModifyTableHelper`)
- [Text helpers (MultiText/HTML)](./text.md#text-helpers-multitexthtml) (`ModifyTextHelper`)
- [Image helpers](./images.md#image-helpers) (`ModifyImageHelper`)

## Shape helpers

Use `ModifyShapeHelper` to quickly adjust common properties of shapes and text frames.

- Solid fill color

```ts
slide.modifyElement('MyShape', [
  // sets the shape's solid fill to theme color "accent6"
  ModifyShapeHelper.setSolidFill,
]);
```

- Outline (line) width, color and dash style

```ts
import { PtToEmu } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  ModifyShapeHelper.setOutline({
    weight: PtToEmu(2), // EMU, 1pt = 12700
    color: { type: 'srgbClr', value: 'FF0000' },
    type: 'sysDash', // any a:prstDash value, e.g. solid, dash, dot
  }),
]);
```

Only the given properties are modified. An outline is created if the shape has
none, which is the case whenever it was never overridden in PowerPoint and is
inherited from the theme or shape style. If the outline was explicitly switched
off in PowerPoint, a `weight` alone stays invisible — pass a `color`, too. Use
`ModifyCleanupHelper.removeBorder` to remove an outline.

- Bullet list from strings

```ts
slide.modifyElement('MyTextBox', [
  ModifyShapeHelper.setBulletList(['Item 1', 'Item 2', 'Item 3']),
]);
```

- Replace tagged text (see [Replace tagged text](./text.md#replace-tagged-text) for the
  tag concept; the default delimiters are `{{` and `}}`)

```ts
slide.modifyElement('MyTextBox', [
  ModifyShapeHelper.replaceText(
    [
      { replace: 'company', by: { text: 'Globex' } },
      { replace: 'year', by: { text: '2025', style: { isBold: true } } },
    ],
    { openingTag: '{{', closingTag: '}}' },
  ),
]);
```

- Position, size, rotation, rounded corners

```ts
import { CmToDxa } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  // set absolute position/size
  ModifyShapeHelper.setPosition({ x: CmToDxa(2), y: CmToDxa(3), w: CmToDxa(6), h: CmToDxa(2) }),
  // update only some props, leave others untouched
  ModifyShapeHelper.updatePosition({ x: CmToDxa(4) }),
  // rotate clockwise in degrees
  ModifyShapeHelper.rotate(15),
  // rounded rectangle corners (0-100000)
  ModifyShapeHelper.roundedCorners(25000),
]);
```

## Cleanup helpers

`ModifyCleanupHelper` helps remove formatting noise from shapes when you need a clean base.

```ts
import { XmlElement } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  ModifyCleanupHelper.removeBackground,
  ModifyCleanupHelper.removeBorder,
  ModifyCleanupHelper.removeEffects,
  // text-level cleanup
  ModifyCleanupHelper.clearTextUnderline,
  ModifyCleanupHelper.clearTextBold,
  ModifyCleanupHelper.clearTextSize,
  // remove all explicit text colors …
  (element: XmlElement) => ModifyCleanupHelper.clearTextColor(element),
  // … or pass a color to set a uniform one instead
  (element: XmlElement) =>
    ModifyCleanupHelper.clearTextColor(element, {
      type: 'srgbClr',
      value: 'FF0000',
    }),
]);
```

Other useful helpers include: `removeTextEffects`, `removeFillEffects`, `remove3dEffects`, `removeShadowEffects`, and `removeExtLst`.

## Unit conversion helpers

PowerPoint stores coordinates and sizes in the `dxa` (EMU) unit. Use the exported converters to work with centimeters instead:

```ts
import { CmToDxa, DxaToCm } from 'pptx-automizer';

// centimeters -> dxa (e.g. when setting position/size)
const widthInDxa = CmToDxa(6); // 2160000

// dxa -> centimeters (e.g. when reading shape coordinates)
const widthInCm = DxaToCm(2160000); // 6
```

Line weights are usually given in points, use `PtToEmu`/`EmuToPt` for those:

```ts
import { PtToEmu, EmuToPt } from 'pptx-automizer';

const weight = PtToEmu(1.5); // 19050
const inPoints = EmuToPt(19050); // 1.5
```

## Generic / debugging helpers

`ModifyHelper` (also available through the `modify` namespace) offers low-level callbacks that are handy for debugging or custom XML tweaks:

```ts
import { modify } from 'pptx-automizer';

slide.modifyElement('MyShape', [
  // print the element's XML to the console
  modify.dump,
  // print the related chart XML to the console
  modify.dumpChart,
  // set an attribute on the first matching tag (optionally by index)
  modify.setAttribute('a:off', 'x', 1000000),
]);
```

## Advanced XML helpers (power users)

For advanced scenarios, you can inspect slide XML and relationships. These are considered expert APIs and may change.

Import them from the package root:

```ts
import { XmlSlideHelper, XmlRelationshipHelper } from 'pptx-automizer';
```

Examples:

- Read all text element IDs on a slide: `new XmlSlideHelper(slideXml).getAllTextElementIds()`
- Get named elements: `new XmlSlideHelper(slideXml).getNamedElements(['p:sp'])`
- Table introspection: `XmlSlideHelper.readTableInfo(element)`
- Relationship targets by type or prefix: `new XmlRelationshipHelper(relsXml).getTargetsByType(type)`

See tests for practical usage:

- [Find all text elements on a slide](https://github.com/singerla/pptx-automizer/blob/main/__tests__/get-all-text-element-ids.test.ts)
- [Read shape/group info](https://github.com/singerla/pptx-automizer/blob/main/__tests__/read-shape-info.test.ts)
- [Read group info](https://github.com/singerla/pptx-automizer/blob/main/__tests__/read-group-info.test.ts)
