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

## Writing raw XML callbacks

Not every OOXML property has a `modify.*` helper (`cap`/`cmpd` line attributes,
custom geometry, `a:effectLst`, …). Any callback you pass to
`modifyElement`/`addElement` receives the shape's XML element — the `<p:sp>`,
`<p:pic>` or `<p:graphicFrame>` node — so you can edit it directly:

```ts
import { XmlElement } from 'pptx-automizer';

slide.modifyElement('MyBox', (element: XmlElement) => {
  // ...manipulate the DOM node...
});
```

Rules for hand-written callbacks:

1. **Scope your lookups.** `element.getElementsByTagName('a:ln')` searches *all*
   descendants of the shape, including text run properties (`a:rPr`) — where
   `a:ln` also occurs. Reach the container first
   (`element.getElementsByTagName('p:spPr')[0]`), then prefer a **direct child**
   scan over another `getElementsByTagName`.
2. **A missing element is the normal case.** If a property was never overridden
   in PowerPoint, it is inherited from the theme/shape style and simply absent
   from the slide XML. Always handle "modify existing" *and* "create new".
3. **Child order follows the schema, not your call order.** Appending in the
   wrong order is what makes PowerPoint show the "repair" prompt on open. For
   `p:spPr` the sequence is `a:xfrm` → geometry (`a:prstGeom`/`a:custGeom`) →
   fill → `a:ln` → `a:effectLst` → `a:scene3d` → `a:sp3d` → `a:extLst`; inside
   `a:ln` it is fill → `a:prstDash` → join (`a:round`/`a:bevel`/`a:miter`) →
   `a:headEnd` → `a:tailEnd`.
4. **Inspect before you guess:** `slide.modifyElement('MyBox', modify.dump)`
   prints the shape's current XML to the console. Do that first when unsure what
   the template actually contains.
5. **A throwing callback rejects `write()`** with a `CallbackError` naming the
   slide and element — unless `continueOnError: true` is set, which logs a
   warning and skips the modification instead (see
   [deferred execution](./concepts.md#deferred-execution)).
6. Use `XmlHelper` (exported) for common DOM chores: `XmlHelper.remove(node)`,
   `insertAfter(new, ref)`, `getClosestParent('p:sp', node)`,
   `appendClone(node, parent)`, `dump(node)`.

If a raw callback you wrote turns out to be generally useful, it is a good
candidate for a new `modify.*` helper — see
[AGENTS.md](https://github.com/singerla/pptx-automizer/blob/main/AGENTS.md) in
the repository.

### Worked example: shape outline (weight + color)

> Outlines have a dedicated modifier — use `ModifyShapeHelper.setOutline` from
> the [shape helpers](#shape-helpers) for real work. The example is kept because
> it shows the general technique on a realistic property.

```ts
import Automizer, { XmlElement } from 'pptx-automizer';

// p:spPr children that must stay AFTER a:ln
const AFTER_LN = ['a:effectLst', 'a:effectDag', 'a:scene3d', 'a:sp3d', 'a:extLst'];
const childByName = (parent: XmlElement, names: string[]) =>
  Array.from(parent.childNodes as any).find((n: any) =>
    names.includes(n.nodeName),
  ) as XmlElement | undefined;

const setOutline =
  (outline: { weight?: number; color?: string }) =>
  (element: XmlElement) => {
    const spPr = element.getElementsByTagName('p:spPr')[0];
    if (!spPr) return;

    // Direct child only — a:ln also lives inside text run properties.
    let ln = childByName(spPr, ['a:ln']);
    if (!ln) {
      ln = spPr.ownerDocument.createElement('a:ln');
      const anchor = childByName(spPr, AFTER_LN);
      anchor ? spPr.insertBefore(ln, anchor) : spPr.appendChild(ln);
    }

    if (outline.weight !== undefined) {
      // a:ln/@w is EMU: 1pt = 12700
      ln.setAttribute('w', String(Math.round(outline.weight * 12700)));
    }

    if (outline.color) {
      const solidFill = ln.ownerDocument.createElement('a:solidFill');
      const srgbClr = ln.ownerDocument.createElement('a:srgbClr');
      srgbClr.setAttribute('val', outline.color.replace('#', '')); // no '#'!
      solidFill.appendChild(srgbClr);

      // Fill is the FIRST child of a:ln — replace whatever fill is there.
      const currentFill = childByName(ln, [
        'a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill',
      ]);
      currentFill
        ? ln.replaceChild(solidFill, currentFill)
        : ln.insertBefore(solidFill, ln.firstChild);
    }
  };

slide.modifyElement('MyBox', setOutline({ weight: 2, color: 'FFFFFF' }));
```

Produces `<a:ln w="25400"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`.
Caveat worth knowing: if the shape's existing `a:ln` contains `<a:noFill/>`
(outline explicitly turned off in PowerPoint), setting only the weight yields
`<a:ln w="…"><a:noFill/></a:ln>` — a thick *invisible* line. Set a color too, or
replace the `a:noFill` node.

### Units reference

OOXML uses no single unit. When writing raw XML:

| What | Unit | Conversion |
|---|---|---|
| Position/size (`a:off`, `a:ext`), line width `a:ln/@w`, corner radius | EMU | 1 cm = 360000 · 1 inch = 914400 · 1 pt = 12700 |
| `modify.setPosition` / `updatePosition` / `setOutline` | same EMU values | helpers `CmToDxa(cm)` / `DxaToCm(v)` (name says Dxa, value is EMU), `PtToEmu(pt)` / `EmuToPt(v)` |
| Rotation (`a:xfrm/@rot`) | 1/60000 degree | 45° = 2700000 |
| Font size (`a:rPr/@sz`, `TextStyle.size`) | 1/100 pt | 18pt = 1800 |
| Percentages (`a:alpha/@val`, `a:lumMod`, …) | 1/1000 % | 50% = 50000 |
| Colors (`a:srgbClr/@val`) | 6-digit hex | `'FF0000'`, never `'#FF0000'` |
| PptxGenJS `slide.generate(...)` | inches | (different world — see [Generate shapes](./generation.md)) |
