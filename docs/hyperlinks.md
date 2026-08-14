---
title: Hyperlink Management
description: Add, update and remove external or internal hyperlinks on existing shapes.
---

PowerPoint presentations often use hyperlinks to connect to external websites or internal slides. The `pptx-automizer` provides simple and powerful functions to manage hyperlinks in your presentations.

## Add Hyperlinks to existing shapes

You can add hyperlinks to template text shapes using the `addHyperlink` helper function. The function accepts either a URL string for external links or a slide number for internal slide links:

```ts
import { XmlElement } from 'pptx-automizer';

// Add an external hyperlink
slide.modifyElement('TextShape', modify.addHyperlink('https://example.com'));

// Add an internal slide link (to slide 3)
slide.modifyElement('TextShape', (element: XmlElement, relation?: XmlElement) => {
  modify.addHyperlink(3)(element, relation);
});
```

The `addHyperlink` function will automatically detect whether the target is an external URL or an internal slide number and set up the appropriate relationship type and attributes.

## Update or remove existing hyperlinks

Use `modify.setHyperlinkTarget` to change the target of a hyperlink that already exists on a shape. The second argument controls whether the new target is external (default, `true`) or an internal slide link (`false`):

```ts
// Point an existing hyperlink to a new external URL
slide.modifyElement('TextShape', modify.setHyperlinkTarget('https://example.com'));

// Point an existing hyperlink to an internal slide (e.g. slide 5)
slide.modifyElement('TextShape', modify.setHyperlinkTarget(5, false));
```

Use `modify.removeHyperlink` to strip the hyperlink from a shape while keeping its text:

```ts
slide.modifyElement('TextShape', modify.removeHyperlink());
```

## Related

- To create a **new** hyperlinked shape from scratch, use the PptxGenJS wrapper — see [Generate shapes with PptxGenJS](./generation.md#create-a-new-hyperlinked-text-shape).
- Inline hyperlinks inside generated text (external targets or slide numbers) are also supported by the [MultiText/HTML text helpers](./text.md#text-helpers-multitexthtml).
