---
title: Requirements and Limitations
description: What pptx-automizer can and cannot do — environment, shape and chart types, animations, PowerPoint version.
---

This generator can only be used on the server-side and requires a [Node.js](https://nodejs.org/en/download/package-manager/) environment.

## Shape Types

At the moment, you might encounter difficulties with special shape types that require additional relations (e.g., hyperlinks, video and audio may not work correctly). However, most shape types, including connection shapes, tables, and charts, are already supported. If you encounter any issues, please feel free to [report any issue](https://github.com/singerla/pptx-automizer/issues/new).

## Chart Types

Extended chart types, like waterfall or map charts, are basically supported. You might need additional modifiers to handle extended properties, which are not implemented yet. Please help to improve `pptx-automizer` and [report](https://github.com/singerla/pptx-automizer/issues/new) issues regarding extended charts.

## Animations

Animations are currently out of scope of this library. You might get errors on opening an output .pptx when there are added or removed shapes. This is because `pptx-automizer` doesn't synchronize `id`-attributes of animations with the existing shapes on a slide.

## Slide Masters and Layouts

Slide masters and their layouts can be imported, but individual slide layouts cannot be added, modified or removed directly, and layouts must not carry complex content like charts and images. See [Slide Masters and Layouts](./masters-layouts.md) for details and workarounds.

## Direct manipulation of elements

It is also important to know that `pptx-automizer` is currently limited to _adding_ things to the output presentation. If you require the ability to, for instance, modify a specific element on a slide within an existing presentation and leave the rest untouched, you will need to include all the other slides in the process. Find some workarounds in [Slide Management](./slide-management.md#loop-through-the-slides-of-a-presentation).

## PowerPoint version

All testing focuses on PowerPoint 2019 .pptx file format.
