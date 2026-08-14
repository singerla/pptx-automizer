---
title: Troubleshooting
description: What to check when PowerPoint shows the repair prompt, and how to run the test suite.
---

If you encounter problems when opening a `.pptx`-file modified by this library, you might worry about PowerPoint not giving any details about the error. It can be hard to find the cause, but there are some things you can check:

- **Broken relation**: There are still unsupported shape types and `pptx-automizer` will not copy required relations of those. You can inflate `.pptx` output and check `ppt/slides/_rels/slide[#].xml.rels` files to find possible missing files.
- **Unsupported media**: You can also take a look at the `ppt/media`-directory of an inflated `.pptx`-file. If you discover any unusual file formats, remove or replace the files by one of the [known types](https://github.com/singerla/pptx-automizer/blob/main/src/enums/content-type-map.ts).
- **Broken animation**: Pay attention to modified/removed shapes which are part of an animation. In case of doubt, (temporarily) remove all animations from your template. (see [#78](https://github.com/singerla/pptx-automizer/issues/78))
- **Proprietary/Binary contents** (e.g. ThinkCell): Walk through all slides, slideMasters and slideLayouts and seek for hidden Objects. Hit `ALT+F10` to toggle the sidebar.
- **Chart datasheet won't open** If you encounter an error message on opening a chart's datasheet, please make sure that the data table (blue bordered rectangle in worksheet view) of your template starts at cell A:1. If not, open worksheet in Excel mode and edit the table size in the table draft tab.

## Bisecting the repair prompt

If the output opens with a PowerPoint repair prompt and the checks above don't point to a cause:

1. Retry with `cleanup: false` and `cleanupPlaceholders: false`.
2. Remove exotic elements (animations, embedded videos, ThinkCell remnants) from the template.
3. Isolate the failing slide by bisecting your `addSlide` calls — comment out half of them, re-run, and narrow down until the offending slide (or modification) is found.

If none of these could help, please don't hesitate to [talk about it](https://github.com/singerla/pptx-automizer/issues/new).

## Testing

You can run all unit tests using these commands:

```
yarn test
yarn test-coverage
```
