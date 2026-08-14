---
title: Output
description: Write a file, stream the archive, or access the underlying JSZip; track status and control cleanup.
---

Calling an output method executes the whole queued modification chain (see [deferred execution](./concepts.md#deferred-execution)) and produces the final archive. Always `await` the output call — nothing happens without it, and one output call per Automizer instance.

## write, stream, getJSZip

```ts
// Write the output file.
pres.write('myPresentation.pptx').then((summary) => {
  console.log(summary);
});

// It is also possible to get a ReadableStream.
// stream() accepts JSZip.JSZipGeneratorOptions for 'nodebuffer' type.
const stream = await pres.stream({
  compressionOptions: {
    level: 9,
  },
});
// You can e.g. output the pptx archive to stdout instead of writing a file:
stream.pipe(process.stdout);

// If you need any other output format, you can eventually access
// the underlying JSZip instance:
const finalJSZip = await pres.getJSZip();
// Convert the output to whatever needed:
const base64 = await finalJSZip.generateAsync({ type: 'base64' });
```

By default, a throwing modification callback or an unresolvable element selector rejects `write()` with a typed error (`CallbackError`, `ElementNotFoundError`). Set `continueOnError: true` to log a warning and skip the failing modification instead.

## Track status of automation process

When creating large presentations, you might want to have some information about the current status. Use a custom status tracker:

```ts
import Automizer, { StatusTracker } from 'pptx-automizer';

// If you want to track the steps of creation process,
// you can use a custom callback:
const myStatusTracker = (status: StatusTracker) => {
  console.log(status.info + ' (' + status.share + '%)');
};

const automizer = new Automizer({
  // ...
  statusTracker: myStatusTracker,
});
```

## Cleanup and compression flags

The `Automizer` constructor takes a few options that affect the output archive:

```ts
const automizer = new Automizer({
  templateDir: `my/pptx/templates`,
  outputDir: `my/pptx/output`,

  // activate `cleanup` to eventually remove unused files from the archive:
  cleanup: false,

  // Remove all unused placeholders to prevent unwanted overlays:
  cleanupPlaceholders: false,

  // Set a value from 0-9 to specify the zip-compression level.
  // The lower the number, the faster your output file will be ready.
  // Higher compression levels produce smaller files.
  compression: 0,
});
```

If an output file triggers PowerPoint's repair prompt, retry with `cleanup: false` and `cleanupPlaceholders: false` first — see [Troubleshooting](./troubleshooting.md).
