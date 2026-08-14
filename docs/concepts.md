---
title: Concepts
description: The mental model behind pptx-automizer — deferred execution, the template/root model, numbering, and the one-instance rule.
---

Four things explain almost everything about how `pptx-automizer` behaves. Read this page before anything else.

## Deferred execution

The single most important thing to know: calls like `addSlide()`, `addElement()`
and `modifyElement()` only **queue** work — including every modification
callback you pass. Nothing is applied to the output presentation until you call
`await pres.write(...)` (or `.stream()` / `.getJSZip()`), which executes the
whole queue in order. Two practical consequences:

- A throwing callback or an unresolvable element selector surfaces as a rejected
  `write()` (with a typed error such as `CallbackError`), not at the line where
  you queued it. Set `continueOnError: true` to log a warning and skip the
  failing modification instead.
- Anything you compute *inside* a callback runs at `write()` time; variables it
  closes over will have their values from that moment, not from when the
  callback was queued.

## The template/root model

Every build starts from a **root template** (`loadRoot`) that the output
presentation is based on, plus any number of further templates (`load`) that
serve as shape and slide sources. This is how it works internally:

- Load a root template to append slides to
- (Probably) load root template again to modify slides
- Load other templates
- Append a loaded slide to (probably truncated) root template
- Modify the recently added slide
- Write root template and appended slides as output presentation

`pptx-automizer` is currently limited to _adding_ things to the output
presentation. If you require the ability to, for instance, modify a specific
element on a slide within an existing presentation and leave the rest
untouched, you will need to include all the other slides in the process — see
[looping through the slides of a presentation](./slide-management.md#loop-through-the-slides-of-a-presentation).

## 1-based numbering

Slide numbers are **1-based**: `addSlide('shapes', 1)` takes the first slide of
the template labelled `shapes`. The number addresses the slide file inside the
.pptx — it is not necessarily the position in the final deck. Template files
are addressed by filename or the label given to `.load()`.

## One instance, one output

Use **one Automizer instance per presentation build**, and call the output
method (`write`/`stream`/`getJSZip`) **once** per instance. If you need
several output files, create a fresh `new Automizer(...)` per file. Separate
instances are isolated from each other — running several builds concurrently
in the same process (e.g. `Promise.all` in a server) is supported and covered
by a regression test.
