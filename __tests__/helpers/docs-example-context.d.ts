/**
 * Ambient context for __tests__/docs-examples.test.ts — and nothing else:
 * this file is passed explicitly to that test's TS program and is excluded
 * from the main tsconfig along with the rest of __tests__.
 *
 * The documentation corpus (README.md, AI-INSTRUCTOR.md, docs/) uses a small
 * set of conventional context variables across its ```ts examples: `pres` /
 * `automizer` for the Automizer instance, `slide` / `master` inside
 * modification callbacks, the `modify` facade and the Modify*Helper classes.
 * Declaring them here — with their real types imported from src — lets every
 * snippet body be fully typechecked without repeating boilerplate in each
 * fenced block. A snippet that declares its own `const pres = …` shadows the
 * global, so complete examples are unaffected.
 *
 * Do NOT add names here to silence a failing docs example unless the name is
 * an established docs-wide convention — otherwise fix the example, or mark
 * it ```ts ignore if it is deliberately partial.
 */

type _PptxAutomizerModule = typeof import('../../src/index');

// The Automizer class (value + instance type), and the two conventional
// names for an instance of it.
declare const Automizer: _PptxAutomizerModule['default'];
type Automizer = import('../../src/index').default;
declare const automizer: Automizer;
declare const pres: Automizer;

// Callback context objects.
declare const slide: import('../../src/index').ISlide;
declare const master: import('../../src/index').IMaster;

// The modify/read facades and the helper classes (static-method surfaces).
declare const modify: _PptxAutomizerModule['modify'];
declare const read: _PptxAutomizerModule['read'];
declare const ModifyHelper: _PptxAutomizerModule['ModifyHelper'];
declare const ModifyShapeHelper: _PptxAutomizerModule['ModifyShapeHelper'];
declare const ModifyTableHelper: _PptxAutomizerModule['ModifyTableHelper'];
declare const ModifyChartHelper: _PptxAutomizerModule['ModifyChartHelper'];
declare const ModifyTextHelper: _PptxAutomizerModule['ModifyTextHelper'];
declare const ModifyImageHelper: _PptxAutomizerModule['ModifyImageHelper'];
declare const ModifyColorHelper: _PptxAutomizerModule['ModifyColorHelper'];
declare const ModifyCleanupHelper: _PptxAutomizerModule['ModifyCleanupHelper'];
// not exported from the index (yet) — typed straight from its module
declare const ModifyHyperlinkHelper: typeof import('../../src/helper/modify-hyperlink-helper').default;
declare const XmlHelper: _PptxAutomizerModule['XmlHelper'];

// Unit conversion helpers.
declare const CmToDxa: _PptxAutomizerModule['CmToDxa'];
declare const DxaToCm: _PptxAutomizerModule['DxaToCm'];
declare const PtToEmu: _PptxAutomizerModule['PtToEmu'];
declare const EmuToPt: _PptxAutomizerModule['EmuToPt'];

// Conventional data variables in chart examples.
type ChartData = import('../../src/index').ChartData;
declare const chartData: ChartData;
declare const myData: ChartData;
