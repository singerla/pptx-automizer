import Automizer from '../src/automizer';

// Regression: an OLE object sitting on a slideLayout (e.g. think-cell data
// objects) must not crash the auto-import of masters. The OLE copier used to
// build a slide path from the layout's target number and failed with
// "Could not find file ppt/slides/slide<n>.xml".
test('Auto-import slideMaster with an OLE object on a slideLayout', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    autoImportSlideMasters: true,
  });

  const pres = await automizer
    .loadRoot(`EmptyTemplate.pptx`)
    .load('SlideLayoutWithOle.pptx')
    .addSlide('SlideLayoutWithOle.pptx', 1)
    .write(`add-slide-master-auto-import-ole.test.pptx`);

  expect(pres.masters).toBe(2);
});
