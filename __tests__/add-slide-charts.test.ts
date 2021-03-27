import Automizer from '../src/automizer';

test('create presentation and append charts to existing charts', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer.loadRoot(`RootTemplateWithCharts.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts');

  pres.addSlide('charts', 1);

  const result = await pres.write(`add-slide-charts.test.pptx`);

  expect(result.slides).toBe(3);
  expect(result.charts).toBe(3);
});
