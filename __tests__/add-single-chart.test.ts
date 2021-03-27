import Automizer from '../src/automizer';

test('create presentation and add some single charts', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty')
    .load(`SlideWithCharts.pptx`, 'charts');

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      slide.addElement('charts', 2, 'PieChart');
      slide.addElement('charts', 1, 'StackedBars');
    })
    .write(`add-single-chart.test.pptx`);

  expect(result.slides).toBe(2);
});
