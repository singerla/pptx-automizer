import Automizer, { modify } from './index';

const run = async () => {
  const outputDir = `${__dirname}/../__tests__/pptx-output`;
  const templateDir = `${__dirname}/../__tests__/pptx-templates`;

  const automizer = new Automizer({
    templateDir,
    outputDir,
    verbosity: 2,
  });

  const dataSmaller = {
    series: [{ label: 'series s1' }, { label: 'series s2' }],
    categories: [
      { label: 'item test r1', values: [10, 16] },
      { label: 'item test r2', values: [12, 18] },
    ],
  };

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`ChartBarsStacked.pptx`, 'charts')
    .load(`ChartWithExternalData.pptx`, 'xml');

  pres
    .addSlide('xml', 1)
    .addSlide('xml', 1, (slide) => {
      slide.removeElement('Gráfico 7');
      slide.addElement('charts', 1, 'BarsStacked');
    })
    .addSlide('charts', 1, (slide) => {
      slide.modifyElement('BarsStacked', [modify.setChartData(dataSmaller)]);
    })
    .addSlide('charts', 1, (slide) => {
      slide.removeElement('BarsStacked');
      slide.addElement('xml', 1, 'Gráfico 7');
    });

  pres.write(`testCharts.pptx`).then((summary) => {
    console.log(summary);
  });
};

run().catch((error) => {
  console.error(error);
});
