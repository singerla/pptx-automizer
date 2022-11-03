import Automizer, {
  ChartData,
  modify,
  TableRow,
  TableRowStyle,
  XmlHelper,
} from './index';
import { vd } from './helper/general-helper';

const automizer = new Automizer({
  templateDir: `${__dirname}/../__tests__/pptx-templates`,
  outputDir: `${__dirname}/../__tests__/pptx-output`,
  removeExistingSlides: true,
});

const run = async () => {
  const ppt = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithCharts.pptx`, 'charts')
    .load(`SlideWithImages.pptx`, 'images');

  ppt.addSlide('charts', 1);
  ppt.addSlide('images', 1);
  ppt.addSlide('images', 2);

  ppt.modify((xml: XMLDocument) => {
    // Dump before to console
    XmlHelper.dump(xml);

    const sldIdLst = xml.getElementsByTagName('p:sldIdLst')[0];
    const existingSlides = sldIdLst.getElementsByTagName('p:sldId');

    // reordering logic here
    const order = [3, 2, 1];
    let id = 256;
    order.forEach((sourceSlideNumber) => {
      const slide = existingSlides[sourceSlideNumber - 1];
      slide.setAttribute('id', String(id));
      sldIdLst.appendChild(slide);
      id++;
    });

    // Dump after to console
    XmlHelper.dump(sldIdLst);
  });

  const summary = await ppt.write('reorder.pptx');
};

run().catch((error) => {
  console.error(error);
});
