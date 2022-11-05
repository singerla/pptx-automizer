import Automizer, {
  ChartData,
  modify,
  TableRow,
  TableRowStyle,
  XmlHelper,
} from './index';
import { vd } from './helper/general-helper';
import ModifyPresentationHelper from './helper/modify-presentation-helper';

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
  ppt.addSlide('charts', 2);
  ppt.addSlide('images', 1);
  ppt.addSlide('images', 2);

  ppt.modify(ModifyPresentationHelper.sortSlides([3, 2, 1]));

  const summary = await ppt.write('reorder.pptx');
};

run().catch((error) => {
  console.error(error);
});
