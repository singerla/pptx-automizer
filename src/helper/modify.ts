import { XmlHelper } from './xml-helper';
import { ChartData, FrameCoordinates, Workbook } from '../types/types';
import { XmlChartHelper } from './xml-chart-helper';

export const setSolidFill = (element: XMLDocument): void => {
  element.getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6');
};

export const setText = (text: string) => (element: XMLDocument): void => {
  element.getElementsByTagName('a:t')[0]
    .firstChild
    .textContent = text;
};

// eslint-disable-next-line @typescript-eslint/no-unused-vars
export const revertElements = (slide: Document): void => {
  // dump(slide)
};

// e.g. setPosition({x: 8000000, h:5000000})
export const setPosition = (pos: FrameCoordinates) => (element: XMLDocument): void => {
  const map = {
    x: {tag: 'a:off', attribute: 'x'},
    y: {tag: 'a:off', attribute: 'y'},
    w: {tag: 'a:ext', attribute: 'cx'},
    h: {tag: 'a:ext', attribute: 'cy'},
  };

  const parent = 'a:xfrm';

  Object.keys(pos)
    .forEach(key => {
      element.getElementsByTagName(parent)[0]
        .getElementsByTagName(map[key].tag)[0]
        .setAttribute(map[key].attribute, pos[key]);
    });
};

export const setAttribute = (tagName: string, attribute: string, value: string | number, count?: number) => (element: XMLDocument): void => {
  element.getElementsByTagName(tagName)[count || 0]
    .setAttribute(attribute, String(value));
};

export const setChartData = (data: ChartData) => (element: XMLDocument, chart: Document, workbook: Workbook): void => {
  XmlChartHelper.setChartData(chart, workbook, data);
};

export const dump = (element: XMLDocument | Document): void => {
  XmlHelper.dump(element);
};
