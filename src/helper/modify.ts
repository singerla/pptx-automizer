import { Workbook } from '../definitions/app';
import XmlHelper from './xml';

export const setSolidFill = (element) => {
  element.getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6');
};

export const setText = (text: string) => (element) => {
  element.getElementsByTagName('a:t')[0]
    .firstChild
    .data = text;
};

export const revertElements = (slide: Document) => {
  // dump(slide)
};

// e.g. setPosition({x: 8000000, h:5000000})
export const setPosition = (pos: any) => (element: HTMLElement) => {
  let map = {
    x: {tag: 'a:off', attribute: 'x'},
    y: {tag: 'a:off', attribute: 'y'},
    w: {tag: 'a:ext', attribute: 'cx'},
    h: {tag: 'a:ext', attribute: 'cy'},
  };

  let parent = 'a:xfrm';

  for (let key in pos) {
    element.getElementsByTagName(parent)[0]
      .getElementsByTagName(map[key].tag)[0]
      .setAttribute(map[key].attribute, pos[key]);
  }
};

export const setAttribute = (tagName: string, attribute: string, value: string | number, count?: number) => (element: HTMLElement) => {
  element.getElementsByTagName(tagName)[count || 0]
    .setAttribute(attribute, String(value));
};

export const setChartData = (data: any) => (element: HTMLElement, chart: Document, workbook: Workbook) => {
  XmlHelper.setChartData(chart, workbook, data);
};

export const dump = (element: HTMLElement | Document) => {
  XmlHelper.dump(element);
};
