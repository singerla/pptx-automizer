import { XmlHelper } from './xml';
import { Frame, Workbook } from '../types/types';

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
export const setPosition = (pos: Frame) => (element: HTMLElement) => {
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
