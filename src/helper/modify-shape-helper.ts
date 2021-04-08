import { ShapeCoordinates } from '../types/shape-types';

export const setSolidFill = (element: XMLDocument): void => {
  element
    .getElementsByTagName('a:solidFill')[0]
    .getElementsByTagName('a:schemeClr')[0]
    .setAttribute('val', 'accent6');
};

export const setText = (text: string) => (element: XMLDocument): void => {
  element.getElementsByTagName('a:t')[0].firstChild.textContent = text;
};

// e.g. setPosition({x: 8000000, h:5000000})
export const setPosition = (pos: ShapeCoordinates) => (
  element: XMLDocument,
): void => {
  const map = {
    x: { tag: 'a:off', attribute: 'x' },
    y: { tag: 'a:off', attribute: 'y' },
    w: { tag: 'a:ext', attribute: 'cx' },
    h: { tag: 'a:ext', attribute: 'cy' },
  };

  const parent = 'a:xfrm';

  Object.keys(pos).forEach((key) => {
    element
      .getElementsByTagName(parent)[0]
      .getElementsByTagName(map[key].tag)[0]
      .setAttribute(map[key].attribute, pos[key]);
  });
};
