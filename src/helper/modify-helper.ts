import { XmlHelper } from './xml-helper';


export const setAttribute = (
  tagName: string,
  attribute: string,
  value: string | number,
  count?: number,
) => (element: XMLDocument): void => {
  element
    .getElementsByTagName(tagName)
    [count || 0].setAttribute(attribute, String(value));
};

export const dump = (element: XMLDocument | Document | Element): void => {
  XmlHelper.dump(element);
};
