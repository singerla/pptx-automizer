import { XmlElement } from './xml-types';

export type ModifyCallback = {
  (element: XmlElement);
};
export type ModifyCollectionCallback = {
  (collection: HTMLCollectionOf<XmlElement>);
};
export type Modification = {
  index?: number;
  last?: boolean;
  collection?: ModifyCollectionCallback;
  children?: ModificationTags;
  modify?: ModifyCallback | ModifyCallback[];
  create?: any;
  isRequired?: boolean;
  fromPrevious?: boolean;
  forceCreate?: boolean;
};
export type ModificationTags = {
  [tag: string]: Modification;
};
export type Color = {
  type?: 'schemeClr' | 'srgbClr';
  value: string;
};

export type TextStyle = {
  size?: number;
  color?: Color;
  isBold?: boolean;
  isItalics?: boolean;
};
export type ReplaceText = {
  replace: string;
  by: ReplaceTextReplacement | ReplaceTextReplacement[];
};
export type ReplaceTextReplacement = {
  text: string;
  style?: TextStyle;
};
export type ReplaceTextOptions = {
  openingTag: string;
  closingTag: string;
};
