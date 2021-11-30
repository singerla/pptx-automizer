export type ModifyCallback = {
  (element: Element);
};
export type ModifyCollectionCallback = {
  (collection: HTMLCollectionOf<Element>);
};
export type Modification = {
  index?: number;
  collection?: ModifyCollectionCallback | ModifyCollectionCallback;
  children?: ModificationTags;
  modify?: ModifyCallback | ModifyCallback[];
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
