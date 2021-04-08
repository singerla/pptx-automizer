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
