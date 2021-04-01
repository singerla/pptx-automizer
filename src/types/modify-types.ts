export type ModifyCallback = {
  (element: Element);
};
export type Modification = {
  index?: number;
  children?: ModificationTags;
  modify?: ModifyCallback | ModifyCallback[];
};
export type ModificationTags = {
  [tag: string]: Modification;
};
