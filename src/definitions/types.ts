export type ModificationCallback = (document: Document) => void;

export type ShapeCallback = (htmlElement: HTMLElement, arg1?: any, arg2?: any) => void;

export type Frame = {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}
