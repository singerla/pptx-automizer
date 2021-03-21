import JSZip from 'jszip';
import { ElementType } from '../enums/element-type';

export type ModificationCallback = (document: Document) => void;

export type ShapeCallback = (htmlElement: HTMLElement, arg1?: any, arg2?: any) => void;

export type Frame = {
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export type AutomizerParams = {
  templateDir?: string;
  outputDir?: string;
}
export type AutomizerSummary = {
  status: string;
  duration: number;
  file: string;
  templates: number;
  slides: number;
  charts: number;
  images: number;
}
export type Target = {
  file: string;
  number: number;
  rId?: string;
}
export type ImportElement = {
  presName: string;
  slideNumber: number;
  selector: string;
  mode: string;
  callback?: Function | Function[];
}
export type ImportedElement = {
  mode: string;
  name?: string;
  sourceArchive: JSZip;
  sourceSlideNumber: number;
  callback?: any;
  target?: AnalyzedElementType['target'];
  type?: AnalyzedElementType['type'];
  sourceElement?: AnalyzedElementType['element'];
}
export type AnalyzedElementType = {
  type: ElementType;
  target?: Target;
  element?: HTMLElement;
}
export type TargetByRelIdMapParam = {
  relRootTag: string;
  relAttribute: string;
  prefix: string;
  expression?: RegExp;
}
export type Workbook = {
  archive: JSZip;
  sheet: Document | any;
  sharedStrings: Document;
  table: Document;
}
