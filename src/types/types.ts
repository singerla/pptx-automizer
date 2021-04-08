import JSZip from 'jszip';
import { ElementType } from '../enums/element-type';

export type SlideModificationCallback = (document: Document) => void;
export type ShapeModificationCallback = (
  XMLDocument: XMLDocument,
  arg1?: Document,
  arg2?: Workbook,
) => void;
export type GetRelationshipsCallback = (
  element: Element,
  rels: Target[],
) => void;

export type AutomizerParams = {
  /**
   * Prefix for all template files. You can set a path here.
   */
  templateDir?: string;

  /**
   * Prefix for the output files for `Automizer` instance.
   * You can set a path here.
   */
  outputDir?: string;
};
export type AutomizerSummary = {
  status: string;
  duration: number;
  file: string;
  templates: number;
  slides: number;
  charts: number;
  images: number;
};
export type Target = {
  file: string;
  number?: number;
  rId?: string;
};
export type ImportElement = {
  presName: string;
  slideNumber: number;
  selector: string;
  mode: string;
  callback?: ShapeModificationCallback | ShapeModificationCallback[];
};
export type ImportedElement = {
  mode: string;
  name?: string;
  sourceArchive: JSZip;
  sourceSlideNumber: number;
  callback?: ImportElement['callback'];
  target?: AnalyzedElementType['target'];
  type?: AnalyzedElementType['type'];
  sourceElement?: AnalyzedElementType['element'];
};
export type AnalyzedElementType = {
  type: ElementType;
  target?: Target;
  element?: XMLDocument;
};
export type TargetByRelIdMapParam = {
  relRootTag: string;
  relAttribute: string;
  prefix: string;
  expression?: RegExp;
};
export type Workbook = {
  archive: JSZip;
  sheet: XMLDocument;
  sharedStrings: XMLDocument;
  table: XMLDocument;
};
