import JSZip from 'jszip';
import { ElementType } from './enums';

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

export interface ICounter {
  set(): void | PromiseLike<void>;

  get(): number;

  name: string;
  count: number;

  _increment(): number;
}

export interface ISlide {
  sourceArchive: JSZip;
  sourceNumber: number;
  modifications: Function[];
  modify: Function;

  append(targetTemplate: RootPresTemplate): Promise<void>;

  addElement(presName: string, slideNumber: number, selector: Function | string): void;
}

export interface IPresentationProps {
  rootTemplate: RootPresTemplate;
  templates: PresTemplate[];
  params: AutomizerParams;
  timer: number;

  template(name: string): PresTemplate;
}

export interface ITemplate {
  location: string;
  file: Promise<Buffer>;
  archive: Promise<JSZip>;
}

export interface RootPresTemplate extends ITemplate {
  slides: ISlide[];
  counter: ICounter[];

  count(name: string): number;

  incrementCounter(name: string): number;

  appendSlide(slide: ISlide): Promise<void>;
}

export interface PresTemplate extends ITemplate {
  name: string;
}

export interface IShape {
  sourceArchive: JSZip;
  targetArchive: JSZip;
}

export interface IChart extends IShape {
  sourceNumber: number;
  targetNumber: number;

  append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<IChart>;
}

export interface IImage extends IShape {
  sourceFile: string;
  targetFile: string;
  contentTypeMap: any;

  append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<IImage>;
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
