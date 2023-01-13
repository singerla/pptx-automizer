import JSZip from 'jszip';
import { ElementSubtype, ElementType } from '../enums/element-type';
import { RelationshipAttribute } from './xml-types';

export type SourceSlideIdentifier = number | string;
export type SlideModificationCallback = (document: Document) => void;
export type ShapeModificationCallback = (
  XMLDocument: XMLDocument | Element,
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
   * Specify a fallback directory if template file was not found
   * in templateDir.
   */
  templateFallbackDir?: string;

  /**
   * Prefix for the output files for `Automizer` instance.
   * You can set a path here.
   */
  outputDir?: string;
  /**
   * Buffer unzipped pptx on disk
   */
  cacheDir?: string;
  /**
   * Zip compression level 0-9
   */
  compression?: number;
  rootTemplate?: string;
  presTemplates?: string[];
  useCreationIds?: boolean;
  /**
   * Delete all existing slides from rootTemplate
   * before automation starts.
   */
  removeExistingSlides?: boolean;
  /**
   * Eventually remove all unnecessary files from archive.
   */
  cleanup?: boolean;
  /**
   * statusTracker will be triggered on each appended slide.
   * You can e.g. attach a custom callback to a progress bar.
   */
  statusTracker?: StatusTracker['next'];
};
export type StatusTracker = {
  current: number;
  max: number;
  share: number;
  info: string | undefined;
  next: (tracker: StatusTracker) => void;
  increment: () => void;
};
export type AutomizerSummary = {
  status: string;
  duration: number;
  file: string;
  filename: string;
  templates: number;
  slides: number;
  charts: number;
  images: number;
};
export type Target = {
  file: string;
  type: string;
  filename: string;
  number?: number;
  rId?: string;
  prefix?: string;
  subtype?: ElementSubtype;
  filenameExt?: string;
  filenameBase?: string;
  getCreatedContent?: () => TrackedRelationInfo;
  getRelatedContent?: () => Promise<Target>;
  relatedContent?: Target;
};
export type FileInfo = {
  base: string;
  extension: string;
  dir: string;
  isDir: boolean;
};
export type TrackedFiles = Record<string, string[]>;
export type TrackedRelationInfo = {
  base: string;
  attributes?: RelationshipAttribute;
};
export type TrackedRelations = Record<string, TrackedRelationInfo[]>;
export type TrackedRelation = {
  tag: string;
  type?: string;
  attribute?: string;
  role?:
    | 'image'
    | 'slideMaster'
    | 'slide'
    | 'chart'
    | 'externalData'
    | 'slideLayout';
  targets?: Target[];
};
export type TrackedRelationTag = {
  source: string;
  relationsKey: string;
  isDir?: boolean;
  tags: TrackedRelation[];
  getTrackedRelations?: (role: string) => TrackedRelation[];
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
  hasCreationId?: boolean;
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
