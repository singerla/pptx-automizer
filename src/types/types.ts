import { ElementSubtype, ElementType } from '../enums/element-type';
import { RelationshipAttribute, XmlDocument, XmlElement } from './xml-types';
import IArchive, { ArchiveMode } from '../interfaces/iarchive';

export type ShapeTargetType = 'slide' | 'slideMaster';
export type SourceIdentifier = number | string;
export type SlideModificationCallback = (document: XmlDocument) => void;
export type ShapeModificationCallback = (
  XmlDocument: XmlDocument | XmlElement,
  arg1?: XmlDocument,
  arg2?: Workbook,
) => void;
export type GetRelationshipsCallback = (
  element: XmlElement,
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
   * Absolute path to cache directory.
   */
  archiveType?: ArchiveParams;
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
  masters: number;
};
export type Target = {
  file: string;
  type: string;
  filename: string;
  number?: number;
  rId?: string;
  prefix?: string;
  element?: XmlElement;
  subtype?: ElementSubtype;
  filenameExt?: string;
  filenameBase?: string;
  getCreatedContent?: () => TrackedRelationInfo;
  getRelatedContent?: () => Promise<Target>;
  relatedContent?: Target;
  copiedTarget?: string;
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
export type ArchiveParams = {
  mode: ArchiveMode;
  baseDir?: string;
  workDir?: string;
  cleanupWorkDir?: boolean;
  name?: string;
};
export type ImportElement = {
  presName: string;
  slideNumber: number;
  selector: FindElementSelector;
  mode: string;
  callback?: ShapeModificationCallback | ShapeModificationCallback[];
  info?: any;
};
export type FindElementSelector =
  | string
  | {
      creationId: string;
      name: string;
    };
export type FindElementStrategy = {
  mode: 'findByElementCreationId' | 'findByElementName';
  selector: string;
};
export type ImportedElement = {
  mode: string;
  name?: string;
  hasCreationId?: boolean;
  sourceArchive: IArchive;
  sourceSlideNumber: number;
  callback?: ImportElement['callback'];
  target?: AnalyzedElementType['target'];
  type?: AnalyzedElementType['type'];
  sourceElement?: AnalyzedElementType['element'];
};
export type AnalyzedElementType = {
  type: ElementType;
  target?: Target;
  element?: XmlDocument;
};
export type TargetByRelIdMapParam = {
  relRootTag: string;
  relAttribute: string;
  prefix: string;
  expression?: RegExp;
};
export type Workbook = {
  archive: IArchive;
  sheet: XmlDocument;
  sharedStrings: XmlDocument;
  table: XmlDocument;
};
