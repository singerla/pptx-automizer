import { ElementSubtype, ElementType } from '../enums/element-type';
import {
  ElementInfo,
  RelationshipAttribute,
  SlideInfo,
  TemplateInfo,
  XmlDocument,
  XmlElement,
} from './xml-types';
import IArchive, { ArchiveMode } from '../interfaces/iarchive';
import { ContentTypeExtension } from '../enums/content-type-map';
import PptxGenJS from 'pptxgenjs';
import { Logger } from '../helper/general-helper';

export type ShapeTargetType = 'slide' | 'slideMaster' | 'slideLayout';
export type SourceIdentifier = number | string;
export type SlideModificationCallback = (document: XmlDocument) => void;
export type SlidePlaceholder = {
  xml: XmlElement;
  type: string;
  id?: number;
};
export type ModificationCallback =
  | ChartModificationCallback
  | ShapeModificationCallback;
export type ShapeModificationCallback = (
  element: XmlElement,
  relation?: XmlElement,
) => void;
export type ChartModificationCallback = (
  element: XmlElement,
  chart?: XmlDocument,
  workbook?: Workbook,
) => void;
export type GetRelationshipsCallback = (
  element: XmlElement,
  rels: Target[],
) => void;

export type AutomizerFile = string | Buffer | Uint8Array;

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
   * Path to media directory, in case you need to import additional
   * files. You can set a path here. Load files with Automizer.loadMedia
   */
  mediaDir?: string;
  /**
   * Absolute path to cache directory.
   */
  archiveType?: ArchiveParams;
  /**
   * Zip compression level 0-9
   */
  compression?: number;
  rootTemplate?: AutomizerFile;
  /**
   * Array of template files to be loaded on initialization.
   * If files are Buffer or Uint8Array, they will be named 0.pptx, 1.pptx, ...
   * according to their order in the array.
   */
  presTemplates?: AutomizerFile[];
  useCreationIds?: boolean;
  /**
   * Turn this to true if you always want to import all required slide masters.
   * You don't need to adjust with slide.useSlideLayout, but it will have a
   * negative impact on performance.
   * It is highly recommended to activate autoImportSlideMasters in case your
   * loaded templates have different sets of slideMasters & -layouts.
   */
  autoImportSlideMasters?: boolean;
  /**
   * In case you encounter weird pptx messages on opening a created presentation,
   * you can turn this to true. It will log a message on missing related contents
   * and help you to locate where it is. Use this along with "assertRelatedContents"
   * to auto-fix broken relations.
   */
  showIntegrityInfo?: boolean;
  /**
   * Pptx-automizer can try to add any missing related content that could not be
   * handled properly by "addElement" or "addMaster" or one of their subroutines.
   * This probably fixes corrupted pptx files.
   */
  assertRelatedContents?: boolean;
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
   * Remove all unused shape placeholders from slide.
   */
  cleanupPlaceholders?: boolean;
  /**
   * statusTracker will be triggered on each appended slide.
   * You can e.g. attach a custom callback to a progress bar.
   */
  statusTracker?: StatusTracker['next'];
  /**
   * Set logging verbosity.
   * 0: no logging at all
   * 1: show warnings
   * 2: show info (e.g. on import & append)
   */
  verbosity?: Logger['verbosity'];
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
export type PresentationInfo = {
  templateByName: (tplName: string) => TemplateInfo;
  slidesByTemplate: (tplName: string) => SlideInfo[];
  slideByNumber: (tplName: string, slideNumber: number) => SlideInfo;
  elementByName: (
    tplName: string,
    slideNumber: number,
    elementName: string,
  ) => ElementInfo;
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
  filenameExt?: ContentTypeExtension;
  filenameBase?: string;
  getCreatedContent?: () => TrackedRelationInfo;
  getRelatedContent?: () => Promise<Target>;
  getTargetValue?: () => string;
  updateTargetValue?: (newTarget: string) => void;
  updateTargetIndex?: (newIndex: number) => void;
  updateId?: (newId: string) => void;
  relatedContent?: Target;
  copiedTarget?: string;
};
export type FileInfo = {
  base: string;
  extension: string;
  dir: string;
  isDir: boolean;
};
export type MediaFile = {
  file: string;
  directory: string;
  filepath: string;
  prefix?: string;
  extension: ContentTypeExtension;
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

export interface SupportedPptxGenJSSlide {
  /**
   * Add chart to Slide
   * @param {CHART_NAME|IChartMulti[]} type - chart type
   * @param {object[]} data - data object
   * @param {IChartOpts} options - chart options
   * @return {Slide} this Slide
   * @type {Function}
   */
  addChart(
    type: PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[],
    data: any[],
    options?: PptxGenJS.IChartOpts,
  ): void;

  /**
   * Add image to Slide
   * @param {ImageProps} options - image options
   * @return {Slide} this Slide
   */
  addImage(options: PptxGenJS.ImageProps): void;

  /**
   * Add shape to Slide
   * @param {SHAPE_NAME} shapeName - shape name
   * @param {ShapeProps} options - shape options
   * @return {Slide} this Slide
   */
  addShape(
    shapeName: PptxGenJS.SHAPE_NAME,
    options?: PptxGenJS.ShapeProps,
  ): void;

  /**
   * Add table to Slide
   * @param {TableRow[]} tableRows - table rows
   * @param {TableProps} options - table options
   * @return {Slide} this Slide
   */
  addTable(
    tableRows: PptxGenJS.TableRow[],
    options?: PptxGenJS.TableProps,
  ): void;

  /**
   * Add text to Slide
   * @param {string|TextProps[]} text - text string or complex object
   * @param {TextPropsOptions} options - text options
   * @return {Slide} this Slide
   */
  addText(
    text: string | PptxGenJS.TextProps[],
    options?: PptxGenJS.TextPropsOptions,
  ): void;
}

export type GenerateOnSlideCallback = (
  pptxGenJSSlide: SupportedPptxGenJSSlide,
  pptxGenJS: PptxGenJS,
) => void;
export type GenerateElements = {
  objectName?: string;
  tmpSlideNumber?: number;
  callback?: GenerateOnSlideCallback;
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
  sourceElement?: XmlElement;
};
export type AnalyzedElementType = {
  type: ElementType;
  target?: Target;
  element?: XmlElement;
};
export type ElementOnSlide = {
  sourceElement: XmlElement;
  selector: string;
  mode?: 'findByElementCreationId' | 'findByElementName';
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
