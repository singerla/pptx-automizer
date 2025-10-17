import IArchive from '../interfaces/iarchive';
import { TableInfo } from './table-types';

export type DefaultAttribute = {
  Extension: string;
  ContentType: string;
};

export type RelationshipAttribute = {
  Id: string;
  Type: string;
  Target: string;
};

export type SlideListAttribute = {
  id: (xml: XmlDocument) => number;
  'r:id': string;
};

export type OverrideAttribute = {
  PartName: string;
  ContentType: string;
};

export type XmlDocument = XMLDocument;
export type XmlElement = Element;

export type HelperElement = {
  archive: IArchive;
  assert?: (xml: XmlDocument) => void;
  clause?: (xml: XmlDocument, element?: XmlElement) => boolean;
  parent?: (xml: XmlDocument) => XmlElement;
  file: string;
  tag: string;
  attributes?:
    | DefaultAttribute
    | OverrideAttribute
    | SlideListAttribute
    | RelationshipAttribute;
};

export type TemplateInfo = {
  name: string;
  slides: SlideInfo[];
};

export type SlideInfo = {
  id: number;
  number: number;
  info: TemplateSlideInfo;
  elements: ElementInfo[];
};

export type TemplateSlideInfo = {
  name: string;
  layoutName: string;
  layoutPlaceholders: PlaceholderInfo[];
};

export type ElementType =
  | 'sp'
  | 'chart'
  | 'chartEx'
  | 'table'
  | 'pic'
  | 'cxnSp';

export type ElementVisualType =
  | 'textField'
  | 'chart'
  | 'smartArt'
  | 'diagram'
  | 'graphicFrame'
  | 'table'
  | '3dModel'
  | 'pictogram'
  | 'picture'
  | 'svgImage'
  | 'imageFilledShape'
  | 'vectorShape'
  | 'vectorLine'
  | 'unknown';

/**
 * body - Body text placeholder
 * title - Title placeholder
 * ctrTitle - Center title placeholder
 * subTitle - Subtitle placeholder
 * dt - Date placeholder
 * ftr - Footer placeholder
 * sldNum - Slide number placeholder
 * pic - Picture placeholder
 * chart - Chart placeholder
 * tbl - Table placeholder
 * media - Media clip placeholder
 * obj - Object placeholder
 * dgm - Diagram placeholder
 * clipArt - Clip art placeholder
 * orgChart - Organization chart placeholder
 * bitmap - Bitmap placeholder
 * hdr - Header placeholder
 * unknown - For cases where the type cannot be determined
 */
export type PlaceholderType =
  | 'body'
  | 'title'
  | 'ctrTitle'
  | 'subTitle'
  | 'dt'
  | 'ftr'
  | 'sldNum'
  | 'pic'
  | 'chart'
  | 'tbl'
  | 'media'
  | 'obj'
  | 'dgm'
  | 'clipArt'
  | 'orgChart'
  | 'bitmap'
  | 'hdr'
  | 'unknown';

export type PlaceholderInfo = {
  type: PlaceholderType;
  sz: string;
  idx: number;
  elementType?: ElementType;
  // We can use the ph layout position as fallback.
  // Some ph shapes on a slide do not have coordinates, but inherit
  // them from the placeholder.
  position?: ElementPosition;
};

export type ElementInfo = {
  name: string;
  type: ElementType;
  id: string;
  creationId: string;
  nameIdx: number;
  position: ElementPosition;
  placeholder: PlaceholderInfo;
  hasTextBody: boolean;
  getText: () => string[];
  getParagraphs: () => TextParagraph[];
  getParagraphGroups: () => TextParagraphGroup[];
  getAltText: () => string;
  getTableInfo: () => TableInfo[];
  getXmlElement: () => XmlElement;
  getGroupInfo: () => GroupInfo;
};

export type LayoutInfo = {
  layoutName: string;
  placeholders: PlaceholderInfo[];
};

/**
 * Result object for tracking placeholder mapping operations
 */
export type PlaceholderMappingResult = {
  /** Array of placeholders that have been assigned to elements */
  usedPlaceholders: PlaceholderInfo[];
  /** All source elements that could be assigned to a target placeholder */
  matchedSourceElements: ElementInfo[];
  /** Array of source placeholder elements that couldn't be matched to target placeholders */
  unmatchedSourcePlaceholderElements: ElementInfo[];
};

export type ElementPosition = {
  x: number;
  y: number;
  cx: number;
  cy: number;
  rot?: number;
};

export type SlideHelperProps = {
  sourceArchive: IArchive;
  slideNumber: number;
  sourceLayoutId?: number;
};

export type GroupInfo = {
  isChild: boolean;
  isParent: boolean;
  getParent: () => XmlElement;
  getChildren: () => XmlElement[];
};

export type TextParagraphProps = {
  isBold?: boolean;
  isItalic?: boolean;
  isUnderlined?: boolean;
  fontSize?: number;
  align?: 'l' | 'ctr' | 'r';
  indent?: number;
  bullet?: string;
  isNumbered?: boolean;
  numberingType?: string;
  startAt?: string;
};

export type TextParagraph = {
  texts: string[];
} & TextParagraphProps;

export type TextParagraphGroup = {
  properties: TextParagraphProps;
  texts: string[];
};

export type ContentMapType = 'slideMaster' | 'slideLayout';
export type ContentMap = {
  type: ContentMapType;
  key: string;
  sourceId: number;
  targetId: number;
  name?: string;
  creationId?: string;
};

export type ModifyXmlCallback = (
  xml: XmlDocument | XmlElement,
  index?: number,
  archive?: IArchive,
) => void;
