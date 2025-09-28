import IArchive from '../interfaces/iarchive';
import { TableData, TableInfo } from './table-types';

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
  layoutPlaceholders: PlaceholderInfo[]
};

export type ElementType =
  | 'sp'
  | 'chart'
  | 'chartEx'
  | 'table'
  | 'pic'
  | 'cxnSp';

export type PlaceholderInfo = {
  type: string;
  sz: string;
  idx: number;
};

export type ElementInfo = {
  name: string;
  type: ElementType;
  id: string;
  creationId: string;
  position: {
    x: number;
    y: number;
    cx: number;
    cy: number;
    rot?: number;
  };
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
