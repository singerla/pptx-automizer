import IArchive from '../interfaces/iarchive';

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
};

export type ElementType = 'sp' | 'chart' | 'table' | 'pic' | 'cxnSp';

export type ElementInfo = {
  name: string;
  type: ElementType;
  id: string;
  position: {
    x: number;
    y: number;
    cx: number;
    cy: number;
  };
  hasTextBody: boolean;
  getXmlElement: () => XmlElement;
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
