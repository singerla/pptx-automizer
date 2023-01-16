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

export type ElementInfo = {
  name: string;
  type: string;
  id: string;
  position: {
    x: number;
    y: number;
    cx: number;
    cy: number;
  };
};

export type ModifyXmlCallback = (
  xml: XmlDocument | XmlElement,
  index?: number,
  archive?: IArchive,
) => void;
