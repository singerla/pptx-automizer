import JSZip from 'jszip';

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
  id: (xml: XMLDocument) => number;
  'r:id': string;
};

export type OverrideAttribute = {
  PartName: string;
  ContentType: string;
};

export type HelperElement = {
  archive: JSZip;
  clause?: (xml: XMLDocument) => boolean;
  parent: (xml: XMLDocument) => Element;
  file: string;
  tag: string;
  attributes:
    | DefaultAttribute
    | OverrideAttribute
    | SlideListAttribute
    | RelationshipAttribute;
};
