import JSZip from 'jszip';
import { FileProxy } from '../helper/file-proxy';

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
  archive: FileProxy;
  assert?: (xml: XMLDocument) => void;
  clause?: (xml: XMLDocument, element?: Element) => boolean;
  parent?: (xml: XMLDocument) => Element;
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
  xml: XMLDocument | Element,
  index?: number,
  archive?: FileProxy,
) => void;
