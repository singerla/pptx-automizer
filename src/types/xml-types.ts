export type DefaultAttribute = {
  Extension: string;
  ContentType: string;
}

export type RelationshipAttribute = {
  Id: string;
  Type: string;
  Target: string;
}

export type SlideListAttribute = {
  id: (xml: HTMLElement) => number;
  'r:id': string;
}

export type OverrideAttribute = {
  PartName: string;
  ContentType: string;
}

export type XMLElement = {
  archive: any;
  clause?: (xml: HTMLElement | Document) => boolean;
  parent: any;
  file: string;
  tag: string;
  attributes: DefaultAttribute | OverrideAttribute | SlideListAttribute | RelationshipAttribute;
}
