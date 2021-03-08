export type RelationshipAttribute = {
  Id: string
  Type: string
  Target: string
}

export type SlideListAttribute = {
  id: Function
  'r:id': string
}

export type OverrideAttribute = {
  PartName: string
  ContentType: string
}

export type XMLElement = {
  archive: any
  parent: any
  file: string
  tag: string
  attributes: OverrideAttribute | SlideListAttribute | RelationshipAttribute
}