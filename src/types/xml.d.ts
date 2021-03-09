export type DefaultAttribute = {
  Extension: string
  ContentType: string
}

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
  clause?: Function
  parent: any
  file: string
  tag: string
  attributes: DefaultAttribute | OverrideAttribute | SlideListAttribute | RelationshipAttribute
}

export type Target = {
  file: string
  number: number
}