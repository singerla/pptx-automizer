import JSZip from "jszip";

export interface ISlide {
  append(): Promise<void>
  setTarget(archive: JSZip, targetTemplate: RootPresTemplate): void
	sourceArchive: JSZip
	sourceNumber: number
	modifications: Function[]
	modify: Function
}

export interface IPresentationProps {
	rootTemplate: RootPresTemplate
	templates: PresTemplate[]
}

export interface ITemplate {
  slideCount: number;
	location: string
	file: Promise<Buffer>
	archive: Promise<JSZip>
	countSlides: Function
	incrementSlideCounter: Function
}

export interface RootPresTemplate extends ITemplate {
  countCharts();
  appendSlide(slide: ISlide): Promise<void>
	appendShape: Function
	slideCount: number
	slides: ISlide[]
}

export interface PresTemplate extends ITemplate {
	name: string
}


export interface IShape {
	sourceArchive: JSZip
	targetArchive: JSZip
}

export interface IChart extends IShape {
  append: Function
  setTarget: Function
	sourceNumber: number
	targetNumber: number
}

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
	rId?: string
  }