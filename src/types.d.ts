import JSZip from "jszip";


type AutomizerParams = {
	templateDir?: string
	outputDir?: string
}

type AutomizerSummary = {
	status: string
	file: string
	templates: number
	slides: number
}

export interface ISlide {
  append(): Promise<void>
  setTarget(archive: JSZip, targetTemplate: RootPresTemplate): void
  addElement(presName: string, slideNumber: number, selector: Function | string): void
	sourceArchive: JSZip
	sourceNumber: number
	modifications: Function[]
	toAppend: any[]
	modify: Function
}

export interface IPresentationProps {
	rootTemplate: RootPresTemplate
	templates: PresTemplate[]
	template(name: string): PresTemplate
	params: AutomizerParams
}

export interface ITemplate {
  slideCount: number;
	location: string
	file: Promise<Buffer>
	archive: Promise<JSZip>
}

export interface RootPresTemplate extends ITemplate {
	countSlides(): Promise<number>
  countImages(): Promise<number>
	countCharts(): Promise<number>
	incrementSlideCounter(): number
  incrementChartCounter(): number
  incrementImageCounter(): number
  appendSlide(slide: ISlide): Promise<void>
	appendChart: Function
	appendImage: Function
	slideCount: number
  chartCount: number
  imageCount: number
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

export interface IImage extends IShape {
  append: Function
  setTarget: Function
	sourceFile: string
	targetFile: string
	contentTypeMap: any
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