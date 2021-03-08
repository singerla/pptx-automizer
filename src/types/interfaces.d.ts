import JSZip from "jszip";

export interface PresSlide {
	template: PresTemplate
	number: number
	modifications: Function[]
	addChart: Function
}

export interface IPresentationProps {
	rootTemplate: RootPresTemplate
	templates: PresTemplate[]
}

export interface ITemplate {
	location: string
	file: Promise<Buffer>
	archive: Promise<JSZip>
	countSlides: Function
	incrementSlideCounter: Function
}

export interface RootPresTemplate extends ITemplate {
	slideCount: number
	slides: PresSlide[]
}

export interface PresTemplate extends ITemplate {
	name: string
}
