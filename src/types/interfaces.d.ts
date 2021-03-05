
export type FILE_TYPE = 'local' | 'remote' | 'browser'

export interface File {
  type: FILE_TYPE
  location: string
}

export interface ISlideRel {
	Target: string
	fileName?: string
	data: any[] | string
	path?: string
	extn?: string
	globalId?: number
	rId: number
}

export interface SlideBaseProps {

}

export interface PresSlide extends SlideBaseProps {
	template: PresTemplate
	number: number
	modifications: Function[]
	addChart: Function
}

export interface IPresentationProps {
	rootTemplate: PresTemplate
	templates: PresTemplate[]
}

export interface IArchive {
	files: Function[]
	file: Function
	commment: any
	root: string
	clone: Function
	generateAsync: Function
}

export interface PresTemplate {
	location: string
	name: string
	file: any
	archive: IArchive
	worksheet: any
	slides: PresSlide[]
}
