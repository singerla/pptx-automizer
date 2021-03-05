import {
	PresSlide, PresTemplate
} from './types/interfaces'


export default class Slide implements PresSlide {
  public modifications: Function[]
  public template: PresTemplate
  public number: number

  constructor(params) {
    this.template = params.template
    this.number = params.number
    this.modifications = []
  }

  modify(callback) {
    this.modifications.push(callback)
  }

  public addChart(name: string, worksheetNumber: number): this {
    // let template = this.presentation.template(name)
    
    // template.worksheet = template.archive
    //     .then(Helper.getWorksheet(worksheetNumber))
    //     .then(Helper.extractFileContent)
    //     .then(Helper.extractWorkbook)
    //     .then(Helper.parseXmlDocument)

    return this
  }
}
