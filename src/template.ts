import {
  ISlide,
  ITemplate, IChart,
  PresTemplate, RootPresTemplate
} from './types/interfaces'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import JSZip from 'jszip'

class Template implements ITemplate {
  location: string
  file: Promise<Buffer>
  archive: Promise<JSZip>
  name: string
  slides: ISlide[]
  slideCount: number

  constructor(location: string) {
    this.location = location
    this.file = FileHelper.readFile(location)
    this.archive = FileHelper.extractFileContent(this.file)
    this.slides = []
    this.slideCount = 0
  }

  static importRoot(location: string): RootPresTemplate {
    let newTemplate = new Template(location)
    newTemplate.countSlides()

    return newTemplate
  }

  static import(location: string, name?:string): PresTemplate {
    let newTemplate = new Template(location)

    if(name) {
      newTemplate.name = name
    }

    return newTemplate
  }

  async appendSlide(slide: ISlide): Promise<void> {
    this.incrementSlideCounter()

    slide.setTarget(await this.archive, this)
    await slide.append()
  }

  async appendShape(shape: IChart): Promise<void> {
    shape.setTarget(await this.archive, shape.sourceNumber)
    await shape.append()
  }

  async countSlides(): Promise<number> {
    this.slideCount = await XmlHelper.countSlides(await this.archive)

    return this.slideCount
  }

  incrementSlideCounter(): number {
    this.slideCount ++

    return this.slideCount;
  }
}


export default Template