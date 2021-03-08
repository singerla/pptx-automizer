import {
  PresSlide,
  ITemplate,
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
  slides: PresSlide[]
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

    return newTemplate
  }

  static import(location: string, name?:string): PresTemplate {
    let newTemplate = new Template(location)

    if(name) {
      newTemplate.name = name
    }

    return newTemplate
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