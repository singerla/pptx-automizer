import {
  ISlide,
  ITemplate, IChart,
  PresTemplate, RootPresTemplate, IImage, ICounter
} from './types'

import FileHelper from './helper/file'
import JSZip from 'jszip'
import CountHelper from './helper/count'

class Template implements ITemplate {
  /**
   * Path to local file
   * @type string
   */
  location: string

  /**
   * An alias name to identify template and simplify
   * @type string
   */
  name: string

  /**
   * Node file buffer
   * @type Promise<Buffer>
   */
  file: Promise<Buffer>

  /**
   * this.file will be passed to JSZip
   * @type Promise<JSZip>
   */
  archive: Promise<JSZip>

  /**
   * Array containing all slides coming from Automizer.addSlide()
   * @type: ISlide[]
   */
  slides: ISlide[]

  /**
   * Array containing all counters
   * @type: ICounter[]
   */
  counter: ICounter[]


  constructor(location: string, name?: string) {
    this.location = location
    this.file = FileHelper.readFile(location)
    this.archive = FileHelper.extractFileContent(this.file)
  }

  static import(location: string, name?:string): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate

    if(name) {
      newTemplate = <PresTemplate> new Template(location, name)
      newTemplate.name = name
    } else {
      newTemplate = <RootPresTemplate> new Template(location)
      newTemplate.slides = []
      newTemplate.counter = [
        new CountHelper('slides', newTemplate),
        new CountHelper('charts', newTemplate),
        new CountHelper('images', newTemplate)
      ]
      newTemplate.counter.forEach(async counter => await counter.set())
    }

    return newTemplate
  }

  async appendSlide(slide: ISlide): Promise<void> {
    this.incrementCounter('slides')
    slide.setTarget(await this.archive, this)
    await slide.append()
  }

  async appendChart(shape: IChart): Promise<void> {
    shape.setTarget(await this.archive, this.count('charts'))
    await shape.append()
  }

  async appendImage(shape: IImage): Promise<void> {
    shape.setTarget(await this.archive, this.count('images'))
    await shape.append()
  }

  incrementCounter(name: string): number {
    return CountHelper.increment(name, this.counter)
  }

  count(name: string): number {
    return CountHelper.count(name, this.counter)
  }
}


export default Template