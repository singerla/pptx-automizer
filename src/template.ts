import {
  ISlide,
  ITemplate, IChart,
  PresTemplate, RootPresTemplate
} from './types'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import JSZip from 'jszip'

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
   * Contains actual number of slides in root template.
   * Has to be incremented before a new slide is appended.
   * @type: number
   */
  slideCount: number
  chartCount: number

  constructor(location: string) {
    this.location = location
    this.file = FileHelper.readFile(location)
    this.archive = FileHelper.extractFileContent(this.file)
    this.slides = []

    this.slideCount = 0
    this.chartCount = 0
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

  async appendSlide(slide: ISlide): Promise<void> {
    this.incrementSlideCounter()
    slide.setTarget(await this.archive, this)
    await slide.append()
  }

  async appendShape(shape: IChart): Promise<void> {
    this.incrementChartCounter()
    shape.setTarget(await this.archive, this.chartCount)
    await shape.append()
  }

  async countSlides(): Promise<number> {
    let slideCount = await XmlHelper.countSlides(await this.archive)
    this.slideCount = slideCount
    return this.slideCount
  }

  incrementSlideCounter(): number {
    this.slideCount++
    return this.slideCount;
  }

  async countCharts(): Promise<number> {
    this.chartCount = await XmlHelper.countCharts(await this.archive)
    return this.chartCount
  }

  incrementChartCounter(): number {
    return ++this.chartCount;
  }
}


export default Template