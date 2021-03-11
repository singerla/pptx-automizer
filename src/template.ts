import {
  ISlide,
  ITemplate, IChart,
  PresTemplate, RootPresTemplate, IImage
} from './types'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
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
   * Contains actual number of slides in root template.
   * Has to be incremented before a new slide is appended.
   * @type: number
   */
  slideCount: number
  chartCount: number
  imageCount: number

  constructor(location: string, name?: string) {
    this.location = location

    if(name !== undefined) {
      this.name = name
    }

    this.file = FileHelper.readFile(location)
    this.archive = FileHelper.extractFileContent(this.file)
    this.slides = []

    this.slideCount = 0
    this.chartCount = 0
    this.imageCount = 0
  }

  static import(location: string, name?:string): PresTemplate | RootPresTemplate {
    let newTemplate: PresTemplate | RootPresTemplate

    if(name) {
      newTemplate = <PresTemplate> new Template(location, name)
    } else {
      newTemplate = <RootPresTemplate> new Template(location)
    }

    return newTemplate
  }

  async appendSlide(slide: ISlide): Promise<void> {
    this.incrementSlideCounter()
    slide.setTarget(await this.archive, this)
    await slide.append()
  }
  
  async countSlides(): Promise<number> {
    let slideCount = await CountHelper.countSlides(await this.archive)
    this.slideCount = slideCount
    return this.slideCount
  }
  
  incrementSlideCounter(): number {
    this.slideCount++
    return this.slideCount;
  }

  async appendChart(shape: IChart): Promise<void> {
    shape.setTarget(await this.archive, this.chartCount)
    await shape.append()
  }

  async countCharts(): Promise<number> {
    this.chartCount = await CountHelper.countCharts(await this.archive)
    return this.chartCount
  }

  incrementChartCounter(): number {
    return ++this.chartCount;
  }

  async appendImage(shape: IImage): Promise<void> {
    shape.setTarget(await this.archive, this.imageCount)
    await shape.append()
  }

  async countImages(): Promise<number> {
    this.imageCount = await CountHelper.countImages(await this.archive)
    return this.imageCount
  }

  incrementImageCounter(): number {
    return ++this.imageCount;
  }
}


export default Template