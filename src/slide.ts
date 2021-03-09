import JSZip from 'jszip'
import Chart from './chart'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import {
	ISlide, ITemplate, PresTemplate
} from './types/interfaces'


export default class Slide implements ISlide {
  modifications: Function[]
  template: PresTemplate
  sourceNumber: number
  targetArchive: any
  targetNumber: number
  slidePath: string
  sourceArchive: JSZip
  relsPath: string
  targetTemplate: any

  constructor(params: any) {
    this.template = params.template
    this.sourceNumber = params.number
    this.modifications = []
  }

  modify(callback: Function): void {
    this.modifications.push(callback)
  }

  setTarget(archive: JSZip, targetTemplate: ITemplate) {
    this.targetTemplate = targetTemplate
    this.targetArchive = archive
    this.targetNumber = targetTemplate.slideCount

    this.slidePath = `ppt/slides/slide${this.targetNumber}.xml`
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`
  }
  
  async append() {
    this.sourceArchive = await this.template.archive
    
    await this.applyModifications()
    await this.copySlideFiles()
    await this.copyRelatedContent()
    await this.addSlideToPresentation()
  }

  async applyModifications(): Promise<void> {
    for(let m in this.modifications) {
      let xml = await XmlHelper.getXmlFromArchive(this.sourceArchive, this.slidePath)
      this.modifications[m](xml)
      await XmlHelper.writeXmlToArchive(this.sourceArchive, this.slidePath, xml)
    }
  }

  async copySlideFiles(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/slides/slide${this.sourceNumber}.xml`, 
      this.targetArchive, `ppt/slides/slide${this.targetNumber}.xml`
    )

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`, 
      this.targetArchive, `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`
    )
  }

  async addSlideToPresentation(): Promise<HTMLElement[]> {
    let relId = await XmlHelper.getNextRelId(this.targetArchive, 'ppt/_rels/presentation.xml.rels')
    let promises = [
      XmlHelper.appendToSlideRel(this.targetArchive, relId, this.targetNumber),
      XmlHelper.appendToSlideList(this.targetArchive, relId),
      XmlHelper.appendToContentType(this.targetArchive, this.targetNumber)
    ]

    return Promise.all(promises)
  }
  
  async copyRelatedContent(): Promise<void> {
    let charts = await XmlHelper.getTargetsFromRelationships(this.sourceArchive, this.relsPath, '../charts/chart')
    
    for(let i in charts) {
      let newChart = new Chart(charts[i], this.sourceArchive)
      
      await this.targetTemplate.appendShape(newChart)
    }
  }


}
