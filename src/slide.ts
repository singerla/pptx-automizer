import JSZip from 'jszip'
import Chart from './chart'
import Image from './image'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import { 
  ISlide, RootPresTemplate, PresTemplate,
  RelationshipAttribute, SlideListAttribute 
} from './types'

export default class Slide implements ISlide {
  sourceTemplate: PresTemplate
  targetTemplate: RootPresTemplate
  targetNumber: number
  sourceNumber: number
  targetArchive: JSZip
  sourceArchive: JSZip
  sourcePath: string
  targetPath: string
  modifications: Function[]
  relsPath: string

  constructor(params: any) {
    this.sourceTemplate = params.template
    this.sourceNumber = params.number
    this.modifications = []
  }

  modify(callback: Function): void {
    this.modifications.push(callback)
  }

  setTarget(archive: JSZip, targetTemplate: RootPresTemplate) {
    this.targetTemplate = targetTemplate
    this.targetArchive = archive
    this.targetNumber = targetTemplate.slideCount

    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`
    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`
  }
  
  async append() {
    this.sourceArchive = await this.sourceTemplate.archive
    
    await this.applyModifications()
    await this.copySlideFiles()
    await this.copyRelatedContent()
    await this.addSlideToPresentation()
  }

  async applyModifications(): Promise<void> {
    for(let m in this.modifications) {
      let xml = await XmlHelper.getXmlFromArchive(this.sourceArchive, this.sourcePath)
      this.modifications[m](xml)
      await XmlHelper.writeXmlToArchive(this.sourceArchive, this.sourcePath, xml)
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

  async addSlideToPresentation(): Promise<void> {
    let relId = await XmlHelper.getNextRelId(this.targetArchive, 'ppt/_rels/presentation.xml.rels')
    await this.appendToSlideRel(this.targetArchive, relId, this.targetNumber),
    await this.appendToSlideList(this.targetArchive, relId),
    await this.appendToContentType(this.targetArchive, this.targetNumber)
  }

  appendToSlideRel(rootArchive: JSZip, relId: string, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/_rels/presentation.xml.rels`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: <RelationshipAttribute> {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`,
        Target: `slides/slide${slideCount}.xml`
      }
    })
  }

  appendToSlideList(rootArchive: JSZip, relId: string): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('p:sldIdLst')[0],
      tag: 'p:sldId',
      attributes: <SlideListAttribute> {
        id: (xml: HTMLElement) => XmlHelper.getMaxId(xml.getElementsByTagName('p:sldId'), 'id', true),
        'r:id': relId
      }
    })
  }

  appendToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
      })
    )
  }

  async copyRelatedContent(): Promise<void> {
    let charts = await XmlHelper.getTargetsFromRelationships(this.sourceArchive, this.relsPath, '../charts/chart')
    for(let i in charts) {
      let newChart = new Chart(charts[i], this.sourceArchive, this.targetNumber)
      this.targetTemplate.incrementChartCounter()
      await this.targetTemplate.appendChart(newChart)
    }

    let images = await XmlHelper.getTargetsByRelationshipType(this.sourceArchive, this.relsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    for(let i in images) {
      let newImage = new Image(images[i], this.sourceArchive, this.targetNumber)
      this.targetTemplate.incrementImageCounter()
      await this.targetTemplate.appendImage(newImage)
    }
  }
}
