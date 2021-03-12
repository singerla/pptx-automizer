import JSZip from 'jszip'
import Chart from './chart'
import Image from './image'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import { 
  ISlide, RootPresTemplate, PresTemplate,
  RelationshipAttribute, SlideListAttribute, IPresentationProps 
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
  appendElements: any[]
  appendRelations: any[]
  relsPath: string
  rootTemplate: RootPresTemplate
  root: IPresentationProps
  targetRelsPath: string

  constructor(params: any) {
    this.sourceTemplate = params.template
    this.sourceNumber = params.number
    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`

    this.modifications = []
    this.appendElements = []
    this.appendRelations = []
  }

  modify(callback: Function): void {
    this.modifications.push(callback)
  }

  async addElement(presName: string, slideNumber: number, selector: string, callback?: Function | Function[]): Promise<this> {
    let template = this.root.template(presName)
    let sourcePath = `ppt/slides/slide${slideNumber}.xml`
    let sourceArchive = await template.archive
    let sourceElement = await XmlHelper.findByElementName(sourceArchive, sourcePath, selector)
    
    if(!sourceElement) {
      throw new Error(`Can't find ${selector} on slide ${slideNumber} in ${presName}`)
    }

    let appendElement = sourceElement.cloneNode(true)

    this.analyzeElement(appendElement, sourceArchive)

    if(callback !== undefined) {
      if(callback instanceof Array) {
        callback.forEach(cb => cb(appendElement))
      } else {
        callback(appendElement)
      }
    }
    
    this.appendElements.push(appendElement)

    return this
  }

  async setTarget(targetTemplate: RootPresTemplate): Promise<void>{
    this.targetTemplate = targetTemplate
    this.targetArchive = await targetTemplate.archive
    this.targetNumber = targetTemplate.count('slides')

    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`
    this.targetRelsPath = `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`
  }
  
  async append() {
    this.sourceArchive = await this.sourceTemplate.archive
    
    await this.copySlideFiles()
    await this.copyRelatedContent()
    await this.addSlideToPresentation()

    if(this.appendRelations.length) {
      await this.appendImportedRelations()
    }

    if(this.appendElements.length) {
      await this.appendImportedElements()
    }

    await this.applyModifications()
  }

  async analyzeElement(appendElement, sourceArchive) {
    let isChart = appendElement.getElementsByTagName('c:chart')
    if(isChart.length) {
      let sourceRid = isChart[0].getAttribute('r:id')
      let chartRels = await XmlHelper.getTargetsFromRelationships(sourceArchive, this.relsPath, '../charts/chart')
      this.appendRelations.push({
        relation: chartRels.find(rel => rel.rId === sourceRid),
        sourceArchive: sourceArchive
      })
    }
  }

  async appendImportedElements() {
    let slideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetPath)
    let tree = slideXml.getElementsByTagName('p:spTree')[0]

    this.appendElements.forEach(element => {
      tree.appendChild(element)
    })

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, slideXml)
  }

  async appendImportedRelations() {
    for(let i in this.appendRelations) {
      let relation = this.appendRelations[i].relation
      let sourceArchive = this.appendRelations[i].sourceArchive

      let newChart = new Chart(relation, sourceArchive, this.targetNumber, true)
      this.targetTemplate.incrementCounter('charts')
      await this.targetTemplate.appendChart(newChart)

      this.modifications.push(chart => {
        chart.getElementsByTagName('c:chart')[0].setAttribute('r:id', newChart.createdRid)
      })
    }
  }

  async applyModifications(): Promise<void> {
    for(let m in this.modifications) {
      let xml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetPath)
      this.modifications[m](xml)
      await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, xml)
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
      this.targetTemplate.incrementCounter('charts')
      await this.targetTemplate.appendChart(newChart)
    }

    let images = await XmlHelper.getTargetsByRelationshipType(this.sourceArchive, this.relsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    for(let i in images) {
      let newImage = new Image(images[i], this.sourceArchive, this.targetNumber)
      this.targetTemplate.incrementCounter('images')
      await this.targetTemplate.appendImage(newImage)
    }
  }
}
