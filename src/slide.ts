import JSZip from 'jszip'
import Chart from './chart'
import Image from './image'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import { 
  ISlide, RootPresTemplate, PresTemplate,
  RelationshipAttribute, SlideListAttribute, IPresentationProps, ImportedElement 
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
  appendElements: ImportedElement[]
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

    if(this.appendElements.length) {
      await this.appendImportedElements()
    }

    await this.applyModifications()
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

    let appendElement = await this.analyzeElement(sourceElement, sourceArchive, slideNumber)

    this.appendElements.push( <ImportedElement>{
      sourceArchive: sourceArchive,
      type: appendElement.type,
      target: appendElement.target,
      element: sourceElement.cloneNode(true),
      callback: callback
    })

    return this
  }

  async analyzeElement(appendElement: any, sourceArchive: JSZip, slideNumber: number): Promise<any> {
    let isChart = appendElement.getElementsByTagName('c:chart')
    let relsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`
    
    if(isChart.length) {
      let sourceRid = isChart[0].getAttribute('r:id')
      let chartRels = await XmlHelper.getTargetsFromRelationships(sourceArchive, relsPath, '../charts/chart')
      
      return {
        type: 'chart',
        target: chartRels.find(rel => rel.rId === sourceRid),
      }
    }

    let isImage = appendElement.getElementsByTagName('p:nvPicPr')
    if(isImage.length) {
      let sourceRid = appendElement.getElementsByTagName('a:blip')[0].getAttribute('r:embed')
      let imageRels = await XmlHelper.getTargetsFromRelationships(sourceArchive, relsPath, '../media/image', /\..+?$/)
      
      return {
        type: 'image',
        target: imageRels.find(rel => rel.rId === sourceRid),
      }
    }
  }

  async appendImportedElements(): Promise<void> {
    let slideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetPath)
    let tree = slideXml.getElementsByTagName('p:spTree')[0]

    for(let i in this.appendElements) {
      let info = this.appendElements[i]

      let element = info.element
      let callbacks = this.arrayify(info.callback)

      switch(info.type) {
        case 'chart' :
          let newChart = await this.appendChart(info)
          
          callbacks.push(element => {
            element.getElementsByTagName('c:chart')[0].setAttribute('r:id', newChart.createdRid)
          })
        break
        case 'image' :
          let newImage = await this.appendImage(info)
          callbacks.push(element => {
            element.getElementsByTagName('a:blip')[0].setAttribute('r:embed', newImage.createdRid)
          })
        break
      }

      callbacks.forEach(callback => callback(element))
      tree.appendChild(element)
    }

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetPath, slideXml)
  }

  async appendChart(relationInfo: any): Promise<Chart> {
    let target = relationInfo.target
    let sourceArchive = relationInfo.sourceArchive
    
    let newChart = new Chart(target, sourceArchive)
    await newChart.append(this.targetTemplate, this.targetNumber, true)

    return newChart
  }

  async appendImage(relationInfo: any): Promise<Image> {
    let target = relationInfo.target
    let sourceArchive = relationInfo.sourceArchive
    
    let newImage = new Image(target, sourceArchive)
    await newImage.append(this.targetTemplate, this.targetNumber, true)

    return newImage
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
      let newChart = new Chart(charts[i], this.sourceArchive)
      await newChart.append(this.targetTemplate, this.targetNumber)
    }

    let images = await XmlHelper.getTargetsByRelationshipType(this.sourceArchive, this.relsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    for(let i in images) {
      let newImage = new Image(images[i], this.sourceArchive)
      await newImage.append(this.targetTemplate, this.targetNumber)
    }
  }


  arrayify(s) {
    if(s instanceof Array) {
      return s
    } else if(s !== undefined) {
      return [s]
    } else {
      return []
    }
  }

}
