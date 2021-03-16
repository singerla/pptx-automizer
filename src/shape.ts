import JSZip from "jszip"
import XmlHelper from "./helper/xml"
import { RootPresTemplate, Target } from "./types/app"

export default class Shape {
  sourceArchive: JSZip
  sourceNumber: number
  sourceFile: string
  sourceRid: string
  targetFile: string
  targetArchive: JSZip
  targetTemplate: RootPresTemplate
  targetSlideNumber: number
  contentTypeMap: any
  targetNumber: number

  createdRid: string
  sourceSlideNumber: number
  sourceSlideFile: string
  targetSlideFile: string
  targetSlideRelFile: string

  relRootTag: string
  relAttribute: string
  relParent: (element: HTMLElement) => HTMLElement

  constructor(relsXmlInfo: Target, sourceArchive: JSZip, sourceSlideNumber?:number) {
    this.sourceNumber = relsXmlInfo.number
    this.sourceRid = relsXmlInfo.rId
    this.sourceArchive = sourceArchive
    this.sourceSlideNumber = sourceSlideNumber
    this.sourceSlideFile = `ppt/slides/slide${this.sourceSlideNumber}.xml`
  }

  async setTarget(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<void> {
    this.targetTemplate = targetTemplate
    this.targetArchive = await this.targetTemplate.archive
    this.targetSlideNumber = targetSlideNumber
    this.targetSlideFile = `ppt/slides/slide${this.targetSlideNumber}.xml`
    this.targetSlideRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`
  }

  async appendToSlideTree(): Promise<void> {
    let sourceSlideXml = await XmlHelper.getXmlFromArchive(this.sourceArchive, this.sourceSlideFile)
    let sourceElement = <HTMLElement> await this.getElementByRid(sourceSlideXml, this.sourceRid)
    let targetElement = sourceElement.cloneNode(true)
    
    let targetSlideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile)
    targetSlideXml.getElementsByTagName('p:spTree')[0].appendChild(targetElement)

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml)
  }

  async getElementByRid(slideXml: Document, rId: string): Promise<HTMLElement> {
    let sourceChartList = slideXml.getElementsByTagName('p:spTree')[0].getElementsByTagName(this.relRootTag)
    let sourceElement = XmlHelper.findByAttributeValue(sourceChartList, this.relAttribute, rId)

    return this.relParent(sourceElement)
  }

  async updateElementRelId() {
    let targetSlideXml = await XmlHelper.getXmlFromArchive(this.targetArchive, this.targetSlideFile)
    let targetElement = await this.getElementByRid(targetSlideXml, this.sourceRid)
    targetElement.getElementsByTagName(this.relRootTag)[0].setAttribute(this.relAttribute, this.createdRid)

    await XmlHelper.writeXmlToArchive(this.targetArchive, this.targetSlideFile, targetSlideXml)
  }
}