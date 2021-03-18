import JSZip from 'jszip'
import Chart from './shapes/chart'
import Image from './shapes/image'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import { ElementType } from './definitions/enums'
import { ISlide, RootPresTemplate, PresTemplate, IPresentationProps, ImportedElement, AnalyzedElementType, Target } from './definitions/app'
import { RelationshipAttribute, SlideListAttribute } from './definitions/xml'
import Generic from './shapes/generic'

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

  async append(targetTemplate: RootPresTemplate) {
    this.targetTemplate = targetTemplate
    this.targetArchive = await targetTemplate.archive
    this.targetNumber = targetTemplate.incrementCounter('slides')
    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`
    this.targetRelsPath = `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`

    this.sourceArchive = await this.sourceTemplate.archive
    
    await this.copySlideFiles()
    await this.copyRelatedContent()
    await this.addSlideToPresentation()

    if(this.hasNotes()) {
      await this.copySlideNoteFiles()
      await this.updateSlideNoteFile()
      await this.appendNotesToContentType(this.targetArchive, this.targetNumber)
    }
    
    if(this.appendElements.length) {
      await this.appendImportedElements()
    }
    
    await this.applyModifications()
  }

  modify(callback: Function): void {
    this.modifications.push(callback)
  }

  async addSlideToPresentation(): Promise<void> {
    let relId = await XmlHelper.getNextRelId(this.targetArchive, 'ppt/_rels/presentation.xml.rels')
    await this.appendToSlideRel(this.targetArchive, relId, this.targetNumber),
    await this.appendToSlideList(this.targetArchive, relId),
    await this.appendSlideToContentType(this.targetArchive, this.targetNumber)
  }

  async addElement(presName: string, slideNumber: number, selector: string, callback?: Function | Function[]): Promise<this> {
    let template = this.root.template(presName)
    let sourcePath = `ppt/slides/slide${slideNumber}.xml`
    let sourceArchive = await template.archive
    let sourceElement = await XmlHelper.findByElementName(sourceArchive, sourcePath, selector)
    
    if(!sourceElement) {
      throw new Error(`Can't find ${selector} on slide ${slideNumber} in ${presName}`)
    }

    let appendElementParams = await this.analyzeElement(sourceElement, sourceArchive, slideNumber)

    this.appendElements.push(<ImportedElement>{
      sourceArchive: sourceArchive,
      sourceSlideNumber: slideNumber,
      callback: callback,
      type: appendElementParams.type,
      target: appendElementParams.target,
      element: appendElementParams.element
    })

    return this
  }

  async analyzeElement(appendElement: any, sourceArchive: JSZip, slideNumber: number): Promise<AnalyzedElementType> {
    let isChart = appendElement.getElementsByTagName('c:chart')
    if(isChart.length) {
      return <AnalyzedElementType> {
        type: ElementType.Chart,
        target: await XmlHelper.getTargetByRelId(sourceArchive, slideNumber, appendElement, 'chart'),
      }
    }

    let isImage = appendElement.getElementsByTagName('p:nvPicPr')
    if(isImage.length) {
      return <AnalyzedElementType> {
        type: ElementType.Image,
        target: await XmlHelper.getTargetByRelId(sourceArchive, slideNumber, appendElement, 'image')
      }
    }

    return <AnalyzedElementType> {
      type: ElementType.Shape,
      element: appendElement
    }
  }

  async appendImportedElements(): Promise<void> {
    for(let i in this.appendElements) {
      let info = this.appendElements[i]

      switch(info.type) {
        case ElementType.Chart:
          await new Chart(info)
            .append(this.targetTemplate, this.targetNumber, true)
        break
        case ElementType.Image:
          await new Image(info)
            .append(this.targetTemplate, this.targetNumber, true)
        break
        case ElementType.Shape:
          await new Generic(info)
            .append(this.targetTemplate, this.targetNumber)
        break
      }
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

  async copySlideNoteFiles(): Promise<void> {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/notesSlides/notesSlide${this.sourceNumber}.xml`, 
      this.targetArchive, `ppt/notesSlides/notesSlide${this.targetNumber}.xml`
    )

    FileHelper.zipCopy(
      this.sourceArchive, `ppt/notesSlides/_rels/notesSlide${this.sourceNumber}.xml.rels`, 
      this.targetArchive, `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`
    )
  }

  async updateSlideNoteFile(): Promise<void> {
    await XmlHelper.replaceAttribute(
      this.targetArchive, 
      `ppt/notesSlides/_rels/notesSlide${this.targetNumber}.xml.rels`, "Relationship", "Target", 
      `../slides/slide${this.sourceNumber}.xml`,
      `../slides/slide${this.targetNumber}.xml`
    )

    await XmlHelper.replaceAttribute(
      this.targetArchive, 
      `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`, "Relationship", "Target", 
      `../notesSlides/notesSlide${this.sourceNumber}.xml`,
      `../notesSlides/notesSlide${this.targetNumber}.xml`
    )
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

  appendSlideToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
      })
    )
  }

  appendNotesToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: `/ppt/notesSlides/notesSlide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml`
      })
    )
  }

  async copyRelatedContent(): Promise<void> {
    let charts = await Chart.getAllOnSlide(this.sourceArchive, this.relsPath)
    for(let i in charts) {
      await new Chart({
        target: charts[i], 
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).append(this.targetTemplate, this.targetNumber)
    }

    let images = await Image.getAllOnSlide(this.sourceArchive, this.relsPath)
    for(let i in images) {
      await new Image({
        target: images[i], 
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceNumber,
      }).append(this.targetTemplate, this.targetNumber)
    }
  }

  hasNotes(): boolean {
    let file = this.sourceArchive.file(`ppt/notesSlides/notesSlide${this.sourceNumber}.xml`)
    return file && file.hasOwnProperty('name')
  }

}
