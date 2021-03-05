import {
	IPresentationProps, PresSlide, PresTemplate
} from './types/interfaces'

import fs from 'fs'

import Template from './template'
import Slide from './slide'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

export default class Automizer implements IPresentationProps {

	private _rootTemplate: PresTemplate
	public get rootTemplate(): PresTemplate {
		return this._rootTemplate
	}

	private _templates: PresTemplate[]
	public get templates(): PresTemplate[] {
		return this._templates
	}

  constructor() {
    this._templates = []
    this._rootTemplate = <PresTemplate> {}
  }

  public importRootTemplate(location: string): this {
    let newTemplate = Template.import(location)
    this._rootTemplate = newTemplate
    return this
  }

  public importTemplate(location: string, name: string): void {
    let newTemplate = Template.import(location, name)
    this._templates.push(newTemplate)
  }

	public template(name): PresTemplate {
		return this._templates.find(template => template.name === name)
	}

  public addSlide(name: string, slideNumber: number): Slide {
    let template = this.template(name)
    
    let newSlide = new Slide({
      presentation: this,
      template: template,
      number: slideNumber
    })
    
    this._rootTemplate.slides.push(newSlide)
    
    return newSlide
  }

  async write(location: string) {
    let rootArchive = await this._rootTemplate.archive
    let presentationXml = await XmlHelper.getXmlFromArchive(rootArchive, 'ppt/presentation.xml')
    let slideCount = presentationXml.getElementsByTagName('p:sldId').length

    for(let i in this._rootTemplate.slides) {
      ++slideCount
      
      let slide = this._rootTemplate.slides[i]
      let slidePath = `ppt/slides/slide${slide.number}.xml`
      let slideTemplate = slide.template.archive


      for(let m in slide.modifications) {
        let callback = slide.modifications[m]
        let archive = await slideTemplate
        let slideXml = await XmlHelper.getXmlFromArchive(archive, slidePath)
        callback(slideXml)
        await XmlHelper.writeXmlToArchive(archive, slidePath, slideXml)
      }

      await this.copyFiles(slideTemplate, slide.number, rootArchive, slideCount)

      let relId = await XmlHelper.getNextRelId(rootArchive, 'ppt/_rels/presentation.xml.rels')
      await this.appendRel(rootArchive, relId, slideCount)
      await this.appendToSlideList(rootArchive, relId)
      await this.appendToContentType(rootArchive, slideCount)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    FileHelper.writeOutputFile(location, content)
  }
  
  async copyFiles(sourceArchive, sourceSlide, targetArchive, targetSlide) {
     FileHelper.zipCopy(
      sourceArchive, `ppt/slides/slide${sourceSlide}.xml`, 
      targetArchive, `ppt/slides/slide${targetSlide}.xml`
    )

    FileHelper.zipCopy(
      sourceArchive, `ppt/slides/_rels/slide${sourceSlide}.xml.rels`, 
      targetArchive, `ppt/slides/_rels/slide${targetSlide}.xml.rels`
    )
  }

  async appendRel(rootArchive, relId: string, slideCount: number) {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/_rels/presentation.xml.rels`,
      parent: (xml) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`,
        Target: `slides/slide${slideCount}.xml`
      }
    })
  }

  async appendToSlideList(rootArchive, relId: string) {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      parent: (xml) => xml.getElementsByTagName('p:sldIdLst')[0],
      tag: 'p:sldId',
      attributes: {
        id: (xml) => XmlHelper.getMaxId(xml.getElementsByTagName('p:sldId'), 'id', true),
        'r:id': relId
      }
    })
  }

  async appendToContentType(rootArchive, slideCount: number) {
    return XmlHelper.append({
      archive: rootArchive,
      file: `[Content_Types].xml`,
      parent: (xml) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes: {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
      }
    })
  }

} 