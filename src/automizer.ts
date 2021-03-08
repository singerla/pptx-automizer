import {
	IPresentationProps, PresTemplate, RootPresTemplate,
} from './types/interfaces'


import Template from './template'
import Slide from './slide'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import JSZip from 'jszip'


export default class Automizer implements IPresentationProps {

	private _rootTemplate: RootPresTemplate
	public get rootTemplate(): RootPresTemplate {
		return this._rootTemplate
	}

	private _templates: PresTemplate[]
	public get templates(): PresTemplate[] {
		return this._templates
	}

  constructor() {
    this._templates = []
  }

  public importRootTemplate(location: string): this {
    let newTemplate = Template.importRoot(location)
    this._rootTemplate = newTemplate
    return this
  }

  public importTemplate(location: string, name: string): void {
    let newTemplate = Template.import(location, name)
    this._templates.push(newTemplate)
  }

	public template(name: string): PresTemplate {
		return this._templates.find(template => template.name === name)
	}

	/**
	 * Search imported templates for given name and make a certain slide available
	 * @param {string} name - Name of template; must be imported by Automizer.importTemplate()
	 * @param {number} slideNumber - Number of slide in template presentation
	 * @return {Slide} Imported slide as an instance of Slide
	 */
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

  async write(location: string): Promise<void> {
    let rootArchive = await this._rootTemplate.archive
    let slideCount = await this._rootTemplate.countSlides()

    for(let i in this._rootTemplate.slides) {
      slideCount = this._rootTemplate.incrementSlideCounter()

      let slide = this._rootTemplate.slides[i]
      let slidePath = `ppt/slides/slide${slide.number}.xml`
      let slideArchive = await slide.template.archive

      await this.applyModifications(slide.modifications, slideArchive, slidePath)
      await this.copySlideFiles(slideArchive, slide.number, rootArchive, slideCount)
      await this.addContentToPresentation(rootArchive, slideCount)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    FileHelper.writeOutputFile(location, content)
  }
  
  async addContentToPresentation(rootArchive: JSZip, slideCount: number): Promise<HTMLElement[]> {
    let relId = await XmlHelper.getNextRelId(rootArchive, 'ppt/_rels/presentation.xml.rels')
    let promises = [
      XmlHelper.appendToSlideRel(rootArchive, relId, slideCount),
      XmlHelper.appendToSlideList(rootArchive, relId),
      XmlHelper.appendToContentType(rootArchive, slideCount)
    ]
    return Promise.all(promises)
  }

  async applyModifications(modifications: Function[], template: JSZip, path: string) {
    for(let m in modifications) {
      let xml = await XmlHelper.getXmlFromArchive(template, path)
      modifications[m](xml)
      await XmlHelper.writeXmlToArchive(template, path, xml)
    }
  }

  async copySlideFiles(sourceArchive: JSZip, sourceSlide: number, targetArchive: JSZip, targetSlide: string): Promise<void> {
    FileHelper.zipCopy(
      sourceArchive, `ppt/slides/slide${sourceSlide}.xml`, 
      targetArchive, `ppt/slides/slide${targetSlide}.xml`
    )

    FileHelper.zipCopy(
      sourceArchive, `ppt/slides/_rels/slide${sourceSlide}.xml.rels`, 
      targetArchive, `ppt/slides/_rels/slide${targetSlide}.xml.rels`
    )
  }
} 