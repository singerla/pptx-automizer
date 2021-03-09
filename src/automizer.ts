import JSZip from 'jszip'

import {
	IPresentationProps, PresTemplate, RootPresTemplate,
} from './types/interfaces'


import Template from './template'
import Slide from './slide'
import Chart from './chart'

import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

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
      let slide = this._rootTemplate.slides[i]
      await this._rootTemplate.appendSlide(slide)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    FileHelper.writeOutputFile(location, content)
  }

} 