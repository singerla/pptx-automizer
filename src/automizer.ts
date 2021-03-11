
import {
	IPresentationProps, PresTemplate, RootPresTemplate,
} from './types'


import Template from './template'
import Slide from './slide'
import FileHelper from './helper/file'

export default class Automizer implements IPresentationProps {

	rootTemplate: RootPresTemplate
	templates: PresTemplate[]

  constructor() {
    this.templates = []
  }

  public importRootTemplate(location: string): this {
    let newTemplate = Template.importRoot(location)
    this.rootTemplate = newTemplate
    return this
  }

  public importTemplate(location: string, name: string): this {
    let newTemplate = Template.import(location, name)
    this.templates.push(newTemplate)
    return this
  }

	public template(name: string): PresTemplate {
		return this.templates.find(template => template.name === name)
	}

	/**
	 * Find imported template by given name and return a certain slide available
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
    
    this.rootTemplate.slides.push(newSlide)
    
    return newSlide
  }

  async write(location: string): Promise<void> {
    await this.rootTemplate.countSlides()
    await this.rootTemplate.countCharts()

    let rootArchive = await this.rootTemplate.archive

    for(let i in this.rootTemplate.slides) {
      let slide = this.rootTemplate.slides[i]
      await this.rootTemplate.appendSlide(slide)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    FileHelper.writeOutputFile(location, content)
  }

} 