
import {
  AutomizerParams,
	IPresentationProps, PresTemplate, RootPresTemplate,
} from './types'


import Template from './template'
import Slide from './slide'
import FileHelper from './helper/file'

export default class Automizer implements IPresentationProps {

	rootTemplate: RootPresTemplate
	templates: PresTemplate[]
	templateDir: string
	outputDir: string
  params: AutomizerParams

  constructor(params?: AutomizerParams) {
    this.templates = []
    this.params = params

    this.templateDir = (params?.templateDir) ? params.templateDir + '/' : ''
    this.outputDir = (params?.outputDir) ? params.outputDir + '/' : ''
  }

  public importRootTemplate(location: string): this {
    location = this.getLocation(location, 'template')
    let newTemplate = Template.importRoot(location)
    this.rootTemplate = newTemplate
    return this
  }

  public importTemplate(location: string, name: string): this {
    location = this.getLocation(location, 'template')
    let newTemplate = Template.import(location, name)
    this.templates.push(newTemplate)
    return this
  }

	public template(name: string): PresTemplate {
		return this.templates.find(template => template.name === name)
	}

  public getLocation(location: string, type?: string): string {
    switch(type) {
      case 'template':
        return this.templateDir + location
      case 'output':
        return this.outputDir + location
      default:
        return location
    }
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

  async write(location: string): Promise<string> {
    await this.rootTemplate.countSlides()
    await this.rootTemplate.countCharts()

    let rootArchive = await this.rootTemplate.archive

    for(let i in this.rootTemplate.slides) {
      let slide = this.rootTemplate.slides[i]
      await this.rootTemplate.appendSlide(slide)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    location = this.getLocation(location, 'output')
    return FileHelper.writeOutputFile(location, content)
  }
}