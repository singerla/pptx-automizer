
import {
  AutomizerParams,
  AutomizerSummary,
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

  /**
   * Parameters for Automizer constructor.
   * @param {string} templateDir - optional: prefix for all templates to load.
   * @param {string} outputDir - optional: prefix for all output files.
   */
  constructor(params?: AutomizerParams) {
    this.templates = []
    this.params = params
    
    this.templateDir = (params?.templateDir) ? params.templateDir + '/' : ''
    this.outputDir = (params?.outputDir) ? params.outputDir + '/' : ''
  }

	/**
	 * Load a pptx file. 
	 * @param {string} location - Filename or path to the template. Will be prefixed with 'templateDir'
	 * @param {string} name - Optional: A short name for the template. If skipped, the template will be set as RootTemplate.
	 * @return {Slide} Instance of Automizer
	 */
  public load(location: string, name?: string): this {
    location = this.getLocation(location, 'template')
    
    let newTemplate = Template.import(location, name)

    if(!this.isPresTemplate(newTemplate)) {
      this.rootTemplate = newTemplate
    } else {
      this.templates.push(newTemplate)
    }

    return this
  }

  isPresTemplate(template: PresTemplate | RootPresTemplate): template is PresTemplate { 
    return 'name' in template; 
  }

	/**
	 * Find imported template by given name and return a certain slide available
	 * @param {string} name - Name of template; must be imported by Automizer.importTemplate()
	 * @param {number} slideNumber - Number of slide in template presentation
	 * @return {Slide} Imported slide as an instance of Slide
	 */
  public addSlide(name: string, slideNumber: number, callback?: Function): this {
    let template = this.template(name)
    
    let newSlide = new Slide({
      presentation: this,
      template: template,
      number: slideNumber
    })
    
    if(callback !== undefined) {
      newSlide.root = this,
      callback(newSlide)
    }

    this.rootTemplate.slides.push(newSlide)
    
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

  async write(location: string): Promise<AutomizerSummary> {
    await this.rootTemplate.countSlides()
    await this.rootTemplate.countCharts()
    await this.rootTemplate.countImages()

    let rootArchive = await this.rootTemplate.archive

    for(let i in this.rootTemplate.slides) {
      let slide = this.rootTemplate.slides[i]
      await this.rootTemplate.appendSlide(slide)
    }

    let content = await rootArchive.generateAsync({type: "nodebuffer"})

    location = this.getLocation(location, 'output')
    return FileHelper.writeOutputFile(location, content, this)
  }
}