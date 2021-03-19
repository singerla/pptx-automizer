import JSZip from 'jszip'
import FileHelper from '../helper/file'
import XmlHelper from '../helper/xml'
import Shape from '../shape'

import { IImage, ImportedElement, RootPresTemplate, Target } from '../definitions/app'
import { RelationshipAttribute } from '../definitions/xml'
import { ImageTypeMap, ElementType } from '../definitions/enums'

export default class Image extends Shape implements IImage {  
  extension: string
  contentTypeMap: object

  constructor(shape: ImportedElement) {
    super(shape)
    
    this.sourceFile = shape.target.file.replace('../media/', '')
    this.extension = FileHelper.getFileExtension(this.sourceFile)
    this.relAttribute = 'r:embed'
    
    switch(this.extension) {
      case 'svg':
        this.relRootTag = 'asvg:svgBlip'
        this.relParent = element => <HTMLElement> element.parentNode
      break
      default:
        this.relRootTag = 'a:blip'
        this.relParent = element => <HTMLElement> element.parentNode.parentNode
      break
    }

    this.contentTypeMap = ImageTypeMap
  }

  async modify(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber)
    await this.updateElementRelId()

    return this
  }

  async modifyOnAddedSlide(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber)
    await this.updateElementRelId()

    return this
  }

  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number): Promise<Image> {
    await this.prepare(targetTemplate, targetSlideNumber)
    await this.setTargetElement()

    this.applyCallbacks(this.callbacks, this.targetElement)

    await this.updateTargetElementRelId()
    await this.appendToSlideTree()

    if(this.hasSvgRelation()) {
      let target = await XmlHelper.getTargetByRelId(this.sourceArchive, this.sourceSlideNumber, this.targetElement, 'image:svg')
      await new Image({
        mode: 'append',
        target: target, 
        sourceArchive: this.sourceArchive,
        sourceSlideNumber: this.sourceSlideNumber,
        type: ElementType.Image,
      }).modify(targetTemplate, targetSlideNumber)
    }

    return this
  }

  async prepare(targetTemplate, targetSlideNumber) {
    await this.setTarget(targetTemplate, targetSlideNumber)
    
    this.targetNumber = this.targetTemplate.incrementCounter('images')
    this.targetFile = 'image' + this.targetNumber + '.' + this.extension

    await this.copyFiles()
    await this.appendTypes()
    await this.appendToSlideRels()
  }

  async copyFiles() {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/media/${this.sourceFile}`, 
      this.targetArchive, `ppt/media/${this.targetFile}`
    )
  }
  
  async appendTypes() {
    await this.appendImageExtensionToContentType()
  }

  async appendToSlideRels(): Promise<HTMLElement> {
    let targetRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`
    this.createdRid = await XmlHelper.getNextRelId(this.targetArchive, targetRelFile)

    let attributes = <RelationshipAttribute> {
      Id: this.createdRid,
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      Target: `../media/image${this.targetNumber}.${this.extension}`
    }

    return XmlHelper.append(
      XmlHelper.createRelationshipChild(this.targetArchive, targetRelFile, attributes)
    )
  }

  appendImageExtensionToContentType(): Promise<HTMLElement | boolean> {
    let extension = this.extension
    let contentType = (this.contentTypeMap[extension]) ? this.contentTypeMap[extension] : 'image/' + extension
    
    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(this.targetArchive, {
        Extension: extension,
        ContentType: contentType
      }),
      tag: 'Default',
      clause: (xml: HTMLElement) => !XmlHelper.findByAttribute(xml, 'Default', 'Extension', extension)
    })
  }

  hasSvgRelation() {
    return (this.targetElement.getElementsByTagName('asvg:svgBlip').length > 0)
  }

  static async getAllOnSlide(archive: JSZip, relsPath: string): Promise<Target[]> {
    return await XmlHelper.getTargetsByRelationshipType(archive, relsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
  }
}
