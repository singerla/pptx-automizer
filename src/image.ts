import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'
import Shape from './shape'

import { IImage, RootPresTemplate, Target } from './types/app'
import { RelationshipAttribute } from './types/xml'

export default class Image extends Shape implements IImage {  

  extension: string

  constructor(relsXmlInfo: Target, sourceArchive: JSZip, sourceSlideNumber?:number) {
    super(relsXmlInfo, sourceArchive, sourceSlideNumber)
    
    this.sourceFile = relsXmlInfo.file.replace('../media/', '')
    this.extension = FileHelper.getFileExtension(this.sourceFile)

    let mapRelRootTags = {
      svg: 'asvg:svgBlip'
    }

    this.relRootTag = (mapRelRootTags[this.extension]) 
      ? mapRelRootTags[this.extension] 
      : 'a:blip'
      
    this.relAttribute = 'r:embed'
    this.relParent = element => <HTMLElement> element.parentNode.parentNode

    this.contentTypeMap = {
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      png: "image/png",
      gif: "image/gif",
      svg: "image/svg+xml",
      m4v: "video/mp4",
      mp4: "video/mp4",
      emf: "image/x-emf"
    }
  }
  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number, appendToTree?: boolean): Promise<Image> {
    await this.setTarget(targetTemplate, targetSlideNumber)
    
    this.targetNumber = this.targetTemplate.incrementCounter('images')
    this.targetFile = 'image' + this.targetNumber + '.' + this.extension

    await this.copyFiles()
    await this.appendTypes()
    await this.appendToSlideRels()

    if(appendToTree) {
      await this.appendToSlideTree()
    }

    await this.updateElementRelId()

    return this
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

  static async getAllOnSlide(archive: JSZip, relsPath: string): Promise<Target[]> {
    return await XmlHelper.getTargetsByRelationshipType(archive, relsPath, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
  }
}
