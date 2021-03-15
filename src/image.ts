import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import {
	IImage, RelationshipAttribute, RootPresTemplate, Target
} from './types'


export default class Image implements IImage {  
  sourceArchive: JSZip
  sourceFile: string
  sourceRid: any
  targetFile: string
  targetArchive: JSZip
  targetTemplate: RootPresTemplate
  targetSlideNumber: number
  contentTypeMap: any
  targetNumber: number
  extension: string
  createdRid: string

  constructor(relsXmlInfo: Target, sourceArchive: JSZip) {
    this.sourceFile = relsXmlInfo.file.replace('../media/', '')
    this.sourceRid = relsXmlInfo.rId
    this.sourceArchive = sourceArchive
    
    this.contentTypeMap = {
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      png: "image/png",
      gif: "image/gif",
      svg: "image/svg+xml",
      m4v: "video/mp4",
      mp4: "video/mp4"
    }
  }

  
  async append(targetTemplate: RootPresTemplate, targetSlideNumber: number, appendMode?: boolean): Promise<void> {
    this.targetTemplate = targetTemplate
    this.targetArchive = await this.targetTemplate.archive
    this.targetSlideNumber = targetSlideNumber
    this.targetNumber = this.targetTemplate.incrementCounter('charts')
    this.extension = FileHelper.getFileExtension(this.sourceFile)
    this.targetFile = 'image' + this.targetNumber + '.' + this.extension

    await this.copyFiles()
    await this.appendTypes()
    
    if(appendMode === true) {
      await this.appendToSlideRels()
    } else {
      await this.editTargetChartRel()
    }
  }

  async copyFiles() {
    FileHelper.zipCopy(
      this.sourceArchive, `ppt/media/${this.sourceFile}`, 
      this.targetArchive, `ppt/media/${this.targetFile}`
    )
  }
  
  async appendTypes() {
    await this.editTargetChartRel()
    await this.appendImageExtensionToContentType()
  }

  async editTargetChartRel() {
    let targetRelFile = `ppt/slides/_rels/slide${this.targetSlideNumber}.xml.rels`
    let relXml = await XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile)
    let relations = relXml.getElementsByTagName('Relationship')

    for(let i in relations) {
      let element = relations[i]
      if(element.getAttribute) {
        let rId = element.getAttribute('Id')
        if(rId === this.sourceRid) {
          element.setAttribute('Target', `../media/${this.targetFile}`)
        }
      }
    }

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml)
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


}
