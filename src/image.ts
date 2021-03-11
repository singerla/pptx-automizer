import JSZip from 'jszip'
import FileHelper from './helper/file'
import XmlHelper from './helper/xml'

import {
	IImage, Target
} from './types'


export default class Image implements IImage {  
  sourceArchive: JSZip
  targetArchive: JSZip
  sourceFile: string
  targetFile: string
  targetSlide: this
  targetSlideNumber: number
  sourceRid: any
  contentTypeMap: any
  targetNumber: number
  extension: string

  constructor(relsXmlInfo: Target, sourceArchive: JSZip, targetSlideNumber: number) {
    this.sourceFile = relsXmlInfo.file.replace('../media/', '')
    this.sourceRid = relsXmlInfo.rId
    this.sourceArchive = sourceArchive
    this.targetSlideNumber = targetSlideNumber
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


  setTarget(archive: JSZip, number: number) {
    this.targetArchive = archive
    this.targetNumber = number
    this.extension = FileHelper.getFileExtension(this.sourceFile)
    this.targetFile = 'image' + number + '.' + this.extension
  }

  async append() {
    await this.copyFiles()
    await this.appendTypes()
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
