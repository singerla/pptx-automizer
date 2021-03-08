import { DOMParser, XMLSerializer } from 'xmldom'
import FileHelper from './file'
import { XMLElement } from '../types/xml'
import JSZip from 'jszip'
import {
	RelationshipAttribute, SlideListAttribute, OverrideAttribute
} from '../types/xml'

export default class XmlHelper {

  static async getXmlFromArchive(archive: JSZip, file: string): Promise<Document> {
    let xmlDocument = await FileHelper.extractFromArchive(archive, file)
    const dom = new DOMParser()
    return dom.parseFromString(xmlDocument)
  }
  
  static async writeXmlToArchive(archive: JSZip, file: string, xml: any): Promise<JSZip> {
    let s = new XMLSerializer()
    let xmlBuffer = s.serializeToString(xml)
    
    return archive.file(file, xmlBuffer)
  }

  static async append(element: XMLElement): Promise<HTMLElement> {
    let xml = await XmlHelper.getXmlFromArchive(element.archive, element.file)

    let newElement = xml.createElement(element.tag)
    for(let attribute in element.attributes) {
      let value = element.attributes[attribute]
      newElement.setAttribute(attribute, (typeof value === 'function') ? value(xml) : value)
    }

    let parent = element.parent(xml)
    parent.appendChild(newElement)

    XmlHelper.writeXmlToArchive(element.archive, element.file, xml)

    return newElement
  }

  static async getNextRelId(rootArchive, file): Promise<string> {
    let presentationRelsXml = await XmlHelper.getXmlFromArchive(rootArchive, file)
    let increment: Function = (max: number) => 'rId' + max
    let rid = XmlHelper.getMaxId(presentationRelsXml.documentElement.childNodes, 'Id', true)

    return increment(rid)
  }

  static getMaxId(rels: { [x: string]: any }, attribute: string, increment?: boolean): number {
    let max = 0
    for(let i in rels) {
      let rel = rels[i]
      if(rel.getAttribute !== undefined) {
        let id = Number(rel.getAttribute(attribute).replace('rId', ''))
        max = (id > max) ? id : max
      }
    }

    switch(typeof increment) {
      case 'boolean' : return ++max
      default: return max
    }
  }


  /**
   * Slide related xml helpers
   */
  static async countSlides(presentation: JSZip): Promise<number> {
    let presentationXml = await XmlHelper.getXmlFromArchive(presentation, 'ppt/presentation.xml')
    let slideCount = presentationXml.getElementsByTagName('p:sldId').length

    return slideCount
  }

  static appendToSlideRel(rootArchive: JSZip, relId: string, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/_rels/presentation.xml.rels`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: <RelationshipAttribute> {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`,
        Target: `slides/slide${slideCount}.xml`
      }
    })
  }

  static appendToSlideList(rootArchive: JSZip, relId: string): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `ppt/presentation.xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('p:sldIdLst')[0],
      tag: 'p:sldId',
      attributes: <SlideListAttribute> {
        id: (xml: HTMLElement) => XmlHelper.getMaxId(xml.getElementsByTagName('p:sldId'), 'id', true),
        'r:id': relId
      }
    })
  }

  static appendToContentType(rootArchive: JSZip, slideCount: number): Promise<HTMLElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: `[Content_Types].xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes: <OverrideAttribute> {
        PartName: `/ppt/slides/slide${slideCount}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`
      }
    })
  }

}