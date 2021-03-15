import { DOMParser, XMLSerializer } from 'xmldom'
import FileHelper from './file'
import { RelationshipAttribute, XMLElement } from '../types'
import JSZip from 'jszip'
import {
  DefaultAttribute, OverrideAttribute, Target
} from '../types'

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

  static async appendIf(element: XMLElement): Promise<HTMLElement | boolean> {
    let xml = await XmlHelper.getXmlFromArchive(element.archive, element.file)
    
    if(element.clause !== undefined) {
      if(!element.clause(xml)) return false
    }

    return XmlHelper.append(element)
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

  static async getTargetsFromRelationships(archive: JSZip, path: string, prefix: string, suffix?: string | RegExp): Promise<Target[]>{
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      let target = element.getAttribute('Target')
      if(target.indexOf(prefix) === 0) {
        rels.push(<Target>{
          file: target,
          rId: element.getAttribute('Id'),
          number: Number(target.replace(prefix, '').replace(suffix || '.xml', ''))
        })
      }
    })
  }

  static async getTargetsByRelationshipType(archive: JSZip, path: string, type: string): Promise<Target[]>{
    return XmlHelper.getRelationships(archive, path, (element: HTMLElement, rels: Target[]) => {
      let target = element.getAttribute('Type')
      if(target === type) {
        rels.push(<Target>{
          file: element.getAttribute('Target'),
          rId: element.getAttribute('Id'),
        })
      }
    })
  }

  static async getRelationships(archive: JSZip, path: string, cb: Function) {
    let xml = await XmlHelper.getXmlFromArchive(archive, path)
    let relationships = xml.getElementsByTagName('Relationship')
    let rels = []
    for(let i in relationships) {
      let element = relationships[i]
      if(element.getAttribute !== undefined) {
        cb(element, rels)
      }
    }
    return rels
  }

  static findByAttribute(xml: HTMLElement, tagName: string, attributeName: string, attributeValue: string): Boolean {
    let elements = xml.getElementsByTagName(tagName)
    for(let i in elements) {
      let element = elements[i]
      if(element.getAttribute !== undefined) {
        if(element.getAttribute(attributeName) === attributeValue) {
          return true
        }
      }
    }
    return false
  }

  static async findByElementName(archive: JSZip, path: string, name: string): Promise<any> {
    let slideXml = await XmlHelper.getXmlFromArchive(archive, path)
    let names = slideXml.getElementsByTagName('p:cNvPr')
    
    for(let i in names) {
      if(names[i].getAttribute && names[i].getAttribute('name') === name) {
        return names[i].parentNode.parentNode
      }
    }

    return null
  }

  static createContentTypeChild(archive: JSZip, attributes: OverrideAttribute | DefaultAttribute): XMLElement {
    return {
      archive: archive,
      file: `[Content_Types].xml`,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Types')[0],
      tag: 'Override',
      attributes: attributes
    }
  }

  static createRelationshipChild(archive: JSZip, targetRelFile:string, attributes: RelationshipAttribute): XMLElement {
    return {
      archive: archive,
      file: targetRelFile,
      parent: (xml: HTMLElement) => xml.getElementsByTagName('Relationships')[0],
      tag: 'Relationship',
      attributes: attributes
    }
  }
}
