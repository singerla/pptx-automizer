import JSZip from "jszip";
import {ElementInfo, SlideInfo} from "../types/xml-types";
import {XmlHelper} from "./xml-helper";

export class XmlTemplateHelper {
  archive: JSZip;
  relType: string;
  path: string;

  constructor(archive: JSZip) {
    this.relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
    this.archive = archive
    this.path = 'ppt/_rels/presentation.xml.rels'
  }

  async getCreationIds(): Promise<SlideInfo[]> {
    const archive = await this.archive
    const relationships = await XmlHelper.getTargetsByRelationshipType(
      archive,
      this.path,
      this.relType
    )

    const creationIds = []
    for(const slideRel of relationships) {
      const slideXml = await XmlHelper.getXmlFromArchive(
        archive,
        'ppt/' + slideRel.file
      )

      const creationIdSlide = slideXml
        .getElementsByTagName('p14:creationId')
        .item(0)
        .getAttribute('val')

      const elementIds = this.elementCreationIds(slideXml)

      creationIds.push({
        id: Number(creationIdSlide),
        number: this.parseSlideRelFile(slideRel.file),
        elements: elementIds
      })
    }

    return creationIds
  }

  parseSlideRelFile(slideRelFile: string): number {
    return Number(slideRelFile
      .replace('slides/slide', '')
      .replace('.xml', '')
    )
  }

  elementCreationIds(slideXml): ElementInfo[] {
    const slideElements = slideXml
      .getElementsByTagName('p:cNvPr')

    const elementIds = <ElementInfo[]> []
    for(const item in slideElements) {
      const slideElement = slideElements[item]
      if(slideElement.getAttribute !== undefined) {
        const creationIdElement = slideElement
          .getElementsByTagName('a16:creationId')
        if(creationIdElement.item(0)) {
          elementIds.push(
            this.getElementInfo(slideElement)
          )
        }
      }
    }
    return elementIds
  }

  getElementInfo(slideElement): ElementInfo {
    const elementName = slideElement.getAttribute('name')
    const type = slideElement.parentNode.parentNode.tagName
    const creationId = slideElement
      .getElementsByTagName('a16:creationId')
      .item(0)
      .getAttribute('id')

    return {
      name: elementName,
      type: type.replace('p:', ''),
      id: creationId
    }
  }
}
