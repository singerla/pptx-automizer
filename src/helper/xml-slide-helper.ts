import { XmlDocument, XmlElement } from '../types/xml-types';

export class XmlSlideHelper {
  private slideXml: XmlDocument;

  constructor(slideXml: XmlDocument) {
    if (!slideXml) { throw Error('Slide XML is not defined'); }
    this.slideXml = slideXml;
  }

  getAllTextElementIds(useCreationIds = false): string[] {
    const shapeNodes = this.slideXml.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'sp');
    const elementIds: string[] = [];

    for (let i = 0; i < shapeNodes.length; i++) {
      const shapeNode = shapeNodes.item(i);
      const txBody = shapeNode.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'txBody').item(0);

      if (txBody) {
        const cNvPr = shapeNode.getElementsByTagNameNS('http://schemas.openxmlformats.org/presentationml/2006/main', 'cNvPr').item(0);

        if (cNvPr) {
          let id: string;
          let creationIds: string | string[] | HTMLCollectionOf<Element>;

          if (useCreationIds) {
            creationIds = this.slideXml.getElementsByTagName('a16:creationId');
          }

          if (useCreationIds && creationIds.length > 1) {
            id = cNvPr.getAttribute('id');
          } else {
            id = cNvPr.getAttribute('name');
          }

          if (id) {
            elementIds.push(id);
          } else {
            console.warn('Element ID is missing for a text element');
          }
        }
      }
    }

    return elementIds;
  }

  // Other slide-related helper functions will go here
}
