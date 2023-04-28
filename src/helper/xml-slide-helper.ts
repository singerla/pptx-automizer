import { XmlDocument, XmlElement } from '../types/xml-types';

/**
 * Class that represents an XML slide helper
 */
export class XmlSlideHelper {
  private slideXml: XmlDocument;

  /**
   * Constructor for the XmlSlideHelper class.
   * @param {XmlDocument} slideXml - The slide XML document to be used by the helper.
   */
  constructor(slideXml: XmlDocument) {
    if (!slideXml) {
      throw Error('Slide XML is not defined');
    }
    this.slideXml = slideXml;
  }

  /**
   * Get all text element IDs from the slide.
   * @param {boolean} [useCreationIds=false] - If true, use creation IDs when available; otherwise, use element names.
   * @return {string[]} An array of text element IDs.
   */
  getAllTextElementIds(useCreationIds = false): string[] {
    const elementIds: string[] = [];

    try {
      const shapeNodes = this.slideXml.getElementsByTagNameNS(
        'http://schemas.openxmlformats.org/presentationml/2006/main',
        'sp',
      );

      for (let i = 0; i < shapeNodes.length; i++) {
        const shapeNode = shapeNodes.item(i);
        const txBody = shapeNode
          .getElementsByTagNameNS(
            'http://schemas.openxmlformats.org/presentationml/2006/main',
            'txBody',
          )
          .item(0);

        // .. if the shape node contains a text body
        if (txBody) {
          const cNvPr = shapeNode
            .getElementsByTagNameNS(
              'http://schemas.openxmlformats.org/presentationml/2006/main',
              'cNvPr',
            )
            .item(0);

          // Check if the shape node contains a non-visual drawing properties element
          if (cNvPr) {
            let id: string;
            let creationIds: HTMLCollectionOf<Element>;

            if (useCreationIds) {
              creationIds =
                this.slideXml.getElementsByTagName('a16:creationId');
            }

            // Use the creation ID if useCreationIds is true and creationIds.length > 1; otherwise, use the element name
            if (useCreationIds && creationIds.length > 1) {
              id = cNvPr.getAttribute('id');
            } else {
              id = cNvPr.getAttribute('name');
            }

            // Add the ID to the elementIds array if it exists, else warn but dont break,
            if (id) {
              elementIds.push(id);
            } else {
              console.warn('Element ID is missing for a text element');
            }
          }
        }
      }
    } catch (error) {
      throw new Error(`Failed to retrieve text element IDs: ${error.message}`);
    }

    return elementIds;
  }

  // Other slide-related helper functions will go here
}
