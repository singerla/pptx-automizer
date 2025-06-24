import HyperlinkElement from './modify-hyperlink-element';
import { ShapeModificationCallback } from '../types/types';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { Logger } from './general-helper';

interface RelationshipData {
  Id: string;
  Target: string;
  Type: string;
  TargetMode?: string;
}

/**
 * Helper class for modifying hyperlinks in PowerPoint elements
 */
export default class ModifyHyperlinkHelper {
  private static createRelationshipData(
    target: string | number,
    isInternal: boolean,
  ): RelationshipData {
    if (isInternal) {
      return {
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
        Target: `../slides/${target}`,
        Id: '', // Will be set later
      };
    }

    return {
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      Target: target.toString(),
      TargetMode: 'External',
      Id: '', // Will be set later
    };
  }

  private static addRelationship(
    relation: XmlDocument | XmlElement,
    relData: RelationshipData,
  ): string {
    const relNodes = relation.getElementsByTagName('Relationship');
    const maxId = XmlHelper.getMaxId(relNodes, 'Id', true);

    const newRelId = `rId${maxId}`;

    const newRel = relation.ownerDocument.createElement('Relationship');
    newRel.setAttribute('Id', newRelId);
    Object.entries(relData).forEach(([key, value]) => {
      if (value) newRel.setAttribute(key, value);
    });

    relNodes.item(0).parentNode.appendChild(newRel);

    return newRelId;
  }

  private static addHyperlinkToTextRuns(
    element: XmlElement,
    hyperlinkElement: HyperlinkElement,
  ): void {
    const textRuns = element.getElementsByTagName('a:r');
    Array.from(textRuns).forEach((run) => {
      let rPr = run.getElementsByTagName('a:rPr')[0];
      if (!rPr) {
        rPr = element.ownerDocument.createElement('a:rPr');
        const textElement = run.getElementsByTagName('a:t')[0];
        if (textElement) {
          run.insertBefore(rPr, textElement);
        } else {
          run.appendChild(rPr);
        }
      }
      rPr.appendChild(hyperlinkElement.createHlinkClick());
    });
  }

  private static addHyperlinkToParagraph(
    paragraph: Element,
    hyperlinkElement: HyperlinkElement,
  ): void {
    const existingText = paragraph.getElementsByTagName('a:t')[0];
    const text = existingText?.textContent || 'Hyperlink';

    if (existingText?.parentNode) {
      paragraph.removeChild(existingText.parentNode);
    }

    const run = hyperlinkElement.createTextRun(text);
    paragraph.appendChild(run);
  }

  private static createNewTextStructure(
    txBody: Element,
    hyperlinkElement: HyperlinkElement,
  ): void {
    const p = txBody.ownerDocument.createElement('a:p');
    const run = hyperlinkElement.createTextRun('Hyperlink');
    p.appendChild(run);
    txBody.appendChild(p);
  }

  /**
   * Set the target URL of a hyperlink
   *
   * @param target The new target URL for the hyperlink
   * @param isExternal Whether the hyperlink is external (true) or internal (false)
   * @returns A callback function that modifies the hyperlink
   */
  static setHyperlinkTarget =
    (target: string | number, isExternal = true): ShapeModificationCallback =>
    async (element: XmlElement, relation?: XmlElement): Promise<void> => {
      if (!element || !relation) {
        Logger.log('SetHyperlinkTarget: Missing element or relation', 2);
        return;
      }

      // Find existing hyperlinks
      const hlinkClicks = element.getElementsByTagName('a:hlinkClick');
      if (hlinkClicks.length === 0) {
        Logger.log('No hyperlinks found to modify', 1);
        return;
      }

      // Get all existing rIds from hyperlinks
      const existingRIds = Array.from(hlinkClicks)
        .map((hlink) => hlink.getAttribute('r:id'))
        .filter(Boolean) as string[];

      if (existingRIds.length === 0) {
        Logger.log('No valid relationship IDs found in hyperlinks', 1);
        return;
      }

      // Create new relationship data
      const relData = this.createRelationshipData(target, !isExternal);
      const newRelId = this.addRelationship(relation, relData);

      // Update all hyperlink elements with new relationship ID
      Array.from(hlinkClicks).forEach((hlink) => {
        // Update relationship ID
        hlink.setAttribute('r:id', newRelId);

        // Update internal/external specific attributes
        if (!isExternal) {
          hlink.setAttribute('action', 'ppaction://hlinksldjump');
          hlink.setAttribute(
            'xmlns:a',
            'http://schemas.openxmlformats.org/drawingml/2006/main',
          );
          hlink.setAttribute(
            'xmlns:p14',
            'http://schemas.microsoft.com/office/powerpoint/2010/main',
          );
        } else {
          hlink.removeAttribute('action');
          // Keep xmlns attributes as they're still needed for the relationship
        }
      });

      // Remove old relationships
      const relationships = relation.getElementsByTagName('Relationship');
      Array.from(relationships).forEach((rel) => {
        const relId = rel.getAttribute('Id');
        if (relId && existingRIds.includes(relId)) {
          relation.removeChild(rel);
        }
      });

      Logger.log(
        'SetHyperlinkTarget: Successfully updated hyperlink target',
        2,
      );
    };

  /**
   * Add a hyperlink to an element
   *
   * @param target The target URL for external links, or slide number for internal links
   * @param isInternalLink
   * @returns A callback function that adds a hyperlink
   */
  static addHyperlink =
    (
      target: string | number,
      isInternalLink?: boolean,
    ): ShapeModificationCallback =>
    (element: XmlElement, relation: XmlElement): void => {
      if (!element || !relation) return;

      if (typeof target === 'number') {
        target = `slide${target}.xml`;
        isInternalLink = true;
      }

      const relData = this.createRelationshipData(target, isInternalLink);
      const newRelId = this.addRelationship(relation, relData);

      const hasHlink = element.getElementsByTagName('a:hlinkClick');
      if (hasHlink.item(0)) {
        // Link has already been set by e.g. pptxGenJs, don't add another link to element
        return;
      }

      const hyperlinkElement = new HyperlinkElement(
        element.ownerDocument,
        newRelId,
        isInternalLink,
      );

      const textRuns = element.getElementsByTagName('a:r');
      if (textRuns.length > 0) {
        this.addHyperlinkToTextRuns(element, hyperlinkElement);
      } else {
        const paragraphs = element.getElementsByTagName('a:p');
        if (paragraphs.length > 0) {
          this.addHyperlinkToParagraph(paragraphs[0], hyperlinkElement);
        } else {
          const txBody =
            element.getElementsByTagName('p:txBody')[0] ||
            element.getElementsByTagName('a:txBody')[0];
          if (txBody) {
            this.createNewTextStructure(txBody, hyperlinkElement);
          } else {
            console.error('No suitable text element found to add hyperlink to');
          }
        }
      }

      Logger.log('AddHyperlink: Successfully completed', 2);
    };

  /**
   * Remove hyperlinks from an element
   *
   * @returns A callback function that removes hyperlinks
   */
  static removeHyperlink =
    (): ShapeModificationCallback =>
    async (element: XmlElement, relation?: XmlElement): Promise<void> => {
      if (!element) return;

      try {
        const hlinkClicks = element.getElementsByTagName('a:hlinkClick');

        Array.from(hlinkClicks).forEach((hlink) =>
          hlink.parentNode?.removeChild(hlink),
        );
        Logger.log('RemoveHyperlink: Successfully completed', 2);
      } catch (error) {
        console.error('Error in RemoveHyperlink:', error);
      }
    };
}
