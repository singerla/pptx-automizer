import HyperlinkElement from './modify-hyperlink-element';
import { HyperlinkProcessor } from './hyperlink-processor';
import { ShapeModificationCallback } from '../types/types';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { log } from './logger';

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
    relId?: string,
  ): string {
    const relNodes = relation.getElementsByTagName('Relationship');
    const newRelId = relId || `rId${XmlHelper.getMaxId(relNodes, 'Id', true)}`;

    const newRel = relation.ownerDocument.createElement('Relationship');
    newRel.setAttribute('Id', newRelId);
    Object.entries(relData).forEach(([key, value]) => {
      if (value) newRel.setAttribute(key, value);
    });

    relNodes.item(0).parentNode.appendChild(newRel);

    return newRelId;
  }

  /**
   * Returns the <Relationship> entry backing a relationship id, if any.
   */
  private static getRelationshipById(
    relation: XmlDocument | XmlElement,
    relId: string,
  ): XmlElement | null {
    const relNodes = relation.getElementsByTagName('Relationship');
    return (
      Array.from(relNodes).find((rel) => rel.getAttribute('Id') === relId) ||
      null
    );
  }

  /**
   * Checks if a relationship id is referenced by a hyperlink outside of the
   * given element, e.g. by another shape on the same slide.
   */
  private static isRelIdUsedElsewhere(
    element: XmlElement,
    relId: string,
  ): boolean {
    const slideXml = element.ownerDocument;
    if (!slideXml) return false;

    return Array.from(slideXml.getElementsByTagName('a:hlinkClick')).some(
      (hlink) =>
        hlink.getAttribute('r:id') === relId && !element.contains(hlink),
    );
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
    paragraph: XmlElement,
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
    txBody: XmlElement,
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
        log.debug('SetHyperlinkTarget: Missing element or relation');
        return;
      }

      // Find existing hyperlinks
      const hlinkClicks = element.getElementsByTagName('a:hlinkClick');
      if (hlinkClicks.length === 0) {
        log.warn('No hyperlinks found to modify');
        return;
      }

      // Get all existing rIds from hyperlinks
      const existingRIds = Array.from(hlinkClicks)
        .map((hlink) => hlink.getAttribute('r:id'))
        .filter(Boolean) as string[];

      if (existingRIds.length === 0) {
        log.warn('No valid relationship IDs found in hyperlinks');
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

      // Remove old relationships, unless another shape still refers to them.
      // Dropping a shared rId would leave r:id attributes with no matching
      // relationship behind, and make PowerPoint ask to repair the file.
      const relationships = relation.getElementsByTagName('Relationship');
      Array.from(relationships).forEach((rel) => {
        const relId = rel.getAttribute('Id');
        if (
          relId &&
          existingRIds.includes(relId) &&
          !this.isRelIdUsedElsewhere(element, relId)
        ) {
          relation.removeChild(rel);
        }
      });

      log.debug('SetHyperlinkTarget: Successfully updated hyperlink target');
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

      const existingHlink = element.getElementsByTagName('a:hlinkClick').item(0);
      if (existingHlink) {
        const existingRid = existingHlink.getAttribute('r:id');
        if (!existingRid) {
          // An hlinkClick without r:id is action-only (e.g. a ppaction jump);
          // nothing to wire up.
          return;
        }

        const existingRel = this.getRelationshipById(relation, existingRid);
        if (
          existingRel &&
          HyperlinkProcessor.isHyperlinkRelType(
            existingRel.getAttribute('Type') || '',
          )
        ) {
          // Link has already been set and its relationship already created
          // by e.g. pptxGenJs, don't add another link to the element.
          return;
        }

        if (existingRel) {
          // The existing r:id collides with an unrelated relationship on this
          // slide (an image, the layout, …). Reusing it would silently point
          // the hyperlink at that part — allocate a fresh id instead and
          // rewrite every hlinkClick in the element carrying the stale id.
          const relData = this.createRelationshipData(target, isInternalLink);
          const freshRelId = this.addRelationship(relation, relData);
          Array.from(element.getElementsByTagName('a:hlinkClick')).forEach(
            (hlink) => {
              if (hlink.getAttribute('r:id') === existingRid) {
                hlink.setAttribute('r:id', freshRelId);
              }
            },
          );
          log.debug(
            'AddHyperlink: existing r:id collided with a non-hyperlink relationship, assigned a fresh id',
          );
          return;
        }

        // The element already carries an <a:hlinkClick>, but its r:id has no
        // backing <Relationship> (e.g. a shape cloned from a template without
        // its relationship). Create a relationship for that existing r:id so
        // it resolves, instead of leaving it unmatched or creating an unused
        // extra one.
        const relData = this.createRelationshipData(target, isInternalLink);
        this.addRelationship(relation, relData, existingRid);
        log.debug('AddHyperlink: Created missing relationship for existing hyperlink');
        return;
      }

      const relData = this.createRelationshipData(target, isInternalLink);
      const newRelId = this.addRelationship(relation, relData);

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
            log.error('No suitable text element found to add hyperlink to');
          }
        }
      }

      log.debug('AddHyperlink: Successfully completed');
    };

  /**
   * Remove hyperlinks from an element
   *
   * @returns A callback function that removes hyperlinks
   */
  static removeHyperlink =
    (): ShapeModificationCallback =>
    async (element: XmlElement, _relation?: XmlElement): Promise<void> => {
      if (!element) return;

      try {
        const hlinkClicks = element.getElementsByTagName('a:hlinkClick');

        Array.from(hlinkClicks).forEach((hlink) =>
          hlink.parentNode?.removeChild(hlink),
        );
        log.debug('RemoveHyperlink: Successfully completed');
      } catch (error) {
        log.error('Error in RemoveHyperlink:', error);
      }
    };
}
