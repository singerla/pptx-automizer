import { XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { log } from './logger';
import IArchive from '../interfaces/iarchive';
import { Target } from '../types/types';

/**
 * Hyperlink processing utilities
 */
export class HyperlinkProcessor {
  private static readonly HYPERLINK_TAG = 'a:hlinkClick';
  private static readonly RELATIONSHIP_ATTRIBUTE = 'r:id';

  // The only relationship Types an a:hlinkClick may legitimately resolve to:
  // an external hyperlink or an internal slide jump.
  private static readonly HYPERLINK_REL_TYPES = [
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  ];

  static isHyperlinkRelType(relType: string): boolean {
    return this.HYPERLINK_REL_TYPES.includes(relType);
  }

  /**
   * Finds all hyperlink elements within a given element
   */
  static findHyperlinks(element: XmlElement): XmlElement[] {
    const hyperlinks: XmlElement[] = [];
    
    try {
      // Find hyperlinks in shape properties (shape-level hyperlinks)
      const shapeProps = element.getElementsByTagName('p:cNvPr');
      for (let i = 0; i < shapeProps.length; i++) {
        const prop = shapeProps[i];
        const shapeHyperlinks = prop.getElementsByTagName(this.HYPERLINK_TAG);
        for (let j = 0; j < shapeHyperlinks.length; j++) {
          hyperlinks.push(shapeHyperlinks[j]);
        }
      }

      // Find hyperlinks in text runs (text-level hyperlinks)
      // Only look for hyperlinks that actually exist, don't process all text runs
      const allHyperlinks = element.getElementsByTagName(this.HYPERLINK_TAG);
      for (let i = 0; i < allHyperlinks.length; i++) {
        const hlink = allHyperlinks[i];
        // Check if this hyperlink is in a text run (has a:rPr as parent)
        const parent = hlink.parentNode;
        if (parent && parent.nodeName === 'a:rPr') {
          // Only add if not already added (avoid duplicates from shape properties)
          if (!hyperlinks.includes(hlink)) {
            hyperlinks.push(hlink);
          }
        }
      }
    } catch (error) {
      log.warn(`Error finding hyperlinks: ${error}`);
    }

    return hyperlinks;
  }

  /**
   * Checks if an element contains hyperlinks
   */
  static hasHyperlinks(element: XmlElement): boolean {
    return this.findHyperlinks(element).length > 0;
  }

  /**
   * Checks if an element contains multiple hyperlinks
   */
  static hasMultipleHyperlinks(element: XmlElement): boolean {
    return this.findHyperlinks(element).length > 1;
  }

  /**
   * Determines if an element should be processed as a hyperlink element
   * @param element - Element to analyze
   * @returns True if element should be processed as hyperlink
   */
  static shouldProcessAsHyperlink(element: XmlElement): boolean {
    const hyperlinks = this.findHyperlinks(element);
    
    // Single hyperlink elements can be processed as hyperlink shapes
    // Multiple hyperlinks (like tables) should be processed as generic shapes
    return hyperlinks.length === 1;
  }

  /**
   * Gets the primary hyperlink target from an element
   * @param element - Element to analyze
   * @returns Target information or null if no hyperlink found
   */
  static getPrimaryHyperlinkTarget(element: XmlElement): Target | null {
    try {
      const hyperlinks = this.findHyperlinks(element);
      
      if (hyperlinks.length === 0) {
        return null;
      }

      // Return the first hyperlink's target
      const firstHyperlink = hyperlinks[0];
      const rId = firstHyperlink.getAttribute(this.RELATIONSHIP_ATTRIBUTE);
      
      if (!rId) {
        return null;
      }

      return {
        rId,
        type: 'hyperlink'
      } as Target;
    } catch (error) {
      log.warn(`Error getting primary hyperlink target: ${error}`);
      return null;
    }
  }

  /**
   * Extracts hyperlink relationship IDs from an element
   * @param element - The XML element to extract from
   * @returns Array of relationship IDs
   */
  static extractHyperlinkRelationshipIds(element: XmlElement): string[] {
    const hyperlinks = this.findHyperlinks(element);
    return hyperlinks
      .map(hlink => hlink.getAttribute(this.RELATIONSHIP_ATTRIBUTE))
      .filter((rId): rId is string => rId !== null);
  }

  /**
   * Updates hyperlink relationship IDs in an element
   */
  static updateHyperlinkRelationshipIds(
    element: XmlElement,
    relationshipMap: Map<string, string>
  ): void {
    try {
      const hyperlinks = this.findHyperlinks(element);
      
      hyperlinks.forEach(hlink => {
        const currentRId = hlink.getAttribute(this.RELATIONSHIP_ATTRIBUTE);
        if (currentRId && relationshipMap.has(currentRId)) {
          const newRId = relationshipMap.get(currentRId);
          if (newRId) {
            hlink.setAttribute(this.RELATIONSHIP_ATTRIBUTE, newRId);
          }
        }
      });
    } catch (error) {
      log.warn(`Error updating hyperlink relationship IDs: ${error}`);
    }
  }

  /**
   * Processes hyperlinks for single-hyperlink elements
   */
  static async processSingleHyperlink(element: XmlElement, newRid: string): Promise<void> {
    const hyperlinks = this.findHyperlinks(element);

    // Only process if there's exactly one hyperlink
    if (hyperlinks.length !== 1) {
      log.warn(`Expected single hyperlink, found ${hyperlinks.length}`);
      return;
    }

    // Update the single hyperlink with the new relationship ID
    const hyperlink = hyperlinks[0];
    hyperlink.setAttribute(this.RELATIONSHIP_ATTRIBUTE, newRid);
  }

  /**
   * Copies multiple hyperlinks from source to target slide
   */
  static async copyMultipleHyperlinks(
    element: XmlElement,
    sourceArchive: IArchive,
    sourceSlideNumber: number,
    targetArchive: IArchive,
    targetSlideRelFile: string
  ): Promise<void> {
    if (!this.hasHyperlinks(element)) {
      return;
    }

    const hyperlinkRIds = this.extractHyperlinkRelationshipIds(element);
    if (hyperlinkRIds.length === 0) {
      return;
    }

    const sourceRelPath = `ppt/slides/_rels/slide${sourceSlideNumber}.xml.rels`;
    const sourceRelDoc = await XmlHelper.getXmlFromArchive(sourceArchive, sourceRelPath);
    
    if (!sourceRelDoc) {
      log.warn(`Source relationships not found: ${sourceRelPath}`);
      return;
    }

    const sourceRelationships = sourceRelDoc.getElementsByTagName('Relationship');
    const targetRelXml = await XmlHelper.getXmlFromArchive(targetArchive, targetSlideRelFile);
    
    if (!targetRelXml) {
      log.warn(`Target relationships not found: ${targetSlideRelFile}`);
      return;
    }

    const relationshipMap = new Map<string, string>();
    const processedTargets = new Set<string>();

    for (let i = 0; i < hyperlinkRIds.length; i++) {
      const rId = hyperlinkRIds[i];
      
      let sourceRel: XmlElement | null = null;
      for (let j = 0; j < sourceRelationships.length; j++) {
        if (sourceRelationships[j].getAttribute('Id') === rId) {
          sourceRel = sourceRelationships[j];
          break;
        }
      }

      if (sourceRel) {
        const relType = sourceRel.getAttribute('Type');
        const target = sourceRel.getAttribute('Target');
        const targetMode = sourceRel.getAttribute('TargetMode');

        // A stale r:id can collide with any relationship on the source slide
        // (rId1 is typically the slideLayout, rId2 often the notesSlide).
        // Cloning such a rel verbatim adds a second slideLayout/notesSlide
        // relationship to the target slide — a package-level corruption that
        // makes PowerPoint offer to repair the file.
        if (relType && target && this.isHyperlinkRelType(relType)) {
          const relationshipKey = `${relType}:${target}:${targetMode || ''}`;
          
          let newRId: string;
          
          if (processedTargets.has(relationshipKey)) {
            const existingRels = targetRelXml.getElementsByTagName('Relationship');
            for (let k = 0; k < existingRels.length; k++) {
              const existingRel = existingRels[k];
              if (existingRel.getAttribute('Type') === relType && 
                  existingRel.getAttribute('Target') === target &&
                  existingRel.getAttribute('TargetMode') === targetMode) {
                newRId = existingRel.getAttribute('Id') || '';
                break;
              }
            }
          } else {
            newRId = await XmlHelper.getNextRelId(targetArchive, targetSlideRelFile);
            
            const newRelationship = targetRelXml.createElement('Relationship');
            newRelationship.setAttribute('Id', newRId);
            newRelationship.setAttribute('Type', relType);
            newRelationship.setAttribute('Target', XmlHelper.sanitizeAttr(target));
            
            if (targetMode) {
              newRelationship.setAttribute('TargetMode', targetMode);
            }

            targetRelXml.documentElement.appendChild(newRelationship);
            processedTargets.add(relationshipKey);
          }

          relationshipMap.set(rId, newRId);
        } else if (relType && !this.isHyperlinkRelType(relType)) {
          log.warn(
            `Hyperlink r:id="${rId}" resolves to a non-hyperlink relationship ` +
              `(${relType}) on source slide ${sourceSlideNumber} — not copied`,
          );
        }
      }
    }

    // Strip hyperlinks whose relationship could not be copied (no source rel,
    // or a rel of the wrong Type). Leaving their r:id in place would either
    // dangle or point at an unrelated relationship on the target slide —
    // both are repair-prompt triggers.
    this.findHyperlinks(element).forEach((hlink) => {
      const rId = hlink.getAttribute(this.RELATIONSHIP_ATTRIBUTE);
      if (rId && !relationshipMap.has(rId)) {
        log.warn(`Dropping hyperlink with unresolvable r:id="${rId}"`);
        hlink.parentNode?.removeChild(hlink);
      }
    });

    this.updateHyperlinkRelationshipIds(element, relationshipMap);
    await XmlHelper.writeXmlToArchive(targetArchive, targetSlideRelFile, targetRelXml);
  }
} 
