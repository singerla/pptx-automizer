import { ElementInfo, PlaceholderInfo, XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { ShapeCoordinates } from '../types/shape-types';

export default class XmlPlaceholderHelper {
  static resetPlaceholderToDefaults(element: XmlElement, layoutPlaceholder: PlaceholderInfo): void {
    // Get the placeholder element
    const ph = element.getElementsByTagName('p:ph').item(0);

    // Set the index to match the layout placeholder
    ph.setAttribute('idx', String(layoutPlaceholder.idx));

    // Reset all positioning information
    const xfrm = element.getElementsByTagName('a:xfrm').item(0);
    if (xfrm) {
      XmlHelper.remove(xfrm);
    }
  }

  static removePlaceholder(element: XmlElement): void {
    const ph = element.getElementsByTagName('p:ph').item(0);
    XmlHelper.remove(ph);

    XmlPlaceholderHelper.assertShapeCoordinates(element, {
      x: 10,
      y: 10,
      w: 5000000,
      h: 1000000
    })
  }

  /**
   * Finds the best matching target placeholder for a source placeholder based on multiple criteria.
   *
   * @returns The best matching target placeholder or null if no suitable match found
   * @param sourceElement
   * @param typeMatches
   */
  static findBestTargetPlaceholder(
    sourceElement: ElementInfo,
    typeMatches: PlaceholderInfo[],
  ): PlaceholderInfo | null {
    const sourcePlaceholder = sourceElement.placeholder

    // Score each potential match based on multiple criteria
    const scoredMatches = typeMatches.map(target => {
      let score = 0;

      // 1. Same size gets a high score
      if (target.sz === sourcePlaceholder.sz) {
        score += 50;
      } else {
        // Smaller difference in size is better
        const sizeDiff = Math.abs(
          parseInt(target.sz || "0") - parseInt(sourcePlaceholder.sz || "0")
        );
        // Inverse relationship - smaller difference gets higher score
        score += Math.max(0, 30 - (sizeDiff / 100));
      }

      // 2. Similar idx values get a small bonus
      // This is lower priority but can be a tiebreaker
      const idxDiff = Math.abs(sourcePlaceholder.idx - target.idx);
      score += Math.max(0, 10 - idxDiff);

      return { target, score };
    });

    // Sort by score, highest first
    scoredMatches.sort((a, b) => b.score - a.score);

    // Return the highest scoring match
    return scoredMatches.length > 0 ? scoredMatches[0].target : null;
  }

  /**
   * Adds or updates coordinates in a shape element
   * @param element The XML element of the shape
   * @param coords The coordinates to set
   */
  static assertShapeCoordinates(element: XmlElement, coords: ShapeCoordinates): void {
    // Find or create the transform element
    let xfrm = element.getElementsByTagName('a:xfrm').item(0) ||
      element.getElementsByTagName('p:xfrm').item(0);

    const spPr = element.getElementsByTagName('p:spPr').item(0);

    if (!spPr) {
      return; // Cannot add coordinates without spPr element
    }

    if (!xfrm) {
      // Create a new transform element
      xfrm = element.ownerDocument.createElement('a:xfrm');
      spPr.appendChild(xfrm);
    }

    // Create or update the offset element (position)
    let off = xfrm.getElementsByTagName('a:off').item(0);
    if (!off) {
      off = element.ownerDocument.createElement('a:off');
      xfrm.appendChild(off);
      if (coords.x !== undefined) off.setAttribute('x', coords.x.toString());
      if (coords.y !== undefined) off.setAttribute('y', coords.y.toString());
    }

    // Create or update the extent element (size)
    let ext = xfrm.getElementsByTagName('a:ext').item(0);
    if (!ext) {
      ext = element.ownerDocument.createElement('a:ext');
      xfrm.appendChild(ext);
      if (coords.w !== undefined) ext.setAttribute('cx', coords.w.toString());
      if (coords.h !== undefined) ext.setAttribute('cy', coords.h.toString());
    }
  }
}
