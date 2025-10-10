import {
  ElementInfo,
  ElementPosition,
  PlaceholderInfo,
  PlaceholderType,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { XmlSlideHelper } from './xml-slide-helper';

export default class XmlPlaceholderHelper {
  private static mapAlternativePlaceholders = {
    // Title-related placeholders
    title: ['ctrTitle', 'subTitle', 'body'],
    ctrTitle: ['title', 'subTitle', 'body'],
    subTitle: ['title', 'ctrTitle', 'body'],

    // Content placeholders
    body: ['title', 'ctrTitle', 'subTitle'],

    // Media and visual content
    pic: ['media', 'obj', 'clipArt', 'bitmap'],
    media: ['pic', 'obj', 'clipArt', 'bitmap'],
    obj: ['pic', 'media', 'clipArt', 'bitmap'],
    clipArt: ['pic', 'media', 'obj', 'bitmap'],
    bitmap: ['pic', 'media', 'obj', 'clipArt'],

    // Data visualization
    chart: ['tbl', 'dgm', 'orgChart', 'obj'],
    tbl: ['chart', 'dgm', 'orgChart'],
    dgm: ['chart', 'orgChart', 'tbl', 'obj'],
    orgChart: ['dgm', 'chart', 'tbl', 'obj'],

    // Footer elements
    ftr: ['dt', 'sldNum', 'hdr'],
    dt: ['ftr', 'sldNum', 'hdr'],
    sldNum: ['ftr', 'dt', 'hdr'],
    hdr: ['ftr', 'dt', 'sldNum'],

    // Fallback for unknown
    unknown: ['body', 'obj', 'pic'],
  };

  /**
   * Extracts placeholder information from an XML element.
   *
   * This method parses a PowerPoint shape element to extract its placeholder properties,
   * including placeholder type, size, index, position, and element type. It also attempts to merge
   * position information from layout placeholders when available.
   *
   * @param element - The XML element representing a PowerPoint shape
   * @param layoutPlaceholders - Optional array of placeholder info from the slide layout
   * @returns PlaceholderInfo object containing all extracted placeholder data, or undefined if no placeholder found
   */
  static getPlaceholderInfo(
    element: XmlElement,
    layoutPlaceholders?: PlaceholderInfo[],
  ): PlaceholderInfo | undefined {
    // Find the placeholder element within the shape
    const placeholderElement = element.getElementsByTagName('p:ph').item(0);

    // Early return if this element doesn't contain a placeholder
    if (!placeholderElement) {
      return undefined;
    }

    // Extract the shape properties element for determining element type
    const slideElementParent = element.getElementsByTagName('p:spPr').item(0)
      ?.parentNode as XmlElement;

    // Parse the placeholder index attribute
    const indexAttribute = placeholderElement.getAttribute('idx');
    const placeholderIndex = indexAttribute?.length
      ? parseInt(indexAttribute, 10)
      : null;

    // Try to find corresponding layout placeholder by matching index
    const matchingLayoutPlaceholder = layoutPlaceholders?.find(
      (layoutPh) =>
        placeholderIndex !== null && layoutPh.idx === placeholderIndex,
    );

    // Determine position - prefer shape's own position, fallback to layout position
    const shapePosition = XmlSlideHelper.parseShapeCoordinates(element, false);
    const finalPosition = shapePosition || matchingLayoutPlaceholder?.position;

    const elementType = XmlSlideHelper.getElementType(slideElementParent);
    const phType = placeholderElement.getAttribute('type') as PlaceholderType;
    const type = !phType && elementType === 'sp' ? 'body' : phType;

    // Build the placeholder info object
    const placeholderInfo: PlaceholderInfo = {
      type: type || 'unknown',
      sz: placeholderElement.getAttribute('sz'),
      idx: placeholderIndex,
      elementType,
      position: finalPosition,
    };

    return placeholderInfo;
  }

  static setPlaceholderDefaults(
    element: XmlElement,
    layoutPlaceholder: PlaceholderInfo,
  ): void {
    // Get the placeholder element
    const ph = element.getElementsByTagName('p:ph').item(0);

    if (ph && layoutPlaceholder.idx) {
      // Set the index to match the layout placeholder
      ph.setAttribute('idx', String(layoutPlaceholder.idx));
    }

    // Reset all positioning information
    const xfrm = element.getElementsByTagName('a:xfrm').item(0);
    if (xfrm) {
      XmlHelper.remove(xfrm);
    }
  }

  static removePlaceholder(
    element: XmlElement,
    fallbackPosition: ElementPosition,
  ): void {
    const ph = element.getElementsByTagName('p:ph').item(0);
    XmlHelper.remove(ph);

    XmlPlaceholderHelper.assertShapeCoordinates(element, fallbackPosition);
  }

  /**
   * Finds the best matching target placeholder for a source placeholder based on multiple criteria.
   *
   * @returns The best matching target placeholder or null if no suitable match found
   * @param sourceElement
   * @param typeMatches
   */
  static findBestMatchingPlaceholder(
    sourceElement: ElementInfo,
    typeMatches: PlaceholderInfo[],
  ): PlaceholderInfo | null {
    // Score each potential match based on multiple criteria
    const scoredMatches = typeMatches.map((target) => {
      const score = XmlPlaceholderHelper.calculatePlaceholderSimilarityScore(
        0,
        target,
        sourceElement,
      );
      return { target, score };
    });

    // Sort by score, highest first
    scoredMatches.sort((a, b) => b.score - a.score);

    // Return the highest scoring match
    return scoredMatches.length > 0 ? scoredMatches[0].target : null;
  }

  static findBestTargetPlaceholderAlternative(
    element: ElementInfo,
    targetPlaceholders: PlaceholderInfo[],
    usedPlaceholders: PlaceholderInfo[],
  ): PlaceholderInfo {
    const originalType = element.placeholder.type;
    const alternatives = this.mapAlternativePlaceholders[originalType] || [];

    let bestMatch = null;
    let bestScore = -1;

    // Try to find the best alternative placeholder in the target layout
    for (const alternativeType of alternatives) {
      // Look for available placeholders of this alternative type
      const availablePlaceholder = targetPlaceholders.find(
        (ph) => ph.type === alternativeType && !usedPlaceholders.includes(ph),
      );

      if (availablePlaceholder) {
        const initScore =
          alternatives.length - alternatives.indexOf(alternativeType);
        const score = XmlPlaceholderHelper.calculatePlaceholderSimilarityScore(
          initScore,
          availablePlaceholder,
          element,
        );

        if (score > bestScore) {
          bestScore = score;
          bestMatch = availablePlaceholder;
        }
      }
    }

    return bestMatch;
  }

  static calculatePlaceholderSimilarityScore(
    score: number,
    availablePlaceholder: PlaceholderInfo,
    element: ElementInfo,
  ) {
    // Bonus points for matching size if available
    if (
      element.placeholder.sz &&
      availablePlaceholder.sz === element.placeholder.sz
    ) {
      score += 10;
    }

    // Bonus points for similar position if available
    if (element.placeholder.position && availablePlaceholder.position) {
      const distanceScore = Math.max(
        0,
        100 -
          Math.sqrt(
            Math.pow(
              element.placeholder.position.x - availablePlaceholder.position.x,
              2,
            ) +
              Math.pow(
                element.placeholder.position.y -
                  availablePlaceholder.position.y,
                2,
              ),
          ) /
            1000,
      );
      score += distanceScore;
    }
    return score;
  }

  /**
   * Adds or updates coordinates in a shape element
   * @param element The XML element of the shape
   * @param coords The coordinates to set
   */
  static assertShapeCoordinates(
    element: XmlElement,
    coords: ElementPosition,
  ): void {
    // Find or create the transform element
    let xfrm =
      element.getElementsByTagName('a:xfrm').item(0) ||
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
      if (coords.cx !== undefined) ext.setAttribute('cx', coords.cx.toString());
      if (coords.cy !== undefined) ext.setAttribute('cy', coords.cy.toString());
    }
  }
}
