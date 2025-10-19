import {
  ElementInfo,
  ElementPosition,
  LayoutInfo,
  PlaceholderInfo,
  PlaceholderMappingResult,
  PlaceholderType,
  XmlElement,
} from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { XmlSlideHelper } from './xml-slide-helper';
import { ShapeModificationCallback } from '../types/types';
import { ISlide } from '../interfaces/islide';
import { vd } from './general-helper';

export type EnrichedElementInfo = ElementInfo & {
  fromTopRank: number;
  fromLeftRank: number;
  sizeRank: number;
};
export type GroupedByType = Record<ElementInfo['type'], EnrichedElementInfo[]>;

export default class XmlPlaceholderHelper {
  private mapAlternativePlaceholders = {
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

  slide: ISlide;
  slideElements: ElementInfo[];
  sourceLayoutInfo: LayoutInfo;
  targetPlaceholders: PlaceholderInfo[];
  mappingResult: PlaceholderMappingResult = {
    usedPlaceholders: [],
    unmatchedSourcePlaceholderElements: [],
    matchedSourceElements: [],
  };

  constructor(
    slide: ISlide,
    slideElements: ElementInfo[],
    sourceLayoutInfo: LayoutInfo,
    targetPlaceholders: PlaceholderInfo[],
  ) {
    this.slide = slide;
    this.slideElements = slideElements;
    this.sourceLayoutInfo = sourceLayoutInfo;
    this.targetPlaceholders = targetPlaceholders;
  }

  run() {
    this.performInitialPlaceholderMatching();
    this.handleUnmatchedPlaceholderElements();
    this.handleForceAssignmentPlaceholders([
      'title',
      'ctrTitle',
      'subTitle',
      'body',
      'pic',
    ]);
    this.cleanupUnmatchedPlaceholders();
  }

  /**
   * Performs the initial placeholder matching between source elements and target placeholders.
   * Elements with exact placeholder type matches are processed first.
   */
  public performInitialPlaceholderMatching(): void {
    this.slideElements.forEach((element: ElementInfo) => {
      if (element.placeholder?.type) {
        const matchesPlaceholder = this.applyPlaceholderToElement(
          this.targetPlaceholders,
          element,
        );
        if (!matchesPlaceholder) {
          this.mappingResult.unmatchedSourcePlaceholderElements.push(element);
        } else {
          this.mappingResult.matchedSourceElements.push(element);
        }
      }
    });
  }

  /**
   * Handles placeholder elements on the source slide that couldn't be matched in the
   * initial pass by finding alternative placeholder matches using best-fit algorithms.
   */
  public handleUnmatchedPlaceholderElements(): void {
    const unmatchedElements = this.mappingResult
      .unmatchedSourcePlaceholderElements as ElementInfo[];

    unmatchedElements.forEach((element) => {
      const bestAlternativeMatch =
        this.findBestTargetPlaceholderAlternative(element);

      if (bestAlternativeMatch) {
        this.applyPlaceholder(element, bestAlternativeMatch);
        this.removeElementFromUnmatched(element, unmatchedElements);
        this.mappingResult.matchedSourceElements.push(element);
      }
    });
  }

  /**
   * Removes an element from the unmatched elements array.
   *
   * @param element - Element to remove
   * @param unmatchedElements - Array to remove from
   * @private
   */
  private removeElementFromUnmatched(
    element: ElementInfo,
    unmatchedElements: ElementInfo[],
  ): ElementInfo[] {
    const index = unmatchedElements.indexOf(element);
    if (index > -1) {
      unmatchedElements.splice(index, 1);
    }
    return unmatchedElements;
  }

  /**
   * Handles force assignment of placeholder types by finding the best matching
   * unmatched slide elements for unassigned target placeholders.
   *
   * @param forceAssignPhTypes - Array of placeholder types that should be force assigned
   */
  public handleForceAssignmentPlaceholders(forceAssignPhTypes: string[]): void {
    const mappingResult = this.mappingResult;
    const usedElements = [];

    forceAssignPhTypes.forEach((phType) => {
      // Find unassigned target placeholders of this type
      const unassignedTargetPlaceholders = this.targetPlaceholders.filter(
        (ph) =>
          ph.type === phType && !mappingResult.usedPlaceholders.includes(ph),
      );
      if (unassignedTargetPlaceholders.length === 0) {
        return; // No unassigned placeholders of this type
      }
      const unmatchedElements = this.slideElements.filter((ele) => {
        return !mappingResult.matchedSourceElements.includes(ele);
      });

      // Recalculate candidate elements for each placeholder to reflect current state
      const unmatchedElementsGroups =
        XmlPlaceholderHelper.groupUnmatchedElements(unmatchedElements);

      unassignedTargetPlaceholders.forEach((ph) => {
        const candidateElements = unmatchedElementsGroups[ph.elementType] || [];
        const filteredCandidates = candidateElements.filter((candidate) => {
          return !usedElements.includes(candidate);
        });

        const bestCandidate = this.findBestCandidateElementForPlaceholder(
          filteredCandidates,
          ph,
        );

        if (bestCandidate) {
          this.applyForceAssignedPlaceholderToElement(
            bestCandidate,
            ph,
            unmatchedElements,
          );
          usedElements.push(bestCandidate);
        }
      });
    });
  }

  private findBestCandidateElementForPlaceholder(
    candidateElements: EnrichedElementInfo[],
    ph: PlaceholderInfo,
  ): ElementInfo {
    if (ph.type === 'title') {
      if (ph.position) {
        return this.findClosestCandidate(ph, candidateElements);
      }
      return this.findFromTopCandidate(candidateElements, 1, true);
    }

    // if (ph.type === 'subTitle' || ph.type === 'ctrTitle') {
    //   return this.findFromTopCandidate(candidateElements, 2, false);
    // }

    if (ph.type === 'body') {
      if (ph.position) {
        return this.findClosestCandidate(ph, candidateElements);
      }
      // return this.findLargestCandidate(candidateElements);
    }

    // if (ph.type === 'pic') {
    //   return this.findLargestCandidate(candidateElements);
    // }

    vd('no candidate for ');
    vd(ph);

    return null;
  }

  private findClosestCandidate(
    ph: PlaceholderInfo,
    candidateElements: EnrichedElementInfo[],
  ) {
    let highestScore = 0;
    let bestCandidate = null;
    candidateElements.forEach((ele) => {
      const closestShapeScore = XmlPlaceholderHelper.calculateDistanceScore(
        ele.position,
        ph.position,
      );
      if (closestShapeScore > highestScore) {
        highestScore = closestShapeScore;
        bestCandidate = ele;
      }
    });
    return bestCandidate;
  }

  private findLargestCandidate(candidateElements: EnrichedElementInfo[]) {
    let minSizeRank = Number.POSITIVE_INFINITY;
    candidateElements.forEach((ele) => {
      minSizeRank = ele.sizeRank < minSizeRank ? ele.sizeRank : minSizeRank;
    });
    return candidateElements.find((ele) => {
      return ele.sizeRank === minSizeRank;
    });
  }

  private findFromTopCandidate(
    candidateElements: EnrichedElementInfo[],
    fromTopRank: number,
    equals?: boolean,
  ) {
    const bestCandidate = candidateElements.find((ele) => {
      if (equals) {
        return ele.fromTopRank === fromTopRank;
      }
      return ele.fromTopRank >= fromTopRank;
    });
    if (bestCandidate) {
      return bestCandidate;
    }
  }

  /**
   * Cleans up elements that still don't have placeholder matches by removing
   * their placeholder properties and applying fallback positioning.
   */
  public cleanupUnmatchedPlaceholders(): void {
    const sourcePlaceholders = this.sourceLayoutInfo.placeholders;
    const unmatchedElements =
      this.mappingResult.unmatchedSourcePlaceholderElements;
    unmatchedElements.forEach((element) => {
      this.clearUnmatchedPlaceholder(element, sourcePlaceholders);
    });
  }

  clearUnmatchedPlaceholder(
    element: ElementInfo,
    sourcePlaceholders: PlaceholderInfo[],
  ) {
    const fallbackPh = sourcePlaceholders.find(
      (ph) => ph.idx === element.placeholder.idx,
    );
    const fallbackPosition = fallbackPh?.position || {
      x: 1000,
      y: 1000,
      cx: 5000000,
      cy: 1000000,
    };

    const callback = (element) => {
      XmlPlaceholderHelper.removePlaceholder(element, fallbackPosition);
    };
    this.postApplyModification(element, callback);
  }

  applyPlaceholderToElement(
    layoutPlaceholders: PlaceholderInfo[],
    element: ElementInfo,
  ): PlaceholderInfo {
    const unusedPlaceholders = layoutPlaceholders.filter(
      (ph) => !this.mappingResult.usedPlaceholders.includes(ph),
    );
    const matchPlaceholders = unusedPlaceholders.filter((ph) => {
      return ph.type === element.placeholder?.type;
    });

    if (matchPlaceholders.length) {
      const bestMatch = XmlPlaceholderHelper.findBestMatchingPlaceholder(
        element,
        matchPlaceholders,
      );
      this.applyPlaceholder(element, bestMatch);
      return bestMatch;
    }

    return null;
  }

  applyPlaceholder(element: ElementInfo, bestMatch: PlaceholderInfo) {
    const callback = (element: XmlElement) => {
      XmlPlaceholderHelper.setPlaceholderDefaults(element, bestMatch);
    };
    this.postApplyModification(element, callback);
    this.mappingResult.usedPlaceholders.push(bestMatch);
  }

  applyForceAssignedPlaceholderToElement(
    bestCandidate: ElementInfo,
    bestMatch: PlaceholderInfo,
    unmatchedElements: ElementInfo[],
  ) {
    this.applyPlaceholder(bestCandidate, bestMatch);
    this.removeElementFromUnmatched(bestCandidate, unmatchedElements);
    this.mappingResult.matchedSourceElements.push(bestCandidate);
  }

  postApplyModification(
    element: ElementInfo,
    callback: ShapeModificationCallback,
  ) {
    const selector = {
      creationId: element.creationId,
      nameIdx: element.nameIdx,
      name: element.name,
    };
    this.slide.modifyElement(selector, callback);
  }

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
    let ph = element.getElementsByTagName('p:ph').item(0);
    if (!ph) {
      // If element has no placeholder, create one
      const nvPr = element.getElementsByTagName('p:nvPr').item(0);
      if (nvPr) {
        ph = element.ownerDocument.createElement('p:ph');

        if (layoutPlaceholder.type) {
          ph.setAttribute('type', layoutPlaceholder.type);
        }
        if (layoutPlaceholder.sz) {
          ph.setAttribute('sz', layoutPlaceholder.sz);
        }
        if (layoutPlaceholder.idx) {
          ph.setAttribute('idx', String(layoutPlaceholder.idx));
        }
        nvPr.appendChild(ph);
      }
    }

    if (ph && layoutPlaceholder.idx) {
      // Set the index to match the layout placeholder
      ph.setAttribute('idx', String(layoutPlaceholder.idx));
    }

    // Force fallback to layout properties
    ['a:xfrm', 'p:style'].forEach(tag => {
      const ele = element.getElementsByTagName(tag).item(0);
      if (ele) {
        XmlHelper.remove(ele);
      }
    });
    ['p:spPr', 'a:bodyPr', 'p:cNvSpPr'].forEach(tag => {
      const ele = element.getElementsByTagName(tag).item(0);
      if (ele) {
        XmlHelper.removeAllChildren(ele);
      }
    })
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

  private findBestTargetPlaceholderAlternative(
    element: ElementInfo,
  ): PlaceholderInfo {
    const originalType = element.placeholder.type;
    const alternatives = this.mapAlternativePlaceholders[originalType] || [];

    let bestMatch = null;
    let bestScore = -1;

    // Try to find the best alternative placeholder in the target layout
    for (const alternativeType of alternatives) {
      // Look for available placeholders of this alternative type
      const availablePlaceholder = this.targetPlaceholders.find(
        (ph) =>
          ph.type === alternativeType &&
          !this.mappingResult.usedPlaceholders.includes(ph),
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
      element.placeholder?.sz &&
      availablePlaceholder.sz === element.placeholder.sz
    ) {
      score += 10;
    }

    // Bonus points for similar position if available
    const position = element.placeholder?.position || element.position;
    if (position && availablePlaceholder.position) {
      score += this.calculateDistanceScore(
        position,
        availablePlaceholder.position,
      );
    }
    return score;
  }

  static calculateDistanceScore(pos1: ElementPosition, pos2: ElementPosition) {
    // Calculate center points of both elements
    const center1X = pos1.x + pos1.cx / 2;
    const center1Y = pos1.y + pos1.cy / 2;
    const center2X = pos2.x + pos2.cx / 2;
    const center2Y = pos2.y + pos2.cy / 2;

    // Calculate distance between center points
    const distance =
      Math.sqrt(
        Math.pow(center1X - center2X, 2) + Math.pow(center1Y - center2Y, 2),
      ) / 100000;

    // Convert distance to score - closer elements get higher scores
    const distanceScore = Math.max(0, 10 - distance); // Max 10 points for very close elements

    // Calculate size similarity bonus (area comparison)
    const area1 = pos1.cx * pos1.cy;
    const area2 = pos2.cx * pos2.cy;
    const sizeRatio = Math.min(area1, area2) / Math.max(area1, area2);
    const sizeBonus = sizeRatio * 10; // Up to 10 points for identical sizes

    return distanceScore + sizeBonus;
  }

  /**
   * @param unmatchedElements - Elements that couldn't be matched initially
   * @returns groupedByType
   * @private
   */
  static groupUnmatchedElements(
    unmatchedElements: ElementInfo[],
  ): GroupedByType {
    const groupedByType = {} as GroupedByType;

    unmatchedElements.forEach((element: EnrichedElementInfo) => {
      if (!groupedByType[element.type]) {
        groupedByType[element.type] = [];
      }

      element.fromTopRank = 0;
      element.fromLeftRank = 0;
      element.sizeRank = 0;

      groupedByType[element.type].push(element);
    });

    // Process each group
    Object.keys(groupedByType).forEach((shapeType) => {
      const elementsOfType = groupedByType[shapeType as ElementInfo['type']];

      // Sort by position from top-left to bottom-right for ranking
      const sortedByPosition = [...elementsOfType].sort((a, b) => {
        // First sort by Y position (top to bottom)
        if (a.position.y !== b.position.y) {
          return a.position.y - b.position.y;
        }
        // Then sort by X position (left to right)
        return a.position.x - b.position.x;
      });

      // Sort by size (area) for size ranking
      const sortedBySize = [...elementsOfType].sort((a, b) => {
        const areaA = a.position.cx * a.position.cy;
        const areaB = b.position.cx * b.position.cy;
        return areaB - areaA; // Descending order (largest first)
      });

      // Create ranking maps
      const topRankMap = new Map<string, number>();
      const leftRankMap = new Map<string, number>();
      const sizeRankMap = new Map<string, number>();

      // Assign fromTopRank and fromLeftRank based on position sorting
      sortedByPosition.forEach((element, index) => {
        const elementKey = element.creationId || element.name + element.nameIdx;
        topRankMap.set(elementKey, index + 1);
        leftRankMap.set(elementKey, index + 1);
      });

      // Assign sizeRank based on size sorting
      sortedBySize.forEach((element, index) => {
        const elementKey = element.creationId || element.name + element.nameIdx;
        sizeRankMap.set(elementKey, index + 1);
      });

      // Create enriched elements with rankings
      groupedByType[shapeType as ElementInfo['type']] = elementsOfType.map(
        (element) => {
          const elementKey =
            element.creationId || element.name + element.nameIdx;

          element.fromTopRank = topRankMap.get(elementKey) || 0;
          element.fromLeftRank = leftRankMap.get(elementKey) || 0;
          element.sizeRank = sizeRankMap.get(elementKey) || 0;

          return element;
        },
      );
    });

    return groupedByType;
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
