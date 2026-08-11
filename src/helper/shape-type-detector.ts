import { ElementType } from '../enums/element-type';
import IArchive from '../interfaces/iarchive';
import { AnalyzedElementType } from '../types/types';
import { XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { HyperlinkProcessor } from './hyperlink-processor';
import { log } from './logger';

export type ShapeTypeDetector = {
  /**
   * Returns true if the given source element is of this shape type.
   * Detectors are evaluated in registry order; first match wins.
   */
  match: (element: XmlElement) => boolean;
  type: ElementType;
  /**
   * Relationship lookup key for XmlHelper.getTargetByRelId; the detected
   * element's related target (chart xml, media file, …) is resolved with it.
   * Omit for shape types without related content.
   */
  relType?: string;
  /**
   * Custom analysis for shape types that need more than a rel lookup.
   * Overrides the default `relType` handling.
   */
  analyze?: (
    element: XmlElement,
    sourceArchive: IArchive,
    relsPath: string,
  ) => Promise<AnalyzedElementType>;
};

const containsTag = (tag: string) => (element: XmlElement) =>
  element.getElementsByTagName(tag).length > 0;

/**
 * Registry of shape type detectors, replacing the former if-chain in
 * `HasShapes.analyzeElement`. To support a new shape type (e.g. video/audio),
 * add one entry here — order matters, more specific detectors first.
 */
export const shapeTypeDetectors: ShapeTypeDetector[] = [
  {
    match: containsTag('c:chart'),
    type: ElementType.Chart,
    relType: 'chart',
  },
  {
    match: containsTag('cx:chart'),
    type: ElementType.Chart,
    relType: 'chartEx',
  },
  {
    match: containsTag('p:nvPicPr'),
    type: ElementType.Image,
    relType: 'image',
  },
  {
    match: containsTag('dgm:relIds'),
    type: ElementType.Diagram,
    relType: 'diagram',
  },
  {
    match: containsTag('p:oleObj'),
    type: ElementType.OLEObject,
    relType: 'oleObject',
  },
  {
    match: (element) => HyperlinkProcessor.hasHyperlinks(element),
    type: ElementType.Hyperlink,
    analyze: async (element, sourceArchive, relsPath) => {
      // Elements with multiple hyperlinks (like tables) are treated as
      // generic shapes; GenericShape copies the hyperlink relationships.
      if (HyperlinkProcessor.hasMultipleHyperlinks(element)) {
        return { type: ElementType.Shape };
      }
      try {
        const target = await XmlHelper.getTargetByRelId(
          sourceArchive,
          relsPath,
          element,
          'hyperlink',
        );
        return {
          type: ElementType.Hyperlink,
          target,
          element,
        };
      } catch (error) {
        log.warn('Error finding hyperlink target:', error);
        return { type: ElementType.Shape };
      }
    },
  },
];

/**
 * Detects the shape type of a source element and resolves its related
 * target (if any). Falls back to a generic shape when no detector matches.
 */
export async function analyzeElement(
  sourceElement: XmlElement,
  sourceArchive: IArchive,
  relsPath: string,
): Promise<AnalyzedElementType> {
  for (const detector of shapeTypeDetectors) {
    if (!detector.match(sourceElement)) {
      continue;
    }

    if (detector.analyze) {
      return detector.analyze(sourceElement, sourceArchive, relsPath);
    }

    return {
      type: detector.type,
      target: detector.relType
        ? await XmlHelper.getTargetByRelId(
            sourceArchive,
            relsPath,
            sourceElement,
            detector.relType,
          )
        : undefined,
    };
  }

  return {
    type: ElementType.Shape,
  };
}
