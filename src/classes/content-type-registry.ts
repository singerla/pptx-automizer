import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import IArchive from '../interfaces/iarchive';
import {
  RelationshipAttribute,
  SlideListAttribute,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import type HasShapes from './has-shapes';

/**
 * Per-document caches that keep `addToPresentation` linear in the number of
 * appended slides. Keyed by the parsed document object: when an archive
 * evicts and re-parses a part, the fresh document misses the cache and the
 * state is rebuilt by a single scan.
 *
 * The rId cache stays a safe upper bound because only this class appends
 * relationships to `ppt/_rels/presentation.xml.rels`; the one other writer,
 * `Template.removeSlideRelations` (removeExistingSlides), only removes.
 */
const nextRelIdByRelsDoc = new WeakMap<XmlDocument, number>();
const listElementsByPresentationDoc = new WeakMap<
  XmlDocument,
  Map<string, XmlElement>
>();

/**
 * Registers appended parts with the presentation: relationship in
 * `ppt/_rels/presentation.xml.rels`, entry in the slide/master list of
 * `ppt/presentation.xml`, and override in `[Content_Types].xml`.
 *
 * Extracted from HasShapes (ROADMAP Phase 2).
 */
export class ContentTypeRegistry {
  constructor(private host: HasShapes) {}

  /**
   * Registers the host's target part (slide, slideMaster or slideLayout)
   * with the presentation.
   */
  async addToPresentation(): Promise<void> {
    const host = this.host;
    const relId = await this.getNextPresentationRelId(host.targetArchive);
    await this.appendToSlideRel(host.targetArchive, relId, host.targetNumber);

    if (host.targetType === 'slide') {
      await this.appendToSlideList(host.targetArchive, relId);
    } else if (host.targetType === 'slideMaster') {
      await this.appendToSlideMasterList(host.targetArchive, relId);
    } else if (host.targetType === 'slideLayout') {
      // No changes to ppt/presentation.xml required for slideLayouts
    }

    await this.appendToContentType(host.targetArchive, host.targetNumber);
  }

  /**
   * Cached equivalent of `XmlHelper.getNextRelId` for
   * `ppt/_rels/presentation.xml.rels`: the max-rId scan runs once per parsed
   * document, afterwards ids are incremented locally (the scan is O(n) in
   * appended slides, so rescanning per append made large decks O(n²)).
   */
  private async getNextPresentationRelId(
    rootArchive: IArchive,
  ): Promise<string> {
    const xml = await XmlHelper.getXmlFromArchive(
      rootArchive,
      PptPaths.presentationRels,
    );
    const max =
      nextRelIdByRelsDoc.get(xml) ??
      XmlHelper.getMaxId(xml.documentElement.childNodes, 'Id');
    const next = max + 1;
    nextRelIdByRelsDoc.set(xml, next);
    return `rId${next}-created`;
  }

  /**
   * Cached lookup of a list element of `ppt/presentation.xml`
   * (`p:sldIdLst` / `p:sldMasterIdLst`). xmldom's live
   * `getElementsByTagName` re-walks the whole document on every access.
   */
  private getPresentationListElement(
    xml: XmlDocument,
    tag: string,
  ): XmlElement | undefined {
    let byTag = listElementsByPresentationDoc.get(xml);
    if (!byTag) {
      byTag = new Map();
      listElementsByPresentationDoc.set(xml, byTag);
    }
    let element = byTag.get(tag);
    if (!element || !element.parentNode) {
      element = xml.getElementsByTagName(tag)[0];
      if (element) {
        byTag.set(tag, element);
      } else {
        byTag.delete(tag);
      }
    }
    return element;
  }

  appendToSlideRel(
    rootArchive: IArchive,
    relId: string,
    slideCount: number,
  ): Promise<XmlElement> {
    const targetType = this.host.targetType;
    return XmlHelper.append({
      archive: rootArchive,
      file: PptPaths.presentationRels,
      // `Relationships` is the document element of a .rels part.
      parent: (xml: XmlDocument) => xml.documentElement as XmlElement,
      tag: 'Relationship',
      attributes: {
        Id: relId,
        Type: `http://schemas.openxmlformats.org/officeDocument/2006/relationships/${targetType}`,
        Target: `${targetType}s/${targetType}${slideCount}.xml`,
      } as RelationshipAttribute,
    });
  }

  /**
   * Appends a new slide to slide list in presentation.xml.
   * If rootArchive has no slides, a new node will be created.
   * "id"-attribute of 'p:sldId'-element must be greater than 255.
   */
  appendToSlideList(rootArchive: IArchive, relId: string): Promise<XmlElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: PptPaths.presentation,
      assert: async (xml: XmlDocument) => {
        if (!this.getPresentationListElement(xml, 'p:sldIdLst')) {
          XmlHelper.insertAfter(
            xml.createElement('p:sldIdLst'),
            this.getPresentationListElement(xml, 'p:sldMasterIdLst'),
          );
        }
      },
      parent: (xml: XmlDocument) =>
        this.getPresentationListElement(xml, 'p:sldIdLst'),
      tag: 'p:sldId',
      attributes: {
        'r:id': relId,
      } as SlideListAttribute,
    });
  }

  /**
   * Appends a new slideMaster to the master list in presentation.xml.
   */
  appendToSlideMasterList(
    rootArchive: IArchive,
    relId: string,
  ): Promise<XmlElement> {
    return XmlHelper.append({
      archive: rootArchive,
      file: PptPaths.presentation,
      parent: (xml: XmlDocument) =>
        this.getPresentationListElement(xml, 'p:sldMasterIdLst'),
      tag: 'p:sldMasterId',
      attributes: {
        'r:id': relId,
      } as SlideListAttribute,
    });
  }

  appendToContentType(
    rootArchive: IArchive,
    count: number,
  ): Promise<XmlElement> {
    const targetType = this.host.targetType;
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(rootArchive, {
        PartName: PptPaths.partName(PptPaths.part(targetType, count)),
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.${targetType}+xml`,
      }),
    );
  }

  appendNotesToContentType(slideCount: number): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.host.targetArchive, {
        PartName: PptPaths.partName(PptPaths.notesSlide(slideCount)),
        ContentType: `application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml`,
      }),
    );
  }

  appendThemeToContentType(themeCount: string | number): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.host.targetArchive, {
        PartName: PptPaths.partName(PptPaths.theme(themeCount)),
        ContentType: `application/vnd.openxmlformats-officedocument.theme+xml`,
      }),
    );
  }
}
