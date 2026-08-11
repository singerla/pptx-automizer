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
    const relId = await XmlHelper.getNextRelId(
      host.targetArchive,
      PptPaths.presentationRels,
    );
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

  appendToSlideRel(
    rootArchive: IArchive,
    relId: string,
    slideCount: number,
  ): Promise<XmlElement> {
    const targetType = this.host.targetType;
    return XmlHelper.append({
      archive: rootArchive,
      file: PptPaths.presentationRels,
      parent: (xml: XmlDocument) =>
        xml.getElementsByTagName('Relationships')[0],
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
        if (xml.getElementsByTagName('p:sldIdLst').length === 0) {
          XmlHelper.insertAfter(
            xml.createElement('p:sldIdLst'),
            xml.getElementsByTagName('p:sldMasterIdLst')[0],
          );
        }
      },
      parent: (xml: XmlDocument) => xml.getElementsByTagName('p:sldIdLst')[0],
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
        xml.getElementsByTagName('p:sldMasterIdLst')[0],
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
