import { XmlHelper } from './xml-helper';
import { XmlRelationshipHelper } from './xml-relationship-helper';
import { PptPaths } from './ppt-paths';
import { FileHelper } from './file-helper';
import IArchive from '../interfaces/iarchive';
import { XmlDocument, XmlElement } from '../types/xml-types';

export default class ModifyPresentationHelper {
  /**
   * Get Collection of slides
   */
  static getSlidesCollection = (xml: XmlDocument) => {
    return xml.getElementsByTagName('p:sldId');
  };
  static getSlideMastersCollection = (xml: XmlDocument) => {
    return xml.getElementsByTagName('p:sldMasterId');
  };

  /**
   * Pass an array of slide numbers to define a target sort order.
   * First slide starts by 1.
   * @param order Array of slide numbers, starting by 1
   */
  static sortSlides = (order: number[]) => (xml: XmlDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    order.map((index, i) => order[i]--);
    XmlHelper.sortCollection(slides, order);
  };

  /**
   * Pass an array of slide numbers to remove from slide sortation.
   * @param numbers Array of slide numbers, starting by 1
   */
  static removeSlides = (numbers: number[]) => (xml: XmlDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    numbers.map((index, i) => numbers[i]--);
    for (let i = 0; i <= slides.length; i++) {
      if (numbers.includes(i)) {
        XmlHelper.remove(slides[i]);
      }
    }
  };

  /**
   * Set ids to prevent corrupted pptx.
   * Must start with 256 and increment by one.
   */
  static normalizeSlideIds = (xml: XmlDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    const firstId = 256;
    XmlHelper.modifyCollection(slides, (slide: XmlElement, i) => {
      slide.setAttribute('id', String(firstId + i));
    });
  };

  /**
   * Smallest id PowerPoint accepts for `p:sldMasterId` / `p:sldLayoutId`
   * (ids below 2^31 trigger the repair prompt; see #85).
   */
  static readonly MIN_SLIDE_MASTER_ID = 2147483648;

  /**
   * Update slideMaster ids to prevent corrupted pptx.
   * - Take first slideMaster id from presentation.xml to start (raised to
   *   `MIN_SLIDE_MASTER_ID` if it is lower),
   * - then update incremental ids of each p:sldLayoutId in the master part
   *   the entry's relationship points at (starting by slideMasterId + 1)
   * - and update next slideMaster id with previous p:sldLayoutId + 1
   *
   * p:sldMasterId-ids and p:sldLayoutId-ids need to be in a row, otherwise
   * PowerPoint will complain on any p:sldLayoutId-id lower than its
   * corresponding slideMaster-id. omg.
   */
  static normalizeSlideMasterIds = async (
    xml: XmlDocument,
    i: number,
    archive: IArchive,
  ) => {
    const slides = ModifyPresentationHelper.getSlideMastersCollection(xml);
    // The list entries and the master part numbers do not line up once a
    // master has been removed: resolve each entry through its relationship.
    const masterTargets =
      await XmlRelationshipHelper.getTargetsByRelationshipType(
        archive,
        PptPaths.presentationRels,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
      );
    let currentId: number;
    await XmlHelper.modifyCollectionAsync(
      slides,
      async (slide: XmlElement, i) => {
        const masterTarget = masterTargets.find(
          (target) => target.rId === slide.getAttribute('r:id'),
        );
        const masterId = masterTarget ? masterTarget.number : i + 1;
        if (i === 0) {
          currentId = Math.max(
            Number(slide.getAttribute('id')) || 0,
            ModifyPresentationHelper.MIN_SLIDE_MASTER_ID,
          );
        }

        slide.setAttribute('id', String(currentId));
        currentId++;

        const slideMasterXml = await XmlHelper.getXmlFromArchive(
          archive,
          PptPaths.slideMaster(masterId),
        );

        const slideLayouts =
          slideMasterXml.getElementsByTagName('p:sldLayoutId');
        XmlHelper.modifyCollection(slideLayouts, (slideLayout: XmlElement) => {
          slideLayout.setAttribute('id', String(currentId));
          currentId++;
        });
      },
    );
  };

  /**
   * The content tracker's files include all files that have been
   * copied to the root template by automizer. We remove all other files.
   */
  static async removeUnusedFiles(
    xml: XmlDocument,
    i: number,
    archive: IArchive,
  ): Promise<void> {
    const tracker = archive.contentTracker;
    if (!tracker) {
      return;
    }
    // Need to skip some dirs until masters and layouts are handled properly.
    // ppt/drawings is tracked for image collection only (see
    // removedUnusedImages); it also holds vml drawings that may belong to
    // masters and layouts, so it is left alone here.
    const skipDirs = [
      'ppt/slideMasters',
      'ppt/slideMasters/_rels',
      'ppt/slideLayouts',
      'ppt/slideLayouts/_rels',
      'ppt/drawings',
    ];
    for (const dir in tracker.files) {
      if (skipDirs.includes(dir)) {
        continue;
      }
      const requiredFiles = tracker.files[dir];
      await FileHelper.removeFromDirectory(archive, dir, (file) => {
        return !requiredFiles.includes(file.relativePath);
      });
    }
  }

  /**
   * PPT won't complain about unused items in [Content_Types].xml,
   * but we remove them anyway in case the file mentioned in PartName-
   * attribute does not exist.
   */
  static async removeUnusedContentTypes(
    xml: XmlDocument,
    i: number,
    archive: IArchive,
  ): Promise<void> {
    await XmlHelper.removeIf({
      archive,
      file: `[Content_Types].xml`,
      tag: 'Override',
      clause: (xml: XmlDocument, element: XmlElement) => {
        const filename = element.getAttribute('PartName').substring(1);
        const exists = FileHelper.fileExistsInArchive(archive, filename);
        return exists ? false : true;
      },
    });
  }

  static async removedUnusedImages(
    xml: XmlDocument,
    i: number,
    archive: IArchive,
  ): Promise<void> {
    const tracker = archive.contentTracker;
    if (!tracker) {
      return;
    }
    await tracker.analyzeContents(archive);

    const extensions = ['jpg', 'jpeg', 'png', 'gif', 'svg', 'emf'];
    const keepFiles: string[] = [];

    await tracker.collect('ppt/slides', 'image', keepFiles);
    await tracker.collect('ppt/slideMasters', 'image', keepFiles);
    await tracker.collect('ppt/slideLayouts', 'image', keepFiles);
    // Images referenced by copied charts (e.g. scatter point markers) and by
    // their user shapes drawings (e.g. a picture in a chart annotation).
    await tracker.collect('ppt/charts', 'image', keepFiles);
    await tracker.collect('ppt/drawings', 'image', keepFiles);

    await FileHelper.removeFromDirectory(archive, 'ppt/media', (file) => {
      const info = FileHelper.getFileInfo(file.name);
      return (
        extensions.includes(info.extension.toLowerCase()) &&
        !keepFiles.includes(info.base)
      );
    });
  }
}
