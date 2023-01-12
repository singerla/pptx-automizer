import { XmlHelper } from './xml-helper';
import { contentTracker as Tracker } from './content-tracker';
import { FileHelper } from './file-helper';
import JSZip from 'jszip';
import { vd } from './general-helper';

export default class ModifyPresentationHelper {
  /**
   * Get Collection of slides
   */
  static getSlidesCollection = (xml: XMLDocument) => {
    return xml.getElementsByTagName('p:sldId');
  };

  /**
   * Pass an array of slide numbers to define a target sort order.
   * First slide starts by 1.
   * @order Array of slide numbers, starting by 1
   */
  static sortSlides = (order: number[]) => (xml: XMLDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    order.map((index, i) => order[i]--);
    XmlHelper.sortCollection(slides, order);
  };

  /**
   * Set ids to prevent corrupted pptx.
   * Must start with 256 and increment by one.
   */
  static normalizeSlideIds = (xml: XMLDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    const firstId = 256;
    XmlHelper.modifyCollection(slides, (slide: Element, i) => {
      slide.setAttribute('id', String(firstId + i));
    });
  };

  /**
   * contentTracker.files includes all files that have been
   * copied into the root template by automizer. We remove all other files.
   */
  static async removeUnusedFiles(
    xml: XMLDocument,
    i: number,
    archive: JSZip,
  ): Promise<void> {
    const skipDirs = ['ppt/slideMasters', 'ppt/slideMasters/_rels'];
    for (const dir in Tracker.files) {
      if (skipDirs.includes(dir)) {
        continue;
      }
      const requiredFiles = Tracker.files[dir];
      archive.folder(dir).forEach((relativePath, file) => {
        if (
          !relativePath.includes('/') &&
          !requiredFiles.includes(relativePath)
        ) {
          FileHelper.removeFromArchive(archive, file.name);
        }
      });
    }
  }

  /**
   * PPT won't complain about unused items in [Content_Types].xml,
   * but we remove them anyway in case the file mentioned in PartName-
   * attribute does not exist.
   */
  static async removeUnusedContentTypes(
    xml: XMLDocument,
    i: number,
    archive: JSZip,
  ): Promise<void> {
    await XmlHelper.removeIf({
      archive,
      file: `[Content_Types].xml`,
      tag: 'Override',
      clause: (xml: XMLDocument, element: Element) => {
        const filename = element.getAttribute('PartName').substring(1);
        return FileHelper.fileExistsInArchive(archive, filename) ? false : true;
      },
    });
  }

  static async removedUnusedImages(
    xml: XMLDocument,
    i: number,
    archive: JSZip,
  ): Promise<void> {
    await Tracker.analyzeContents(archive);
    const extensions = ['jpg', 'jpeg', 'png', 'gif', 'svg', 'emf'];
    const keepFiles = [];
    const addFiles = async (section: string) => {
      const imagesInfo =
        Tracker.getRelationTag(section).getRelationTargets('image');
      const images = await Tracker.getRelatedContents(imagesInfo);
      images.forEach((image) => keepFiles.push(image.filename));
    };

    await addFiles('ppt/slides');
    await addFiles('ppt/slideMasters');

    archive.folder('ppt/media').forEach((relativePath, file) => {
      if (!relativePath.includes('/')) {
        const info = FileHelper.getFileInfo(file.name);
        if (
          extensions.includes(info.extension.toLowerCase()) &&
          !keepFiles.includes(info.base)
        ) {
          archive.remove(file.name);
          FileHelper.removeFromArchive(archive, file.name);
        }
      }
    });
  }
}
