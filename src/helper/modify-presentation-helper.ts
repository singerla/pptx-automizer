import { XmlHelper } from './xml-helper';
import { vd } from './general-helper';
import { contentTracker } from './content-tracker';
import { FileHelper } from './file-helper';
import JSZip from 'jszip';

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

  static async removeUnusedFiles(
    xml: XMLDocument,
    i: number,
    archive: JSZip,
  ): Promise<void> {
    for (const dir in contentTracker.files) {
      const requiredFiles = contentTracker.files[dir];

      archive.folder(dir).forEach((relativePath) => {
        if (
          !relativePath.includes('/') &&
          !requiredFiles.includes(relativePath)
        ) {
          FileHelper.removeFromArchive(archive, dir + '/' + relativePath);
        }
      });
    }
  }

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
}
