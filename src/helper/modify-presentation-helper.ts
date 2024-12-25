import { XmlHelper } from './xml-helper';
import { contentTracker as Tracker } from './content-tracker';
import { FileHelper } from './file-helper';
import IArchive from '../interfaces/iarchive';
import { XmlDocument, XmlElement } from '../types/xml-types';
import { log, vd } from './general-helper';
import { XmlRelationshipHelper } from './xml-relationship-helper';
import { Target } from '../types/types';
import Automizer from '../automizer';

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
   * @order Array of slide numbers, starting by 1
   */
  static sortSlides = (order: number[]) => (xml: XmlDocument) => {
    const slides = ModifyPresentationHelper.getSlidesCollection(xml);
    order.map((index, i) => order[i]--);
    XmlHelper.sortCollection(slides, order);
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
   * Update slideMaster ids to prevent corrupted pptx.
   * - Take first slideMaster id from presentation.xml to start,
   * - then update incremental ids of each p:sldLayoutId in slideMaster[i].xml
   *   (starting by slideMasterId + 1)
   * - and update next slideMaster id with previous p:sldLayoutId + 1
   *
   * p:sldMasterId-ids and p:sldLayoutId-ids need to be in a row, otherwise
   * PowerPoint will complain on any p:sldLayoutId-id lower than its
   * corresponding slideMaster-id. omg.
   */
  static normalizeSlideMasterIds =
    (currentId: number) =>
    async (
      presXml: XmlDocument,
      i: number,
      archive: IArchive,
      pres: Automizer,
    ) => {
      const slides =
        ModifyPresentationHelper.getSlideMastersCollection(presXml);

      const deletedIds = pres.content.deleted['ppt/slideMasters'].map(
        (deleted: any) => deleted.targetMasterId,
      );

      await XmlHelper.modifyCollectionAsync(
        slides,
        async (slide: XmlElement, i) => {
          const masterId = i + 1;
          if (deletedIds.includes(masterId)) {
            return;
          }

          const slideMasterXml = await XmlHelper.getXmlFromArchive(
            archive,
            `ppt/slideMasters/slideMaster${masterId}.xml`,
          );

          slide.setAttribute('id', String(currentId));
          currentId++;

          const slideLayouts =
            slideMasterXml.getElementsByTagName('p:sldLayoutId');

          XmlHelper.modifyCollection(
            slideLayouts,
            (slideLayout: XmlElement) => {
              slideLayout.setAttribute('id', String(currentId));
              currentId++;
            },
          );
        },
      );

      deletedIds.forEach((deletedId) => {
        const existingMasters = presXml.getElementsByTagName('p:sldMasterId');
        XmlHelper.sliceCollection(existingMasters, 1, deletedId - 1);
      });

      XmlHelper.dump(slides.item(0));
    };

  static getFirstSlideMasterId = async (pres: Automizer) => {
    const presXml = await XmlHelper.getXmlFromArchive(
      pres.rootTemplate.archive,
      'ppt/presentation.xml',
    );
    const slides = ModifyPresentationHelper.getSlideMastersCollection(presXml);
    const first = slides.item(0);
    return Number(first.getAttribute('id'));
  };

  /**
   * Tracker.files includes all files that have been
   * copied to the root template by automizer. We remove all other files.
   */
  static async removeUnusedFiles(
    xml: XmlDocument,
    i: number,
    archive: IArchive,
  ): Promise<void> {
    // Need to skip some dirs until masters and layouts are handled properly
    const skipDirs = [
      'ppt/slideMasters',
      'ppt/slideMasters/_rels',
      'ppt/slideLayouts',
      'ppt/slideLayouts/_rels',
    ];
    for (const dir in Tracker.files) {
      if (skipDirs.includes(dir)) {
        continue;
      }
      const requiredFiles = Tracker.files[dir];
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
    await Tracker.analyzeContents(archive);

    const extensions = ['jpg', 'jpeg', 'png', 'gif', 'svg', 'emf'];
    const keepFiles = [];

    await Tracker.collect('ppt/slides', 'image', keepFiles);
    await Tracker.collect('ppt/slideMasters', 'image', keepFiles);
    await Tracker.collect('ppt/slideLayouts', 'image', keepFiles);

    await FileHelper.removeFromDirectory(archive, 'ppt/media', (file) => {
      const info = FileHelper.getFileInfo(file.name);
      return (
        extensions.includes(info.extension.toLowerCase()) &&
        !keepFiles.includes(info.base)
      );
    });
  }

  static removeSlideMaster =
    (length: number, from: number, pres: Automizer) =>
    async (presXml: XmlDocument) => {
      for (let i = 0; i < length; i += 1) {
        const targetMasterId = from + i + 1;
        const masterToRemove = `slideMaster${targetMasterId}.xml`;

        const layouts = (await new XmlRelationshipHelper().initialize(
          pres.rootTemplate.archive,
          `${masterToRemove}.rels`,
          `ppt/slideMasters/_rels`,
          '../slideLayouts/slideLayout',
        )) as Target[];

        const layoutFiles = layouts.map(
          (f) => f.file, // path.resolve(`ppt/presentation.xml/${f.file}`),
        );

        const removedLayouts = await FileHelper.removeFromDirectory(
          pres.rootTemplate.archive,
          'ppt/slideLayouts/',
          (file) => {
            return !layoutFiles.includes(file.relativePath);
          },
        );

        const themes = (await new XmlRelationshipHelper().initialize(
          pres.rootTemplate.archive,
          `${masterToRemove}.rels`,
          `ppt/slideMasters/_rels`,
          '../theme/theme',
        )) as Target[];

        const themesFiles = themes.map(
          (f) => f.file, // path.resolve(`ppt/presentation.xml/${f.file}`),
        );

        // const removedThemes = await FileHelper.removeFromDirectory(
        //   pres.rootTemplate.archive,
        //   'ppt/theme/',
        //   (file) => {
        //     return !themesFiles.includes(file.relativePath);
        //   },
        // );

        const removedMasters = await FileHelper.removeFromDirectory(
          pres.rootTemplate.archive,
          'ppt/slideMasters/',
          (file) => {
            return file.relativePath === masterToRemove;
          },
        );

        const removedMasterRels = await FileHelper.removeFromDirectory(
          pres.rootTemplate.archive,
          'ppt/slideMasters/_rels',
          (file) => {
            return file.relativePath === `${masterToRemove}.rels`;
          },
        );

        log('removed Layouts:' + removedLayouts.length, 2);
        // log('removed Themes:' + removedThemes.length, 2);
        log('removed Masters:' + removedMasters.length, 2);
        log('removed MasterRels:' + removedMasterRels.length, 2);

        pres.content.deletedFile('ppt/slideMasters', {
          masterToRemove,
          targetMasterId: targetMasterId,
        });
      }
    };
}
