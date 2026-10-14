import { AutomizerError } from '../errors';
import { CountHelper } from '../helper/count-helper';
import { log } from '../helper/logger';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlDocument, XmlElement } from '../types/xml-types';

const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/**
 * Directories whose `_rels` parts can reference a theme. A theme is only
 * removed with its master when none of these still points at it (the
 * presentation itself usually references the first theme).
 */
const THEME_REFERRER_DIRS = [
  'ppt/_rels',
  'ppt/slideMasters/_rels',
  'ppt/notesMasters/_rels',
  'ppt/handoutMasters/_rels',
];

/**
 * Removes a slide master from the output presentation together with its
 * layouts, its theme (if no other part uses it), its relationships and
 * its entries in `ppt/presentation.xml` and `[Content_Types].xml`
 * (`Automizer.removeMaster`, #85).
 *
 * Runs after all slides have been written: a master whose layout is still
 * used by a listed slide cannot be removed — point the slide at another
 * layout with `useSlideLayout()` first.
 *
 * The removal steps follow the manual recipe that produced a valid file in
 * the #85 discussion: delete the theme, the master and its rels, the
 * layouts and their rels, the presentation's relationships to them and the
 * `p:sldMasterId` entry.
 */
export class SlideMasterRemover {
  constructor(private root: RootPresTemplate) {}

  get archive(): IArchive {
    return this.root.archive;
  }

  async remove(masterNumber: number): Promise<void> {
    const archive = this.archive;
    const masterPath = PptPaths.slideMaster(masterNumber);
    const masterRelsPath = PptPaths.slideMasterRels(masterNumber);

    // ensure the archive is open before fileExists() is consulted
    await XmlHelper.getXmlFromArchive(archive, PptPaths.presentation);

    if (!archive.fileExists(masterPath)) {
      throw new AutomizerError(
        `Cannot remove slide master ${masterNumber}: ${masterPath} does not ` +
          `exist. Masters are numbered 1-based by file name, as in the root ` +
          `template; imported masters follow the existing ones.`,
      );
    }

    const layoutNumbers = (
      await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
        archive,
        masterRelsPath,
        '../slideLayouts/slideLayout',
      )
    ).map((target) => target.number);

    const themeTargets =
      await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
        archive,
        masterRelsPath,
        '../theme/theme',
      );
    const themeNumber = themeTargets[0]?.number;

    const slidesInUse = await this.findSlidesUsingLayouts(layoutNumbers);
    if (slidesInUse.length) {
      throw new AutomizerError(
        `Cannot remove slide master ${masterNumber}: ` +
          slidesInUse
            .map((use) => `slide ${use.slide} uses slideLayout${use.layout}`)
            .join(', ') +
          `. Point these slides at a layout of another master with ` +
          `useSlideLayout() first.`,
      );
    }

    log.info(
      `Removing slideMaster ${masterNumber} with ${layoutNumbers.length} layout(s)`,
    );

    await this.removeFromPresentation(masterNumber, layoutNumbers);

    const removedParts = [masterPath, masterRelsPath];
    for (const layoutNumber of layoutNumbers) {
      removedParts.push(
        PptPaths.slideLayout(layoutNumber),
        PptPaths.slideLayoutRels(layoutNumber),
      );
    }
    for (const part of removedParts) {
      await this.removePart(part);
    }

    let themeRemoved = false;
    if (themeNumber && !(await this.isThemeReferenced(themeNumber))) {
      await this.removePart(PptPaths.theme(themeNumber));
      themeRemoved = true;
    }

    await this.removeContentTypes(
      removedParts.concat(themeRemoved ? [PptPaths.theme(themeNumber)] : []),
    );

    CountHelper.decrement('masters', this.root.counter, 1);
    CountHelper.decrement('layouts', this.root.counter, layoutNumbers.length);
    if (themeRemoved) {
      CountHelper.decrement('themes', this.root.counter, 1);
    }
  }

  /**
   * Slides listed in `p:sldIdLst` whose layout relationship points at one
   * of the given layouts.
   */
  async findSlidesUsingLayouts(
    layoutNumbers: number[],
  ): Promise<{ slide: number; layout: number }[]> {
    const archive = this.archive;
    const presentation = await XmlHelper.getXmlFromArchive(
      archive,
      PptPaths.presentation,
    );
    const slideTargets =
      await XmlRelationshipHelper.getTargetsByRelationshipType(
        archive,
        PptPaths.presentationRels,
        REL_TYPE + 'slide',
      );

    const uses: { slide: number; layout: number }[] = [];
    const slideIds = presentation.getElementsByTagName('p:sldId');
    for (let i = 0; i < slideIds.length; i++) {
      const rId = slideIds.item(i).getAttribute('r:id');
      const slideTarget = slideTargets.find((target) => target.rId === rId);
      if (!slideTarget) {
        continue;
      }
      const slideRelsPath = PptPaths.slideRels(slideTarget.number);
      if (!archive.fileExists(slideRelsPath)) {
        continue;
      }
      const layouts =
        await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
          archive,
          slideRelsPath,
          '../slideLayouts/slideLayout',
        );
      for (const layout of layouts) {
        if (layoutNumbers.includes(layout.number)) {
          uses.push({ slide: slideTarget.number, layout: layout.number });
        }
      }
    }
    return uses;
  }

  /**
   * Drops the `p:sldMasterId` entry of the master and every relationship
   * of `ppt/presentation.xml` that points at the master or one of its
   * layouts (layouts may only be related to their master; a presentation
   * relationship to a layout is a leftover of master auto-import).
   */
  async removeFromPresentation(
    masterNumber: number,
    layoutNumbers: number[],
  ): Promise<void> {
    const archive = this.archive;
    const removedTargets = [
      `slideMasters/slideMaster${masterNumber}.xml`,
      ...layoutNumbers.map((n) => `slideLayouts/slideLayout${n}.xml`),
    ];

    const removedRels = await XmlHelper.removeIf({
      archive,
      file: PptPaths.presentationRels,
      tag: 'Relationship',
      clause: (xml: XmlDocument, element: XmlElement) =>
        removedTargets.includes(element.getAttribute('Target')),
    });
    const removedRelIds = removedRels.map((rel) => rel.getAttribute('Id'));

    await XmlHelper.removeIf({
      archive,
      file: PptPaths.presentation,
      tag: 'p:sldMasterId',
      clause: (xml: XmlDocument, element: XmlElement) =>
        removedRelIds.includes(element.getAttribute('r:id')),
    });
  }

  /**
   * `true` if any `_rels` part of the output still points at the theme.
   */
  async isThemeReferenced(themeNumber: number): Promise<boolean> {
    const archive = this.archive;
    const themeFile = `theme/theme${themeNumber}.xml`;

    for (const dir of THEME_REFERRER_DIRS) {
      const relsFiles = await archive.folder(dir);
      for (const relsFile of relsFiles) {
        if (!relsFile.name.endsWith('.rels')) {
          continue;
        }
        const targets =
          await XmlRelationshipHelper.getTargetsByRelationshipType(
            archive,
            relsFile.name,
            REL_TYPE + 'theme',
          );
        if (targets.some((target) => target.file.endsWith(themeFile))) {
          return true;
        }
      }
    }
    return false;
  }

  async removeContentTypes(parts: string[]): Promise<void> {
    const partNames = parts.map((part) => PptPaths.partName(part));
    await XmlHelper.removeIf({
      archive: this.archive,
      file: PptPaths.contentTypes,
      tag: 'Override',
      clause: (xml: XmlDocument, element: XmlElement) =>
        partNames.includes(element.getAttribute('PartName')),
    });
  }

  async removePart(part: string): Promise<void> {
    if (this.archive.fileExists(part)) {
      await this.archive.remove(part);
    }
  }
}
