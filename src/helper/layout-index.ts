import { SlideLayoutNotFoundError } from '../errors';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { SlideLayoutSelector } from '../types/types';
import { ROOT_TEMPLATE_ALIAS } from '../constants/constants';
import { PptPaths } from './ppt-paths';
import { XmlHelper } from './xml-helper';
import { XmlRelationshipHelper } from './xml-relationship-helper';


const REL_TYPE_SLIDE_MASTER =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster';

/**
 * One layout of the output presentation, as resolved by `LayoutIndex`.
 */
export type SlideLayoutInfo = {
  /** 1-based number of `ppt/slideLayouts/slideLayout<n>.xml` in the output */
  layoutNumber: number;
  /** `p:cSld@name` of the layout — the name PowerPoint's layout gallery shows */
  layoutName: string;
  /** 1-based number of `ppt/slideMasters/slideMaster<n>.xml` in the output */
  masterNumber: number;
  /**
   * Name of the master as PowerPoint shows it: the master's own
   * `p:cSld@name` if set, otherwise the name of its theme (`a:theme@name`).
   */
  masterName: string;
  /**
   * `'root'` for a layout that came with the root template, otherwise the
   * alias of the template it was imported from.
   */
  template: string;
};

type CachedIndex = { key: string; layouts: SlideLayoutInfo[] };

/**
 * Per-output cache; a change in the master or layout count invalidates it.
 * Keyed by the root template instance, so concurrent Automizer instances
 * never share state.
 */
const indexByRoot = new WeakMap<RootPresTemplate, CachedIndex>();

/**
 * Name index over the masters and layouts of the output presentation
 * (root template as loaded plus everything imported so far), built from
 * `ppt/presentation.xml` → master rels → layouts at write time. Backs
 * `slide.useSlideLayout({ layout, master?, template? })` (#136).
 */
export class LayoutIndex {
  /**
   * The layouts currently in the output, in master and layout order.
   */
  static async read(root: RootPresTemplate): Promise<SlideLayoutInfo[]> {
    const key = `${root.count('masters')}:${root.count('layouts')}`;
    const cached = indexByRoot.get(root);
    if (cached?.key === key) {
      return cached.layouts;
    }
    const layouts = await LayoutIndex.build(root);
    indexByRoot.set(root, { key, layouts });
    return layouts;
  }

  static async build(root: RootPresTemplate): Promise<SlideLayoutInfo[]> {
    const archive = root.archive;
    const layouts: SlideLayoutInfo[] = [];

    const presentation = await XmlHelper.getXmlFromArchive(
      archive,
      PptPaths.presentation,
    );
    const masterTargets =
      await XmlRelationshipHelper.getTargetsByRelationshipType(
        archive,
        PptPaths.presentationRels,
        REL_TYPE_SLIDE_MASTER,
      );

    const masterIds = presentation.getElementsByTagName('p:sldMasterId');
    for (let i = 0; i < masterIds.length; i++) {
      const rId = masterIds.item(i).getAttribute('r:id');
      const masterTarget = masterTargets.find((target) => target.rId === rId);
      if (!masterTarget) {
        continue;
      }
      const masterNumber = masterTarget.number;
      const masterName = await LayoutIndex.readMasterName(
        archive,
        masterNumber,
      );

      const layoutTargets =
        await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
          archive,
          PptPaths.slideMasterRels(masterNumber),
          '../slideLayouts/slideLayout',
        );

      for (const layoutTarget of layoutTargets) {
        const layoutNumber = layoutTarget.number;
        const layoutName = await LayoutIndex.readCSldName(
          archive,
          PptPaths.slideLayout(layoutNumber),
        );
        const imported = root.contentMap?.find(
          (map) => map.type === 'slideLayout' && map.targetId === layoutNumber,
        );

        layouts.push({
          layoutNumber,
          layoutName,
          masterNumber,
          masterName,
          template: imported ? imported.key : ROOT_TEMPLATE_ALIAS,
        });
      }
    }

    return layouts;
  }

  /**
   * The master's own `p:cSld@name` if it has one, else its theme's name.
   */
  static async readMasterName(
    archive: IArchive,
    masterNumber: number,
  ): Promise<string> {
    const ownName = await LayoutIndex.readCSldName(
      archive,
      PptPaths.slideMaster(masterNumber),
    );
    if (ownName) {
      return ownName;
    }

    const themeTargets =
      await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
        archive,
        PptPaths.slideMasterRels(masterNumber),
        '../theme/theme',
      );
    if (!themeTargets.length) {
      return '';
    }

    const theme = await XmlHelper.getXmlFromArchive(
      archive,
      PptPaths.theme(themeTargets[0].number),
    );
    return theme.getElementsByTagName('a:theme').item(0)?.getAttribute('name') ?? '';
  }

  static async readCSldName(archive: IArchive, path: string): Promise<string> {
    const xml = await XmlHelper.getXmlFromArchive(archive, path);
    return xml.getElementsByTagName('p:cSld').item(0)?.getAttribute('name') ?? '';
  }

  /**
   * All layouts matching the selector: equal layout name, and — if given —
   * equal master name and origin template.
   */
  static select(
    layouts: SlideLayoutInfo[],
    selector: SlideLayoutSelector,
  ): SlideLayoutInfo[] {
    return layouts.filter(
      (layout) =>
        layout.layoutName === selector.layout &&
        (selector.master === undefined ||
          layout.masterName === selector.master) &&
        (selector.template === undefined ||
          layout.template === selector.template),
    );
  }

  /**
   * Exactly one layout matching the selector; throws
   * `SlideLayoutNotFoundError` for no match or more than one.
   */
  static resolve(
    layouts: SlideLayoutInfo[],
    selector: SlideLayoutSelector,
  ): SlideLayoutInfo {
    const matches = LayoutIndex.select(layouts, selector);

    if (matches.length === 1) {
      return matches[0];
    }

    if (matches.length === 0) {
      throw new SlideLayoutNotFoundError(
        `Slide layout ${LayoutIndex.describeSelector(selector)} not found. ` +
          `Available layouts: ${LayoutIndex.describeLayouts(layouts)}`,
        { selector, candidates: [] },
      );
    }

    throw new SlideLayoutNotFoundError(
      `Slide layout ${LayoutIndex.describeSelector(selector)} is ambiguous: ` +
        `${LayoutIndex.describeLayouts(matches)}. ` +
        `Add "master" (or "template") to the selector.`,
      { selector, candidates: matches },
    );
  }

  static describeSelector(selector: SlideLayoutSelector): string {
    const parts = [`layout "${selector.layout}"`];
    if (selector.master !== undefined) {
      parts.push(`master "${selector.master}"`);
    }
    if (selector.template !== undefined) {
      parts.push(`template "${selector.template}"`);
    }
    return parts.join(', ');
  }

  static describeLayout(layout: SlideLayoutInfo): string {
    return (
      `"${layout.masterName} / ${layout.layoutName}" ` +
      `(slideLayout${layout.layoutNumber}, ${layout.template})`
    );
  }

  static describeLayouts(layouts: SlideLayoutInfo[]): string {
    if (!layouts.length) {
      return '(none)';
    }
    return layouts.map(LayoutIndex.describeLayout).join(', ');
  }
}
