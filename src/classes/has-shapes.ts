import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import IArchive from '../interfaces/iarchive';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import {
  AutomizerParams,
  ImportElement,
  ShapeTargetType,
  SlideModificationCallback,
  StatusTracker,
} from '../types/types';
import { ContentTracker } from '../helper/content-tracker';
import { vd } from '../helper/general-helper';

export default class HasShapes {
  /**
   * Source template of slide
   * @internal
   */
  sourceTemplate: PresTemplate;
  /**
   * Target template of slide
   * @internal
   */
  targetTemplate: RootPresTemplate;
  /**
   * Target number of slide
   * @internal
   */
  targetNumber: number;
  /**
   * Source number of slide
   * @internal
   */
  sourceNumber: number;
  /**
   * Target archive of slide
   * @internal
   */
  targetArchive: IArchive;
  /**
   * Source archive of slide
   * @internal
   */
  sourceArchive: IArchive;
  /**
   * Source path of slide
   * @internal
   */
  sourcePath: string;
  /**
   * Target path of slide
   * @internal
   */
  targetPath: string;
  /**
   * Root template of slide
   * @internal
   */
  modifications: SlideModificationCallback[];
  /**
   * Modifications of slide relations
   * @internal
   */
  relModifications: SlideModificationCallback[];
  /**
   * Import elements of slide
   * @internal
   */
  importElements: ImportElement[];
  /**
   * Rels path of slide
   * @internal
   */
  relsPath: string;
  /**
   * Target rels path of slide
   * @internal
   */
  targetRelsPath: string;
  /**
   * Root  of slide
   * @internal
   */
  root: IPresentationProps;
  status: StatusTracker;
  content: ContentTracker;
  /**
   * List of unsupported tags in slide xml
   * @internal
   */
  unsupportedTags = [
    'p:custDataLst',
    // 'mc:AlternateContent',
    //'a14:imgProps',
  ];
  targetType: ShapeTargetType;
  params: AutomizerParams;

  constructor() {}

  async checkIntegrity(): Promise<void> {
    const params = this.targetTemplate.automizer.params;

    const info = params.showIntegrityInfo;
    const assert = params.assertRelatedContents;

    if (info || assert) {
      const masterRels = await new XmlRelationshipHelper().initialize(
        this.targetArchive,
        `${this.targetType}${this.targetNumber}.xml.rels`,
        `ppt/${this.targetType}s/_rels`,
      );
      await masterRels.assertRelatedContent(this.sourceArchive, info, assert);
    }
  }
}
