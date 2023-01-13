import JSZip from 'jszip';

import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import {
  AnalyzedElementType,
  ImportedElement,
  ImportElement,
  SlideModificationCallback,
  ShapeModificationCallback,
} from '../types/types';
import { ISlide } from '../interfaces/islide';
import { IPresentationProps } from '../interfaces/ipresentation-props';
import { PresTemplate } from '../interfaces/pres-template';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { ElementType } from '../enums/element-type';
import {
  RelationshipAttribute,
  SlideListAttribute,
  HelperElement,
} from '../types/xml-types';
import { Image } from '../shapes/image';
import { Chart } from '../shapes/chart';
import { GenericShape } from '../shapes/generic';
import { GeneralHelper, vd } from '../helper/general-helper';
import { FileProxy } from '../helper/file-proxy';

export class Master {
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
  targetArchive: FileProxy;
  /**
   * Source archive of slide
   * @internal
   */
  sourceArchive: FileProxy;
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
   * Modifications  of slide
   * @internal
   */
  modifications: SlideModificationCallback[];
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
   * Root template of slide
   * @internal
   */
  rootTemplate: RootPresTemplate;
  /**
   * Root  of slide
   * @internal
   */
  root: IPresentationProps;
  /**
   * Target rels path of slide
   * @internal
   */
  targetRelsPath: string;

  constructor(params: {
    presentation: IPresentationProps;
    template: PresTemplate;
    masterNumber: number;
  }) {
    this.sourceTemplate = params.template;
    this.sourceNumber = params.masterNumber;

    this.sourcePath = `ppt/slides/slide${this.sourceNumber}.xml`;
    this.relsPath = `ppt/slides/_rels/slide${this.sourceNumber}.xml.rels`;
  }

  /**
   * Appends slide
   * @internal
   * @param targetTemplate
   * @returns append
   */
  async append(targetTemplate: RootPresTemplate): Promise<void> {
    this.targetTemplate = targetTemplate;

    this.targetArchive = await targetTemplate.archive;
    this.targetNumber = targetTemplate.incrementCounter('slides');
    this.targetPath = `ppt/slides/slide${this.targetNumber}.xml`;
    this.targetRelsPath = `ppt/slides/_rels/slide${this.targetNumber}.xml.rels`;
    this.sourceArchive = await this.sourceTemplate.archive;

    console.log('Appending slide ' + this.targetNumber);
  }
}
