import IArchive from '../interfaces/iarchive';

import { ICounter } from '../interfaces/icounter';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlHelper } from './xml-helper';
import { XmlElement } from '../types/xml-types';
import { FileHelper } from './file-helper';

export class CountHelper implements ICounter {
  template: RootPresTemplate;
  name: string;
  count: number;

  constructor(name: string, template: RootPresTemplate) {
    this.name = name;
    this.template = template;
  }

  static increment(name: string, counters: ICounter[]): number | null {
    return CountHelper.getCounterByName(name, counters)._increment();
  }

  static count(name: string, counters: ICounter[]): number {
    return CountHelper.getCounterByName(name, counters).get();
  }

  static reset(counters: ICounter[]): void {
    counters.forEach((counter) => (counter.count = undefined));
  }

  static getCounterByName(name: string, counters: ICounter[]): ICounter {
    const counter = counters.find((c) => c.name === name);
    if (counter === undefined) {
      throw new Error(`Counter ${name} not found.`);
    }
    return counter;
  }

  _increment(): number {
    this.count++;
    return this.count;
  }

  async set(): Promise<void> {
    this.count = await this.calculateCount(await this.template.archive);
  }

  get(): number {
    return this.count;
  }

  private calculateCount(presentation: IArchive): Promise<number> {
    switch (this.name) {
      case 'slides':
        return CountHelper.countSlides(presentation);
      case 'masters':
        return CountHelper.countMasters(presentation);
      case 'layouts':
        return CountHelper.countLayouts(presentation);
      case 'themes':
        return CountHelper.countThemes(presentation);
      case 'charts':
        return CountHelper.countCharts(presentation);
      case 'images':
        return CountHelper.countImages(presentation);
        case 'oleObjects':  
        return CountHelper.countOleObjects(presentation);  
  
    }

    throw new Error(`No way to count ${this.name}.`);
  }

  private static async countSlides(presentation: IArchive): Promise<number> {
    const presentationXml = await XmlHelper.getXmlFromArchive(
      presentation,
      'ppt/presentation.xml',
    );
    return presentationXml.getElementsByTagName('p:sldId').length;
  }

  private static async countMasters(presentation: IArchive): Promise<number> {
    const presentationXml = await XmlHelper.getXmlFromArchive(
      presentation,
      'ppt/presentation.xml',
    );
    return presentationXml.getElementsByTagName('p:sldMasterId').length;
  }

  private static async countLayouts(presentation: IArchive): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');

    return Object.keys(overrides)
      .map((key) => overrides[key] as XmlElement)
      .filter(
        (o) =>
          o.getAttribute &&
          o.getAttribute('ContentType') ===
            `application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml`,
      ).length;
  }

  private static async countThemes(presentation: IArchive): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');

    return Object.keys(overrides)
      .map((key) => overrides[key] as XmlElement)
      .filter(
        (o) =>
          o.getAttribute &&
          o.getAttribute('ContentType') ===
            `application/vnd.openxmlformats-officedocument.theme+xml`,
      ).length;
  }

  private static async countCharts(presentation: IArchive): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');

    return Object.keys(overrides)
      .map((key) => overrides[key] as XmlElement)
      .filter(
        (o) =>
          o.getAttribute &&
          o.getAttribute('ContentType') ===
            `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`,
      ).length;
  }

  private static async countOleObjects(presentation: IArchive): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');
  
    return Object.keys(overrides)
      .map((key) => overrides[key] as XmlElement)
      .filter(
        (o) =>
          o.getAttribute &&
          o.getAttribute('ContentType') ===
            `application/vnd.openxmlformats-officedocument.oleObject`
      ).length;
  }
  

  private static async countImages(presentation: IArchive): Promise<number> {
    const mediaFiles = await presentation.folder('ppt/media');
    const count = mediaFiles.filter(
      (file) => file.relativePath.indexOf('image') === 0,
    ).length;
    return count;
  }
}
