import IArchive from '../interfaces/iarchive';

import { ICounter } from '../interfaces/icounter';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { XmlHelper } from './xml-helper';

export class CountHelper implements ICounter {
  template: RootPresTemplate;
  name: string;
  count: number;

  /**
   * Distance between `count` and the number of the next part to create,
   * see `getHighestPartNumber()`.
   */
  private numberOffset = 0;

  constructor(name: string, template: RootPresTemplate) {
    this.name = name;
    this.template = template;
  }

  static increment(name: string, counters: ICounter[]): number | null {
    return CountHelper.getCounterByName(name, counters)._increment();
  }

  /**
   * Lower a counter, e.g. after slides have been removed from the archive.
   */
  static decrement(name: string, counters: ICounter[], by = 1): number {
    const counter = CountHelper.getCounterByName(name, counters);
    counter.count = Math.max(0, counter.count - by);
    return counter.count;
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

  /**
   * Increments the counter and returns the number for the next part.
   */
  _increment(): number {
    this.count++;
    return this.count + this.numberOffset;
  }

  async set(): Promise<void> {
    this.count = await this.calculateCount(this.template.archive);

    const highestNumber = await this.getHighestPartNumber(
      this.template.archive,
    );
    this.numberOffset = Math.max(0, highestNumber - this.count);
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
      case 'diagrams':
        return CountHelper.countDiagrams(presentation);
      case 'oleObjects':
        return CountHelper.countOleObjects(presentation);
    }

    throw new Error(`No way to count ${this.name}.`);
  }

  /**
   * Highest number in the names of the parts this counter is used to name.
   *
   * New parts are numbered after the counter, but the number of existing
   * parts is not necessarily the highest number in use: the numbering can
   * have gaps (e.g. in an output of automizer used as root template), or
   * a part is not counted at all (`media<n>` files, theme overrides of
   * charts). A new part would then overwrite an existing part that is
   * still in use. Numbering continues above the highest number instead.
   */
  private getHighestPartNumber(presentation: IArchive): Promise<number> {
    switch (this.name) {
      case 'charts':
        // One number names a chart and all parts copied along with it,
        // see Chart.copyFiles() and Chart.copyChartStyleFiles().
        return CountHelper.getHighestNumber(presentation, [
          ['ppt/charts', /^(?:chart|chartEx|colors|style)(\d+)\.xml$/i],
          ['ppt/drawings', /^drawing(\d+)\.xml$/i],
          ['ppt/theme', /^themeOverride(\d+)\.xml$/i],
          ['ppt/embeddings', /-created-(\d+)\.[^.]+$/i],
          ['ppt/media', /-chart-(\d+)\.[^.]+$/i],
        ]);
      case 'images':
        return CountHelper.getHighestNumber(presentation, [
          ['ppt/media', /^(?:image|media)(\d+)\.[^.]+$/i],
        ]);
      case 'diagrams':
        return CountHelper.getHighestNumber(presentation, [
          [
            'ppt/diagrams',
            /^(?:data|colors|layout|quickStyle|drawing)(\d+)\.xml$/i,
          ],
        ]);
      case 'oleObjects':
        return CountHelper.getHighestNumber(presentation, [
          ['ppt/embeddings', /^oleObject(\d+)\.[^.]+$/i],
        ]);
    }

    return Promise.resolve(0);
  }

  /**
   * Highest number captured by one of the file name patterns in its
   * folder, 0 if no file matches.
   */
  private static async getHighestNumber(
    presentation: IArchive,
    patterns: [string, RegExp][],
  ): Promise<number> {
    let highest = 0;
    for (const [folder, pattern] of patterns) {
      const files = await presentation.folder(folder);
      for (const file of files) {
        const match = file.relativePath.match(pattern);
        if (match) {
          highest = Math.max(highest, Number(match[1]));
        }
      }
    }
    return highest;
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

  /**
   * Counts the Override entries of [Content_Types].xml with the given
   * ContentType.
   */
  private static async countContentTypeOverrides(
    presentation: IArchive,
    contentType: string,
  ): Promise<number> {
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      presentation,
      '[Content_Types].xml',
    );
    const overrides = contentTypesXml.getElementsByTagName('Override');

    let count = 0;
    for (let i = 0; i < overrides.length; i++) {
      if (overrides.item(i)?.getAttribute('ContentType') === contentType) {
        count++;
      }
    }
    return count;
  }

  private static async countLayouts(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    );
  }

  private static async countThemes(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.theme+xml',
    );
  }

  private static async countCharts(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    );
  }

  private static async countOleObjects(
    presentation: IArchive,
  ): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.oleObject',
    );
  }

  private static async countImages(presentation: IArchive): Promise<number> {
    const mediaFiles = await presentation.folder('ppt/media');
    const count = mediaFiles.filter(
      (file) => file.relativePath.indexOf('image') === 0,
    ).length;
    return count;
  }

  private static async countDiagrams(presentation: IArchive): Promise<number> {
    return CountHelper.countContentTypeOverrides(
      presentation,
      'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
    );
  }
}
