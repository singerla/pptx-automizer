import { FileHelper } from '../helper/file-helper';
import { PptPaths } from '../helper/ppt-paths';
import { XmlHelper } from '../helper/xml-helper';
import { XmlRelationshipHelper } from '../helper/xml-relationship-helper';
import { Shape } from '../classes/shape';
import path from 'path';

import {
  RelationshipAttribute,
  XmlDocument,
  XmlElement,
} from '../types/xml-types';
import {
  ChartModificationCallback,
  ImportedElement,
  ShapeTargetType,
  Target,
  Workbook,
} from '../types/types';
import { IChart } from '../interfaces/ichart';
import { IShapeAction } from '../interfaces/ishape-action';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import IArchive from '../interfaces/iarchive';
import { ContentTypeExtension } from '../enums/content-type-map';
import { log } from '../helper/logger';

export class Chart extends Shape implements IChart, IShapeAction {
  sourceWorksheet: number | string;
  targetWorksheet: number | string;
  worksheetFilePrefix: string;
  wbEmbeddingsPath: string;
  wbExtension: string;
  relTypeChartColorStyle: string;
  relTypeChartUserShapes: string;
  relTypeChartStyle: string;
  relTypeChartImage: string;
  relTypeChartThemeOverride: string;
  wbRelsPath: string;
  styleRelationFiles: {
    [key: string]: string[];
  };
  callbacks: ChartModificationCallback[];
  // Will be false if the chart has no workbook (e.g. driven by external data)
  hasWorkbook: boolean;

  constructor(shape: ImportedElement, targetType: ShapeTargetType) {
    super(shape, targetType);

    this.relRootTag = this.subtype === 'chart' ? 'c:chart' : 'cx:chart';
    this.relAttribute = 'r:id';
    this.relParent =
      this.subtype === 'chart'
        ? (element: XmlElement) =>
            element.parentNode.parentNode.parentNode as XmlElement
        : (element: XmlElement) =>
            element.parentNode.parentNode.parentNode.parentNode
              .parentNode as XmlElement;

    this.wbEmbeddingsPath = `../embeddings/`;
    this.wbExtension = '.xlsx';
    this.relTypeChartColorStyle =
      'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
    this.relTypeChartStyle =
      'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
    this.relTypeChartImage =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
    this.relTypeChartThemeOverride =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride';
    this.relTypeChartUserShapes =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes';
    this.styleRelationFiles = {};
    this.hasWorkbook = true;
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Chart> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.clone();
    await this.replaceIntoSlideTree();

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Chart> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.clone();
    await this.appendToSlideTree();

    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Chart> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();

    return this;
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Chart> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateElementsRelId();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    this.targetNumber = this.targetTemplate.incrementCounter('charts');
    this.wbRelsPath = PptPaths.chartPartRels(this.subtype, this.sourceNumber);

    await this.copyFiles();
    await this.copyChartStyleFiles();
    await this.appendTypes();
    await this.appendToSlideRels();
  }

  async clone(): Promise<void> {
    await this.setTargetElement();
    await this.modifyChartData();
    await this.updateTargetElementRelId();
  }

  async modifyChartData(): Promise<void> {
    if (!this.hasWorkbook) {
      return;
    }

    const chartXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      PptPaths.chartPart(this.subtype, this.targetNumber),
    );

    const workbook = await this.readWorkbook();

    this.applyChartCallbacks(
      this.callbacks as ChartModificationCallback[],
      this.targetElement,
      chartXml,
      workbook,
    );

    XmlHelper.writeXmlToArchive(
      this.targetArchive,
      PptPaths.chartPart(this.subtype, this.targetNumber),
      chartXml,
    );
    await this.writeWorkbook(workbook);
  }

  async readWorkbook(): Promise<Workbook> {
    const workbookFilename = PptPaths.embedding(
      `${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
    );
    const archive = await this.targetArchive.extract(workbookFilename);

    const sheet = await XmlHelper.getXmlFromArchive(
      archive,
      'xl/worksheets/sheet1.xml',
    );

    const table = FileHelper.fileExistsInArchive(
      archive,
      'xl/tables/table1.xml',
    )
      ? await XmlHelper.getXmlFromArchive(archive, 'xl/tables/table1.xml')
      : undefined;

    const sharedStrings = await XmlHelper.getXmlFromArchive(
      archive,
      'xl/sharedStrings.xml',
    );

    return {
      archive,
      sheet,
      sharedStrings,
      table,
    };
  }

  async writeWorkbook(workbook: Workbook): Promise<void> {
    XmlHelper.writeXmlToArchive(
      workbook.archive,
      'xl/worksheets/sheet1.xml',
      workbook.sheet,
    );

    if (workbook.table) {
      XmlHelper.writeXmlToArchive(
        workbook.archive,
        'xl/tables/table1.xml',
        workbook.table,
      );
    }

    XmlHelper.writeXmlToArchive(
      workbook.archive,
      'xl/sharedStrings.xml',
      workbook.sharedStrings,
    );

    const worksheet = await workbook.archive.getContent({});
    await this.targetArchive.write(
      PptPaths.embedding(
      `${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
    ),
      worksheet,
    );
  }

  async copyFiles(): Promise<void> {
    await this.copyChartFiles();

    this.worksheetFilePrefix = await this.getWorksheetFilePrefix(
      this.wbRelsPath,
    );

    if (this.hasWorkbook) {
      const worksheets = await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
        this.sourceArchive,
        this.wbRelsPath,
        `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}`,
      );

      const worksheet = worksheets[0];

      this.sourceWorksheet = worksheet.number === 0 ? '' : worksheet.number;
      this.targetWorksheet = '-created-' + this.targetNumber;

      await this.copyWorksheetFile();
    } else {
      log.info('Chart has no worksheet: ' + this.wbRelsPath);
    }

    await this.editTargetWorksheetRel();
  }

  async getWorksheetFilePrefix(targetRelFile: string): Promise<string> {
    const relationTargets = await XmlRelationshipHelper.getRelationshipTargetsByPrefix(
      this.sourceArchive,
      targetRelFile,
      this.wbEmbeddingsPath,
    );

    if (!relationTargets[0]) {
      this.hasWorkbook = false;
      return '';
    }

    return relationTargets[0].filenameBase;
  }

  /**
   * Registers the copied chart parts in `[Content_Types].xml`.
   *
   * All style parts are optional: `copyChartStyleFiles` copies colors, style,
   * themeOverride and user shapes only if the source chart relates to them.
   * An Override for a part that was never copied stays behind as a dangling
   * entry — PowerPoint ignores it, but OPC validators report one warning per
   * orphan. `prepare` runs `copyChartStyleFiles` first, so the relations are
   * known here.
   */
  async appendTypes(): Promise<void> {
    if (this.hasWorkbook) {
      await this.appendChartExtensionToContentType();
    }

    await this.appendChartToContentType();

    if (this.hasStyleRelation('relTypeChartUserShapes')) {
      await this.appendChartUserShapesToContentType();
    }
    if (this.hasStyleRelation('relTypeChartColorStyle')) {
      await this.appendColorToContentType();
    }
    if (this.hasStyleRelation('relTypeChartStyle')) {
      await this.appendStyleToContentType();
    }
    if (this.hasStyleRelation('relTypeChartThemeOverride')) {
      await this.appendThemeOverrideToContentType();
    }
  }

  hasStyleRelation(styleType: string): boolean {
    return this.styleRelationFiles[styleType]?.length > 0;
  }

  async copyChartFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.chartPart(this.subtype, this.sourceNumber),
      this.targetArchive,
      PptPaths.chartPart(this.subtype, this.targetNumber),
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      PptPaths.chartPartRels(this.subtype, this.sourceNumber),
      this.targetArchive,
      PptPaths.chartPartRels(this.subtype, this.targetNumber),
    );
  }

  async copyChartStyleFiles(): Promise<void> {
    await this.getChartStyles();

    if (this.styleRelationFiles.relTypeChartStyle?.length) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartStyle[0]}`,
        this.targetArchive,
        PptPaths.chartPart('style', this.targetNumber),
      );
    }

    if (this.styleRelationFiles.relTypeChartColorStyle?.length) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartColorStyle[0]}`,
        this.targetArchive,
        PptPaths.chartPart('colors', this.targetNumber),
      );
    }

    if (this.styleRelationFiles.relTypeChartImage) {
      for (const relTypeChartImage of this.styleRelationFiles
        .relTypeChartImage) {
        const imageInfo = this.getTargetChartImageUri(relTypeChartImage);
        await this.appendImageExtensionToContentType(imageInfo.extension);
        await FileHelper.zipCopy(
          this.sourceArchive,
          imageInfo.source,
          this.targetArchive,
          imageInfo.target,
        );
      }
    }

    if (this.styleRelationFiles.relTypeChartUserShapes?.length) {
      const sourceFile =
        this.styleRelationFiles.relTypeChartUserShapes[0].replace(
          '../drawings/',
          '',
        );
      const sourceDrawing = `${PptPaths.drawingsDir}/${sourceFile}`;

      await FileHelper.zipCopy(
        this.sourceArchive,
        sourceDrawing,
        this.targetArchive,
        PptPaths.drawing(this.targetNumber),
      );

      await this.copyChartUserShapesRels(sourceDrawing);
    }

    if (this.styleRelationFiles.relTypeChartThemeOverride?.length) {
      const sourceFile =
        this.styleRelationFiles.relTypeChartThemeOverride[0].replace(
          '../theme/',
          '',
        );
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/theme/${sourceFile}`,
        this.targetArchive,
        `ppt/theme/themeOverride${this.targetNumber}.xml`,
      );
    }
  }

  /**
   * Copies the `_rels` file of a chart's `chartUserShapes` drawing, if it has
   * one, and copies+renames any image it relates to — e.g. a picture embedded
   * in a chart annotation, as think-cell produces — the same way
   * `editTargetWorksheetRel` handles the chart's own `relTypeChartImage`
   * relations. Without it, the copied drawing's `r:embed` resolves to nothing
   * and PowerPoint reports "The picture can't be displayed" (#215).
   *
   * The rels file is zip-copied to the target first and the *target* copy is
   * rewritten. The source archive caches one parsed DOM per path, so editing
   * the source rels in place would leak the renamed targets into every later
   * clone of the same chart.
   *
   * Only image relations are renamed. External relations (hyperlinks to URLs)
   * need no rewrite. Any other internal relation type is copied as-is with a
   * warning: its target is still relative to the source archive and may not
   * exist or may point at a different part in the output.
   *
   * @param sourceDrawing Path of the drawing part inside the source archive
   */
  private async copyChartUserShapesRels(sourceDrawing: string): Promise<void> {
    const sourceRels = PptPaths.relsOf(sourceDrawing);
    if (!FileHelper.fileExistsInArchive(this.sourceArchive, sourceRels)) {
      return;
    }

    const targetRels = PptPaths.drawingRels(this.targetNumber);
    await FileHelper.zipCopy(
      this.sourceArchive,
      sourceRels,
      this.targetArchive,
      targetRels,
    );

    const relsXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRels,
    );
    const relationships = relsXml.getElementsByTagName('Relationship');

    for (let i = 0; i < relationships.length; i++) {
      const relationship = relationships[i];
      const type = relationship.getAttribute('Type');
      const isExternal = relationship.getAttribute('TargetMode') === 'External';

      if (!isExternal && type === this.relTypeChartImage) {
        const imageInfo = this.getTargetChartImageUri(
          relationship.getAttribute('Target'),
        );
        await this.appendImageExtensionToContentType(imageInfo.extension);
        await FileHelper.zipCopy(
          this.sourceArchive,
          imageInfo.source,
          this.targetArchive,
          imageInfo.target,
        );
        relationship.setAttribute('Target', imageInfo.rel);
      } else if (!isExternal) {
        log.warn(
          `Chart user shapes drawing ${sourceDrawing} relates to ` +
            `${relationship.getAttribute('Target')} (${type}); this ` +
            `relationship type is copied unchanged and may not resolve in the output.`,
        );
      }

      this.targetArchive.contentTracker?.trackRelation(targetRels, {
        Id: relationship.getAttribute('Id'),
        Target: relationship.getAttribute('Target'),
        Type: type,
      });
    }

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRels, relsXml);
  }

  async getChartStyles(): Promise<void> {
    const styleTypes = [
      'relTypeChartStyle',
      'relTypeChartColorStyle',
      'relTypeChartImage',
      'relTypeChartThemeOverride',
      'relTypeChartUserShapes',
    ];

    for (const i in styleTypes) {
      const styleType = styleTypes[i];
      const styleRelation = await XmlRelationshipHelper.getTargetsByRelationshipType(
        this.sourceArchive,
        this.wbRelsPath,
        (this as unknown as Record<string, string>)[styleType],
      );

      this.styleRelationFiles[styleType] =
        this.styleRelationFiles[styleType] || [];
      if (styleRelation.length) {
        styleRelation.forEach((styleRelation) => {
          this.styleRelationFiles[styleType].push(styleRelation.file);
        });
      }
    }
  }

  async appendToSlideRels(): Promise<XmlElement> {
    this.createdRid = await XmlRelationshipHelper.getNextRelId(
      this.targetArchive,
      this.targetSlideRelFile,
    );

    const type =
      this.subtype === 'chart'
        ? 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
        : 'http://schemas.microsoft.com/office/2014/relationships/chartEx';

    const attributes = {
      Id: this.createdRid,
      Type: type,
      Target: `../charts/${this.subtype}${this.targetNumber}.xml`,
    } as RelationshipAttribute;

    return XmlHelper.append(
      XmlRelationshipHelper.createRelationshipChild(
        this.targetArchive,
        this.targetSlideRelFile,
        attributes,
      ),
    );
  }

  async editTargetWorksheetRel(): Promise<void> {
    const targetRelFile = PptPaths.chartPartRels(this.subtype, this.targetNumber);
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relations = relXml.getElementsByTagName('Relationship');

    Object.keys(relations)
      .map((key) => relations[Number(key)])
      .filter((element) => element && element.getAttribute)
      .forEach((element) => {
        const type = element.getAttribute('Type');
        switch (type) {
          case 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package':
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
            );
            break;
          case this.relTypeChartColorStyle:
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              `colors${this.targetNumber}.xml`,
            );
            break;
          case this.relTypeChartStyle:
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              `style${this.targetNumber}.xml`,
            );
            break;
          case this.relTypeChartImage:
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              this.getTargetChartImageUri(element.getAttribute('Target')).rel,
            );
            break;
          case this.relTypeChartThemeOverride:
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              `../theme/themeOverride${this.targetNumber}.xml`,
            );
            break;
          case this.relTypeChartUserShapes:
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              PptPaths.relative.drawing(this.targetNumber),
            );
            break;
        }
        this.targetArchive.contentTracker?.trackRelation(targetRelFile, {
          Id: element.getAttribute('Id'),
          Target: element.getAttribute('Target'),
          Type: element.getAttribute('Type'),
        });
      });

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
  }

  updateTargetWorksheetRelation(
    targetRelFile: string,
    element: XmlElement,
    attribute: string,
    value: string,
  ) {
    element.setAttribute(attribute, value);
  }

  getTargetChartImageUri(origin: string): {
    source: string;
    target: string;
    rel: string;
    extension: ContentTypeExtension;
  } {
    const file = origin.replace('../media/', '');
    const extension = path
      .extname(file)
      .replace('.', '') as ContentTypeExtension;

    return {
      source: PptPaths.media(file),
      target: PptPaths.media(
        `${file}-chart-${this.targetNumber}.${extension}`,
      ),
      rel: `../media/${file}-chart-${this.targetNumber}.${extension}`,
      extension: extension,
    };
  }

  async copyWorksheetFile(): Promise<void> {
    const sourceFile = PptPaths.embedding(
      `${this.worksheetFilePrefix}${this.sourceWorksheet}${this.wbExtension}`,
    );
    const targetFile = PptPaths.embedding(
      `${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      sourceFile,
      this.targetArchive,
      targetFile,
    ).catch((e) => {
      log.warn(e);
    });
  }

  appendChartExtensionToContentType(): Promise<XmlElement | boolean> {
    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(this.targetArchive, {
        Extension: `xlsx`,
        ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
      }),
      tag: 'Default',
      clause: (xml: XmlDocument) =>
        !XmlHelper.findByAttribute(xml, 'Default', 'Extension', 'xlsx'),
    });
  }

  appendChartToContentType(): Promise<XmlElement> {
    const contentType =
      this.subtype === 'chart'
        ? 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
        : 'application/vnd.ms-office.chartex+xml';

    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: PptPaths.partName(
          PptPaths.chartPart(this.subtype, this.targetNumber),
        ),
        ContentType: contentType,
      }),
    );
  }

  appendColorToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: PptPaths.partName(
          PptPaths.chartPart('colors', this.targetNumber),
        ),
        ContentType: `application/vnd.ms-office.chartcolorstyle+xml`,
      }),
    );
  }

  appendStyleToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: PptPaths.partName(
          PptPaths.chartPart('style', this.targetNumber),
        ),
        ContentType: `application/vnd.ms-office.chartstyle+xml`,
      }),
    );
  }

  appendThemeOverrideToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/theme/themeOverride${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.themeOverride+xml`,
      }),
    );
  }

  appendChartUserShapesToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: PptPaths.partName(PptPaths.drawing(this.targetNumber)),
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml`,
      }),
    );
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    return await XmlRelationshipHelper.getRelationshipTargetsByPrefix(archive, relsPath, [
      '../charts/chart',
      '../charts/chartEx',
    ]);
  }
}
