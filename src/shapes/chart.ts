import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
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
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { contentTracker } from '../helper/content-tracker';
import IArchive from '../interfaces/iarchive';
import { ContentTypeExtension } from '../enums/content-type-map';
import { vd } from '../helper/general-helper';

export class Chart extends Shape implements IChart {
  sourceWorksheet: number | string;
  targetWorksheet: number | string;
  worksheetFilePrefix: string;
  wbEmbeddingsPath: string;
  wbExtension: string;
  relTypeChartColorStyle: string;
  relTypeChartStyle: string;
  relTypeChartImage: string;
  relTypeChartThemeOverride: string;
  wbRelsPath: string;
  styleRelationFiles: {
    [key: string]: string[];
  };
  callbacks: ChartModificationCallback[];

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
    this.styleRelationFiles = {};
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
    this.wbRelsPath = `ppt/charts/_rels/${this.subtype}${this.sourceNumber}.xml.rels`;

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
    const chartXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      `ppt/charts/${this.subtype}${this.targetNumber}.xml`,
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
      `ppt/charts/${this.subtype}${this.targetNumber}.xml`,
      chartXml,
    );
    await this.writeWorkbook(workbook);
  }

  async readWorkbook(): Promise<Workbook> {
    const workbookFilename = `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`;
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
      `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
      worksheet,
    );
  }

  async copyFiles(): Promise<void> {
    await this.copyChartFiles();

    this.worksheetFilePrefix = await this.getWorksheetFilePrefix(
      this.wbRelsPath,
    );

    const worksheets = await XmlHelper.getRelationshipTargetsByPrefix(
      this.sourceArchive,
      this.wbRelsPath,
      `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}`,
    );

    const worksheet = worksheets[0];

    this.sourceWorksheet = worksheet.number === 0 ? '' : worksheet.number;
    this.targetWorksheet = '-created-' + this.targetNumber;

    await this.copyWorksheetFile();
    await this.editTargetWorksheetRel();
  }

  async getWorksheetFilePrefix(targetRelFile: string): Promise<string> {
    const relationTargets = await XmlHelper.getRelationshipTargetsByPrefix(
      this.sourceArchive,
      targetRelFile,
      this.wbEmbeddingsPath,
    );

    if (!relationTargets[0]) {
      throw new Error(
        `Could not find a related worksheet pointing to ${this.wbEmbeddingsPath}@${targetRelFile}`,
      );
    }

    return relationTargets[0].filenameBase;
  }

  async appendTypes(): Promise<void> {
    await this.appendChartExtensionToContentType();
    await this.appendChartToContentType();
    await this.appendColorToContentType();
    await this.appendStyleToContentType();
    await this.appendThemeOverrideToContentType();
  }

  async copyChartFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/charts/${this.subtype}${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/charts/${this.subtype}${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/charts/_rels/${this.subtype}${this.sourceNumber}.xml.rels`,
      this.targetArchive,
      `ppt/charts/_rels/${this.subtype}${this.targetNumber}.xml.rels`,
    );
  }

  async copyChartStyleFiles(): Promise<void> {
    await this.getChartStyles();

    if (this.styleRelationFiles.relTypeChartStyle?.length) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartStyle[0]}`,
        this.targetArchive,
        `ppt/charts/style${this.targetNumber}.xml`,
      );
    }

    if (this.styleRelationFiles.relTypeChartColorStyle?.length) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartColorStyle[0]}`,
        this.targetArchive,
        `ppt/charts/colors${this.targetNumber}.xml`,
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

  async getChartStyles(): Promise<void> {
    const styleTypes = [
      'relTypeChartStyle',
      'relTypeChartColorStyle',
      'relTypeChartImage',
      'relTypeChartThemeOverride',
    ];

    for (const i in styleTypes) {
      const styleType = styleTypes[i];
      const styleRelation = await XmlHelper.getTargetsByRelationshipType(
        this.sourceArchive,
        this.wbRelsPath,
        this[styleType],
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
    this.createdRid = await XmlHelper.getNextRelId(
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
      XmlHelper.createRelationshipChild(
        this.targetArchive,
        this.targetSlideRelFile,
        attributes,
      ),
    );
  }

  async editTargetWorksheetRel(): Promise<void> {
    const targetRelFile = `ppt/charts/_rels/${this.subtype}${this.targetNumber}.xml.rels`;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relations = relXml.getElementsByTagName('Relationship');

    Object.keys(relations)
      .map((key) => relations[key])
      .filter((element) => element.getAttribute)
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
            const target = element.getAttribute('Target');
            const imageInfo = this.getTargetChartImageUri(target);
            this.updateTargetWorksheetRelation(
              targetRelFile,
              element,
              'Target',
              imageInfo.rel,
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
        }
        contentTracker.trackRelation(targetRelFile, {
          Id: element.getAttribute('Id'),
          Target: element.getAttribute('Target'),
          Type: element.getAttribute('Type'),
        });
      });

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
  }

  updateTargetWorksheetRelation(targetRelFile, element, attribute, value) {
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
      source: `ppt/media/${file}`,
      target: `ppt/media/${file}-chart-${this.targetNumber}.${extension}`,
      rel: `../media/${file}-chart-${this.targetNumber}.${extension}`,
      extension: extension,
    };
  }

  async copyWorksheetFile(): Promise<void> {
    const targetFile = `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`;
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/embeddings/${this.worksheetFilePrefix}${this.sourceWorksheet}${this.wbExtension}`,
      this.targetArchive,
      targetFile,
    );
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
        PartName: `/ppt/charts/${this.subtype}${this.targetNumber}.xml`,
        ContentType: contentType,
      }),
    );
  }

  appendColorToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/colors${this.targetNumber}.xml`,
        ContentType: `application/vnd.ms-office.chartcolorstyle+xml`,
      }),
    );
  }

  appendStyleToContentType(): Promise<XmlElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/style${this.targetNumber}.xml`,
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

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    return await XmlHelper.getRelationshipTargetsByPrefix(archive, relsPath, [
      '../charts/chart',
      '../charts/chartEx',
    ]);
  }
}
