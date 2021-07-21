import JSZip from 'jszip';
import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';

import { RelationshipAttribute, HelperElement } from '../types/xml-types';
import { ImportedElement, Target, Workbook } from '../types/types';
import { IChart } from '../interfaces/ichart';
import { RootPresTemplate } from '../interfaces/root-pres-template';

export class Chart extends Shape implements IChart {
  sourceWorksheet: number | string;
  targetWorksheet: number | string;
  worksheetFilePrefix: string;
  wbEmbeddingsPath: string;
  wbExtension: string;
  relTypeChartColorStyle: string;
  relTypeChartStyle: string;
  wbRelsPath: string;
  styleRelationFiles: {
    [key: string]: string;
  };

  constructor(shape: ImportedElement) {
    super(shape);

    this.relRootTag = 'c:chart';
    this.relAttribute = 'r:id';
    this.relParent = (element: Element) =>
      element.parentNode.parentNode.parentNode as Element;

    this.wbEmbeddingsPath = `../embeddings/`;
    this.wbExtension = '.xlsx';
    this.relTypeChartColorStyle =
      'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
    this.relTypeChartStyle =
      'http://schemas.microsoft.com/office/2011/relationships/chartStyle';
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

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Chart> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.updateElementRelId();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    this.targetNumber = this.targetTemplate.incrementCounter('charts');
    this.wbRelsPath = `ppt/charts/_rels/chart${this.sourceNumber}.xml.rels`;

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
      `ppt/charts/chart${this.targetNumber}.xml`,
    );
    const workbook = await this.readWorkbook();

    this.applyCallbacks(this.callbacks, this.targetElement, chartXml, workbook);

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      `ppt/charts/chart${this.targetNumber}.xml`,
      chartXml,
    );
    await this.writeWorkbook(workbook);
  }

  async readWorkbook(): Promise<Workbook> {
    const worksheet = await FileHelper.extractFromArchive(
      this.targetArchive,
      `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
      'nodebuffer',
    );
    const archive = await FileHelper.extractFileContent(
      (worksheet as unknown) as Buffer,
    );
    const sheet = await XmlHelper.getXmlFromArchive(
      archive,
      'xl/worksheets/sheet1.xml',
    );
    const table = await XmlHelper.getXmlFromArchive(
      archive,
      'xl/tables/table1.xml',
    );
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
    await XmlHelper.writeXmlToArchive(
      workbook.archive,
      'xl/worksheets/sheet1.xml',
      workbook.sheet,
    );
    await XmlHelper.writeXmlToArchive(
      workbook.archive,
      'xl/tables/table1.xml',
      workbook.table,
    );
    await XmlHelper.writeXmlToArchive(
      workbook.archive,
      'xl/sharedStrings.xml',
      workbook.sharedStrings,
    );

    const worksheet = await workbook.archive.generateAsync({
      type: 'nodebuffer',
    });
    this.targetArchive.file(
      `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
      worksheet,
    );
  }

  async copyFiles(): Promise<void> {
    await this.copyChartFiles();

    this.worksheetFilePrefix = await this.getWorksheetFilePrefix(
      this.wbRelsPath,
    );

    const worksheets = await XmlHelper.getTargetsFromRelationships(
      this.sourceArchive,
      this.wbRelsPath,
      `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}`,
      this.wbExtension,
    );
    const worksheet = worksheets[0];

    this.sourceWorksheet = worksheet.number === 0 ? '' : worksheet.number;
    this.targetWorksheet = this.targetNumber;

    await this.copyWorksheetFile();
    await this.editTargetWorksheetRel();
  }

  async getWorksheetFilePrefix(targetRelFile: string): Promise<string> {
    const relationTargets = await XmlHelper.getTargetsFromRelationships(
      this.sourceArchive,
      targetRelFile,
      this.wbEmbeddingsPath,
      this.wbExtension,
    );

    const wbPath = relationTargets[0].file
      .replace(this.wbEmbeddingsPath, '')
      .replace(this.wbExtension, '');

    const wbTitle = wbPath.match(/^(.+?)(\d+)*$/);

    return wbTitle[0];
  }

  async appendTypes(): Promise<void> {
    await this.appendChartExtensionToContentType();
    await this.appendChartToContentType();
    await this.appendColorToContentType();
    await this.appendStyleToContentType();
  }

  async copyChartFiles(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/charts/chart${this.sourceNumber}.xml`,
      this.targetArchive,
      `ppt/charts/chart${this.targetNumber}.xml`,
    );

    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/charts/_rels/chart${this.sourceNumber}.xml.rels`,
      this.targetArchive,
      `ppt/charts/_rels/chart${this.targetNumber}.xml.rels`,
    );
  }

  async copyChartStyleFiles(): Promise<void> {
    await this.getChartStyles();

    if (this.styleRelationFiles.relTypeChartStyle) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartStyle}`,
        this.targetArchive,
        `ppt/charts/style${this.targetNumber}.xml`,
      );
    }

    if (this.styleRelationFiles.relTypeChartColorStyle) {
      await FileHelper.zipCopy(
        this.sourceArchive,
        `ppt/charts/${this.styleRelationFiles.relTypeChartColorStyle}`,
        this.targetArchive,
        `ppt/charts/colors${this.targetNumber}.xml`,
      );
    }
  }

  async getChartStyles(): Promise<void> {
    const styleTypes = ['relTypeChartStyle', 'relTypeChartColorStyle'];

    for (const i in styleTypes) {
      const styleType = styleTypes[i];
      const styleRelation = await XmlHelper.getTargetsByRelationshipType(
        this.sourceArchive,
        this.wbRelsPath,
        this[styleType],
      );

      if (styleRelation.length) {
        this.styleRelationFiles[styleType] = styleRelation[0].file;
      }
    }
  }

  async appendToSlideRels(): Promise<HelperElement> {
    this.createdRid = await XmlHelper.getNextRelId(
      this.targetArchive,
      this.targetSlideRelFile,
    );
    const attributes = {
      Id: this.createdRid,
      Type:
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
      Target: `../charts/chart${this.targetNumber}.xml`,
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
    const targetRelFile = `ppt/charts/_rels/chart${this.targetNumber}.xml.rels`;
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
            element.setAttribute(
              'Target',
              `${this.wbEmbeddingsPath}${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
            );
            break;
          case this.relTypeChartColorStyle:
            element.setAttribute('Target', `colors${this.targetNumber}.xml`);
            break;
          case this.relTypeChartStyle:
            element.setAttribute('Target', `style${this.targetNumber}.xml`);
            break;
        }
      });

    XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
  }

  async copyWorksheetFile(): Promise<void> {
    await FileHelper.zipCopy(
      this.sourceArchive,
      `ppt/embeddings/${this.worksheetFilePrefix}${this.sourceWorksheet}${this.wbExtension}`,
      this.targetArchive,
      `ppt/embeddings/${this.worksheetFilePrefix}${this.targetWorksheet}${this.wbExtension}`,
    );
  }

  appendChartExtensionToContentType(): Promise<HelperElement | boolean> {
    return XmlHelper.appendIf({
      ...XmlHelper.createContentTypeChild(this.targetArchive, {
        Extension: `xlsx`,
        ContentType: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
      }),
      tag: 'Default',
      clause: (xml: XMLDocument) =>
        !XmlHelper.findByAttribute(xml, 'Default', 'Extension', 'xlsx'),
    });
  }

  appendChartToContentType(): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/chart${this.targetNumber}.xml`,
        ContentType: `application/vnd.openxmlformats-officedocument.drawingml.chart+xml`,
      }),
    );
  }

  appendColorToContentType(): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/colors${this.targetNumber}.xml`,
        ContentType: `application/vnd.ms-office.chartcolorstyle+xml`,
      }),
    );
  }

  appendStyleToContentType(): Promise<HelperElement> {
    return XmlHelper.append(
      XmlHelper.createContentTypeChild(this.targetArchive, {
        PartName: `/ppt/charts/style${this.targetNumber}.xml`,
        ContentType: `application/vnd.ms-office.chartstyle+xml`,
      }),
    );
  }

  static async getAllOnSlide(
    archive: JSZip,
    relsPath: string,
  ): Promise<Target[]> {
    return await XmlHelper.getTargetsFromRelationships(
      archive,
      relsPath,
      '../charts/chart',
    );
  }
}
