import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import path from 'path';

export class OLEObject extends Shape {
  private readonly oleObjectPath: string;

  constructor(
    shape: ImportedElement,
    targetType: ShapeTargetType,
    sourceArchive: IArchive,
  ) {
    super(shape, targetType);
    this.sourceArchive = sourceArchive;
    this.oleObjectPath = `ppt/embeddings/${
      this.sourceRid
    }${this.getFileExtension(shape.target?.file)}`;
    this.relRootTag = 'p:oleObj';
    this.relAttribute = 'r:id';
  }

  private getFileExtension(file?: string): string {
    if (!file) return '.bin';
    const ext = path.extname(file).toLowerCase();
    return ['.bin', '.xls', '.xlsx', '.doc', '.docx', '.ppt', '.pptx'].includes(
      ext,
    )
      ? ext
      : '.bin';
  }

  // NOTE: modify() and append() won't be implemented.

  // TODO: remove is not currently properly implemented,
  //  suggest we delete the file from the archive as well as removing the relationship.
  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<OLEObject> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.removeFromSlideTree();
    await this.removeOleObjectFile();
    await this.removeFromContentTypes();
    await this.removeFromSlideRels();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    oleObjects?: Target[],
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    const allOleObjects =
      oleObjects ||
      (await OLEObject.getAllOnSlide(
        this.sourceArchive,
        this.targetSlideRelFile,
      ));

    const oleObject = allOleObjects.find((obj) => obj.rId === this.sourceRid);
    if (!oleObject) {
      throw new Error(`OLE object with rId ${this.sourceRid} not found.`);
    }

    const sourceFilePath = `ppt/embeddings/${oleObject.file.split('/').pop()}`;

    this.createdRid = await XmlHelper.getNextRelId(
      this.targetArchive,
      this.targetSlideRelFile,
    );

    await this.copyOleObjectFile(sourceFilePath);
    await this.appendToContentTypes();
    await this.updateSlideRels();
    await this.updateSlideXml();
  }

  private async copyOleObjectFile(sourceFilePath: string): Promise<void> {
    const fileExtension = this.getFileExtension(sourceFilePath);
    const targetFileName = `ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;

    try {
      await FileHelper.zipCopy(
        this.sourceArchive,
        sourceFilePath,
        this.targetArchive,
        targetFileName,
      );
    } catch (error) {
      console.error('Error copying OLE object file:', error);
      throw error;
    }
  }

  private async appendToContentTypes(): Promise<void> {
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      contentTypesPath,
    );

    const types = contentTypesXml.getElementsByTagName('Types')[0];
    const fileExtension = this.getFileExtension(this.oleObjectPath);
    const partName = `/ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
    const existingOverride = Array.from(
      types.getElementsByTagName('Override'),
    ).find((override) => override.getAttribute('PartName') === partName);

    if (!existingOverride) {
      const newOverride = contentTypesXml.createElement('Override');
      newOverride.setAttribute('PartName', partName);
      newOverride.setAttribute(
        'ContentType',
        this.getContentType(fileExtension),
      );
      types.appendChild(newOverride);

      await XmlHelper.writeXmlToArchive(
        this.targetArchive,
        contentTypesPath,
        contentTypesXml,
      );
    }
  }

  private async updateSlideRels(): Promise<void> {
    const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');

    const fileExtension = this.getFileExtension(this.oleObjectPath);
    const newTarget = `../embeddings/oleObject${this.createdRid}${fileExtension}`;

    // Update or create the relationship
    let relationshipUpdated = false;
    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.sourceRid) {
        relationships[i].setAttribute('Id', this.createdRid);
        relationships[i].setAttribute('Target', newTarget);
        relationshipUpdated = true;
        break;
      }
    }

    if (!relationshipUpdated) {
      const newRel = relXml.createElement('Relationship');
      newRel.setAttribute('Id', this.createdRid);
      newRel.setAttribute(
        'Type',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
      );
      newRel.setAttribute('Target', newTarget);
      relXml.documentElement.appendChild(newRel);
    }

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  private async updateSlideXml(): Promise<void> {
    const slideXmlPath = `ppt/slides/slide${this.targetSlideNumber}.xml`;
    const slideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      slideXmlPath,
    );

    const oleObjs = Array.from(slideXml.getElementsByTagName('p:oleObj'));
    oleObjs.forEach((oleObj) => {
      if (oleObj.getAttribute('r:id') === this.sourceRid) {
        oleObj.setAttribute('r:id', this.createdRid);
        const oleObjPr = oleObj.getElementsByTagName('p:oleObjPr')[0];
        if (oleObjPr) {
          const links = Array.from(oleObjPr.getElementsByTagName('a:link'));
          links.forEach((link) => {
            link.setAttribute('r:id', this.createdRid);
          });
        }
      }
    });

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      slideXmlPath,
      slideXml,
    );
  }

  private getContentType(fileExtension: string): string {
    const contentTypes: { [key: string]: string } = {
      '.bin': 'application/vnd.openxmlformats-officedocument.oleObject',
      '.xls': 'application/vnd.ms-excel',
      '.xlsx':
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      '.doc': 'application/msword',
      '.docx':
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      '.ppt': 'application/vnd.ms-powerpoint',
      '.pptx':
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    };
    return (
      contentTypes[fileExtension.toLowerCase()] ||
      'application/vnd.openxmlformats-officedocument.oleObject'
    );
  }

  private async removeOleObjectFile(): Promise<void> {
    const fileExtension = this.getFileExtension(this.oleObjectPath);
    const fileName = `ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
    await this.targetArchive.remove(fileName);
  }

  private async removeFromContentTypes(): Promise<void> {
    const contentTypesPath = '[Content_Types].xml';
    const contentTypesXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      contentTypesPath,
    );

    const types = contentTypesXml.getElementsByTagName('Types')[0];
    const fileExtension = this.getFileExtension(this.oleObjectPath);
    const partName = `/ppt/embeddings/oleObject${this.createdRid}${fileExtension}`;
    const overrideToRemove = Array.from(
      types.getElementsByTagName('Override'),
    ).find((override) => override.getAttribute('PartName') === partName);

    if (overrideToRemove) {
      types.removeChild(overrideToRemove);
      await XmlHelper.writeXmlToArchive(
        this.targetArchive,
        contentTypesPath,
        contentTypesXml,
      );
    }
  }

  private async removeFromSlideRels(): Promise<void> {
    const targetRelFile = `ppt/${this.targetType}s/_rels/${this.targetType}${this.targetSlideNumber}.xml.rels`;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');

    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.createdRid) {
        relationships[i].parentNode.removeChild(relationships[i]);
        break;
      }
    }

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    const oleObjectType =
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';

    return XmlHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');

        if (type === oleObjectType) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            element: element,
          } as Target);
        }
      },
    );
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    oleObjects: Target[],
  ): Promise<void> {
    await this.prepare(targetTemplate, targetSlideNumber, oleObjects);
  }
}
