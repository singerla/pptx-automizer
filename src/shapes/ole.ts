import { FileHelper } from '../helper/file-helper';
import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { contentTracker } from '../helper/content-tracker';
import path from 'path';

export class OLEObject extends Shape {
	private readonly oleObjectPath: string;

	constructor(shape: ImportedElement, targetType: ShapeTargetType, sourceArchive: IArchive) {
		super(shape, targetType);
		this.sourceArchive = sourceArchive;
		this.oleObjectPath = `ppt/embeddings/${this.sourceRid}${this.getFileExtension(shape.target?.file)}`;
		this.relRootTag = 'p:oleObj';
		this.relAttribute = 'r:id';
	}

	private getFileExtension(file?: string): string {
		if (!file) return '.bin';
		const ext = path.extname(file).toLowerCase();
		return ['.bin', '.xls', '.xlsx', '.doc', '.docx', '.ppt', '.pptx'].includes(ext) ? ext : '.bin';
	}

	// NOTE: modify is not currently implemented
	// async modify(
	// 	targetTemplate: RootPresTemplate,
	// 	targetSlideNumber: number,
	// ): Promise<OLEObject> {
	// 	await this.prepare(targetTemplate, targetSlideNumber);
	// 	await this.replaceIntoSlideTree();
	// 	await this.editTargetOleObjectRel();

	// 	return this;
	// }

	// NOTE: append is not currently implemented
	// async append(
	// 	targetTemplate: RootPresTemplate,
	// 	targetSlideNumber: number,
	// ): Promise<OLEObject> {
	// 	await this.prepare(targetTemplate, targetSlideNumber);
	// 	await this.appendToSlideTree();

	// 	return this;
	// }

    // TODO: remove is not currently properly implemented, suggest we delete the file from the archive as well as removing the relationship.
	async remove(
		targetTemplate: RootPresTemplate,
		targetSlideNumber: number,
	): Promise<OLEObject> {
		await this.prepare(targetTemplate, targetSlideNumber);
		await this.removeFromSlideTree();

		return this;
	}

	async prepare(
		targetTemplate: RootPresTemplate,
		targetSlideNumber: number,
		oleObjects?: Target[]
	): Promise<void> {
		await this.setTarget(targetTemplate, targetSlideNumber);

		this.targetNumber = this.targetTemplate.incrementCounter('oleObjects');

		const allOleObjects = oleObjects || await OLEObject.getAllOnSlide(this.sourceArchive, this.targetSlideRelFile);

		const oleObject = allOleObjects.find(obj => obj.rId === this.sourceRid);
		if (!oleObject) {
			throw new Error(`OLE object with rId ${this.sourceRid} not found.`);
		}

		const sourceFilePath = `ppt/embeddings/${oleObject.file.split('/').pop()}`;

		await this.copyFiles(sourceFilePath);
		await this.appendTypes();
	}

	private async copyFiles(sourceFilePath: string): Promise<void> {
		if (!this.createdRid) {
			this.createdRid = await XmlHelper.getNextRelId(this.targetArchive, this.targetSlideRelFile);
		}

		const fileExtension = this.getFileExtension(sourceFilePath);
		const targetFileName = `ppt/embeddings/${this.createdRid}${fileExtension}`;

		try {
			await FileHelper.zipCopy(
				this.sourceArchive,
				sourceFilePath,
				this.targetArchive,
				targetFileName
			);
		} catch (error) {
			console.error("Error copying OLE object file:", error);
			throw error;
		}
	}

	private async appendTypes(): Promise<void> {
		await this.appendOleObjectToContentType();
	}

	private async editTargetOleObjectRel(): Promise<void> {
		const targetRelFile = this.targetSlideRelFile;
		const relXml = await XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
		const relations = relXml.getElementsByTagName('Relationship');

		Array.from(relations).forEach((element) => {
			const type = element.getAttribute('Type');
			if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject') {
				const fileExtension = this.getFileExtension(element.getAttribute('Target'));
				this.updateTargetOleObjectRelation(
					element,
					'Target',
					`../embeddings/${this.createdRid}${fileExtension}`,
				);
			}
		});

		XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);
	}

	private updateTargetOleObjectRelation(element: Element, attribute: string, value: string): void {
		element.setAttribute(attribute, value);
		contentTracker.trackRelation(this.targetSlideRelFile, {
			Id: element.getAttribute('Id') || '',
			Target: value,
			Type: element.getAttribute('Type') || '',
		});
	}

	private async appendOleObjectToContentType(): Promise<void> {
		const contentTypesPath = '[Content_Types].xml';
		const contentTypesXml = await XmlHelper.getXmlFromArchive(this.targetArchive, contentTypesPath);
		
		const types = contentTypesXml.getElementsByTagName('Types')[0];
		const fileExtension = this.getFileExtension(this.oleObjectPath);
		const existingOverride = Array.from(types.getElementsByTagName('Override')).find(
			(override) => override.getAttribute('PartName').endsWith(fileExtension)
		);

		if (!existingOverride) {
			const newOverride = contentTypesXml.createElement('Override');
			newOverride.setAttribute('PartName', `/ppt/embeddings/${this.createdRid}${fileExtension}`);
			newOverride.setAttribute('ContentType', this.getContentType(fileExtension));
			types.appendChild(newOverride);

			XmlHelper.writeXmlToArchive(this.targetArchive, contentTypesPath, contentTypesXml);
		}
	}

	private getContentType(fileExtension: string): string {
		const contentTypes: { [key: string]: string } = {
			'.bin': 'application/vnd.openxmlformats-officedocument.oleObject',
			'.xls': 'application/vnd.ms-excel',
			'.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			'.doc': 'application/msword',
			'.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
			'.ppt': 'application/vnd.ms-powerpoint',
			'.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		};
		return contentTypes[fileExtension.toLowerCase()] || 'application/vnd.openxmlformats-officedocument.oleObject';
	}

	static async getAllOnSlide(
		archive: IArchive,
		relsPath: string,
	): Promise<Target[]> {
		const oleObjectType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject';

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
			}
		);
	}

	async modifyOnAddedSlide(
		targetTemplate: RootPresTemplate,
		targetSlideNumber: number,
		oleObjects: Target[]
	): Promise<void> {
		await this.prepare(targetTemplate, targetSlideNumber, oleObjects);
		await this.editTargetOleObjectRel();
	}
}
