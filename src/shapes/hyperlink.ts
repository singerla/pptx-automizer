import { XmlHelper } from '../helper/xml-helper';
import { Shape } from '../classes/shape';
import { ImportedElement, ShapeTargetType, Target } from '../types/types';
import { XmlElement } from '../types/xml-types';
import IArchive from '../interfaces/iarchive';
import { RootPresTemplate } from '../interfaces/root-pres-template';
import { contentTracker } from '../helper/content-tracker';

export class Hyperlink extends Shape {
  private hyperlinkType: 'internal' | 'external';
  private hyperlinkTarget: string;

  constructor(
    shape: ImportedElement,
    targetType: ShapeTargetType,
    sourceArchive: IArchive,
    hyperlinkType: 'internal' | 'external' = 'external',
    hyperlinkTarget: string =  '',
  ) {
    super(shape, targetType);
    this.sourceArchive = sourceArchive;
    this.hyperlinkType = hyperlinkType;
    this.hyperlinkTarget = hyperlinkTarget;
    this.relRootTag = 'a:hlinkClick';
    this.relAttribute = 'r:id';
  }

  async modify(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.editTargetHyperlinkRel();
    await this.replaceIntoSlideTree();

    return this;
  }

  async append(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);
    await this.setTargetElement();
    await this.appendToSlideTree();
    await this.editTargetHyperlinkRel();
    await this.updateHyperlinkInSlide();

    return this;
  }

  async remove(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
  ): Promise<Hyperlink> {
    await this.prepare(targetTemplate, targetSlideNumber);
    if (this.target && this.target.rId) {
      this.sourceRid = this.target.rId;
    }
    await this.removeFromSlideTree();
    await this.removeFromSlideRels();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number
  ): Promise<void> {
    await this.setTarget(targetTemplate, targetSlideNumber);

    if (!this.createdRid) {
      const baseId = await XmlHelper.getNextRelId(
        this.targetArchive,
        this.targetSlideRelFile,
      );
      // Strip '-created' suffix more robustly
      this.createdRid = baseId.endsWith('-created') 
        ? baseId.slice(0, -8) 
        : baseId;
    }
    if (this.shape && this.shape.target && this.shape.target.rId) {
      this.sourceRid = this.shape.target.rId;
    }
    // If hyperlinkTarget is not set, try to get it from the original rel target
    if (!this.hyperlinkTarget && this.shape && this.shape.target && this.shape.target.file) {
      this.hyperlinkTarget = this.shape.target.file;
      this.hyperlinkType = (this.shape.target.isExternal || this.shape.target.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') ? 'external' : 'internal';
    }
  }

  private async editTargetHyperlinkRel(): Promise<void> {
    const targetRelFile = this.targetSlideRelFile;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');

    // Check if the relationship already exists
    let relationshipExists = false;
    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.createdRid) {
        this.updateHyperlinkRelation(relationships[i]);
        relationshipExists = true;
        break;
      }
    }

    // If the relationship doesn't exist, create it
    if (!relationshipExists) {
      const newRel = relXml.createElement('Relationship');
      newRel.setAttribute('Id', this.createdRid);
      newRel.setAttribute('Type', this.getRelationshipType());
      newRel.setAttribute('Target', this.getRelationshipTarget());
      if (this.hyperlinkType === 'external') {
        newRel.setAttribute('TargetMode', 'External');
      }
      relXml.documentElement.appendChild(newRel);
      
      // Track the relationship for content integrity
      contentTracker.trackRelation(targetRelFile, {
        Id: this.createdRid,
        Target: this.getRelationshipTarget(),
        Type: this.getRelationshipType(),
      });
    }

    XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  // Add a method to update the hyperlink in the slide XML
  private async updateHyperlinkInSlide(): Promise<void> {
    if (!this.targetElement && this.sourceRid && this.createdRid) {
      const slideXml = await XmlHelper.getXmlFromArchive(
        this.targetArchive,
        this.targetSlideFile,
      );
      const allHyperlinkElements = slideXml.getElementsByTagName('a:hlinkClick');
      let foundAndUpdatedInSlide = false;
      for (let i = 0; i < allHyperlinkElements.length; i++) {
        const hlinkClick = allHyperlinkElements[i];
        if (hlinkClick.getAttribute('r:id') === this.sourceRid) {
          hlinkClick.setAttribute('r:id', this.createdRid);
          hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
          foundAndUpdatedInSlide = true;
          break;
        }
      }
      if (foundAndUpdatedInSlide) {
        XmlHelper.writeXmlToArchive(
          this.targetArchive,
          this.targetSlideFile,
          slideXml,
        );
      }
    }
  }

  private updateHyperlinkRelation(element: Element): void {
    element.setAttribute('Type', this.getRelationshipType());
    element.setAttribute('Target', this.getRelationshipTarget());
    
    if (this.hyperlinkType === 'external') {
      element.setAttribute('TargetMode', 'External');
    } else if (element.hasAttribute('TargetMode')) {
      element.removeAttribute('TargetMode');
    }

    contentTracker.trackRelation(this.targetSlideRelFile, {
      Id: element.getAttribute('Id') || '',
      Target: this.getRelationshipTarget(),
      Type: this.getRelationshipType(),
    });
  }

  private getRelationshipType(): string {
    if (this.hyperlinkType === 'internal') {
      return 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
    }
    return 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  }

  private getRelationshipTarget(): string {
    if (this.hyperlinkType === 'internal') {
      // Enhanced internal slide link handling
      const slideNumber = this.hyperlinkTarget?.match(/\d+/)?.[0] || this.targetSlideNumber.toString();
      // Ensure proper slide path format
      return `../slides/slide${slideNumber}.xml`;
    }
    return this.hyperlinkTarget || 'https://example.com';
  }

  private async removeFromSlideRels(): Promise<void> {
    const targetRelFile = this.targetSlideRelFile;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    const relationships = relXml.getElementsByTagName('Relationship');
    const ridToRemove = this.sourceRid || this.createdRid;
    if (ridToRemove) {
      for (let i = relationships.length - 1; i >= 0; i--) {
        if (relationships[i].getAttribute('Id') === ridToRemove) {
          relationships[i].parentNode.removeChild(relationships[i]);
          break;
        }
      }
      XmlHelper.writeXmlToArchive(
        this.targetArchive,
        targetRelFile,
        relXml,
      );
    }
  }

  static async getAllOnSlide(
    archive: IArchive,
    relsPath: string,
  ): Promise<Target[]> {
    const hyperlinkRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
    const slideRelType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
    return XmlHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');

        if (type === hyperlinkRelType || type === slideRelType) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            filename: element.getAttribute('Target'),
            element: element,
            isExternal: element.getAttribute('TargetMode') === 'External' || type === hyperlinkRelType,
          } as Target);
        }
      }
    );
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number
  ): Promise<void> {
    if (!this.target || !this.target.rId) {
      console.warn('modifyOnAddedSlide called on Hyperlink without a valid source target/rId.');
      return;
    }
    this.sourceRid = this.target.rId;
    this.hyperlinkTarget = this.target.file;
    this.hyperlinkType = (this.target.isExternal || this.target.type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') ? 'external' : 'internal';


    await this.prepare(targetTemplate, targetSlideNumber);

    // 1. Modify the copied _rels file:
    const targetRelFile = this.targetSlideRelFile;
    const relXml = await XmlHelper.getXmlFromArchive(this.targetArchive, targetRelFile);
    const relationships = relXml.getElementsByTagName('Relationship');
    let relFoundAndUpdated = false;

    for (let i = 0; i < relationships.length; i++) {
      const relElement = relationships[i];
      if (relElement.getAttribute('Id') === this.sourceRid) { // Find by original rId
        relElement.setAttribute('Id', this.createdRid); // Update Id to new unique rId

        relElement.setAttribute('Target', this.getRelationshipTarget());

        if (this.hyperlinkType === 'external') {
          relElement.setAttribute('TargetMode', 'External');
        } else {
          if (relElement.hasAttribute('TargetMode')) relElement.removeAttribute('TargetMode');
        }
        relFoundAndUpdated = true;

        contentTracker.trackRelation(targetRelFile, {
          Id: this.createdRid,
          Target: relElement.getAttribute('Target') || '',
          Type: relElement.getAttribute('Type') || '',
        });
        break;
      }
    }

    if (!relFoundAndUpdated) {
      console.warn(`modifyOnAddedSlide: Relationship with sourceRId ${this.sourceRid} not found in target rels ${targetRelFile}. It might have been processed by another instance or was missing in the copied rels.`);
      const newRel = relXml.createElement('Relationship');
      newRel.setAttribute('Id', this.createdRid);
      newRel.setAttribute('Type', this.getRelationshipType());
      newRel.setAttribute('Target', this.getRelationshipTarget());
      if (this.hyperlinkType === 'external') {
        newRel.setAttribute('TargetMode', 'External');
      }
      relXml.documentElement.appendChild(newRel);
      contentTracker.trackRelation(targetRelFile, {
        Id: this.createdRid, Target: this.getRelationshipTarget(), Type: this.getRelationshipType()
      });
    }
    await XmlHelper.writeXmlToArchive(this.targetArchive, targetRelFile, relXml);

    // 2. Modify the copied slide content XML
    await this.updateHyperlinkInSlide();
  }

  // Helper method to create a hyperlink in a shape
  static async addHyperlinkToShape(
    archive: IArchive,
    slidePath: string,
    slideRelsPath: string,
    shapeId: string,
    hyperlinkTarget: string | number
  ): Promise<string> {
    const slideXml = await XmlHelper.getXmlFromArchive(archive, slidePath);
    
    // Find the shape by ID or name
    const shape = XmlHelper.isElementCreationId(shapeId)
      ? XmlHelper.findByCreationId(slideXml, shapeId)
      : XmlHelper.findByName(slideXml, shapeId);
    
    if (!shape) {
      throw new Error(`Shape with ID/name "${shapeId}" not found`);
    }

    // Create a new relationship ID
    const relId = await XmlHelper.getNextRelId(archive, slideRelsPath);
    
    // Add the hyperlink relationship to the slide relationships
    const relXml = await XmlHelper.getXmlFromArchive(archive, slideRelsPath);
    const newRel = relXml.createElement('Relationship');
    newRel.setAttribute('Id', relId);

    // Improved internal link detection
    const isInternalLink = typeof hyperlinkTarget === 'number' || 
      (typeof hyperlinkTarget === 'string' && /^\d+$/.test(hyperlinkTarget));
    
    if (isInternalLink) {
      // Enhanced internal slide link handling
      const slideNumber = typeof hyperlinkTarget === 'number' ? 
        hyperlinkTarget : 
        parseInt(hyperlinkTarget, 10);
      newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
      newRel.setAttribute('Target', `../slides/slide${slideNumber}.xml`);
    } else {
      newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
      newRel.setAttribute('Target', hyperlinkTarget.toString());
      newRel.setAttribute('TargetMode', 'External');
    }
    
    relXml.documentElement.appendChild(newRel);
    await XmlHelper.writeXmlToArchive(archive, slideRelsPath, relXml);

    // Add the hyperlink to the shape
    const txBody = shape.getElementsByTagName('p:txBody')[0];
    if (txBody) {
      const paragraphs = txBody.getElementsByTagName('a:p');
      
      if (paragraphs.length > 0) {
        const paragraph = paragraphs[0];
        const runs = paragraph.getElementsByTagName('a:r');
        
        if (runs.length > 0) {
          const run = runs[0];
          const rPr = run.getElementsByTagName('a:rPr')[0];
          
          if (rPr) {
            const hlinkClick = slideXml.createElement('a:hlinkClick');
            hlinkClick.setAttribute('r:id', relId);
            if (isInternalLink) {
              hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
            }
            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            rPr.appendChild(hlinkClick);
          }
        }
      }
    }

    await XmlHelper.writeXmlToArchive(archive, slidePath, slideXml);
    
    return relId;
  }
} 