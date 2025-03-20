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
    hyperlinkTarget: string = '',
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
    await this.updateHyperlinkInSlide();

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
    await this.removeFromSlideRels();

    return this;
  }

  async prepare(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    hyperlinks?: Target[]
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

    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
  }

  // Add a method to update the hyperlink in the slide XML
  private async updateHyperlinkInSlide(): Promise<void> {
    // Get the slide XML
    const slideXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      this.targetSlideFile,
    );

    // When copying a full slide with hyperlinks, we may not have a targetElement
    // but we still need to update hyperlinks in the slide XML
    if (this.sourceRid && this.createdRid) {
      // Find all a:hlinkClick elements in the entire slide
      const allHyperlinkElements = slideXml.getElementsByTagName('a:hlinkClick');
      
      for (let i = 0; i < allHyperlinkElements.length; i++) {
        const hlinkClick = allHyperlinkElements[i];
        
        // Check if this element references our source relationship ID
        if (hlinkClick.getAttribute('r:id') === this.sourceRid) {
          // Update to use the new relationship ID
          hlinkClick.setAttribute('r:id', this.createdRid);
          
          // Ensure the xmlns:r attribute is set
          hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        }
      }
    }
    
    // For element-specific hyperlinks (when targetElement is set)
    if (this.targetElement) {
      // Find all text runs in the element
      const runs = this.targetElement.getElementsByTagName('a:r');
      
      for (let i = 0; i < runs.length; i++) {
        const run = runs[i];
        const rPr = run.getElementsByTagName('a:rPr')[0];
        
        if (rPr) {
          // Find hyperlink elements
          const hlinkClicks = rPr.getElementsByTagName('a:hlinkClick');
          
          for (let j = 0; j < hlinkClicks.length; j++) {
            const hlinkClick = hlinkClicks[j];
            
            // Update the r:id attribute to use the created relationship ID
            hlinkClick.setAttribute('r:id', this.createdRid);
            
            // Ensure the xmlns:r attribute is set
            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
          }
        }
      }
    }

    // Write the updated XML back to the archive
    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      this.targetSlideFile,
      slideXml,
    );
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

    for (let i = 0; i < relationships.length; i++) {
      if (relationships[i].getAttribute('Id') === this.createdRid) {
        relXml.documentElement.removeChild(relationships[i]);
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
    const hyperlinkType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

    return XmlHelper.getRelationshipItems(
      archive,
      relsPath,
      (element: XmlElement, rels: Target[]) => {
        const type = element.getAttribute('Type');

        if (type === hyperlinkType) {
          rels.push({
            rId: element.getAttribute('Id'),
            type: element.getAttribute('Type'),
            file: element.getAttribute('Target'),
            filename: element.getAttribute('Target'),
            element: element,
            isExternal: element.getAttribute('TargetMode') === 'External',
          } as Target);
        }
      }
    );
  }

  async modifyOnAddedSlide(
    targetTemplate: RootPresTemplate,
    targetSlideNumber: number,
    hyperlinks?: Target[]
  ): Promise<void> {
    await this.prepare(targetTemplate, targetSlideNumber, hyperlinks);
    
    // Enhanced internal link type detection
    if (this.target && this.target.file) {
      this.hyperlinkTarget = this.target.file;
      this.hyperlinkType = this.target.file.includes('/slides/slide') ? 'internal' : 'external';
    }
    
    // Update the relationship in the slide's relationships file
    const targetRelFile = this.targetSlideRelFile;
    const relXml = await XmlHelper.getXmlFromArchive(
      this.targetArchive,
      targetRelFile,
    );
    
    // Find the relationship with the source rId
    const relationships = relXml.getElementsByTagName('Relationship');
    let relationshipUpdated = false;
    
    // First, check if we need to update an existing relationship
    if (this.sourceRid) {
      for (let i = 0; i < relationships.length; i++) {
        const relationship = relationships[i];
        if (relationship.getAttribute('Id') === this.sourceRid) {
          // Update the existing relationship
          relationship.setAttribute('Id', this.createdRid);
          
          // Set the relationship type
          relationship.setAttribute('Type', this.getRelationshipType());
          
          // For external links, preserve the original target URL and ensure TargetMode is External
          if (this.hyperlinkType === 'external') {
            relationship.setAttribute('Target', this.hyperlinkTarget);
            relationship.setAttribute('TargetMode', 'External');
          } else {
            // For internal links, set the target appropriately
            relationship.setAttribute('Target', this.getRelationshipTarget());
            if (relationship.hasAttribute('TargetMode')) {
              relationship.removeAttribute('TargetMode');
            }
          }
          
          relationshipUpdated = true;
          break;
        }
      }
    }
    
    // If the relationship wasn't found or updated, create a new one
    if (!relationshipUpdated) {
      const newRel = relXml.createElement('Relationship');
      newRel.setAttribute('Id', this.createdRid);
      newRel.setAttribute('Type', this.getRelationshipType());
      
      if (this.hyperlinkType === 'external') {
        // For external links, use the original URL and set TargetMode
        newRel.setAttribute('Target', this.hyperlinkTarget);
        newRel.setAttribute('TargetMode', 'External');
      } else {
        // For internal links
        newRel.setAttribute('Target', this.getRelationshipTarget());
      }
      
      relXml.documentElement.appendChild(newRel);
    }
    
    // Write the updated XML back to the archive
    await XmlHelper.writeXmlToArchive(
      this.targetArchive,
      targetRelFile,
      relXml,
    );
    
    // Track the relationship for content integrity
    contentTracker.trackRelation(targetRelFile, {
      Id: this.createdRid,
      Target: this.hyperlinkType === 'external' ? this.hyperlinkTarget : this.getRelationshipTarget(),
      Type: this.getRelationshipType(),
    });
    
    // Now update the hyperlink reference in the slide XML
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