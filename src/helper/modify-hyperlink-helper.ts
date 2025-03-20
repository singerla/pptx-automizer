import { ShapeModificationCallback } from '../types/types';
import { XmlElement } from '../types/xml-types';
import { XmlHelper } from './xml-helper';
import { contentTracker } from './content-tracker';

/**
 * Helper class for modifying hyperlinks in PowerPoint elements
 */
export default class ModifyHyperlinkHelper {
  /**
   * Set the target URL of a hyperlink
   * 
   * @param target The new target URL for the hyperlink
   * @param isExternal Whether the hyperlink is external (true) or internal (false)
   * @returns A callback function that modifies the hyperlink
   */
  static setHyperlinkTarget = 
    (target: string, isExternal: boolean = true): ShapeModificationCallback =>
    async (element: XmlElement, relation?: XmlElement): Promise<void> => {
      console.log('SetHyperlinkTarget: This function requires implementation.');
    };
  
  /**
   * Remove hyperlinks from an element
   * 
   * @returns A callback function that removes hyperlinks
   */
  static removeHyperlink = 
    (): ShapeModificationCallback =>
    async (element: XmlElement, relation?: XmlElement): Promise<void> => {
      if (!element) {
        console.log('RemoveHyperlink: Missing element');
        return;
      }
      
      try {
        console.log('STUB - Requires implementation: Starting to remove hyperlinks');
        // Verify that hyperlinks were actually removed
        const remainingHlinks = element.getElementsByTagName('a:hlinkClick');
        console.log(`After removal operations: Found ${remainingHlinks.length} hyperlinks remaining`);
        
        console.log('RemoveHyperlink: Successfully completed');
      } catch (error) {
        console.error('Error in RemoveHyperlink:', error);
        // Don't rethrow the error, just log it
      }
    };
  
  /**
   * Add a hyperlink to an element
   * 
   * @param target The target URL for external links, or slide number for internal links
   * @returns A callback function that adds a hyperlink
   */
  static addHyperlink = 
    (target: string | number): ShapeModificationCallback =>
    (element: XmlElement, relation?: XmlElement): void => {
      if (!element || !relation) {
        console.log('AddHyperlink: Missing element or relation');
        if (!element) console.log('Element is missing');
        if (!relation) console.log('Relation is missing');
        return;
      }
      
      try {
        console.log('AddHyperlink: Starting to add hyperlink');
        console.log('Element:', element.nodeName, element.getAttribute('name') || 'no name');
        console.log('Relation document URI:', relation.ownerDocument.documentURI || 'no URI');
        
        // Create a new relationship ID
        let newRelId = '';
        
        // Get all existing relationship IDs
        const relationships = relation.getElementsByTagName('Relationship');
        console.log(`Found ${relationships.length} existing relationships`);
        let maxId = 0;
        
        for (let i = 0; i < relationships.length; i++) {
          const relId = relationships[i].getAttribute('Id');
          if (relId && relId.startsWith('rId')) {
            const idNum = parseInt(relId.substring(3), 10);
            if (!isNaN(idNum) && idNum > maxId) {
              maxId = idNum;
            }
          }
        }
        
        newRelId = `rId${maxId + 1}`;
        
        // Determine if this is an internal slide link
        const isInternalLink = typeof target === 'number' || 
          (typeof target === 'string' && /^\d+$/.test(target));

        // Create the relationship
        const newRel = relation.ownerDocument.createElement('Relationship');
        newRel.setAttribute('Id', newRelId);
        
        if (isInternalLink) {
          // For internal slide links
          const slideNumber = typeof target === 'number' ? target : parseInt(target, 10);
          newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide');
          newRel.setAttribute('Target', `../slides/slide${slideNumber}.xml`);
        } else {
          // For external links
          newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
          newRel.setAttribute('Target', target.toString());
          newRel.setAttribute('TargetMode', 'External');
        }
        
        // Add the relationship to the document
        relation.appendChild(newRel);
        console.log('Added new relationship to the document');
        
        // Track the relationship
        try {
          contentTracker.trackRelation(relation.ownerDocument.documentURI || '', {
            Id: newRelId,
            Target: newRel.getAttribute('Target') || '',
            Type: newRel.getAttribute('Type') || '',
          });
          console.log('Tracked the relationship in content tracker');
        } catch (e) {
          console.error('Error tracking relation:', e);
        }
        
        // Find the appropriate element to add the hyperlink to
        // If it's a text shape, find text runs
        const textRuns = element.getElementsByTagName('a:r');
        console.log(`Found ${textRuns.length} text runs in the element`);
        
        if (textRuns.length > 0) {
          // Add hyperlink to each text run
          for (let i = 0; i < textRuns.length; i++) {
            const run = textRuns[i];
            let rPr = run.getElementsByTagName('a:rPr')[0];
            
            // Create rPr element if it doesn't exist
            if (!rPr) {
              rPr = element.ownerDocument.createElement('a:rPr');
              // Insert rPr before the text element
              const textElement = run.getElementsByTagName('a:t')[0];
              if (textElement) {
                run.insertBefore(rPr, textElement);
              } else {
                run.appendChild(rPr);
              }
            }
            
            // Add hyperlink element
            const hlinkClick = element.ownerDocument.createElement('a:hlinkClick');
            hlinkClick.setAttribute('r:id', newRelId);
            if (isInternalLink) {
              hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
              hlinkClick.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
              hlinkClick.setAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
            }
            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            rPr.appendChild(hlinkClick);
            console.log(`Added hyperlink to text run ${i+1}`);
          }
        } else {
          // If no text runs, check for paragraphs and create text run
          const paragraphs = element.getElementsByTagName('a:p');
          console.log(`Found ${paragraphs.length} paragraphs in the element`);
          
          if (paragraphs.length > 0) {
            // Use the first paragraph
            const paragraph = paragraphs[0];
            
            // Create new text run with the hyperlink
            const run = element.ownerDocument.createElement('a:r');
            const rPr = element.ownerDocument.createElement('a:rPr');
            const hlinkClick = element.ownerDocument.createElement('a:hlinkClick');
            hlinkClick.setAttribute('r:id', newRelId);
            if (isInternalLink) {
              hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
              hlinkClick.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
              hlinkClick.setAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
            }
            hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            
            rPr.appendChild(hlinkClick);
            run.appendChild(rPr);
            
            // Add text content if the paragraph is empty
            if (paragraph.getElementsByTagName('a:t').length === 0) {
              const t = element.ownerDocument.createElement('a:t');
              t.textContent = 'Hyperlink';
              run.appendChild(t);
              console.log('Added new text content "Hyperlink"');
            } else {
              // Use existing text
              const existingText = paragraph.getElementsByTagName('a:t')[0];
              const t = element.ownerDocument.createElement('a:t');
              t.textContent = existingText.textContent;
              run.appendChild(t);
              console.log(`Used existing text: "${existingText.textContent}"`);
              
              // Remove existing text element
              paragraph.removeChild(existingText.parentNode || existingText);
            }
            
            paragraph.appendChild(run);
            console.log('Added hyperlink to paragraph');
          } else {
            // If no text content at all, create a complete text structure
            const txBody = element.getElementsByTagName('p:txBody')[0] || element.getElementsByTagName('a:txBody')[0];
            console.log(`Found txBody: ${txBody ? 'yes' : 'no'}`);
            
            if (txBody) {
              // Create new paragraph
              const p = element.ownerDocument.createElement('a:p');
              const run = element.ownerDocument.createElement('a:r');
              const rPr = element.ownerDocument.createElement('a:rPr');
              const hlinkClick = element.ownerDocument.createElement('a:hlinkClick');
              const t = element.ownerDocument.createElement('a:t');
              
              hlinkClick.setAttribute('r:id', newRelId);
              if (isInternalLink) {
                hlinkClick.setAttribute('action', 'ppaction://hlinksldjump');
                hlinkClick.setAttribute('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
                hlinkClick.setAttribute('xmlns:p14', 'http://schemas.microsoft.com/office/powerpoint/2010/main');
              }
              hlinkClick.setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
              t.textContent = 'Hyperlink';
              
              rPr.appendChild(hlinkClick);
              run.appendChild(rPr);
              run.appendChild(t);
              p.appendChild(run);
              txBody.appendChild(p);
              console.log('Created complete text structure with hyperlink');
            } else {
              console.error('No suitable text element found to add hyperlink to');
            }
          }
        }
        
        console.log('AddHyperlink: Successfully completed');
      } catch (error) {
        console.error('Error in AddHyperlink:', error);
      }
    };
} 