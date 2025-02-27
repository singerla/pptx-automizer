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
      if (!element) {
        console.log('SetHyperlinkTarget: Missing element');
        return;
      }
      
      try {
        console.log('SetHyperlinkTarget: Starting to update hyperlink target to:', target);
        console.log('Element:', element.nodeName, element.getAttribute('name') || 'no name');
        
        // Find all hyperlink elements in the shape
        const hlinkElements = element.getElementsByTagName('a:hlinkClick');
        console.log(`SetHyperlinkTarget found ${hlinkElements.length} hyperlinks to update with target: ${target}`);
        
        if (hlinkElements.length === 0) {
          // If no hyperlinks found, add one instead
          console.log('No existing hyperlinks found. Adding a new one.');
          ModifyHyperlinkHelper.addHyperlink(target, isExternal)(element, relation);
          return;
        }
        
        // Get the document URI from the element to determine the slide number
        const documentURI = element.ownerDocument?.documentURI;
        if (!documentURI) {
          console.error('Cannot determine document URI from element');
          return;
        }
        
        // Extract the slide number from the document URI
        // Example: ppt/slides/slide1.xml
        const slideMatch = documentURI.match(/slide(\d+)\.xml/);
        if (!slideMatch) {
          console.error('Cannot determine slide number from document URI:', documentURI);
          return;
        }
        
        const slideNumber = slideMatch[1];
        const slideRelsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
        
        console.log(`Working with slide ${slideNumber}, relationships file: ${slideRelsPath}`);
        
        // Process all hyperlinks found in the element
        for (let i = 0; i < hlinkElements.length; i++) {
          const hlinkElement = hlinkElements[i];
          const rId = hlinkElement.getAttribute('r:id');
          
          console.log(`Processing hyperlink with r:id=${rId}`);
          
          if (rId) {
            // Create a new relationship ID
            const newRelId = `rId${Date.now().toString().substring(8)}`;
            
            // Update the hyperlink element to use the new relationship ID
            hlinkElement.setAttribute('r:id', newRelId);
            console.log(`Updated hyperlink element to use new relationship ID: ${newRelId}`);
            
            // Get the slide relationships XML
            const slideRelXml = await XmlHelper.getXmlFromArchive(
              // We don't have direct access to the archive, so we'll need to use a different approach
              // This is a workaround and might not work in all cases
              // In a real implementation, we would need to pass the archive to this function
              global.__PPTX_CURRENT_ARCHIVE__,
              slideRelsPath
            );
            
            if (!slideRelXml) {
              console.error(`Could not get relationships XML for slide ${slideNumber}`);
              return;
            }
            
            // Create a new relationship element
            const newRel = slideRelXml.createElement('Relationship');
            newRel.setAttribute('Id', newRelId);
            newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
            newRel.setAttribute('Target', target);
            
            if (isExternal) {
              newRel.setAttribute('TargetMode', 'External');
            }
            
            // Add the new relationship to the relationships XML
            slideRelXml.documentElement.appendChild(newRel);
            console.log(`Added new relationship with Id=${newRelId} and target=${target}`);
            
            // Remove the old relationship if it exists
            const relationships = slideRelXml.getElementsByTagName('Relationship');
            for (let j = 0; j < relationships.length; j++) {
              const rel = relationships[j];
              const relId = rel.getAttribute('Id');
              
              if (relId === rId) {
                console.log(`Found old relationship with Id=${rId}, removing it`);
                rel.parentNode?.removeChild(rel);
                break;
              }
            }
            
            // Write the updated relationships XML back to the archive
            await XmlHelper.writeXmlToArchive(
              // We don't have direct access to the archive, so we'll need to use a different approach
              // This is a workaround and might not work in all cases
              // In a real implementation, we would need to pass the archive to this function
              global.__PPTX_CURRENT_ARCHIVE__,
              slideRelsPath,
              slideRelXml
            );
            
            // Track the relationship for content integrity
            try {
              contentTracker.trackRelation(slideRelsPath, {
                Id: newRelId,
                Target: target,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
              });
              console.log('Tracked the new relationship in content tracker');
            } catch (e) {
              console.error('Error tracking relation:', e);
            }
          }
        }
        
        console.log('SetHyperlinkTarget: Successfully completed');
      } catch (error) {
        console.error('Error in SetHyperlinkTarget:', error);
      }
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
        console.log('RemoveHyperlink: Starting to remove hyperlinks for element:', element.nodeName);
        console.log('Element name:', element.getAttribute('name') || 'no name');
        console.log('Element owner document:', element.ownerDocument ? 'Available' : 'Not available');
        
        // Find all hyperlink elements in the shape
        const hlinkElements = element.getElementsByTagName('a:hlinkClick');
        console.log(`RemoveHyperlink found ${hlinkElements.length} hyperlinks to remove`);
        
        if (hlinkElements.length === 0) {
          console.log('No hyperlinks found to remove');
          return;
        }
        
        // Collect all relationship IDs to remove
        const rIds: string[] = [];
        
        // First, collect all r:id values
        for (let i = 0; i < hlinkElements.length; i++) {
          const rId = hlinkElements[i].getAttribute('r:id');
          if (rId) {
            console.log(`Found hyperlink with r:id=${rId}`);
            rIds.push(rId);
          }
        }
        
        // Remove all hyperlink elements from parent nodes
        // We need to iterate backwards because the collection will change as we remove items
        for (let i = hlinkElements.length - 1; i >= 0; i--) {
          const hlink = hlinkElements[i];
          if (hlink.parentNode) {
            console.log(`Removing hyperlink element at index ${i}`);
            hlink.parentNode.removeChild(hlink);
          }
        }
        
        // Remove the relationships if relation XML is provided directly
        if (relation && rIds.length > 0) {
          console.log(`Removing hyperlinks from provided relation XML. Relation type:`, relation.nodeName);
          const relationships = relation.getElementsByTagName('Relationship');
          console.log(`Found ${relationships.length} relationships in total`);
          let removedCount = 0;
          
          // Iterate backwards to avoid collection modification issues
          for (let i = relationships.length - 1; i >= 0; i--) {
            const rel = relationships[i];
            const relId = rel.getAttribute('Id');
            console.log(`Checking relationship ${i} with Id=${relId}`);
            
            if (relId && rIds.includes(relId)) {
              console.log(`Removing relationship with Id=${relId}`);
              if (rel.parentNode) {
                rel.parentNode.removeChild(rel);
                removedCount++;
              } else {
                console.log(`Cannot remove relationship ${relId} - no parent node`);
              }
            }
          }
          
          console.log(`Removed ${removedCount} hyperlink relationships`);
        } else {
          if (!relation) {
            console.log('No relation XML provided - cannot remove relationship entries');
            
            // If no relation XML is provided but we have document URI, try to remove from file
            // This is a fallback to ensure hyperlinks are removed from the relationship file
            if (element.ownerDocument?.documentURI) {
              const documentURI = element.ownerDocument.documentURI;
              console.log('Trying to remove relationships from document URI:', documentURI);
              
              // Extract the slide number from the document URI
              // Example: ppt/slides/slide1.xml
              const slideMatch = documentURI.match(/slide(\d+)\.xml/);
              if (slideMatch) {
                const slideNumber = slideMatch[1];
                const slideRelsPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
                
                console.log(`Found slide number: ${slideNumber}, trying to get relationships from: ${slideRelsPath}`);
                
                try {
                  // We need to access the global archive to modify the relations file
                  if (global.__PPTX_CURRENT_ARCHIVE__) {
                    const archive = global.__PPTX_CURRENT_ARCHIVE__;
                    const relXml = await XmlHelper.getXmlFromArchive(archive, slideRelsPath);
                    
                    if (relXml) {
                      console.log('Successfully got relationship XML, removing hyperlink relationships');
                      const relationships = relXml.getElementsByTagName('Relationship');
                      let removedCount = 0;
                      
                      // Iterate backwards to avoid collection modification issues
                      for (let i = relationships.length - 1; i >= 0; i--) {
                        const rel = relationships[i];
                        const relId = rel.getAttribute('Id');
                        const type = rel.getAttribute('Type');
                        
                        if (relId && rIds.includes(relId) && 
                            type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
                          console.log(`Removing relationship with Id=${relId}`);
                          if (rel.parentNode) {
                            rel.parentNode.removeChild(rel);
                            removedCount++;
                          }
                        }
                      }
                      
                      console.log(`Removed ${removedCount} hyperlink relationships from file`);
                      
                      // Write the updated relationships XML back to the archive
                      await XmlHelper.writeXmlToArchive(archive, slideRelsPath, relXml);
                    } else {
                      console.log('Could not get relationship XML from archive');
                    }
                  } else {
                    console.log('No archive available in global context');
                  }
                } catch (e) {
                  console.error('Error removing relationships from file:', e);
                }
              }
            }
          } else if (rIds.length === 0) {
            console.log('No rIds collected - nothing to remove from relationships');
          }
        }

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
   * @param target The target URL for the hyperlink
   * @param isExternal Whether the hyperlink is external (true) or internal (false)
   * @returns A callback function that adds a hyperlink
   */
  static addHyperlink = 
    (target: string, isExternal: boolean = true): ShapeModificationCallback =>
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
        console.log(`Adding hyperlink with new rel ID: ${newRelId}, target: ${target}`);
        
        // Create the relationship
        const newRel = relation.ownerDocument.createElement('Relationship');
        newRel.setAttribute('Id', newRelId);
        newRel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
        newRel.setAttribute('Target', target);
        
        if (isExternal) {
          newRel.setAttribute('TargetMode', 'External');
        }
        
        // Add the relationship to the document
        relation.appendChild(newRel);
        console.log('Added new relationship to the document');
        
        // Track the relationship
        try {
          contentTracker.trackRelation(relation.ownerDocument.documentURI || '', {
            Id: newRelId,
            Target: target,
            Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
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
  
  /**
   * Process pending hyperlink modifications stored on an element
   * This is a placeholder implementation to maintain API compatibility
   */
  static async processPendingHyperlinkModifications(
    element: XmlElement,
    targetArchive: any,
    targetSlideFile: string,
    targetSlideRelFile: string
  ): Promise<void> {
    // This is a simplified version that's just here to maintain compatibility
    // The actual hyperlink processing is done via the callbacks
    console.log('processPendingHyperlinkModifications called - using simplified implementation');
    
    // No-op implementation as we're using the callback pattern now
    return Promise.resolve();
  }
} 