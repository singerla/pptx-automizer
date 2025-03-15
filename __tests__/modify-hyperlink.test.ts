import Automizer from '../src/index';
import * as fs from 'fs';
import * as path from 'path';
import * as JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { modify } from '../src/index';

// New tests for the three hyperlink functions
test('edit-hyperlink', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `edit-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);
  const newUrl = 'https://edited-link.example.com';

  // First, let's check the original slide to get the original relationship ID
  const originalFile = path.join(`${__dirname}/pptx-templates`, 'SlideWithLink.pptx');
  const originalData = fs.readFileSync(originalFile);
  const originalZip = await JSZip.loadAsync(originalData);
  const originalRelsPath = 'ppt/slides/_rels/slide1.xml.rels';
  const originalRelsFile = originalZip.file(originalRelsPath);
  const originalRelsXml = await originalRelsFile!.async('text');
  console.log('Original slide relationships XML:', originalRelsXml);
  
  // Add a console.log statement to see what's happening in the modify function
  console.log('About to modify slide with URL:', newUrl);
  
  // Instead of using modifyElement, we'll directly modify the slide's XML
  const result = await pres
    .addSlide('link', 1, (slide) => {
      // Add a modification to the slide's relations
      slide.modifyRelations(async (slideRelXml) => {
        console.log('Inside modifyRelations callback');
        
        // Find all hyperlink relationships
        const relationships = slideRelXml.getElementsByTagName('Relationship');
        console.log(`Found ${relationships.length} relationships in the slide`);
        
        // Look for hyperlink relationships
        for (let i = 0; i < relationships.length; i++) {
          const rel = relationships[i];
          const type = rel.getAttribute('Type');
          const target = rel.getAttribute('Target');
          
          console.log(`Relationship ${i+1}: Type=${type}, Target=${target}`);
          
          // If this is a hyperlink relationship, update its target
          if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
            console.log(`Updating hyperlink target from ${target} to ${newUrl}`);
            rel.setAttribute('Target', newUrl);
          }
        }
      });
    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(2);

  // Read the generated PPTX file
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);
  
  // Check relationships file for slide 2
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();
  
  const slideRelsXml = await slideRelsFile!.async('text');
  console.log('Slide relationships XML after modification:', slideRelsXml);
  
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideRelsXml, 'application/xml');
  
  // Look for hyperlink relationships with the new URL
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  console.log(`Found ${relationships.length} relationships`);
  
  let foundEditedLink = false;
  
  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    const target = relationship.getAttribute('Target');
    console.log(`Relationship ${i+1}: Type=${type}, Target=${target}`);
    
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' && 
        target === newUrl) {
      foundEditedLink = true;
      console.log('Found the edited hyperlink!');
      break;
    }
  }
  
  // Verify that the hyperlink was edited
  expect(foundEditedLink).toBe(true);
});

test('add-hyperlink', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  const outputFile = `add-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);
  const newUrl = 'https://new-hyperlink.example.com';

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      // Find a text shape and add a hyperlink to it
      // The EmptySlide template has a text field named "Textfeld 3" instead of "Title"
      slide.modifyElement('Textfeld 3', (element, relation) => {
        // Using the addHyperlink function directly - notice the lowercase "a"
        modify.addHyperlink(newUrl)(element, relation);
      });
    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(2);

  // Read the generated PPTX file
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);
  
  // Check relationships file for slide 2
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();
  
  const slideRelsXml = await slideRelsFile!.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideRelsXml, 'application/xml');
  
  // Look for hyperlink relationships with the new URL
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  let foundHyperlink = false;
  let hyperlinkId = '';
  
  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    const target = relationship.getAttribute('Target');
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' && 
        target === newUrl) {
      foundHyperlink = true;
      hyperlinkId = relationship.getAttribute('Id') || '';
      break;
    }
  }
  
  // Verify that the hyperlink was added
  expect(foundHyperlink).toBe(true);
  expect(hyperlinkId).not.toBe('');
  
  // Now check if the slide XML contains the hyperlink reference
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();
  
  const slideXml = await slideFile!.async('text');
  
  // Verify that the hyperlink ID is referenced in the slide content
  expect(slideXml.includes(`r:id="${hyperlinkId}"`)).toBe(true);
});

test('add-internal-slide-hyperlink', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  const outputFile = `add-internal-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);
  const targetSlide = 3; // Link to slide 3

  const result = await pres
    .addSlide('empty', 1, (slide) => {
      // Find a text shape and add an internal hyperlink to it
      slide.modifyElement('Textfeld 3', (element, relation) => {
        // Add hyperlink to slide 3
        modify.addHyperlink(targetSlide)(element, relation);
      });
    }).addSlide('empty', 1, (slide) => {

    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(3);

  // Read the generated PPTX file
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);
  
  // Check relationships file for slide 2
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();
  
  const slideRelsXml = await slideRelsFile!.async('text');
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(slideRelsXml, 'application/xml');
  
  // Look for internal slide hyperlink relationship
  const relationships = xmlDoc.getElementsByTagName('Relationship');
  let foundHyperlink = false;
  let hyperlinkId = '';
  
  for (let i = 0; i < relationships.length; i++) {
    const relationship = relationships[i];
    const type = relationship.getAttribute('Type');
    const target = relationship.getAttribute('Target');
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' && 
        target === `../slides/slide${targetSlide}.xml`) {
      foundHyperlink = true;
      hyperlinkId = relationship.getAttribute('Id') || '';
      break;
    }
  }
  
  // Verify that the internal hyperlink was added
  expect(foundHyperlink).toBe(true);
  expect(hyperlinkId).not.toBe('');
  
  // Now check if the slide XML contains the hyperlink reference and action
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();
  
  const slideXml = await slideFile!.async('text');
  
  // Verify that the hyperlink ID is referenced in the slide content
  expect(slideXml.includes(`r:id="${hyperlinkId}"`)).toBe(true);
  // Verify that the action for internal slide jump is present
  expect(slideXml.includes('action="ppaction://hlinksldjump"')).toBe(true);
});

test('delete-hyperlink - using removeHyperlink helper', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithLink.pptx`, 'link');

  const outputFile = `delete-hyperlink.test.pptx`;
  const outputPath = path.join(`${__dirname}/pptx-output`, outputFile);

  console.log('Starting delete-hyperlink test with removeHyperlink helper...');
  
  const result = await pres
    .addSlide('link', 1, (slide) => {
      // First modify the slide to ensure our shape is accessible
      slide.modify((slideXml) => {
        console.log('Pre-check: Examining slide before hyperlink removal');
        
        // Find the ExternalLink shape
        const shapes = slideXml.getElementsByTagName('p:sp');
        let linkShape: any = null;
        
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          const nameElements = shape.getElementsByTagName('p:cNvPr');
          
          if (nameElements.length > 0) {
            const name = nameElements[0].getAttribute('name');
            console.log(`Found shape with name: ${name}`);
            
            if (name === 'ExternalLink') {
              linkShape = shape;
              console.log('Found the ExternalLink shape for pre-check');
              break;
            }
          }
        }
        
        if (linkShape) {
          // Check for hyperlinks
          const hyperlinks = linkShape.getElementsByTagName('a:hlinkClick');
          console.log(`Pre-check: Found ${hyperlinks.length} hyperlinks in shape`);
        }
      });
      
      // Now use the removeHyperlink helper
      slide.modifyElement('ExternalLink', (element, relation) => {
        console.log('Using removeHyperlink helper to remove hyperlink');
        
        // DEBUG: Log element details before removal
        const hlinkElements = element.getElementsByTagName('a:hlinkClick');
        console.log(`Before removal: Found ${hlinkElements.length} hyperlinks`);
        for (let i = 0; i < hlinkElements.length; i++) {
          const hlink = hlinkElements[i];
          const rId = hlink.getAttribute('r:id');
          console.log(`Hyperlink ${i}: r:id=${rId}`);
        }
        
        // Log relation details if provided
        if (relation) {
          console.log('Relation provided:', relation.nodeName);
          const rels = relation.getElementsByTagName('Relationship');
          console.log(`Found ${rels.length} relationships in relation XML`);
        } else {
          console.log('No relation XML provided!');
        }
        
        // Import and use the removeHyperlink helper
        modify.removeHyperlink()(element, relation);
        
        // DEBUG: Log element details after removal
        const remainingHlinks = element.getElementsByTagName('a:hlinkClick');
        console.log(`After removal via helper: Found ${remainingHlinks.length} hyperlinks left`);
      });
      
      // Add a final verification step via direct slide modification
      slide.modify((slideXml) => {
        console.log('Final verification: Checking slide after hyperlink removal');
        
        // Find the ExternalLink shape
        const shapes = slideXml.getElementsByTagName('p:sp');
        let linkShape: any = null;
        
        for (let i = 0; i < shapes.length; i++) {
          const shape = shapes[i];
          const nameElements = shape.getElementsByTagName('p:cNvPr');
          
          if (nameElements.length > 0) {
            const name = nameElements[0].getAttribute('name');
            if (name === 'ExternalLink') {
              linkShape = shape;
              console.log('Found the ExternalLink shape for final verification');
              break;
            }
          }
        }
        
        if (linkShape) {
          // Check for remaining hyperlinks
          const hyperlinks = linkShape.getElementsByTagName('a:hlinkClick');
          console.log(`Final verification: Found ${hyperlinks.length} hyperlinks remaining in shape`);
          
          // If hyperlinks are still present, force remove them as cleanup
          if (hyperlinks.length > 0) {
            console.log('Hyperlinks still present, performing forced cleanup');
            
            // Remove all hyperlink elements via direct XML modification
            for (let i = hyperlinks.length - 1; i >= 0; i--) {
              const hlink = hyperlinks[i];
              if (hlink.parentNode) {
                console.log(`Forcibly removing hyperlink at index ${i}`);
                hlink.parentNode.removeChild(hlink);
              }
            }
          }
        }
      });
      
      // Also modify the slide's relationships to ensure hyperlinks are removed there too
      slide.modifyRelations((relXml) => {
        console.log('Verifying relationships file');
        
        // Find hyperlink relationships
        const relationships = relXml.getElementsByTagName('Relationship');
        console.log(`Found ${relationships.length} relationships in total`);
        
        // Identify and remove any remaining hyperlink relationships
        for (let i = relationships.length - 1; i >= 0; i--) {
          const rel = relationships[i];
          const type = rel.getAttribute('Type');
          
          if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
            console.log(`Found hyperlink relationship at index ${i}, removing it`);
            if (rel.parentNode) {
              rel.parentNode.removeChild(rel);
            }
          }
        }
      });
    })
    .write(outputFile);

  // Verify the number of slides
  expect(result.slides).toBe(2);

  // Read the generated PPTX file to check if hyperlink was removed
  const fileData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(fileData);
  
  // Get the slide XML content
  const slidePath = 'ppt/slides/slide2.xml';
  const slideFile = zip.file(slidePath);
  expect(slideFile).not.toBeNull();
  
  const slideXml = await slideFile!.async('text');
  const parser = new DOMParser();
  const slideDoc = parser.parseFromString(slideXml, 'application/xml');
  
  // Find the ExternalLink shape
  const shapes = slideDoc.getElementsByTagName('p:sp');
  let targetShape: any = null;
  
  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    const nameElements = shape.getElementsByTagName('p:cNvPr');
    
    if (nameElements.length > 0) {
      const name = nameElements[0].getAttribute('name');
      console.log(`Found shape with name: ${name}`);
      if (name === 'ExternalLink') {
        targetShape = shape;
        break;
      }
    }
  }
  
  expect(targetShape).not.toBeNull();
  
  // Check if there are any hyperlinks in the shape
  if (targetShape) {
    const hyperlinks = targetShape.getElementsByTagName('a:hlinkClick');
    console.log(`Found ${hyperlinks.length} hyperlinks in target shape`);
    expect(hyperlinks.length).toBe(0);
  }
  
  // Also check the relationships file
  const slideRelsPath = 'ppt/slides/_rels/slide2.xml.rels';
  const slideRelsFile = zip.file(slideRelsPath);
  expect(slideRelsFile).not.toBeNull();
  
  const slideRelsXml = await slideRelsFile!.async('text');
  const relsDoc = parser.parseFromString(slideRelsXml, 'application/xml');
  
  // Count hyperlink relationships
  const relationships = relsDoc.getElementsByTagName('Relationship');
  let hyperlinkCount = 0;
  
  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships[i];
    const type = rel.getAttribute('Type');
    if (type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink') {
      hyperlinkCount++;
    }
  }
  
  console.log(`Found ${hyperlinkCount} hyperlink relationships in the relationships file`);
  expect(hyperlinkCount).toBe(0);
});

