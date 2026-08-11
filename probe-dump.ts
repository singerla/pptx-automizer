import JSZip from 'jszip';
import fs from 'fs';

(async () => {
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/__tests__/pptx-output/probe-html.pptx`),
  );
  const xml = await archive.file('ppt/slides/slide2.xml').async('string');
  const shapes = xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
  const target = shapes.find((shape) => shape.includes('name="setText"'));
  const body = target.match(/<p:txBody>[\s\S]*?<\/p:txBody>/);
  console.log(body[0].replace(/></g, '>\n<'));
  const rels = await archive.file('ppt/slides/_rels/slide2.xml.rels').async('string');
  console.log('\n--- rels ---\n' + rels.replace(/></g, '>\n<'));
})();
