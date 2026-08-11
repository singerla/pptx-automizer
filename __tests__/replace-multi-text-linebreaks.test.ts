import Automizer, { modify } from '../src/index';
import JSZip from 'jszip';
import fs from 'fs';

// Vertical tab: what PowerPoint returns for a soft line break (Shift+Enter).
// Passing it through unescaped used to produce a file PowerPoint had to repair.
// See https://github.com/singerla/pptx-automizer/issues/186
const VT = '\u000B';

const readSlideXml = async (file: string, slide: number): Promise<string> => {
  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${file}`),
  );
  return archive.file(`ppt/slides/slide${slide}.xml`).async('string');
};

test('setMultiText converts soft line breaks into <a:br/>', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        modify.setMultiText([
          {
            paragraph: {},
            textRuns: [{ text: `Line A${VT}Line B` }, { text: `\nLine C` }],
          },
        ]),
      );
    })
    .write(`modify-multi-text-linebreaks.test.pptx`);

  const xml = await readSlideXml('modify-multi-text-linebreaks.test.pptx', 2);

  // The control character must not survive into the XML, it is not valid XML 1.0
  expect(xml).not.toContain(VT);

  // Each break becomes an <a:br/> *between* two runs, never text inside <a:t>
  expect(xml).toMatch(
    /<a:t>Line A<\/a:t><\/a:r><a:br>.*?<\/a:br><a:r>.*?<a:t>Line B<\/a:t>/,
  );
  expect(xml).toMatch(
    /<a:t>Line B<\/a:t><\/a:r><a:br>.*?<\/a:br><a:r>.*?<a:t>Line C<\/a:t>/,
  );
  expect(xml.match(/<a:br>/g)).toHaveLength(2);

  // No line break may remain inside a text node
  expect(xml).not.toMatch(/<a:t>[^<]*[\r\n][^<]*<\/a:t>/);

  // The <a:br/> inherits the run style, so the line height stays intact
  expect(xml).toContain('<a:br><a:rPr sz="2000" i="1"/></a:br>');
});

test('setMultiText keeps a hyperlink across a soft line break', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer.loadRoot(`RootTemplate.pptx`).load(`TextReplace.pptx`);

  await pres
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        modify.setMultiText([
          {
            paragraph: {},
            textRuns: [
              {
                text: `Broken${VT}Link`,
                style: { hyperlink: { target: 'https://github.com' } },
              },
            ],
          },
        ]),
      );
    })
    .write(`modify-multi-text-linebreaks-hyperlink.test.pptx`);

  const xml = await readSlideXml(
    'modify-multi-text-linebreaks-hyperlink.test.pptx',
    2,
  );

  expect(xml).not.toContain(VT);

  // Both halves of the split run keep pointing at the same relationship
  const relIds = Array.from(xml.matchAll(/<a:hlinkClick r:id="([^"]+)"/g)).map(
    (match) => match[1],
  );
  expect(relIds).toHaveLength(2);
  expect(relIds[0]).toEqual(relIds[1]);
});
