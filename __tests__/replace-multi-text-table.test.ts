import Automizer, { modify, XmlElement, XmlHelper } from '../src/index';
import { MultiTextHelper } from '../src/helper/multitext-helper';
import { vd } from '../src/helper/general-helper';
import { MultiTextParagraph } from '../src/interfaces/imulti-text';

test('replace multi text in a table cell (creates a:txBody when missing).', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    // Use an explicit alias because addSlide() references the template by name
    .load(`SlideWithTables.pptx`, 'tables');

  await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableDefault', [
        (tableShape: XmlElement) => {
          // Find the first table cell
          const cell = tableShape.getElementsByTagName('a:tc')[0] as XmlElement;

          // Run MultiText on the *cell* element (not the shape)
          new MultiTextHelper(cell).run([
            { text: 'Hello table cell' },
            {
              text: 'Hello paragraph',
              style: {
                color: { type: 'schemeClr', value: 'bg1' },
              },
            },
          ] as MultiTextParagraph[]);

          // Assertions: must create a:txBody (not p:txBody) and put text inside
          const newATxBody = cell.getElementsByTagName('a:txBody')[0];
          expect(newATxBody).toBeTruthy();

          const createdText = newATxBody.textContent || '';
          expect(createdText).toContain('Hello table cell');

          // Also ensure we didn't accidentally create a shape txBody inside the cell
          const forbiddenPTxBody = cell.getElementsByTagName('p:txBody')[0];
          expect(forbiddenPTxBody).toBeFalsy();
        },
      ]);
    })
    .write(`replace-multi-text-table.test.pptx`);
});
