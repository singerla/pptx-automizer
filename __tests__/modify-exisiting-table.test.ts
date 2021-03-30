import Automizer, { modify } from '../src/index';

test('create presentation, add and modify an existing table.', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const data = [
    [
      { label: 'my header 1' },
      { label: 'my header 2' },
      { label: 'my header 3' },
      { label: 'my header 4' }
    ],
    [
      { label: 'my cell 1-1' },
      { label: 'my cell 1-2' },
    ],
    [], // we don't want to change body row 2
    [
      {}, // we also want to skip body row3/col1 and row3/col2
      {},
      { label: 'my cell 3-3' },
      { label: 'my cell 3-4' },
    ]
  ]

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`SlideWithTable.pptx`, 'tables');

  const result = await pres
    .addSlide('tables', 1, (slide) => {
      slide.modifyElement('TableHeader', (table) => {
        
        // uncomment next line to dump table's xml in console 
        // modify.dump(table)
        
        data.forEach((row,r) => {
          const tabRow = table.getElementsByTagName('a:tr')[r]
          row.forEach((cell,c) => {
            const tabCell = tabRow.getElementsByTagName('a:tc')[c]

            // We could hopefully find a corresponding text node
            // and replace textContent with our own label.
            if(cell.label !== undefined) {
              tabCell.getElementsByTagName('a:t')
                [0].firstChild.textContent = String(cell.label)
            }
          })
        })
      });
    })
    .write(`modify-existing-table.test.pptx`);

  // expect(result.tables).toBe(2); // tbd
});
