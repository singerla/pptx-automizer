import Automizer, { XmlElement, XmlHelper } from './index';
import { vd } from './helper/general-helper';

const run = async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/../__tests__/pptx-templates`,
    outputDir: `${__dirname}/../__tests__/pptx-output`,
  });

  const updateId = (element: XmlElement, tag: string, id: number) => {
    element.getElementsByTagName(tag).item(0).setAttribute('val', String(id));
  };

  const rows = ['row 1', 'row 2'];
  const subs = ['sub 1'];

  automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`NestedTables.pptx`, 'tables')
    .addSlide('tables', 1, async (slide) => {
      const info = await slide.getElement('NestedTable2');
      const data = info.getTableData().body;

      slide.modifyElement('NestedTable2', (element: XmlElement) => {
        const table = element.getElementsByTagName('a:tbl').item(0);
        const tblGrid = element.getElementsByTagName('a:tblGrid').item(0);

        data.forEach((tplCell: any) => {
          if (tplCell.text === '{{each:row}}') {
            rows.forEach((rowKey, r) => {
              const newRow = XmlHelper.appendClone(tplCell.rowXml, table);
              updateId(newRow, 'a16:rowId', r);
            });
          }
        });

        data.forEach((tplCell: any) => {
          if (tplCell.text === '{{each:sub}}') {
            subs.forEach((colKey, c) => {
              if (tplCell.gridSpan) {
                const rows = element.getElementsByTagName('a:tr');
                for (let r = 0; r < rows.length; r++) {
                  const row = rows.item(r);
                  const columns = row.getElementsByTagName('a:tc');
                  const maxC = tplCell.column + tplCell.gridSpan;
                  for (let c = tplCell.column; c < maxC; c++) {
                    const sourceCell = columns.item(c);
                    const newCell = XmlHelper.appendClone(sourceCell, row);
                  }
                  XmlHelper.moveChild(
                    row.getElementsByTagName('a:extLst').item(0),
                  );
                }
              }
            });

            subs.forEach((colKey, ci) => {
              const maxC = tplCell.column + tplCell.gridSpan;
              for (let c = tplCell.column; c < maxC; c++) {
                const sourceTblGridCol = tblGrid
                  .getElementsByTagName('a:gridCol')
                  .item(c);
                const newCol = XmlHelper.appendClone(sourceTblGridCol, tblGrid);
                updateId(newCol, 'a16:colId', c * (ci + 1));
              }
            });
          }
        });

        // XmlHelper.dump(element);
      });
    })
    .write(`dev.pptx`)
    .then((summary) => {
      console.log(summary);
    });
};

run().catch((error) => {
  console.error(error);
});
