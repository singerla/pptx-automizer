import { ChartData, Workbook } from '../types/types';

export class XmlChartHelper {
  static setChartData(
    chart: XMLDocument,
    workbook: Workbook,
    data: ChartData,
  ): void {
    const series = chart.getElementsByTagName('c:ser');

    for (const c in data.categories) {
      for (const s in data.categories[c].values) {
        series[s].getElementsByTagName('c:cat')[0].getElementsByTagName('c:v')[
          c
        ].firstChild.textContent = data.categories[c].label;

        series[s].getElementsByTagName('c:v')[0].firstChild.textContent =
          data.series[s].label;

        series[s].getElementsByTagName('c:val')[0].getElementsByTagName('c:v')[
          c
        ].firstChild.textContent = String(data.categories[c].values[s]);
      }
    }

    XmlChartHelper.setWorkbookData(workbook, data);
  }

  static setWorkbookData(workbook: Workbook, data: ChartData): void {
    const rows = workbook.sheet.getElementsByTagName('row');

    for (const c in data.categories) {
      const r = Number(c) + 1;
      const stringId = XmlChartHelper.appendSharedString(
        workbook.sharedStrings,
        data.categories[c].label,
      );
      const rowLabel = rows[r]
        .getElementsByTagName('c')[0]
        .getElementsByTagName('v')[0];
      rowLabel.firstChild.textContent = String(stringId);

      for (const s in data.categories[c].values) {
        const v = Number(s) + 1;
        rows[r]
          .getElementsByTagName('c')
          [v].getElementsByTagName('v')[0].firstChild.textContent = String(
          data.categories[c].values[s],
        );
      }
    }

    for (const s in data.series) {
      const c = Number(s) + 1;
      const colLabel = rows[0]
        .getElementsByTagName('c')
        [c].getElementsByTagName('v')[0];
      const stringId = XmlChartHelper.appendSharedString(
        workbook.sharedStrings,
        data.series[s].label,
      );

      colLabel.firstChild.textContent = String(stringId);

      workbook.table
        .getElementsByTagName('tableColumn')
        [c].setAttribute('name', data.series[s].label);
    }
  }

  static appendSharedString(
    sharedStrings: Document,
    stringValue: string,
  ): number {
    const strings = sharedStrings.getElementsByTagName('sst')[0];
    const newLabel = sharedStrings.createTextNode(stringValue);
    const newText = sharedStrings.createElement('t');
    newText.appendChild(newLabel);

    const newString = sharedStrings.createElement('si');
    newString.appendChild(newText);

    strings.appendChild(newString);

    return strings.getElementsByTagName('si').length - 1;
  }
}
