---
title: Modify Charts
description: Inject chart data, fine-tune axes, legends and labels, handle extended chart types, and read chart data back.
---

All data and styles of a chart can be modified. Please note that if your template contains more data than your data object, Automizer will remove these extra nodes. Conversely, if you provide more data, new nodes will be cloned from the first existing one in the template.

Note that the chart has to be a **native chart** in the template.

```ts
// Modify an existing chart on an added slide.
pres.addSlide('charts', 2, (slide) => {
  slide.modifyElement('ColumnChart', [
    // Use an object like this to inject the new chart data.
    // Additional series and categories will be copied from
    // previous sibling.
    modify.setChartData({
      series: [
        { label: 'series 1' },
        { label: 'series 2' },
        { label: 'series 3' },
      ],
      categories: [
        { label: 'cat 2-1', values: [50, 50, 20] },
        { label: 'cat 2-2', values: [14, 50, 20] },
        { label: 'cat 2-3', values: [15, 50, 20] },
        { label: 'cat 2-4', values: [26, 50, 20] },
      ],
    }),
  ]);
});
```

Find out more about modifying charts:

- [Modify chart axis](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-axis.test.ts)
- [Dealing with bubble charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-bubbles.test.ts)
- [Vertical line charts](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-chart-vertical-lines.test.ts)
- [Style chart series and data points](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-chart-styled.test.ts)

## Modify Extended Charts

If you need to modify extended chart types, such like waterfall or map charts, you need to use `modify.setExtendedChartData`.

```ts
// Add and modify a waterfall chart on slide.
pres.addSlide('charts', 2, (slide) => {
  slide.addElement('ChartWaterfall.pptx', 1, 'Waterfall 1', [
    modify.setExtendedChartData(<ChartData>{
      series: [{ label: 'series 1' }],
      categories: [
        { label: 'cat 2-1', values: [100] },
        { label: 'cat 2-3', values: [50] },
        { label: 'cat 2-4', values: [-40] },
        // ...
      ],
    }),
  ]);
});
```

## Additional chart modifiers

Besides `modify.setChartData` and `modify.setExtendedChartData`, the `modify` namespace exposes a number of helpers to fine-tune chart appearance and special chart types. Each of them returns a modification callback that can be passed to `slide.modifyElement()`.

- Chart title (requires an already existing, manually edited title)

```ts
slide.modifyElement('ColumnChart', [modify.setChartTitle('My new title')]);
```

- Axis range and number format. Only manually scaled (non-"Auto") min/max values can be altered.

```ts
slide.modifyElement('ColumnChart', [
  modify.setAxisRange({
    axisIndex: 0, // index of c:valAx, defaults to 0
    min: 0,
    max: 100,
    majorUnit: 20,
    minorUnit: 5,
    formatCode: '0.0',
    sourceLinked: false,
  }),
]);
```

- Legend position and visibility. Legend coordinates are shares of the chart coordinates (e.g. `w: 0.5` means "half of chart width").

```ts
slide.modifyElement('ColumnChart', [
  // move/resize the legend
  modify.setLegendPosition({ x: 0.8, y: 0.1, w: 0.2, h: 0.3 }),
  // set legend coordinates to zero so a user can maximize it easily
  modify.minimizeChartLegend(),
  // completely remove the legend (PowerPoint will maximize the chart space)
  modify.removeChartLegend(),
]);
```

- Plot area. Requires a `c:manualLayout` element, which only exists if the plot area was edited manually in PowerPoint before. Coordinates are shares of the chart coordinates.

```ts
slide.modifyElement('ColumnChart', [
  modify.setPlotArea({ x: 0.1, y: 0.1, w: 0.8, h: 0.8 }),
]);
```

- Data labels

```ts
import { LabelPosition } from 'pptx-automizer';

slide.modifyElement('ColumnChart', [
  // format the data labels of all series (or a single series by index)
  modify.setDataLabelAttributes({
    applyToSeries: 0, // omit to apply to all series
    dLblPos: LabelPosition.OutsideEnd,
    formatCode: '0.0%',
    sourceLinked: false,
    showVal: true,
    showSerName: false,
    showCatName: false,
    showPercent: false,
    showLegendKey: false,
  }),
  // or remove all data labels
  modify.removeDataLabels(),
]);
```

- Waterfall total column. Mark the last column (or a specific index) of an extended waterfall chart as a total.

```ts
slide.modifyElement('Waterfall 1', [
  modify.setWaterFallColumnTotalToLast(), // or pass an index, e.g. (3)
]);
```

- Special chart types. Use dedicated modifiers for scatter, combo, bubble and vertical line charts. They accept the same `ChartData` object as `setChartData`.

```ts
slide.modifyElement('ScatterChart', [modify.setChartScatter(chartData)]);
slide.modifyElement('ComboChart', [modify.setChartCombo(chartData)]);
slide.modifyElement('BubbleChart', [modify.setChartBubbles(chartData)]);
slide.modifyElement('VerticalLineChart', [
  modify.setChartVerticalLines(chartData),
]);
```

## Read chart data

The `read` namespace provides callbacks to read information out of a chart without modifying it. They populate the object you pass in (see `__tests__/read-chart-data.test.js`).

```ts
import { read, WorkbookData, ChartInfo } from 'pptx-automizer';

const workbookData: WorkbookData = [];
const chartInfo: ChartInfo = { series: [] };

slide.modifyElement('ColumnChart', [
  // read the raw rows of the embedded workbook into workbookData
  read.readWorkbookData(workbookData),
  // read series colors and the detected chart type into chartInfo
  read.readChartInfo(chartInfo),
]);
```
