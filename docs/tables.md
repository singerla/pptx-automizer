---
title: Modify Tables
description: Fill, expand and style PowerPoint tables with setTable and the table helpers.
---

You can use a PowerPoint table and add/modify data and style. It is also possible to add rows and columns and to style cells.

```ts
const pres = automizer
  .loadRoot(`RootTemplate.pptx`)
  .load(`SlideWithTables.pptx`, 'tables');

const result = await pres.addSlide('tables', 3, (slide) => {
  slide.modifyElement('TableWithEmptyCells', [
    modify.setTable({
      // Use an array of rows to insert data.
      // use `label` key for your information only
      body: [
        { label: 'item test r1', values: ['test1', 10, 16, 12, 11] },
        { label: 'item test r2', values: ['test2', 12, 18, 15, 12] },
        { label: 'item test r3', values: ['test3', 14, 12, 11, 14] },
      ],
    }),
  ]);
});
```

Note that the table has to be a **native table** in the template — a grouped-shape "fake table" cannot be filled with `setTable`.

## Table helpers

`ModifyTableHelper` provides rich control over existing tables.

- Fill table data and auto-adjust size

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTable({
    body: [
      { label: 'r1', values: ['A', 1] },
      { label: 'r2', values: ['B', 2] },
    ],
  }),
]);
```

- Expand rows/columns by tag before filling

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTable(
    {
      body: [ /* ... */ ],
    },
    {
      expand: [
        { tag: '<<ROW>>', count: 3, mode: 'row' },
        { tag: '<<COL>>', count: 2, mode: 'column' },
      ],
      adjustHeight: true,
      adjustWidth: true,
    }
  ),
]);
```

- Set fixed row heights / column widths

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.updateRowHeight(0, CmToDxa(1)),
  ModifyTableHelper.updateColumnWidth(1, CmToDxa(3)),
]);
```

- Apply a table style and header/column banding flags

```ts
slide.modifyElement('MyTable', [
  ModifyTableHelper.setTableStyle('TableStyleMedium2', [
    'firstRow', 'bandRow',
  ]),
]);
```

Additional convenience methods:

- `ModifyTableHelper.setTableData(data)` – just set data without sizing
- `ModifyTableHelper.adjustHeight(data)` / `adjustWidth(data)` – recompute sizes only

## Find out more

- [Modify and style table cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table.test.ts)
- [Insert data into table with empty cells](https://github.com/singerla/pptx-automizer/blob/main/__tests__/modify-existing-table-create-text.test.ts)
