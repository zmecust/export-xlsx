# export-xlsx &middot; [![GitHub license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/zmecust/export-xlsx/blob/master/LICENCE) [![Build Status](https://travis-ci.org/zmecust/export-xlsx.svg)](https://travis-ci.org/zmecust/export-xlsx) [![npm version](https://badge.fury.io/js/export-xlsx.svg)](https://badge.fury.io/js/export-xlsx)

## Export Excel

- support formula
- hierarchy structure
- multi-headers


## Getting started

    $ npm i -S export-xlsx


## Usage

### Make excel setting

```
import { alignment, defaultDataType } from 'export-xlsx';

// Export settings
export const SETTINGS_FOR_EXPORT = {
  // Table settings
  fileName: 'example',
  workSheets: [
    {
      sheetName: 'example',
      startingRowNumber: 2,
      gapBetweenTwoTables: 2,
      tableSettings: {
        data: {
          importable: true,
          tableTitle: 'Score',
          notification: 'Notify: only yellow background cell could edit!',
          headerGroups: [
            {
              name: 'Score',
              key: 'score',
            },
          ],
          headerDefinition: [
            {
              name: 'Id',
              key: 'id',
              width: 25,
              hierarchy: true,
              checkable: true,
            },
            {
              name: 'Number',
              key: 'number',
              width: 18,
              checkable: true,
              style: { alignment: alignment.middleCenter },
            },
            {
              name: 'Name',
              key: 'name',
              width: 18,
              style: { alignment: alignment.middleCenter },
            },
            {
              name: 'A',
              key: 'a',
              width: 18,
              groupKey: 'score',
              dataType: defaultDataType.number,
              selfSum: true,
              editable: true,
            },
            {
              name: 'B',
              key: 'b',
              width: 18,
              groupKey: 'score',
              dataType: defaultDataType.number,
              selfSum: true,
              editable: true,
            },
            {
              name: 'Total',
              key: 'total',
              width: 18,
              dataType: defaultDataType.number,
              selfSum: true,
              rowFormula: '{a}+{b}',
            },
          ],
        },
      },
    },
  ],
};
```

### How to use

```
import ExcelExport from 'export-xlsx';
import { SETTINGS_FOR_EXPORT } from './setting';

const data = [
    {
      data: [
        {
          id: 1,
          level: 0,
          number: '0001',
          name: '0001',
          a: 50,
          b: 45,
          total: 95,
        },
        {
          id: 2,
          parentId: 1,
          level: 1,
          number: '0001-1',
          name: '0001-1',
          a: 20,
          b: 25,
          total: 45,
        },
        {
          id: 3,
          parentId: 2,
          level: 1,
          number: '0001-2',
          name: '0001-2',
          a: 30,
          b: 20,
          total: 50,
        },
        {
          id: 4,
          level: 0,
          number: '0002',
          name: '0002',
          a: 40,
          b: 40,
          total: 80,
        }
      ]
    }
];

const excelExport = new ExcelExport();
excelExport.downloadExcel(SETTINGS_FOR_EXPORT, data);
```

### Result

![](https://note.laravue.org/images/note/20190901-201545-557.png)

## Contact

If you have any questions, please contact me **`root@laravue.org`**

## License
The MIT license
