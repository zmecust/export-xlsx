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
        table1: {
          tableTitle: "Score",
          headerGroups: [
            {
              name: '',
              key: 'void',
              groupKey: 'directions',
            },
            {
              name: 'Science',
              key: 'science',
              groupKey: 'directions',
            },
            {
              name: 'Directions',
              key: 'directions',
            },
          ],
          headerDefinition: [
            {
              name: '#',
              key: 'number',
            },
            {
              name: 'Name',
              key: 'name',
            },
            {
              name: 'SUM',
              key: 'sum',
              groupKey: 'void',
              rowFormula: '{math}+{physics}+{chemistry}+{informatics}+{literature}+{foreignLang}',
            },
            {
              name: 'Mathematics',
              key: 'math',
              groupKey: 'science',
            },
            {
              name: 'Physics',
              key: 'physics',
              groupKey: 'science',
            },
            {
              name: 'Chemistry',
              key: 'chemistry',
              groupKey: 'science',
            },
            {
              name: 'Informatics',
              key: 'informatics',
              groupKey: 'science',
            },
            {
              name: 'Literature',
              key: 'literature',
              groupKey: 'science',
            },
            {
              name: 'Foreign lang.',
              key: 'foreignLang',
              groupKey: 'science',
            },
            {
              name: 'AVG',
              key: 'avg',
              groupKey: 'void',
              rowFormula: '{sum}/6',
            }
          ],
        }
      }
    },
  ],
};
```

### How to use

```
import ExcelExport from 'export-xlsx';
import { SETTINGS_FOR_EXPORT } from './setting';

const data = {
  table1: [
    {
      number: 1,
      name: 'Jack',
      sum: '',
      math: 1,
      physics: 2,
      chemistry: 2,
      informatics: 1,
      literature: 2,
      foreignLang: 1,
      avg: '',
    },
    {
      number: 2,
      name: 'Peter',
      sum: '',
      math: 2,
      physics: 2,
      chemistry: 1,
      informatics: 2,
      literature: 2,
      foreignLang: 1,
      avg: '',
    },
  ]
};

const excelExport = new ExcelExport();
excelExport.downloadExcel(SETTINGS_FOR_EXPORT, data);
```

### Result

![](https://note.laravue.org/images/note/20190901-201545-557.png)

## Contact

If you have any questions, please contact me **`root@laravue.org`**

## License
The MIT license
