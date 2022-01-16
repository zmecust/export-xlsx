import { alignment, defaultDataType } from '..';

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
        table1: data => ({
          tableTitle: 'Total',
          rowsDefinition: Object.keys(data).map(key => [
            key,
            {
              value: data[index],
              style: {
                dataType: defaultDataType.number,
                cellFormula: '{table1,4,-1}+{table1,5,-1}',
              },
            },
          ]),
        }),
        table2: {
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
