import { alignment, defaultDataType } from '..';

// Export settings
export const SETTINGS_FOR_EXPORT = {
  // Table settings
  fileName: "example",
  workSheets: [
    {
      sheetName: "example",
      startingRowNumber: 2,
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
    }
  ]
};
