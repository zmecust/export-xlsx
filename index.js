import FileSaver from 'file-saver';
import ExcelJS from 'exceljs/dist/exceljs.min.js';

// Default border style
export const borderStyle = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  right: { style: 'thin' },
  bottom: { style: 'thin' },
};

const generateExcelColumnIndex = () => {
  const str = [];
  for (let i = 65; i < 91; i++) {
    str.push(String.fromCharCode(i));
  }
  return str;
};

// Excel grid index
export const excelColumnIndex = generateExcelColumnIndex().reduce(
  (array, item) => array.concat(generateExcelColumnIndex().map(v => `${item}${v}`)),
  generateExcelColumnIndex()
);

// Define alignment style
export const alignment = {
  middleLeft: { vertical: 'middle', horizontal: 'left' },
  middleRight: { vertical: 'middle', horizontal: 'right' },
  middleCenter: { vertical: 'middle', horizontal: 'center' },
};

// Define data type
export const defaultDataType = {
  date: 'Date',
  number: 'Number',
  string: 'String',
  currency: 'Currency',
  percentage: 'Percentage',
};

const defaultGapBetweenTwoTables = 4;
const defaultNumberFormat = '#,##0.00';
const defaultPercentageFormat = '0.00%';
const tableTitleStyle = {
  font: { bold: true },
  alignment: alignment.middleLeft,
};
const notificationStyle = { font: { bold: true, color: { argb: 'FFFF0000' } } };
const defaultHeaderStyle = {
  font: { bold: true },
  alignment: { wrapText: true, ...alignment.middleCenter },
};
const fillStyleForEditableCell = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFFFF00' },
  bgColor: { indexed: 64 },
};

/**
 * Creates an instance of ExcelExport.
 */
export default class ExcelExport {
  /**
   * @constructor
   */
  constructor() {
    /**
     * Execl worksheet
     * @type {Object}
     */
    this._ws = null;

    /**
     * Current actived row index
     * @type {Number}
     */
    this._activedRowIndex = 1;

    /**
     * Store start and end row index for all table contents
     * @type {Array}
     */
    this._tableContentIndex = [];

    /**
     * Store Callback function to set cell formula
     * @type {Array}
     */
    this._callCellFormula = [];
  }

  /**
   * Resolve formula at row level
   *
   * @param {String} formula
   * @returns {Object}
   */
  resolveRowFormula(formula) {
    // formulaKeys: ['{key1}', '{key2}']
    const formulaKeys = Array.from(new Set(formula.match(/\{[A-Za-z]+\}/g)));
    // columnKeys: ['key1', 'key2']
    const columnKeys = formulaKeys.map(key => key.slice(1, -1));
    return { formulaKeys, columnKeys };
  }

  /**
   * Set formula for single cell
   *
   * @param {Object} cell
   * @param {String} formula
   */
  resolveCellFormula(cell, formula) {
    // formula: '{tableName, tableColumnName, tableRowIndex}'
    const formulaKeys = Array.from(new Set(formula.match(/\{.*?\}/g)));
    const callbackFormula = tableSettings => {
      // Resolve cell formula
      formula = formulaKeys.reduce((result, formulaKey) => {
        const [tableName, tableColumnKey, cellIndex] = formulaKey.slice(1, -1).split(',');
        const tableContentIndex = this._tableContentIndex[Object.keys(tableSettings).findIndex(v => v === tableName)];
        let columnIndex;
        // If tableColumnKey is number type , it is the column index, or else it is the column key which defined by headerDefinition
        if (isNaN(tableColumnKey)) {
          const index = tableSettings[tableName].headerDefinition.findIndex(v => v.key === tableColumnKey);
          columnIndex = excelColumnIndex[index];
        } else {
          columnIndex = excelColumnIndex[tableColumnKey - 1];
        }
        // If cellIndex < 0, it means cell index started with the last one, e.g. '-1' is the last row index.
        const rowIndex =
          cellIndex > 0 ? tableContentIndex.start + Number(cellIndex) - 1 : tableContentIndex.end + Number(cellIndex);
        return result.replace(new RegExp(formulaKey, 'g'), `${columnIndex}${rowIndex}`);
      }, formula);
      // Set formula for this cell
      cell.value = { formula, result: cell.value };
    };
    this._callCellFormula.push(callbackFormula);
  }

  /**
   * Get table headers
   *
   * @param {Object}
   * @returns {Array}
   */
  getTableHeaders({ headerDefinition, headerGroups = [] }) {
    const headers = [headerDefinition];
    let columns = headerDefinition.map(column => headerGroups.find(group => group.key === column.groupKey));
    while (columns.filter(v => !!v).length > 0) {
      headers.unshift(columns);
      columns = columns.map(column => column && headerGroups.find(group => group.key === column.groupKey));
    }
    return headers;
  }

  /**
   * Get number format
   *
   * @param {String} dataType
   * @returns {String}
   */
  getNumberFormat(dataType) {
    let numFmt = '';
    if (dataType === defaultDataType.number) {
      numFmt = defaultNumberFormat;
    }
    if (dataType === defaultDataType.percentage) {
      numFmt = defaultPercentageFormat;
    }
    return numFmt;
  }

  /**
   * Add table title (default alignment is middle and left)
   *
   * @param {String} title
   */
  addTableTitle(title) {
    const row = this._ws.getRow(this._activedRowIndex);
    row.values = [title];
    // Set title font to bold, and alignment to middle left
    row.eachCell(cell => {
      cell.style = tableTitleStyle;
    });
    this._activedRowIndex = this._activedRowIndex + 1;
  }

  /**
   * Add hierarchy structure (middle left alignment)
   *
   * @param {Array} data
   * @param {Object} tableSetting
   */
  addHierarchy(data, { headerDefinition }) {
    const hierarchyColumn = headerDefinition.find(v => v.hierarchy);
    // Add hierarchy structure if table has hierarchy column
    if (hierarchyColumn) {
      const idCol = this._ws.getColumn(hierarchyColumn.key);
      idCol.eachCell((cell, rowNumber) => {
        if (rowNumber >= this._activedRowIndex) {
          const level = data[rowNumber - this._activedRowIndex].level;
          // Hierarchy column must have middle left alignment
          if (level > 0) {
            cell.alignment = {
              ...cell.alignment,
              ...alignment.middleLeft,
              indent: 2 * level,
            };
          } else {
            cell.alignment = { ...cell.alignment, ...alignment.middleLeft };
          }
        }
      });
    }
  }

  /**
   * Add vertical border line for the table
   *
   * @param {Array} data
   * @param {Object} tableSetting
   */
  addVerticalBorderLineForTables(data, headerHeight, { headerDefinition }) {
    for (let i = this._activedRowIndex - headerHeight; i < this._activedRowIndex + data.length; i++) {
      headerDefinition.forEach((column, index) => {
        const colNumber = excelColumnIndex[index];
        const cell = this._ws.getCell(`${colNumber}${i}`);
        cell.border = {
          ...cell.border,
          left: borderStyle.left,
          right: borderStyle.right,
        };
      });
    }
  }

  /**
   * Add horizontal border line for the table content
   *
   * @param {Object} row
   * @param {Array} rowValues
   * @param {Array} tableContent
   * @param {Number} tableContentIndex
   */
  addHorizontalBorderLineForTableContent(row, rowValues, tableContent, tableContentIndex) {
    // Add border line at the top of table content
    if (tableContentIndex === 0) {
      row.eachCell({ includeEmpty: true }, cell => {
        cell.border = { ...cell.border, top: borderStyle.top };
      });
    }
    // Add border line at the bottom of table content
    if (tableContentIndex === tableContent.length - 1) {
      // Special case, when all of the row values is null
      if (rowValues.length === 0) {
        const lastRow = this._activedRowIndex + tableContent.length - 1;
        for (let i = 0; i < tableContent[0].length; i++) {
          const cell = this._ws.getCell(`${excelColumnIndex[i]}${lastRow}`);
          cell.border = { ...cell.border, bottom: borderStyle.bottom };
        }
      } else {
        row.eachCell({ includeEmpty: true }, cell => {
          cell.border = { ...cell.border, bottom: borderStyle.bottom };
        });
      }
    }
  }

  /**
   * Add formula
   *
   * @param {Array} data
   * @param {Object} tableSetting
   */
  addFormula(data, { headerDefinition }) {
    const hierarchyColumn = headerDefinition.find(column => column.hierarchy);
    headerDefinition.forEach((column, index) => {
      const columnIndex = excelColumnIndex[index];
      // Add formula at column level
      if (hierarchyColumn && column.selfSum) {
        for (let level = 0; level < data.length; level++) {
          data.forEach((v, rowIndex) => {
            if (v.level !== undefined && v.level === level) {
              const rowNums = data.map((t, i) => (t.parentId === v.id ? i : undefined)).filter(t => t !== undefined);
              const formula = rowNums.reduce((result, rowNum) => {
                result += `+${columnIndex}${this._activedRowIndex + rowNum}`;
                return result;
              }, '');
              if (formula) {
                const cell = this._ws.getCell(`${columnIndex}${this._activedRowIndex + rowIndex}`);
                cell.value = { formula: formula.slice(1), result: cell.value };
              }
            }
          });
        }
      }
      // Add formula of row level
      if (column.rowFormula) {
        const { formulaKeys, columnKeys } = this.resolveRowFormula(column.rowFormula);
        data.forEach((v, index) => {
          const currentRowIndex = this._activedRowIndex + index;
          const leafDataNorNot = hierarchyColumn ? data.findIndex(t => t.parentId === v.id) === -1 : true;
          // Leaf node or cell with percentage data-type need to add row level formula
          if (column.dataType === defaultDataType.percentage || leafDataNorNot) {
            const formula = formulaKeys.reduce((result, formulaKey, index) => {
              const columnKeyIndex = headerDefinition.findIndex(column => column.key === columnKeys[index]);
              return result.replace(
                new RegExp(formulaKey, 'g'),
                `${excelColumnIndex[columnKeyIndex]}${currentRowIndex}`
              );
            }, column.rowFormula);
            // Add formula style
            const cell = this._ws.getCell(`${columnIndex}${currentRowIndex}`);
            cell.value = { formula, result: cell.value };
          }
        });
      }
    });
  }

  /**
   * Merge cell for headers
   *
   * @param {Array} header
   * @param {Number} numberIndex
   */
  mergeCellForHeader(header, numberIndex) {
    for (let i = 0; i < header.length; ) {
      if (header[i]) {
        let end = excelColumnIndex[i];
        const start = excelColumnIndex[i];
        const currentName = header[i];
        i++;
        while (i < header.length) {
          if (currentName === header[i]) {
            end = excelColumnIndex[i];
            i++;
          } else {
            break;
          }
        }
        if (start !== end) {
          this._ws.mergeCells(`${start}${numberIndex}`, `${end}${numberIndex}`);
        }
      } else {
        i++;
      }
    }
  }

  /**
   * Add table columns
   *
   * @param {Array} headerDefinition
   */
  addTableColumns(headerDefinition) {
    // Set table column style
    this._ws.columns = headerDefinition.map(column => {
      // If data type is String, set cell to middle center alignment
      if (!column.dataType) {
        column.style = { ...column.style, alignment: alignment.middleCenter };
      }
      // If data type is Number and Percentage, set number format
      if ([defaultDataType.number, defaultDataType.percentage].includes(column.dataType)) {
        column.style = {
          ...column.style,
          numFmt: this.getNumberFormat(column.dataType),
        };
      }
      return column;
    });
  }

  /**
   * Add table headers
   *
   * @param {Object} tableSetting
   * @return {Number}
   */
  addTableHeaders({ headerDefinition, headerGroups }) {
    const headers = this.getTableHeaders({ headerDefinition, headerGroups });
    // Add table headers
    headers.forEach((header, index) => {
      let row;
      if (index === 0) {
        row = this._ws.getRow(this._activedRowIndex);
        // Add display name for the first header
        row.values = header.map(v => (v ? v.name : null));
        // Add crosswise border line at the top of table header
        row.eachCell({ includeEmpty: true }, cell => {
          cell.border = { top: borderStyle.top };
        });
      } else {
        row = this._ws.addRow(header.map(v => (v ? v.name : null)));
      }
      // Set style for all headers
      row.eachCell((cell, colNumber) => {
        const headerStyle = header[colNumber - 1].headerStyle || {};
        cell.style = { ...defaultHeaderStyle, ...cell.style, ...headerStyle };
        cell.border = { ...cell.border, bottom: borderStyle.bottom };
      });
    });
    // Merge cell for headers
    headers.forEach((header, index) =>
      this.mergeCellForHeader(header.map(v => (v && v.name) || null), this._activedRowIndex + index)
    );
    this.addTableColumns(headerDefinition);
    const headerHeight = headers.length;
    this._activedRowIndex = this._activedRowIndex + headerHeight;
    return headerHeight;
  }

  /**
   * Add column width
   *
   * @param {Number} columnWidth
   * @param {Number} columns
   */
  addColumnWidth(columnWidth, columns) {
    if (columnWidth) {
      for (let i = 1; i <= columns; i++) {
        this._ws.getColumn(i).width = columnWidth;
      }
    }
  }

  /**
   * Add notification for table
   *
   * @param {Object} tableSetting
   */
  addNotification(tableSetting) {
    if (tableSetting.notification) {
      const row = this._ws.addRow([tableSetting.notification]);
      // Set notification style
      row.eachCell(cell => {
        cell.style = notificationStyle;
      });
      this._activedRowIndex = this._activedRowIndex + 1;
    }
  }

  /**
   * Add style for single cell
   *
   * @param {Object} cell
   * @param {Array} data
   */
  addCellStyle(cell, data) {
    const { font, dataType, cellFormula } = data.style;
    // Add number format
    cell.numFmt = this.getNumberFormat(dataType);
    // Add font style
    cell.font = { ...cell.font, ...font };
    // Add formula
    if (cellFormula) {
      this.resolveCellFormula(cell, cellFormula);
    }
  }

  /**
   * Add table value
   *
   * @param {Array} data
   * @param {Boolean} noHeader
   */
  addTableValues(data, noHeader) {
    data.forEach((value, index) => {
      const cellsHasStyleSetting = [];
      if (noHeader) {
        value = value.map((v, index) => {
          if (v !== null && typeof v === 'object') {
            cellsHasStyleSetting.push(index + 1);
            return v.value;
          }
          return v;
        });
      }
      const row = this._ws.addRow(value);
      // Add border line for table content
      this.addHorizontalBorderLineForTableContent(row, value, data, index);
      // Add cell style if needed
      if (cellsHasStyleSetting.length > 0) {
        row.eachCell({ includeEmpty: true }, (cell, cellNumber) => {
          const objectData = data[index][cellNumber - 1];
          if (cellsHasStyleSetting.includes(cellNumber) && objectData !== null && typeof objectData === 'object') {
            this.addCellStyle(cell, objectData);
          }
        });
      }
    });
  }

  /**
   * Leaf node and editable cells could add fill style (yellow background)
   *
   * @param {Array} data
   * @param {Object} tableSetting
   */
  addFillStyleForEditableCells(data, { headerDefinition }) {
    headerDefinition.forEach((column, index) => {
      const columnIndex = excelColumnIndex[index];
      data.forEach((v, rowIndex) => {
        // Leaf node need to add these formula
        const leafDataIndex = data.findIndex(t => t.parentId === v.id);
        if (leafDataIndex === -1 && column.editable) {
          const cell = this._ws.getCell(`${columnIndex}${this._activedRowIndex + rowIndex}`);
          cell.fill = { ...cell.fill, ...fillStyleForEditableCell };
        }
      });
    });
  }

  /**
   * Add table content index
   *
   * @param {Number} contentLength
   */
  addTableContentIndex(contentLength) {
    this._tableContentIndex.push({
      start: this._activedRowIndex,
      end: this._activedRowIndex + contentLength,
    });
    this._activedRowIndex = this._activedRowIndex + contentLength;
  }

  /**
   * Set space between two tables
   *
   * @param {Number} gapBetweenTwoTables
   */
  setGapBetweenTwoTables(gapBetweenTwoTables) {
    this._activedRowIndex = this._activedRowIndex + gapBetweenTwoTables || defaultGapBetweenTwoTables;
  }

  /**
   * Create work sheet
   *
   * @param {Array} data
   * @param {Object} wsSetting
   * @param {Number} wsIndex
   */
  createWorkSheet(data, wsSetting, wsIndex) {
    const { columnWidth, gapBetweenTwoTables, tableSettings } = wsSetting;
    // Create tables for this work sheet
    Object.keys(tableSettings).forEach(index => {
      let tableSetting = tableSettings[index];
      if (typeof tableSetting === 'function') {
        tableSetting = tableSetting(data[wsIndex][index]);
      }
      // Add table title
      if (tableSetting.tableTitle) {
        this.addTableTitle(tableSetting.tableTitle);
      }
      if (tableSetting.headerDefinition) {
        const tableData = data[wsIndex][index];
        // Add table headers
        const headerHeight = this.addTableHeaders(tableSetting);
        // Add table value
        this.addTableValues(tableData, false);
        // Add hierarchy structure for column which has hierarchy setting
        this.addHierarchy(tableData, tableSetting);
        // Add vertical border line
        this.addVerticalBorderLineForTables(tableData, headerHeight, tableSetting);
        // Add fill style for cells which could editable
        this.addFillStyleForEditableCells(tableData, tableSetting);
        // Add table column formula
        this.addFormula(tableData, tableSetting);
        // Push start and end row index of table content to '_tableContentIndex'
        this.addTableContentIndex(tableData.length);
        // Add notify at table buttom
        this.addNotification(tableSetting);
      } else {
        const rows = tableSetting.rowsDefinition;
        // Set column width if table setting has not header definition
        this.addColumnWidth(columnWidth, rows.length);
        // Add table value
        this.addTableValues(rows, true);
        // Push start and end row index of table content to '_tableContentIndex'
        this.addTableContentIndex(rows.length);
      }
      // Set space between two tables
      this.setGapBetweenTwoTables(gapBetweenTwoTables);
    });
    // Callback function to set cell formula
    this._callCellFormula.forEach(callbackFunc => callbackFunc(tableSettings));
  }

  /**
   * Generate excel file
   *
   * @param {Object} worksheet
   * @param {String} filename
   */
  async generateExcel(wb, filename) {
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.ms-excel' });
    FileSaver.saveAs(blob, `${filename}.xlsx`);
  }

  async downloadExcel(settings = {}, data = []) {
    const wb = new ExcelJS.Workbook();
    const { fileName, workSheets } = settings;
    // Create worksheets
    workSheets.forEach((wsSetting, wsIndex) => {
      const { sheetName, startingRowNumber } = wsSetting;
      this._ws = wb.addWorksheet(sheetName);
      // Actived row index
      this._activedRowIndex = startingRowNumber;
      this.createWorkSheet(data, wsSetting, wsIndex);
    });
    // Generate excel file
    await this.generateExcel(wb, fileName);
  }
}
