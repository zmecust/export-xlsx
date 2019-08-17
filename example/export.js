import ExcelExport from "..";
import { SETTINGS_FOR_EXPORT } from "./setting";

const data = {
  table1: { total: 80 },
  table2: [
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
};

const excelExport = new ExcelExport();
excelExport.downloadExcel(SETTINGS_FOR_EXPORT, data);
