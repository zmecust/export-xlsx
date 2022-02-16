import ExcelExport from "..";
import { SETTINGS_FOR_EXPORT } from "./setting";

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
