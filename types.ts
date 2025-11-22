
export type ExcelRow = {
  [key: string]: string | number | boolean | null;
};

export type ExcelSheetData = ExcelRow[];
