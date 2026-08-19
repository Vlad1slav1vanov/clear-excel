export type ExcelFileFormat = "xls" | "xlsx" | "xlsm";

export function excelFileFormat(name: unknown): ExcelFileFormat | undefined {
  if (typeof name !== "string") {
    return undefined;
  }
  const extension = /\.(xls|xlsx|xlsm)$/i.exec(name)?.[1]?.toLowerCase();
  return extension === "xls" || extension === "xlsx" || extension === "xlsm"
    ? extension
    : undefined;
}

export function isExcelFilename(name: unknown): name is string {
  return excelFileFormat(name) !== undefined;
}

export function isXlsxFilename(name: unknown): name is string {
  return excelFileFormat(name) === "xlsx";
}
