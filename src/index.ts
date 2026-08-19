export { sanitizeExcel } from "./api/sanitize-excel.js";
export { isExcelFilename, isXlsxFilename, sanitizeExcelFile } from "./api/sanitize-excel-file.js";
export { DEFAULT_LIMITS } from "./constants.js";
export { ExcelSanitizationError } from "./errors.js";
export type { ExcelSanitizationErrorCode, ExcelSanitizationErrorOptions } from "./errors.js";
export type { ExcelFileFormat } from "./formats.js";
export type {
  ExcelSanitizationLimits,
  SanitizeExcelFileOptions,
  SanitizeExcelOptions,
} from "./types.js";
