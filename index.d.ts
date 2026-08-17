/// <reference types="node" />

export interface ExcelSanitizationLimits {
  maxEntryCount: number;
  maxUncompressedBytes: number;
  maxXmlBytes: number;
  maxWorksheetXmlBytes: number;
  maxStylesXmlBytes: number;
  maxMetadataXmlBytes: number;
  maxRelationshipsXmlBytes: number;
}

export interface SanitizeExcelOptions {
  requireSingleWorksheet?: boolean;
  limits?: Partial<ExcelSanitizationLimits>;
}

export interface SanitizeExcelFileOptions extends SanitizeExcelOptions {
  overwrite?: boolean;
}

export type ExcelSanitizationErrorCode =
  | "ERR_INVALID_ARGUMENT"
  | "ERR_UNSUPPORTED_EXTENSION"
  | "ERR_INVALID_XLSX"
  | "ERR_LIMIT_EXCEEDED"
  | "ERR_WORKSHEET_COUNT"
  | "ERR_SAME_PATH"
  | "ERR_OUTPUT_EXISTS"
  | "ERR_FILE_READ"
  | "ERR_FILE_WRITE";

export interface ExcelSanitizationErrorOptions {
  cause?: unknown;
  details?: Readonly<Record<string, unknown>>;
}

export class ExcelSanitizationError extends Error {
  constructor(
    code: ExcelSanitizationErrorCode,
    message: string,
    options?: ExcelSanitizationErrorOptions,
  );
  readonly name: "ExcelSanitizationError";
  readonly code: ExcelSanitizationErrorCode;
  readonly details?: Readonly<Record<string, unknown>>;
}

export const DEFAULT_LIMITS: Readonly<ExcelSanitizationLimits>;

export function isXlsxFilename(name: unknown): name is string;

export function sanitizeExcel(
  buffer: Buffer,
  options?: SanitizeExcelOptions,
): Promise<Buffer>;

export function sanitizeExcelFile(
  inputPath: string,
  outputPath: string,
  options?: SanitizeExcelFileOptions,
): Promise<void>;
