import type { ExcelSanitizationLimits } from "./types.js";

export type ExcelSanitizationErrorCode =
  | "ERR_INVALID_ARGUMENT"
  | "ERR_UNSUPPORTED_EXTENSION"
  | "ERR_INVALID_XLS"
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

interface CreateErrorOptions extends ExcelSanitizationErrorOptions {
  message?: string;
}

export type SafetyLimitName = keyof ExcelSanitizationLimits | "partBytes";

const ERROR_MESSAGES: Readonly<Record<ExcelSanitizationErrorCode, string>> = Object.freeze({
  ERR_INVALID_ARGUMENT: "Invalid argument",
  ERR_UNSUPPORTED_EXTENSION:
    "Supported extensions are .xls, .xlsx, and .xlsm; input and output formats must match",
  ERR_INVALID_XLS: "Invalid or unsupported XLS workbook",
  ERR_INVALID_XLSX: "Invalid or unsupported OOXML workbook",
  ERR_LIMIT_EXCEEDED: "Excel workbook exceeds a configured safety limit",
  ERR_WORKSHEET_COUNT: "The workbook must contain exactly one worksheet",
  ERR_SAME_PATH: "Input and output paths must be different",
  ERR_OUTPUT_EXISTS: "The output file already exists",
  ERR_FILE_READ: "Unable to read the input Excel file",
  ERR_FILE_WRITE: "Unable to write the sanitized Excel file",
});

export class ExcelSanitizationError extends Error {
  override readonly name = "ExcelSanitizationError";
  readonly code: ExcelSanitizationErrorCode;
  readonly details?: Readonly<Record<string, unknown>>;

  constructor(
    code: ExcelSanitizationErrorCode,
    message: string,
    options: ExcelSanitizationErrorOptions = {},
  ) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.code = code;
    if (options.details !== undefined) {
      this.details = Object.freeze({ ...options.details });
    }
  }
}

export function createError(
  code: ExcelSanitizationErrorCode,
  options: CreateErrorOptions = {},
): ExcelSanitizationError {
  return new ExcelSanitizationError(code, options.message ?? ERROR_MESSAGES[code], options);
}

export function limitExceeded(
  limit: SafetyLimitName,
  configured: number,
  actual: number,
  partName?: string,
): ExcelSanitizationError {
  return createError("ERR_LIMIT_EXCEEDED", {
    details: {
      limit,
      configured,
      actual,
      ...(partName === undefined ? {} : { partName }),
    },
  });
}
