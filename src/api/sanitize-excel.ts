import { hasCompoundFileSignature, sanitizeBiffWorkbook } from "../biff/workbook.js";
import { ExcelSanitizationError, createError } from "../errors.js";
import { sanitizeOoxmlWorkbook } from "../ooxml/workbook.js";
import { normalizeOptions } from "../options.js";
import type { SanitizeExcelOptions } from "../types.js";

export async function sanitizeExcel(
  buffer: Buffer,
  options?: SanitizeExcelOptions,
): Promise<Buffer> {
  if (!Buffer.isBuffer(buffer)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "buffer must be a Node.js Buffer",
    });
  }

  const normalized = normalizeOptions(options);
  const isBiff = hasCompoundFileSignature(buffer);
  try {
    return isBiff
      ? await sanitizeBiffWorkbook(buffer, normalized)
      : await sanitizeOoxmlWorkbook(buffer, normalized);
  } catch (error) {
    if (error instanceof ExcelSanitizationError) {
      throw error;
    }
    throw createError(isBiff ? "ERR_INVALID_XLS" : "ERR_INVALID_XLSX", { cause: error });
  }
}
