import { randomUUID } from "node:crypto";
import { open, readFile, realpath, rename, stat, unlink } from "node:fs/promises";
import { basename, dirname, join, resolve } from "node:path";

import { hasCompoundFileSignature } from "../biff/workbook.js";
import { createError } from "../errors.js";
import { excelFileFormat, isExcelFilename, isXlsxFilename } from "../formats.js";
import type { ExcelFileFormat } from "../formats.js";
import { normalizeOptions } from "../options.js";
import type { SanitizeExcelFileOptions } from "../types.js";
import { sanitizeExcel } from "./sanitize-excel.js";

export { isExcelFilename, isXlsxFilename };

export async function sanitizeExcelFile(
  inputPath: string,
  outputPath: string,
  options?: SanitizeExcelFileOptions,
): Promise<void> {
  const format = validateFilePaths(inputPath, outputPath);
  const normalized = normalizeOptions(options, true);
  const absoluteInputPath = resolve(inputPath);
  const absoluteOutputPath = resolve(outputPath);
  if (absoluteInputPath === absoluteOutputPath) {
    throw createError("ERR_SAME_PATH", {
      details: { inputPath: absoluteInputPath, outputPath: absoluteOutputPath },
    });
  }

  const input = await readInputFile(absoluteInputPath);
  await validateDistinctFiles(absoluteInputPath, absoluteOutputPath);
  validateInputFormat(input, format);
  const sanitized = await sanitizeExcel(input, {
    requireSingleWorksheet: normalized.requireSingleWorksheet,
    limits: normalized.limits,
  });

  if (normalized.overwrite) {
    await replaceFileAtomically(absoluteOutputPath, sanitized);
  } else {
    await writeNewFile(absoluteOutputPath, sanitized);
  }
}

function validateFilePaths(inputPath: string, outputPath: string): ExcelFileFormat {
  if (
    typeof inputPath !== "string" ||
    inputPath.length === 0 ||
    typeof outputPath !== "string" ||
    outputPath.length === 0
  ) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "inputPath and outputPath must be non-empty strings",
    });
  }
  const inputFormat = excelFileFormat(inputPath);
  const outputFormat = excelFileFormat(outputPath);
  if (!inputFormat || !outputFormat || inputFormat !== outputFormat) {
    throw createError("ERR_UNSUPPORTED_EXTENSION", {
      details: { inputPath, outputPath, inputFormat, outputFormat },
    });
  }
  return inputFormat;
}

function validateInputFormat(input: Buffer, format: ExcelFileFormat): void {
  const isXls = hasCompoundFileSignature(input);
  if ((format === "xls") !== isXls) {
    throw createError(format === "xls" ? "ERR_INVALID_XLS" : "ERR_INVALID_XLSX", {
      details: { format },
    });
  }
}

async function readInputFile(inputPath: string): Promise<Buffer> {
  try {
    return await readFile(inputPath);
  } catch (error) {
    throw createError("ERR_FILE_READ", {
      cause: error,
      details: { inputPath },
    });
  }
}

async function writeNewFile(outputPath: string, contents: Buffer): Promise<void> {
  try {
    await writeExclusiveFile(outputPath, contents);
  } catch (error) {
    if (hasErrorCode(error, "EEXIST")) {
      throw createError("ERR_OUTPUT_EXISTS", {
        cause: error,
        details: { outputPath },
      });
    }
    throw createError("ERR_FILE_WRITE", {
      cause: error,
      details: { outputPath },
    });
  }
}

async function replaceFileAtomically(outputPath: string, contents: Buffer): Promise<void> {
  const temporaryPath = `${outputPath}.${process.pid}.${randomUUID()}.tmp`;
  let temporaryCreated = false;
  try {
    const existingOutput = await stat(outputPath).catch((error: unknown) => {
      if (hasErrorCode(error, "ENOENT")) {
        return undefined;
      }
      throw error;
    });
    const outputMode = existingOutput ? existingOutput.mode & 0o777 : undefined;
    await writeExclusiveFile(temporaryPath, contents, outputMode);
    temporaryCreated = true;
    await rename(temporaryPath, outputPath);
  } catch (error) {
    if (temporaryCreated) {
      await removeIfPresent(temporaryPath);
    }
    throw createError("ERR_FILE_WRITE", {
      cause: error,
      details: { outputPath },
    });
  }
}

async function writeExclusiveFile(path: string, contents: Buffer, mode?: number): Promise<void> {
  // A failed open does not establish ownership, so it must never trigger unlink.
  const handle = await open(path, "wx", mode);
  try {
    await handle.writeFile(contents);
    if (mode !== undefined) {
      await handle.chmod(mode);
    }
    await handle.close();
  } catch (error) {
    await handle.close().catch(() => undefined);
    await removeIfPresent(path);
    throw error;
  }
}

async function validateDistinctFiles(inputPath: string, outputPath: string): Promise<void> {
  let inputRealPath: string;
  let inputStat: Awaited<ReturnType<typeof stat>>;
  try {
    inputRealPath = await realpath(inputPath);
    inputStat = await stat(inputPath);
  } catch (error) {
    throw createError("ERR_FILE_READ", { cause: error, details: { inputPath } });
  }

  let sameFile: boolean;
  try {
    const outputStat = await stat(outputPath).catch((error: unknown) => {
      if (hasErrorCode(error, "ENOENT")) {
        return undefined;
      }
      throw error;
    });
    const outputRealPath = outputStat
      ? await realpath(outputPath)
      : await realpath(dirname(outputPath))
          .then((parent) => join(parent, basename(outputPath)))
          .catch((error: unknown) => {
            if (hasErrorCode(error, "ENOENT")) {
              return outputPath;
            }
            throw error;
          });
    sameFile =
      inputRealPath === outputRealPath ||
      (outputStat !== undefined &&
        inputStat.dev === outputStat.dev &&
        inputStat.ino === outputStat.ino);
  } catch (error) {
    throw createError("ERR_FILE_WRITE", { cause: error, details: { outputPath } });
  }

  if (sameFile) {
    throw createError("ERR_SAME_PATH", { details: { inputPath, outputPath } });
  }
}

function hasErrorCode(error: unknown, expectedCode: string): boolean {
  return (
    typeof error === "object" && error !== null && "code" in error && error.code === expectedCode
  );
}

async function removeIfPresent(path: string): Promise<void> {
  await unlink(path).catch(() => undefined);
}
