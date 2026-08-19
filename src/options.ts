import {
  DEFAULT_LIMITS,
  FILE_OPTION_NAMES,
  LIMIT_NAMES,
  SANITIZE_OPTION_NAMES,
} from "./constants.js";
import { createError } from "./errors.js";
import type {
  ExcelSanitizationLimits,
  NormalizedFileOptions,
  NormalizedSanitizeOptions,
} from "./types.js";

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function validateKnownKeys(
  value: Readonly<Record<string, unknown>>,
  knownKeys: ReadonlySet<string>,
  label: string,
): void {
  for (const key of Object.keys(value)) {
    if (!knownKeys.has(key)) {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: `Unknown ${label}: ${key}`,
        details: { key },
      });
    }
  }
}

function isLimitName(name: string): name is keyof ExcelSanitizationLimits {
  return LIMIT_NAMES.has(name);
}

export function normalizeLimits(overrides: unknown): Readonly<ExcelSanitizationLimits> {
  if (overrides === undefined) {
    return DEFAULT_LIMITS;
  }
  if (!isRecord(overrides)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options.limits must be an object",
    });
  }

  validateKnownKeys(overrides, LIMIT_NAMES, "limit");
  const limits: ExcelSanitizationLimits = { ...DEFAULT_LIMITS };
  for (const [name, value] of Object.entries(overrides)) {
    if (!isLimitName(name)) {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: `Unknown limit: ${name}`,
        details: { key: name },
      });
    }
    if (typeof value !== "number" || !Number.isSafeInteger(value) || value <= 0) {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: `options.limits.${name} must be a positive safe integer`,
        details: { limit: name, value },
      });
    }
    limits[name] = value;
  }
  return Object.freeze(limits);
}

export function normalizeOptions(options: unknown, fileOptions: true): NormalizedFileOptions;
export function normalizeOptions(options?: unknown, fileOptions?: false): NormalizedSanitizeOptions;
export function normalizeOptions(
  options?: unknown,
  fileOptions = false,
): NormalizedSanitizeOptions | NormalizedFileOptions {
  if (options !== undefined && !isRecord(options)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options must be an object",
    });
  }

  const rawOptions = options ?? {};
  validateKnownKeys(rawOptions, fileOptions ? FILE_OPTION_NAMES : SANITIZE_OPTION_NAMES, "option");
  if (
    rawOptions.requireSingleWorksheet !== undefined &&
    typeof rawOptions.requireSingleWorksheet !== "boolean"
  ) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options.requireSingleWorksheet must be a boolean",
    });
  }

  let normalizedOverwrite = false;
  if (fileOptions) {
    const overwrite = rawOptions.overwrite;
    if (overwrite !== undefined && typeof overwrite !== "boolean") {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: "options.overwrite must be a boolean",
      });
    }
    normalizedOverwrite = overwrite ?? false;
  }

  const normalized: NormalizedSanitizeOptions = {
    requireSingleWorksheet: rawOptions.requireSingleWorksheet ?? false,
    limits: normalizeLimits(rawOptions.limits),
  };
  if (!fileOptions) {
    return normalized;
  }
  return { ...normalized, overwrite: normalizedOverwrite };
}
