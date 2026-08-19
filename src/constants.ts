import type { ExcelSanitizationLimits } from "./types.js";

export const DEFAULT_LIMITS: Readonly<ExcelSanitizationLimits> = Object.freeze({
  maxEntryCount: 10_000,
  maxUncompressedBytes: 512 * 1024 * 1024,
  maxXmlBytes: 512 * 1024 * 1024,
  maxWorksheetXmlBytes: 256 * 1024 * 1024,
  maxStylesXmlBytes: 64 * 1024 * 1024,
  maxMetadataXmlBytes: 16 * 1024 * 1024,
  maxRelationshipsXmlBytes: 8 * 1024 * 1024,
});

export const SANITIZE_OPTION_NAMES: ReadonlySet<string> = new Set([
  "requireSingleWorksheet",
  "limits",
]);

export const FILE_OPTION_NAMES: ReadonlySet<string> = new Set([
  "requireSingleWorksheet",
  "limits",
  "overwrite",
]);

export const LIMIT_NAMES: ReadonlySet<string> = new Set([
  "maxEntryCount",
  "maxUncompressedBytes",
  "maxXmlBytes",
  "maxWorksheetXmlBytes",
  "maxStylesXmlBytes",
  "maxMetadataXmlBytes",
  "maxRelationshipsXmlBytes",
]);
