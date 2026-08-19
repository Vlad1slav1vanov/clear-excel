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

export interface NormalizedSanitizeOptions {
  requireSingleWorksheet: boolean;
  limits: Readonly<ExcelSanitizationLimits>;
}

export interface NormalizedFileOptions extends NormalizedSanitizeOptions {
  overwrite: boolean;
}

export interface ReadBudget {
  consumedBytes: number;
  limitBytes: number;
}
