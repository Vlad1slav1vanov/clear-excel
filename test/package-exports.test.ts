import assert from "node:assert/strict";
import test from "node:test";

import * as api from "../src/index.js";

test("ESM entry point exposes the stable public implementation", () => {
  assert.equal(typeof api.sanitizeExcel, "function");
  assert.equal(typeof api.sanitizeExcelFile, "function");
  assert.equal(typeof api.isXlsxFilename, "function");
  assert.equal(typeof api.ExcelSanitizationError, "function");
  assert.equal(api.DEFAULT_LIMITS.maxEntryCount, 10_000);
});
