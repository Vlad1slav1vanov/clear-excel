import assert from "node:assert/strict";
import test from "node:test";

import { DEFAULT_LIMITS, ExcelSanitizationError, sanitizeExcel } from "../src/index.js";
import { addEscapingRelationship, createXlsx } from "./fixtures/workbook.js";
import {
  addUnsafeCentralDirectoryName,
  createRuntimeXmlBomb,
  inflateDeclaredZipSize,
  markAsMultiDisk,
  markAsZip64,
} from "./fixtures/zip.js";
import { expectError } from "./helpers.js";

test("rejects declared and actual XML expansion bombs", async () => {
  await expectError(
    async () => sanitizeExcel(inflateDeclaredZipSize(await createXlsx())),
    "ERR_LIMIT_EXCEEDED",
  );
  await expectError(async () => sanitizeExcel(await createRuntimeXmlBomb()), "ERR_LIMIT_EXCEEDED");
});

test("supports safe custom limits and validates limit options", async () => {
  assert.equal(Object.isFrozen(DEFAULT_LIMITS), true);
  await expectError(
    async () => sanitizeExcel(await createXlsx(), { limits: { maxEntryCount: 1 } }),
    "ERR_LIMIT_EXCEEDED",
  );
  await expectError(
    async () => sanitizeExcel(await createXlsx(), { limits: { maxEntryCount: 0 } }),
    "ERR_INVALID_ARGUMENT",
  );
  await expectError(
    async () =>
      Reflect.apply(sanitizeExcel, undefined, [await createXlsx(), { limits: { typo: 100 } }]),
    "ERR_INVALID_ARGUMENT",
  );
});

test("enforces the optional single-sheet contract", async () => {
  await expectError(
    async () => sanitizeExcel(await createXlsx(2), { requireSingleWorksheet: true }),
    "ERR_WORKSHEET_COUNT",
  );
  await sanitizeExcel(await createXlsx(2));
});

test("rejects malformed and unsupported ZIP structures", async () => {
  await assert.rejects(
    () => sanitizeExcel(Buffer.from("not an OOXML workbook")),
    (error: unknown) => {
      assert.ok(error instanceof ExcelSanitizationError);
      assert.equal(error.code, "ERR_INVALID_XLSX");
      assert.ok(error.cause instanceof Error);
      return true;
    },
  );
  await expectError(
    async () => sanitizeExcel(markAsMultiDisk(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
  await expectError(async () => sanitizeExcel(markAsZip64(await createXlsx())), "ERR_INVALID_XLSX");
});

test("rejects unsafe package paths and relationship targets", async () => {
  await expectError(
    async () => sanitizeExcel(addUnsafeCentralDirectoryName(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
  await expectError(
    async () => sanitizeExcel(await addEscapingRelationship(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
});
