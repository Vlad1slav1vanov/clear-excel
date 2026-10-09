import assert from "node:assert/strict";
import test from "node:test";

import { DEFAULT_LIMITS, ExcelSanitizationError, sanitizeExcel } from "../src/index.js";
import JSZip from "jszip";

import { addEscapingRelationship, createXlsm, createXlsx } from "./fixtures/workbook.js";
import {
  addUnsafeCentralDirectoryName,
  createRuntimeXmlBomb,
  forgeZipEntrySizes,
  inflateDeclaredZipSize,
  markAsMultiDisk,
  markAsZip64,
} from "./fixtures/zip.js";
import { expectError } from "./helpers.js";

test("enforces actual aggregate expansion limits for retained binary parts", async () => {
  const zip = await JSZip.loadAsync(await createXlsm());
  zip.file("xl/vbaProject.bin", Buffer.alloc(2 * 1024 * 1024, 0x41));
  const compressed = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  const forged = forgeZipEntrySizes(compressed, "xl/vbaProject.bin", 1);
  await assert.rejects(
    () => sanitizeExcel(forged, { limits: { maxUncompressedBytes: 1024 * 1024 } }),
    (error: unknown) => {
      assert.ok(error instanceof ExcelSanitizationError);
      assert.equal(error.code, "ERR_LIMIT_EXCEEDED");
      assert.equal(error.details?.limit, "maxUncompressedBytes");
      assert.equal(error.details?.partName, "xl/vbaProject.bin");
      return true;
    },
  );
});

test("rejects false binary sizes and preserves valid binary contents", async () => {
  const zip = await JSZip.loadAsync(await createXlsm());
  const vba = Buffer.alloc(32 * 1024, 0x42);
  zip.file("xl/vbaProject.bin", vba);
  const input = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await expectError(
    () => sanitizeExcel(forgeZipEntrySizes(input, "xl/vbaProject.bin", 1)),
    "ERR_INVALID_XLSX",
  );
  const output = await JSZip.loadAsync(await sanitizeExcel(input));
  assert.deepEqual(await output.file("xl/vbaProject.bin")?.async("nodebuffer"), vba);
});

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
