import assert from "node:assert/strict";
import test from "node:test";

import * as CFB from "cfb";
import JSZip from "jszip";
import { read } from "@e965/xlsx";

import { isExcelFilename, isXlsxFilename, sanitizeExcel } from "../src/index.js";
import {
  addXlsCompoundStorages,
  createXls,
  createXlsm,
  markXlsEncrypted,
} from "./fixtures/workbook.js";
import { expectError } from "./helpers.js";

test("validates public arguments and supported Excel filenames", async () => {
  assert.equal(isExcelFilename("report.xls"), true);
  assert.equal(isExcelFilename("REPORT.XLSM"), true);
  assert.equal(isExcelFilename("report.xlsx"), true);
  assert.equal(isExcelFilename("report.csv"), false);
  assert.equal(isExcelFilename(null), false);
  assert.equal(isXlsxFilename("report.xlsx"), true);
  assert.equal(isXlsxFilename("REPORT.XLSX"), true);
  assert.equal(isXlsxFilename("report.xls"), false);
  assert.equal(isXlsxFilename("report.xlsm"), false);
  assert.equal(isXlsxFilename(null), false);
  await expectError(
    () => Reflect.apply(sanitizeExcel, undefined, [new Uint8Array()]),
    "ERR_INVALID_ARGUMENT",
  );
  await expectError(
    () => Reflect.apply(sanitizeExcel, undefined, [Buffer.alloc(0), { unknown: true }]),
    "ERR_INVALID_ARGUMENT",
  );
});

test("sanitizes BIFF8 XLS workbooks without converting their data model", async () => {
  const input = addXlsCompoundStorages(createXls());
  const output = await sanitizeExcel(input);
  const workbook = read(output, { type: "buffer" });
  const worksheet = workbook.Sheets["Sheet 1"];
  const cfb = CFB.read(output, { type: "buffer" });

  assert.equal(worksheet?.A1?.v, "kept value");
  assert.equal(worksheet?.A1?.c, undefined);
  assert.deepEqual(worksheet?.["!merges"], [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }]);
  assert.equal(
    cfb.FullPaths.some((path) => path.includes("ObjectPool")),
    false,
  );
  assert.equal(
    Buffer.from(CFB.find(cfb, "R/_VBA_PROJECT_CUR/VBA/dir")?.content ?? []).toString(),
    "legacy VBA project fixture",
  );
});

test("preserves the VBA project while sanitizing XLSM workbooks", async () => {
  const output = await sanitizeExcel(await createXlsm());
  const zip = await JSZip.loadAsync(output);

  assert.equal(await zip.file("xl/vbaProject.bin")?.async("string"), "test VBA project");
  assert.equal(zip.file("xl/drawings/drawing1.xml"), null);
});

test("applies worksheet and size limits to XLS workbooks", async () => {
  await expectError(
    () => sanitizeExcel(createXls(2), { requireSingleWorksheet: true }),
    "ERR_WORKSHEET_COUNT",
  );
  await expectError(
    () => sanitizeExcel(createXls(), { limits: { maxUncompressedBytes: 1024 } }),
    "ERR_LIMIT_EXCEEDED",
  );
  await expectError(() => sanitizeExcel(markXlsEncrypted(createXls())), "ERR_INVALID_XLS");
});
