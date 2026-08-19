import assert from "node:assert/strict";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";

import ExcelJS from "exceljs";
import JSZip from "jszip";
import { read as readSpreadsheet } from "@e965/xlsx";

import { sanitizeExcelFile } from "../src/index.js";
import { createXls, createXlsm, createXlsx } from "./fixtures/workbook.js";
import { expectError, toArrayBuffer } from "./helpers.js";

test("file API writes safely, refuses replacement by default, and supports overwrite", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-test-"));
  const inputPath = join(directory, "input.XLSX");
  const outputPath = join(directory, "output.xlsx");
  try {
    await writeFile(inputPath, await createXlsx());
    await sanitizeExcelFile(inputPath, outputPath);
    const firstOutput = await readFile(outputPath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(toArrayBuffer(firstOutput));
    assert.equal(workbook.getWorksheet(1)?.getCell("A1").value, "kept value");

    await expectError(() => sanitizeExcelFile(inputPath, outputPath), "ERR_OUTPUT_EXISTS");
    await sanitizeExcelFile(inputPath, outputPath, { overwrite: true });
    assert.ok((await readFile(outputPath)).length > 0);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("file API supports matching XLS and XLSM paths", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-formats-"));
  const xlsInputPath = join(directory, "input.XLS");
  const xlsOutputPath = join(directory, "output.xls");
  const xlsmInputPath = join(directory, "input.xlsm");
  const xlsmOutputPath = join(directory, "output.XLSM");
  try {
    await writeFile(xlsInputPath, createXls());
    await sanitizeExcelFile(xlsInputPath, xlsOutputPath);
    assert.equal(
      readSpreadsheet(await readFile(xlsOutputPath), { type: "buffer" }).Sheets["Sheet 1"]?.A1?.v,
      "kept value",
    );

    await writeFile(xlsmInputPath, await createXlsm());
    await sanitizeExcelFile(xlsmInputPath, xlsmOutputPath);
    const xlsm = await JSZip.loadAsync(await readFile(xlsmOutputPath));
    assert.equal(await xlsm.file("xl/vbaProject.bin")?.async("string"), "test VBA project");
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("file API exposes path, extension, read, and write errors", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-errors-"));
  const inputPath = join(directory, "input.xlsx");
  const disguisedXlsPath = join(directory, "disguised.xlsx");
  const disguisedXlsxPath = join(directory, "disguised.xls");
  try {
    await writeFile(inputPath, await createXlsx());
    await writeFile(disguisedXlsPath, createXls());
    await writeFile(disguisedXlsxPath, await createXlsx());
    await expectError(() => sanitizeExcelFile(inputPath, inputPath), "ERR_SAME_PATH");
    await expectError(
      () => sanitizeExcelFile(join(directory, "input.xls"), join(directory, "out.xlsx")),
      "ERR_UNSUPPORTED_EXTENSION",
    );
    await expectError(
      () => sanitizeExcelFile(join(directory, "input.xlsm"), join(directory, "out.xlsx")),
      "ERR_UNSUPPORTED_EXTENSION",
    );
    await expectError(
      () => sanitizeExcelFile(join(directory, "missing.xlsx"), join(directory, "out.xlsx")),
      "ERR_FILE_READ",
    );
    await expectError(
      () => sanitizeExcelFile(disguisedXlsPath, join(directory, "disguised-output.xlsx")),
      "ERR_INVALID_XLSX",
    );
    await expectError(
      () => sanitizeExcelFile(disguisedXlsxPath, join(directory, "disguised-output.xls")),
      "ERR_INVALID_XLS",
    );
    await expectError(
      () => sanitizeExcelFile(inputPath, join(directory, "missing-directory", "out.xlsx")),
      "ERR_FILE_WRITE",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
