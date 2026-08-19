import assert from "node:assert/strict";
import test from "node:test";

import ExcelJS from "exceljs";
import JSZip from "jszip";

import { sanitizeExcel } from "../src/index.js";
import {
  addInvalidStyleReference,
  addThreadedCommentParts,
  createXlsx,
  usePairedRelationshipTags,
} from "./fixtures/workbook.js";
import { cellXfCount, expectError, toArrayBuffer } from "./helpers.js";

test("removes drawing content and unused styles while preserving sheet semantics", async () => {
  const input = await createXlsx();
  const inputCopy = Buffer.from(input);
  const inputZip = await JSZip.loadAsync(input);
  const inputStyles = await inputZip.file("xl/styles.xml")?.async("string");
  assert.ok(inputStyles);

  const output = await sanitizeExcel(input, { requireSingleWorksheet: true });
  assert.deepEqual(input, inputCopy);

  const outputZip = await JSZip.loadAsync(output);
  const fileNames = Object.keys(outputZip.files);
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/drawings/") && !outputZip.files[name]?.dir),
    false,
  );
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/media/") && !outputZip.files[name]?.dir),
    false,
  );
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/comments")),
    false,
  );

  const worksheetXml = await outputZip.file("xl/worksheets/sheet1.xml")?.async("string");
  const worksheetRels = await outputZip
    .file("xl/worksheets/_rels/sheet1.xml.rels")
    ?.async("string");
  assert.ok(worksheetXml);
  assert.ok(worksheetRels);
  assert.doesNotMatch(worksheetXml, /<(?:\w+:)?(?:drawing|legacyDrawing)\b/i);
  assert.doesNotMatch(worksheetRels, /\/(?:comments|drawing|image|vmlDrawing)"/i);

  const outputStyles = await outputZip.file("xl/styles.xml")?.async("string");
  assert.ok(outputStyles);
  assert.ok(cellXfCount(outputStyles) < cellXfCount(inputStyles));

  const cleanedWorkbook = new ExcelJS.Workbook();
  await cleanedWorkbook.xlsx.load(toArrayBuffer(output));
  const sheet = cleanedWorkbook.getWorksheet(1);
  assert.equal(sheet?.getCell("A1").value, "kept value");
  assert.equal(sheet?.getCell("A1").font.bold, true);
  assert.deepEqual(sheet?.getCell("B1").value, {
    formula: "SUM(A2:A3)",
    result: 3,
  });
  assert.equal(sheet?.getCell("C1").value, "merged");
  assert.equal(sheet?.getCell("D1").isMerged, true);
  assert.equal(sheet?.getCell("E1").note, undefined);
});

test("supports paired relationship tags", async () => {
  const output = await sanitizeExcel(await usePairedRelationshipTags(await createXlsx()));
  const zip = await JSZip.loadAsync(output);
  const fileNames = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir);
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/drawings/")),
    false,
  );
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/comments")),
    false,
  );
  const cleanedWorkbook = new ExcelJS.Workbook();
  await cleanedWorkbook.xlsx.load(toArrayBuffer(output));
  assert.equal(cleanedWorkbook.getWorksheet(1)?.getCell("A1").font.bold, true);
});

test("removes threaded comments and person metadata but preserves other extensions", async () => {
  const output = await sanitizeExcel(await addThreadedCommentParts(await createXlsx()));
  const zip = await JSZip.loadAsync(output);
  const fileNames = Object.keys(zip.files).filter((name) => !zip.files[name]?.dir);
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/threadedComments/")),
    false,
  );
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/persons/")),
    false,
  );
  const worksheet = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
  const workbookRels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  assert.ok(worksheet);
  assert.ok(workbookRels);
  assert.doesNotMatch(worksheet, /threadedComments/);
  assert.match(worksheet, /keep:marker/);
  assert.doesNotMatch(workbookRels, /\/person"/i);
});

test("rejects invalid worksheet style indexes", async () => {
  await expectError(
    async () => sanitizeExcel(await addInvalidStyleReference(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
});
