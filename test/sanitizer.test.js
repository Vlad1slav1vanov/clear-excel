import assert from "node:assert/strict";
import { createRequire } from "node:module";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import {
  DEFAULT_LIMITS,
  ExcelSanitizationError,
  isXlsxFilename,
  sanitizeExcel,
  sanitizeExcelFile,
} from "../index.js";

const require = createRequire(import.meta.url);
const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nWQAAAAASUVORK5CYII=";

async function expectError(operation, code) {
  await assert.rejects(operation, (error) => {
    assert.ok(error instanceof ExcelSanitizationError);
    assert.equal(error.code, code);
    return true;
  });
}

async function createXlsx(worksheetCount = 1) {
  const workbook = new ExcelJS.Workbook();
  for (let index = 0; index < worksheetCount; index += 1) {
    const worksheet = workbook.addWorksheet(`Sheet ${index + 1}`);
    worksheet.getCell("A1").value = "kept value";
    worksheet.getCell("A1").font = {
      bold: true,
      color: { argb: "FFFF0000" },
    };
    worksheet.getCell("A2").value = 1;
    worksheet.getCell("A3").value = 2;
    worksheet.getCell("B1").value = { formula: "SUM(A2:A3)", result: 3 };
    worksheet.mergeCells("C1:D1");
    worksheet.getCell("C1").value = "merged";
    worksheet.getCell("E1").note = "drawing-backed comment";
  }

  const imageId = workbook.addImage({ base64: PNG_BASE64, extension: "png" });
  workbook.getWorksheet(1)?.addImage(imageId, {
    tl: { col: 5, row: 0 },
    ext: { width: 16, height: 16 },
  });

  const zip = await JSZip.loadAsync(Buffer.from(await workbook.xlsx.writeBuffer()));
  const stylesFile = zip.file("xl/styles.xml");
  if (!stylesFile) {
    throw new Error("Test workbook has no styles part");
  }
  const styles = await stylesFile.async("string");
  const withUnusedStyle = styles
    .replace(
      /(<cellXfs\b[^>]*\bcount=")(\d+)(")/,
      (_match, start, count, end) => `${start}${Number(count) + 1}${end}`,
    )
    .replace(
      "</cellXfs>",
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>',
    );
  zip.file("xl/styles.xml", withUnusedStyle);
  return zip.generateAsync({ type: "nodebuffer" });
}

function cellXfCount(styles) {
  const count = styles.match(/<cellXfs\b[^>]*\bcount="(\d+)"/)?.[1];
  if (!count) {
    throw new Error("styles.xml has no cellXfs count");
  }
  return Number(count);
}

async function usePairedRelationshipTags(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  for (const partName of [
    "xl/_rels/workbook.xml.rels",
    "xl/worksheets/_rels/sheet1.xml.rels",
  ]) {
    const file = zip.file(partName);
    if (!file) {
      throw new Error(`Test workbook has no ${partName}`);
    }
    zip.file(
      partName,
      (await file.async("string")).replace(
        /<Relationship\b([^>]*)\/>/g,
        "<Relationship$1></Relationship>",
      ),
    );
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

async function addThreadedCommentParts(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const worksheetPath = "xl/worksheets/sheet1.xml";
  const worksheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const worksheet = await zip.file(worksheetPath)?.async("string");
  const worksheetRels = await zip.file(worksheetRelsPath)?.async("string");
  const workbookRels = await zip.file(workbookRelsPath)?.async("string");
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (!worksheet || !worksheetRels || !workbookRels || !contentTypes) {
    throw new Error("Test workbook is missing required OOXML parts");
  }

  zip.file(
    worksheetPath,
    worksheet.replace(
      "</worksheet>",
      '<extLst><ext uri="{threaded-comment}"><x15:threadedComments xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdThreaded"/></ext><ext uri="{keep-me}"><keep:marker xmlns:keep="urn:test:keep"/></ext></extLst></worksheet>',
    ),
  );
  zip.file(
    worksheetRelsPath,
    worksheetRels.replace(
      "</Relationships>",
      '<Relationship Id="rIdThreaded" Type="http://schemas.microsoft.com/office/2017/10/relationships/threadedComment" Target="../threadedComments/threadedComment1.xml"/></Relationships>',
    ),
  );
  zip.file(
    workbookRelsPath,
    workbookRels.replace(
      "</Relationships>",
      '<Relationship Id="rIdPerson" Type="http://schemas.microsoft.com/office/2017/10/relationships/person" Target="persons/person.xml"/></Relationships>',
    ),
  );
  zip.file(
    "xl/threadedComments/threadedComment1.xml",
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A1" personId="person-1" id="comment-1"><text>remove me</text></threadedComment></ThreadedComments>',
  );
  zip.file(
    "xl/persons/person.xml",
    '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="Author" id="person-1"/></personList>',
  );
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(
      "</Types>",
      '<Override PartName="/xl/threadedComments/threadedComment1.xml" ContentType="application/vnd.ms-excel.threadedcomments+xml"/><Override PartName="/xl/persons/person.xml" ContentType="application/vnd.ms-excel.person+xml"/></Types>',
    ),
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

function findEndRecord(buffer) {
  for (let offset = buffer.length - 22; offset >= 0; offset -= 1) {
    if (buffer.readUInt32LE(offset) === 0x06054b50) {
      return offset;
    }
  }
  throw new Error("Test ZIP has no end-of-central-directory record");
}

function inflateDeclaredZipSize(buffer) {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const firstEntryOffset = manipulated.readUInt32LE(endOffset + 16);
  manipulated.writeUInt32LE(600 * 1024 * 1024, firstEntryOffset + 24);
  return manipulated;
}

function forgeZipEntrySizes(buffer, targetName, declaredSize) {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const entryCount = manipulated.readUInt16LE(endOffset + 10);
  let entryOffset = manipulated.readUInt32LE(endOffset + 16);
  for (let index = 0; index < entryCount; index += 1) {
    const fileNameLength = manipulated.readUInt16LE(entryOffset + 28);
    const extraLength = manipulated.readUInt16LE(entryOffset + 30);
    const commentLength = manipulated.readUInt16LE(entryOffset + 32);
    const fileName = manipulated
      .subarray(entryOffset + 46, entryOffset + 46 + fileNameLength)
      .toString("utf8");
    if (fileName === targetName) {
      const localHeaderOffset = manipulated.readUInt32LE(entryOffset + 42);
      manipulated.writeUInt32LE(declaredSize, entryOffset + 24);
      manipulated.writeUInt32LE(declaredSize, localHeaderOffset + 22);
      return manipulated;
    }
    entryOffset += 46 + fileNameLength + extraLength + commentLength;
  }
  throw new Error(`Test ZIP entry was not found: ${targetName}`);
}

async function createRuntimeXmlBomb() {
  const relationshipPath = "xl/_rels/workbook.xml.rels";
  const zip = await JSZip.loadAsync(await createXlsx());
  const relationships = await zip.file(relationshipPath)?.async("string");
  if (!relationships) {
    throw new Error("Test workbook has no workbook relationships");
  }
  zip.file(
    relationshipPath,
    relationships.replace(
      "</Relationships>",
      `${" ".repeat(9 * 1024 * 1024)}</Relationships>`,
    ),
  );
  const compressed = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  return forgeZipEntrySizes(compressed, relationshipPath, 1);
}

function markAsMultiDisk(buffer) {
  const manipulated = Buffer.from(buffer);
  manipulated.writeUInt16LE(1, findEndRecord(manipulated) + 4);
  return manipulated;
}

function markAsZip64(buffer) {
  const manipulated = Buffer.from(buffer);
  manipulated.writeUInt16LE(0xffff, findEndRecord(manipulated) + 10);
  return manipulated;
}

function addUnsafeCentralDirectoryName(buffer) {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const entryOffset = manipulated.readUInt32LE(endOffset + 16);
  const fileNameLength = manipulated.readUInt16LE(entryOffset + 28);
  const unsafeName = Buffer.from(`../${"x".repeat(fileNameLength - 3)}`);
  unsafeName.copy(manipulated, entryOffset + 46);
  return manipulated;
}

async function addEscapingRelationship(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const relationshipPath = "xl/worksheets/_rels/sheet1.xml.rels";
  const relationships = await zip.file(relationshipPath)?.async("string");
  if (!relationships) {
    throw new Error("Test workbook has no worksheet relationships");
  }
  zip.file(
    relationshipPath,
    relationships.replace(
      /Target="\.\.\/drawings\/drawing1\.xml"/,
      'Target="../../../../escape.xml"',
    ),
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

async function addInvalidStyleReference(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const worksheetPath = "xl/worksheets/sheet1.xml";
  const worksheet = await zip.file(worksheetPath)?.async("string");
  if (!worksheet) {
    throw new Error("Test workbook has no worksheet XML");
  }
  const corrupted = worksheet.replace(
    /<c\b[^>]*\br="A1"[^>]*>/,
    (tag) => tag.replace(/\bs=(["'])\d+\1/, 's="999"'),
  );
  if (corrupted === worksheet) {
    throw new Error("Test workbook has no styled A1 cell");
  }
  zip.file(worksheetPath, corrupted);
  return zip.generateAsync({ type: "nodebuffer" });
}

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
    fileNames.some(
      (name) => name.startsWith("xl/drawings/") && !outputZip.files[name].dir,
    ),
    false,
  );
  assert.equal(
    fileNames.some(
      (name) => name.startsWith("xl/media/") && !outputZip.files[name].dir,
    ),
    false,
  );
  assert.equal(fileNames.some((name) => name.startsWith("xl/comments")), false);

  const worksheetXml = await outputZip
    .file("xl/worksheets/sheet1.xml")
    ?.async("string");
  const worksheetRels = await outputZip
    .file("xl/worksheets/_rels/sheet1.xml.rels")
    ?.async("string");
  assert.doesNotMatch(worksheetXml, /<(?:\w+:)?(?:drawing|legacyDrawing)\b/i);
  assert.doesNotMatch(worksheetRels, /\/(?:comments|drawing|image|vmlDrawing)"/i);

  const outputStyles = await outputZip.file("xl/styles.xml")?.async("string");
  assert.ok(outputStyles);
  assert.ok(cellXfCount(outputStyles) < cellXfCount(inputStyles));

  const cleanedWorkbook = new ExcelJS.Workbook();
  await cleanedWorkbook.xlsx.load(output);
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
  const output = await sanitizeExcel(
    await usePairedRelationshipTags(await createXlsx()),
  );
  const zip = await JSZip.loadAsync(output);
  const fileNames = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  assert.equal(fileNames.some((name) => name.startsWith("xl/drawings/")), false);
  assert.equal(fileNames.some((name) => name.startsWith("xl/comments")), false);
  const cleanedWorkbook = new ExcelJS.Workbook();
  await cleanedWorkbook.xlsx.load(output);
  assert.equal(cleanedWorkbook.getWorksheet(1)?.getCell("A1").font.bold, true);
});

test("removes threaded comments and person metadata but preserves other extensions", async () => {
  const output = await sanitizeExcel(
    await addThreadedCommentParts(await createXlsx()),
  );
  const zip = await JSZip.loadAsync(output);
  const fileNames = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  assert.equal(
    fileNames.some((name) => name.startsWith("xl/threadedComments/")),
    false,
  );
  assert.equal(fileNames.some((name) => name.startsWith("xl/persons/")), false);
  const worksheet = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
  const workbookRels = await zip
    .file("xl/_rels/workbook.xml.rels")
    ?.async("string");
  assert.doesNotMatch(worksheet, /threadedComments/);
  assert.match(worksheet, /keep:marker/);
  assert.doesNotMatch(workbookRels, /\/person"/i);
});

test("rejects declared and actual XML expansion bombs", async () => {
  await expectError(
    async () => sanitizeExcel(inflateDeclaredZipSize(await createXlsx())),
    "ERR_LIMIT_EXCEEDED",
  );
  await expectError(
    async () => sanitizeExcel(await createRuntimeXmlBomb()),
    "ERR_LIMIT_EXCEEDED",
  );
});

test("supports safe custom limits and validates limit options", async () => {
  assert.equal(Object.isFrozen(DEFAULT_LIMITS), true);
  await expectError(
    async () =>
      sanitizeExcel(await createXlsx(), { limits: { maxEntryCount: 1 } }),
    "ERR_LIMIT_EXCEEDED",
  );
  await expectError(
    async () =>
      sanitizeExcel(await createXlsx(), { limits: { maxEntryCount: 0 } }),
    "ERR_INVALID_ARGUMENT",
  );
  await expectError(
    async () => sanitizeExcel(await createXlsx(), { limits: { typo: 100 } }),
    "ERR_INVALID_ARGUMENT",
  );
});

test("enforces the optional single-sheet contract", async () => {
  await expectError(
    async () =>
      sanitizeExcel(await createXlsx(2), { requireSingleWorksheet: true }),
    "ERR_WORKSHEET_COUNT",
  );
  await sanitizeExcel(await createXlsx(2));
});

test("rejects malformed and unsupported ZIP structures", async () => {
  await assert.rejects(
    () => sanitizeExcel(Buffer.from("not an OOXML workbook")),
    (error) => {
      assert.equal(error.code, "ERR_INVALID_XLSX");
      assert.ok(error.cause instanceof Error);
      return true;
    },
  );
  await expectError(
    async () => sanitizeExcel(markAsMultiDisk(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
  await expectError(
    async () => sanitizeExcel(markAsZip64(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
});

test("rejects unsafe package paths and relationship targets", async () => {
  await expectError(
    async () =>
      sanitizeExcel(addUnsafeCentralDirectoryName(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
  await expectError(
    async () => sanitizeExcel(await addEscapingRelationship(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
});

test("rejects invalid worksheet style indexes", async () => {
  await expectError(
    async () => sanitizeExcel(await addInvalidStyleReference(await createXlsx())),
    "ERR_INVALID_XLSX",
  );
});

test("validates public arguments and XLSX filenames", async () => {
  assert.equal(isXlsxFilename("report.xlsx"), true);
  assert.equal(isXlsxFilename("REPORT.XLSX"), true);
  assert.equal(isXlsxFilename("report.xls"), false);
  assert.equal(isXlsxFilename("report.xlsm"), false);
  assert.equal(isXlsxFilename(null), false);
  await expectError(() => sanitizeExcel(new Uint8Array()), "ERR_INVALID_ARGUMENT");
  await expectError(
    () => sanitizeExcel(Buffer.alloc(0), { unknown: true }),
    "ERR_INVALID_ARGUMENT",
  );
});

test("file API writes safely, refuses replacement by default, and supports overwrite", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-test-"));
  const inputPath = join(directory, "input.XLSX");
  const outputPath = join(directory, "output.xlsx");
  try {
    await writeFile(inputPath, await createXlsx());
    await sanitizeExcelFile(inputPath, outputPath);
    const firstOutput = await readFile(outputPath);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(firstOutput);
    assert.equal(workbook.getWorksheet(1)?.getCell("A1").value, "kept value");

    await expectError(
      () => sanitizeExcelFile(inputPath, outputPath),
      "ERR_OUTPUT_EXISTS",
    );
    await sanitizeExcelFile(inputPath, outputPath, { overwrite: true });
    assert.ok((await readFile(outputPath)).length > 0);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("file API exposes path, extension, read, and write errors", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-errors-"));
  const inputPath = join(directory, "input.xlsx");
  try {
    await writeFile(inputPath, await createXlsx());
    await expectError(
      () => sanitizeExcelFile(inputPath, inputPath),
      "ERR_SAME_PATH",
    );
    await expectError(
      () => sanitizeExcelFile(join(directory, "input.xls"), join(directory, "out.xlsx")),
      "ERR_UNSUPPORTED_EXTENSION",
    );
    await expectError(
      () => sanitizeExcelFile(join(directory, "missing.xlsx"), join(directory, "out.xlsx")),
      "ERR_FILE_READ",
    );
    await expectError(
      () =>
        sanitizeExcelFile(
          inputPath,
          join(directory, "missing-directory", "out.xlsx"),
        ),
      "ERR_FILE_WRITE",
    );
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test("ESM and CommonJS entry points expose the same implementation", () => {
  const commonJsApi = require("../index.cjs");
  assert.equal(commonJsApi.sanitizeExcel, sanitizeExcel);
  assert.equal(commonJsApi.sanitizeExcelFile, sanitizeExcelFile);
  assert.equal(commonJsApi.ExcelSanitizationError, ExcelSanitizationError);
  assert.equal(commonJsApi.DEFAULT_LIMITS, DEFAULT_LIMITS);
});
