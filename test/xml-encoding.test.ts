import assert from "node:assert/strict";
import test from "node:test";

import ExcelJS from "exceljs";
import JSZip from "jszip";

import { sanitizeExcel } from "../src/index.js";
import { createXlsx } from "./fixtures/workbook.js";
import { expectError, toArrayBuffer } from "./helpers.js";

for (const encoding of ["UTF-16LE", "UTF-16BE"] as const) {
  test(`preserves worksheet values and styles from ${encoding} XML`, async () => {
    const zip = await JSZip.loadAsync(await createXlsx());
    await Promise.all(
      ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml"].map(async (path) => {
        const xml = await zip.file(path)?.async("string");
        assert.ok(xml);
        const bytes = Buffer.from(xml.replace(/encoding="UTF-8"/i, 'encoding="UTF-16"'), "utf16le");
        const encoded = encoding === "UTF-16BE" ? bytes.swap16() : bytes;
        const bom = encoding === "UTF-16BE" ? Buffer.from([0xfe, 0xff]) : Buffer.from([0xff, 0xfe]);
        zip.file(path, Buffer.concat([bom, encoded]));
      }),
    );
    const output = await sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" }));
    const cleaned = new ExcelJS.Workbook();
    await cleaned.xlsx.load(toArrayBuffer(output));
    const sheet = cleaned.getWorksheet(1);
    assert.equal(sheet?.getCell("A1").value, "kept value");
    assert.equal(sheet?.getCell("A1").font.bold, true);
    assert.deepEqual(sheet?.getCell("B1").value, { formula: "SUM(A2:A3)", result: 3 });
    assert.equal(sheet?.getCell("E1").note, undefined);
    const outputZip = await JSZip.loadAsync(output);
    const xml = await outputZip.file("xl/worksheets/sheet1.xml")?.async("string");
    assert.ok(xml);
    assert.match(xml, /encoding="UTF-8"/);
    assert.equal(xml.includes("\u0000"), false);
    assert.equal(xml.includes("\ufffd"), false);
  });
}

test("rejects malformed text and inconsistent XML encodings", async () => {
  await Promise.all(
    [
      Buffer.from('<?xml version="1.0" encoding="UTF-16"?><worksheet/>'),
      Buffer.from([0xff, 0xfe, 0x3c]),
    ].map(async (contents) => {
      const zip = await JSZip.loadAsync(await createXlsx());
      zip.file("xl/worksheets/sheet1.xml", contents);
      await expectError(
        async () => sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" })),
        "ERR_INVALID_XLSX",
      );
    }),
  );
});
