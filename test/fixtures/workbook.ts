import * as CFB from "cfb";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { utils, write } from "@e965/xlsx";

const PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nWQAAAAASUVORK5CYII=";

export async function createXlsx(worksheetCount = 1): Promise<Buffer> {
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
      (_match: string, start: string, count: string, end: string) =>
        `${start}${Number(count) + 1}${end}`,
    )
    .replace(
      "</cellXfs>",
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf></cellXfs>',
    );
  zip.file("xl/styles.xml", withUnusedStyle);
  return zip.generateAsync({ type: "nodebuffer" });
}

export function createXls(worksheetCount = 1): Buffer {
  const workbook = utils.book_new();
  for (let index = 0; index < worksheetCount; index += 1) {
    const worksheet = utils.aoa_to_sheet([
      ["kept value", 1],
      [2, 3],
    ]);
    worksheet.A1.c = [{ a: "Test author", t: "drawing-backed comment" }];
    worksheet["!merges"] = [{ s: { r: 1, c: 0 }, e: { r: 1, c: 1 } }];
    utils.book_append_sheet(workbook, worksheet, `Sheet ${index + 1}`);
  }
  return Buffer.from(write(workbook, { bookType: "biff8", type: "buffer" }));
}

export function addXlsCompoundStorages(buffer: Buffer): Buffer {
  const cfb = CFB.read(buffer, { type: "buffer" });
  CFB.utils.cfb_add(
    cfb,
    "R/ObjectPool/MBD12345678/contents",
    Buffer.from("embedded object fixture"),
  );
  CFB.utils.cfb_add(cfb, "R/_VBA_PROJECT_CUR/VBA/dir", Buffer.from("legacy VBA project fixture"));
  return Buffer.from(CFB.write(cfb, { type: "buffer" }));
}

export function markXlsEncrypted(buffer: Buffer): Buffer {
  const cfb = CFB.read(buffer, { type: "buffer" });
  const workbook = CFB.find(cfb, "Workbook");
  if (!workbook) {
    throw new Error("Test workbook has no BIFF Workbook stream");
  }
  const stream = Buffer.from(workbook.content);
  const firstRecordEnd = 4 + stream.readUInt16LE(2);
  const filePassRecord = Buffer.from([0x2f, 0x00, 0x00, 0x00]);
  workbook.content = Buffer.concat([
    stream.subarray(0, firstRecordEnd),
    filePassRecord,
    stream.subarray(firstRecordEnd),
  ]);
  workbook.size = workbook.content.length;
  return Buffer.from(CFB.write(cfb, { type: "buffer" }));
}

export async function createXlsm(): Promise<Buffer> {
  const zip = await JSZip.loadAsync(await createXlsx());
  const relationshipsPath = "xl/_rels/workbook.xml.rels";
  const relationships = await zip.file(relationshipsPath)?.async("string");
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (!relationships || !contentTypes) {
    throw new Error("Test workbook is missing OOXML metadata");
  }

  zip.file("xl/vbaProject.bin", Buffer.from("test VBA project"));
  zip.file(
    relationshipsPath,
    relationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdVbaProject" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>',
    ),
  );
  zip.file(
    "[Content_Types].xml",
    contentTypes
      .replace(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
      )
      .replace(
        "</Types>",
        '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>',
      ),
  );
  return zip.generateAsync({ type: "nodebuffer" });
}

export async function usePairedRelationshipTags(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  for (const partName of ["xl/_rels/workbook.xml.rels", "xl/worksheets/_rels/sheet1.xml.rels"]) {
    const file = zip.file(partName);
    if (!file) {
      throw new Error(`Test workbook has no ${partName}`);
    }
    zip.file(
      partName,
      // The fixture mutates a single JSZip instance in deterministic order.
      // oxlint-disable-next-line no-await-in-loop
      (await file.async("string")).replace(
        /<Relationship\b([^>]*)\/>/g,
        "<Relationship$1></Relationship>",
      ),
    );
  }
  return zip.generateAsync({ type: "nodebuffer" });
}

export async function addThreadedCommentParts(buffer: Buffer): Promise<Buffer> {
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

export async function addEscapingRelationship(buffer: Buffer): Promise<Buffer> {
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

export async function addInvalidStyleReference(buffer: Buffer): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buffer);
  const worksheetPath = "xl/worksheets/sheet1.xml";
  const worksheet = await zip.file(worksheetPath)?.async("string");
  if (!worksheet) {
    throw new Error("Test workbook has no worksheet XML");
  }
  const corrupted = worksheet.replace(/<c\b[^>]*\br="A1"[^>]*>/, (tag) =>
    tag.replace(/\bs=(["'])\d+\1/, 's="999"'),
  );
  if (corrupted === worksheet) {
    throw new Error("Test workbook has no styled A1 cell");
  }
  zip.file(worksheetPath, corrupted);
  return zip.generateAsync({ type: "nodebuffer" });
}
