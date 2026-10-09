import assert from "node:assert/strict";
import test from "node:test";

import ExcelJS from "exceljs";
import JSZip from "jszip";

import { sanitizeExcel } from "../src/index.js";
import {
  removeThreadedCommentExtensions,
  removeWorksheetDrawingMarkup,
} from "../src/ooxml/cleanup.js";
import { parseRelationships, removeRelationships } from "../src/ooxml/relationships.js";
import { collectStyleReferences, remapWorksheetStyles } from "../src/ooxml/styles.js";
import { decodeXml, findElements, getAttribute, replaceCollection } from "../src/ooxml/xml.js";
import { createXlsx } from "./fixtures/workbook.js";
import { cellXfCount, expectError, toArrayBuffer } from "./helpers.js";

const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

test("XML cleanup preserves literal markup in CDATA, comments, and attribute values", () => {
  const literal = '<drawing>kept</drawing><c s="999"/><col style="888"/>';
  const xml = `<worksheet><sheetData><row s="2"><c s="1" note='fake s="999"'><is><t><![CDATA[${literal}]]></t></is></c></row></sheetData><!-- ${literal} --><drawing note="a > b"></drawing><extLst><ext uri="keep"><![CDATA[<threadedComments/>]]><!-- <threadedComments/> --></ext><ext uri="remove"><threadedComments/></ext></extLst></worksheet>`;
  const cleaned = removeThreadedCommentExtensions(removeWorksheetDrawingMarkup(xml));
  assert.match(cleaned, /<!\[CDATA\[<drawing>kept<\/drawing><c s="999"\/><col style="888"\/>\]\]>/);
  assert.ok(cleaned.includes(`<!-- ${literal} -->`));
  assert.ok(cleaned.includes('<ext uri="keep">'));
  assert.ok(!cleaned.includes('uri="remove"'));
  assert.ok(!cleaned.includes('<drawing note="a > b">'));

  const references = new Set<number>();
  collectStyleReferences(cleaned, references);
  assert.deepEqual(references, new Set([2, 1]));
  const remapped = remapWorksheetStyles(
    cleaned,
    new Map([
      [1, 0],
      [2, 1],
    ]),
  );
  assert.ok(remapped.includes(`<c s="0" note='fake s="999"'>`));
  assert.ok(remapped.includes(`<![CDATA[${literal}]]>`));
  assert.ok(remapped.includes(`<!-- ${literal} -->`));
  const unrelated =
    '<worksheet><extLst><ext uri="keep"><foreign:threadedComments xmlns:foreign="urn:test"/></ext></extLst></worksheet>';
  assert.equal(removeThreadedCommentExtensions(unrelated), unrelated);
});

test("relationship parsing and removal ignore literal markup", () => {
  const literal = '<Relationship Id="fake" Type="urn:test/drawing" Target="../must-keep.xml"/>';
  const xml = `<Relationships><!-- ${literal} --><![CDATA[${literal}]]><Relationship Id="real" Type="urn:test/drawing" Target="drawings/a.xml"/></Relationships>`;
  assert.deepEqual(
    parseRelationships(xml, "xl/workbook.xml").map((relationship) => relationship.id),
    ["real"],
  );
  const removed = removeRelationships(xml, "xl/workbook.xml", ["/drawing"]);
  assert.deepEqual(removed.targets, new Set(["xl/drawings/a.xml"]));
  assert.equal(
    removed.xml,
    `<Relationships><!-- ${literal} --><![CDATA[${literal}]]></Relationships>`,
  );
  assert.throws(
    () => parseRelationships('<Relationship Id="bad" Target="sheet.xml"/>', "xl/workbook.xml"),
    /Invalid relationship/,
  );
  assert.throws(
    () =>
      parseRelationships(
        '<Relationship Id="a" Type="urn:test/worksheet" Target="a.xml"/><Relationship Id="a" Type="urn:test/worksheet" Target="b.xml"/>',
        "xl/workbook.xml",
      ),
    /Duplicate relationship ID/,
  );
});

test("attribute parsing respects quotes and decodes character references once", () => {
  assert.equal(getAttribute('<c note="x > y" s="&#49;"/>', "s"), "1");
  assert.equal(getAttribute('Id="relationship"', "Id"), "relationship");
  assert.equal(decodeXml("&amp;lt;&#x41;&#65;"), "&lt;AA");
  assert.throws(() => decodeXml("&constructor;"), /Invalid XML character reference/);
  assert.throws(() => decodeXml("&toString;"), /Invalid XML character reference/);
  assert.equal(
    replaceCollection(
      '<styleSheet><!-- <cellXfs count="999"/> --><cellXfs count="2"><xf/><xf/></cellXfs></styleSheet>',
      "cellXfs",
      "<xf/>",
      1,
    ),
    '<styleSheet><!-- <cellXfs count="999"/> --><cellXfs count="1"><xf/></cellXfs></styleSheet>',
  );
  assert.throws(
    () => removeWorksheetDrawingMarkup("<worksheet><drawing></worksheet>"),
    /Mismatched XML/,
  );
  assert.equal(
    replaceCollection('<styleSheet><cellXfs count="0"/></styleSheet>', "cellXfs", "<xf/>", 1),
    '<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>',
  );
  assert.throws(
    () =>
      removeWorksheetDrawingMarkup(
        '<worksheet xmlns:a="urn:test" xmlns:b="urn:test" a:id="1" b:id="2"/>',
      ),
    /Duplicate XML attribute/,
  );
});

test("supports inherited prefixes in styles XML while ignoring comment styles", async () => {
  const zip = await JSZip.loadAsync(await createXlsx());
  const styles = await zip.file("xl/styles.xml")?.async("string");
  assert.ok(styles);
  const prefixed = styles
    .replace(/(<\/?)([A-Za-z][\w]*)(?=[\s/>])/g, "$1main:$2")
    .replace(
      "<main:styleSheet",
      '<main:styleSheet xmlns:main="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    )
    .replace("<main:cellXfs", '<!-- <cellXfs count="999"><xf/></cellXfs> --><main:cellXfs');
  zip.file("xl/styles.xml", prefixed);
  const output = await sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" }));
  const cleaned = await JSZip.loadAsync(output);
  const cleanedStyles = await cleaned.file("xl/styles.xml")?.async("string");
  assert.ok(cleanedStyles);
  assert.ok(cleanedStyles.includes('<!-- <cellXfs count="999"><xf/></cellXfs> -->'));
  const stylesCollection = findElements(cleanedStyles, "cellXfs")[0];
  assert.ok(stylesCollection);
  assert.equal(
    getAttribute(cleanedStyles.slice(stylesCollection.start, stylesCollection.openingEnd), "count"),
    String(cellXfCount(styles) - 1),
  );
});

for (const scope of ["workbook", "sheets", "sheet"]) {
  test(`resolves renamed relationship prefixes declared on ${scope}`, async () => {
    const zip = await JSZip.loadAsync(await createXlsx());
    const workbook = await zip.file("xl/workbook.xml")?.async("string");
    const originalStyles = await zip.file("xl/styles.xml")?.async("string");
    assert.ok(workbook);
    assert.ok(originalStyles);
    const renamed = workbook
      .replace(` xmlns:r="${RELATIONSHIP_NAMESPACE}"`, "")
      .replaceAll("r:id=", "rel:id=")
      .replace(new RegExp(`<${scope}(?=[\\s>])`), `<${scope} xmlns:rel="${RELATIONSHIP_NAMESPACE}"`)
      .replace("<sheets", '<!-- <sheet name="fake" r:id="fake"/> --><sheets');
    zip.file("xl/workbook.xml", renamed);
    const output = await sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" }), {
      requireSingleWorksheet: true,
    });
    const cleaned = await JSZip.loadAsync(output);
    const worksheet = await cleaned.file("xl/worksheets/sheet1.xml")?.async("string");
    const styles = await cleaned.file("xl/styles.xml")?.async("string");
    assert.ok(worksheet);
    assert.ok(styles);
    assert.doesNotMatch(worksheet, /<(?:drawing|legacyDrawing)\b/);
    assert.ok(
      Object.keys(cleaned.files).every(
        (name) => cleaned.files[name]?.dir || !name.startsWith("xl/drawings/"),
      ),
    );
    assert.equal(cellXfCount(styles), cellXfCount(originalStyles) - 1);
    const outputStyle = getAttribute(worksheet.match(/<c\b[^>]*r="A1"[^>]*>/)?.[0] ?? "", "s");
    assert.ok(outputStyle);
    const sourceXfs = findElements(originalStyles, "xf").filter(
      (element) => element.parent?.localName === "cellXfs",
    );
    const outputXfs = findElements(styles, "xf").filter(
      (element) => element.parent?.localName === "cellXfs",
    );
    const sourceXf = sourceXfs[Number(outputStyle)];
    const outputXf = outputXfs[Number(outputStyle)];
    assert.ok(sourceXf);
    assert.ok(outputXf);
    assert.equal(
      styles.slice(outputXf.start, outputXf.end),
      originalStyles.slice(sourceXf.start, sourceXf.end),
    );
  });
}

test("preserves literal inline cell markup while sanitizing an OOXML workbook", async () => {
  const zip = await JSZip.loadAsync(await createXlsx());
  const worksheet = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
  assert.ok(worksheet);
  const text = '<drawing>kept value</drawing><c s="999"/>';
  zip.file(
    "xl/worksheets/sheet1.xml",
    worksheet.replace(
      "</row>",
      `<c r="F1" t="inlineStr"><is><t><![CDATA[${text}]]></t></is></c></row>`,
    ),
  );
  const output = await sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" }));
  const cleaned = await JSZip.loadAsync(output);
  const cleanedWorksheet = await cleaned.file("xl/worksheets/sheet1.xml")?.async("string");
  assert.ok(cleanedWorksheet?.includes(`<![CDATA[${text}]]>`));
  const excel = new ExcelJS.Workbook();
  await excel.xlsx.load(toArrayBuffer(output));
  assert.equal(excel.getWorksheet(1)?.getCell("A1").font.bold, true);
});

for (const mutation of ["missing ID", "wrong namespace", "missing target", "external target"]) {
  test(`rejects a worksheet with ${mutation} instead of silently skipping cleanup`, async () => {
    const zip = await JSZip.loadAsync(await createXlsx());
    const workbook = await zip.file("xl/workbook.xml")?.async("string");
    const relationships = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    assert.ok(workbook);
    assert.ok(relationships);
    if (mutation === "missing ID") {
      zip.file("xl/workbook.xml", workbook.replace(/r:id="[^"]+"/, 'r:id="missing"'));
    } else if (mutation === "wrong namespace") {
      zip.file("xl/workbook.xml", workbook.replace(RELATIONSHIP_NAMESPACE, "urn:wrong"));
    } else if (mutation === "missing target") {
      zip.file(
        "xl/_rels/workbook.xml.rels",
        relationships.replace(/<Relationship\b[^>]*Type="[^"]*\/worksheet"[^>]*\/\s*>/, ""),
      );
    } else {
      zip.file(
        "xl/_rels/workbook.xml.rels",
        relationships.replace(
          /(<Relationship\b[^>]*Type="[^"]*\/worksheet"[^>]*)(\/\s*>)/,
          '$1 TargetMode="External"$2',
        ),
      );
    }
    await expectError(
      async () => sanitizeExcel(await zip.generateAsync({ type: "nodebuffer" })),
      "ERR_INVALID_XLSX",
    );
  });
}
