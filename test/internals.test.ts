import assert from "node:assert/strict";
import test from "node:test";

import {
  collectReachableParts,
  parseRelationships,
  removeRelationships,
  resolveTarget,
  type RelationshipGraph,
} from "../src/ooxml/relationships.js";
import { collectStyleReferences, remapWorksheetStyles } from "../src/ooxml/styles.js";
import { validatePackageEntryName } from "../src/zip/central-directory.js";

test("ZIP entry and relationship helpers reject package escapes", () => {
  assert.throws(() => validatePackageEntryName("../escape.xml"), /Unsafe ZIP entry name/);
  assert.throws(() => validatePackageEntryName("C:/escape.xml"), /Unsafe ZIP entry name/);
  assert.throws(() => resolveTarget("xl/workbook.xml", "../../escape.xml"), /escapes package/);
  assert.equal(
    resolveTarget("xl/workbook.xml", "worksheets/sheet1.xml"),
    "xl/worksheets/sheet1.xml",
  );
});

test("relationship helpers handle internal, external, and paired tags", () => {
  const xml = [
    '<Relationship Id="internal" Type="urn:test/drawing" Target="drawings/a.xml"></Relationship>',
    '<Relationship Id="external" Type="urn:test/drawing" Target="https://example.com/a" TargetMode="External"/>',
  ].join("");
  const parsed = parseRelationships(xml, "xl/workbook.xml");
  assert.equal(parsed[0]?.target, "xl/drawings/a.xml");
  assert.equal(parsed[1]?.target, "https://example.com/a");

  const removed = removeRelationships(xml, "xl/workbook.xml", ["/drawing"]);
  assert.equal(removed.targets.has("xl/drawings/a.xml"), true);
  assert.match(removed.xml, /TargetMode="External"/);
  assert.doesNotMatch(removed.xml, /Id="internal"/);
});

test("reachable-part traversal excludes external relationships", () => {
  const graph: RelationshipGraph = new Map([
    [
      "",
      [
        { id: "root", target: "xl/workbook.xml", targetMode: undefined, type: "root" },
        { id: "external", target: "https://example.com", targetMode: "External", type: "link" },
      ],
    ],
  ]);
  assert.deepEqual(collectReachableParts(new Set([""]), graph), new Set(["", "xl/workbook.xml"]));
});

test("style helpers collect and remap cell, row, and column references", () => {
  const xml =
    '<worksheet><cols><col style="3"/></cols><sheetData><row s="2"><c s="1"/></row></sheetData></worksheet>';
  const references = new Set<number>();
  collectStyleReferences(xml, references);
  assert.deepEqual(references, new Set([2, 1, 3]));
  assert.equal(
    remapWorksheetStyles(
      xml,
      new Map([
        [1, 4],
        [2, 5],
        [3, 6],
      ]),
    ),
    '<worksheet><cols><col style="6"/></cols><sheetData><row s="5"><c s="4"/></row></sheetData></worksheet>',
  );
});
