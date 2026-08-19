import type JSZip from "jszip";

import type { ExcelSanitizationLimits, ReadBudget } from "../types.js";
import { readXmlPart } from "./parts.js";
import { findCollection, getAttribute, matchElements, replaceCollection } from "./xml.js";

export async function removeUnusedCellStyles(
  zip: JSZip,
  worksheets: ReadonlyMap<string, string>,
  readBudget: ReadBudget,
  limits: Readonly<ExcelSanitizationLimits>,
): Promise<void> {
  const stylesFile = zip.file("xl/styles.xml");
  if (!stylesFile) {
    return;
  }

  const usedStyles = new Set([0]);
  for (const xml of worksheets.values()) {
    collectStyleReferences(xml, usedStyles);
  }

  const stylesXml = await readXmlPart(
    stylesFile,
    "xl/styles.xml",
    readBudget,
    limits.maxStylesXmlBytes,
  );
  const cellXfs = findCollection(stylesXml, "cellXfs");
  if (!cellXfs) {
    return;
  }

  const xfs = matchElements(cellXfs.body, "xf");
  // The spread already creates a fresh array, so this sort cannot mutate the source set.
  // oxlint-disable-next-line unicorn/no-array-sort
  const retainedIndexes = [...usedStyles].sort((left, right) => left - right);
  validateStyleIndexes(retainedIndexes, xfs.length);
  const indexMap = new Map(retainedIndexes.map((oldIndex, newIndex) => [oldIndex, newIndex]));
  const retainedStyles = retainedIndexes.map((index) => {
    const style = xfs[index];
    if (style === undefined) {
      throw new Error(`Worksheet references an invalid style index: ${index}`);
    }
    return style;
  });

  zip.file(
    "xl/styles.xml",
    replaceCollection(stylesXml, "cellXfs", retainedStyles.join(""), retainedIndexes.length),
  );
  for (const [worksheetPath, xml] of worksheets) {
    zip.file(worksheetPath, remapWorksheetStyles(xml, indexMap));
  }
}

function validateStyleIndexes(indexes: readonly number[], styleCount: number): void {
  if (indexes.some((index) => !Number.isInteger(index) || index < 0 || index >= styleCount)) {
    throw new Error(
      `Worksheet references an invalid style index (${indexes.join(",")} of ${styleCount})`,
    );
  }
}

export function collectStyleReferences(xml: string, references: Set<number>): void {
  for (const tag of xml.match(/<(?:\w+:)?(?:c|row)\b[^>]*>/gi) ?? []) {
    const value = getAttribute(tag, "s");
    if (value !== undefined) {
      references.add(Number(value));
    }
  }
  for (const tag of xml.match(/<(?:\w+:)?col\b[^>]*>/gi) ?? []) {
    const value = getAttribute(tag, "style");
    if (value !== undefined) {
      references.add(Number(value));
    }
  }
}

export function remapWorksheetStyles(xml: string, indexMap: ReadonlyMap<number, number>): string {
  return xml
    .replace(/<(?:\w+:)?(?:c|row)\b[^>]*>/gi, (tag) => remapNumericAttribute(tag, "s", indexMap))
    .replace(/<(?:\w+:)?col\b[^>]*>/gi, (tag) => remapNumericAttribute(tag, "style", indexMap));
}

export function remapNumericAttribute(
  tag: string,
  attributeName: string,
  indexMap: ReadonlyMap<number, number>,
): string {
  const expression = new RegExp(`(\\s${attributeName}\\s*=\\s*)(["'])(\\d+)\\2`, "i");
  return tag.replace(expression, (match: string, prefix: string, quote: string, value: string) => {
    const remapped = indexMap.get(Number(value));
    return remapped === undefined ? match : `${prefix}${quote}${remapped}${quote}`;
  });
}
