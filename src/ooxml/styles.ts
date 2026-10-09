import type JSZip from "jszip";

import type { ExcelSanitizationLimits, ReadBudget } from "../types.js";
import { readXmlPart } from "./parts.js";
import {
  findCollection,
  findElements,
  getAttribute,
  isSpreadsheetElement,
  mapOpeningElements,
  replaceAttribute,
  replaceCollection,
} from "./xml.js";

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

  const xfs = findElements(stylesXml, "xf")
    .filter(
      (element) => isSpreadsheetElement(element) && element.parent?.start === cellXfs.element.start,
    )
    .map((element) => stylesXml.slice(element.start, element.end));
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
  mapOpeningElements(xml, (element, tag) => {
    const attribute = element.localName === "col" ? "style" : "s";
    if (isSpreadsheetElement(element) && ["c", "row", "col"].includes(element.localName)) {
      const value = getAttribute(tag, attribute);
      if (value !== undefined) {
        references.add(Number(value));
      }
    }
    return tag;
  });
}

export function remapWorksheetStyles(xml: string, indexMap: ReadonlyMap<number, number>): string {
  return mapOpeningElements(xml, (element, tag) => {
    if (!isSpreadsheetElement(element) || !["c", "row", "col"].includes(element.localName)) {
      return tag;
    }
    return remapNumericAttribute(tag, element.localName === "col" ? "style" : "s", indexMap);
  });
}

export function remapNumericAttribute(
  tag: string,
  attributeName: string,
  indexMap: ReadonlyMap<number, number>,
): string {
  const value = getAttribute(tag, attributeName);
  const remapped = value === undefined ? undefined : indexMap.get(Number(value));
  return remapped === undefined ? tag : replaceAttribute(tag, attributeName, String(remapped));
}
