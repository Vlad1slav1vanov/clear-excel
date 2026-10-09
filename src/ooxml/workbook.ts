import JSZip from "jszip";

import { createError } from "../errors.js";
import type { NormalizedSanitizeOptions, ReadBudget } from "../types.js";
import { hasZipSignature, validateZipCentralDirectory } from "../zip/central-directory.js";
import { removeThreadedCommentExtensions, removeWorksheetDrawingMarkup } from "./cleanup.js";
import { DRAWING_RELATIONSHIP_SUFFIXES, PERSON_RELATIONSHIP_SUFFIXES } from "./constants.js";
import { readXmlPart, requireXml, validateUncompressedSize } from "./parts.js";
import {
  parseRelationships,
  relationshipFileForPart,
  removeRelationships,
  removeUnreachableParts,
  type Relationship,
} from "./relationships.js";
import { removeUnusedCellStyles } from "./styles.js";
import { findElements, getNamespacedAttribute, isSpreadsheetElement } from "./xml.js";

const WORKBOOK_PATH = "xl/workbook.xml";
const WORKBOOK_RELATIONSHIPS_PATH = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES_PATH = "[Content_Types].xml";

interface WorkbookContext {
  zip: JSZip;
  readBudget: ReadBudget;
  workbookXml: string;
  workbookRelationshipsXml: string;
  workbookRelationships: Relationship[];
}

export async function sanitizeOoxmlWorkbook(
  buffer: Buffer,
  options: NormalizedSanitizeOptions,
): Promise<Buffer> {
  if (!hasZipSignature(buffer)) {
    throw new Error("Invalid OOXML ZIP container");
  }
  validateZipCentralDirectory(buffer, options.limits);

  const context = await loadWorkbookContext(buffer, options);
  const worksheetPaths = findWorksheetPaths(
    context.workbookXml,
    context.workbookRelationships,
    options.requireSingleWorksheet,
  );
  const detachedRoots = cleanWorkbookRelationships(context);
  const worksheetXmlByPath = await cleanWorksheets(context, worksheetPaths, detachedRoots, options);

  await removeUnreachableParts(context.zip, detachedRoots, context.readBudget, options.limits);
  await removeUnusedCellStyles(context.zip, worksheetXmlByPath, context.readBudget, options.limits);
  await validateUncompressedSize(context.zip, options.limits.maxUncompressedBytes);
  return context.zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

async function loadWorkbookContext(
  buffer: Buffer,
  options: NormalizedSanitizeOptions,
): Promise<WorkbookContext> {
  const readBudget: ReadBudget = {
    consumedBytes: 0,
    limitBytes: options.limits.maxXmlBytes,
  };

  // CRC verification eagerly inflates every archive entry. Required XML parts
  // are instead inflated lazily below and constrained by per-part budgets.
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await requireXml(
    zip,
    WORKBOOK_PATH,
    readBudget,
    options.limits.maxMetadataXmlBytes,
  );
  const workbookRelationshipsXml = await requireXml(
    zip,
    WORKBOOK_RELATIONSHIPS_PATH,
    readBudget,
    options.limits.maxRelationshipsXmlBytes,
  );
  await requireXml(zip, CONTENT_TYPES_PATH, readBudget, options.limits.maxMetadataXmlBytes);
  // readXmlPart normalizes encoding declarations; persist the decoded metadata as UTF-8.
  zip.file(WORKBOOK_PATH, workbookXml);

  return {
    zip,
    readBudget,
    workbookXml,
    workbookRelationshipsXml,
    workbookRelationships: parseRelationships(workbookRelationshipsXml, WORKBOOK_PATH),
  };
}

function findWorksheetPaths(
  workbookXml: string,
  workbookRelationships: readonly Relationship[],
  requireSingleWorksheet: boolean,
): string[] {
  const sheets = findElements(workbookXml, "sheet").filter(
    (element) =>
      isSpreadsheetElement(element) &&
      element.parent !== undefined &&
      isSpreadsheetElement(element.parent, "sheets"),
  );
  enforceWorksheetCount(sheets.length, requireSingleWorksheet);
  if (sheets.length === 0) {
    throw new Error("Workbook contains no sheets");
  }

  const relationshipsById = new Map(
    workbookRelationships.map((relationship) => [relationship.id, relationship]),
  );
  const worksheetPaths = sheets.flatMap((sheet) => {
    const id = getNamespacedAttribute(sheet, "id", [
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "http://purl.oclc.org/ooxml/officeDocument/relationships",
    ]);
    const relationship = id ? relationshipsById.get(id) : undefined;
    if (!relationship || relationship.targetMode?.toLowerCase() === "external") {
      throw new Error(
        `Missing or invalid workbook sheet relationship: ${id ?? "no relationship ID"}`,
      );
    }
    if (
      !["/worksheet", "/chartsheet", "/dialogsheet", "/macrosheet", "/intlMacrosheet"].some(
        (suffix) => relationship.type.endsWith(suffix),
      )
    ) {
      throw new Error(`Invalid workbook sheet relationship type: ${relationship.type}`);
    }
    return relationship.type.endsWith("/worksheet") ? [relationship.target] : [];
  });
  if (worksheetPaths.length === 0) {
    throw new Error("Workbook contains no worksheets");
  }
  return worksheetPaths;
}

function cleanWorkbookRelationships(context: WorkbookContext): Set<string> {
  const cleaned = removeRelationships(
    context.workbookRelationshipsXml,
    WORKBOOK_PATH,
    PERSON_RELATIONSHIP_SUFFIXES,
  );
  context.zip.file(WORKBOOK_RELATIONSHIPS_PATH, cleaned.xml);
  return cleaned.targets;
}

async function cleanWorksheets(
  context: WorkbookContext,
  worksheetPaths: readonly string[],
  detachedRoots: Set<string>,
  options: NormalizedSanitizeOptions,
): Promise<Map<string, string>> {
  const worksheetXmlByPath = new Map<string, string>();
  for (const worksheetPath of worksheetPaths) {
    // Reads intentionally stay sequential because they share a mutable byte budget.
    // oxlint-disable-next-line no-await-in-loop
    const worksheetXml = await requireXml(
      context.zip,
      worksheetPath,
      context.readBudget,
      options.limits.maxWorksheetXmlBytes,
    );
    const cleanedWorksheetXml = removeThreadedCommentExtensions(
      removeWorksheetDrawingMarkup(worksheetXml),
    );
    worksheetXmlByPath.set(worksheetPath, cleanedWorksheetXml);
    context.zip.file(worksheetPath, cleanedWorksheetXml);

    // Relationship reads consume the same shared byte budget as worksheet reads.
    // oxlint-disable-next-line no-await-in-loop
    await cleanWorksheetRelationships(context, worksheetPath, detachedRoots, options);
  }
  return worksheetXmlByPath;
}

async function cleanWorksheetRelationships(
  context: WorkbookContext,
  worksheetPath: string,
  detachedRoots: Set<string>,
  options: NormalizedSanitizeOptions,
): Promise<void> {
  const relationshipsPath = relationshipFileForPart(worksheetPath);
  const relationshipsFile = context.zip.file(relationshipsPath);
  if (!relationshipsFile) {
    return;
  }

  const cleaned = removeRelationships(
    await readXmlPart(
      relationshipsFile,
      relationshipsPath,
      context.readBudget,
      options.limits.maxRelationshipsXmlBytes,
    ),
    worksheetPath,
    DRAWING_RELATIONSHIP_SUFFIXES,
  );
  context.zip.file(relationshipsPath, cleaned.xml);
  for (const target of cleaned.targets) {
    detachedRoots.add(target);
  }
}

function enforceWorksheetCount(count: number, required: boolean): void {
  if (required && count !== 1) {
    throw createError("ERR_WORKSHEET_COUNT", {
      details: { expected: 1, actual: count },
    });
  }
}
