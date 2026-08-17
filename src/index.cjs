"use strict";

const { randomUUID } = require("node:crypto");
const fs = require("node:fs/promises");
const { posix: packagePath, resolve: resolveFilePath, dirname } = require("node:path");
const JSZip = require("jszip");

const DRAWING_RELATIONSHIP_SUFFIXES = [
  "/comments",
  "/control",
  "/drawing",
  "/image",
  "/oleObject",
  "/threadedComment",
  "/vmlDrawing",
];
const PERSON_RELATIONSHIP_SUFFIXES = ["/person"];
const DRAWING_ELEMENTS = [
  "controls",
  "drawing",
  "legacyDrawing",
  "legacyDrawingHF",
  "oleObjects",
  "picture",
];

const ERROR_MESSAGES = Object.freeze({
  ERR_INVALID_ARGUMENT: "Invalid argument",
  ERR_UNSUPPORTED_EXTENSION: "Only .xlsx files are supported",
  ERR_INVALID_XLSX: "Invalid or unsupported XLSX workbook",
  ERR_LIMIT_EXCEEDED: "XLSX workbook exceeds a configured safety limit",
  ERR_WORKSHEET_COUNT: "The workbook must contain exactly one worksheet",
  ERR_SAME_PATH: "Input and output paths must be different",
  ERR_OUTPUT_EXISTS: "The output file already exists",
  ERR_FILE_READ: "Unable to read the input XLSX file",
  ERR_FILE_WRITE: "Unable to write the sanitized XLSX file",
});

const DEFAULT_LIMITS = Object.freeze({
  maxEntryCount: 10_000,
  maxUncompressedBytes: 512 * 1024 * 1024,
  maxXmlBytes: 512 * 1024 * 1024,
  maxWorksheetXmlBytes: 256 * 1024 * 1024,
  maxStylesXmlBytes: 64 * 1024 * 1024,
  maxMetadataXmlBytes: 16 * 1024 * 1024,
  maxRelationshipsXmlBytes: 8 * 1024 * 1024,
});

const SANITIZE_OPTION_NAMES = new Set(["requireSingleWorksheet", "limits"]);
const FILE_OPTION_NAMES = new Set([
  "requireSingleWorksheet",
  "limits",
  "overwrite",
]);
const LIMIT_NAMES = new Set(Object.keys(DEFAULT_LIMITS));

class ExcelSanitizationError extends Error {
  constructor(code, message, options = {}) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.name = "ExcelSanitizationError";
    this.code = code;
    if (options.details !== undefined) {
      this.details = Object.freeze({ ...options.details });
    }
  }
}

function createError(code, options = {}) {
  return new ExcelSanitizationError(
    code,
    options.message ?? ERROR_MESSAGES[code] ?? "Excel sanitization failed",
    options,
  );
}

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function validateKnownKeys(value, knownKeys, label) {
  for (const key of Object.keys(value)) {
    if (!knownKeys.has(key)) {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: `Unknown ${label}: ${key}`,
        details: { key },
      });
    }
  }
}

function normalizeLimits(overrides) {
  if (overrides === undefined) {
    return DEFAULT_LIMITS;
  }
  if (!isRecord(overrides)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options.limits must be an object",
    });
  }
  validateKnownKeys(overrides, LIMIT_NAMES, "limit");
  const limits = { ...DEFAULT_LIMITS };
  for (const [name, value] of Object.entries(overrides)) {
    if (!Number.isSafeInteger(value) || value <= 0) {
      throw createError("ERR_INVALID_ARGUMENT", {
        message: `options.limits.${name} must be a positive safe integer`,
        details: { limit: name, value },
      });
    }
    limits[name] = value;
  }
  return Object.freeze(limits);
}

function normalizeOptions(options, fileOptions = false) {
  if (options === undefined) {
    return {
      requireSingleWorksheet: false,
      limits: DEFAULT_LIMITS,
      ...(fileOptions ? { overwrite: false } : {}),
    };
  }
  if (!isRecord(options)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options must be an object",
    });
  }
  validateKnownKeys(
    options,
    fileOptions ? FILE_OPTION_NAMES : SANITIZE_OPTION_NAMES,
    "option",
  );
  if (
    options.requireSingleWorksheet !== undefined &&
    typeof options.requireSingleWorksheet !== "boolean"
  ) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options.requireSingleWorksheet must be a boolean",
    });
  }
  if (
    fileOptions &&
    options.overwrite !== undefined &&
    typeof options.overwrite !== "boolean"
  ) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "options.overwrite must be a boolean",
    });
  }
  return {
    requireSingleWorksheet: options.requireSingleWorksheet ?? false,
    limits: normalizeLimits(options.limits),
    ...(fileOptions ? { overwrite: options.overwrite ?? false } : {}),
  };
}

function isXlsxFilename(name) {
  return typeof name === "string" && name.toLowerCase().endsWith(".xlsx");
}

async function sanitizeExcel(buffer, options) {
  if (!Buffer.isBuffer(buffer)) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "buffer must be a Node.js Buffer",
    });
  }
  const normalized = normalizeOptions(options);
  try {
    return await sanitizeOoxmlWorkbook(buffer, normalized);
  } catch (error) {
    if (error instanceof ExcelSanitizationError) {
      throw error;
    }
    throw createError("ERR_INVALID_XLSX", { cause: error });
  }
}

async function sanitizeExcelFile(inputPath, outputPath, options) {
  if (
    typeof inputPath !== "string" ||
    inputPath.length === 0 ||
    typeof outputPath !== "string" ||
    outputPath.length === 0
  ) {
    throw createError("ERR_INVALID_ARGUMENT", {
      message: "inputPath and outputPath must be non-empty strings",
    });
  }
  if (!isXlsxFilename(inputPath) || !isXlsxFilename(outputPath)) {
    throw createError("ERR_UNSUPPORTED_EXTENSION", {
      details: { inputPath, outputPath },
    });
  }
  const normalized = normalizeOptions(options, true);
  const absoluteInputPath = resolveFilePath(inputPath);
  const absoluteOutputPath = resolveFilePath(outputPath);
  if (absoluteInputPath === absoluteOutputPath) {
    throw createError("ERR_SAME_PATH", {
      details: { inputPath: absoluteInputPath, outputPath: absoluteOutputPath },
    });
  }

  let input;
  try {
    input = await fs.readFile(absoluteInputPath);
  } catch (error) {
    throw createError("ERR_FILE_READ", {
      cause: error,
      details: { inputPath: absoluteInputPath },
    });
  }

  const sanitized = await sanitizeExcel(input, {
    requireSingleWorksheet: normalized.requireSingleWorksheet,
    limits: normalized.limits,
  });

  if (!normalized.overwrite) {
    try {
      await fs.writeFile(absoluteOutputPath, sanitized, { flag: "wx" });
      return;
    } catch (error) {
      if (error?.code === "EEXIST") {
        throw createError("ERR_OUTPUT_EXISTS", {
          cause: error,
          details: { outputPath: absoluteOutputPath },
        });
      }
      await fs.unlink(absoluteOutputPath).catch(() => undefined);
      throw createError("ERR_FILE_WRITE", {
        cause: error,
        details: { outputPath: absoluteOutputPath },
      });
    }
  }

  const temporaryPath = `${absoluteOutputPath}.${process.pid}.${randomUUID()}.tmp`;
  try {
    await fs.writeFile(temporaryPath, sanitized, { flag: "wx" });
    await fs.rename(temporaryPath, absoluteOutputPath);
  } catch (error) {
    await fs.unlink(temporaryPath).catch(() => undefined);
    throw createError("ERR_FILE_WRITE", {
      cause: error,
      details: { outputPath: absoluteOutputPath },
    });
  }
}

async function sanitizeOoxmlWorkbook(buffer, options) {
  const { limits, requireSingleWorksheet } = options;
  if (!hasZipSignature(buffer)) {
    throw new Error("Invalid OOXML ZIP container");
  }
  validateZipCentralDirectory(buffer, limits);
  const readBudget = { consumedBytes: 0, limitBytes: limits.maxXmlBytes };

  // CRC verification eagerly inflates every archive entry. Required XML parts
  // are instead inflated lazily below and constrained by per-part budgets.
  const zip = await JSZip.loadAsync(buffer);
  const workbookXml = await requireXml(
    zip,
    "xl/workbook.xml",
    readBudget,
    limits.maxMetadataXmlBytes,
  );
  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const workbookRelsXml = await requireXml(
    zip,
    workbookRelsPath,
    readBudget,
    limits.maxRelationshipsXmlBytes,
  );
  const workbookRels = parseRelationships(workbookRelsXml, "xl/workbook.xml");
  await requireXml(
    zip,
    "[Content_Types].xml",
    readBudget,
    limits.maxMetadataXmlBytes,
  );

  const sheetTags = workbookXml.match(/<(?:\w+:)?sheet\b[^>]*\/?>/gi) ?? [];
  enforceWorksheetCount(sheetTags.length, requireSingleWorksheet);
  if (sheetTags.length === 0) {
    throw new Error("Workbook contains no sheets");
  }
  const relsById = new Map(workbookRels.map((relationship) => [relationship.id, relationship]));
  const worksheetPaths = sheetTags.flatMap((sheetTag) => {
    const id = getAttribute(sheetTag, "r:id");
    const relationship = id ? relsById.get(id) : undefined;
    return relationship?.type.endsWith("/worksheet")
      ? [relationship.target]
      : [];
  });

  const detachedRoots = new Set();
  const cleanedWorkbookRels = removeRelationships(
    workbookRelsXml,
    "xl/workbook.xml",
    PERSON_RELATIONSHIP_SUFFIXES,
  );
  zip.file(workbookRelsPath, cleanedWorkbookRels.xml);
  for (const target of cleanedWorkbookRels.targets) {
    detachedRoots.add(target);
  }

  const worksheetXmlByPath = new Map();
  for (const worksheetPath of worksheetPaths) {
    const worksheetXml = await requireXml(
      zip,
      worksheetPath,
      readBudget,
      limits.maxWorksheetXmlBytes,
    );
    const cleanedWorksheetXml = removeThreadedCommentExtensions(
      removeWorksheetDrawingMarkup(worksheetXml),
    );
    worksheetXmlByPath.set(worksheetPath, cleanedWorksheetXml);
    zip.file(worksheetPath, cleanedWorksheetXml);

    const relationshipsPath = relationshipFileForPart(worksheetPath);
    const relationshipsFile = zip.file(relationshipsPath);
    if (relationshipsFile) {
      const cleaned = removeRelationships(
        await readXmlPart(
          relationshipsFile,
          relationshipsPath,
          readBudget,
          limits.maxRelationshipsXmlBytes,
        ),
        worksheetPath,
        DRAWING_RELATIONSHIP_SUFFIXES,
      );
      zip.file(relationshipsPath, cleaned.xml);
      for (const target of cleaned.targets) {
        detachedRoots.add(target);
      }
    }
  }

  await removeUnreachableParts(zip, detachedRoots, readBudget, limits);
  await removeUnusedCellStyles(zip, worksheetXmlByPath, readBudget, limits);
  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

function hasZipSignature(buffer) {
  return buffer.length >= 2 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

function validateZipCentralDirectory(buffer, limits) {
  const endOffset = findEndOfCentralDirectory(buffer);
  if (endOffset === undefined || endOffset + 22 > buffer.length) {
    throw new Error("ZIP end-of-central-directory record is missing");
  }
  const diskNumber = buffer.readUInt16LE(endOffset + 4);
  const centralDirectoryDisk = buffer.readUInt16LE(endOffset + 6);
  const entriesOnDisk = buffer.readUInt16LE(endOffset + 8);
  const totalEntries = buffer.readUInt16LE(endOffset + 10);
  const centralDirectorySize = buffer.readUInt32LE(endOffset + 12);
  const centralDirectoryOffset = buffer.readUInt32LE(endOffset + 16);
  const commentLength = buffer.readUInt16LE(endOffset + 20);
  if (endOffset + 22 + commentLength > buffer.length) {
    throw new Error("Truncated ZIP comment");
  }
  if (
    diskNumber !== 0 ||
    centralDirectoryDisk !== 0 ||
    entriesOnDisk !== totalEntries
  ) {
    throw new Error("Multi-disk ZIP files are not supported");
  }
  if (
    totalEntries === 0xffff ||
    centralDirectorySize === 0xffffffff ||
    centralDirectoryOffset === 0xffffffff ||
    (endOffset >= 20 && buffer.readUInt32LE(endOffset - 20) === 0x07064b50)
  ) {
    throw new Error("ZIP64 files are not supported");
  }
  if (totalEntries > limits.maxEntryCount) {
    throw limitExceeded("maxEntryCount", limits.maxEntryCount, totalEntries);
  }

  const centralDirectoryEnd = centralDirectoryOffset + centralDirectorySize;
  if (
    centralDirectoryOffset > endOffset ||
    centralDirectoryEnd > endOffset ||
    centralDirectoryEnd > buffer.length
  ) {
    throw new Error("Invalid ZIP central-directory bounds");
  }

  let entryOffset = centralDirectoryOffset;
  let totalUncompressedSize = 0;
  for (let index = 0; index < totalEntries; index += 1) {
    if (
      entryOffset + 46 > centralDirectoryEnd ||
      buffer.readUInt32LE(entryOffset) !== 0x02014b50
    ) {
      throw new Error("Malformed ZIP central-directory entry");
    }
    const compressedSize = buffer.readUInt32LE(entryOffset + 20);
    const uncompressedSize = buffer.readUInt32LE(entryOffset + 24);
    const fileNameLength = buffer.readUInt16LE(entryOffset + 28);
    const extraLength = buffer.readUInt16LE(entryOffset + 30);
    const fileCommentLength = buffer.readUInt16LE(entryOffset + 32);
    const entryDisk = buffer.readUInt16LE(entryOffset + 34);
    const localHeaderOffset = buffer.readUInt32LE(entryOffset + 42);
    if (
      compressedSize === 0xffffffff ||
      uncompressedSize === 0xffffffff ||
      localHeaderOffset === 0xffffffff ||
      entryDisk !== 0
    ) {
      throw new Error("ZIP64 or multi-disk entries are not supported");
    }

    const fileNameOffset = entryOffset + 46;
    const extraOffset = fileNameOffset + fileNameLength;
    const nextEntryOffset = extraOffset + extraLength + fileCommentLength;
    if (nextEntryOffset > centralDirectoryEnd) {
      throw new Error("Truncated ZIP central-directory entry");
    }
    validatePackageEntryName(
      buffer.subarray(fileNameOffset, extraOffset).toString("utf8"),
    );
    rejectZip64ExtraFields(buffer, extraOffset, extraLength);
    if (localHeaderOffset >= centralDirectoryOffset) {
      throw new Error("Invalid ZIP local-header offset");
    }

    totalUncompressedSize += uncompressedSize;
    if (totalUncompressedSize > limits.maxUncompressedBytes) {
      throw limitExceeded(
        "maxUncompressedBytes",
        limits.maxUncompressedBytes,
        totalUncompressedSize,
      );
    }
    entryOffset = nextEntryOffset;
  }
  if (entryOffset !== centralDirectoryEnd) {
    throw new Error("ZIP central-directory entry count is inconsistent");
  }
}

function validatePackageEntryName(name) {
  const normalized = name.replaceAll("\\", "/");
  if (
    normalized.includes("\0") ||
    normalized.startsWith("/") ||
    /^[A-Za-z]:\//.test(normalized) ||
    normalized.split("/").includes("..")
  ) {
    throw new Error(`Unsafe ZIP entry name: ${name}`);
  }
}

function findEndOfCentralDirectory(buffer) {
  const minimumOffset = Math.max(0, buffer.length - 65_557);
  for (let offset = buffer.length - 22; offset >= minimumOffset; offset -= 1) {
    if (
      buffer.readUInt32LE(offset) === 0x06054b50 &&
      offset + 22 + buffer.readUInt16LE(offset + 20) === buffer.length
    ) {
      return offset;
    }
  }
  return undefined;
}

function rejectZip64ExtraFields(buffer, extraOffset, extraLength) {
  const extraEnd = extraOffset + extraLength;
  let offset = extraOffset;
  while (offset < extraEnd) {
    if (offset + 4 > extraEnd) {
      throw new Error("Malformed ZIP extra field");
    }
    const headerId = buffer.readUInt16LE(offset);
    const fieldLength = buffer.readUInt16LE(offset + 2);
    offset += 4;
    if (offset + fieldLength > extraEnd) {
      throw new Error("Truncated ZIP extra field");
    }
    if (headerId === 0x0001) {
      throw new Error("ZIP64 files are not supported");
    }
    offset += fieldLength;
  }
}

function limitExceeded(limit, configured, actual, partName) {
  return createError("ERR_LIMIT_EXCEEDED", {
    details: {
      limit,
      configured,
      actual,
      ...(partName === undefined ? {} : { partName }),
    },
  });
}

function enforceWorksheetCount(count, required) {
  if (required && count !== 1) {
    throw createError("ERR_WORKSHEET_COUNT", {
      details: { expected: 1, actual: count },
    });
  }
}

async function requireXml(zip, partName, budget, partLimitBytes) {
  const file = zip.file(partName);
  if (!file) {
    throw new Error(`Required workbook part is missing: ${partName}`);
  }
  return readXmlPart(file, partName, budget, partLimitBytes);
}

function readXmlPart(file, partName, budget, partLimitBytes) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    const stream = file.internalStream("nodebuffer");
    let partBytes = 0;
    let settled = false;

    const fail = (error) => {
      if (settled) {
        return;
      }
      settled = true;
      stream.pause();
      reject(error);
    };

    stream
      .on("data", (chunk) => {
        if (settled) {
          return;
        }
        partBytes += chunk.length;
        budget.consumedBytes += chunk.length;
        if (partBytes > partLimitBytes) {
          fail(limitExceeded("partBytes", partLimitBytes, partBytes, partName));
          return;
        }
        if (budget.consumedBytes > budget.limitBytes) {
          fail(
            limitExceeded(
              "maxXmlBytes",
              budget.limitBytes,
              budget.consumedBytes,
              partName,
            ),
          );
          return;
        }
        chunks.push(Buffer.from(chunk));
      })
      .on("error", (error) => fail(error))
      .on("end", () => {
        if (settled) {
          return;
        }
        settled = true;
        resolve(Buffer.concat(chunks, partBytes).toString("utf8"));
      })
      .resume();
  });
}

function removeWorksheetDrawingMarkup(xml) {
  return DRAWING_ELEMENTS.reduce((result, name) => {
    const element = new RegExp(
      `<(?:\\w+:)?${name}\\b[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/(?:\\w+:)?${name}\\s*>)`,
      "gi",
    );
    return result.replace(element, "");
  }, xml);
}

function removeThreadedCommentExtensions(xml) {
  return xml.replace(
    /<(?:\w+:)?ext\b[^>]*>[\s\S]*?<\/(?:\w+:)?ext\s*>/gi,
    (extension) =>
      /<(?:\w+:)?threadedComments?\b/i.test(extension) ? "" : extension,
  );
}

function removeRelationships(xml, source, relationshipSuffixes) {
  const targets = new Set();
  const cleaned = xml.replace(
    /<(?:\w+:)?Relationship\b[^>]*?(?:\/\s*>|>\s*<\/(?:\w+:)?Relationship\s*>)/gi,
    (tag) => {
      const type = getAttribute(tag, "Type") ?? "";
      const target = getAttribute(tag, "Target");
      const targetMode = getAttribute(tag, "TargetMode");
      if (
        !target ||
        targetMode?.toLowerCase() === "external" ||
        !relationshipSuffixes.some((suffix) => type.endsWith(suffix))
      ) {
        return tag;
      }
      targets.add(resolveTarget(source, target));
      return "";
    },
  );
  return { xml: cleaned, targets };
}

function parseRelationships(xml, source) {
  return Array.from(
    xml.matchAll(
      /<(?:\w+:)?Relationship\b([^>]*?)(?:\/\s*>|>\s*<\/(?:\w+:)?Relationship\s*>)/gi,
    ),
  ).flatMap((match) => {
    const attributes = match[1];
    const id = getAttribute(attributes, "Id");
    const target = getAttribute(attributes, "Target");
    const type = getAttribute(attributes, "Type");
    if (!id || !target || !type) {
      return [];
    }
    const targetMode = getAttribute(attributes, "TargetMode");
    return [
      {
        id,
        target:
          targetMode?.toLowerCase() === "external"
            ? target
            : resolveTarget(source, target),
        targetMode,
        type,
      },
    ];
  });
}

function getAttribute(xml, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = xml.match(
    new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*(["'])(.*?)\\1`, "i"),
  );
  return match ? decodeXml(match[2]) : undefined;
}

function decodeXml(value) {
  return value
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">");
}

function resolveTarget(source, target) {
  const resolved = target.startsWith("/")
    ? packagePath.normalize(target.slice(1))
    : packagePath.normalize(packagePath.join(packagePath.dirname(source), target));
  if (resolved === ".." || resolved.startsWith("../")) {
    throw new Error(`Relationship target escapes package: ${target}`);
  }
  return resolved;
}

function relationshipFileForPart(partName) {
  return packagePath.join(
    packagePath.dirname(partName),
    "_rels",
    `${packagePath.basename(partName)}.rels`,
  );
}

function sourcePartForRelationshipFile(fileName) {
  if (fileName === "_rels/.rels") {
    return "";
  }
  const match = fileName.match(/^(.*\/)_rels\/([^/]+)\.rels$/);
  return match ? `${match[1]}${match[2]}` : undefined;
}

async function removeUnreachableParts(zip, detachedRoots, readBudget, limits) {
  if (detachedRoots.size === 0) {
    return;
  }
  const graph = new Map();
  for (const fileName of Object.keys(zip.files)) {
    if (!fileName.endsWith(".rels")) {
      continue;
    }
    const source = sourcePartForRelationshipFile(fileName);
    const file = zip.file(fileName);
    if (source !== undefined && file) {
      graph.set(
        source,
        parseRelationships(
          await readXmlPart(
            file,
            fileName,
            readBudget,
            limits.maxRelationshipsXmlBytes,
          ),
          source,
        ),
      );
    }
  }

  const candidates = collectReachableParts(detachedRoots, graph);
  const packageReachable = collectReachableParts(new Set([""]), graph);
  packageReachable.delete("");
  const deletedParts = new Set(
    [...candidates].filter((partName) => !packageReachable.has(partName)),
  );
  for (const partName of deletedParts) {
    zip.remove(partName);
    zip.remove(relationshipFileForPart(partName));
  }

  const contentTypes = await requireXml(
    zip,
    "[Content_Types].xml",
    readBudget,
    limits.maxMetadataXmlBytes,
  );
  zip.file(
    "[Content_Types].xml",
    contentTypes.replace(/<(?:\w+:)?Override\b[^>]*\/>/gi, (tag) => {
      const partName = getAttribute(tag, "PartName")?.replace(/^\//, "");
      return partName && deletedParts.has(partName) ? "" : tag;
    }),
  );
}

function collectReachableParts(initial, graph) {
  const reachable = new Set(initial);
  const pending = [...initial];
  while (pending.length > 0) {
    const source = pending.pop();
    if (source === undefined) {
      continue;
    }
    for (const relationship of graph.get(source) ?? []) {
      if (
        relationship.targetMode?.toLowerCase() !== "external" &&
        !reachable.has(relationship.target)
      ) {
        reachable.add(relationship.target);
        pending.push(relationship.target);
      }
    }
  }
  return reachable;
}

async function removeUnusedCellStyles(zip, worksheets, readBudget, limits) {
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
  const retainedIndexes = [...usedStyles].sort((left, right) => left - right);
  if (
    retainedIndexes.some(
      (index) => !Number.isInteger(index) || index < 0 || index >= xfs.length,
    )
  ) {
    throw new Error(
      `Worksheet references an invalid style index (${retainedIndexes.join(",")} of ${xfs.length})`,
    );
  }
  const indexMap = new Map(
    retainedIndexes.map((oldIndex, newIndex) => [oldIndex, newIndex]),
  );
  zip.file(
    "xl/styles.xml",
    replaceCollection(
      stylesXml,
      "cellXfs",
      retainedIndexes.map((index) => xfs[index]).join(""),
      retainedIndexes.length,
    ),
  );
  for (const [worksheetPath, xml] of worksheets) {
    zip.file(worksheetPath, remapWorksheetStyles(xml, indexMap));
  }
}

function collectStyleReferences(xml, references) {
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

function remapWorksheetStyles(xml, indexMap) {
  return xml
    .replace(/<(?:\w+:)?(?:c|row)\b[^>]*>/gi, (tag) =>
      remapNumericAttribute(tag, "s", indexMap),
    )
    .replace(/<(?:\w+:)?col\b[^>]*>/gi, (tag) =>
      remapNumericAttribute(tag, "style", indexMap),
    );
}

function remapNumericAttribute(tag, attributeName, indexMap) {
  const expression = new RegExp(
    `(\\s${attributeName}\\s*=\\s*)(["'])(\\d+)\\2`,
    "i",
  );
  return tag.replace(expression, (match, prefix, quote, value) => {
    const remapped = indexMap.get(Number(value));
    return remapped === undefined
      ? match
      : `${prefix}${quote}${remapped}${quote}`;
  });
}

function findCollection(xml, name) {
  const match = xml.match(
    new RegExp(
      `<(?:\\w+:)?${name}\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?${name}\\s*>`,
      "i",
    ),
  );
  if (!match) {
    return undefined;
  }
  return {
    body: match[1],
    full: match[0],
    opening: match[0].slice(0, match[0].indexOf(">") + 1),
  };
}

function replaceCollection(xml, name, body, count) {
  const collection = findCollection(xml, name);
  if (!collection) {
    return xml;
  }
  const opening = /\bcount\s*=/.test(collection.opening)
    ? collection.opening.replace(
        /(\bcount\s*=\s*)(["'])\d+\2/i,
        `$1"${count}"`,
      )
    : collection.opening.replace(/>$/, ` count="${count}">`);
  const prefix = collection.opening.match(/^<(\w+:)/)?.[1] ?? "";
  return xml.replace(collection.full, `${opening}${body}</${prefix}${name}>`);
}

function matchElements(xml, name) {
  return (
    xml.match(
      new RegExp(
        `<(?:\\w+:)?${name}\\b[^>]*\\/\\s*>|<(?:\\w+:)?${name}\\b[^>]*>[\\s\\S]*?<\\/(?:\\w+:)?${name}\\s*>`,
        "gi",
      ),
    ) ?? []
  );
}

module.exports = {
  DEFAULT_LIMITS,
  ExcelSanitizationError,
  isXlsxFilename,
  sanitizeExcel,
  sanitizeExcelFile,
};
