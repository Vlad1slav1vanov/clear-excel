import * as CFB from "cfb";

import { createError, limitExceeded } from "../errors.js";
import type { NormalizedSanitizeOptions } from "../types.js";

const COMPOUND_FILE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
const RECORD_HEADER_BYTES = 4;
const CFB_STORAGE_ENTRY = 1;
const CFB_STREAM_ENTRY = 2;
const RECORD_CONTINUE = 0x003c;
const RECORD_BOUNDSHEET = 0x0085;
const RECORD_BOF_IDS = new Set([0x0009, 0x0209, 0x0409, 0x0809]);

// INDEX, DBCELL, and ExtSST cache absolute stream offsets. They are optional
// acceleration records and must be discarded after other records move.
const OFFSET_CACHE_RECORD_IDS = new Set([0x00d7, 0x00ff, 0x020b]);

// Drawing-layer records cover OfficeArt drawings, images, notes/comments,
// controls, embedded OLE objects, and header/footer or background pictures.
const DRAWING_RECORD_IDS = new Set([
  0x001c, // Note
  0x005d, // Obj
  0x007f, // ImData
  0x00e9, // BkHim
  0x00eb, // MsoDrawingGroup
  0x00ec, // MsoDrawing
  0x00ed, // MsoDrawingSelection
  0x01b6, // TxO
  0x0866, // HFPicture
]);

interface BiffRecord {
  id: number;
  originalOffset: number;
  payload: Buffer;
  remove: boolean;
}

export function hasCompoundFileSignature(buffer: Buffer): boolean {
  return (
    buffer.length >= COMPOUND_FILE_SIGNATURE.length &&
    buffer.subarray(0, COMPOUND_FILE_SIGNATURE.length).equals(COMPOUND_FILE_SIGNATURE)
  );
}

export async function sanitizeBiffWorkbook(
  buffer: Buffer,
  options: NormalizedSanitizeOptions,
): Promise<Buffer> {
  enforceInputSize(buffer, options);
  const cfb = CFB.read(buffer, { type: "buffer" });
  enforceContainerLimits(cfb, options);

  const workbookEntry = CFB.find(cfb, "Workbook") ?? CFB.find(cfb, "Book");
  if (!workbookEntry || workbookEntry.type !== CFB_STREAM_ENTRY) {
    throw new Error("Compound file has no BIFF Workbook stream");
  }

  const workbookStream = Buffer.from(workbookEntry.content);
  const records = parseBiffRecords(workbookStream);
  rejectEncryptedWorkbook(records);
  enforceWorksheetCount(records, options.requireSingleWorksheet);

  const sanitizedStream = rebuildWorkbookStream(records);
  workbookEntry.content = sanitizedStream;
  workbookEntry.size = sanitizedStream.length;
  removeEmbeddedObjectEntries(cfb);

  return Buffer.from(CFB.write(cfb, { type: "buffer", fileType: "cfb" }));
}

function enforceInputSize(buffer: Buffer, options: NormalizedSanitizeOptions): void {
  const configured = options.limits.maxUncompressedBytes;
  if (buffer.length > configured) {
    throw limitExceeded("maxUncompressedBytes", configured, buffer.length);
  }
}

function enforceContainerLimits(cfb: CFB.CFB$Container, options: NormalizedSanitizeOptions): void {
  const entries = cfb.FileIndex.filter(
    (entry) => entry.type === CFB_STORAGE_ENTRY || entry.type === CFB_STREAM_ENTRY,
  );
  if (entries.length > options.limits.maxEntryCount) {
    throw limitExceeded("maxEntryCount", options.limits.maxEntryCount, entries.length);
  }

  const totalBytes = entries.reduce(
    (total, entry) => (entry.type === CFB_STREAM_ENTRY ? total + entry.content.length : total),
    0,
  );
  if (totalBytes > options.limits.maxUncompressedBytes) {
    throw limitExceeded("maxUncompressedBytes", options.limits.maxUncompressedBytes, totalBytes);
  }
}

function parseBiffRecords(stream: Buffer): BiffRecord[] {
  const records: BiffRecord[] = [];
  let offset = 0;
  let removeContinuation = false;

  while (offset < stream.length) {
    if (offset + RECORD_HEADER_BYTES > stream.length) {
      throw new Error("Truncated BIFF record header");
    }

    const id = stream.readUInt16LE(offset);
    const length = stream.readUInt16LE(offset + 2);
    const nextOffset = offset + RECORD_HEADER_BYTES + length;
    if (nextOffset > stream.length) {
      throw new Error("Truncated BIFF record payload");
    }

    const remove: boolean =
      id === RECORD_CONTINUE
        ? removeContinuation
        : DRAWING_RECORD_IDS.has(id) || OFFSET_CACHE_RECORD_IDS.has(id);
    records.push({
      id,
      originalOffset: offset,
      payload: stream.subarray(offset + RECORD_HEADER_BYTES, nextOffset),
      remove,
    });
    if (id !== RECORD_CONTINUE) {
      removeContinuation = remove;
    }
    offset = nextOffset;
  }

  if (records.length === 0 || !RECORD_BOF_IDS.has(records[0]?.id ?? -1)) {
    throw new Error("Workbook stream does not start with a BIFF BOF record");
  }
  return records;
}

function rejectEncryptedWorkbook(records: readonly BiffRecord[]): void {
  if (records.some((record) => record.id === 0x002f)) {
    throw createError("ERR_INVALID_XLS", {
      message: "Encrypted XLS workbooks are not supported",
    });
  }
}

function enforceWorksheetCount(records: readonly BiffRecord[], required: boolean): void {
  const count = records.filter((record) => record.id === RECORD_BOUNDSHEET).length;
  if (count === 0) {
    throw new Error("Workbook contains no sheets");
  }
  if (required && count !== 1) {
    throw createError("ERR_WORKSHEET_COUNT", {
      details: { expected: 1, actual: count },
    });
  }
}

function rebuildWorkbookStream(records: readonly BiffRecord[]): Buffer {
  const newOffsets = new Map<number, number>();
  let outputLength = 0;
  for (const record of records) {
    if (record.remove) {
      continue;
    }
    newOffsets.set(record.originalOffset, outputLength);
    outputLength += RECORD_HEADER_BYTES + record.payload.length;
  }

  const output = Buffer.allocUnsafe(outputLength);
  let outputOffset = 0;
  for (const record of records) {
    if (record.remove) {
      continue;
    }

    const payload = patchBoundSheetOffset(record, newOffsets);
    output.writeUInt16LE(record.id, outputOffset);
    output.writeUInt16LE(payload.length, outputOffset + 2);
    payload.copy(output, outputOffset + RECORD_HEADER_BYTES);
    outputOffset += RECORD_HEADER_BYTES + payload.length;
  }
  return output;
}

function patchBoundSheetOffset(
  record: BiffRecord,
  newOffsets: ReadonlyMap<number, number>,
): Buffer {
  if (record.id !== RECORD_BOUNDSHEET) {
    return record.payload;
  }
  if (record.payload.length < 4) {
    throw new Error("Invalid BoundSheet record");
  }

  const originalSheetOffset = record.payload.readUInt32LE(0);
  const newSheetOffset = newOffsets.get(originalSheetOffset);
  if (newSheetOffset === undefined) {
    throw new Error("BoundSheet record points outside the Workbook stream");
  }

  const payload = Buffer.from(record.payload);
  payload.writeUInt32LE(newSheetOffset, 0);
  return payload;
}

function removeEmbeddedObjectEntries(cfb: CFB.CFB$Container): void {
  const paths = cfb.FullPaths.filter((path) =>
    path
      .split("/")
      .filter(Boolean)
      .some(
        (segment) =>
          segment.toLowerCase() === "objectpool" ||
          segment.toLowerCase() === "ctls" ||
          /^mbd[0-9a-f]{8}$/i.test(segment),
      ),
  );

  for (let index = paths.length - 1; index >= 0; index -= 1) {
    CFB.utils.cfb_del(cfb, paths[index] ?? "");
  }
  if (paths.length > 0) {
    CFB.utils.cfb_gc(cfb);
  }
}
