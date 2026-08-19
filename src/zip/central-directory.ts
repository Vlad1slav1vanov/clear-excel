import { limitExceeded } from "../errors.js";
import type { ExcelSanitizationLimits } from "../types.js";

const END_RECORD_MINIMUM_SIZE = 22;
const CENTRAL_DIRECTORY_ENTRY_MINIMUM_SIZE = 46;
const MAX_ZIP_COMMENT_SIZE = 65_535;
const END_RECORD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_ENTRY_SIGNATURE = 0x02014b50;
const ZIP64_END_LOCATOR_SIGNATURE = 0x07064b50;

interface EndRecord {
  offset: number;
  totalEntries: number;
  centralDirectoryOffset: number;
  centralDirectorySize: number;
}

interface CentralDirectoryEntry {
  nextOffset: number;
  uncompressedSize: number;
}

export function hasZipSignature(buffer: Buffer): boolean {
  return buffer.length >= 2 && buffer[0] === 0x50 && buffer[1] === 0x4b;
}

export function validateZipCentralDirectory(
  buffer: Buffer,
  limits: Readonly<ExcelSanitizationLimits>,
): void {
  const endRecord = readEndRecord(buffer, limits);
  const centralDirectoryEnd = endRecord.centralDirectoryOffset + endRecord.centralDirectorySize;

  if (
    endRecord.centralDirectoryOffset > endRecord.offset ||
    centralDirectoryEnd > endRecord.offset ||
    centralDirectoryEnd > buffer.length
  ) {
    throw new Error("Invalid ZIP central-directory bounds");
  }

  let entryOffset = endRecord.centralDirectoryOffset;
  let totalUncompressedSize = 0;
  for (let index = 0; index < endRecord.totalEntries; index += 1) {
    const entry = readCentralDirectoryEntry(
      buffer,
      entryOffset,
      centralDirectoryEnd,
      endRecord.centralDirectoryOffset,
    );
    totalUncompressedSize += entry.uncompressedSize;
    if (totalUncompressedSize > limits.maxUncompressedBytes) {
      throw limitExceeded(
        "maxUncompressedBytes",
        limits.maxUncompressedBytes,
        totalUncompressedSize,
      );
    }
    entryOffset = entry.nextOffset;
  }

  if (entryOffset !== centralDirectoryEnd) {
    throw new Error("ZIP central-directory entry count is inconsistent");
  }
}

function readEndRecord(buffer: Buffer, limits: Readonly<ExcelSanitizationLimits>): EndRecord {
  const offset = findEndOfCentralDirectory(buffer);
  if (offset === undefined || offset + END_RECORD_MINIMUM_SIZE > buffer.length) {
    throw new Error("ZIP end-of-central-directory record is missing");
  }

  const diskNumber = buffer.readUInt16LE(offset + 4);
  const centralDirectoryDisk = buffer.readUInt16LE(offset + 6);
  const entriesOnDisk = buffer.readUInt16LE(offset + 8);
  const totalEntries = buffer.readUInt16LE(offset + 10);
  const centralDirectorySize = buffer.readUInt32LE(offset + 12);
  const centralDirectoryOffset = buffer.readUInt32LE(offset + 16);
  const commentLength = buffer.readUInt16LE(offset + 20);

  if (offset + END_RECORD_MINIMUM_SIZE + commentLength > buffer.length) {
    throw new Error("Truncated ZIP comment");
  }
  if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesOnDisk !== totalEntries) {
    throw new Error("Multi-disk ZIP files are not supported");
  }
  if (
    totalEntries === 0xffff ||
    centralDirectorySize === 0xffffffff ||
    centralDirectoryOffset === 0xffffffff ||
    (offset >= 20 && buffer.readUInt32LE(offset - 20) === ZIP64_END_LOCATOR_SIGNATURE)
  ) {
    throw new Error("ZIP64 files are not supported");
  }
  if (totalEntries > limits.maxEntryCount) {
    throw limitExceeded("maxEntryCount", limits.maxEntryCount, totalEntries);
  }

  return { offset, totalEntries, centralDirectoryOffset, centralDirectorySize };
}

function readCentralDirectoryEntry(
  buffer: Buffer,
  entryOffset: number,
  centralDirectoryEnd: number,
  centralDirectoryOffset: number,
): CentralDirectoryEntry {
  if (
    entryOffset + CENTRAL_DIRECTORY_ENTRY_MINIMUM_SIZE > centralDirectoryEnd ||
    buffer.readUInt32LE(entryOffset) !== CENTRAL_DIRECTORY_ENTRY_SIGNATURE
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

  const fileNameOffset = entryOffset + CENTRAL_DIRECTORY_ENTRY_MINIMUM_SIZE;
  const extraOffset = fileNameOffset + fileNameLength;
  const nextOffset = extraOffset + extraLength + fileCommentLength;
  if (nextOffset > centralDirectoryEnd) {
    throw new Error("Truncated ZIP central-directory entry");
  }

  validatePackageEntryName(buffer.subarray(fileNameOffset, extraOffset).toString("utf8"));
  rejectZip64ExtraFields(buffer, extraOffset, extraLength);
  if (localHeaderOffset >= centralDirectoryOffset) {
    throw new Error("Invalid ZIP local-header offset");
  }

  return { nextOffset, uncompressedSize };
}

export function validatePackageEntryName(name: string): void {
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

export function findEndOfCentralDirectory(buffer: Buffer): number | undefined {
  const minimumOffset = Math.max(0, buffer.length - MAX_ZIP_COMMENT_SIZE - END_RECORD_MINIMUM_SIZE);
  for (let offset = buffer.length - END_RECORD_MINIMUM_SIZE; offset >= minimumOffset; offset -= 1) {
    if (
      buffer.readUInt32LE(offset) === END_RECORD_SIGNATURE &&
      offset + END_RECORD_MINIMUM_SIZE + buffer.readUInt16LE(offset + 20) === buffer.length
    ) {
      return offset;
    }
  }
  return undefined;
}

export function rejectZip64ExtraFields(
  buffer: Buffer,
  extraOffset: number,
  extraLength: number,
): void {
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
