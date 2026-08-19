import JSZip from "jszip";

import { createXlsx } from "./workbook.js";

export function findEndRecord(buffer: Buffer): number {
  for (let offset = buffer.length - 22; offset >= 0; offset -= 1) {
    if (buffer.readUInt32LE(offset) === 0x06054b50) {
      return offset;
    }
  }
  throw new Error("Test ZIP has no end-of-central-directory record");
}

export function inflateDeclaredZipSize(buffer: Buffer): Buffer {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const firstEntryOffset = manipulated.readUInt32LE(endOffset + 16);
  manipulated.writeUInt32LE(600 * 1024 * 1024, firstEntryOffset + 24);
  return manipulated;
}

export function forgeZipEntrySizes(
  buffer: Buffer,
  targetName: string,
  declaredSize: number,
): Buffer {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const entryCount = manipulated.readUInt16LE(endOffset + 10);
  let entryOffset = manipulated.readUInt32LE(endOffset + 16);
  for (let index = 0; index < entryCount; index += 1) {
    const fileNameLength = manipulated.readUInt16LE(entryOffset + 28);
    const extraLength = manipulated.readUInt16LE(entryOffset + 30);
    const commentLength = manipulated.readUInt16LE(entryOffset + 32);
    const fileName = manipulated
      .subarray(entryOffset + 46, entryOffset + 46 + fileNameLength)
      .toString("utf8");
    if (fileName === targetName) {
      const localHeaderOffset = manipulated.readUInt32LE(entryOffset + 42);
      manipulated.writeUInt32LE(declaredSize, entryOffset + 24);
      manipulated.writeUInt32LE(declaredSize, localHeaderOffset + 22);
      return manipulated;
    }
    entryOffset += 46 + fileNameLength + extraLength + commentLength;
  }
  throw new Error(`Test ZIP entry was not found: ${targetName}`);
}

export async function createRuntimeXmlBomb(): Promise<Buffer> {
  const relationshipPath = "xl/_rels/workbook.xml.rels";
  const zip = await JSZip.loadAsync(await createXlsx());
  const relationships = await zip.file(relationshipPath)?.async("string");
  if (!relationships) {
    throw new Error("Test workbook has no workbook relationships");
  }
  zip.file(
    relationshipPath,
    relationships.replace("</Relationships>", `${" ".repeat(9 * 1024 * 1024)}</Relationships>`),
  );
  const compressed = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
  return forgeZipEntrySizes(compressed, relationshipPath, 1);
}

export function markAsMultiDisk(buffer: Buffer): Buffer {
  const manipulated = Buffer.from(buffer);
  manipulated.writeUInt16LE(1, findEndRecord(manipulated) + 4);
  return manipulated;
}

export function markAsZip64(buffer: Buffer): Buffer {
  const manipulated = Buffer.from(buffer);
  manipulated.writeUInt16LE(0xffff, findEndRecord(manipulated) + 10);
  return manipulated;
}

export function addUnsafeCentralDirectoryName(buffer: Buffer): Buffer {
  const manipulated = Buffer.from(buffer);
  const endOffset = findEndRecord(manipulated);
  const entryOffset = manipulated.readUInt32LE(endOffset + 16);
  const fileNameLength = manipulated.readUInt16LE(entryOffset + 28);
  const unsafeName = Buffer.from(`../${"x".repeat(fileNameLength - 3)}`);
  unsafeName.copy(manipulated, entryOffset + 46);
  return manipulated;
}
