import { TextDecoder } from "node:util";

import type JSZip from "jszip";

import { limitExceeded } from "../errors.js";
import type { ReadBudget } from "../types.js";

interface StreamHelper {
  on(event: "data", listener: (chunk: Buffer) => void): StreamHelper;
  on(event: "error", listener: (error: unknown) => void): StreamHelper;
  on(event: "end", listener: () => void): StreamHelper;
  pause(): StreamHelper;
  resume(): StreamHelper;
}

interface StreamableZipObject {
  internalStream(type: "nodebuffer"): StreamHelper;
}

function getNodeBufferStream(file: JSZip.JSZipObject): StreamHelper {
  // JSZip's declarations omit its public internalStream helper, while the
  // runtime API is required to enforce byte budgets before full inflation.
  // oxlint-disable-next-line typescript/no-unsafe-type-assertion
  return (file as unknown as StreamableZipObject).internalStream("nodebuffer");
}

export async function requireXml(
  zip: JSZip,
  partName: string,
  budget: ReadBudget,
  partLimitBytes: number,
): Promise<string> {
  const file = zip.file(partName);
  if (!file) {
    throw new Error(`Required workbook part is missing: ${partName}`);
  }
  return readXmlPart(file, partName, budget, partLimitBytes);
}

export async function readXmlPart(
  file: JSZip.JSZipObject,
  partName: string,
  budget: ReadBudget,
  partLimitBytes: number,
): Promise<string> {
  const chunks: Buffer[] = [];
  await consumePart(
    file,
    partName,
    budget,
    "maxXmlBytes",
    (chunk) => chunks.push(Buffer.from(chunk)),
    partLimitBytes,
  );
  return decodeXmlBytes(Buffer.concat(chunks), partName);
}

export async function validateUncompressedSize(zip: JSZip, limitBytes: number): Promise<void> {
  const budget: ReadBudget = { consumedBytes: 0, limitBytes };
  for (const file of Object.values(zip.files)) {
    if (!file.dir) {
      // Reads stay sequential to enforce the shared budget without retaining file contents.
      // oxlint-disable-next-line no-await-in-loop
      await consumePart(file, file.name, budget, "maxUncompressedBytes");
    }
  }
}

function consumePart(
  file: JSZip.JSZipObject,
  partName: string,
  budget: ReadBudget,
  budgetLimitName: "maxXmlBytes" | "maxUncompressedBytes",
  onChunk?: (chunk: Buffer) => void,
  partLimitBytes?: number,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const stream = getNodeBufferStream(file);
    let partBytes = 0;
    let settled = false;

    const fail = (error: unknown): void => {
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
        if (partLimitBytes !== undefined && partBytes > partLimitBytes) {
          fail(limitExceeded("partBytes", partLimitBytes, partBytes, partName));
          return;
        }
        if (budget.consumedBytes > budget.limitBytes) {
          fail(limitExceeded(budgetLimitName, budget.limitBytes, budget.consumedBytes, partName));
          return;
        }
        onChunk?.(chunk);
      })
      .on("error", (error) => fail(error))
      .on("end", () => {
        if (settled) {
          return;
        }
        settled = true;
        resolve();
      })
      .resume();
  });
}

function decodeXmlBytes(bytes: Buffer, partName: string): string {
  let encoding = "utf-8";
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0x00)) {
    encoding = "utf-16le";
  } else if ((bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0x00 && bytes[1] === 0x3c)) {
    encoding = "utf-16be";
  }
  const xml = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  const declaration = /^<\?xml\s[^?]*\?>/.exec(xml)?.[0];
  const declaredEncoding = declaration
    ?.match(/\bencoding\s*=\s*(["'])([^"']+)\1/i)?.[2]
    ?.toLowerCase();
  if (
    declaredEncoding !== undefined &&
    declaredEncoding !== encoding &&
    !(declaredEncoding === "utf-16" && encoding.startsWith("utf-16"))
  ) {
    throw new Error(`Unsupported or inconsistent XML encoding in ${partName}: ${declaredEncoding}`);
  }
  return declaration
    ? xml.replace(
        declaration,
        declaration.replace(/(\bencoding\s*=\s*)(["'])[^"']+\2/i, '$1"UTF-8"'),
      )
    : xml;
}
