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

export function readXmlPart(
  file: JSZip.JSZipObject,
  partName: string,
  budget: ReadBudget,
  partLimitBytes: number,
): Promise<string> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
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
        if (partBytes > partLimitBytes) {
          fail(limitExceeded("partBytes", partLimitBytes, partBytes, partName));
          return;
        }
        if (budget.consumedBytes > budget.limitBytes) {
          fail(limitExceeded("maxXmlBytes", budget.limitBytes, budget.consumedBytes, partName));
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
