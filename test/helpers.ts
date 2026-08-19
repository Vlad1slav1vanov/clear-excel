import assert from "node:assert/strict";

import { ExcelSanitizationError, type ExcelSanitizationErrorCode } from "../src/index.js";

export async function expectError(
  operation: () => Promise<unknown>,
  code: ExcelSanitizationErrorCode,
): Promise<void> {
  await assert.rejects(operation, (error: unknown) => {
    assert.ok(error instanceof ExcelSanitizationError);
    assert.equal(error.code, code);
    return true;
  });
}

export function cellXfCount(styles: string): number {
  const count = styles.match(/<cellXfs\b[^>]*\bcount="(\d+)"/)?.[1];
  if (!count) {
    throw new Error("styles.xml has no cellXfs count");
  }
  return Number(count);
}

export function toArrayBuffer(buffer: Buffer): ArrayBuffer {
  return Uint8Array.from(buffer).buffer;
}
