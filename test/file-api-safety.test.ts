import assert from "node:assert/strict";
import fsPromises, {
  chmod,
  link,
  mkdir,
  mkdtemp,
  readFile,
  rm,
  stat,
  symlink,
  writeFile,
} from "node:fs/promises";
import { syncBuiltinESMExports } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import test from "node:test";

import { sanitizeExcelFile } from "../src/index.js";
import { createXlsx } from "./fixtures/workbook.js";
import { expectError } from "./helpers.js";

test("a failed exclusive open never removes an existing output", async (context) => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-open-failure-"));
  const inputPath = join(directory, "input.xlsx");
  const outputPath = join(directory, "output.xlsx");
  const originalOpen = fsPromises.open;
  try {
    await writeFile(inputPath, await createXlsx());
    await writeFile(outputPath, "existing output");
    context.mock.method(fsPromises, "open", async (...args: Parameters<typeof originalOpen>) => {
      if (args[0] === outputPath) {
        throw Object.assign(new Error("Too many open files"), { code: "EMFILE" });
      }
      return originalOpen(...args);
    });
    syncBuiltinESMExports();

    await expectError(() => sanitizeExcelFile(inputPath, outputPath), "ERR_FILE_WRITE");
    assert.equal(await readFile(outputPath, "utf8"), "existing output");
  } finally {
    context.mock.restoreAll();
    syncBuiltinESMExports();
    await rm(directory, { recursive: true, force: true });
  }
});

test("a failed write removes the partial output that it created", async (context) => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-partial-write-"));
  const inputPath = join(directory, "input.xlsx");
  const outputPath = join(directory, "output.xlsx");
  const originalOpen = fsPromises.open;
  try {
    await writeFile(inputPath, await createXlsx());
    context.mock.method(fsPromises, "open", async (...args: Parameters<typeof originalOpen>) => {
      const handle = await originalOpen(...args);
      if (args[0] === outputPath) {
        await handle.writeFile("partial output");
        context.mock.method(handle, "writeFile", async () => {
          throw Object.assign(new Error("No space left"), { code: "ENOSPC" });
        });
      }
      return handle;
    });
    syncBuiltinESMExports();

    await expectError(() => sanitizeExcelFile(inputPath, outputPath), "ERR_FILE_WRITE");
    await assert.rejects(stat(outputPath), { code: "ENOENT" });
  } finally {
    context.mock.restoreAll();
    syncBuiltinESMExports();
    await rm(directory, { recursive: true, force: true });
  }
});

test("file API refuses source aliases through directories, symlinks, and hardlinks", async () => {
  const directory = await mkdtemp(join(tmpdir(), "clear-excel-source-alias-"));
  const actualDirectory = join(directory, "actual");
  const aliasDirectory = join(directory, "alias");
  const inputPath = join(actualDirectory, "input.xlsx");
  try {
    await mkdir(actualDirectory);
    await symlink(actualDirectory, aliasDirectory, "dir");
    const original = await createXlsx();
    await writeFile(inputPath, original);
    const symlinkPath = join(directory, "symlink.xlsx");
    const hardlinkPath = join(directory, "hardlink.xlsx");
    await symlink(inputPath, symlinkPath, "file");
    await link(inputPath, hardlinkPath);

    await Promise.all(
      [join(aliasDirectory, "input.xlsx"), symlinkPath, hardlinkPath].map((outputPath) =>
        expectError(
          () => sanitizeExcelFile(inputPath, outputPath, { overwrite: true }),
          "ERR_SAME_PATH",
        ),
      ),
    );
    assert.deepEqual(await readFile(inputPath), original);

    const outputPath = join(aliasDirectory, "output.xlsx");
    await sanitizeExcelFile(inputPath, outputPath);
    assert.ok((await readFile(join(actualDirectory, "output.xlsx"))).length > 0);
    assert.deepEqual(await readFile(inputPath), original);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});

test(
  "overwriting an output preserves its file permissions",
  { skip: process.platform === "win32" },
  async () => {
    const directory = await mkdtemp(join(tmpdir(), "clear-excel-output-mode-"));
    const inputPath = join(directory, "input.xlsx");
    try {
      await writeFile(inputPath, await createXlsx());
      await Promise.all(
        [0o600, 0o664].map(async (mode) => {
          const outputPath = join(directory, `output-${mode.toString(8)}.xlsx`);
          await writeFile(outputPath, "existing output");
          await chmod(outputPath, mode);
          await sanitizeExcelFile(inputPath, outputPath, { overwrite: true });
          assert.equal((await stat(outputPath)).mode & 0o777, mode);
        }),
      );
    } finally {
      await rm(directory, { recursive: true, force: true });
    }
  },
);
