import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { isAbsolute, join } from "node:path";
import { fileURLToPath } from "node:url";

import * as CFB from "cfb";
import JSZip from "jszip";
import { read as readSpreadsheet } from "@e965/xlsx";

import {
  addXlsCompoundStorages,
  createXls,
  createXlsm,
  createXlsx,
} from "../test/fixtures/workbook.js";

interface RunOptions {
  capture?: boolean;
  cwd?: string;
}

interface PackFile {
  path: string;
}

interface PackResult {
  filename: string;
  files: PackFile[];
}

const projectDirectory = fileURLToPath(new URL("..", import.meta.url));
const temporaryDirectory = mkdtempSync(join(tmpdir(), "clear-excel-package-"));
const pnpmCommand = process.platform === "win32" ? "pnpm.cmd" : "pnpm";
const childEnvironment: NodeJS.ProcessEnv = { ...process.env };
delete childEnvironment.npm_config_dry_run;
delete childEnvironment.NPM_CONFIG_DRY_RUN;

function run(command: string, args: readonly string[], options: RunOptions = {}): string {
  return execFileSync(command, args, {
    cwd: options.cwd ?? projectDirectory,
    encoding: "utf8",
    env: childEnvironment,
    stdio: options.capture ? "pipe" : "inherit",
  });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function parsePackResult(output: string): PackResult {
  const parsed: unknown = JSON.parse(output);
  const candidate = Array.isArray(parsed) ? parsed[0] : parsed;
  if (!isRecord(candidate)) {
    throw new Error("pnpm pack returned an unexpected result");
  }
  if (typeof candidate.filename !== "string" || !Array.isArray(candidate.files)) {
    throw new Error("pnpm pack returned an unexpected result");
  }
  const files = candidate.files.map((file) => {
    if (!isRecord(file) || typeof file.path !== "string") {
      throw new Error("pnpm pack returned an invalid file entry");
    }
    return { path: file.path };
  });
  return { filename: candidate.filename, files };
}

function writeTypeConsumer(): void {
  writeFileSync(
    join(temporaryDirectory, "consumer.mts"),
    [
      'import { DEFAULT_LIMITS, ExcelSanitizationError, isExcelFilename, sanitizeExcel, sanitizeExcelFile, type ExcelFileFormat, type SanitizeExcelOptions } from "clear-excel";',
      "const options: SanitizeExcelOptions = { limits: { maxEntryCount: DEFAULT_LIMITS.maxEntryCount } };",
      'const format: ExcelFileFormat = "xlsm";',
      "const supported: boolean = isExcelFilename(`input.${format}`);",
      "const bufferResult: Promise<Buffer> = sanitizeExcel(Buffer.alloc(0), options);",
      'const fileResult: Promise<void> = sanitizeExcelFile("input.xlsx", "output.xlsx");',
      'const errorCode: string = new ExcelSanitizationError("ERR_INVALID_ARGUMENT", "message").code;',
      "void bufferResult;",
      "void fileResult;",
      "void errorCode;",
      "void supported;",
      "",
    ].join("\n"),
  );
  writeFileSync(
    join(temporaryDirectory, "tsconfig.json"),
    `${JSON.stringify(
      {
        compilerOptions: {
          module: "NodeNext",
          moduleResolution: "NodeNext",
          noEmit: true,
          strict: true,
          target: "ES2022",
          types: ["node"],
        },
        include: ["consumer.mts"],
      },
      undefined,
      2,
    )}\n`,
  );
}

async function writeRuntimeConsumer(): Promise<void> {
  const fixtures = [
    { format: "xls", contents: addXlsCompoundStorages(createXls()) },
    { format: "xlsx", contents: await createXlsx() },
    { format: "xlsm", contents: await createXlsm() },
  ];
  for (const fixture of fixtures) {
    writeFileSync(join(temporaryDirectory, `input.${fixture.format}`), fixture.contents);
  }
  writeFileSync(
    join(temporaryDirectory, "consumer.mjs"),
    [
      'import assert from "node:assert/strict";',
      'import { readFile, writeFile } from "node:fs/promises";',
      'import { sanitizeExcel, sanitizeExcelFile, isExcelFilename, ExcelSanitizationError, DEFAULT_LIMITS } from "clear-excel";',
      'assert.ok(isExcelFilename("input.xlsm"));',
      'assert.equal(typeof ExcelSanitizationError, "function");',
      "assert.equal(DEFAULT_LIMITS.maxEntryCount, 10000);",
      'await Promise.all(["xls", "xlsx", "xlsm"].map(async (format) => {',
      "  const inputPath = `input.${format}`;",
      "  const outputPath = `file.${format}`;",
      "  const input = await readFile(inputPath);",
      "  const original = Buffer.from(input);",
      "  const output = await sanitizeExcel(input);",
      "  assert.ok(Buffer.isBuffer(output) && output.length > 0);",
      "  assert.deepEqual(input, original);",
      "  assert.notDeepEqual(output, original);",
      "  await writeFile(`buffer.${format}`, output);",
      "  await sanitizeExcelFile(inputPath, outputPath);",
      '  await assert.rejects(sanitizeExcelFile(inputPath, outputPath), { code: "ERR_OUTPUT_EXISTS" });',
      "  await sanitizeExcelFile(inputPath, outputPath, { overwrite: true });",
      "  assert.deepEqual(await readFile(inputPath), original);",
      "}));",
      "",
    ].join("\n"),
  );
}

async function verifySanitizedOutput(format: string, api: string): Promise<void> {
  const output = readFileSync(join(temporaryDirectory, `${api}.${format}`));
  const workbook = readSpreadsheet(output, { type: "buffer" });
  const worksheet = workbook.Sheets["Sheet 1"];
  assert.equal(worksheet?.A1?.v, "kept value", `${api} ${format}: cell value`);
  assert.equal(worksheet?.E1?.c, undefined, `${api} ${format}: comment removed`);
  assert.equal(worksheet?.A1?.c, undefined, `${api} ${format}: legacy comment removed`);
  assert.equal(worksheet?.["!merges"]?.length, 1, `${api} ${format}: merged cells`);

  if (format === "xls") {
    const container = CFB.read(output, { type: "buffer" });
    assert.ok(container.FullPaths.some((path) => path.includes("_VBA_PROJECT_CUR/VBA/dir")));
    assert.ok(!container.FullPaths.some((path) => path.includes("ObjectPool/")));
    return;
  }

  assert.equal(worksheet?.B1?.f, "SUM(A2:A3)", `${api} ${format}: formula`);
  assert.equal(worksheet?.B1?.v, 3, `${api} ${format}: cached formula result`);
  const zip = await JSZip.loadAsync(output);
  assert.ok(
    !Object.values(zip.files).some(
      (file) =>
        !file.dir &&
        /^xl\/(?:drawings|media|embeddings)\/|^xl\/comments[^/]*\.xml$/.test(file.name),
    ),
    `${api} ${format}: drawing and comment parts removed`,
  );
  if (format === "xlsm") {
    assert.equal(await zip.file("xl/vbaProject.bin")?.async("string"), "test VBA project");
  }
}

try {
  const packResult = parsePackResult(
    run(
      pnpmCommand,
      ["--config.ignore-scripts=true", "pack", "--json", "--pack-destination", temporaryDirectory],
      { capture: true },
    ),
  );
  const packagedPaths = new Set(packResult.files.map((file) => file.path));
  const requiredPaths = [
    "CHANGELOG.md",
    "LICENSE",
    "README.md",
    "package.json",
    ...readdirSync(join(projectDirectory, "src"), { recursive: true, encoding: "utf8" })
      .filter((path) => path.endsWith(".ts") && !path.endsWith(".d.ts"))
      .flatMap((path) => {
        const modulePath = path.slice(0, -3).replaceAll("\\", "/");
        return [`dist/${modulePath}.js`, `dist/${modulePath}.js.map`, `dist/${modulePath}.d.ts`];
      }),
  ];
  for (const requiredPath of requiredPaths) {
    if (!packagedPaths.has(requiredPath)) {
      throw new Error(`Published package is missing ${requiredPath}`);
    }
  }
  if (
    [...packagedPaths].some(
      (path) => path.startsWith("src/") || path.startsWith("test/") || path.startsWith("scripts/"),
    )
  ) {
    throw new Error("Published package unexpectedly contains development sources");
  }
  if ([...packagedPaths].some((path) => path.endsWith(".d.ts.map"))) {
    throw new Error("Published package unexpectedly contains stale declaration maps");
  }

  const tarballPath = isAbsolute(packResult.filename)
    ? packResult.filename
    : join(temporaryDirectory, packResult.filename);
  const projectPackage: unknown = JSON.parse(
    readFileSync(join(projectDirectory, "package.json"), "utf8"),
  );
  if (!isRecord(projectPackage) || typeof projectPackage.packageManager !== "string") {
    throw new Error("Project package.json must specify the package manager used for verification");
  }
  writeFileSync(
    join(temporaryDirectory, "package.json"),
    `${JSON.stringify({ private: true, type: "module", packageManager: projectPackage.packageManager }, undefined, 2)}\n`,
  );
  run(
    pnpmCommand,
    ["add", "--ignore-scripts", "--save-exact", tarballPath, "@types/node@22", "typescript@7"],
    { cwd: temporaryDirectory },
  );

  await writeRuntimeConsumer();
  run(process.execPath, ["consumer.mjs"], { cwd: temporaryDirectory });
  await Promise.all(
    ["xls", "xlsx", "xlsm"].flatMap((format) =>
      ["buffer", "file"].map((api) => verifySanitizedOutput(format, api)),
    ),
  );

  writeTypeConsumer();
  run(
    process.execPath,
    [join(temporaryDirectory, "node_modules/typescript/bin/tsc"), "--project", "tsconfig.json"],
    { cwd: temporaryDirectory },
  );
} finally {
  rmSync(temporaryDirectory, { recursive: true, force: true });
}
