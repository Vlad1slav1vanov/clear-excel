import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { isAbsolute, join } from "node:path";
import { fileURLToPath } from "node:url";

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
    "dist/index.d.ts",
    "dist/index.d.ts.map",
    "dist/index.js",
    "dist/index.js.map",
    "package.json",
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

  const tarballPath = isAbsolute(packResult.filename)
    ? packResult.filename
    : join(temporaryDirectory, packResult.filename);
  writeFileSync(
    join(temporaryDirectory, "package.json"),
    `${JSON.stringify({ private: true, type: "module" }, undefined, 2)}\n`,
  );
  run(
    pnpmCommand,
    ["add", "--ignore-scripts", "--save-exact", tarballPath, "@types/node@22", "typescript@7"],
    { cwd: temporaryDirectory },
  );

  run(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      "import { sanitizeExcel, sanitizeExcelFile, isExcelFilename, ExcelSanitizationError, DEFAULT_LIMITS } from 'clear-excel'; if (typeof sanitizeExcel !== 'function' || typeof sanitizeExcelFile !== 'function' || !isExcelFilename('input.xlsm') || typeof ExcelSanitizationError !== 'function' || DEFAULT_LIMITS.maxEntryCount !== 10000) process.exit(1);",
    ],
    { cwd: temporaryDirectory },
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
