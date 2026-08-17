"use strict";

const { execFileSync } = require("node:child_process");
const { mkdtempSync, rmSync } = require("node:fs");
const { tmpdir } = require("node:os");
const { join } = require("node:path");

const projectDirectory = join(__dirname, "..");
const temporaryDirectory = mkdtempSync(join(tmpdir(), "clear-excel-package-"));
const npmCommand = process.platform === "win32" ? "npm.cmd" : "npm";
const childEnvironment = { ...process.env };
delete childEnvironment.npm_config_dry_run;
delete childEnvironment.NPM_CONFIG_DRY_RUN;
const run = (command, args, options = {}) =>
  execFileSync(command, args, {
    cwd: options.cwd ?? projectDirectory,
    encoding: "utf8",
    env: childEnvironment,
    stdio: options.capture ? "pipe" : "inherit",
  });

try {
  const packOutput = run(
    npmCommand,
    ["pack", "--json", "--ignore-scripts", "--pack-destination", temporaryDirectory],
    { capture: true },
  );
  const [packResult] = JSON.parse(packOutput);
  if (!packResult?.filename || !Array.isArray(packResult.files)) {
    throw new Error("npm pack returned an unexpected result");
  }

  const packagedPaths = new Set(packResult.files.map((file) => file.path));
  const requiredPaths = [
    "LICENSE",
    "README.md",
    "index.cjs",
    "index.d.ts",
    "index.js",
    "package.json",
    "src/index.cjs",
  ];
  for (const requiredPath of requiredPaths) {
    if (!packagedPaths.has(requiredPath)) {
      throw new Error(`Published package is missing ${requiredPath}`);
    }
  }
  if ([...packagedPaths].some((path) => path.startsWith("test/"))) {
    throw new Error("Published package unexpectedly contains tests");
  }

  const tarballPath = join(temporaryDirectory, packResult.filename);
  run(
    npmCommand,
    [
      "install",
      "--ignore-scripts",
      "--no-audit",
      "--no-fund",
      "--package-lock=false",
      tarballPath,
    ],
    { cwd: temporaryDirectory },
  );

  run(
    process.execPath,
    [
      "--input-type=module",
      "--eval",
      "import { sanitizeExcel, sanitizeExcelFile, ExcelSanitizationError, DEFAULT_LIMITS } from 'clear-excel'; if (typeof sanitizeExcel !== 'function' || typeof sanitizeExcelFile !== 'function' || typeof ExcelSanitizationError !== 'function' || DEFAULT_LIMITS.maxEntryCount !== 10000) process.exit(1);",
    ],
    { cwd: temporaryDirectory },
  );
  run(
    process.execPath,
    [
      "--eval",
      "const { sanitizeExcel, sanitizeExcelFile, ExcelSanitizationError, DEFAULT_LIMITS } = require('clear-excel'); if (typeof sanitizeExcel !== 'function' || typeof sanitizeExcelFile !== 'function' || typeof ExcelSanitizationError !== 'function' || DEFAULT_LIMITS.maxEntryCount !== 10000) process.exit(1);",
    ],
    { cwd: temporaryDirectory },
  );
} finally {
  rmSync(temporaryDirectory, { recursive: true, force: true });
}
