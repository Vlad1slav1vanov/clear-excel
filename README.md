# clear-excel

`clear-excel` safely removes drawing-backed content from legacy `.xls` and
OOXML `.xlsx`/`.xlsm` workbooks while preserving worksheet data and structure.
It is an ESM-only Node.js library with no framework or consumer-side build
requirement.

## What it removes

- drawings, images, comments, controls, and embedded OLE objects;
- VML drawings, threaded comments, and person metadata in OOXML workbooks;
- unused `cellXfs` style records in OOXML workbooks;
- BIFF drawing, note, object, and picture records plus embedded object storages
  in legacy `.xls` workbooks.

Values, formulas and cached results, merged cells, worksheets, and styles that
are still referenced by cells, rows, or columns are retained. `.xlsm` VBA
projects and legacy `.xls` VBA storages are preserved; the library does not
execute or remove macros.

## Requirements

- Node.js 22 or newer;
- Excel 97-2003 BIFF8 `.xls`, OOXML `.xlsx`, or macro-enabled `.xlsm` input.

Encrypted `.xls` workbooks are not supported.

## Installation

```sh
npm install clear-excel
```

## Buffer API

```js
import { readFile, writeFile } from "node:fs/promises";
import { sanitizeExcel } from "clear-excel";

const input = await readFile("report.xlsx");
const output = await sanitizeExcel(input);
await writeFile("report.cleaned.xlsx", output);
```

`sanitizeExcel` accepts a Node.js `Buffer`, returns a new `Buffer`, and never
modifies the input buffer. It detects legacy CFB/BIFF workbooks by their file
signature and keeps the input format in the returned buffer.

## Version 1 migration

Version 1 is ESM-only. Replace CommonJS `require("clear-excel")` calls with
ECMAScript imports. Runtime exports, option types, error codes, and sanitizing
behavior otherwise remain compatible with version 0.1.

## File API

```js
import { sanitizeExcelFile } from "clear-excel";

await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx");
```

The input and output paths must be different and use the same supported
extension: `.xls`, `.xlsx`, or `.xlsm` (case-insensitive). The file API does not
convert between formats. The output directory must already exist. An existing
output file is not replaced unless `overwrite: true` is explicitly supplied:

```js
await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx", {
  overwrite: true,
  requireSingleWorksheet: true,
});
```

## Options

Both sanitizing functions accept these options:

| Option                   | Type      | Default   | Description                                             |
| ------------------------ | --------- | --------- | ------------------------------------------------------- |
| `requireSingleWorksheet` | `boolean` | `false`   | Reject workbooks that do not contain exactly one sheet. |
| `limits`                 | `object`  | See below | Override individual ZIP/XML safety limits.              |

`sanitizeExcelFile` additionally accepts `overwrite`, which defaults to
`false`.

The exported `DEFAULT_LIMITS` object contains the defaults:

| Limit                      |                      Default |
| -------------------------- | ---------------------------: |
| `maxEntryCount`            |               10,000 entries |
| `maxUncompressedBytes`     |                      512 MiB |
| `maxXmlBytes`              | 512 MiB aggregate XML output |
| `maxWorksheetXmlBytes`     |        256 MiB per worksheet |
| `maxStylesXmlBytes`        |                       64 MiB |
| `maxMetadataXmlBytes`      |     16 MiB per metadata part |
| `maxRelationshipsXmlBytes` | 8 MiB per relationships part |

Only specified values are overridden:

```js
const output = await sanitizeExcel(input, {
  limits: {
    maxUncompressedBytes: 128 * 1024 * 1024,
    maxWorksheetXmlBytes: 64 * 1024 * 1024,
  },
});
```

Every limit must be a positive safe integer. Raising limits can significantly
increase memory usage. For `.xls`, `maxEntryCount` and
`maxUncompressedBytes` apply to the CFB container; the XML-specific limits apply
only to `.xlsx` and `.xlsm`.

## Errors

Public functions reject with `ExcelSanitizationError`. Use its stable `code`
instead of matching the English message:

```js
import { ExcelSanitizationError, sanitizeExcel } from "clear-excel";

try {
  await sanitizeExcel(input);
} catch (error) {
  if (error instanceof ExcelSanitizationError) {
    console.error(error.code, error.details);
  }
  throw error;
}
```

| Code                        | Meaning                                             |
| --------------------------- | --------------------------------------------------- |
| `ERR_INVALID_ARGUMENT`      | An argument or option has an invalid type or value. |
| `ERR_UNSUPPORTED_EXTENSION` | Paths use an unsupported or mismatched extension.   |
| `ERR_INVALID_XLS`           | The CFB/BIFF workbook is malformed or unsupported.  |
| `ERR_INVALID_XLSX`          | The ZIP/OOXML package is malformed or unsupported.  |
| `ERR_LIMIT_EXCEEDED`        | A configured ZIP or XML limit was exceeded.         |
| `ERR_WORKSHEET_COUNT`       | Exactly one worksheet was required.                 |
| `ERR_SAME_PATH`             | Input and output resolve to the same path.          |
| `ERR_OUTPUT_EXISTS`         | Output exists and `overwrite` is false.             |
| `ERR_FILE_READ`             | The input could not be read.                        |
| `ERR_FILE_WRITE`            | The sanitized output could not be written.          |

The original technical error is available through `error.cause` when one
exists. Limit and worksheet-count failures also include structured `details`.

## Security and memory model

Before loading an OOXML archive, `clear-excel` validates the ZIP central
directory, rejects ZIP64 and multi-disk archives, rejects unsafe entry paths,
and checks declared entry counts and expansion sizes. Required XML parts are
inflated lazily with both per-part and aggregate runtime budgets, including
protection when ZIP size metadata has been forged.

For `.xls`, the library validates CFB entry and byte limits, parses the BIFF
record stream without evaluating formulas, removes drawing-layer records, and
rewrites worksheet offsets after records move. Workbook records are not
converted through an OOXML or generic tabular representation.

JSZip still loads the ZIP directory and regenerates the complete workbook in
memory. Plan for memory usage above the input size, especially when increasing
the defaults. The sanitizer does not evaluate formulas, execute embedded
content, or execute VBA projects.

## Publishing

Maintainers can verify the package locally with:

```sh
npm ci
npm run check
npm login
npm publish
```

`npm run check` verifies Oxfmt formatting, runs type-aware Oxlint and strict
TypeScript checks, executes the tests, builds the ESM bundle and declarations,
inspects `npm pack --dry-run`, and installs a temporary tarball to smoke-test
both runtime imports and a TypeScript consumer. `prepack` creates `dist/`, and
`prepublishOnly` runs the complete check automatically. npm credentials and
publishing automation are deliberately not stored in this repository.

Useful development commands:

```sh
npm run format
npm run lint
npm run check:types
npm test
npm run build
```

## License

MIT
