# clear-excel

`clear-excel` safely removes drawing-backed content and unused cell styles from
OOXML `.xlsx` workbooks while preserving worksheet data and structure. It is a
plain Node.js library with no framework or build-time requirement.

## What it removes

- drawings and images;
- classic and threaded comments;
- controls, OLE objects, VML drawings, and worksheet pictures;
- threaded-comment person metadata;
- unused `cellXfs` style records.

Values, formulas and cached results, merged cells, worksheets, and styles that
are still referenced by cells, rows, or columns are retained.

## Requirements

- Node.js 22 or newer;
- OOXML `.xlsx` input.

Legacy `.xls` files and macro-enabled `.xlsm` workbooks are intentionally not
supported.

## Installation

```sh
npm install clear-excel
```

## Buffer API

ES modules:

```js
import { readFile, writeFile } from "node:fs/promises";
import { sanitizeExcel } from "clear-excel";

const input = await readFile("report.xlsx");
const output = await sanitizeExcel(input);
await writeFile("report.cleaned.xlsx", output);
```

CommonJS:

```js
const { readFile, writeFile } = require("node:fs/promises");
const { sanitizeExcel } = require("clear-excel");

const input = await readFile("report.xlsx");
const output = await sanitizeExcel(input);
await writeFile("report.cleaned.xlsx", output);
```

`sanitizeExcel` accepts a Node.js `Buffer`, returns a new `Buffer`, and never
modifies the input buffer.

## File API

```js
import { sanitizeExcelFile } from "clear-excel";

await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx");
```

The input and output paths must be different and end in `.xlsx`. The output
directory must already exist. An existing output file is not replaced unless
`overwrite: true` is explicitly supplied:

```js
await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx", {
  overwrite: true,
  requireSingleWorksheet: true,
});
```

## Options

Both sanitizing functions accept these options:

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `requireSingleWorksheet` | `boolean` | `false` | Reject workbooks that do not contain exactly one sheet. |
| `limits` | `object` | See below | Override individual ZIP/XML safety limits. |

`sanitizeExcelFile` additionally accepts `overwrite`, which defaults to
`false`.

The exported `DEFAULT_LIMITS` object contains the defaults:

| Limit | Default |
| --- | ---: |
| `maxEntryCount` | 10,000 entries |
| `maxUncompressedBytes` | 512 MiB |
| `maxXmlBytes` | 512 MiB aggregate XML output |
| `maxWorksheetXmlBytes` | 256 MiB per worksheet |
| `maxStylesXmlBytes` | 64 MiB |
| `maxMetadataXmlBytes` | 16 MiB per metadata part |
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
increase memory usage.

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

| Code | Meaning |
| --- | --- |
| `ERR_INVALID_ARGUMENT` | An argument or option has an invalid type or value. |
| `ERR_UNSUPPORTED_EXTENSION` | A file path does not end in `.xlsx`. |
| `ERR_INVALID_XLSX` | The ZIP/OOXML package is malformed or unsupported. |
| `ERR_LIMIT_EXCEEDED` | A configured ZIP or XML limit was exceeded. |
| `ERR_WORKSHEET_COUNT` | Exactly one worksheet was required. |
| `ERR_SAME_PATH` | Input and output resolve to the same path. |
| `ERR_OUTPUT_EXISTS` | Output exists and `overwrite` is false. |
| `ERR_FILE_READ` | The input could not be read. |
| `ERR_FILE_WRITE` | The sanitized output could not be written. |

The original technical error is available through `error.cause` when one
exists. Limit and worksheet-count failures also include structured `details`.

## Security and memory model

Before loading an archive, `clear-excel` validates the ZIP central directory,
rejects ZIP64 and multi-disk archives, rejects unsafe entry paths, and checks
declared entry counts and expansion sizes. Required XML parts are inflated
lazily with both per-part and aggregate runtime budgets, including protection
when ZIP size metadata has been forged.

JSZip still loads the ZIP directory and regenerates the complete workbook in
memory. Plan for memory usage above the input size, especially when increasing
the defaults. The sanitizer does not evaluate formulas or execute embedded
content.

## Publishing

Maintainers can verify the package locally with:

```sh
npm ci
npm run check
npm login
npm publish
```

`npm run check` validates syntax, runs the tests, inspects `npm pack --dry-run`,
and installs a temporary tarball to smoke-test both ESM and CommonJS exports.
`prepublishOnly` runs the same checks automatically. npm credentials and
publishing automation are deliberately not stored in this repository.

## License

MIT
