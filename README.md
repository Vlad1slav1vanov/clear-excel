# clear-excel

Remove images, drawings, comments, controls, and embedded objects from `.xls`,
`.xlsx`, and `.xlsm` workbooks. Cell values, formulas, merged cells, worksheets,
and used styles are preserved. Unused styles are also removed from `.xlsx` and
`.xlsm` workbooks.

## Install

```sh
npm install clear-excel
```

Requires Node.js 22 or newer. ESM only; TypeScript declarations are included.

## Usage

### Files

```js
import { sanitizeExcelFile } from "clear-excel";

await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx");
```

`sanitizeExcelFile(inputPath, outputPath, options?)` returns `Promise<void>`.
Input and output must refer to different files and use the same supported
extension (`.xls`, `.xlsx`, or `.xlsm`, case-insensitive). The output directory
must already exist. To replace an existing output file, set `overwrite: true`.

### Buffers

```js
import { readFile, writeFile } from "node:fs/promises";
import { sanitizeExcel } from "clear-excel";

const input = await readFile("report.xlsx");
const output = await sanitizeExcel(input);
await writeFile("report.cleaned.xlsx", output);
```

`sanitizeExcel(buffer, options?)` accepts a Node.js `Buffer` and returns
`Promise<Buffer>`. The input buffer is not modified; the output keeps the
original workbook format.

### Filename checks

The package also exports `isExcelFilename(name)` for `.xls`, `.xlsx`, and `.xlsm`,
and `isXlsxFilename(name)` for `.xlsx` only. Both return a boolean, ignore case,
and check the extension without reading the file.

## Options

| Option                   | Default          | Description                                             |
| ------------------------ | ---------------- | ------------------------------------------------------- |
| `requireSingleWorksheet` | `false`          | Reject workbooks that do not contain exactly one sheet. |
| `limits`                 | `DEFAULT_LIMITS` | Override individual size and entry-count limits.        |
| `overwrite`              | `false`          | Replace an existing output; `sanitizeExcelFile` only.   |

```js
import { sanitizeExcelFile } from "clear-excel";

await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx", {
  overwrite: true,
  requireSingleWorksheet: true,
  limits: {
    maxUncompressedBytes: 128 * 1024 * 1024,
  },
});
```

### Limits

The exported `DEFAULT_LIMITS` contains these defaults. Only specified values
are overridden; each must be a positive safe integer. Sizes are in bytes.

| Limit                      |                      Default |
| -------------------------- | ---------------------------: |
| `maxEntryCount`            |               10,000 entries |
| `maxUncompressedBytes`     |                      512 MiB |
| `maxXmlBytes`              |  512 MiB aggregate XML reads |
| `maxWorksheetXmlBytes`     |        256 MiB per worksheet |
| `maxStylesXmlBytes`        |                       64 MiB |
| `maxMetadataXmlBytes`      |     16 MiB per metadata part |
| `maxRelationshipsXmlBytes` | 8 MiB per relationships part |

For `.xls`, only `maxEntryCount` and `maxUncompressedBytes` apply. XML limits
apply to `.xlsx` and `.xlsm`.

## Errors

Both sanitizing functions reject with `ExcelSanitizationError`. Check `code`
instead of matching the error message. `details` contains additional context
when available, and `cause` holds the original error when one exists.

```js
import { ExcelSanitizationError, sanitizeExcelFile } from "clear-excel";

try {
  await sanitizeExcelFile("report.xlsx", "report.cleaned.xlsx");
} catch (error) {
  if (error instanceof ExcelSanitizationError) {
    console.error(error.code, error.details);
  }
  throw error;
}
```

| Code                        | Meaning                                              |
| --------------------------- | ---------------------------------------------------- |
| `ERR_INVALID_ARGUMENT`      | An argument or option has an invalid type or value.  |
| `ERR_UNSUPPORTED_EXTENSION` | Paths use an unsupported or mismatched extension.    |
| `ERR_INVALID_XLS`           | The XLS workbook is malformed or unsupported.        |
| `ERR_INVALID_XLSX`          | The XLSX/XLSM workbook is malformed or unsupported.  |
| `ERR_LIMIT_EXCEEDED`        | A configured size or entry-count limit was exceeded. |
| `ERR_WORKSHEET_COUNT`       | Exactly one worksheet was required.                  |
| `ERR_SAME_PATH`             | Input and output refer to the same file.             |
| `ERR_OUTPUT_EXISTS`         | Output exists and `overwrite` is false.              |
| `ERR_FILE_READ`             | The input could not be read.                         |
| `ERR_FILE_WRITE`            | The output could not be written.                     |

## Limitations

- `.xls` input must use Excel 97–2003 BIFF8. Encrypted `.xls` files are not supported.
- VBA macros are preserved. The library does not execute macros or evaluate formulas.
- Workbooks are processed in memory. Memory usage can exceed the input size,
  especially when increasing limits.

## License

[MIT](./LICENSE)
