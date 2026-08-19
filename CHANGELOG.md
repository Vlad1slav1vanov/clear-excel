# Changelog

All notable changes to this project are documented in this file. The project
follows [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-08-19

### Changed

- Rebuilt the package as a strict TypeScript, ESM-only library with bundled
  declarations.
- Documented the CommonJS-to-ESM migration while keeping the public sanitizer
  APIs, options, error codes, and behavior compatible with 0.1.
- Added package-content verification and CI coverage for supported Node.js
  versions.

## [0.1.0] - 2026-08-17

### Added

- Initial release with buffer and file APIs for sanitizing `.xls`, `.xlsx`,
  and `.xlsm` workbooks.

[1.0.0]: https://www.npmjs.com/package/clear-excel/v/1.0.0
[0.1.0]: https://www.npmjs.com/package/clear-excel/v/0.1.0
