# Changelog

All notable changes to this project are documented in this file. The project
follows [Semantic Versioning](https://semver.org/).

## [1.0.4] - 2026-10-09

### Fixed

- Preserve existing output files when exclusive creation fails and remove only
  partial files created by the current write.
- Reject source-file aliases through symbolic links and hard links, and preserve
  existing output permissions during atomic overwrite.
- Preserve worksheet text, CDATA, XML comments, and unrelated elements during
  namespace-aware drawing and comment cleanup.
- Resolve sheet relationship IDs by namespace URI and reject missing or invalid
  relationships before pruning styles.
- Decode UTF-8 and UTF-16 XML consistently and update encoding declarations when
  writing UTF-8 output.
- Enforce the actual aggregate uncompressed size of all retained ZIP parts,
  including binary VBA projects with forged ZIP size metadata.

### Changed

- Replace Vite with one NodeNext TypeScript build for modular ESM output and
  declarations, eliminating the duplicate bundled `cfb` dependency.
- Clean build artifacts before compilation, embed sources in JavaScript maps,
  and stop publishing declaration maps that reference unavailable source files.
- Move the transitive `uuid` override to pnpm workspace configuration.
- Verify the installed package with real `.xls`, `.xlsx`, and `.xlsm` workbooks
  through both APIs, including retained values, formulas, merges, and macros.

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

[1.0.4]: https://www.npmjs.com/package/clear-excel/v/1.0.4
[1.0.0]: https://www.npmjs.com/package/clear-excel/v/1.0.0
[0.1.0]: https://www.npmjs.com/package/clear-excel/v/0.1.0
