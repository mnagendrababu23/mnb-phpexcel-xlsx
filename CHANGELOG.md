# Changelog

## 2.0.3
- Added native XLSX active-worksheet detection from workbook `activeTab` metadata.
- `inspect()` now reports `active_sheet` with a 1-based index and name.
- Added integration tests for sheet existence, active-sheet selection, and empty worksheets.
- Raised the minimum core dependency to `^2.0.3`.


## 2.0.2
- Added actionable worksheet-not-found diagnostics for direct XLSX reader usage.
- Invalid worksheet indexes and empty names now use the shared core worksheet-selection exception.
- Raised the minimum core dependency to `^2.0.2`.

## 2.0.0
- Declared the required `ext-libxml` and `ext-zlib` runtime dependencies used by native XLSX XML and ZIP processing.

- Coordinated MNB PHPExcel v2 release.
- Internal MNB dependencies aligned to `^2.0`.
- Package boundaries validated for independent installation.
