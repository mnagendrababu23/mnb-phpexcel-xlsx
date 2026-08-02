# Changelog

## 2.0.4
- Added unified XLSX metadata schema `1.0` through `Xlsx::metaInfo()` and `ReadSession::metaInfo()`.
- Added quick, standard, full, and forensic metadata profiles covering properties, workbook structure, security, macros, objects, links, hidden content, comments, embedded objects, calculations, print settings, validation, pivots, and package XML.
- Added atomic `Xlsx::updateMetaInfo()` and `Xlsx::removePersonalInfo()` APIs with typed custom properties, workbook visibility/active-sheet settings, calculation settings, encrypted-file preservation, and unknown-part preservation.
- Added missing `fileInfo()`, `sheetsInfo()`, `sheetInfo()`, `rowCount()`, and `rowCounts()` facade methods.
- Preserved arbitrary OOXML namespace prefixes during metadata updates and extended integrity validation to recognize prefixed content-type and relationship elements.
- Added ordinary, encrypted, in-place, unknown-part, personal-cleanup, and namespace-prefix regression tests.
- Fixed metadata-part registration for minimal OOXML packages and added integrity checks for orphaned document-property parts.
- Preserved untouched custom-property native types, property IDs, vectors, self-closing values, and unsigned 64-bit values without lossy coercion.
- Added strict metadata value validation, accurate unscanned counts, encrypted-container reporting, and `xlsx_password` compatibility.
- Raised the minimum core dependency to `^2.0.5`.

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
