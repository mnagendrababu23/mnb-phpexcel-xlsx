<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use DateTimeImmutable;
use DateTimeInterface;
use Mnb\PHPExcel\Application\LoggerBridge;
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Events\EventDispatcher;
use Mnb\PHPExcel\Security\FormulaGuard;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\ValueSanitizer;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use Throwable;
use ZipArchive;

/**
 * Low-memory XLSX/CSV-ZIP writer for very large tabular exports.
 *
 * It streams worksheet XML to temporary files and then packages those parts into
 * XLSX. It intentionally supports a focused, safe subset of Excel features:
 * inline strings, basic values, formulas, basic number/date/percent/currency/text
 * formats, header rows, auto-filter/freeze panes, progress callbacks, and sheet
 * splitting. Rich workbook objects belong to the normal small-file writer.
 */
final class LargeXlsxStreamingWriter
{
    private const EXCEL_MAX_ROWS = 1048576;
    private const EXCEL_MAX_COLUMNS = 16384;

    /** @param iterable<array<int|string,mixed>> $rows @param array<string,mixed> $options @return array<string,mixed> */
    public function write(iterable $rows, string $path, array $options = []): array
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode(
                'ext-zip is required to stream XLSX exports.',
                ErrorCode::EXTENSION_MISSING,
                ['extension' => 'zip']
            );
        }

        $summary = [];
        AtomicFileWriter::writeViaTemp(
            $path,
            function (string $tmp) use ($rows, $options, &$summary): void {
                $summary = $this->writePackage($rows, $tmp, $options);
            },
            function (string $tmp) use ($options): void {
                if (($options['validate_integrity'] ?? true) === false) {
                    return;
                }
                (new XlsxIntegrityValidator())->assertValid($tmp, is_array($options['integrity_options'] ?? null) ? $options['integrity_options'] : []);
            }
        );

        $summary['status'] = 'completed';
        $summary['path'] = $path;
        return $summary;
    }

    /** @param iterable<array<int|string,mixed>> $rows @param array<string,mixed> $options @return array<string,mixed> */
    public function saveCsvZip(iterable $rows, string $path, array $options = []): array
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode(
                'ext-zip is required to write CSV ZIP exports.',
                ErrorCode::EXTENSION_MISSING,
                ['extension' => 'zip']
            );
        }

        $summary = [];
        AtomicFileWriter::writeViaTemp($path, function (string $tmp) use ($rows, $options, &$summary): void {
            $summary = $this->writeCsvZipPackage($rows, $tmp, $options);
        });

        $summary['status'] = 'completed';
        $summary['path'] = $path;
        return $summary;
    }

    /** @param iterable<array<int|string,mixed>> $rows @param array<string,mixed> $options @return array<string,mixed> */
    private function writePackage(iterable $rows, string $path, array $options): array
    {
        $workDir = $this->makeWorkDir($path);
        $sheets = [];
        $startedAt = microtime(true);
        $rowsExported = 0;
        $dataRowsOnCurrentSheet = 0;
        $currentSheet = null;
        $headers = $this->normalizeHeader($options['headers'] ?? $options['header'] ?? null);
        $withHeader = (bool) ($options['with_header'] ?? ($headers !== []));
        $options['with_header'] = $withHeader;
        $headerResolved = $headers !== [];
        $sheetBaseName = (string) ($options['sheet_name'] ?? $options['sheet'] ?? 'Sheet');
        $autoSplit = (bool) ($options['auto_split_sheets'] ?? true);
        $maxRowsPerSheet = min(self::EXCEL_MAX_ROWS, max(1, (int) ($options['max_rows_per_sheet'] ?? self::EXCEL_MAX_ROWS)));
        $maxDataRowsPerSheet = $withHeader ? max(1, $maxRowsPerSheet - 1) : $maxRowsPerSheet;
        $progressEvery = max(1, (int) ($options['progress_every'] ?? $options['progress_interval'] ?? 1000));
        $columnFormats = is_array($options['column_formats'] ?? null) ? $options['column_formats'] : [];

        EventDispatcher::safeDispatch('before_large_export', ['path' => $path, 'format' => 'xlsx']);
        LoggerBridge::info('Large XLSX streaming export started.', ['path' => $path]);

        try {
            foreach ($rows as $sourceRow) {
                $sourceRow = $this->normalizeRow($sourceRow);
                if (!$headerResolved && $withHeader && $this->isAssociative($sourceRow)) {
                    $headers = array_map(static fn(mixed $key): string => (string) $key, array_keys($sourceRow));
                    $headerResolved = true;
                }

                if ($currentSheet === null) {
                    $currentSheet = $this->openSheet($workDir, count($sheets) + 1, $this->sheetName($sheetBaseName, count($sheets) + 1), $options);
                    if ($withHeader && $headers !== []) {
                        $this->writeRow($currentSheet, $headers, $headers, $columnFormats, $options, true);
                    }
                }

                if ($dataRowsOnCurrentSheet >= $maxDataRowsPerSheet) {
                    if (!$autoSplit) {
                        throw MnbExcelException::withCode(
                            'Large XLSX export exceeds Excel row limit for one sheet. Enable auto_split_sheets or use CSV ZIP fallback.',
                            ErrorCode::XLSX_WRITE_FAILED,
                            ['max_rows_per_sheet' => $maxRowsPerSheet]
                        );
                    }
                    $sheets[] = $this->closeSheet($currentSheet, $options);
                    $currentSheet = $this->openSheet($workDir, count($sheets) + 1, $this->sheetName($sheetBaseName, count($sheets) + 1), $options);
                    $dataRowsOnCurrentSheet = 0;
                    if ($withHeader && $headers !== []) {
                        $this->writeRow($currentSheet, $headers, $headers, $columnFormats, $options, true);
                    }
                }

                $ordered = $this->orderRow($sourceRow, $headers);
                $this->writeRow($currentSheet, $ordered, $headers, $columnFormats, $options, false);
                $rowsExported++;
                $dataRowsOnCurrentSheet++;

                if ($rowsExported % $progressEvery === 0) {
                    $this->emitProgress($options, [
                        'path' => $path,
                        'format' => 'xlsx',
                        'rows_exported' => $rowsExported,
                        'sheets_created' => count($sheets) + 1,
                        'current_sheet_rows' => $dataRowsOnCurrentSheet,
                        'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
                    ]);
                }
            }

            if ($currentSheet === null) {
                $currentSheet = $this->openSheet($workDir, 1, $this->sheetName($sheetBaseName, 1), $options);
                if ($withHeader && $headers !== []) {
                    $this->writeRow($currentSheet, $headers, $headers, $columnFormats, $options, true);
                }
            }
            $sheets[] = $this->closeSheet($currentSheet, $options);

            $this->createZipPackage($path, $workDir, $sheets, $options);
        } catch (Throwable $e) {
            EventDispatcher::safeDispatch('on_large_export_failed', ['path' => $path, 'format' => 'xlsx', 'exception' => $e]);
            LoggerBridge::error('Large XLSX streaming export failed.', ['path' => $path, 'message' => $e->getMessage()]);
            throw $e;
        } finally {
            $this->removeDirectory($workDir);
        }

        $summary = [
            'status' => 'completed',
            'format' => 'xlsx',
            'rows_exported' => $rowsExported,
            'sheets_created' => count($sheets),
            'headers' => $headers,
            'auto_split_sheets' => $autoSplit,
            'max_rows_per_sheet' => $maxRowsPerSheet,
            'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
            'memory_peak_mb' => round(memory_get_peak_usage(true) / 1048576, 2),
        ];
        EventDispatcher::safeDispatch('after_large_export', $summary + ['path' => $path]);
        LoggerBridge::info('Large XLSX streaming export completed.', $summary + ['path' => $path]);
        return $summary;
    }

    /** @param iterable<array<int|string,mixed>> $rows @param array<string,mixed> $options @return array<string,mixed> */
    private function writeCsvZipPackage(iterable $rows, string $path, array $options): array
    {
        $workDir = $this->makeWorkDir($path);
        $startedAt = microtime(true);
        $rowsExported = 0;
        $parts = [];
        $partRows = 0;
        $partHandle = null;
        $partPath = '';
        $partNumber = 0;
        $headers = $this->normalizeHeader($options['headers'] ?? $options['header'] ?? null);
        $withHeader = (bool) ($options['with_header'] ?? ($headers !== []));
        $options['with_header'] = $withHeader;
        $headerResolved = $headers !== [];
        $rowsPerFile = max(1, (int) ($options['rows_per_file'] ?? $options['max_rows_per_file'] ?? 1000000));
        $delimiter = (string) ($options['delimiter'] ?? ',');
        $enclosure = (string) ($options['enclosure'] ?? '"');
        $escape = (string) ($options['escape'] ?? '\\');
        $progressEvery = max(1, (int) ($options['progress_every'] ?? $options['progress_interval'] ?? 1000));

        $openPart = function () use (&$partHandle, &$partPath, &$partNumber, &$partRows, &$parts, $workDir, $withHeader, &$headers, $delimiter, $enclosure, $escape): void {
            if (is_resource($partHandle)) {
                fclose($partHandle);
                $parts[] = ['path' => $partPath, 'name' => basename($partPath), 'rows' => $partRows];
            }
            $partNumber++;
            $partRows = 0;
            $partPath = $workDir . DIRECTORY_SEPARATOR . 'part-' . str_pad((string) $partNumber, 4, '0', STR_PAD_LEFT) . '.csv';
            $partHandle = fopen($partPath, 'wb');
            if (!is_resource($partHandle)) {
                throw MnbExcelException::withCode('Unable to create CSV ZIP part: ' . $partPath, ErrorCode::CSV_WRITE_FAILED, ['path' => $partPath]);
            }
            fwrite($partHandle, "\xEF\xBB\xBF");
            if ($withHeader && $headers !== []) {
                fputcsv($partHandle, $headers, $delimiter, $enclosure, $escape);
            }
        };

        EventDispatcher::safeDispatch('before_large_export', ['path' => $path, 'format' => 'csv_zip']);
        try {
            foreach ($rows as $sourceRow) {
                $sourceRow = $this->normalizeRow($sourceRow);
                if (!$headerResolved && $withHeader && $this->isAssociative($sourceRow)) {
                    $headers = array_map(static fn(mixed $key): string => (string) $key, array_keys($sourceRow));
                    $headerResolved = true;
                }
                if (!is_resource($partHandle) || $partRows >= $rowsPerFile) {
                    $openPart();
                }
                $ordered = $this->orderRow($sourceRow, $headers);
                $csvRow = array_map(fn(mixed $value): mixed => ValueSanitizer::sanitizeFormulaLikeText($this->displayValue($value), (string) ($options['csv_injection_policy'] ?? 'escape')), $ordered);
                if (fputcsv($partHandle, $csvRow, $delimiter, $enclosure, $escape) === false) {
                    throw MnbExcelException::withCode('Unable to write CSV ZIP row.', ErrorCode::CSV_WRITE_FAILED);
                }
                $rowsExported++;
                $partRows++;
                if ($rowsExported % $progressEvery === 0) {
                    $this->emitProgress($options, [
                        'path' => $path,
                        'format' => 'csv_zip',
                        'rows_exported' => $rowsExported,
                        'parts_created' => $partNumber,
                        'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
                    ]);
                }
            }
            if (!is_resource($partHandle)) {
                $openPart();
            }
            if (is_resource($partHandle)) {
                fclose($partHandle);
                $partHandle = null;
                $parts[] = ['path' => $partPath, 'name' => basename($partPath), 'rows' => $partRows];
            }

            $zip = new ZipArchive();
            $opened = $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
            if ($opened !== true) {
                throw MnbExcelException::withCode('Unable to create CSV ZIP file: ' . $path, ErrorCode::FILE_WRITE_FAILED, ['zip_status' => $opened]);
            }
            foreach ($parts as $part) {
                if ($zip->addFile((string) $part['path'], (string) $part['name']) !== true) {
                    throw MnbExcelException::withCode('Unable to add CSV part to ZIP: ' . $part['name'], ErrorCode::FILE_WRITE_FAILED);
                }
            }
            $manifest = json_encode([
                'format' => 'csv_zip',
                'rows_exported' => $rowsExported,
                'parts' => array_map(static fn(array $p): array => ['name' => $p['name'], 'rows' => $p['rows']], $parts),
                'created_at' => gmdate('c'),
            ], JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
            $zip->addFromString('manifest.json', $manifest ?: '{}');
            if ($zip->close() !== true) {
                throw MnbExcelException::withCode('Unable to finalize CSV ZIP file: ' . $path, ErrorCode::FILE_WRITE_FAILED);
            }
        } catch (Throwable $e) {
            if (is_resource($partHandle)) {
                fclose($partHandle);
            }
            EventDispatcher::safeDispatch('on_large_export_failed', ['path' => $path, 'format' => 'csv_zip', 'exception' => $e]);
            throw $e;
        } finally {
            $this->removeDirectory($workDir);
        }

        return [
            'status' => 'completed',
            'format' => 'csv_zip',
            'rows_exported' => $rowsExported,
            'parts_created' => count($parts),
            'rows_per_file' => $rowsPerFile,
            'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
            'memory_peak_mb' => round(memory_get_peak_usage(true) / 1048576, 2),
        ];
    }

    /** @param resource $handle */
    private function writeRow(array &$sheet, array $values, array $headers, array $columnFormats, array $options, bool $isHeader): void
    {
        $sheet['row']++;
        $rowNumber = (int) $sheet['row'];
        $handle = $sheet['handle'];
        $this->writeXml($handle, '<row r="' . $rowNumber . '">');
        $colIndex = 0;
        foreach (array_values($values) as $value) {
            $colIndex++;
            if ($colIndex > self::EXCEL_MAX_COLUMNS) {
                throw MnbExcelException::withCode('XLSX export exceeds Excel column limit.', ErrorCode::XLSX_WRITE_FAILED);
            }
            $cellRef = Coordinate::columnIndexToName($colIndex) . $rowNumber;
            $format = $isHeader ? null : $this->formatForColumn($colIndex, $headers[$colIndex - 1] ?? null, $columnFormats);
            $this->writeCell($handle, $cellRef, $value, $format, $options, $isHeader);
        }
        $this->writeXml($handle, '</row>');
        $sheet['max_col'] = max((int) $sheet['max_col'], $colIndex);
    }

    /** @param resource $handle */
    private function writeCell($handle, string $cellRef, mixed $value, ?string $format, array $options, bool $isHeader): void
    {
        $value = ValueSanitizer::normalizeScalar($value);
        if ($value instanceof CellValue && $value->type() === CellValue::TYPE_BLANK) {
            return;
        }
        if ($value === null || $value === '') {
            return;
        }

        $style = $isHeader ? 8 : $this->styleIdForFormat($format);
        $styleAttr = $style > 0 ? ' s="' . $style . '"' : '';

        if ($value instanceof CellValue) {
            match ($value->type()) {
                CellValue::TYPE_TEXT => $this->writeInlineString($handle, $cellRef, (string) $value->value(), $styleAttr, $options),
                CellValue::TYPE_NUMBER => $this->writeNumberCell($handle, $cellRef, (string) $value->value(), $styleAttr),
                CellValue::TYPE_BOOLEAN => $this->writeXml($handle, '<c r="' . $cellRef . '" t="b"' . $styleAttr . '><v>' . ($value->value() ? '1' : '0') . '</v></c>'),
                CellValue::TYPE_DATE => $this->writeNumberCell($handle, $cellRef, $this->excelDateSerial($value->value()), ' s="' . ($this->styleIdForFormat((string) ($value->options()['format'] ?? 'date')) ?: 4) . '"'),
                CellValue::TYPE_FORMULA => $this->writeFormulaCell($handle, $cellRef, (string) $value->value(), $value->cachedValue(), $styleAttr, $options),
                CellValue::TYPE_ERROR => $this->writeXml($handle, '<c r="' . $cellRef . '" t="e"' . $styleAttr . '><v>' . $this->esc((string) $value->value()) . '</v></c>'),
                default => null,
            };
            return;
        }

        if ($format === 'text') {
            $this->writeInlineString($handle, $cellRef, (string) $value, ' s="7"', $options);
            return;
        }
        if (in_array($format, ['date', 'datetime'], true)) {
            $this->writeNumberCell($handle, $cellRef, $this->excelDateSerial($value), ' s="' . $this->styleIdForFormat($format) . '"');
            return;
        }
        if (in_array($format, ['integer', 'int', 'decimal', 'number', 'float', 'currency', 'money', 'percent', 'percentage'], true) && is_string($value) && is_numeric($value)) {
            $this->writeNumberCell($handle, $cellRef, $value, $styleAttr);
            return;
        }
        if (is_bool($value)) {
            $this->writeXml($handle, '<c r="' . $cellRef . '" t="b"' . $styleAttr . '><v>' . ($value ? '1' : '0') . '</v></c>');
            return;
        }
        if (is_int($value) || is_float($value)) {
            $this->writeNumberCell($handle, $cellRef, $this->serializeNumber($value), $styleAttr);
            return;
        }
        $this->writeInlineString($handle, $cellRef, (string) $value, $styleAttr, $options);
    }

    /** @param resource $handle */
    private function writeFormulaCell($handle, string $cellRef, string $formula, mixed $cached, string $styleAttr, array $options): void
    {
        FormulaGuard::assertSafe($formula, $options);
        if (str_starts_with($formula, '=')) {
            $formula = substr($formula, 1);
        }
        if ($cached === null || $cached === '') {
            $this->writeXml($handle, '<c r="' . $cellRef . '"' . $styleAttr . '><f>' . $this->esc($formula) . '</f></c>');
            return;
        }
        if (is_bool($cached)) {
            $this->writeXml($handle, '<c r="' . $cellRef . '" t="b"' . $styleAttr . '><f>' . $this->esc($formula) . '</f><v>' . ($cached ? '1' : '0') . '</v></c>');
            return;
        }
        if (is_int($cached) || is_float($cached) || (is_string($cached) && is_numeric($cached))) {
            $this->writeXml($handle, '<c r="' . $cellRef . '"' . $styleAttr . '><f>' . $this->esc($formula) . '</f><v>' . $this->serializeNumber($cached) . '</v></c>');
            return;
        }
        $this->writeXml($handle, '<c r="' . $cellRef . '" t="str"' . $styleAttr . '><f>' . $this->esc($formula) . '</f><v>' . $this->esc((string) $cached) . '</v></c>');
    }

    /** @param resource $handle */
    private function writeInlineString($handle, string $cellRef, string $value, string $styleAttr, array $options): void
    {
        $policy = (string) ($options['formula_text_policy'] ?? $options['csv_injection_policy'] ?? 'escape');
        $value = (string) ValueSanitizer::sanitizeFormulaLikeText($value, $policy);
        $value = ValueSanitizer::sanitizeCellText($value, $options);
        $space = trim($value) !== $value ? ' xml:space="preserve"' : '';
        $this->writeXml($handle, '<c r="' . $cellRef . '" t="inlineStr"' . $styleAttr . '><is><t' . $space . '>' . $this->esc($value) . '</t></is></c>');
    }

    /** @param resource $handle */
    private function writeNumberCell($handle, string $cellRef, int|float|string $number, string $styleAttr): void
    {
        $this->writeXml($handle, '<c r="' . $cellRef . '"' . $styleAttr . '><v>' . $this->serializeNumber($number) . '</v></c>');
    }

    /** @return array<string,mixed> */
    private function openSheet(string $workDir, int $sheetNumber, string $name, array $options): array
    {
        $sheetPath = $workDir . DIRECTORY_SEPARATOR . 'sheet' . $sheetNumber . '.xml';
        $handle = fopen($sheetPath, 'wb');
        if (!is_resource($handle)) {
            throw MnbExcelException::withCode('Unable to create temporary sheet XML: ' . $sheetPath, ErrorCode::FILE_WRITE_FAILED);
        }
        $this->writeXml($handle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
        $this->writeXml($handle, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
        if (($options['freeze_header'] ?? true) !== false && ($options['with_header'] ?? true)) {
            $this->writeXml($handle, '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>');
        }
        $this->writeXml($handle, '<sheetData>');
        return ['number' => $sheetNumber, 'name' => $name, 'path' => $sheetPath, 'handle' => $handle, 'row' => 0, 'max_col' => 0];
    }

    /** @param array<string,mixed> $sheet @return array<string,mixed> */
    private function closeSheet(array $sheet, array $options): array
    {
        $handle = $sheet['handle'];
        $this->writeXml($handle, '</sheetData>');
        if (($options['auto_filter'] ?? true) !== false && ($options['with_header'] ?? false) && (int) $sheet['row'] > 0 && (int) $sheet['max_col'] > 0) {
            $end = Coordinate::columnIndexToName((int) $sheet['max_col']) . max(1, (int) $sheet['row']);
            $this->writeXml($handle, '<autoFilter ref="A1:' . $end . '"/>');
        }
        $this->writeXml($handle, '</worksheet>');
        fclose($handle);
        unset($sheet['handle']);
        return $sheet;
    }

    /** @param list<array<string,mixed>> $sheets @param array<string,mixed> $options */
    private function createZipPackage(string $path, string $workDir, array $sheets, array $options): void
    {
        $zip = new ZipArchive();
        $opened = $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        if ($opened !== true) {
            throw MnbExcelException::withCode('Unable to create streaming XLSX zip package: ' . $path, ErrorCode::XLSX_WRITE_FAILED, ['zip_status' => $opened]);
        }

        try {
            $this->zipString($zip, '[Content_Types].xml', $this->contentTypesXml(count($sheets)));
            $this->zipString($zip, '_rels/.rels', $this->rootRelsXml());
            $this->zipString($zip, 'docProps/core.xml', $this->corePropsXml($options));
            $this->zipString($zip, 'docProps/app.xml', $this->appPropsXml());
            $this->zipString($zip, 'xl/workbook.xml', $this->workbookXml($sheets));
            $this->zipString($zip, 'xl/_rels/workbook.xml.rels', $this->workbookRelsXml(count($sheets)));
            $this->zipString($zip, 'xl/styles.xml', $this->stylesXml());
            foreach ($sheets as $sheet) {
                $entry = 'xl/worksheets/sheet' . (int) $sheet['number'] . '.xml';
                if ($zip->addFile((string) $sheet['path'], $entry) !== true) {
                    throw MnbExcelException::withCode('Unable to add streamed sheet XML to XLSX: ' . $entry, ErrorCode::XLSX_ZIP_ENTRY_FAILED, ['entry' => $entry]);
                }
            }
        } catch (Throwable $e) {
            $zip->close();
            throw $e;
        }

        if ($zip->close() !== true) {
            throw MnbExcelException::withCode('Unable to finalize streaming XLSX package: ' . $path, ErrorCode::XLSX_ZIP_CLOSE_FAILED);
        }
    }

    private function zipString(ZipArchive $zip, string $entry, string $xml): void
    {
        if ($zip->addFromString($entry, $xml) !== true) {
            throw MnbExcelException::withCode('Unable to add XLSX zip entry: ' . $entry, ErrorCode::XLSX_ZIP_ENTRY_FAILED, ['entry' => $entry]);
        }
    }

    private function contentTypesXml(int $sheetCount): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            . '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            . '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            . '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        for ($i = 1; $i <= $sheetCount; $i++) {
            $xml .= '<Override PartName="/xl/worksheets/sheet' . $i . '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        return $xml . '</Types>';
    }

    private function rootRelsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            . '</Relationships>';
    }

    /** @param list<array<string,mixed>> $sheets */
    private function workbookXml(array $sheets): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>';
        foreach ($sheets as $index => $sheet) {
            $sheetId = $index + 1;
            $xml .= '<sheet name="' . $this->esc((string) $sheet['name']) . '" sheetId="' . $sheetId . '" r:id="rId' . $sheetId . '"/>';
        }
        return $xml . '</sheets></workbook>';
    }

    private function workbookRelsXml(int $sheetCount): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        for ($i = 1; $i <= $sheetCount; $i++) {
            $xml .= '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }
        $xml .= '<Relationship Id="rId' . ($sheetCount + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        return $xml . '</Relationships>';
    }

    private function stylesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font></fonts>'
            . '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
            . '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . '<cellXfs count="9">'
            . '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            . '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="44" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            . '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
            . '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
    }

    private function corePropsXml(array $options): string
    {
        $now = gmdate('Y-m-d\TH:i:s\Z');
        $creator = (string) ($options['creator'] ?? 'MNB PHPExcel');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            . '<dc:creator>' . $this->esc($creator) . '</dc:creator><cp:lastModifiedBy>' . $this->esc($creator) . '</cp:lastModifiedBy>'
            . '<dcterms:created xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:modified></cp:coreProperties>';
    }

    private function appPropsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>MNB PHPExcel</Application></Properties>';
    }

    /** @param resource $handle */
    private function writeXml($handle, string $xml): void
    {
        if (fwrite($handle, $xml) === false) {
            throw MnbExcelException::withCode('Unable to write streaming XLSX XML.', ErrorCode::FILE_WRITE_FAILED);
        }
    }

    /** @param array<string,mixed>|array<int,mixed>|mixed $row @return array<int|string,mixed> */
    private function normalizeRow(mixed $row): array
    {
        if ($row instanceof \stdClass) {
            $row = (array) $row;
        }
        if (!is_array($row)) {
            return [$row];
        }
        return $row;
    }

    /** @return list<string> */
    private function normalizeHeader(mixed $header): array
    {
        if ($header === null || $header === false || $header === '') {
            return [];
        }
        if (is_string($header)) {
            $header = array_map('trim', explode(',', $header));
        }
        if (!is_array($header)) {
            return [];
        }
        return array_values(array_map(static fn(mixed $value): string => (string) $value, $header));
    }

    /** @param array<int|string,mixed> $row @param list<string> $headers @return list<mixed> */
    private function orderRow(array $row, array $headers): array
    {
        if ($headers === []) {
            return array_values($row);
        }
        $ordered = [];
        foreach ($headers as $header) {
            $ordered[] = $row[$header] ?? null;
        }
        return $ordered;
    }

    /** @param array<int|string,mixed> $row */
    private function isAssociative(array $row): bool
    {
        return array_keys($row) !== range(0, count($row) - 1);
    }

    /** @param array<int|string,mixed> $columnFormats */
    private function formatForColumn(int $columnIndex, ?string $header, array $columnFormats): ?string
    {
        $name = Coordinate::columnIndexToName($columnIndex);
        $format = $columnFormats[$columnIndex] ?? $columnFormats[(string) $columnIndex] ?? $columnFormats[$name] ?? $columnFormats[strtolower($name)] ?? null;
        if ($header !== null) {
            $format = $columnFormats[$header] ?? $columnFormats[strtolower($header)] ?? $format;
        }
        return $format === null ? null : strtolower((string) $format);
    }

    private function styleIdForFormat(?string $format): int
    {
        return match (strtolower((string) $format)) {
            'integer', 'int' => 1,
            'decimal', 'number', 'float' => 2,
            'currency', 'money' => 3,
            'date', 'yyyy-mm-dd' => 4,
            'datetime', 'date_time' => 5,
            'percent', 'percentage' => 6,
            'text', 'string' => 7,
            default => 0,
        };
    }

    private function displayValue(mixed $value): mixed
    {
        if ($value instanceof CellValue) {
            return $value->displayValue();
        }
        if ($value instanceof DateTimeInterface) {
            return $value->format('Y-m-d H:i:s');
        }
        return $value;
    }

    private function serializeNumber(int|float|string $number): string
    {
        if (is_string($number)) {
            $number = trim($number);
            if ($number === '' || !is_numeric($number)) {
                return '0';
            }
            return preg_replace('/[^0-9Ee+\-.]/', '', $number) ?: '0';
        }
        if (is_int($number)) {
            return (string) $number;
        }
        if (!is_finite($number)) {
            return '0';
        }
        $formatted = sprintf('%.15G', $number);
        return str_replace(',', '.', $formatted);
    }

    private function excelDateSerial(mixed $value): string
    {
        if ($value instanceof DateTimeInterface) {
            $timestamp = (float) $value->getTimestamp();
        } else {
            $date = new DateTimeImmutable((string) $value);
            $timestamp = (float) $date->getTimestamp();
        }
        return $this->serializeNumber(25569 + ($timestamp / 86400));
    }

    private function sheetName(string $base, int $number): string
    {
        $base = trim(str_replace(['\\', '/', '?', '*', '[', ']', ':'], ' ', $base));
        $base = $base === '' ? 'Sheet' : $base;
        $suffix = $number > 1 ? ' ' . $number : '';
        return substr($base, 0, max(1, 31 - strlen($suffix))) . $suffix;
    }

    private function makeWorkDir(string $path): string
    {
        $dir = dirname($path) . DIRECTORY_SEPARATOR . '.mnb-large-export-' . bin2hex(random_bytes(8));
        if (!@mkdir($dir, 0775, true) && !is_dir($dir)) {
            throw MnbExcelException::withCode('Unable to create large export temp directory: ' . $dir, ErrorCode::DIRECTORY_CREATE_FAILED);
        }
        return $dir;
    }

    private function removeDirectory(string $dir): void
    {
        if (!is_dir($dir)) {
            return;
        }
        $items = scandir($dir) ?: [];
        foreach ($items as $item) {
            if ($item === '.' || $item === '..') {
                continue;
            }
            $path = $dir . DIRECTORY_SEPARATOR . $item;
            if (is_dir($path)) {
                $this->removeDirectory($path);
            } else {
                @unlink($path);
            }
        }
        @rmdir($dir);
    }

    /** @param array<string,mixed> $options @param array<string,mixed> $state */
    private function emitProgress(array $options, array $state): void
    {
        EventDispatcher::safeDispatch('large_export_progress', $state);
        if (is_callable($options['progress'] ?? null)) {
            ($options['progress'])($state);
        }
    }

    private function esc(string $value): string
    {
        return htmlspecialchars($value, ENT_XML1 | ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }
}
