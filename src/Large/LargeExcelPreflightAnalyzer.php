<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use XMLReader;
use ZipArchive;

/**
 * Safely inspects XLSX size/shape without loading worksheet rows into PHP arrays.
 */
final class LargeExcelPreflightAnalyzer
{
    private const EXCEL_MAX_ROWS = 1048576;
    private const EXCEL_MAX_COLUMNS = 16384;

    public function __construct(private ?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /**
     * @param array<string,mixed> $options accurate_row_count, scan_features, time_budget_seconds, sheet
     * @return array<string,mixed>
     */
    public function analyze(string $path, array $options = []): array
    {
        $this->ensureExtensions();
        $realPath = realpath($path);
        if ($realPath === false || !is_file($path)) {
            throw MnbExcelException::withCode('XLSX file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $startedAt = microtime(true);
        $timeBudget = isset($options['time_budget_seconds']) ? (int) $options['time_budget_seconds'] : 0;
        $accurate = (bool) ($options['accurate_row_count'] ?? true);
        $scanFeatures = (bool) ($options['scan_features'] ?? true);
        $selectedSheet = $options['sheet'] ?? null;

        $zip = $this->openZip($realPath);
        $sheets = $this->resolver->sheets($realPath);
        if ($selectedSheet !== null) {
            $sheets = array_values(array_filter($sheets, fn (array $sheet): bool => $sheet['name'] === $selectedSheet || (string) $sheet['index'] === (string) $selectedSheet));
        }

        $sharedStrings = $this->sharedStringSummary($realPath, $zip);
        $styles = $this->stylesSummary($zip);
        $workbookFeatures = $this->workbookFeatureInventory($zip);

        $sheetProfiles = [];
        $totals = [
            'rows' => 0,
            'cells' => 0,
            'uncompressed_sheet_xml_size' => 0,
            'compressed_sheet_xml_size' => 0,
            'max_columns' => 0,
        ];
        $features = [
            'formulas' => 0,
            'comments' => $workbookFeatures['comments'],
            'hyperlinks' => 0,
            'merged_cells' => 0,
            'drawings' => $workbookFeatures['drawings'],
            'charts' => $workbookFeatures['charts'],
            'pivot_tables' => $workbookFeatures['pivot_tables'],
            'external_links' => $workbookFeatures['external_links'],
            'macros' => $workbookFeatures['macros'],
            'tables' => $workbookFeatures['tables'],
        ];
        $warnings = [];
        $partial = false;

        foreach ($sheets as $sheet) {
            if (!$sheet['exists']) {
                $warnings[] = 'Worksheet XML missing for sheet: ' . $sheet['name'];
                continue;
            }

            if ($timeBudget > 0 && microtime(true) - $startedAt > $timeBudget) {
                $partial = true;
                $warnings[] = 'Preflight stopped early because time budget was reached.';
                break;
            }

            $sheetProfile = $this->analyzeSheet($realPath, $zip, $sheet, $accurate, $scanFeatures, $startedAt, $timeBudget);
            $sheetProfiles[] = $sheetProfile;
            $totals['rows'] += (int) $sheetProfile['rows'];
            $totals['cells'] += (int) $sheetProfile['cells'];
            $totals['uncompressed_sheet_xml_size'] += (int) $sheetProfile['uncompressed_xml_size'];
            $totals['compressed_sheet_xml_size'] += (int) $sheetProfile['compressed_xml_size'];
            $totals['max_columns'] = max($totals['max_columns'], (int) $sheetProfile['columns']);
            foreach (['formulas', 'hyperlinks', 'merged_cells'] as $key) {
                $features[$key] += (int) ($sheetProfile['features'][$key] ?? 0);
            }
            if (!empty($sheetProfile['partial'])) {
                $partial = true;
            }
        }

        $zip->close();

        $profile = [
            'status' => $partial ? 'partial' : 'ok',
            'path' => $realPath,
            'file_size_bytes' => filesize($realPath) ?: 0,
            'file_size_mb' => round((filesize($realPath) ?: 0) / 1048576, 2),
            'sheet_count' => count($sheetProfiles),
            'sheets' => $sheetProfiles,
            'total_rows' => $totals['rows'],
            'total_estimated_rows' => $totals['rows'],
            'total_cells' => $totals['cells'],
            'total_estimated_cells' => $totals['cells'],
            'max_columns' => $totals['max_columns'],
            'total_uncompressed_sheet_xml_size' => $totals['uncompressed_sheet_xml_size'],
            'total_uncompressed_sheet_xml_mb' => round($totals['uncompressed_sheet_xml_size'] / 1048576, 2),
            'total_compressed_sheet_xml_size' => $totals['compressed_sheet_xml_size'],
            'shared_strings' => $sharedStrings,
            'styles' => $styles,
            'features' => $features,
            'complexity' => $this->complexity($totals['rows'], $totals['cells'], (filesize($realPath) ?: 0) / 1048576, $totals['uncompressed_sheet_xml_size'] / 1048576, $features),
            'risk' => $this->risk($totals['rows'], $totals['cells'], $totals['uncompressed_sheet_xml_size'], $features),
            'excel_limits' => [
                'max_rows_per_sheet' => self::EXCEL_MAX_ROWS,
                'max_columns_per_sheet' => self::EXCEL_MAX_COLUMNS,
                'within_limits' => $this->withinExcelLimits($sheetProfiles),
            ],
            'warnings' => $warnings,
            'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
            'partial' => $partial,
        ];

        $profile['method_advice'] = (new ImportMethodAdvisor())->recommendFromProfile($profile, (array) ($options['server'] ?? []));

        return $profile;
    }

    /** @param array<string,mixed> $sheet @return array<string,mixed> */
    private function analyzeSheet(string $realPath, ZipArchive $zip, array $sheet, bool $accurate, bool $scanFeatures, float $startedAt, int $timeBudget): array
    {
        $stat = $zip->statName($sheet['path']) ?: [];
        $dimension = $this->dimensionFromStream($zip, $sheet['path']);
        $rows = $dimension['rows'];
        $columns = $dimension['columns'];
        $cells = $rows * max(1, $columns);
        $features = ['formulas' => 0, 'hyperlinks' => 0, 'merged_cells' => 0];
        $partial = false;
        $physicalRows = 0;
        $filledRows = 0;
        $physicalCells = 0;

        if ($accurate || $scanFeatures || $dimension['rows'] === 0) {
            $uri = $this->zipUri($realPath, (string) $sheet['path']);
            $reader = new XMLReader();
            if (!@$reader->open($uri, null, LIBXML_NONET | LIBXML_COMPACT)) {
                throw MnbExcelException::withCode('Unable to open worksheet XML for preflight: ' . $sheet['path'], ErrorCode::XLSX_INVALID, ['sheet' => $sheet['path']]);
            }

            $inRow = false;
            $rowHasCell = false;
            $currentRowNumber = 0;
            while ($reader->read()) {
                if ($timeBudget > 0 && microtime(true) - $startedAt > $timeBudget) {
                    $partial = true;
                    break;
                }

                if ($reader->nodeType === XMLReader::ELEMENT) {
                    if ($reader->name === 'dimension' && $dimension['ref'] === '') {
                        $dimension = $this->dimensionFromRef((string) ($reader->getAttribute('ref') ?? ''));
                    } elseif ($reader->name === 'row') {
                        $inRow = true;
                        $rowHasCell = false;
                        $physicalRows++;
                        $currentRowNumber = (int) ($reader->getAttribute('r') ?: $physicalRows);
                        $rows = max($rows, $currentRowNumber);
                    } elseif ($reader->name === 'c') {
                        $rowHasCell = true;
                        $physicalCells++;
                        $ref = (string) ($reader->getAttribute('r') ?? '');
                        if ($ref !== '' && preg_match('/^\$?([A-Z]{1,3})\$?([0-9]+)$/i', $ref, $m)) {
                            $columns = max($columns, $this->columnIndex($m[1]));
                            $rows = max($rows, (int) $m[2]);
                        }
                    } elseif ($scanFeatures && $reader->name === 'f') {
                        $features['formulas']++;
                    } elseif ($scanFeatures && $reader->name === 'hyperlink') {
                        $features['hyperlinks']++;
                    } elseif ($scanFeatures && $reader->name === 'mergeCell') {
                        $features['merged_cells']++;
                    }
                } elseif ($reader->nodeType === XMLReader::END_ELEMENT && $reader->name === 'row' && $inRow) {
                    if ($rowHasCell) {
                        $filledRows++;
                    }
                    $inRow = false;
                }
            }
            $reader->close();

            if ($physicalRows > 0) {
                $rows = max($rows, $physicalRows);
            }
            if ($physicalCells > 0) {
                $cells = $physicalCells;
            } else {
                $cells = $rows * max(1, $columns);
            }
        }

        return [
            'index' => $sheet['index'],
            'name' => $sheet['name'],
            'state' => $sheet['state'],
            'path' => $sheet['path'],
            'rows' => $rows,
            'physical_rows' => $physicalRows > 0 ? $physicalRows : $rows,
            'filled_rows' => $filledRows > 0 ? $filledRows : $rows,
            'columns' => $columns,
            'cells' => $cells,
            'dimension_ref' => $dimension['ref'],
            'uncompressed_xml_size' => (int) ($stat['size'] ?? 0),
            'uncompressed_xml_mb' => round(((int) ($stat['size'] ?? 0)) / 1048576, 2),
            'compressed_xml_size' => (int) ($stat['comp_size'] ?? 0),
            'features' => $features,
            'partial' => $partial,
        ];
    }

    /** @return array{ref:string,rows:int,columns:int} */
    private function dimensionFromStream(ZipArchive $zip, string $path): array
    {
        $stream = $zip->getStream($path);
        if (!is_resource($stream)) {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        $chunk = stream_get_contents($stream, 65536);
        fclose($stream);
        if (!is_string($chunk) || $chunk === '') {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        if (preg_match('/<dimension\b[^>]*\bref\s*=\s*("([^"]+)"|\'([^\']+)\')/i', $chunk, $m)) {
            return $this->dimensionFromRef($m[2] !== '' ? $m[2] : $m[3]);
        }
        return ['ref' => '', 'rows' => 0, 'columns' => 0];
    }

    /** @return array{ref:string,rows:int,columns:int} */
    private function dimensionFromRef(string $ref): array
    {
        $ref = trim($ref);
        if ($ref === '') {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        $parts = explode(':', $ref);
        $last = end($parts) ?: $ref;
        if (!preg_match('/^\$?([A-Z]{1,3})\$?([0-9]+)$/i', $last, $m)) {
            return ['ref' => $ref, 'rows' => 0, 'columns' => 0];
        }
        return ['ref' => $ref, 'rows' => (int) $m[2], 'columns' => $this->columnIndex($m[1])];
    }

    /** @return array<string,mixed> */
    private function sharedStringSummary(string $realPath, ZipArchive $zip): array
    {
        $stat = $zip->statName('xl/sharedStrings.xml') ?: null;
        if ($stat === null) {
            return ['exists' => false, 'count' => 0, 'unique_count' => 0, 'size_bytes' => 0, 'size_mb' => 0.0];
        }

        $count = 0;
        $uniqueCount = 0;
        $stream = $zip->getStream('xl/sharedStrings.xml');
        if (is_resource($stream)) {
            $chunk = stream_get_contents($stream, 8192);
            fclose($stream);
            if (is_string($chunk) && preg_match('/<sst\b[^>]*>/i', $chunk, $m)) {
                if (preg_match('/\bcount\s*=\s*("([0-9]+)"|\'([0-9]+)\')/i', $m[0], $countMatch)) {
                    $count = (int) ($countMatch[2] !== '' ? $countMatch[2] : $countMatch[3]);
                }
                if (preg_match('/\buniqueCount\s*=\s*("([0-9]+)"|\'([0-9]+)\')/i', $m[0], $uniqueMatch)) {
                    $uniqueCount = (int) ($uniqueMatch[2] !== '' ? $uniqueMatch[2] : $uniqueMatch[3]);
                }
            }
        }

        return [
            'exists' => true,
            'count' => $count,
            'unique_count' => $uniqueCount,
            'size_bytes' => (int) ($stat['size'] ?? 0),
            'size_mb' => round(((int) ($stat['size'] ?? 0)) / 1048576, 2),
            'compressed_size_bytes' => (int) ($stat['comp_size'] ?? 0),
        ];
    }

    /** @return array<string,mixed> */
    private function stylesSummary(ZipArchive $zip): array
    {
        $xml = $zip->getFromName('xl/styles.xml');
        if ($xml === false) {
            return ['exists' => false, 'cell_xfs' => 0, 'fonts' => 0, 'fills' => 0];
        }
        return [
            'exists' => true,
            'cell_xfs' => $this->countAttr($xml, 'cellXfs'),
            'fonts' => $this->countAttr($xml, 'fonts'),
            'fills' => $this->countAttr($xml, 'fills'),
        ];
    }

    /** @return array<string,int|bool> */
    private function workbookFeatureInventory(ZipArchive $zip): array
    {
        $features = [
            'comments' => 0,
            'drawings' => 0,
            'charts' => 0,
            'pivot_tables' => 0,
            'external_links' => 0,
            'macros' => false,
            'tables' => 0,
        ];
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = (string) ($zip->getNameIndex($i) ?: '');
            if (preg_match('#^xl/comments\d*\.xml$#', $name)) {
                $features['comments']++;
            } elseif (str_starts_with($name, 'xl/drawings/')) {
                $features['drawings']++;
            } elseif (str_starts_with($name, 'xl/charts/')) {
                $features['charts']++;
            } elseif (str_starts_with($name, 'xl/pivotTables/') || str_starts_with($name, 'xl/pivotCache/')) {
                $features['pivot_tables']++;
            } elseif (str_starts_with($name, 'xl/externalLinks/')) {
                $features['external_links']++;
            } elseif ($name === 'xl/vbaProject.bin') {
                $features['macros'] = true;
            } elseif (str_starts_with($name, 'xl/tables/')) {
                $features['tables']++;
            }
        }
        return $features;
    }

    private function countAttr(string $xml, string $tag): int
    {
        if (preg_match('/<' . preg_quote($tag, '/') . '\b[^>]*\bcount\s*=\s*("([0-9]+)"|\'([0-9]+)\')/i', $xml, $m)) {
            return (int) ($m[2] !== '' ? $m[2] : $m[3]);
        }
        return 0;
    }

    /** @param list<array<string,mixed>> $sheets */
    private function withinExcelLimits(array $sheets): bool
    {
        foreach ($sheets as $sheet) {
            if ((int) ($sheet['rows'] ?? 0) > self::EXCEL_MAX_ROWS || (int) ($sheet['columns'] ?? 0) > self::EXCEL_MAX_COLUMNS) {
                return false;
            }
        }
        return true;
    }

    /** @param array<string,int|bool> $features */
    private function complexity(int $rows, int $cells, float $fileSizeMb, float $uncompressedMb, array $features): string
    {
        $advice = (new ImportMethodAdvisor())->recommendFromProfile([
            'total_rows' => $rows,
            'total_cells' => $cells,
            'file_size_mb' => $fileSizeMb,
            'total_uncompressed_sheet_xml_mb' => $uncompressedMb,
            'features' => $features,
        ]);
        return (string) $advice['level'];
    }

    /** @param array<string,int|bool> $features */
    private function risk(int $rows, int $cells, int $uncompressedSize, array $features): string
    {
        $featureCount = 0;
        foreach ($features as $value) {
            if ($value === true || (is_int($value) && $value > 0)) {
                $featureCount++;
            }
        }
        if ($rows > 500000 || $cells > 5000000 || $uncompressedSize > 900 * 1048576) {
            return $featureCount >= 3 ? 'very_high' : 'high';
        }
        if ($rows > 150000 || $cells > 1500000 || $uncompressedSize > 350 * 1048576) {
            return $featureCount >= 3 ? 'high' : 'medium';
        }
        if ($rows > 50000 || $cells > 500000) {
            return 'medium';
        }
        return 'low';
    }

    private function columnIndex(string $letters): int
    {
        $letters = strtoupper($letters);
        $index = 0;
        for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
            $index = $index * 26 + (ord($letters[$i]) - 64);
        }
        return $index;
    }


    private function zipUri(string $realPath, string $entry): string
    {
        return 'zip://' . str_replace('\\', '/', $realPath) . '#' . $entry;
    }

    private function ensureExtensions(): void
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode('ext-zip is required for XLSX preflight analysis.', ErrorCode::EXTENSION_MISSING);
        }
        if (!class_exists(XMLReader::class)) {
            throw MnbExcelException::withCode('ext-xmlreader is required for accurate XLSX preflight analysis.', ErrorCode::EXTENSION_MISSING);
        }
    }

    private function openZip(string $path): ZipArchive
    {
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            throw MnbExcelException::withCode('Unable to open XLSX zip package for preflight.', ErrorCode::XLSX_INVALID, ['path' => $path]);
        }
        return $zip;
    }
}
