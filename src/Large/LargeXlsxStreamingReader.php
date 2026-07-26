<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Large;

use Mnb\PHPExcel\Reader\XlsxStyleMap;
use Mnb\PHPExcel\Reader\SharedStrings\SharedStringProviderInterface;
use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Xml\XmlReader;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/**
 * Forward-only XLSX worksheet reader for large imports.
 *
 * It never returns the whole workbook. Rows are yielded to callbacks in chunks.
 */
final class LargeXlsxStreamingReader
{
    private XlsxWorkbookResolver $resolver;

    public function __construct(?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /**
     * @param callable(array<int,array<int|string,mixed>>, array<string,mixed>): (bool|void) $callback
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function chunk(string $path, int $chunkSize, callable $callback, array $options = []): array
    {
        $this->ensureExtensions();
        if ($chunkSize < 1) {
            throw MnbExcelException::withCode('Chunk size must be greater than zero.', ErrorCode::INVALID_ARGUMENT);
        }
        $realPath = realpath($path);
        if ($realPath === false || !is_file($path)) {
            throw MnbExcelException::withCode('XLSX file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $startedAt = microtime(true);
        $sheet = $options['sheet'] ?? 1;
        $sheetPath = $this->resolver->resolveSheetPath($realPath, is_int($sheet) || ctype_digit((string) $sheet) ? (int) $sheet : (string) $sheet);
        $withHeader = (bool) ($options['with_header'] ?? false);
        $headers = is_array($options['headers'] ?? null) ? array_values($options['headers']) : [];
        $headerConsumed = $headers !== [] || (bool) ($options['header_consumed'] ?? false);
        $skipRows = (int) ($options['skip_rows'] ?? 0);
        $startAfterRowNumber = isset($options['start_after_row_number']) ? (int) $options['start_after_row_number'] : null;
        $limitRows = isset($options['limit_rows']) ? (int) $options['limit_rows'] : null;
        $includeEmptyRows = (bool) ($options['include_empty_rows'] ?? false);
        $onlyColumns = $this->normalizeOnlyColumns((array) ($options['only_columns'] ?? []));
        $formulaMode = (string) ($options['formula_cells'] ?? 'formula');
        $preserveNumericStrings = (bool) ($options['preserve_numeric_strings'] ?? false);
        $convertDates = (bool) ($options['convert_dates'] ?? false);
        $dateOutputFormat = isset($options['date_output_format']) ? (string) $options['date_output_format'] : null;
        $styleMap = $convertDates ? $this->loadStyleMap($realPath) : XlsxStyleMap::fromXml(null);
        $date1904 = $convertDates ? $this->detectDate1904($realPath) : false;
        $timeBudget = isset($options['time_budget_seconds']) ? (int) $options['time_budget_seconds'] : 0;
        $memoryGuardRatio = (float) ($options['memory_guard_ratio'] ?? 0.85);
        $progress = $options['progress'] ?? null;
        $stopBeforeTimeout = (bool) ($options['stop_before_timeout'] ?? true);

        $sharedStrings = LargeSharedStringCache::fromXlsx($realPath, $options);
        $reader = new XMLReader();
        $uri = $this->zipUri($realPath, $sheetPath);
        if (!@$reader->open($uri, null, LIBXML_NONET | LIBXML_COMPACT)) {
            $sharedStrings->close();
            throw MnbExcelException::withCode('Unable to open worksheet XML for streaming import: ' . $sheetPath, ErrorCode::XLSX_INVALID, ['sheet' => $sheetPath]);
        }

        $chunk = [];
        $chunkRowNumbers = [];
        $rowsRead = 0;
        $rowsDelivered = 0;
        $chunksDelivered = 0;
        $physicalRows = 0;
        $stopped = false;
        $stopReason = null;
        $lastDeliveredRowNumber = null;

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'row') {
                    continue;
                }

                $physicalRows++;
                if ($physicalRows <= $skipRows) {
                    $this->skipCurrentElement($reader);
                    continue;
                }

                if ($stopBeforeTimeout && $this->timeBudgetReached($startedAt, $timeBudget)) {
                    $stopped = true;
                    $stopReason = 'time_budget_reached';
                    break;
                }
                $this->assertMemorySafe($memoryGuardRatio);

                $rowNumber = (int) ($reader->getAttribute('r') ?: $physicalRows);
                if ($startAfterRowNumber !== null && $rowNumber <= $startAfterRowNumber) {
                    $this->skipCurrentElement($reader);
                    continue;
                }
                $row = $this->readRow($reader, $sharedStrings, $formulaMode, $onlyColumns, $styleMap, $convertDates, $date1904, $preserveNumericStrings, $dateOutputFormat);
                if (!$includeEmptyRows && $this->isEmptyRow($row)) {
                    continue;
                }

                if ($withHeader && !$headerConsumed) {
                    $headers = $this->normalizeHeaders($row);
                    $headerConsumed = true;
                    continue;
                }

                if ($limitRows !== null && $rowsRead >= $limitRows) {
                    $stopped = true;
                    $stopReason = 'row_limit_reached';
                    break;
                }
                $rowsRead++;

                $chunk[] = $withHeader ? $this->associateRow($headers, $row) : $row;
                $chunkRowNumbers[] = $rowNumber;
                $rowsDelivered++;

                if (count($chunk) >= $chunkSize) {
                    $chunksDelivered++;
                    $state = $this->state($path, $sheetPath, $rowsRead, $rowsDelivered, $chunksDelivered, $rowNumber, $startedAt, $sharedStrings->mode(), false, null, $chunkRowNumbers);
                    $result = $callback($chunk, $state);
                    $lastDeliveredRowNumber = $state['chunk_last_row_number'] ?? $rowNumber;
                    $chunk = [];
                    $chunkRowNumbers = [];
                    if (is_callable($progress)) {
                        $progress($state);
                    }
                    if ($result === false) {
                        $stopped = true;
                        $stopReason = 'callback_stopped';
                        break;
                    }
                }
            }

            if ($chunk !== []) {
                $chunksDelivered++;
                $state = $this->state($path, $sheetPath, $rowsRead, $rowsDelivered, $chunksDelivered, $chunkRowNumbers !== [] ? max($chunkRowNumbers) : null, $startedAt, $sharedStrings->mode(), $stopped, $stopReason, $chunkRowNumbers);
                $result = $callback($chunk, $state);
                $lastDeliveredRowNumber = $state['chunk_last_row_number'] ?? $lastDeliveredRowNumber;
                if (is_callable($progress)) {
                    $progress($state);
                }
                if ($result === false) {
                    $stopped = true;
                    $stopReason = 'callback_stopped';
                }
            }
        } finally {
            $reader->close();
            $sharedStrings->close();
        }

        return $this->state($path, $sheetPath, $rowsRead, $rowsDelivered, $chunksDelivered, is_int($lastDeliveredRowNumber) ? $lastDeliveredRowNumber : null, $startedAt, 'closed', $stopped, $stopReason, []);
    }

    /**
     * @param callable(array<int|string,mixed>, array<string,mixed>): (bool|void) $callback
     * @param array<string,mixed> $options
     * @return array<string,mixed>
     */
    public function eachRow(string $path, callable $callback, array $options = []): array
    {
        return $this->chunk($path, 1, static function (array $rows, array $state) use ($callback): bool|null {
            foreach ($rows as $row) {
                $result = $callback($row, $state);
                if ($result === false) {
                    return false;
                }
            }
            return null;
        }, $options);
    }

    /** @param array<string,bool> $onlyColumns @return array<int,mixed> */
    private function readRow(XMLReader $reader, SharedStringProviderInterface $sharedStrings, string $formulaMode, array $onlyColumns, XlsxStyleMap $styleMap, bool $convertDates, bool $date1904, bool $preserveNumericStrings, ?string $dateOutputFormat): array
    {
        $rowDepth = $reader->depth;
        $cells = [];
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->localName === 'row' && $reader->depth === $rowDepth) {
                break;
            }
            if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'c') {
                continue;
            }
            $ref = (string) ($reader->getAttribute('r') ?? '');
            $columnIndex = $ref !== '' ? $this->columnIndexFromCellRef($ref) : count($cells) + 1;
            if ($onlyColumns !== [] && !isset($onlyColumns[$this->columnLetters($columnIndex)]) && !isset($onlyColumns[(string) $columnIndex])) {
                $this->skipCurrentElement($reader);
                continue;
            }
            $type = (string) ($reader->getAttribute('t') ?? '');
            $styleIndex = $reader->getAttribute('s');
            $xml = $reader->readOuterXml();
            $cells[$columnIndex - 1] = $this->cellValue($xml, $type, $sharedStrings, $formulaMode, $styleMap, $styleIndex !== null ? (int) $styleIndex : null, $convertDates, $date1904, $preserveNumericStrings, $dateOutputFormat);
        }
        if ($cells === []) {
            return [];
        }
        ksort($cells);
        $max = max(array_keys($cells));
        $row = [];
        for ($i = 0; $i <= $max; $i++) {
            $row[] = $cells[$i] ?? null;
        }
        return $row;
    }

    private function skipCurrentElement(XMLReader $reader): void
    {
        if ($reader->isEmptyElement) {
            return;
        }
        $depth = $reader->depth;
        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::END_ELEMENT && $reader->depth === $depth) {
                return;
            }
        }
    }

    private function cellValue(string $xml, string $type, SharedStringProviderInterface $sharedStrings, string $formulaMode, XlsxStyleMap $styleMap, ?int $styleIndex, bool $convertDates, bool $date1904, bool $preserveNumericStrings, ?string $dateOutputFormat): mixed
    {
        if ($formulaMode !== 'cached_value' && preg_match('/<f\b[^>]*>(.*?)<\/f>/su', $xml, $m)) {
            return '=' . html_entity_decode($m[1], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        if ($type === 'inlineStr') {
            return $this->textFromRichXml($xml);
        }
        $value = $this->readTag($xml, 'v');
        if ($value === null) {
            return null;
        }
        if ($type === 's') {
            return $sharedStrings->get((int) $value);
        }
        if ($type === 'b') {
            return $value === '1';
        }
        if ($type === 'str' || $type === 'e') {
            return $value;
        }
        if (is_numeric($value)) {
            if ($convertDates && $styleMap->isDateStyle($styleIndex)) {
                return $this->excelSerialToDateString((float) $value, $date1904, $dateOutputFormat);
            }
            if ($preserveNumericStrings) {
                return $value;
            }
            return str_contains($value, '.') ? (float) $value : (int) $value;
        }
        return $value;
    }

    private function readTag(string $xml, string $tag): ?string
    {
        if (preg_match('/<' . preg_quote($tag, '/') . '\b[^>]*>(.*?)<\/' . preg_quote($tag, '/') . '>/su', $xml, $m)) {
            return html_entity_decode($m[1], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return null;
    }

    private function textFromRichXml(string $xml): string
    {
        preg_match_all('/<t\b[^>]*>(.*?)<\/t>/su', $xml, $matches);
        if (($matches[1] ?? []) !== []) {
            return html_entity_decode(implode('', $matches[1]), ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return html_entity_decode(strip_tags($xml), ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    /** @return array<string,string> */
    private function normalizeHeaders(array $row): array
    {
        $headers = [];
        foreach ($row as $i => $value) {
            $name = trim((string) $value);
            $name = $name !== '' ? $name : 'column_' . ((int) $i + 1);
            $base = $name;
            $suffix = 2;
            while (in_array($name, $headers, true)) {
                $name = $base . '_' . $suffix;
                $suffix++;
            }
            $headers[$i] = $name;
        }
        return $headers;
    }

    /** @param array<int,string> $headers @param array<int,mixed> $row @return array<string,mixed> */
    private function associateRow(array $headers, array $row): array
    {
        $out = [];
        $max = max(count($headers), count($row));
        for ($i = 0; $i < $max; $i++) {
            $key = $headers[$i] ?? ('column_' . ($i + 1));
            $out[$key] = $row[$i] ?? null;
        }
        return $out;
    }

    private function isEmptyRow(array $row): bool
    {
        foreach ($row as $value) {
            if ($value !== null && $value !== '') {
                return false;
            }
        }
        return true;
    }

    /** @param list<int|string> $columns @return array<string,bool> */
    private function normalizeOnlyColumns(array $columns): array
    {
        $normalized = [];
        foreach ($columns as $column) {
            if (is_int($column) || ctype_digit((string) $column)) {
                $normalized[(string) (int) $column] = true;
                $normalized[$this->columnLetters((int) $column)] = true;
            } else {
                $normalized[strtoupper((string) $column)] = true;
            }
        }
        return $normalized;
    }

    private function columnIndexFromCellRef(string $ref): int
    {
        if (!preg_match('/^\$?([A-Z]{1,3})/i', $ref, $m)) {
            return 1;
        }
        $letters = strtoupper($m[1]);
        $index = 0;
        for ($i = 0, $len = strlen($letters); $i < $len; $i++) {
            $index = $index * 26 + (ord($letters[$i]) - 64);
        }
        return $index;
    }

    private function columnLetters(int $index): string
    {
        $letters = '';
        while ($index > 0) {
            $mod = ($index - 1) % 26;
            $letters = chr(65 + $mod) . $letters;
            $index = intdiv($index - 1, 26);
        }
        return $letters;
    }

    private function timeBudgetReached(float $startedAt, int $timeBudget): bool
    {
        return $timeBudget > 0 && (microtime(true) - $startedAt) >= $timeBudget;
    }

    private function assertMemorySafe(float $ratio): void
    {
        $limit = $this->memoryLimitBytes();
        if ($limit <= 0) {
            return;
        }
        if (memory_get_usage(true) >= (int) floor($limit * $ratio)) {
            throw MnbExcelException::withCode(
                'Large XLSX import stopped before reaching the PHP memory limit.',
                ErrorCode::FILE_READ_FAILED,
                ['memory_usage' => memory_get_usage(true), 'memory_limit' => $limit, 'guard_ratio' => $ratio],
                null,
                'The import was stopped safely before the server memory limit was reached.'
            );
        }
    }

    private function memoryLimitBytes(): int
    {
        $limit = trim((string) ini_get('memory_limit'));
        if ($limit === '' || $limit === '-1') {
            return 0;
        }
        if (!preg_match('/^([0-9.]+)\s*([KMG])?B?$/i', $limit, $m)) {
            return 0;
        }
        $amount = (float) $m[1];
        return (int) round(match (strtoupper($m[2] ?? 'M')) {
            'K' => $amount * 1024,
            'G' => $amount * 1073741824,
            default => $amount * 1048576,
        });
    }

    /** @param list<int> $chunkRowNumbers @return array<string,mixed> */
    private function state(string $path, string $sheetPath, int $rowsRead, int $rowsDelivered, int $chunksDelivered, ?int $currentRowNumber, float $startedAt, string $sharedStringMode, bool $stopped, ?string $stopReason, array $chunkRowNumbers = []): array
    {
        return [
            'path' => $path,
            'sheet_path' => $sheetPath,
            'rows_read' => $rowsRead,
            'rows_delivered' => $rowsDelivered,
            'chunks_delivered' => $chunksDelivered,
            'current_row_number' => $currentRowNumber,
            'chunk_first_row_number' => $chunkRowNumbers !== [] ? min($chunkRowNumbers) : null,
            'chunk_last_row_number' => $chunkRowNumbers !== [] ? max($chunkRowNumbers) : $currentRowNumber,
            'chunk_row_numbers' => $chunkRowNumbers,
            'elapsed_seconds' => round(microtime(true) - $startedAt, 3),
            'memory_usage_mb' => round(memory_get_usage(true) / 1048576, 2),
            'shared_string_cache' => $sharedStringMode,
            'stopped' => $stopped,
            'stop_reason' => $stopReason,
        ];
    }


    private function loadStyleMap(string $realPath): XlsxStyleMap
    {
        $zip = new ZipArchive();
        $openResult = $zip->open($realPath);
        if ($openResult !== true) {
            throw MnbExcelException::withCode('Unable to open XLSX styles for large import.', ErrorCode::XLSX_INVALID, ['zip_status' => $openResult]);
        }
        try {
            $stylesXml = $zip->getFromName('xl/styles.xml');
            return XlsxStyleMap::fromXml(is_string($stylesXml) ? $stylesXml : null);
        } finally {
            $zip->close();
        }
    }

    private function detectDate1904(string $realPath): bool
    {
        $zip = new ZipArchive();
        $openResult = $zip->open($realPath);
        if ($openResult !== true) {
            return false;
        }
        try {
            $workbookXml = $zip->getFromName('xl/workbook.xml');
            return is_string($workbookXml) && preg_match('/<workbookPr\b[^>]*date1904=["\']1["\']/i', $workbookXml) === 1;
        } finally {
            $zip->close();
        }
    }

    private function excelSerialToDateString(float $serial, bool $date1904, ?string $format): string
    {
        $timezone = new \DateTimeZone('UTC');
        $base = new \DateTimeImmutable($date1904 ? '1904-01-01 00:00:00' : '1899-12-30 00:00:00', $timezone);
        $seconds = (int) round($serial * 86400);
        $modifier = ($seconds >= 0 ? '+' : '') . $seconds . ' seconds';
        $date = $base->modify($modifier);
        if (!$date instanceof \DateTimeImmutable) {
            return (string) $serial;
        }
        if ($format !== null && $format !== '') {
            return $date->format($format);
        }
        $hasTime = abs($serial - floor($serial)) > 0.0000001;
        return $date->format($hasTime ? 'Y-m-d H:i:s' : 'Y-m-d');
    }

    private function zipUri(string $realPath, string $entry): string
    {
        return 'zip://' . str_replace('\\', '/', $realPath) . '#' . $entry;
    }

    private function ensureExtensions(): void
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode('ext-zip is required for large XLSX streaming import.', ErrorCode::EXTENSION_MISSING);
        }
        if (!class_exists(XMLReader::class)) {
            throw MnbExcelException::withCode('ext-xmlreader is required for large XLSX streaming import.', ErrorCode::EXTENSION_MISSING);
        }
    }
}
