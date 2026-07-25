<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use DateTimeImmutable;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\LocaleNormalizer;
use Mnb\PHPExcel\Reader\SharedStrings\InMemorySharedStringProvider;
use Mnb\PHPExcel\Reader\SharedStrings\SharedStringProviderInterface;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use XMLReader;
use ZipArchive;

final class XlsxReader implements IterableReaderInterface
{
    private XlsxWorkbookResolver $resolver;

    public function __construct(?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        $rows = iterator_to_array($this->iterateSheet($path, $sheet, $options), true);
        ksort($rows);
        return $rows;
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        $this->ensureExtensions();

        $realPath = realpath($path);
        if ($realPath === false || !is_file($path)) {
            throw MnbExcelException::withCode('Invalid XLSX path: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }

        $size = filesize($realPath);
        $maxFileBytes = isset($options['max_file_bytes']) ? max(0, (int) $options['max_file_bytes']) : null;
        if ($maxFileBytes !== null && $size !== false && $size > $maxFileBytes) {
            throw MnbExcelException::withCode(
                'XLSX file exceeds max_file_bytes. Size: ' . $size . ', max_file_bytes: ' . $maxFileBytes,
                ErrorCode::FILE_READ_FAILED,
                ['path' => $path, 'size_bytes' => $size, 'max_file_bytes' => $maxFileBytes]
            );
        }

        if ((bool) ($options['validate_package'] ?? false)) {
            (new XlsxIntegrityValidator())->assertValid($realPath, (array) ($options['validation_options'] ?? []));
        }

        $formulaMode = (string) ($options['formula_cells'] ?? 'formula');
        if (!in_array($formulaMode, ['formula', 'cached_value', 'both'], true)) {
            throw new MnbExcelException('formula_cells must be "formula", "cached_value", or "both".');
        }

        $sheetXml = $this->resolver->resolveSheetPath($realPath, $sheet);
        $this->assertZipEntrySize($realPath, $sheetXml, $options, 'max_worksheet_bytes');
        $sharedStrings = $this->readSharedStrings($realPath, $options);
        $styleMap = $this->readStyleMap($realPath);
        $date1904 = $this->usesDate1904($realPath);
        $hiddenColumns = ((bool) ($options['include_hidden_columns'] ?? true)) ? [] : $this->hiddenColumnIndexes($realPath, $sheetXml);
        $projection = ColumnProjection::fromOptions($options);
        $mergedMode = strtolower((string) ($options['merged_cells'] ?? 'anchor'));
        if (!in_array($mergedMode, ['anchor', 'expand', 'metadata'], true)) {
            throw new MnbExcelException('merged_cells must be "anchor", "expand", or "metadata".');
        }
        $mergedCells = $mergedMode === 'expand' ? MergedCellMap::fromXlsx($realPath, $sheetXml, $options) : null;
        $sourceProjection = $mergedCells?->active() ? ColumnProjection::fromOptions([]) : $projection;
        $mergedAnchors = [];
        $uri = $this->zipUri($realPath, $sheetXml);

        $reader = new XMLReader();
        if (!@$reader->open($uri, null, LIBXML_NONET | LIBXML_COMPACT)) {
            $sharedStrings->close();
            throw MnbExcelException::withCode('Unable to open worksheet XML: ' . $sheetXml, ErrorCode::XLSX_INVALID, ['sheet' => $sheetXml]);
        }

        $includeHiddenRows = (bool) ($options['include_hidden_rows'] ?? true);
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceSkipRows = max(0, (int) ($options['source_skip_rows'] ?? 0));
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $maxRows = isset($options['max_rows']) ? max(0, (int) $options['max_rows']) : null;
        $maxCells = isset($options['max_cells']) ? max(0, (int) $options['max_cells']) : null;
        $maxColumns = isset($options['max_columns']) ? max(1, (int) $options['max_columns']) : null;
        $cellCount = 0;
        $delivered = 0;
        $physicalRows = 0;

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'row') {
                    continue;
                }

                $physicalRows++;
                $rowNumber = (int) ($reader->getAttribute('r') ?: $physicalRows);
                if ($rowNumber < $startRow || $physicalRows <= $sourceSkipRows) {
                    if ($mergedCells?->active()) {
                        $anchorRow = $this->readRow($reader, $sharedStrings, $styleMap, $date1904, $options, $hiddenColumns, $sourceProjection);
                        $mergedCells->expandRow($rowNumber, $anchorRow, $mergedAnchors);
                    } else {
                        $this->skipCurrentElement($reader);
                    }
                    continue;
                }
                if ($endRow !== null && $rowNumber > $endRow) {
                    break;
                }
                if ($sourceLimitRows !== null && $delivered >= $sourceLimitRows) {
                    break;
                }

                $hidden = $reader->getAttribute('hidden') === '1' || $reader->getAttribute('hidden') === 'true';
                if (!$includeHiddenRows && $hidden) {
                    if ($mergedCells?->active()) {
                        $anchorRow = $this->readRow($reader, $sharedStrings, $styleMap, $date1904, $options, $hiddenColumns, $sourceProjection);
                        $mergedCells->expandRow($rowNumber, $anchorRow, $mergedAnchors);
                    } else {
                        $this->skipCurrentElement($reader);
                    }
                    continue;
                }

                $row = $this->readRow($reader, $sharedStrings, $styleMap, $date1904, $options, $hiddenColumns, $sourceProjection);
                if ($mergedCells?->active()) {
                    $row = $mergedCells->expandRow($rowNumber, $row, $mergedAnchors);
                    $row = array_values($projection->project($row));
                }
                if ($maxColumns !== null && count($row) > $maxColumns) {
                    throw MnbExcelException::withCode(
                        'XLSX column limit exceeded on row ' . $rowNumber . '. Columns: ' . count($row) . ', max_columns: ' . $maxColumns,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'row' => $rowNumber, 'columns' => count($row), 'max_columns' => $maxColumns]
                    );
                }

                $cellCount += count($row);
                if ($maxCells !== null && $cellCount > $maxCells) {
                    throw new MnbExcelException('Cell limit exceeded while reading XLSX. Cells: ' . $cellCount . ', max_cells: ' . $maxCells);
                }

                $delivered++;
                if ($maxRows !== null && $delivered > $maxRows) {
                    throw new MnbExcelException('Row limit exceeded while reading XLSX. Rows: ' . $delivered . ', max_rows: ' . $maxRows);
                }

                yield $rowNumber - 1 => $row;
            }
        } finally {
            $reader->close();
            $sharedStrings->close();
        }
    }


    /**
     * Read non-tabular XLSX metadata for a sheet: rich text runs, comments, hyperlinks, and advanced object inventory.
     *
     * @return array<string,mixed>
     */
    public function readSheetMetadata(string $path, int|string $sheet = 1, array $options = []): array
    {
        $metadata = (new XlsxMetadataExtractor($this->resolver))->readSheetMetadata($path, $sheet);
        $realPath = realpath($path);
        if ($realPath !== false && class_exists(ZipArchive::class)) {
            $sheetXml = $this->resolver->resolveSheetPath($realPath, $sheet);
            $merged = MergedCellMap::fromXlsx($realPath, $sheetXml, $options)->ranges();
            $metadata['merged_cells'] = $merged;
            $metadata['summary']['merged_ranges'] = count($merged);
        }
        return $metadata;
    }

    private function readSharedStrings(string $realPath, array $options = []): SharedStringProviderInterface
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package.');
        }
        $stat = $zip->statName('xl/sharedStrings.xml');
        $maxBytes = isset($options['max_shared_strings_bytes']) ? max(0, (int) $options['max_shared_strings_bytes']) : null;
        if ($maxBytes !== null && is_array($stat) && (int) ($stat['size'] ?? 0) > $maxBytes) {
            $zip->close();
            throw MnbExcelException::withCode('XLSX shared strings exceed max_shared_strings_bytes.', ErrorCode::FILE_READ_FAILED, ['size_bytes' => (int) ($stat['size'] ?? 0), 'max_shared_strings_bytes' => $maxBytes]);
        }
        $exists = $zip->locateName('xl/sharedStrings.xml') !== false;
        $zip->close();
        if (!$exists) {
            return new InMemorySharedStringProvider();
        }

        // Streaming mode may opt into the disk-backed provider shared with the
        // large reader without making it a hard dependency of lightweight XLSX.
        $providerClass = 'Mnb\PHPExcel\Large\LargeSharedStringCache';
        $mode = strtolower((string) ($options['shared_strings_mode'] ?? 'memory'));
        if (in_array($mode, ['auto', 'sqlite', 'disk'], true) && class_exists($providerClass)) {
            /** @var SharedStringProviderInterface $provider */
            $provider = $providerClass::fromXlsx($realPath, $options);
            return $provider;
        }

        $reader = new XMLReader();
        if (!@$reader->open($this->zipUri($realPath, 'xl/sharedStrings.xml'), null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to open XLSX shared strings.', ErrorCode::XLSX_INVALID);
        }
        $strings = [];
        $maxStrings = isset($options['max_shared_strings']) ? max(0, (int) $options['max_shared_strings']) : null;
        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || $reader->localName !== 'si') {
                    continue;
                }
                $strings[] = $this->textFromRichXml($reader->readOuterXml());
                if ($maxStrings !== null && count($strings) > $maxStrings) {
                    throw MnbExcelException::withCode('XLSX shared string count exceeds max_shared_strings.', ErrorCode::FILE_READ_FAILED, ['shared_strings' => count($strings), 'max_shared_strings' => $maxStrings]);
                }
            }
        } finally {
            $reader->close();
        }
        return new InMemorySharedStringProvider($strings);
    }

    private function readStyleMap(string $realPath): XlsxStyleMap
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package.');
        }

        $stylesXml = $zip->getFromName('xl/styles.xml');
        $zip->close();

        return XlsxStyleMap::fromXml($stylesXml !== false ? $stylesXml : null);
    }

    private function usesDate1904(string $realPath): bool
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            return false;
        }

        $workbookXml = $zip->getFromName('xl/workbook.xml');
        $zip->close();

        if ($workbookXml === false) {
            return false;
        }

        return preg_match('/<workbookPr\b[^>]*date1904\s*=\s*("1"|"true"|\'1\'|\'true\')/i', $workbookXml) === 1;
    }

    /**
     * Reads one <row> element. The XMLReader must be positioned on the row start element.
     *
     * @param list<string> $sharedStrings
     * @return list<mixed>
     */
    private function readRow(XMLReader $reader, SharedStringProviderInterface $sharedStrings, XlsxStyleMap $styleMap, bool $date1904, array $options, array $hiddenColumns, ColumnProjection $projection): array
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

            $cellRef = (string) $reader->getAttribute('r');
            $type = (string) $reader->getAttribute('t');
            $styleIndexAttr = $reader->getAttribute('s');
            $styleIndex = $styleIndexAttr !== null && $styleIndexAttr !== '' ? (int) $styleIndexAttr : null;
            $columnIndex = $cellRef !== '' ? Coordinate::splitCellRef(preg_replace('/\$/', '', $cellRef) ?: $cellRef)[0] : count($cells) + 1;
            if (isset($hiddenColumns[$columnIndex]) || ($projection->active() && !$projection->includesIndex($columnIndex))) {
                $this->skipCurrentElement($reader);
                continue;
            }
            $xml = $reader->readOuterXml();
            $cells[$columnIndex - 1] = $this->cellValueFromXml($xml, $type, $styleIndex, $sharedStrings, $styleMap, $date1904, $options);
        }

        if ($cells === []) {
            return [];
        }

        ksort($cells);
        if ($projection->active() && $projection->compact()) {
            return array_values($cells);
        }
        $maxIndex = max(array_keys($cells));
        $row = [];
        for ($i = 0; $i <= $maxIndex; $i++) {
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

    private function cellValueFromXml(string $xml, string $type, ?int $styleIndex, SharedStringProviderInterface|array $sharedStrings, XlsxStyleMap $styleMap, bool $date1904, array $options): mixed
    {
        $formulaMode = (string) ($options['formula_cells'] ?? 'formula');
        $formula = $this->formulaFromXml($xml);
        if ($formula !== null && $formulaMode === 'formula') {
            return '=' . $formula['expression'];
        }

        $cached = $this->cachedCellValue($xml, $type, $styleIndex, $sharedStrings, $styleMap, $date1904, $options);
        if ($formula !== null && $formulaMode === 'both') {
            return new FormulaResult(
                '=' . $formula['expression'],
                $cached,
                $this->formulaResultType($type, $cached),
                $formula['metadata']
            );
        }
        return $cached;
    }

    /** @return array{expression:string,metadata:array<string,mixed>}|null */
    private function formulaFromXml(string $xml): ?array
    {
        if (preg_match('/<f\b([^>]*)>(.*?)<\/f>/s', $xml, $match) !== 1
            && preg_match('/<f\b([^>]*)\/>/s', $xml, $match) !== 1) {
            return null;
        }
        $attributes = (string) ($match[1] ?? '');
        $expression = html_entity_decode(strip_tags((string) ($match[2] ?? '')), ENT_QUOTES | ENT_XML1, 'UTF-8');
        $metadata = [];
        foreach (['t', 'si', 'ref'] as $name) {
            if (preg_match('/\b' . $name . '\s*=\s*(?:"([^"]*)"|\'([^\']*)\')/i', $attributes, $attributeMatch) === 1) {
                $metadata[$name] = html_entity_decode((string) ($attributeMatch[1] !== '' ? $attributeMatch[1] : $attributeMatch[2]), ENT_QUOTES | ENT_XML1, 'UTF-8');
            }
        }
        if ($expression === '' && isset($metadata['si'])) {
            $expression = '#SHARED_FORMULA[' . $metadata['si'] . ']';
        }
        return ['expression' => $expression, 'metadata' => $metadata];
    }

    private function cachedCellValue(string $xml, string $type, ?int $styleIndex, SharedStringProviderInterface|array $sharedStrings, XlsxStyleMap $styleMap, bool $date1904, array $options): mixed
    {
        if ($type === 'inlineStr') {
            return $this->textFromRichXml($xml);
        }
        if ($type === 'str' || $type === 'e') {
            return $this->readV($xml) ?? ($type === 'str' ? '' : null);
        }
        $value = $this->readV($xml);
        if ($value === null) {
            return null;
        }
        if ($type === 's') {
            return $sharedStrings instanceof SharedStringProviderInterface
                ? $sharedStrings->get((int) $value)
                : ($sharedStrings[(int) $value] ?? '');
        }
        if ($type === 'b') {
            return $value === '1' || strtolower($value) === 'true';
        }
        if ($type === 'd') {
            return $this->formatIsoDate($value, $options);
        }
        if (is_numeric($value)) {
            if (($options['format_dates'] ?? true) && $styleMap->isDateStyle($styleIndex)) {
                return $this->formatExcelSerialDate((float) $value, $date1904, $options);
            }
            if (($options['preserve_numeric_strings'] ?? false) === true) {
                return (string) $value;
            }
            return LocaleNormalizer::parseCanonicalNumber((string) $value, $options);
        }
        return $value;
    }

    private function formulaResultType(string $cellType, mixed $value): string
    {
        if ($cellType === 'e') {
            return 'error';
        }
        return match (true) {
            is_bool($value) => 'boolean',
            is_int($value), is_float($value) => 'number',
            $value instanceof \DateTimeInterface => 'datetime',
            is_string($value) => 'string',
            $value === null => 'blank',
            default => 'mixed',
        };
    }

    private function readV(string $xml): ?string
    {
        if (!preg_match('/<v[^>]*>(.*?)<\/v>/s', $xml, $match)) {
            return null;
        }

        return html_entity_decode(strip_tags($match[1]), ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    private function textFromRichXml(string $xml): string
    {
        preg_match_all('/<t\b[^>]*>(.*?)<\/t>/s', $xml, $matches);
        $value = '';
        foreach ($matches[1] ?? [] as $part) {
            $value .= html_entity_decode(strip_tags($part), ENT_QUOTES | ENT_XML1, 'UTF-8');
        }

        return $value;
    }

    private function formatIsoDate(string $value, array $options): mixed
    {
        try {
            $date = new DateTimeImmutable($value);
        } catch (\Throwable) {
            return $value;
        }

        if (($options['return_datetime'] ?? false) === true) {
            return $date;
        }

        $format = str_contains($value, 'T') ? (string) ($options['datetime_format'] ?? 'Y-m-d H:i:s') : (string) ($options['date_format'] ?? 'Y-m-d');
        return $date->format($format);
    }

    private function formatExcelSerialDate(float $serial, bool $date1904, array $options): mixed
    {
        if ($serial < 0) {
            return $serial;
        }

        $wholeDays = (int) floor($serial);
        $seconds = (int) round(($serial - $wholeDays) * 86400);

        if ($date1904) {
            $base = new DateTimeImmutable('1904-01-01 00:00:00');
            $date = $base->modify('+' . $wholeDays . ' days')->modify('+' . $seconds . ' seconds');
        } else {
            // Excel's 1900 date system includes the historic fake 1900-02-29 day.
            $base = new DateTimeImmutable('1899-12-31 00:00:00');
            $offset = $wholeDays > 59 ? $wholeDays - 1 : $wholeDays;
            $date = $base->modify('+' . $offset . ' days')->modify('+' . $seconds . ' seconds');
        }

        if (($options['return_datetime'] ?? false) === true) {
            return $date;
        }

        $hasTime = $seconds !== 0;
        $format = $hasTime ? (string) ($options['datetime_format'] ?? 'Y-m-d H:i:s') : (string) ($options['date_format'] ?? 'Y-m-d');
        return $date->format($format);
    }

    /** @return array<int,true> */
    private function hiddenColumnIndexes(string $realPath, string $sheetXml): array
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            return [];
        }

        $xml = $zip->getFromName($sheetXml);
        $zip->close();
        if ($xml === false || $xml === '') {
            return [];
        }

        $hidden = [];
        preg_match_all('/<col\b[^>]*>/i', $xml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            if (preg_match('/\bhidden\s*=\s*("1"|"true"|\'1\'|\'true\')/i', $tag) !== 1) {
                continue;
            }
            $min = $this->intAttribute($tag, 'min') ?? 1;
            $max = $this->intAttribute($tag, 'max') ?? $min;
            for ($i = max(1, $min); $i <= max($min, $max); $i++) {
                $hidden[$i] = true;
            }
        }

        return $hidden;
    }

    private function intAttribute(string $tag, string $name): ?int
    {
        if (preg_match('/\b' . preg_quote($name, '/') . '\s*=\s*("(\d+)"|\'(\d+)\')/i', $tag, $match) !== 1) {
            return null;
        }

        return (int) ($match[2] !== '' ? $match[2] : $match[3]);
    }


    /** @param list<int|string> $columns @return array<int,true> */
    private function normalizeOnlyColumns(array $columns): array
    {
        $normalized = [];
        foreach ($columns as $column) {
            if (is_int($column) || ctype_digit((string) $column)) {
                $index = (int) $column;
            } else {
                $index = Coordinate::columnNameToIndex((string) $column);
            }
            if ($index < 1) {
                throw new MnbExcelException('Selected column indexes must be greater than zero.');
            }
            $normalized[$index] = true;
        }
        return $normalized;
    }

    private function assertZipEntrySize(string $realPath, string $entry, array $options, string $optionName): void
    {
        if (!isset($options[$optionName])) {
            return;
        }

        $maxBytes = max(0, (int) $options[$optionName]);
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw MnbExcelException::withCode('Unable to open XLSX zip package.', ErrorCode::XLSX_INVALID);
        }
        try {
            $stat = $zip->statName($entry);
            if (is_array($stat) && (int) ($stat['size'] ?? 0) > $maxBytes) {
                throw MnbExcelException::withCode(
                    'XLSX entry ' . $entry . ' exceeds ' . $optionName . '.',
                    ErrorCode::FILE_READ_FAILED,
                    ['entry' => $entry, 'size_bytes' => (int) ($stat['size'] ?? 0), $optionName => $maxBytes]
                );
            }
        } finally {
            $zip->close();
        }
    }

    private function zipUri(string $realPath, string $entry): string
    {
        return 'zip://' . str_replace('\\', '/', $realPath) . '#' . $entry;
    }

    private function zipEntryExists(string $realPath, string $entry): bool
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package.');
        }

        $exists = $zip->locateName($entry) !== false;
        $zip->close();
        return $exists;
    }

    private function ensureExtensions(): void
    {
        if (!class_exists(ZipArchive::class)) {
            throw new MnbExcelException('ext-zip is required to read XLSX files.');
        }
        if (!class_exists(XMLReader::class)) {
            throw new MnbExcelException('ext-xmlreader is required to read XLSX files.');
        }
    }
}
