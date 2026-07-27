<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Xml\XmlReader;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/**
 * Lightweight XLSX package information without hydrating worksheet cell rows.
 *
 * Only ZIP package metadata, workbook relationships, worksheet dimensions, and
 * optionally streamed worksheet XML tags are inspected.
 */
final class XlsxQuickInfo
{
    public function __construct(private ?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /**
     * Return file/package information without loading workbook cells.
     *
     * @param array<string,mixed> $options password
     * @return array<string,mixed>
     */
    public function fileInfo(string $path, array $options = []): array
    {
        $realPath = $this->requireFile($path);
        $encryption = new XlsxEncryption();
        $encrypted = $encryption->isEncryptedFile($realPath);
        $base = $this->baseFileInfo($realPath, $encrypted, $encryption->encryptionMode($realPath));

        if ($encrypted) {
            $password = (string) ($options['password'] ?? '');
            if ($password === '') {
                return $base + [
                    'status' => 'password_required',
                    'sheet_count' => null,
                    'sheet_names' => [],
                    'zip_entries' => null,
                    'compressed_package_bytes' => null,
                    'uncompressed_package_bytes' => null,
                    'compression_ratio' => null,
                    'has_macros' => null,
                    'properties' => [],
                    'warnings' => ['A password is required to inspect workbook package metadata.'],
                ];
            }

            $temporary = $encryption->decryptToTemporary($realPath, $password, $options);
            try {
                $package = $this->packageFileInfo($temporary);
                return array_replace($package, $base, [
                    'status' => 'ok',
                    'encrypted' => true,
                    'encryption_mode' => $encryption->encryptionMode($realPath),
                    'warnings' => ['Workbook metadata was inspected from a temporary decrypted package.'],
                ]);
            } finally {
                @unlink($temporary);
            }
        }

        return array_replace($this->packageFileInfo($realPath), $base, [
            'status' => 'ok',
            'warnings' => [],
        ]);
    }

    /**
     * Return information for every worksheet without loading cell values.
     *
     * Set accurate_row_count=true to stream worksheet XML and calculate
     * physical_rows, filled_rows, cells, and the last used row/column.
     *
     * @param array<string,mixed> $options password, accurate_row_count, include_hidden
     * @return list<array<string,mixed>>
     */
    public function sheetsInfo(string $path, array $options = []): array
    {
        return $this->withReadablePackage($path, $options, function (string $realPath) use ($options): array {
            $zip = $this->openZip($realPath);
            try {
                $accurate = (bool) ($options['accurate_row_count'] ?? false);
                $includeHidden = (bool) ($options['include_hidden'] ?? true);
                $result = [];
                foreach ($this->resolver->sheets($realPath) as $sheet) {
                    if (!$includeHidden && ($sheet['state'] ?? 'visible') !== 'visible') {
                        continue;
                    }
                    $result[] = $this->inspectSheet($realPath, $zip, $sheet, $accurate);
                }
                return $result;
            } finally {
                $zip->close();
            }
        });
    }

    /**
     * Return one worksheet's information without loading cell values.
     *
     * @param array<string,mixed> $options password, accurate_row_count
     * @return array<string,mixed>
     */
    public function sheetInfo(string $path, int|string $sheet = 1, array $options = []): array
    {
        return $this->withReadablePackage($path, $options, function (string $realPath) use ($sheet, $options): array {
            $zip = $this->openZip($realPath);
            try {
                $selected = $this->selectSheet($this->resolver->sheets($realPath), $sheet);
                return $this->inspectSheet(
                    $realPath,
                    $zip,
                    $selected,
                    (bool) ($options['accurate_row_count'] ?? false)
                );
            } finally {
                $zip->close();
            }
        });
    }

    /**
     * Return a row count without converting worksheet rows to PHP arrays.
     *
     * Supported modes:
     * - filled (default): rows containing at least one cell
     * - physical: physical <row> elements, including empty styled rows
     * - last_row: highest row index referenced in worksheet XML
     * - declared: last row declared by the worksheet dimension
     *
     * @param array<string,mixed> $options password, mode
     */
    public function rowCount(string $path, int|string $sheet = 1, array $options = []): int
    {
        $mode = strtolower((string) ($options['mode'] ?? 'filled'));
        if (!in_array($mode, ['filled', 'physical', 'last_row', 'declared'], true)) {
            throw new MnbExcelException('Row count mode must be filled, physical, last_row, or declared.');
        }

        $accurate = $mode !== 'declared';
        $info = $this->sheetInfo($path, $sheet, array_replace($options, [
            'accurate_row_count' => $accurate,
        ]));

        return match ($mode) {
            'physical' => (int) ($info['physical_rows'] ?? 0),
            'last_row' => (int) ($info['last_row'] ?? 0),
            'declared' => (int) ($info['declared_last_row'] ?? 0),
            default => (int) ($info['filled_rows'] ?? 0),
        };
    }

    /**
     * Return row counts for all worksheets without loading cell values.
     *
     * @param array<string,mixed> $options password, mode, include_hidden
     * @return array<string,int>
     */
    public function rowCounts(string $path, array $options = []): array
    {
        $mode = strtolower((string) ($options['mode'] ?? 'filled'));
        if (!in_array($mode, ['filled', 'physical', 'last_row', 'declared'], true)) {
            throw new MnbExcelException('Row count mode must be filled, physical, last_row, or declared.');
        }

        $infos = $this->sheetsInfo($path, array_replace($options, [
            'accurate_row_count' => $mode !== 'declared',
        ]));
        $counts = [];
        foreach ($infos as $info) {
            $counts[(string) $info['name']] = match ($mode) {
                'physical' => (int) ($info['physical_rows'] ?? 0),
                'last_row' => (int) ($info['last_row'] ?? 0),
                'declared' => (int) ($info['declared_last_row'] ?? 0),
                default => (int) ($info['filled_rows'] ?? 0),
            };
        }
        return $counts;
    }

    /** @return array<string,mixed> */
    private function packageFileInfo(string $realPath): array
    {
        $zip = $this->openZip($realPath);
        try {
            $compressed = 0;
            $uncompressed = 0;
            $hasMacros = false;
            for ($index = 0; $index < $zip->numFiles; $index++) {
                $name = (string) ($zip->getNameIndex($index) ?: '');
                $stat = $zip->statIndex($index) ?: [];
                $compressed += (int) ($stat['comp_size'] ?? 0);
                $uncompressed += (int) ($stat['size'] ?? 0);
                if (strtolower($name) === 'xl/vbaproject.bin') {
                    $hasMacros = true;
                }
            }

            $sheets = $this->resolver->sheets($realPath);
            return [
                'status' => 'ok',
                'sheet_count' => count($sheets),
                'sheet_names' => array_values(array_map(static fn (array $sheet): string => (string) $sheet['name'], $sheets)),
                'zip_entries' => $zip->numFiles,
                'compressed_package_bytes' => $compressed,
                'uncompressed_package_bytes' => $uncompressed,
                'compression_ratio' => $uncompressed > 0 ? round($compressed / $uncompressed, 4) : null,
                'has_macros' => $hasMacros,
                'properties' => $this->coreProperties($zip),
            ];
        } finally {
            $zip->close();
        }
    }

    /** @return array<string,mixed> */
    private function baseFileInfo(string $realPath, bool $encrypted, ?string $encryptionMode): array
    {
        $size = filesize($realPath) ?: 0;
        $extension = strtolower((string) pathinfo($realPath, PATHINFO_EXTENSION));
        $modified = filemtime($realPath);
        return [
            'file' => $realPath,
            'name' => basename($realPath),
            'extension' => $extension,
            'format' => in_array($extension, ['xlsm', 'xltm'], true) ? 'xlsm' : 'xlsx',
            'mime_type' => in_array($extension, ['xlsm', 'xltm'], true)
                ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
                : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'size_bytes' => (int) $size,
            'size_mb' => round(((int) $size) / 1048576, 2),
            'modified_at' => $modified === false ? null : date(DATE_ATOM, $modified),
            'readable' => is_readable($realPath),
            'encrypted' => $encrypted,
            'encryption_mode' => $encryptionMode,
        ];
    }

    /** @param array<string,mixed> $sheet @return array<string,mixed> */
    private function inspectSheet(string $realPath, ZipArchive $zip, array $sheet, bool $accurate): array
    {
        $stat = $zip->statName((string) $sheet['path']) ?: [];
        $dimension = $this->dimensionFromStream($zip, (string) $sheet['path']);
        $mustScan = $accurate || $dimension['ref'] === '';
        $scan = $mustScan
            ? $this->scanSheet($realPath, (string) $sheet['path'])
            : [
                'physical_rows' => null,
                'filled_rows' => null,
                'last_row' => $dimension['rows'],
                'last_column_index' => $dimension['columns'],
                'cells' => null,
                'scan_mode' => 'not_scanned',
            ];

        $lastColumnIndex = $mustScan
            ? (int) ($scan['last_column_index'] ?? 0)
            : (int) $dimension['columns'];
        $lastRow = $mustScan
            ? (int) ($scan['last_row'] ?? 0)
            : (int) $dimension['rows'];

        return [
            'index' => (int) $sheet['index'],
            'name' => (string) $sheet['name'],
            'sheet_id' => (int) $sheet['sheet_id'],
            'state' => (string) $sheet['state'],
            'path' => (string) $sheet['path'],
            'exists' => (bool) $sheet['exists'],
            'dimension' => $dimension['ref'] !== '' ? $dimension['ref'] : null,
            'declared_last_row' => (int) $dimension['rows'],
            'declared_last_column' => $this->columnLetters((int) $dimension['columns']),
            'declared_columns' => (int) $dimension['columns'],
            'last_row' => $lastRow,
            'last_column' => $this->columnLetters($lastColumnIndex),
            'columns' => $lastColumnIndex,
            'physical_rows' => $scan['physical_rows'],
            'filled_rows' => $scan['filled_rows'],
            'cells' => $scan['cells'],
            'accurate_row_count' => $mustScan,
            'scan_mode' => (string) $scan['scan_mode'],
            'worksheet_xml_size_bytes' => (int) ($stat['size'] ?? 0),
            'worksheet_xml_compressed_bytes' => (int) ($stat['comp_size'] ?? 0),
        ];
    }

    /** @return array{ref:string,rows:int,columns:int} */
    private function dimensionFromStream(ZipArchive $zip, string $path): array
    {
        $stream = $zip->getStream($path);
        if (!is_resource($stream)) {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        try {
            $chunk = stream_get_contents($stream, 65536);
        } finally {
            fclose($stream);
        }
        if (!is_string($chunk) || $chunk === '') {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        if (preg_match('/<dimension\b[^>]*\bref\s*=\s*("([^"]+)"|\'([^\']+)\')/i', $chunk, $matches) !== 1) {
            return ['ref' => '', 'rows' => 0, 'columns' => 0];
        }
        return $this->dimensionFromRef($matches[2] !== '' ? $matches[2] : $matches[3]);
    }

    /** @return array{ref:string,rows:int,columns:int} */
    private function dimensionFromRef(string $ref): array
    {
        $ref = trim($ref);
        $last = str_contains($ref, ':') ? substr($ref, strrpos($ref, ':') + 1) : $ref;
        if ($ref === '' || preg_match('/^\$?([A-Z]{1,3})\$?([0-9]+)$/i', $last, $matches) !== 1) {
            return ['ref' => $ref, 'rows' => 0, 'columns' => 0];
        }
        return [
            'ref' => $ref,
            'rows' => (int) $matches[2],
            'columns' => $this->columnIndex($matches[1]),
        ];
    }

    /** @return array{physical_rows:int,filled_rows:int,last_row:int,last_column_index:int,cells:int,scan_mode:string} */
    private function scanSheet(string $realPath, string $sheetPath): array
    {
        $reader = new XmlReader();
        $uri = 'zip://' . str_replace('\\', '/', $realPath) . '#' . $sheetPath;
        if (!@$reader->open($uri, null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to inspect worksheet XML: ' . $sheetPath, ErrorCode::XLSX_INVALID, ['sheet' => $sheetPath]);
        }

        $physicalRows = 0;
        $filledRows = 0;
        $lastRow = 0;
        $lastColumnIndex = 0;
        $cells = 0;
        $inRow = false;
        $rowHasCell = false;

        while ($reader->read()) {
            if ($reader->nodeType === XmlReader::ELEMENT) {
                if ($reader->localName === 'row') {
                    $inRow = true;
                    $rowHasCell = false;
                    $physicalRows++;
                    $lastRow = max($lastRow, (int) ($reader->getAttribute('r') ?: $physicalRows));
                } elseif ($reader->localName === 'c') {
                    $cells++;
                    $rowHasCell = true;
                    $ref = strtoupper((string) ($reader->getAttribute('r') ?? ''));
                    if (preg_match('/^\$?([A-Z]{1,3})\$?([0-9]+)$/', $ref, $matches) === 1) {
                        $lastColumnIndex = max($lastColumnIndex, $this->columnIndex($matches[1]));
                        $lastRow = max($lastRow, (int) $matches[2]);
                    }
                }
            } elseif ($reader->nodeType === XmlReader::END_ELEMENT && $reader->localName === 'row' && $inRow) {
                if ($rowHasCell) {
                    $filledRows++;
                }
                $inRow = false;
            }
        }
        $reader->close();

        return [
            'physical_rows' => $physicalRows,
            'filled_rows' => $filledRows,
            'last_row' => $lastRow,
            'last_column_index' => $lastColumnIndex,
            'cells' => $cells,
            'scan_mode' => XmlReader::nativeAvailable() ? 'streaming_xmlreader' : 'buffered_xml_fallback',
        ];
    }

    /** @param list<array<string,mixed>> $sheets @return array<string,mixed> */
    private function selectSheet(array $sheets, int|string $selected): array
    {
        if (is_int($selected) || ctype_digit((string) $selected)) {
            $index = (int) $selected;
            if ($index < 1 || !isset($sheets[$index - 1])) {
                throw new MnbExcelException('Worksheet not found: ' . (string) $selected);
            }
            return $sheets[$index - 1];
        }

        foreach ($sheets as $sheet) {
            if ((string) $sheet['name'] === (string) $selected) {
                return $sheet;
            }
        }
        $needle = strtolower((string) $selected);
        $matches = array_values(array_filter($sheets, static fn (array $sheet): bool => strtolower((string) $sheet['name']) === $needle));
        if (count($matches) > 1) {
            throw new MnbExcelException('Multiple worksheets match the name: ' . (string) $selected);
        }
        if ($matches === []) {
            throw new MnbExcelException('Worksheet not found: ' . (string) $selected);
        }
        return $matches[0];
    }

    /** @return array<string,string|null> */
    private function coreProperties(ZipArchive $zip): array
    {
        $xml = $zip->getFromName('docProps/core.xml');
        if (!is_string($xml) || $xml === '') {
            return [];
        }
        $properties = [];
        foreach ([
            'title' => 'title',
            'subject' => 'subject',
            'creator' => 'creator',
            'keywords' => 'keywords',
            'description' => 'description',
            'last_modified_by' => 'lastModifiedBy',
            'created' => 'created',
            'modified' => 'modified',
        ] as $key => $tag) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '>/is', $xml, $matches) === 1) {
                $properties[$key] = html_entity_decode(trim(strip_tags($matches[1])), ENT_QUOTES | ENT_XML1, 'UTF-8');
            }
        }
        return $properties;
    }

    /** @template T @param callable(string):T $callback @return T */
    private function withReadablePackage(string $path, array $options, callable $callback): mixed
    {
        $realPath = $this->requireFile($path);
        $encryption = new XlsxEncryption();
        if (!$encryption->isEncryptedFile($realPath)) {
            return $callback($realPath);
        }

        $password = (string) ($options['password'] ?? '');
        if ($password === '') {
            throw new MnbExcelException('A password is required to inspect this encrypted XLSX file.');
        }
        $temporary = $encryption->decryptToTemporary($realPath, $password, $options);
        try {
            return $callback($temporary);
        } finally {
            @unlink($temporary);
        }
    }

    private function requireFile(string $path): string
    {
        $realPath = realpath($path);
        if ($realPath === false || !is_file($realPath)) {
            throw MnbExcelException::withCode('XLSX file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        return $realPath;
    }

    private function openZip(string $path): ZipArchive
    {
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            throw MnbExcelException::withCode('Unable to open XLSX package: ' . $path, ErrorCode::XLSX_INVALID, ['path' => $path]);
        }
        return $zip;
    }

    private function columnIndex(string $letters): int
    {
        $index = 0;
        foreach (str_split(strtoupper($letters)) as $letter) {
            $index = $index * 26 + ord($letter) - 64;
        }
        return $index;
    }

    private function columnLetters(int $index): ?string
    {
        if ($index < 1) {
            return null;
        }
        $letters = '';
        while ($index > 0) {
            $index--;
            $letters = chr(65 + ($index % 26)) . $letters;
            $index = intdiv($index, 26);
        }
        return $letters;
    }
}
