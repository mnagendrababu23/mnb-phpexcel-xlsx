<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Xml\XmlReader;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Security\XlsxEncryption;

final class XlsxInspector
{
    private XlsxWorkbookResolver $resolver;

    public function __construct(?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /**
     * Inspect workbook structure without converting the workbook to arrays.
     *
     * @return array{
     *   status:string,
     *   file:string,
     *   size_bytes:int,
     *   encrypted:bool,
     *   sheets:list<array<string,mixed>>,
     *   warnings:list<string>,
     *   errors:list<string>
     * }
     */
    public function inspect(string $path, array $options = []): array
    {
        $encryption = new XlsxEncryption();
        if ($encryption->isEncryptedFile($path)) {
            $password = (string) ($options['password'] ?? '');
            if ($password === '') {
                return [
                    'status' => 'password_required',
                    'file' => $path,
                    'size_bytes' => is_file($path) ? (int) filesize($path) : 0,
                    'encrypted' => true,
                    'sheets' => [],
                    'warnings' => [],
                    'errors' => ['A password is required to inspect this encrypted XLSX file.'],
                ];
            }
            $temporary = $encryption->decryptToTemporary($path, $password, $options);
            try {
                $result = $this->inspect($temporary, []);
                $result['file'] = $path;
                $result['size_bytes'] = is_file($path) ? (int) filesize($path) : 0;
                $result['encrypted'] = true;
                return $result;
            } finally {
                @unlink($temporary);
            }
        }
        $realPath = realpath($path);
        $warnings = [];
        $errors = [];

        $result = [
            'status' => 'ok',
            'file' => $path,
            'size_bytes' => is_file($path) ? (int) filesize($path) : 0,
            'encrypted' => false,
            'sheets' => [],
            'warnings' => &$warnings,
            'errors' => &$errors,
        ];

        if ($realPath === false || !is_file($realPath)) {
            $errors[] = 'File does not exist.';
            $result['status'] = 'failed';
            return $result;
        }

        if (!class_exists(ZipArchive::class)) {
            $errors[] = 'ext-zip is required to inspect XLSX files.';
            $result['status'] = 'failed';
            return $result;
        }

        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            $errors[] = 'Unable to open XLSX zip package. File may be corrupted or not a valid XLSX file.';
            $result['status'] = 'failed';
            return $result;
        }

        $result['encrypted'] = $zip->locateName('EncryptionInfo') !== false || $zip->locateName('EncryptedPackage') !== false;
        if ($result['encrypted']) {
            $errors[] = 'Unexpected encrypted entries were found inside the OOXML ZIP package.';
        }

        foreach (['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'] as $entry) {
            if ($zip->locateName($entry) === false) {
                $errors[] = 'Missing required XLSX part: ' . $entry;
            }
        }

        if ($zip->locateName('xl/styles.xml') === false) {
            $warnings[] = 'styles.xml is missing. Date style detection may be limited.';
        }

        if ($zip->locateName('xl/sharedStrings.xml') === false) {
            $warnings[] = 'sharedStrings.xml is missing. This is valid for inline-string workbooks.';
        }

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = (string) $zip->getNameIndex($i);
            $lower = strtolower($name);
            if (str_ends_with($lower, 'vbaproject.bin')) {
                $warnings[] = 'VBA macro project detected. Macros are never executed. Use preserveAdvancedObjectsFrom() only when you intentionally need to copy advanced package parts forward.';
            }
            if (str_starts_with($lower, 'xl/externallinks/')) {
                $warnings[] = 'External workbook links detected: ' . $name;
            }
            if (str_starts_with($lower, 'xl/comments') || str_contains($lower, '/comments')) {
                $warnings[] = 'Comments/notes detected. Comment content is available through sheetMetadata() and structured metadata output.';
            }
            if ($lower === 'xl/calcchain.xml') {
                $warnings[] = 'Formula calculation chain detected. Formula results may depend on Excel recalculation.';
            }
            if (str_starts_with($lower, 'xl/pivottables/') || str_starts_with($lower, 'xl/pivotcache/')) {
                $warnings[] = 'Pivot table/cache parts detected. Pivot data is not imported as rows; package parts can be preserved from a source template with preserveAdvancedObjectsFrom().';
            }
        }

        $warnings = array_values(array_unique($warnings));
        $zip->close();

        try {
            $sheets = $this->resolver->sheets($realPath);
            foreach ($sheets as $sheet) {
                $result['sheets'][] = array_merge($sheet, $this->inspectSheetXml($realPath, $sheet['path'], $sheet['exists']));
                if (!$sheet['exists']) {
                    $errors[] = 'Missing worksheet XML for sheet "' . $sheet['name'] . '": ' . $sheet['path'];
                }
                if ($sheet['state'] !== 'visible') {
                    $warnings[] = 'Sheet "' . $sheet['name'] . '" is ' . $sheet['state'] . '.';
                }
            }

            if ($sheets === []) {
                $errors[] = 'Workbook does not contain any sheets.';
            }
        } catch (MnbExcelException $e) {
            $errors[] = $e->getMessage();
        }

        if ($this->xmlReaderAvailable()) {
            foreach (['xl/workbook.xml', 'xl/styles.xml', 'xl/sharedStrings.xml'] as $entry) {
                $this->checkXmlEntry($realPath, $entry, $warnings, $errors);
            }
        } else {
            $warnings[] = 'ext-xmlreader is not available. XML validity checks were skipped.';
        }

        $result['status'] = $errors === [] ? ($warnings === [] ? 'ok' : 'warning') : 'failed';
        return $result;
    }

    /** @return list<string> */
    public function sheetNames(string $path, array $options = []): array
    {
        $encryption = new XlsxEncryption();
        if ($encryption->isEncryptedFile($path)) {
            $password = (string) ($options['password'] ?? '');
            if ($password === '') {
                throw new MnbExcelException('A password is required to read this encrypted XLSX file.');
            }
            $temporary = $encryption->decryptToTemporary($path, $password, $options);
            try {
                return $this->sheetNames($temporary, []);
            } finally {
                @unlink($temporary);
            }
        }
        $realPath = realpath($path);
        if ($realPath === false) {
            throw new MnbExcelException('Invalid XLSX path: ' . $path);
        }

        return array_map(static fn(array $sheet): string => $sheet['name'], $this->resolver->sheets($realPath));
    }

    /** @return array<string,mixed> */
    private function inspectSheetXml(string $realPath, string $sheetPath, bool $exists): array
    {
        $base = [
            'dimension' => null,
            'declared_last_row' => null,
            'declared_last_column' => null,
            'row_tag_count' => null,
            'hidden_row_count' => null,
            'hidden_column_count' => null,
            'has_merge_cells' => false,
            'has_auto_filter' => false,
            'has_drawing' => false,
        ];

        if (!$exists) {
            return $base;
        }

        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            return $base;
        }

        $xml = $zip->getFromName($sheetPath);
        $zip->close();
        if ($xml === false) {
            return $base;
        }

        if (preg_match('/<dimension\b[^>]*ref\s*=\s*("([^"]*)"|\'([^\']*)\')/i', $xml, $dimension)) {
            $ref = $dimension[2] !== '' ? $dimension[2] : $dimension[3];
            $base['dimension'] = $ref;
            $lastCell = str_contains($ref, ':') ? substr($ref, strrpos($ref, ':') + 1) : $ref;
            if (preg_match('/([A-Z]+)(\d+)$/i', $lastCell, $match)) {
                $base['declared_last_column'] = strtoupper($match[1]);
                $base['declared_last_row'] = (int) $match[2];
            }
        }

        $base['row_tag_count'] = preg_match_all('/<row\b/i', $xml);
        $base['hidden_row_count'] = preg_match_all('/<row\b[^>]*hidden\s*=\s*("1"|"true"|\'1\'|\'true\')/i', $xml);
        $base['hidden_column_count'] = preg_match_all('/<col\b[^>]*hidden\s*=\s*("1"|"true"|\'1\'|\'true\')/i', $xml);
        $base['has_merge_cells'] = str_contains($xml, '<mergeCells');
        $base['has_auto_filter'] = str_contains($xml, '<autoFilter');
        $base['has_drawing'] = str_contains($xml, '<drawing');

        return $base;
    }

    /** @param list<string> $warnings @param list<string> $errors */
    private function checkXmlEntry(string $realPath, string $entry, array &$warnings, array &$errors): void
    {
        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            return;
        }
        $exists = $zip->locateName($entry) !== false;
        $zip->close();
        if (!$exists) {
            return;
        }

        $reader = new XMLReader();
        $uri = 'zip://' . $realPath . '#' . $entry;
        if (!$reader->open($uri, null, LIBXML_NONET)) {
            $errors[] = 'Invalid XML or unable to open XML part: ' . $entry;
            return;
        }

        try {
            while ($reader->read()) {
                // Reading fully is enough to trigger libxml parse errors.
            }
        } catch (\Throwable $e) {
            $errors[] = 'Invalid XML in ' . $entry . ': ' . $e->getMessage();
        } finally {
            $reader->close();
        }
    }

    private function xmlReaderAvailable(): bool
    {
        return class_exists(XMLReader::class);
    }
}
