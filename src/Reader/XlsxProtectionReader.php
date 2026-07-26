<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class XlsxProtectionReader
{
    public function __construct(private readonly XlsxWorkbookResolver $resolver = new XlsxWorkbookResolver())
    {
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function read(string $path, int|string $sheet = 1, array $options = []): array
    {
        $temporary = null;
        $encryption = new XlsxEncryption();
        if ($encryption->isEncryptedFile($path)) {
            $password = (string) ($options['password'] ?? '');
            if ($password === '') {
                throw new MnbExcelException('A password is required to inspect protection in this encrypted XLSX file.');
            }
            $temporary = $encryption->decryptToTemporary($path, $password, $options);
            $path = $temporary;
        }

        try {
            $real = realpath($path);
            if ($real === false) {
                throw new MnbExcelException('Invalid XLSX path: ' . $path);
            }
            $zip = new ZipArchive();
            if ($zip->open($real) !== true) {
                throw new MnbExcelException('Unable to open XLSX package.');
            }
            try {
                $workbookXml = (string) ($zip->getFromName('xl/workbook.xml') ?: '');
                $sheetPath = $this->resolver->resolveSheetPath($real, $sheet);
                $sheetXml = (string) ($zip->getFromName($sheetPath) ?: '');
            } finally {
                $zip->close();
            }

            $workbook = $this->tagAttributes($workbookXml, 'workbookProtection');
            $worksheet = $this->tagAttributes($sheetXml, 'sheetProtection');
            return [
                'file_encrypted' => $temporary !== null || $encryption->isEncryptedFile($path),
                'workbook_protected' => $workbook !== [],
                'worksheet_protected' => $worksheet !== [],
                'workbook' => $this->sanitize($workbook),
                'worksheet' => $this->sanitize($worksheet),
            ];
        } finally {
            if ($temporary !== null) {
                @unlink($temporary);
            }
        }
    }

    /** @return array<string,string> */
    private function tagAttributes(string $xml, string $tag): array
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b([^>]*)\/?\s*>/is', $xml, $match) !== 1) {
            return [];
        }
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $match[1], $matches, PREG_SET_ORDER);
        foreach ($matches as $item) {
            $name = str_contains($item[1], ':') ? substr($item[1], strrpos($item[1], ':') + 1) : $item[1];
            $attrs[$name] = html_entity_decode($item[3] !== '' ? $item[3] : $item[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $attrs;
    }

    /** @param array<string,string> $attributes @return array<string,mixed> */
    private function sanitize(array $attributes): array
    {
        $result = [];
        foreach ($attributes as $name => $value) {
            if (str_contains(strtolower($name), 'hash') || str_contains(strtolower($name), 'salt') || str_contains(strtolower($name), 'password')) {
                $result[$name] = '[present]';
                continue;
            }
            $result[$name] = in_array($value, ['0', '1', 'true', 'false'], true)
                ? in_array($value, ['1', 'true'], true)
                : $value;
        }
        return $result;
    }
}
