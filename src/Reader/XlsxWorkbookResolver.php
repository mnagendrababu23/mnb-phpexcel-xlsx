<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\SheetSelectionException;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class XlsxWorkbookResolver
{
    /** @return list<array{index:int,name:string,sheet_id:int,relationship_id:string,state:string,path:string,exists:bool}> */
    public function sheets(string $realPath): array
    {
        $zip = $this->openZip($realPath);
        $workbookXml = $zip->getFromName('xl/workbook.xml');
        if ($workbookXml === false || trim($workbookXml) === '') {
            $zip->close();
            throw new MnbExcelException('Invalid XLSX structure: xl/workbook.xml is missing.');
        }

        $relsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
        $relationshipTargets = $relsXml !== false ? $this->parseWorkbookRelationshipTargets($relsXml) : [];
        $sheets = $this->parseWorkbookSheets($workbookXml);

        foreach ($sheets as $i => $sheet) {
            $target = $relationshipTargets[$sheet['relationship_id']] ?? ('worksheets/sheet' . ($i + 1) . '.xml');
            $path = $this->normalizeWorkbookTarget($target);
            $sheets[$i]['path'] = $path;
            $sheets[$i]['exists'] = $zip->locateName($path) !== false;
        }

        $zip->close();
        return $sheets;
    }

    public function resolveSheetPath(string $realPath, int|string $sheet): string
    {
        if (is_string($sheet)) {
            $sheet = trim($sheet);
            if ($sheet === '') {
                throw SheetSelectionException::emptyName($realPath);
            }
        }
        if ((is_int($sheet) || ctype_digit((string) $sheet)) && (int) $sheet < 1) {
            throw SheetSelectionException::invalidIndex((int) $sheet, $realPath);
        }

        $sheets = $this->sheets($realPath);
        if ($sheets === []) {
            throw new MnbExcelException('Workbook does not contain any sheets.');
        }

        $selected = null;
        if (is_int($sheet) || ctype_digit((string) $sheet)) {
            $index = (int) $sheet;
            $selected = $sheets[$index - 1] ?? null;
        } else {
            $selected = $this->findSheetByName($sheets, (string) $sheet, $realPath);
        }

        if ($selected === null) {
            throw SheetSelectionException::notFound($sheet, $realPath, array_column($sheets, 'name'));
        }

        if (!$selected['exists']) {
            throw new MnbExcelException('Invalid XLSX structure: worksheet XML is missing for sheet "' . $selected['name'] . '" (' . $selected['path'] . ').');
        }

        return $selected['path'];
    }

    /** @return array<string, string> */
    private function parseWorkbookRelationshipTargets(string $xml): array
    {
        $targets = [];
        foreach ($this->matchTags($xml, 'Relationship') as $tag) {
            $attrs = $this->parseAttributes($tag);
            $type = $attrs['Type'] ?? '';
            if (!str_contains($type, '/worksheet')) {
                continue;
            }
            if (isset($attrs['Id'], $attrs['Target'])) {
                $targets[$attrs['Id']] = $attrs['Target'];
            }
        }

        return $targets;
    }

    /** @return list<array{index:int,name:string,sheet_id:int,relationship_id:string,state:string,path:string,exists:bool}> */
    private function parseWorkbookSheets(string $xml): array
    {
        $sheets = [];
        foreach ($this->matchTags($xml, 'sheet') as $tag) {
            $attrs = $this->parseAttributes($tag);
            $index = count($sheets) + 1;
            $sheets[] = [
                'index' => $index,
                'name' => $this->decode($attrs['name'] ?? ('Sheet' . $index)),
                'sheet_id' => isset($attrs['sheetId']) ? (int) $attrs['sheetId'] : $index,
                'relationship_id' => $this->attributeByLocalName($attrs, 'id') ?? ('rId' . $index),
                'state' => $attrs['state'] ?? 'visible',
                'path' => '',
                'exists' => false,
            ];
        }

        return $sheets;
    }

    /** @param list<array{index:int,name:string,sheet_id:int,relationship_id:string,state:string,path:string,exists:bool}> $sheets */
    private function findSheetByName(array $sheets, string $name, string $realPath): ?array
    {
        foreach ($sheets as $sheet) {
            if ($sheet['name'] === $name) {
                return $sheet;
            }
        }

        $needle = strtolower($name);
        $matches = [];
        foreach ($sheets as $sheet) {
            if (strtolower($sheet['name']) === $needle) {
                $matches[] = $sheet;
            }
        }

        if (count($matches) > 1) {
            throw SheetSelectionException::ambiguousName($name, $realPath, array_column($sheets, 'name'));
        }

        return $matches[0] ?? null;
    }

    private function normalizeWorkbookTarget(string $target): string
    {
        $target = str_replace('\\', '/', trim($target));
        if ($target === '') {
            return 'xl/worksheets/sheet1.xml';
        }

        if (str_starts_with($target, '/')) {
            $target = ltrim($target, '/');
            return $this->normalizePath($target);
        }

        if (!str_starts_with($target, 'xl/')) {
            $target = 'xl/' . $target;
        }

        return $this->normalizePath($target);
    }

    private function normalizePath(string $path): string
    {
        $parts = [];
        foreach (explode('/', $path) as $part) {
            if ($part === '' || $part === '.') {
                continue;
            }
            if ($part === '..') {
                array_pop($parts);
                continue;
            }
            $parts[] = $part;
        }

        return implode('/', $parts);
    }

    /** @return list<string> */
    private function matchTags(string $xml, string $tag): array
    {
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*>/i', $xml, $matches);
        return $matches[0] ?? [];
    }

    /** @return array<string, string> */
    private function parseAttributes(string $tag): array
    {
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $tag, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs[$match[1]] = $this->decode($match[3] !== '' ? $match[3] : $match[4]);
        }

        return $attrs;
    }

    /** @param array<string,string> $attributes */
    private function attributeByLocalName(array $attributes, string $localName): ?string
    {
        foreach ($attributes as $name => $value) {
            $separator = strrpos($name, ':');
            $candidate = $separator === false ? $name : substr($name, $separator + 1);
            if (strcasecmp($candidate, $localName) === 0) {
                return $value;
            }
        }

        return null;
    }

    private function decode(string $value): string
    {
        return html_entity_decode($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    private function openZip(string $realPath): ZipArchive
    {
        if (!class_exists(ZipArchive::class)) {
            throw new MnbExcelException('ext-zip is required to read XLSX files.');
        }

        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw new MnbExcelException('Unable to open XLSX zip package. The file may be corrupted or not a valid XLSX file.');
        }

        return $zip;
    }
}
