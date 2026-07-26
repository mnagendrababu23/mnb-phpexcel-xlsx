<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\Xml\XmlReader;

use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class XlsxIntegrityValidator
{
    /** @var list<string> */
    private array $errors = [];

    /** @var list<string> */
    private array $warnings = [];

    /** @var list<array{name:string,status:string,message:string}> */
    private array $checks = [];

    /** @var array<string,true> */
    private array $entries = [];

    /**
     * Validate an XLSX package for the corruption causes that most often trigger
     * Excel's "We found a problem with some content" repair dialog.
     *
     * @param array<string,mixed> $options
     * @return array{status:string,valid:bool,errors:list<string>,warnings:list<string>,checks:list<array{name:string,status:string,message:string>>,summary:array{passed:int,warning:int,failed:int},path:string}
     */
    public function validate(string $path, array $options = []): array
    {
        $this->errors = [];
        $this->warnings = [];
        $this->checks = [];
        $this->entries = [];

        $path = (string) $path;
        $originalPath = $path;
        $temporary = null;
        if ($path === '' || !is_file($path)) {
            $this->fail('file_exists', 'XLSX file not found: ' . $path);
            return $this->result($path);
        }

        $encryption = new XlsxEncryption();
        if ($encryption->isEncryptedFile($path)) {
            $password = (string) ($options['password'] ?? $options['xlsx_password'] ?? '');
            if ($password === '') {
                $this->fail('password_required', 'A password is required to validate this encrypted XLSX file.');
                return $this->result($originalPath);
            }
            try {
                $temporary = $encryption->decryptToTemporary($path, $password, $options);
                $path = $temporary;
                $this->pass('encrypted_package', 'Encrypted XLSX package was opened for integrity validation.');
            } catch (\Throwable $e) {
                $this->fail('encrypted_package', 'Encrypted XLSX package could not be opened with the supplied password.');
                return $this->result($originalPath);
            }
        }

        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            $this->fail('zip_open', 'Unable to open XLSX as a ZIP package: ' . $originalPath);
            if ($temporary !== null) { @unlink($temporary); }
            return $this->result($originalPath);
        }

        try {
            $this->indexEntries($zip);
            $this->checkRequiredPackageParts($zip);
            $contentTypes = $this->checkContentTypes($zip);
            $relationships = $this->checkRelationshipTargets($zip);
            $this->checkXmlWellFormed($zip, $options);
            $this->checkWorkbookRelationshipIds($zip, $relationships);
            $this->checkWorksheetRelationshipIds($zip, $relationships);
            $this->checkContentTypeCoverage($contentTypes);
        } finally {
            $zip->close();
        }

        $result = $this->result($originalPath);
        if ($temporary !== null) { @unlink($temporary); }
        return $result;
    }

    /**
     * Throw when validation fails. Intended for save-time protection.
     *
     * @param array<string,mixed> $options
     */
    public function assertValid(string $path, array $options = []): void
    {
        $result = $this->validate($path, $options);
        if (($result['status'] ?? 'fail') === 'fail') {
            $message = 'XLSX integrity validation failed.';
            if (($result['errors'] ?? []) !== []) {
                $message .= ' ' . implode(' ', array_slice($result['errors'], 0, 5));
            }
            throw MnbExcelException::withCode($message, ErrorCode::XLSX_INTEGRITY_FAILED, ['path' => $path, 'errors' => $result['errors'] ?? []]);
        }
    }

    private function indexEntries(ZipArchive $zip): void
    {
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $entry = $zip->getNameIndex($i);
            if (!is_string($entry) || $entry === '' || str_ends_with($entry, '/')) {
                continue;
            }
            $this->entries[$this->normalizeEntryName($entry)] = true;
        }

        if ($this->entries === []) {
            $this->fail('zip_entries', 'XLSX ZIP package has no file entries.');
            return;
        }

        $this->pass('zip_entries', 'XLSX ZIP package contains ' . count($this->entries) . ' file entries.');
    }

    private function checkRequiredPackageParts(ZipArchive $zip): void
    {
        $required = [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/workbook.xml',
            'xl/_rels/workbook.xml.rels',
            'xl/styles.xml',
        ];

        foreach ($required as $entry) {
            if (!$this->hasEntry($entry)) {
                $this->fail('required_part', 'Missing required XLSX package part: ' . $entry);
            }
        }

        $worksheetEntries = $this->entriesMatching('/^xl\/worksheets\/sheet\d+\.xml$/i');
        if ($worksheetEntries === []) {
            $this->fail('worksheet_parts', 'Workbook has no worksheet XML parts.');
        } else {
            $this->pass('worksheet_parts', 'Workbook has ' . count($worksheetEntries) . ' worksheet XML part(s).');
        }

        if ($this->hasEntry('docProps/core.xml') && $this->hasEntry('docProps/app.xml')) {
            $this->pass('document_properties', 'Core and app document properties are present.');
        } else {
            $this->warn('document_properties', 'Document property parts are optional but one or more are missing.');
        }
    }

    /** @return array{defaults:array<string,string>,overrides:array<string,string>} */
    private function checkContentTypes(ZipArchive $zip): array
    {
        $xml = $this->readEntry($zip, '[Content_Types].xml');
        if ($xml === null) {
            return ['defaults' => [], 'overrides' => []];
        }

        $defaults = [];
        preg_match_all('/<Default\b[^>]*\/?>/iu', $xml, $defaultMatches);
        foreach ($defaultMatches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            if (isset($attrs['Extension'], $attrs['ContentType'])) {
                $defaults[strtolower((string) $attrs['Extension'])] = (string) $attrs['ContentType'];
            }
        }

        $overrides = [];
        preg_match_all('/<Override\b[^>]*\/?>/iu', $xml, $overrideMatches);
        foreach ($overrideMatches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            if (isset($attrs['PartName'], $attrs['ContentType'])) {
                $partName = '/' . ltrim((string) $attrs['PartName'], '/');
                $overrides[$partName] = (string) $attrs['ContentType'];
            }
        }

        if (($defaults['rels'] ?? '') !== 'application/vnd.openxmlformats-package.relationships+xml') {
            $this->fail('content_type_rels', 'Content types must declare the .rels package relationship content type.');
        }
        if (!isset($defaults['xml'])) {
            $this->warn('content_type_xml_default', 'Content types do not include a default XML mapping. Explicit overrides may still be valid.');
        }

        foreach (['/xl/workbook.xml', '/xl/styles.xml'] as $requiredOverride) {
            if (!isset($overrides[$requiredOverride])) {
                $this->fail('content_type_required_override', 'Missing content type override for ' . $requiredOverride);
            }
        }

        foreach ($this->entriesMatching('/^xl\/worksheets\/sheet\d+\.xml$/i') as $sheetEntry) {
            if (!isset($overrides['/' . $sheetEntry]) && !isset($defaults['xml'])) {
                $this->fail('content_type_worksheet', 'Missing content type for worksheet part: /' . $sheetEntry);
            }
        }

        $this->pass('content_types_parse', 'Parsed ' . count($defaults) . ' default and ' . count($overrides) . ' override content type declarations.');
        return ['defaults' => $defaults, 'overrides' => $overrides];
    }

    /**
     * @return array<string,array<string,array{target:string,mode:string,type:string}>>
     */
    private function checkRelationshipTargets(ZipArchive $zip): array
    {
        $relationships = [];
        foreach ($this->entriesMatching('/(^|\/)\.rels$|\/_rels\/[^\/]+\.rels$/i') as $relsEntry) {
            $xml = $this->readEntry($zip, $relsEntry);
            if ($xml === null) {
                continue;
            }

            $relationshipRows = $this->parseRelationships($xml);
            $relationships[$relsEntry] = $relationshipRows;
            foreach ($relationshipRows as $id => $rel) {
                $target = (string) $rel['target'];
                $mode = strtolower((string) $rel['mode']);
                if ($target === '') {
                    $this->fail('relationship_target', 'Relationship ' . $id . ' in ' . $relsEntry . ' has an empty Target.');
                    continue;
                }

                if ($mode === 'external' || $this->looksExternalTarget($target)) {
                    continue;
                }

                $resolved = $this->resolveRelationshipTarget($relsEntry, $target);
                if ($resolved === '' || !$this->hasEntry($resolved)) {
                    $this->fail('relationship_target', 'Relationship target missing: ' . $relsEntry . ' #' . $id . ' -> ' . $target . ' resolved as ' . ($resolved !== '' ? $resolved : '[invalid]'));
                }
            }
        }

        if ($relationships === []) {
            $this->fail('relationships', 'No XLSX relationship files were found.');
        } else {
            $this->pass('relationships', 'Checked relationship targets in ' . count($relationships) . ' .rels file(s).');
        }

        return $relationships;
    }

    /** @param array<string,mixed> $options */
    private function checkXmlWellFormed(ZipArchive $zip, array $options): void
    {
        $xmlEntries = $this->entriesMatching('/\.(xml|rels)$/i');
        if ($xmlEntries === []) {
            $this->fail('xml_parts', 'No XML package parts found.');
            return;
        }

        $checked = 0;
        $xmlReaderAvailable = class_exists('XMLReader');
        $requireXmlReader = (bool) ($options['require_xmlreader'] ?? false);

        if (!$xmlReaderAvailable && $requireXmlReader) {
            $this->fail('xmlreader_extension', 'ext-xmlreader is required for strict XML well-formed validation.');
            return;
        }

        if (!$xmlReaderAvailable) {
            $this->warn('xml_well_formed', 'ext-xmlreader is not available; using basic XML structure checks only.');
        }

        foreach ($xmlEntries as $entry) {
            $xml = $this->readEntry($zip, $entry);
            if ($xml === null) {
                continue;
            }

            if ($xmlReaderAvailable) {
                $error = $this->strictXmlError($xml);
                if ($error !== null) {
                    $this->fail('xml_well_formed', 'Malformed XML in ' . $entry . ': ' . $error);
                }
            } elseif (!$this->basicXmlLooksWellFormed($xml)) {
                $this->fail('xml_well_formed', 'Malformed or suspicious XML structure in ' . $entry . '.');
            }
            $checked++;
        }

        if ($checked > 0) {
            $this->pass('xml_well_formed', 'Checked XML structure for ' . $checked . ' package part(s).');
        }
    }

    /**
     * @param array<string,array<string,array{target:string,mode:string,type:string}>> $relationships
     */
    private function checkWorkbookRelationshipIds(ZipArchive $zip, array $relationships): void
    {
        $workbookXml = $this->readEntry($zip, 'xl/workbook.xml');
        if ($workbookXml === null) {
            return;
        }

        $ids = $this->relationshipIdsUsedByXml($workbookXml);
        $rels = $relationships['xl/_rels/workbook.xml.rels'] ?? [];
        if ($ids !== [] && $rels === []) {
            $this->fail('workbook_relationship_ids', 'xl/workbook.xml uses relationship IDs but xl/_rels/workbook.xml.rels is missing or empty.');
            return;
        }

        foreach ($ids as $id) {
            if (!isset($rels[$id])) {
                $this->fail('workbook_relationship_ids', 'xl/workbook.xml references missing relationship ID: ' . $id);
            }
        }

        $this->pass('workbook_relationship_ids', 'Workbook relationship IDs are declared in workbook.xml.rels.');
    }

    /**
     * @param array<string,array<string,array{target:string,mode:string,type:string}>> $relationships
     */
    private function checkWorksheetRelationshipIds(ZipArchive $zip, array $relationships): void
    {
        $checked = 0;
        foreach ($this->entriesMatching('/^xl\/worksheets\/sheet\d+\.xml$/i') as $sheetEntry) {
            $xml = $this->readEntry($zip, $sheetEntry);
            if ($xml === null) {
                continue;
            }

            $ids = $this->relationshipIdsUsedByXml($xml);
            $relsEntry = $this->relsPathForPart($sheetEntry);
            $rels = $relationships[$relsEntry] ?? [];

            if ($ids !== [] && $rels === []) {
                $this->fail('worksheet_relationship_ids', $sheetEntry . ' uses relationship IDs but ' . $relsEntry . ' is missing or empty.');
                continue;
            }

            foreach ($ids as $id) {
                if (!isset($rels[$id])) {
                    $this->fail('worksheet_relationship_ids', $sheetEntry . ' references missing relationship ID: ' . $id);
                }
            }
            $checked++;
        }

        if ($checked > 0) {
            $this->pass('worksheet_relationship_ids', 'Worksheet relationship IDs were checked for ' . $checked . ' worksheet(s).');
        }
    }

    /** @param array{defaults:array<string,string>,overrides:array<string,string>} $contentTypes */
    private function checkContentTypeCoverage(array $contentTypes): void
    {
        $defaults = $contentTypes['defaults'] ?? [];
        $overrides = $contentTypes['overrides'] ?? [];

        if ($defaults === [] && $overrides === []) {
            return;
        }

        foreach ($overrides as $partName => $_contentType) {
            $entry = ltrim((string) $partName, '/');
            if (!$this->hasEntry($entry)) {
                $this->fail('content_type_orphan_override', 'Content type override points to a missing package part: ' . $partName);
            }
        }

        foreach (array_keys($this->entries) as $entry) {
            if (str_ends_with(strtolower($entry), '.rels')) {
                continue;
            }
            $partName = '/' . $entry;
            $extension = strtolower(pathinfo($entry, PATHINFO_EXTENSION));
            if (!isset($overrides[$partName]) && ($extension === '' || !isset($defaults[$extension]))) {
                $this->fail('content_type_coverage', 'Package part has no matching content type declaration: ' . $partName);
            }
        }

        $this->pass('content_type_coverage', 'Checked content type coverage for package entries.');
    }

    /** @return array<string,array{target:string,mode:string,type:string}> */
    private function parseRelationships(string $xml): array
    {
        $relationships = [];
        preg_match_all('/<Relationship\b[^>]*\/?>/iu', $xml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            if (!isset($attrs['Id'])) {
                continue;
            }
            $relationships[(string) $attrs['Id']] = [
                'target' => (string) ($attrs['Target'] ?? ''),
                'mode' => (string) ($attrs['TargetMode'] ?? ''),
                'type' => (string) ($attrs['Type'] ?? ''),
            ];
        }
        return $relationships;
    }

    /** @return list<string> */
    private function relationshipIdsUsedByXml(string $xml): array
    {
        $ids = [];
        preg_match_all('/\br:id\s*=\s*("([^"]+)"|\'([^\']+)\')/iu', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $id = (string) (($match[2] ?? '') !== '' ? $match[2] : ($match[3] ?? ''));
            if ($id !== '') {
                $ids[$id] = true;
            }
        }
        return array_keys($ids);
    }

    private function strictXmlError(string $xml): ?string
    {
        $previous = libxml_use_internal_errors(true);
        libxml_clear_errors();

        $reader = new XmlReader();
        $opened = @$reader->XML($xml, null, LIBXML_NONET);
        if (!$opened) {
            $errors = libxml_get_errors();
            libxml_clear_errors();
            libxml_use_internal_errors($previous);
            return $this->formatLibxmlError($errors[0] ?? null);
        }

        try {
            while (@$reader->read()) {
                // Reading the full stream forces XMLReader to parse the complete document.
            }
        } catch (\Throwable $e) {
            $reader->close();
            libxml_clear_errors();
            libxml_use_internal_errors($previous);
            return $e->getMessage();
        }

        $errors = libxml_get_errors();
        $reader->close();
        libxml_clear_errors();
        libxml_use_internal_errors($previous);

        if ($errors !== []) {
            return $this->formatLibxmlError($errors[0]);
        }

        return null;
    }

    private function basicXmlLooksWellFormed(string $xml): bool
    {
        $xml = trim($this->stripUtf8Bom($xml));
        if ($xml === '' || !str_starts_with($xml, '<')) {
            return false;
        }

        if (preg_match('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', $xml) === 1) {
            return false;
        }

        $withoutSpecial = preg_replace('/<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>/su', '', $xml);
        if (!is_string($withoutSpecial)) {
            return false;
        }

        $stack = [];
        preg_match_all('/<\/?([A-Za-z_][A-Za-z0-9_.:\-]*)(?:\s[^<>]*)?>/u', $withoutSpecial, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $tag = $match[0];
            $name = $match[1];
            if (str_starts_with($tag, '<!') || str_starts_with($tag, '<?') || str_ends_with($tag, '/>')) {
                continue;
            }
            if (str_starts_with($tag, '</')) {
                if ($stack === [] || array_pop($stack) !== $name) {
                    return false;
                }
                continue;
            }
            $stack[] = $name;
        }

        return $stack === [];
    }

    private function formatLibxmlError(mixed $error): string
    {
        if (!is_object($error) || !isset($error->message)) {
            return 'Unknown XML parser error.';
        }

        $message = trim((string) $error->message);
        $line = isset($error->line) ? (int) $error->line : 0;
        $column = isset($error->column) ? (int) $error->column : 0;
        return $message . ($line > 0 ? ' at line ' . $line . ', column ' . $column : '');
    }

    /** @return array<string,string> */
    private function parseXmlAttributes(string $tag): array
    {
        $attrs = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $tag, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs[$match[1]] = html_entity_decode($match[3] !== '' ? $match[3] : $match[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $attrs;
    }

    private function resolveRelationshipTarget(string $relsEntry, string $target): string
    {
        $target = trim(str_replace('\\', '/', $target));
        if ($target === '' || str_starts_with($target, '#')) {
            return '';
        }

        $target = preg_replace('/[#?].*$/', '', $target) ?? $target;
        if ($target === '') {
            return '';
        }

        if (str_starts_with($target, '/')) {
            return $this->normalizeEntryName(ltrim($target, '/'));
        }

        $sourcePart = $this->sourcePartForRels($relsEntry);
        $baseDir = dirname($sourcePart);
        $combined = ($baseDir === '.' ? '' : $baseDir . '/') . $target;
        return $this->normalizeEntryName($combined);
    }

    private function sourcePartForRels(string $relsEntry): string
    {
        $relsEntry = $this->normalizeEntryName($relsEntry);
        if ($relsEntry === '_rels/.rels') {
            return '';
        }

        $pos = strrpos($relsEntry, '/_rels/');
        if ($pos === false) {
            return '';
        }

        $dir = substr($relsEntry, 0, $pos);
        $file = substr($relsEntry, $pos + strlen('/_rels/'));
        if (str_ends_with($file, '.rels')) {
            $file = substr($file, 0, -5);
        }
        return $dir . '/' . $file;
    }

    private function relsPathForPart(string $partPath): string
    {
        $partPath = $this->normalizeEntryName($partPath);
        $dir = dirname($partPath);
        $base = basename($partPath);
        return ($dir === '.' ? '_rels/' : $dir . '/_rels/') . $base . '.rels';
    }

    private function looksExternalTarget(string $target): bool
    {
        return preg_match('/^[a-z][a-z0-9+.-]*:/i', $target) === 1;
    }

    private function readEntry(ZipArchive $zip, string $entry): ?string
    {
        $entry = $this->normalizeEntryName($entry);
        if (!$this->hasEntry($entry)) {
            return null;
        }
        $content = $zip->getFromName($entry);
        return is_string($content) ? $content : null;
    }

    private function hasEntry(string $entry): bool
    {
        return isset($this->entries[$this->normalizeEntryName($entry)]);
    }

    /** @return list<string> */
    private function entriesMatching(string $pattern): array
    {
        $matches = [];
        foreach (array_keys($this->entries) as $entry) {
            if (preg_match($pattern, $entry) === 1) {
                $matches[] = $entry;
            }
        }
        sort($matches);
        return $matches;
    }

    private function normalizeEntryName(string $entry): string
    {
        $entry = str_replace('\\', '/', trim($entry));
        $entry = ltrim($entry, '/');
        $parts = [];
        foreach (explode('/', $entry) as $part) {
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

    private function stripUtf8Bom(string $value): string
    {
        return str_starts_with($value, "\xEF\xBB\xBF") ? substr($value, 3) : $value;
    }

    private function pass(string $name, string $message): void
    {
        $this->checks[] = ['name' => $name, 'status' => 'pass', 'message' => $message];
    }

    private function warn(string $name, string $message): void
    {
        $this->warnings[] = $message;
        $this->checks[] = ['name' => $name, 'status' => 'warning', 'message' => $message];
    }

    private function fail(string $name, string $message): void
    {
        $this->errors[] = $message;
        $this->checks[] = ['name' => $name, 'status' => 'fail', 'message' => $message];
    }

    /**
     * @return array{status:string,valid:bool,errors:list<string>,warnings:list<string>,checks:list<array{name:string,status:string,message:string}>,summary:array{passed:int,warning:int,failed:int},path:string}
     */
    private function result(string $path): array
    {
        $summary = ['passed' => 0, 'warning' => 0, 'failed' => 0];
        foreach ($this->checks as $check) {
            if ($check['status'] === 'pass') {
                $summary['passed']++;
            } elseif ($check['status'] === 'warning') {
                $summary['warning']++;
            } elseif ($check['status'] === 'fail') {
                $summary['failed']++;
            }
        }

        return [
            'status' => $this->errors === [] ? ($this->warnings === [] ? 'pass' : 'warning') : 'fail',
            'valid' => $this->errors === [],
            'errors' => array_values(array_unique($this->errors)),
            'warnings' => array_values(array_unique($this->warnings)),
            'checks' => $this->checks,
            'summary' => $summary,
            'path' => $path,
        ];
    }
}
