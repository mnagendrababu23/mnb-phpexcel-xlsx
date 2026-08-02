<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use DateTimeImmutable;
use DateTimeInterface;
use DateTimeZone;
use Mnb\PHPExcel\Metadata\MetadataWriterInterface;
use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/**
 * Atomic OOXML metadata updater.
 *
 * Unknown package parts are copied byte-for-byte. Only the metadata-related
 * XML parts requested by the caller are regenerated or patched.
 */
final class XlsxMetadataWriter implements MetadataWriterInterface
{
    /** @var list<string> */
    private const ALLOWED_SECTIONS = [
        'document',
        'revision',
        'application',
        'custom_properties',
        'workbook',
        'calculation',
    ];

    /** @var array<string,list<string>> */
    private const ALLOWED_FIELDS = [
        'document' => [
            'title', 'subject', 'creator', 'keywords', 'description', 'category',
            'content_status', 'identifier', 'language', 'document_version',
        ],
        'revision' => [
            'last_saved_by', 'revision_number', 'total_editing_time_seconds',
            'last_printed_at', 'document_created_at', 'document_modified_at',
        ],
        'application' => [
            'application', 'application_version', 'company', 'manager',
            'operating_system_hint', 'document_security', 'scale_crop',
            'links_up_to_date', 'shared_document', 'hyperlinks_changed',
        ],
        'custom_properties' => [],
        'workbook' => [
            'active_sheet', 'sheet_visibility', 'date1904', 'code_name',
            'read_only_recommended',
        ],
        'calculation' => [
            'mode', 'calc_id', 'full_calc_on_load', 'force_full_calc',
            'calc_on_save', 'iterate', 'iterate_count', 'iterate_delta',
        ],
    ];

    /** @param array<string,mixed> $changes @param array<string,mixed> $options */
    public function updateMetaInfo(string $source, string $destination, array $changes, array $options = []): void
    {
        $this->assertChanges($changes, (bool) ($options['strict'] ?? true));
        $sourcePath = realpath($source);
        if ($sourcePath === false || !is_file($sourcePath)) {
            throw new MnbExcelException('Invalid XLSX source path: ' . $source);
        }
        if (trim($destination) === '') {
            throw new MnbExcelException('XLSX metadata destination path cannot be empty.');
        }

        $encryption = new XlsxEncryption();
        $encrypted = $encryption->isEncryptedFile($sourcePath);
        $password = (string) ($options['password'] ?? $options['xlsx_password'] ?? '');
        if ($encrypted && $password === '') {
            throw new MnbExcelException('A password is required to update metadata in this encrypted XLSX file.');
        }

        $decrypted = null;
        $updatedPlain = null;
        try {
            $plainSource = $sourcePath;
            if ($encrypted) {
                $decrypted = $encryption->decryptToTemporary($sourcePath, $password, $options);
                $plainSource = $decrypted;
            }

            if ($encrypted) {
                $updatedPlain = $this->temporaryPath('mnb_xlsx_meta_');
                $this->writeUpdatedPackage($plainSource, $updatedPlain, $changes, $options);
                if ((bool) ($options['validate_package'] ?? true)) {
                    (new XlsxIntegrityValidator())->assertValid($updatedPlain);
                }
                $mode = $encryption->encryptionMode($sourcePath);
                if (!in_array($mode, ['agile', 'standard'], true)) {
                    $mode = (string) ($options['encryption_mode'] ?? 'agile');
                }
                $encryptionOptions = array_replace((array) ($options['encryption_options'] ?? []), ['mode' => $mode]);
                $encryption->encryptFile($updatedPlain, $destination, $password, $encryptionOptions);
                return;
            }

            AtomicFileWriter::writeViaTemp(
                $destination,
                function (string $temporary) use ($plainSource, $changes, $options): void {
                    $this->writeUpdatedPackage($plainSource, $temporary, $changes, $options);
                },
                (bool) ($options['validate_package'] ?? true)
                    ? static function (string $temporary): void {
                        (new XlsxIntegrityValidator())->assertValid($temporary);
                    }
                    : null
            );
        } finally {
            if ($decrypted !== null) {
                @unlink($decrypted);
            }
            if ($updatedPlain !== null) {
                @unlink($updatedPlain);
            }
        }
    }

    /** @param array<string,mixed> $options */
    public function removePersonalInfo(string $source, string $destination, array $options = []): void
    {
        $changes = [
            'document' => [
                'creator' => null,
            ],
            'revision' => [
                'last_saved_by' => null,
            ],
            'application' => [
                'manager' => null,
                'company' => null,
            ],
            'custom_properties' => [],
        ];
        $options = array_replace([
            'replace_custom_properties' => (bool) ($options['remove_custom_properties'] ?? true),
            'anonymize_comment_authors' => true,
            'anonymized_author_name' => 'Author',
        ], $options);

        if ((bool) ($options['remove_descriptive_properties'] ?? false)) {
            $changes['document'] += [
                'title' => null,
                'subject' => null,
                'keywords' => null,
                'description' => null,
                'category' => null,
            ];
        }

        $this->updateMetaInfo($source, $destination, $changes, $options + ['_remove_personal_info' => true]);
    }

    /** @param array<string,mixed> $changes @param array<string,mixed> $options */
    private function writeUpdatedPackage(string $source, string $destination, array $changes, array $options): void
    {
        $input = new ZipArchive();
        if ($input->open($source) !== true) {
            throw new MnbExcelException('Unable to open XLSX source package.');
        }

        $entries = [];
        try {
            for ($index = 0; $index < $input->numFiles; $index++) {
                $name = (string) ($input->getNameIndex($index) ?: '');
                if ($name === '') {
                    continue;
                }
                $bytes = $input->getFromName($name);
                if ($bytes === false) {
                    throw new MnbExcelException('Unable to read XLSX package part: ' . $name);
                }
                $entries[$name] = $bytes;
            }
        } finally {
            $input->close();
        }

        $core = $entries['docProps/core.xml'] ?? $this->newCorePropertiesXml();
        $app = $entries['docProps/app.xml'] ?? $this->newApplicationPropertiesXml();
        $workbook = $entries['xl/workbook.xml'] ?? null;
        if ($workbook === null || trim($workbook) === '') {
            throw new MnbExcelException('Invalid XLSX package: xl/workbook.xml is missing.');
        }

        if (isset($changes['document']) || isset($changes['revision'])) {
            $core = $this->updateCoreProperties(
                $core,
                (array) ($changes['document'] ?? []),
                (array) ($changes['revision'] ?? [])
            );
            $entries['docProps/core.xml'] = $core;
        }

        if (isset($changes['application']) || isset($changes['revision'])) {
            $app = $this->updateApplicationProperties(
                $app,
                (array) ($changes['application'] ?? []),
                (array) ($changes['revision'] ?? [])
            );
            $entries['docProps/app.xml'] = $app;
        }

        $existingCustomXml = (string) ($entries['docProps/custom.xml'] ?? '');
        $customExists = array_key_exists('docProps/custom.xml', $entries);
        if (array_key_exists('custom_properties', $changes)) {
            $customItems = $this->mergeCustomProperties(
                $existingCustomXml,
                $changes['custom_properties'],
                (bool) ($options['replace_custom_properties'] ?? false)
            );
            if ($customItems === []) {
                unset($entries['docProps/custom.xml']);
                $customExists = false;
            } else {
                $entries['docProps/custom.xml'] = $this->customPropertiesXml($customItems, $existingCustomXml);
                $customExists = true;
            }
        }

        if (isset($changes['workbook']) || isset($changes['calculation'])) {
            $workbook = $this->updateWorkbookXml(
                $workbook,
                (array) ($changes['workbook'] ?? []),
                (array) ($changes['calculation'] ?? [])
            );
            $entries['xl/workbook.xml'] = $workbook;
        }

        $coreExists = array_key_exists('docProps/core.xml', $entries);
        $applicationExists = array_key_exists('docProps/app.xml', $entries);
        $entries['[Content_Types].xml'] = $this->updateContentTypes(
            $entries['[Content_Types].xml'] ?? '',
            $coreExists,
            $applicationExists,
            $customExists
        );
        $entries['_rels/.rels'] = $this->updateRootRelationships(
            $entries['_rels/.rels'] ?? '',
            $coreExists,
            $applicationExists,
            $customExists
        );

        if ((bool) ($options['anonymize_comment_authors'] ?? false)) {
            $replacement = trim((string) ($options['anonymized_author_name'] ?? 'Author'));
            $replacement = $replacement !== '' ? $replacement : 'Author';
            foreach ($entries as $name => $bytes) {
                if (preg_match('#^xl/comments[^/]*\.xml$#i', $name) === 1) {
                    $entries[$name] = $this->anonymizeLegacyComments($bytes, $replacement);
                } elseif (preg_match('#^xl/persons/.*\.xml$#i', $name) === 1) {
                    $entries[$name] = $this->anonymizePersons($bytes, $replacement);
                }
            }
        }

        $output = new ZipArchive();
        if ($output->open($destination, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new MnbExcelException('Unable to create updated XLSX package.');
        }
        try {
            foreach ($entries as $name => $bytes) {
                if (!$output->addFromString($name, $bytes)) {
                    throw new MnbExcelException('Unable to write XLSX package part: ' . $name);
                }
            }
        } finally {
            $output->close();
        }
    }

    /** @param array<string,mixed> $document @param array<string,mixed> $revision */
    private function updateCoreProperties(string $xml, array $document, array $revision): string
    {
        [$xml, $corePrefix] = $this->ensureNamespace(
            $xml,
            'coreProperties',
            'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
            'cp'
        );
        [$xml, $dcPrefix] = $this->ensureNamespace($xml, 'coreProperties', 'http://purl.org/dc/elements/1.1/', 'dc');
        [$xml, $termsPrefix] = $this->ensureNamespace($xml, 'coreProperties', 'http://purl.org/dc/terms/', 'dcterms');
        [$xml, $xsiPrefix] = $this->ensureNamespace($xml, 'coreProperties', 'http://www.w3.org/2001/XMLSchema-instance', 'xsi');

        $documentMap = [
            'title' => [$this->qualifiedName($dcPrefix, 'title'), 'title', ''],
            'subject' => [$this->qualifiedName($dcPrefix, 'subject'), 'subject', ''],
            'creator' => [$this->qualifiedName($dcPrefix, 'creator'), 'creator', ''],
            'keywords' => [$this->qualifiedName($corePrefix, 'keywords'), 'keywords', ''],
            'description' => [$this->qualifiedName($dcPrefix, 'description'), 'description', ''],
            'category' => [$this->qualifiedName($corePrefix, 'category'), 'category', ''],
            'content_status' => [$this->qualifiedName($corePrefix, 'contentStatus'), 'contentStatus', ''],
            'identifier' => [$this->qualifiedName($dcPrefix, 'identifier'), 'identifier', ''],
            'language' => [$this->qualifiedName($dcPrefix, 'language'), 'language', ''],
            'document_version' => [$this->qualifiedName($corePrefix, 'version'), 'version', ''],
        ];
        foreach ($documentMap as $key => [$qualified, $local, $attrs]) {
            if (array_key_exists($key, $document)) {
                $value = $document[$key];
                if ($key === 'keywords' && is_array($value)) {
                    $value = implode(', ', array_map('strval', $value));
                }
                $xml = $this->setElement($xml, $qualified, $local, $value, $attrs, 'coreProperties');
            }
        }

        $dateAttributes = ' ' . $this->qualifiedName($xsiPrefix, 'type') . '="'
            . $this->qualifiedName($termsPrefix, 'W3CDTF') . '"';
        $revisionMap = [
            'last_saved_by' => [$this->qualifiedName($corePrefix, 'lastModifiedBy'), 'lastModifiedBy', ''],
            'revision_number' => [$this->qualifiedName($corePrefix, 'revision'), 'revision', ''],
            'last_printed_at' => [$this->qualifiedName($corePrefix, 'lastPrinted'), 'lastPrinted', ''],
            'document_created_at' => [$this->qualifiedName($termsPrefix, 'created'), 'created', $dateAttributes],
            'document_modified_at' => [$this->qualifiedName($termsPrefix, 'modified'), 'modified', $dateAttributes],
        ];
        foreach ($revisionMap as $key => [$qualified, $local, $attrs]) {
            if (!array_key_exists($key, $revision)) {
                continue;
            }
            $value = $revision[$key];
            if (in_array($key, ['last_printed_at', 'document_created_at', 'document_modified_at'], true) && $value !== null && $value !== '') {
                $value = $this->dateValue($value);
            }
            $xml = $this->setElement($xml, $qualified, $local, $value, $attrs, 'coreProperties');
        }

        return $xml;
    }

    /** @param array<string,mixed> $application @param array<string,mixed> $revision */
    private function updateApplicationProperties(string $xml, array $application, array $revision): string
    {
        [$xml, $appPrefix] = $this->ensureNamespace(
            $xml,
            'Properties',
            'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
            'ep'
        );

        $map = [
            'application' => 'Application',
            'application_version' => 'AppVersion',
            'company' => 'Company',
            'manager' => 'Manager',
            'operating_system_hint' => 'OperatingSystem',
            'document_security' => 'DocSecurity',
            'scale_crop' => 'ScaleCrop',
            'links_up_to_date' => 'LinksUpToDate',
            'shared_document' => 'SharedDoc',
            'hyperlinks_changed' => 'HyperlinksChanged',
        ];
        foreach ($map as $key => $local) {
            if (!array_key_exists($key, $application)) {
                continue;
            }
            $value = $application[$key];
            if ($value !== null) {
                if (in_array($key, ['scale_crop', 'links_up_to_date', 'shared_document', 'hyperlinks_changed'], true)) {
                    $value = $this->booleanValue($value) ? 'true' : 'false';
                } elseif ($key === 'document_security') {
                    $value = $this->integerText($value);
                } else {
                    $value = $this->stringValue($value);
                }
            }
            $xml = $this->setElement($xml, $this->qualifiedName($appPrefix, $local), $local, $value, '', 'Properties');
        }

        if (array_key_exists('total_editing_time_seconds', $revision)) {
            $seconds = $revision['total_editing_time_seconds'];
            $minutes = $seconds === null ? null : (int) ceil(max(0, (int) $seconds) / 60);
            $xml = $this->setElement(
                $xml,
                $this->qualifiedName($appPrefix, 'TotalTime'),
                'TotalTime',
                $minutes,
                '',
                'Properties'
            );
        }

        return $xml;
    }

    /** @param array<string,mixed> $workbook @param array<string,mixed> $calculation */
    private function updateWorkbookXml(string $xml, array $workbook, array $calculation): string
    {
        $sheets = $this->workbookSheets($xml);

        if (array_key_exists('sheet_visibility', $workbook)) {
            $visibility = (array) $workbook['sheet_visibility'];
            foreach ($visibility as $selector => $stateValue) {
                $state = (string) $stateValue;
                if (!in_array($state, ['visible', 'hidden', 'veryHidden'], true)) {
                    throw new MnbExcelException('Worksheet visibility must be visible, hidden, or veryHidden.');
                }
                $index = $this->resolveSheetIndex($sheets, is_int($selector) ? $selector : (string) $selector);
                $xml = $this->setSheetState($xml, $sheets[$index - 1]['name'], $state);
                $sheets[$index - 1]['state'] = $state;
            }
            if (count(array_filter($sheets, static fn(array $sheet): bool => $sheet['state'] === 'visible')) === 0) {
                throw new MnbExcelException('At least one worksheet must remain visible.');
            }
        }

        if (array_key_exists('active_sheet', $workbook)) {
            $activeIndex = $this->resolveSheetIndex($sheets, $workbook['active_sheet']);
            if ($sheets[$activeIndex - 1]['state'] !== 'visible') {
                throw new MnbExcelException('The active worksheet must be visible.');
            }
            $xml = $this->setWorkbookViewAttribute($xml, 'activeTab', (string) ($activeIndex - 1));
            $xml = $this->setWorkbookViewAttribute($xml, 'firstSheet', (string) ($activeIndex - 1));
        } elseif (array_key_exists('sheet_visibility', $workbook)) {
            $currentActive = $this->activeSheetIndexFromXml($xml, count($sheets));
            if ($sheets[$currentActive - 1]['state'] !== 'visible') {
                foreach ($sheets as $index => $sheet) {
                    if ($sheet['state'] === 'visible') {
                        $xml = $this->setWorkbookViewAttribute($xml, 'activeTab', (string) $index);
                        $xml = $this->setWorkbookViewAttribute($xml, 'firstSheet', (string) $index);
                        break;
                    }
                }
            }
        }

        if (array_key_exists('date1904', $workbook)) {
            $xml = $this->setSingletonAttribute($xml, 'workbookPr', 'date1904', $this->booleanValue($workbook['date1904']) ? '1' : '0', 'sheets');
        }
        if (array_key_exists('code_name', $workbook)) {
            $value = $workbook['code_name'];
            $xml = $value === null || $value === ''
                ? $this->removeSingletonAttribute($xml, 'workbookPr', 'codeName')
                : $this->setSingletonAttribute($xml, 'workbookPr', 'codeName', (string) $value, 'sheets');
        }
        if (array_key_exists('read_only_recommended', $workbook)) {
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?fileSharing\b[^>]*\/?\s*>/i', $xml) === 1) {
                $xml = $this->setAttributeOnFirstTag($xml, 'fileSharing', 'readOnlyRecommended', $this->booleanValue($workbook['read_only_recommended']) ? '1' : '0');
            } else {
                $qualified = $this->qualifiedName(
                    $this->namespacePrefix($xml, 'workbook', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),
                    'fileSharing'
                );
                $element = '<' . $qualified . ' readOnlyRecommended="' . ($this->booleanValue($workbook['read_only_recommended']) ? '1' : '0') . '"/>';
                $xml = $this->insertBeforeFirstAvailableTag($xml, ['workbookPr', 'workbookProtection', 'bookViews', 'sheets'], $element);
            }
        }

        $calcMap = [
            'mode' => 'calcMode',
            'calc_id' => 'calcId',
            'full_calc_on_load' => 'fullCalcOnLoad',
            'force_full_calc' => 'forceFullCalc',
            'calc_on_save' => 'calcOnSave',
            'iterate' => 'iterate',
            'iterate_count' => 'iterateCount',
            'iterate_delta' => 'iterateDelta',
        ];
        foreach ($calcMap as $key => $attribute) {
            if (!array_key_exists($key, $calculation)) {
                continue;
            }
            $value = $calculation[$key];
            if ($key === 'mode') {
                if (!in_array($value, ['auto', 'manual', 'autoNoTable'], true)) {
                    throw new MnbExcelException('Calculation mode must be auto, manual, or autoNoTable.');
                }
                $value = (string) $value;
            } elseif (in_array($key, ['full_calc_on_load', 'force_full_calc', 'calc_on_save', 'iterate'], true)) {
                $value = $this->booleanValue($value) ? '1' : '0';
            } elseif (in_array($key, ['calc_id', 'iterate_count'], true)) {
                $value = $this->integerText($value);
            } elseif ($key === 'iterate_delta') {
                $value = $this->number($this->floatValue($value));
            }
            $xml = $this->setSingletonAttribute($xml, 'calcPr', $attribute, (string) $value, 'workbook');
        }

        return $xml;
    }

    /** @param mixed $changes @return list<array<string,mixed>> */
    private function mergeCustomProperties(string $existingXml, mixed $changes, bool $replace): array
    {
        $properties = $replace ? [] : $this->parseCustomProperties($existingXml);
        $indexed = [];
        foreach ($properties as $property) {
            $indexed[(string) $property['name']] = $property;
        }

        foreach ($this->normalizeCustomChanges($changes) as $change) {
            $name = trim((string) ($change['name'] ?? ''));
            if ($name === '') {
                throw new MnbExcelException('Custom property name cannot be empty.');
            }
            if (!array_key_exists('value', $change) || $change['value'] === null) {
                unset($indexed[$name]);
                continue;
            }
            $type = strtolower((string) ($change['type'] ?? $this->inferCustomType($change['value'])));
            if (!in_array($type, ['string', 'integer', 'float', 'boolean', 'datetime'], true)) {
                throw new MnbExcelException('Unsupported custom property type: ' . $type);
            }
            $existing = $indexed[$name] ?? [];
            $indexed[$name] = [
                'name' => $name,
                'type' => $type,
                'value' => $change['value'],
                'linked' => (bool) ($change['linked'] ?? ($existing['linked'] ?? false)),
                'link_target' => $change['link_target'] ?? ($existing['link_target'] ?? null),
                'property_id' => $existing['property_id'] ?? null,
                'fmtid' => $existing['fmtid'] ?? null,
            ];
        }

        return array_values($indexed);
    }

    /** @return list<array<string,mixed>> */
    private function parseCustomProperties(string $xml): array
    {
        $items = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?property\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?property\s*>/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?([A-Za-z0-9_]+)\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?\1\s*>/is', $match[2], $valueMatch) !== 1) {
                $items[] = [
                    'name' => (string) ($attrs['name'] ?? ''),
                    'type' => 'opaque',
                    'value' => null,
                    'linked' => isset($attrs['linkTarget']),
                    'link_target' => $attrs['linkTarget'] ?? null,
                    'property_id' => isset($attrs['pid']) ? (int) $attrs['pid'] : null,
                    'fmtid' => $attrs['fmtid'] ?? null,
                    'raw_xml' => $match[0],
                ];
                continue;
            }
            $native = strtolower($valueMatch[1]);
            $raw = $this->decodeText($valueMatch[2]);
            [$type, $value] = match ($native) {
                'i1', 'i2', 'i4', 'i8', 'int', 'ui1', 'ui2', 'ui4', 'ui8', 'uint' => ['integer', $this->integerMetadataValue($raw)],
                'r4', 'r8', 'decimal', 'cy' => ['float', (float) $raw],
                'bool' => ['boolean', in_array(strtolower($raw), ['1', 'true'], true)],
                'filetime', 'date' => ['datetime', $raw],
                'lpstr', 'lpwstr', 'bstr' => ['string', $raw],
                default => ['opaque', null],
            };
            $items[] = [
                'name' => (string) ($attrs['name'] ?? ''),
                'type' => $type,
                'value' => $value,
                'native_type' => $native,
                'linked' => isset($attrs['linkTarget']),
                'link_target' => $attrs['linkTarget'] ?? null,
                'property_id' => isset($attrs['pid']) ? (int) $attrs['pid'] : null,
                'fmtid' => $attrs['fmtid'] ?? null,
                'raw_xml' => $match[0],
            ];
        }
        return $items;
    }

    /** @return list<array<string,mixed>> */
    private function normalizeCustomChanges(mixed $changes): array
    {
        if (!is_array($changes)) {
            throw new MnbExcelException('custom_properties changes must be an array.');
        }
        $items = [];
        foreach ($changes as $key => $value) {
            if (is_int($key)) {
                if (!is_array($value)) {
                    throw new MnbExcelException('Each custom property list item must be an array.');
                }
                $items[] = $value;
                continue;
            }
            if (is_array($value) && (array_key_exists('value', $value) || array_key_exists('type', $value))) {
                $items[] = ['name' => (string) $key] + $value;
            } else {
                $items[] = ['name' => (string) $key, 'value' => $value];
            }
        }
        return $items;
    }

    /** @param list<array<string,mixed>> $properties */
    private function customPropertiesXml(array $properties, string $existingXml = ''): string
    {
        if (trim($existingXml) === '') {
            $head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
                . 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
            $tail = '</Properties>';
            $propertyPrefix = '';
            $valuePrefix = 'vt';
        } else {
            [$existingXml, $propertyPrefix] = $this->ensureNamespace(
                $existingXml,
                'Properties',
                'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties',
                'cp'
            );
            [$existingXml, $valuePrefix] = $this->ensureNamespace(
                $existingXml,
                'Properties',
                'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
                'vt'
            );
            if (preg_match('/^(.*?<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?Properties)\b[^>]*>).*?(<\/\2\s*>\s*)$/is', $existingXml, $root) !== 1) {
                throw new MnbExcelException('Invalid custom-properties XML root.');
            }
            $head = $root[1];
            $tail = $root[3];
        }

        $usedPids = [];
        $nextPid = 2;
        foreach ($properties as $property) {
            $pid = (int) ($property['property_id'] ?? 0);
            if ($pid >= 2 && !isset($usedPids[$pid])) {
                $usedPids[$pid] = true;
                $nextPid = max($nextPid, $pid + 1);
            }
        }

        $body = '';
        foreach ($properties as $property) {
            if (isset($property['raw_xml']) && is_string($property['raw_xml']) && $property['raw_xml'] !== '') {
                $body .= $property['raw_xml'];
                continue;
            }

            $pid = (int) ($property['property_id'] ?? 0);
            if ($pid < 2 || isset($usedPids[$pid]) && $pid < $nextPid && count(array_filter($properties, static fn(array $item): bool => (int) ($item['property_id'] ?? 0) === $pid)) > 1) {
                while (isset($usedPids[$nextPid])) {
                    $nextPid++;
                }
                $pid = $nextPid++;
                $usedPids[$pid] = true;
            }

            $name = (string) $property['name'];
            $type = (string) $property['type'];
            $value = $property['value'];
            $link = (bool) ($property['linked'] ?? false) && (string) ($property['link_target'] ?? '') !== ''
                ? ' linkTarget="' . $this->esc((string) $property['link_target']) . '"'
                : '';
            $fmtid = trim((string) ($property['fmtid'] ?? ''));
            if ($fmtid === '') {
                $fmtid = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
            }
            $propertyTag = $this->qualifiedName($propertyPrefix, 'property');
            $body .= '<' . $propertyTag . ' fmtid="' . $this->esc($fmtid) . '" pid="' . $pid . '" name="' . $this->esc($name) . '"' . $link . '>';
            $body .= match ($type) {
                'integer' => $this->customIntegerXml($value, $valuePrefix),
                'float' => '<' . $this->qualifiedName($valuePrefix, 'r8') . '>' . $this->number($this->floatValue($value)) . '</' . $this->qualifiedName($valuePrefix, 'r8') . '>',
                'boolean' => '<' . $this->qualifiedName($valuePrefix, 'bool') . '>' . ($this->booleanValue($value) ? 'true' : 'false') . '</' . $this->qualifiedName($valuePrefix, 'bool') . '>',
                'datetime' => '<' . $this->qualifiedName($valuePrefix, 'filetime') . '>' . $this->esc($this->dateValue($value)) . '</' . $this->qualifiedName($valuePrefix, 'filetime') . '>',
                'string' => '<' . $this->qualifiedName($valuePrefix, 'lpwstr') . '>' . $this->esc($this->stringValue($value)) . '</' . $this->qualifiedName($valuePrefix, 'lpwstr') . '>',
                default => throw new MnbExcelException('Unsupported custom property type: ' . $type),
            };
            $body .= '</' . $propertyTag . '>';
        }
        return $head . $body . $tail;
    }

    private function updateContentTypes(string $xml, bool $coreExists, bool $applicationExists, bool $customExists): string
    {
        if (trim($xml) === '') {
            throw new MnbExcelException('Invalid XLSX package: [Content_Types].xml is missing.');
        }

        $parts = [
            '/docProps/core.xml' => [$coreExists, 'application/vnd.openxmlformats-package.core-properties+xml'],
            '/docProps/app.xml' => [$applicationExists, 'application/vnd.openxmlformats-officedocument.extended-properties+xml'],
            '/docProps/custom.xml' => [$customExists, 'application/vnd.openxmlformats-officedocument.custom-properties+xml'],
        ];
        foreach ($parts as $partName => [$exists, $contentType]) {
            $xml = $this->setContentTypeOverride($xml, $partName, (string) $contentType, (bool) $exists);
        }
        return $xml;
    }

    private function setContentTypeOverride(string $xml, string $partName, string $contentType, bool $exists): string
    {
        $quotedPart = preg_quote($partName, '/');
        $pattern = '/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Override\b(?=[^>]*\bPartName=("|\')' . $quotedPart . '\1)[^>]*\/>/i';
        if (!$exists) {
            return preg_replace($pattern, '', $xml) ?? $xml;
        }
        if (preg_match($pattern, $xml) === 1) {
            return $xml;
        }
        $qualified = $this->qualifiedName(
            $this->namespacePrefix($xml, 'Types', 'http://schemas.openxmlformats.org/package/2006/content-types'),
            'Override'
        );
        $entry = '<' . $qualified . ' PartName="' . $this->esc($partName) . '" ContentType="' . $this->esc($contentType) . '"/>';
        return $this->insertBeforeClosing($xml, 'Types', $entry);
    }

    private function updateRootRelationships(string $xml, bool $coreExists, bool $applicationExists, bool $customExists): string
    {
        if (trim($xml) === '') {
            throw new MnbExcelException('Invalid XLSX package: _rels/.rels is missing.');
        }

        $relationships = [
            'core-properties' => [$coreExists, 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'],
            'extended-properties' => [$applicationExists, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', 'docProps/app.xml'],
            'custom-properties' => [$customExists, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties', 'docProps/custom.xml'],
        ];
        foreach ($relationships as $suffix => [$exists, $type, $target]) {
            $xml = $this->setRootRelationship($xml, (string) $suffix, (string) $type, (string) $target, (bool) $exists);
        }
        return $xml;
    }

    private function setRootRelationship(string $xml, string $typeSuffix, string $type, string $target, bool $exists): string
    {
        $pattern = '/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Relationship\b(?=[^>]*\bType=("|\')[^"\']*\/' . preg_quote($typeSuffix, '/') . '\1)[^>]*\/>/i';
        if (!$exists) {
            return preg_replace($pattern, '', $xml) ?? $xml;
        }
        if (preg_match($pattern, $xml) === 1) {
            return $xml;
        }

        preg_match_all('/\bId=("|\')rId([0-9]+)\1/i', $xml, $matches);
        $numbers = array_map('intval', $matches[2] ?? []);
        $id = 'rId' . (($numbers === []) ? 1 : (max($numbers) + 1));
        $qualified = $this->qualifiedName(
            $this->namespacePrefix($xml, 'Relationships', 'http://schemas.openxmlformats.org/package/2006/relationships'),
            'Relationship'
        );
        $entry = '<' . $qualified . ' Id="' . $id . '" Type="' . $this->esc($type) . '" Target="' . $this->esc($target) . '"/>';
        return $this->insertBeforeClosing($xml, 'Relationships', $entry);
    }

    private function anonymizeLegacyComments(string $xml, string $replacement): string
    {
        return preg_replace_callback(
            '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?author)\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?author\s*>/is',
            fn(array $match): string => '<' . $match[1] . '>' . $this->esc($replacement) . '</' . $match[1] . '>',
            $xml
        ) ?? $xml;
    }

    private function anonymizePersons(string $xml, string $replacement): string
    {
        return preg_replace_callback('/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?person)\b([^>]*?)\/?\s*>/i', function (array $match) use ($replacement): string {
            $attrs = $this->attributes($match[2]);
            $attrs['displayName'] = $replacement;
            if (isset($attrs['userId'])) {
                $attrs['userId'] = '';
            }
            $built = '';
            foreach ($attrs as $name => $value) {
                if ($value === '') {
                    continue;
                }
                $built .= ' ' . $name . '="' . $this->esc($value) . '"';
            }
            return '<' . $match[1] . $built . '/>';
        }, $xml) ?? $xml;
    }

    /** @return list<array{name:string,state:string}> */
    private function workbookSheets(string $xml): array
    {
        $sheets = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheet\b([^>]*)\/?\s*>/i', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $sheets[] = [
                'name' => (string) ($attrs['name'] ?? ('Sheet' . (count($sheets) + 1))),
                'state' => (string) ($attrs['state'] ?? 'visible'),
            ];
        }
        if ($sheets === []) {
            throw new MnbExcelException('Workbook does not contain any worksheets.');
        }
        return $sheets;
    }

    /** @param list<array{name:string,state:string}> $sheets */
    private function resolveSheetIndex(array $sheets, mixed $selector): int
    {
        if (is_int($selector) || (is_string($selector) && ctype_digit(trim($selector)))) {
            $index = (int) $selector;
            if ($index < 1 || !isset($sheets[$index - 1])) {
                throw new MnbExcelException('Worksheet index is out of range: ' . $index);
            }
            return $index;
        }
        $name = trim((string) $selector);
        foreach ($sheets as $index => $sheet) {
            if ($sheet['name'] === $name) {
                return $index + 1;
            }
        }
        $matches = [];
        foreach ($sheets as $index => $sheet) {
            if (strcasecmp($sheet['name'], $name) === 0) {
                $matches[] = $index + 1;
            }
        }
        if (count($matches) === 1) {
            return $matches[0];
        }
        throw new MnbExcelException('Worksheet not found: ' . $name);
    }

    private function activeSheetIndexFromXml(string $xml, int $sheetCount): int
    {
        $index = 1;
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?workbookView\b([^>]*)\/?\s*>/i', $xml, $match) === 1) {
            $attrs = $this->attributes($match[1]);
            if (isset($attrs['activeTab']) && is_numeric($attrs['activeTab'])) {
                $index = (int) $attrs['activeTab'] + 1;
            }
        }
        return min(max(1, $index), max(1, $sheetCount));
    }

    private function setSheetState(string $xml, string $sheetName, string $state): string
    {
        return preg_replace_callback(
            '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheet)\b([^>]*?)\/?\s*>/i',
            function (array $match) use ($sheetName, $state): string {
                $parsed = $this->attributes($match[2]);
                if (($parsed['name'] ?? null) !== $sheetName) {
                    return $match[0];
                }
                $attrs = preg_replace('/\s+state=("|\')[^"\']*\1/i', '', $match[2]) ?? $match[2];
                if ($state !== 'visible') {
                    $attrs .= ' state="' . $state . '"';
                }
                return '<' . $match[1] . $attrs . '/>';
            },
            $xml
        ) ?? $xml;
    }

    private function setWorkbookViewAttribute(string $xml, string $attribute, string $value): string
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?workbookView\b[^>]*\/?\s*>/i', $xml) === 1) {
            return $this->setAttributeOnFirstTag($xml, 'workbookView', $attribute, $value);
        }
        $prefix = $this->namespacePrefix(
            $xml,
            'workbook',
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        );
        $bookViews = $this->qualifiedName($prefix, 'bookViews');
        $workbookView = $this->qualifiedName($prefix, 'workbookView');
        $view = '<' . $bookViews . '><' . $workbookView . ' ' . $attribute . '="' . $this->esc($value) . '"/></' . $bookViews . '>';
        return $this->insertBeforeFirstTag($xml, 'sheets', $view);
    }

    private function setSingletonAttribute(string $xml, string $tag, string $attribute, string $value, string $insertBefore): string
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*\/?\s*>/i', $xml) === 1) {
            return $this->setAttributeOnFirstTag($xml, $tag, $attribute, $value);
        }
        $qualified = $this->qualifiedName(
            $this->namespacePrefix($xml, 'workbook', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'),
            $tag
        );
        $element = '<' . $qualified . ' ' . $attribute . '="' . $this->esc($value) . '"/>';
        return $insertBefore === 'workbook'
            ? $this->insertBeforeClosing($xml, 'workbook', $element)
            : $this->insertBeforeFirstTag($xml, $insertBefore, $element);
    }

    private function removeSingletonAttribute(string $xml, string $tag, string $attribute): string
    {
        return preg_replace_callback(
            '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . ')\b([^>]*?)(\/?)\s*>/i',
            function (array $match) use ($attribute): string {
                $attrs = preg_replace('/\s+' . preg_quote($attribute, '/') . '=("|\')[^"\']*\1/i', '', $match[2]) ?? $match[2];
                return '<' . $match[1] . $attrs . $match[3] . '>';
            },
            $xml,
            1
        ) ?? $xml;
    }

    private function setAttributeOnFirstTag(string $xml, string $tag, string $attribute, string $value): string
    {
        return preg_replace_callback(
            '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . ')\b([^>]*?)(\/?)\s*>/i',
            function (array $match) use ($attribute, $value): string {
                $attrs = preg_replace('/\s+' . preg_quote($attribute, '/') . '=("|\')[^"\']*\1/i', '', $match[2]) ?? $match[2];
                return '<' . $match[1] . $attrs . ' ' . $attribute . '="' . $this->esc($value) . '"' . $match[3] . '>';
            },
            $xml,
            1
        ) ?? $xml;
    }

    private function setElement(string $xml, string $qualified, string $local, mixed $value, string $attributes, string $rootLocal): string
    {
        $pattern = '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($local, '/') . ')\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($local, '/') . '\s*>/is';
        if ($value === null || $value === '') {
            return preg_replace($pattern, '', $xml) ?? $xml;
        }
        if (preg_match($pattern, $xml) === 1) {
            return preg_replace_callback(
                $pattern,
                fn(array $match): string => '<' . $match[1] . $attributes . '>' . $this->esc((string) $value) . '</' . $match[1] . '>',
                $xml,
                1
            ) ?? $xml;
        }
        $element = '<' . $qualified . $attributes . '>' . $this->esc((string) $value) . '</' . $qualified . '>';
        return $this->insertBeforeClosing($xml, $rootLocal, $element);
    }

    private function insertBeforeClosing(string $xml, string $localTag, string $fragment): string
    {
        $pattern = '/<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($localTag, '/') . '\s*>/i';
        if (preg_match($pattern, $xml) !== 1) {
            throw new MnbExcelException('Unable to locate XML closing tag: ' . $localTag);
        }
        return preg_replace($pattern, $fragment . '$0', $xml, 1) ?? $xml;
    }

    private function insertBeforeFirstTag(string $xml, string $localTag, string $fragment): string
    {
        $pattern = '/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($localTag, '/') . '\b/i';
        if (preg_match($pattern, $xml) !== 1) {
            return $this->insertBeforeClosing($xml, 'workbook', $fragment);
        }
        return preg_replace($pattern, $fragment . '$0', $xml, 1) ?? $xml;
    }

    /** @param list<string> $localTags */
    private function insertBeforeFirstAvailableTag(string $xml, array $localTags, string $fragment): string
    {
        foreach ($localTags as $localTag) {
            $pattern = '/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($localTag, '/') . '\b/i';
            if (preg_match($pattern, $xml) === 1) {
                return preg_replace($pattern, $fragment . '$0', $xml, 1) ?? $xml;
            }
        }
        return $this->insertBeforeClosing($xml, 'workbook', $fragment);
    }

    /** @return array{0:string,1:string} */
    private function ensureNamespace(string $xml, string $rootLocal, string $uri, string $preferredPrefix): array
    {
        $pattern = '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($rootLocal, '/') . ')\b([^>]*?)(\/?)\s*>/i';
        if (preg_match($pattern, $xml, $root) !== 1) {
            throw new MnbExcelException('Unable to locate XML root element: ' . $rootLocal);
        }

        preg_match_all('/\bxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $root[2], $matches, PREG_SET_ORDER);
        $used = [];
        foreach ($matches as $match) {
            $prefix = (string) ($match[1] ?? '');
            $value = html_entity_decode($match[3] !== '' ? $match[3] : $match[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
            $used[$prefix] = $value;
            if ($value === $uri) {
                return [$xml, $prefix];
            }
        }

        $prefix = $preferredPrefix;
        $suffix = 2;
        while (array_key_exists($prefix, $used)) {
            $prefix = $preferredPrefix . $suffix++;
        }
        $declaration = ' xmlns:' . $prefix . '="' . $this->esc($uri) . '"';
        $xml = preg_replace_callback(
            $pattern,
            static fn(array $match): string => '<' . $match[1] . $match[2] . $declaration . $match[3] . '>',
            $xml,
            1
        ) ?? $xml;
        return [$xml, $prefix];
    }

    private function namespacePrefix(string $xml, string $rootLocal, string $uri): string
    {
        $pattern = '/<((?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($rootLocal, '/') . ')\b([^>]*?)(\/?)\s*>/i';
        if (preg_match($pattern, $xml, $root) !== 1) {
            return '';
        }
        preg_match_all('/\bxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $root[2], $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $value = html_entity_decode($match[3] !== '' ? $match[3] : $match[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
            if ($value === $uri) {
                return (string) ($match[1] ?? '');
            }
        }
        return str_contains($root[1], ':') ? strstr($root[1], ':', true) : '';
    }

    private function qualifiedName(string $prefix, string $local): string
    {
        return $prefix === '' ? $local : $prefix . ':' . $local;
    }

    /** @return array<string,string> */
    private function attributes(string $source): array
    {
        $attributes = [];
        preg_match_all('/([A-Za-z0-9_:\-]+)\s*=\s*("([^"]*)"|\'([^\']*)\')/u', $source, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $name = str_contains($match[1], ':') ? substr($match[1], strrpos($match[1], ':') + 1) : $match[1];
            $attributes[$name] = html_entity_decode($match[3] !== '' ? $match[3] : $match[4], ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $attributes;
    }

    private function decodeText(string $value): string
    {
        return trim(html_entity_decode(strip_tags($value), ENT_QUOTES | ENT_XML1, 'UTF-8'));
    }

    private function inferCustomType(mixed $value): string
    {
        return match (true) {
            is_bool($value) => 'boolean',
            is_int($value) => 'integer',
            is_float($value) => 'float',
            $value instanceof DateTimeInterface => 'datetime',
            default => 'string',
        };
    }

    private function customIntegerXml(mixed $value, string $valuePrefix): string
    {
        $integer = $this->integerText($value);
        $native = $this->fitsSignedInteger($integer, '-2147483648', '2147483647') ? 'i4' : 'i8';
        $tag = $this->qualifiedName($valuePrefix, $native);
        return '<' . $tag . '>' . $integer . '</' . $tag . '>';
    }

    private function integerText(mixed $value): string
    {
        if (is_int($value)) {
            $text = (string) $value;
        } elseif (is_string($value) && preg_match('/^[+-]?\d+$/', trim($value)) === 1) {
            $text = trim($value);
        } else {
            throw new MnbExcelException('Metadata integer value must be an integer or integer string.');
        }

        $negative = str_starts_with($text, '-');
        $digits = ltrim($text, '+-0');
        if ($digits === '') {
            return '0';
        }
        $normalized = ($negative ? '-' : '') . $digits;
        if (!$this->fitsSignedInteger($normalized, '-9223372036854775808', '9223372036854775807')) {
            throw new MnbExcelException('Metadata integer value exceeds the signed 64-bit range.');
        }
        return $normalized;
    }

    private function fitsSignedInteger(string $value, string $minimum, string $maximum): bool
    {
        if (str_starts_with($value, '-')) {
            $digits = substr($value, 1);
            $limit = substr($minimum, 1);
            return strlen($digits) < strlen($limit) || (strlen($digits) === strlen($limit) && strcmp($digits, $limit) <= 0);
        }
        $digits = ltrim($value, '+');
        return strlen($digits) < strlen($maximum) || (strlen($digits) === strlen($maximum) && strcmp($digits, $maximum) <= 0);
    }

    private function integerMetadataValue(string $value): int|string
    {
        $value = trim($value);
        if (preg_match('/^[+-]?\d+$/', $value) !== 1) {
            return $value;
        }
        $negative = str_starts_with($value, '-');
        $digits = ltrim($value, '+-0');
        if ($digits === '') {
            return 0;
        }
        $normalized = ($negative ? '-' : '') . $digits;
        $limit = $negative ? substr((string) PHP_INT_MIN, 1) : (string) PHP_INT_MAX;
        $absolute = $negative ? substr($normalized, 1) : $normalized;
        if (strlen($absolute) < strlen($limit) || (strlen($absolute) === strlen($limit) && strcmp($absolute, $limit) <= 0)) {
            return (int) $normalized;
        }
        return $normalized;
    }

    private function floatValue(mixed $value): float
    {
        if (!is_int($value) && !is_float($value) && !(is_string($value) && is_numeric(trim($value)))) {
            throw new MnbExcelException('Metadata numeric value must be an integer, float, or numeric string.');
        }
        $number = (float) $value;
        if (!is_finite($number)) {
            throw new MnbExcelException('Metadata numeric value must be finite.');
        }
        return $number;
    }

    private function booleanValue(mixed $value): bool
    {
        if (is_bool($value)) {
            return $value;
        }
        if (is_int($value) && ($value === 0 || $value === 1)) {
            return $value === 1;
        }
        if (is_string($value)) {
            return match (strtolower(trim($value))) {
                '1', 'true' => true,
                '0', 'false' => false,
                default => throw new MnbExcelException('Metadata boolean value must be true, false, 1, or 0.'),
            };
        }
        throw new MnbExcelException('Metadata boolean value must be true, false, 1, or 0.');
    }

    private function stringValue(mixed $value): string
    {
        if (is_string($value) || is_int($value) || is_float($value) || is_bool($value) || $value instanceof \Stringable) {
            return (string) $value;
        }
        throw new MnbExcelException('Metadata string value must be scalar or Stringable.');
    }

    private function dateValue(mixed $value): string
    {
        if ($value instanceof DateTimeInterface) {
            $date = DateTimeImmutable::createFromInterface($value);
        } else {
            try {
                $date = new DateTimeImmutable((string) $value);
            } catch (\Throwable $error) {
                throw new MnbExcelException('Invalid metadata date value: ' . (string) $value, 0, $error);
            }
        }
        return $date->setTimezone(new DateTimeZone('UTC'))->format('Y-m-d\TH:i:s\Z');
    }

    private function number(float $value): string
    {
        $text = sprintf('%.15F', $value);
        return rtrim(rtrim($text, '0'), '.');
    }

    private function esc(string $value): string
    {
        return htmlspecialchars($value, ENT_QUOTES | ENT_XML1, 'UTF-8');
    }

    /** @param array<string,mixed> $changes */
    private function assertChanges(array $changes, bool $strict): void
    {
        if ($changes === []) {
            throw new MnbExcelException('Metadata changes cannot be empty.');
        }

        foreach ($changes as $section => $values) {
            $section = (string) $section;
            if (!in_array($section, self::ALLOWED_SECTIONS, true)) {
                if ($strict) {
                    throw new MnbExcelException('Unsupported metadata update section: ' . $section);
                }
                continue;
            }
            if (!is_array($values)) {
                throw new MnbExcelException('Metadata update section must be an array: ' . $section);
            }
            if (!$strict || $section === 'custom_properties') {
                if ($section === 'custom_properties') {
                    $this->assertCustomPropertyShape($values, $strict);
                }
                continue;
            }
            foreach (array_keys($values) as $field) {
                if (!in_array((string) $field, self::ALLOWED_FIELDS[$section], true)) {
                    throw new MnbExcelException('Unsupported metadata field: ' . $section . '.' . $field);
                }
            }
        }

        $document = (array) ($changes['document'] ?? []);
        foreach (['title', 'subject', 'creator', 'description', 'category', 'content_status', 'identifier', 'language', 'document_version'] as $field) {
            if (array_key_exists($field, $document) && $document[$field] !== null) {
                $this->stringValue($document[$field]);
            }
        }
        if (array_key_exists('keywords', $document) && $document['keywords'] !== null) {
            if (is_array($document['keywords'])) {
                foreach ($document['keywords'] as $keyword) {
                    $this->stringValue($keyword);
                }
            } else {
                $this->stringValue($document['keywords']);
            }
        }

        $revision = (array) ($changes['revision'] ?? []);
        foreach (['last_saved_by', 'revision_number'] as $field) {
            if (array_key_exists($field, $revision) && $revision[$field] !== null) {
                $this->stringValue($revision[$field]);
            }
        }
        if (array_key_exists('total_editing_time_seconds', $revision) && $revision['total_editing_time_seconds'] !== null) {
            $seconds = $this->integerText($revision['total_editing_time_seconds']);
            if (str_starts_with($seconds, '-')) {
                throw new MnbExcelException('Revision total_editing_time_seconds cannot be negative.');
            }
        }
        foreach (['last_printed_at', 'document_created_at', 'document_modified_at'] as $field) {
            if (array_key_exists($field, $revision) && $revision[$field] !== null && $revision[$field] !== '') {
                $this->dateValue($revision[$field]);
            }
        }

        $application = (array) ($changes['application'] ?? []);
        foreach (['application', 'application_version', 'company', 'manager', 'operating_system_hint'] as $field) {
            if (array_key_exists($field, $application) && $application[$field] !== null) {
                $this->stringValue($application[$field]);
            }
        }
        if (array_key_exists('document_security', $application) && $application['document_security'] !== null) {
            $security = $this->integerText($application['document_security']);
            if (str_starts_with($security, '-')) {
                throw new MnbExcelException('Application document_security cannot be negative.');
            }
        }
        foreach (['scale_crop', 'links_up_to_date', 'shared_document', 'hyperlinks_changed'] as $field) {
            if (array_key_exists($field, $application) && $application[$field] !== null) {
                $this->booleanValue($application[$field]);
            }
        }

        $workbook = (array) ($changes['workbook'] ?? []);
        if (array_key_exists('sheet_visibility', $workbook) && !is_array($workbook['sheet_visibility'])) {
            throw new MnbExcelException('Workbook sheet_visibility must be an array.');
        }
        if (array_key_exists('active_sheet', $workbook) && !is_int($workbook['active_sheet']) && !is_string($workbook['active_sheet'])) {
            throw new MnbExcelException('Workbook active_sheet must be a worksheet index or name.');
        }
        if (array_key_exists('code_name', $workbook) && $workbook['code_name'] !== null) {
            $this->stringValue($workbook['code_name']);
        }
        foreach (['date1904', 'read_only_recommended'] as $field) {
            if (array_key_exists($field, $workbook)) {
                $this->booleanValue($workbook[$field]);
            }
        }

        $calculation = (array) ($changes['calculation'] ?? []);
        if (array_key_exists('mode', $calculation) && !in_array($calculation['mode'], ['auto', 'manual', 'autoNoTable'], true)) {
            throw new MnbExcelException('Calculation mode must be auto, manual, or autoNoTable.');
        }
        foreach (['calc_id', 'iterate_count'] as $field) {
            if (array_key_exists($field, $calculation)) {
                $value = $this->integerText($calculation[$field]);
                if (str_starts_with($value, '-')) {
                    throw new MnbExcelException('Calculation ' . $field . ' cannot be negative.');
                }
            }
        }
        foreach (['full_calc_on_load', 'force_full_calc', 'calc_on_save', 'iterate'] as $field) {
            if (array_key_exists($field, $calculation)) {
                $this->booleanValue($calculation[$field]);
            }
        }
        if (array_key_exists('iterate_delta', $calculation)) {
            $delta = $this->floatValue($calculation['iterate_delta']);
            if ($delta < 0) {
                throw new MnbExcelException('Calculation iterate_delta cannot be negative.');
            }
        }
    }

    /** @param array<mixed> $changes */
    private function assertCustomPropertyShape(array $changes, bool $strict): void
    {
        if (!$strict) {
            return;
        }
        foreach ($changes as $key => $value) {
            if (is_int($key)) {
                if (!is_array($value)) {
                    throw new MnbExcelException('Each custom property list item must be an array.');
                }
                foreach (array_keys($value) as $field) {
                    if (!in_array((string) $field, ['name', 'type', 'value', 'linked', 'link_target'], true)) {
                        throw new MnbExcelException('Unsupported custom property field: ' . $field);
                    }
                }
                continue;
            }
            if (trim((string) $key) === '') {
                throw new MnbExcelException('Custom property name cannot be empty.');
            }
            if (is_array($value) && (array_key_exists('value', $value) || array_key_exists('type', $value))) {
                foreach (array_keys($value) as $field) {
                    if (!in_array((string) $field, ['type', 'value', 'linked', 'link_target'], true)) {
                        throw new MnbExcelException('Unsupported custom property field: ' . $field);
                    }
                }
            }
        }
    }

    private function temporaryPath(string $prefix): string
    {
        $path = tempnam(sys_get_temp_dir(), $prefix);
        if ($path === false) {
            throw new MnbExcelException('Unable to allocate temporary metadata file.');
        }
        return $path;
    }

    private function newCorePropertiesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            . 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            . 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            . '</cp:coreProperties>';
    }

    private function newApplicationPropertiesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            . 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            . '</Properties>';
    }
}
