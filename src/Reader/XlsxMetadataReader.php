<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Metadata\MetadataCapabilities;
use Mnb\PHPExcel\Metadata\MetadataOptions;
use Mnb\PHPExcel\Metadata\MetadataProfile;
use Mnb\PHPExcel\Metadata\MetadataReport;
use Mnb\PHPExcel\Metadata\MetadataSectionState;
use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/**
 * Consolidated XLSX metadata collector.
 *
 * This class is intentionally read-only. It inventories workbook features and
 * never executes formulas, VBA, external connections, ActiveX, or embedded data.
 */
final class XlsxMetadataReader
{
    public function __construct(private ?XlsxWorkbookResolver $resolver = null)
    {
        $this->resolver = $resolver ?? new XlsxWorkbookResolver();
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function metaInfo(string $path, array $options = []): array
    {
        $settings = MetadataOptions::fromArray($options);
        $realPath = realpath($path);
        if ($realPath === false || !is_file($realPath)) {
            throw new MnbExcelException('Invalid XLSX path: ' . $path);
        }

        $encryption = new XlsxEncryption();
        $encrypted = $encryption->isEncryptedFile($realPath);
        $encryptionMode = $encryption->encryptionMode($realPath);

        if ($encrypted && $settings->password() === '') {
            return $this->passwordRequiredReport($realPath, $encryptionMode, $settings);
        }

        $temporary = null;
        $readPath = $realPath;
        if ($encrypted) {
            $temporary = $encryption->decryptToTemporary($realPath, $settings->password(), $settings->toArray());
            $readPath = $temporary;
        }

        try {
            $report = $this->readPlainPackage($readPath, $realPath, $settings, $encrypted, $encryptionMode);
            if ($encrypted) {
                $report['warnings'][] = 'Workbook metadata was inspected from a temporary decrypted package.';
                $report['warnings'] = array_values(array_unique($report['warnings']));
            }
            return $report;
        } finally {
            if ($temporary !== null) {
                @unlink($temporary);
            }
        }
    }

    /** @return array<string,mixed> */
    private function readPlainPackage(
        string $readPath,
        string $sourcePath,
        MetadataOptions $settings,
        bool $encrypted,
        ?string $encryptionMode
    ): array {
        $extension = strtolower((string) pathinfo($sourcePath, PATHINFO_EXTENSION));
        $variant = $this->variantFromExtension($extension);
        $mime = $this->mimeForVariant($variant);
        $report = new MetadataReport('xlsx', $variant, $mime, $settings->profile());

        $quick = (new XlsxQuickInfo($this->resolver))->fileInfo($readPath);
        $sheets = (new XlsxQuickInfo($this->resolver))->sheetsInfo($readPath, [
            'accurate_row_count' => $settings->accurateSheetCounts(),
            'include_hidden' => true,
        ]);
        $activeSheet = $this->resolver->activeSheet($readPath);
        $inspection = $settings->atLeast(MetadataProfile::STANDARD)
            ? (new XlsxInspector($this->resolver))->inspect($readPath)
            : ['warnings' => [], 'errors' => [], 'sheets' => []];

        $zip = $this->openZip($readPath);
        try {
            $package = $this->packageInventory($zip, $settings);
            $core = $this->readCoreProperties((string) ($zip->getFromName('docProps/core.xml') ?: ''));
            $application = $this->readApplicationProperties((string) ($zip->getFromName('docProps/app.xml') ?: ''));
            $custom = $this->readCustomProperties((string) ($zip->getFromName('docProps/custom.xml') ?: ''));
            $workbookXml = (string) ($zip->getFromName('xl/workbook.xml') ?: '');
            $workbookSettings = $this->workbookSettings($workbookXml);
            $definedNames = $this->definedNames($workbookXml, $sheets);
            $featureData = $this->featureInventory($readPath, $zip, $sheets, $inspection, $settings);
        } finally {
            $zip->close();
        }

        $file = $this->fileInformation($sourcePath, $settings);
        $file += [
            'container' => $encrypted ? 'ole-encrypted-ooxml' : 'zip',
            'inner_package_container' => 'zip',
            'zip_entries' => (int) ($quick['zip_entries'] ?? 0),
            'compressed_package_bytes' => $quick['compressed_package_bytes'] ?? null,
            'uncompressed_package_bytes' => $quick['uncompressed_package_bytes'] ?? null,
            'compression_ratio' => $quick['compression_ratio'] ?? null,
        ];
        $report->setSection('file', MetadataSectionState::AVAILABLE, $file);

        $report->setSection('format_details', MetadataSectionState::AVAILABLE, [
            'ooxml_variant' => $variant,
            'container' => $encrypted ? 'ole-encrypted-ooxml' : 'zip',
            'inner_package_container' => 'zip',
            'zip_entries' => (int) ($quick['zip_entries'] ?? 0),
            'content_types' => $package['content_types'],
            'shared_strings_present' => $package['shared_strings_present'],
            'styles_present' => $package['styles_present'],
            'workbook_relationships_present' => $package['workbook_relationships_present'],
            'package_part_count' => count($package['parts']),
            'count' => count($package['parts']),
            'items' => $settings->includePackageParts() ? $package['parts_limited'] : [],
            'truncated' => $settings->includePackageParts() && $package['parts_truncated'],
        ]);

        $document = [
            'title' => $core['title'] ?? null,
            'subject' => $core['subject'] ?? null,
            'creator' => $core['creator'] ?? null,
            'authors' => isset($core['creator']) && $core['creator'] !== '' ? [$core['creator']] : [],
            'manager' => $application['manager'] ?? null,
            'company' => $application['company'] ?? null,
            'category' => $core['category'] ?? null,
            'keywords' => $this->keywords((string) ($core['keywords'] ?? '')),
            'keywords_original' => $core['keywords'] ?? null,
            'description' => $core['description'] ?? null,
            'content_status' => $core['content_status'] ?? null,
            'identifier' => $core['identifier'] ?? null,
            'language' => $core['language'] ?? null,
            'document_version' => $core['version'] ?? null,
        ];
        $report->setSection('document', MetadataSectionState::AVAILABLE, $document + [
            'count' => $this->nonEmptyCount($document),
        ]);

        $revision = [
            'last_saved_by' => $core['last_modified_by'] ?? null,
            'revision_number' => $core['revision'] ?? null,
            'total_editing_time_seconds' => $application['total_time_minutes'] === null
                ? null
                : ((int) $application['total_time_minutes'] * 60),
            'last_printed_at' => $core['last_printed'] ?? null,
            'document_created_at' => $core['created'] ?? null,
            'document_modified_at' => $core['modified'] ?? null,
        ];
        $report->setSection('revision', MetadataSectionState::AVAILABLE, $revision + [
            'count' => $this->nonEmptyCount($revision),
            'warnings' => ['Office revision properties describe the current document state, not a complete version history.'],
        ]);

        $report->setSection('application', MetadataSectionState::AVAILABLE, $application + [
            'count' => $this->nonEmptyCount($application),
            'warnings' => ['Application and operating-system values are producer-supplied hints and are not independently verified.'],
        ]);

        [$customItems, $customTruncated] = $this->limit($custom, $settings->maxItems());
        $report->setSection('custom_properties', MetadataSectionState::AVAILABLE, [
            'count' => count($custom),
            'items' => $customItems,
            'truncated' => $customTruncated,
        ]);

        $sheetItems = [];
        $hiddenSheets = [];
        foreach ($sheets as $sheet) {
            $item = [
                'index' => (int) ($sheet['index'] ?? 0),
                'name' => (string) ($sheet['name'] ?? ''),
                'sheet_id' => (int) ($sheet['sheet_id'] ?? 0),
                'type' => 'worksheet',
                'state' => (string) ($sheet['state'] ?? 'visible'),
                'path' => (string) ($sheet['path'] ?? ''),
                'exists' => (bool) ($sheet['exists'] ?? false),
                'dimension' => $sheet['dimension'] ?? null,
                'declared_last_row' => $sheet['declared_last_row'] ?? null,
                'declared_last_column' => $sheet['declared_last_column'] ?? null,
                'physical_rows' => $sheet['physical_rows'] ?? null,
                'filled_rows' => $sheet['filled_rows'] ?? null,
                'last_row' => $sheet['last_row'] ?? null,
                'last_column' => $sheet['last_column'] ?? null,
                'columns' => $sheet['columns'] ?? null,
                'cells' => $sheet['cells'] ?? null,
                'worksheet_xml_size_bytes' => $sheet['worksheet_xml_size_bytes'] ?? null,
                'worksheet_xml_compressed_bytes' => $sheet['worksheet_xml_compressed_bytes'] ?? null,
            ];
            $sheetItems[] = $item;
            if ($item['state'] !== 'visible') {
                $hiddenSheets[] = $item;
            }
        }
        $workbook = [
            'name' => pathinfo($sourcePath, PATHINFO_FILENAME),
            'sheet_count' => count($sheetItems),
            'active_sheet' => $activeSheet,
            'first_visible_sheet' => $this->firstVisibleSheet($sheetItems),
            'selected_sheets' => $settings->atLeast(MetadataProfile::STANDARD) ? $featureData['selected_sheets'] : null,
            'sheets' => $sheetItems,
            'date_system' => $workbookSettings['date1904'] ? '1904' : '1900',
            'date1904' => $workbookSettings['date1904'],
            'workbook_code_name' => $workbookSettings['code_name'],
            'workbook_view' => $workbookSettings['view'],
            'count' => count($sheetItems),
            'items' => $sheetItems,
        ];
        $report->setSection('workbook', MetadataSectionState::AVAILABLE, $workbook);

        $security = $this->securityInformation($readPath, $sheetItems, $settings, $encrypted, $encryptionMode, $featureData);
        $report->setSection('security', $security['state'], $security['data']);

        $report->setSection('macros', $featureData['states']['macros'], $featureData['macros']);

        $namedObjects = array_merge($definedNames, $featureData['named_objects']);
        [$namedItems, $namedTruncated] = $this->limit($namedObjects, $settings->maxItems());
        $report->setSection('named_objects', $featureData['states']['named_objects'], [
            'count' => count($namedObjects),
            'items' => $namedItems,
            'truncated' => $namedTruncated,
            'defined_name_count' => count($definedNames),
            'table_count' => $featureData['counts']['tables'],
            'chart_count' => $featureData['counts']['charts'],
            'pivot_table_count' => $featureData['counts']['pivot_tables'],
            'slicer_count' => $featureData['counts']['slicers'],
            'warnings' => $featureData['states']['named_objects'] === MetadataSectionState::PARTIAL
                ? ['Object inventory is complete at package-part level; some object-specific fields are not yet decoded.']
                : [],
        ]);

        $report->setSection('links', $featureData['states']['links'], $featureData['links']);

        $hiddenItems = array_merge($hiddenSheets, $featureData['hidden_items']);
        [$hiddenLimited, $hiddenTruncated] = $this->limit($hiddenItems, $settings->maxItems());
        $report->setSection('hidden_content', $featureData['states']['hidden_content'], [
            'count' => count($hiddenItems),
            'items' => $hiddenLimited,
            'truncated' => $hiddenTruncated,
            'hidden_sheet_count' => count(array_filter($hiddenSheets, static fn(array $item): bool => $item['state'] === 'hidden')),
            'very_hidden_sheet_count' => count(array_filter($hiddenSheets, static fn(array $item): bool => $item['state'] === 'veryHidden')),
            'hidden_row_count' => $featureData['counts']['hidden_rows'],
            'hidden_column_range_count' => $featureData['counts']['hidden_columns'],
            'warnings' => $settings->atLeast(MetadataProfile::STANDARD)
                ? []
                : ['Quick profile reports hidden worksheets but does not scan hidden rows or columns.'],
        ]);

        $report->setSection('comments_notes', $featureData['states']['comments_notes'], $featureData['comments_notes']);
        $report->setSection('tracked_changes', $featureData['tracked_changes']['state'], $featureData['tracked_changes']['data']);
        $report->setSection('embedded_objects', $featureData['states']['embedded_objects'], $featureData['embedded_objects']);
        $report->setSection('calculation', $featureData['states']['calculation'], $featureData['calculation'] + [
            'settings' => $workbookSettings['calculation'],
        ]);
        $report->setSection('print_settings', $featureData['states']['print_settings'], $featureData['print_settings']);
        $report->setSection('validation', $featureData['states']['validation'], $featureData['validation']);
        $report->setSection('pivot_metadata', $featureData['states']['pivot_metadata'], $featureData['pivot_metadata']);
        $report->setSection('xml_metadata', MetadataSectionState::AVAILABLE, $featureData['xml_metadata'] + [
            'content_types' => $package['content_types'],
            'package_part_count' => count($package['parts']),
        ]);

        $statistics = [
            'sheet_count' => count($sheetItems),
            'declared_rows_total' => array_sum(array_map(static fn(array $item): int => (int) ($item['declared_last_row'] ?? 0), $sheetItems)),
            'physical_rows_total' => $settings->accurateSheetCounts()
                ? array_sum(array_map(static fn(array $item): int => (int) ($item['physical_rows'] ?? 0), $sheetItems))
                : null,
            'filled_rows_total' => $settings->accurateSheetCounts()
                ? array_sum(array_map(static fn(array $item): int => (int) ($item['filled_rows'] ?? 0), $sheetItems))
                : null,
            'cell_count_total' => $settings->accurateSheetCounts()
                ? array_sum(array_map(static fn(array $item): int => (int) ($item['cells'] ?? 0), $sheetItems))
                : null,
            'formula_count' => $settings->atLeast(MetadataProfile::STANDARD) ? $featureData['counts']['formulas'] : null,
            'comment_count' => $settings->atLeast(MetadataProfile::FULL) ? $featureData['counts']['comments'] : null,
            'hyperlink_count' => $settings->atLeast(MetadataProfile::STANDARD) ? $featureData['counts']['hyperlinks'] : null,
            'image_count' => $settings->atLeast(MetadataProfile::FULL) ? $featureData['counts']['images'] : null,
            'custom_property_count' => count($custom),
        ];
        $report->setSection('statistics', MetadataSectionState::AVAILABLE, $statistics + [
            'count' => $this->nonEmptyCount($statistics),
        ]);

        foreach ((array) ($quick['warnings'] ?? []) as $warning) {
            $report->warning((string) $warning);
        }
        foreach ((array) ($inspection['warnings'] ?? []) as $warning) {
            $report->warning((string) $warning);
        }
        foreach ((array) ($inspection['errors'] ?? []) as $error) {
            $report->error((string) $error);
        }
        foreach ($featureData['warnings'] as $warning) {
            $report->warning($warning);
        }

        $array = $report->toArray();
        $report->capabilities(MetadataCapabilities::fromReport($array, [
            'document', 'revision', 'application', 'custom_properties', 'workbook', 'calculation',
        ]));

        return $report->toArray();
    }

    /** @return array<string,mixed> */
    private function passwordRequiredReport(string $path, ?string $mode, MetadataOptions $settings): array
    {
        $extension = strtolower((string) pathinfo($path, PATHINFO_EXTENSION));
        $variant = $this->variantFromExtension($extension);
        $report = new MetadataReport('xlsx', $variant, $this->mimeForVariant($variant), $settings->profile());
        $report->status('password_required')
            ->setSection('file', MetadataSectionState::AVAILABLE, $this->fileInformation($path, $settings) + [
                'container' => 'ole-encrypted-ooxml',
            ])
            ->setSection('format_details', MetadataSectionState::ENCRYPTED, [
                'ooxml_variant' => $variant,
                'container' => 'ole-encrypted-ooxml',
            ])
            ->setSection('security', MetadataSectionState::PASSWORD_REQUIRED, [
                'file_encrypted' => true,
                'encryption_mode' => $mode,
                'password_required' => true,
                'count' => 1,
            ]);

        foreach (MetadataReport::SECTIONS as $section) {
            if (in_array($section, ['file', 'format_details', 'security'], true)) {
                continue;
            }
            $report->setSection($section, MetadataSectionState::ENCRYPTED, [
                'warnings' => ['A password is required to inspect this metadata section.'],
            ]);
        }
        $report->warning('A password is required to inspect workbook package metadata.');
        $array = $report->toArray();
        $report->capabilities(MetadataCapabilities::fromReport($array));
        return $report->toArray();
    }

    /** @return array<string,mixed> */
    private function fileInformation(string $path, MetadataOptions $settings): array
    {
        $size = filesize($path);
        $modified = filemtime($path);
        $accessed = fileatime($path);
        $changed = filectime($path);
        $permissions = fileperms($path);

        return [
            'path' => $path,
            'resolved_path' => realpath($path) ?: $path,
            'name' => basename($path),
            'extension' => strtolower((string) pathinfo($path, PATHINFO_EXTENSION)),
            'size_bytes' => $size === false ? null : (int) $size,
            'filesystem_created_at' => PHP_OS_FAMILY === 'Windows' && $changed !== false ? date(DATE_ATOM, $changed) : null,
            'filesystem_status_changed_at' => $changed === false ? null : date(DATE_ATOM, $changed),
            'filesystem_modified_at' => $modified === false ? null : date(DATE_ATOM, $modified),
            'filesystem_accessed_at' => $accessed === false ? null : date(DATE_ATOM, $accessed),
            'readable' => is_readable($path),
            'writable' => is_writable($path),
            'permissions_octal' => $permissions === false ? null : substr(sprintf('%o', $permissions), -4),
            'sha256' => $settings->includeHash() ? (hash_file('sha256', $path) ?: null) : null,
            'count' => 1,
        ];
    }

    /** @return array<string,mixed> */
    private function packageInventory(ZipArchive $zip, MetadataOptions $settings): array
    {
        $parts = [];
        for ($index = 0; $index < $zip->numFiles; $index++) {
            $name = (string) ($zip->getNameIndex($index) ?: '');
            if ($name === '') {
                continue;
            }
            $stat = $zip->statIndex($index) ?: [];
            $item = [
                'path' => $name,
                'size_bytes' => (int) ($stat['size'] ?? 0),
                'compressed_size_bytes' => (int) ($stat['comp_size'] ?? 0),
            ];
            if ($settings->profile() === MetadataProfile::FORENSIC) {
                $bytes = $zip->getFromName($name);
                $item['sha256'] = $bytes === false ? null : hash('sha256', $bytes);
            }
            $parts[] = $item;
        }
        [$limited, $truncated] = $this->limit($parts, $settings->maxItems());
        $contentTypes = $this->contentTypes((string) ($zip->getFromName('[Content_Types].xml') ?: ''));

        return [
            'parts' => $parts,
            'parts_limited' => $limited,
            'parts_truncated' => $truncated,
            'content_types' => $contentTypes,
            'shared_strings_present' => $zip->locateName('xl/sharedStrings.xml') !== false,
            'styles_present' => $zip->locateName('xl/styles.xml') !== false,
            'workbook_relationships_present' => $zip->locateName('xl/_rels/workbook.xml.rels') !== false,
        ];
    }

    /** @return array<string,mixed> */
    private function featureInventory(string $packagePath, ZipArchive $zip, array $sheets, array $inspection, MetadataOptions $settings): array
    {
        $entries = [];
        for ($index = 0; $index < $zip->numFiles; $index++) {
            $name = (string) ($zip->getNameIndex($index) ?: '');
            if ($name !== '') {
                $entries[] = $name;
            }
        }

        $namedObjects = [];
        foreach ($entries as $entry) {
            $type = match (true) {
                preg_match('#^xl/tables/[^/]+\.xml$#i', $entry) === 1 => 'table',
                preg_match('#^xl/charts/[^/]+\.xml$#i', $entry) === 1 => 'chart',
                preg_match('#^xl/pivotTables/[^/]+\.xml$#i', $entry) === 1 => 'pivot_table',
                preg_match('#^xl/slicers/[^/]+\.xml$#i', $entry) === 1 => 'slicer',
                preg_match('#^xl/timelines/[^/]+\.xml$#i', $entry) === 1 => 'timeline',
                preg_match('#^xl/ctrlProps/[^/]+\.xml$#i', $entry) === 1 => 'form_control',
                default => null,
            };
            if ($type === null) {
                continue;
            }
            $xml = (string) ($zip->getFromName($entry) ?: '');
            $rootTag = match ($type) {
                'table' => 'table',
                'pivot_table' => 'pivotTableDefinition',
                'slicer' => 'slicerList',
                'timeline' => 'timelineCacheDefinition',
                default => '',
            };
            $attrs = $rootTag === '' ? [] : $this->firstTagAttributes($xml, $rootTag);
            $location = $type === 'pivot_table' ? $this->firstTagAttributes($xml, 'location') : [];
            $namedObjects[] = [
                'type' => $type,
                'path' => $entry,
                'name' => $attrs['displayName'] ?? $attrs['name'] ?? pathinfo($entry, PATHINFO_FILENAME),
                'source_range' => $attrs['ref'] ?? $location['ref'] ?? null,
                'cache_id' => isset($attrs['cacheId']) ? (int) $attrs['cacheId'] : null,
            ];
        }

        $macroParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#(^|/)vbaProject\.bin$#i', $entry) === 1
            || preg_match('#vbaProjectSignature#i', $entry) === 1
        ));
        $signatureParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#^_xmlsignatures/#i', $entry) === 1
            || preg_match('#signature#i', $entry) === 1 && preg_match('#\.(xml|bin)$#i', $entry) === 1
        ));
        $commentParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#^xl/comments[^/]*\.xml$#i', $entry) === 1
            || preg_match('#^xl/threadedComments/#i', $entry) === 1
        ));
        $mediaParts = $this->paths($entries, '#^xl/media/#i');
        $macros = [
            'present' => $macroParts !== [],
            'project_part' => $this->firstMatching($macroParts, '#vbaProject\.bin$#i'),
            'signature_present' => $this->firstMatching($signatureParts, '#vbaProjectSignature#i') !== null,
            'module_count' => null,
            'module_names' => [],
            'count' => count($macroParts),
            'items' => array_map(static fn(string $path): array => ['path' => $path], $macroParts),
            'warnings' => $macroParts === [] ? [] : ['VBA content was inventoried but never executed or decompiled.'],
        ];

        $externalLinks = $this->paths($entries, '#^xl/externalLinks/[^/]+\.xml$#i');
        $connections = $this->paths($entries, '#^xl/connections\.xml$#i');
        $queryTables = $this->paths($entries, '#^xl/queryTables/[^/]+\.xml$#i');
        $powerQuery = array_values(array_filter($entries, static fn(string $entry): bool =>
            stripos($entry, 'customXml') !== false && stripos($entry, 'mashup') !== false
            || stripos($entry, 'connections') !== false && str_ends_with(strtolower($entry), '.bin')
        ));

        $comments = [];
        $hyperlinks = [];
        $images = [];
        $hiddenItems = [];
        $printSettings = [];
        $validations = [];
        $conditionalFormatting = [];
        $formulas = [];
        $warnings = [];
        $selectedSheets = [];
        $hiddenRowCount = 0;
        $hiddenColumnCount = 0;

        $inspectionByName = [];
        foreach ((array) ($inspection['sheets'] ?? []) as $sheetInspection) {
            if (is_array($sheetInspection)) {
                $inspectionByName[(string) ($sheetInspection['name'] ?? '')] = $sheetInspection;
            }
        }

        if ($settings->atLeast(MetadataProfile::STANDARD)) {
            foreach ($sheets as $sheet) {
                $sheetName = (string) ($sheet['name'] ?? '');
                $sheetPath = (string) ($sheet['path'] ?? '');
                if ($sheetPath === '') {
                    continue;
                }
                $sheetXml = (string) ($zip->getFromName($sheetPath) ?: '');
                $sheetScan = $this->scanWorksheetXml($sheetXml, $sheetName, $settings);
                if ($sheetScan['tab_selected']) {
                    $selectedSheets[] = [
                        'index' => (int) ($sheet['index'] ?? 0),
                        'name' => $sheetName,
                    ];
                }
                $hiddenItems = array_merge($hiddenItems, $sheetScan['hidden_items']);
                $hiddenRowCount += $sheetScan['hidden_row_count'];
                $hiddenColumnCount += $sheetScan['hidden_column_count'];
                $printSettings = array_merge($printSettings, $sheetScan['print_settings']);
                $validations = array_merge($validations, $sheetScan['validations']);
                $conditionalFormatting = array_merge($conditionalFormatting, $sheetScan['conditional_formatting']);
                $formulas = array_merge($formulas, $sheetScan['formulas']);
                if (!$settings->atLeast(MetadataProfile::FULL)) {
                    $hyperlinks = array_merge($hyperlinks, $sheetScan['hyperlinks']);
                }

                $inspected = $inspectionByName[$sheetName] ?? [];
                if (!$settings->atLeast(MetadataProfile::FULL)) {
                    $hiddenRowCount = max($hiddenRowCount, (int) ($inspected['hidden_row_count'] ?? 0));
                    $hiddenColumnCount = max($hiddenColumnCount, (int) ($inspected['hidden_column_count'] ?? 0));
                }
            }
        }

        // Rich per-sheet objects need a package file path. The caller only invokes this
        // on an ordinary or temporary decrypted package, so resolver paths are safe.
        if ($settings->atLeast(MetadataProfile::FULL)) {
            foreach ($sheets as $sheet) {
                    $sheetName = (string) ($sheet['name'] ?? '');
                    try {
                        $meta = (new XlsxMetadataExtractor($this->resolver))->readSheetMetadata($packagePath, $sheetName);
                    } catch (\Throwable $error) {
                        $warnings[] = 'Unable to read detailed metadata for sheet "' . $sheetName . '": ' . $error->getMessage();
                        continue;
                    }
                    foreach ((array) ($meta['comments'] ?? []) as $item) {
                        if (is_array($item)) {
                            $comments[] = ['sheet' => $sheetName] + $item;
                        }
                    }
                    foreach ((array) ($meta['hyperlinks'] ?? []) as $item) {
                        if (is_array($item)) {
                            $hyperlinks[] = ['sheet' => $sheetName] + $item;
                        }
                    }
                    foreach ((array) ($meta['images'] ?? []) as $item) {
                        if (is_array($item)) {
                            unset($item['bytes']);
                            $images[] = ['sheet' => $sheetName] + $item;
                        }
                    }
            }
        }

        $oleParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#^xl/embeddings/#i', $entry) === 1
            || preg_match('#^xl/activeX/#i', $entry) === 1
            || preg_match('#^xl/ctrlProps/#i', $entry) === 1
        ));
        $embedded = array_merge(
            $images,
            array_map(static fn(string $path): array => [
                'type' => str_starts_with(strtolower($path), 'xl/activex/') ? 'activex' : 'ole_or_embedded_document',
                'path' => $path,
            ], $oleParts)
        );

        $trackedParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#^xl/revisions/#i', $entry) === 1
            || preg_match('#^xl/userNames\.xml$#i', $entry) === 1
            || preg_match('#revisionLog#i', $entry) === 1
        ));
        $trackedState = $trackedParts === [] ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL;

        $pivotParts = array_values(array_filter($entries, static fn(string $entry): bool =>
            preg_match('#^xl/pivotTables/#i', $entry) === 1
            || preg_match('#^xl/pivotCache/#i', $entry) === 1
        ));

        $xmlParts = $this->paths($entries, '#^customXml/#i');
        $relationshipScan = $settings->includeRelationships()
            ? $this->relationships($zip, $entries, $settings->maxItems())
            : ['items' => [], 'count' => null, 'truncated' => false];
        $relationshipItems = $relationshipScan['items'];

        [$linkItems, $linksTruncated] = $this->limit(array_merge(
            array_map(static fn(string $path): array => ['type' => 'external_workbook', 'path' => $path], $externalLinks),
            array_map(static fn(string $path): array => ['type' => 'connection', 'path' => $path], $connections),
            array_map(static fn(string $path): array => ['type' => 'query_table', 'path' => $path], $queryTables),
            array_map(static fn(string $path): array => ['type' => 'power_query_or_mashup', 'path' => $path], $powerQuery),
            $hyperlinks
        ), $settings->maxItems());
        [$commentItems, $commentsTruncated] = $this->limit($comments, $settings->maxItems());
        [$embeddedItems, $embeddedTruncated] = $this->limit($embedded, $settings->maxItems());
        [$printItems, $printTruncated] = $this->limit($printSettings, $settings->maxItems());
        [$validationItems, $validationTruncated] = $this->limit(array_merge($validations, $conditionalFormatting), $settings->maxItems());
        [$formulaItems, $formulaTruncated] = $this->limit($formulas, $settings->maxItems());
        [$pivotItems, $pivotTruncated] = $this->limit(array_map(static fn(string $path): array => ['path' => $path], $pivotParts), $settings->maxItems());
        [$xmlItems, $xmlTruncated] = $this->limit(array_map(static fn(string $path): array => ['path' => $path], $xmlParts), $settings->maxItems());

        return [
            'selected_sheets' => $selectedSheets,
            'named_objects' => $namedObjects,
            'macros' => $macros,
            'links' => [
                'count' => $settings->atLeast(MetadataProfile::STANDARD)
                    ? count($externalLinks) + count($connections) + count($queryTables) + count($powerQuery) + count($hyperlinks)
                    : null,
                'known_package_link_count' => count($externalLinks) + count($connections) + count($queryTables) + count($powerQuery),
                'items' => $linkItems,
                'truncated' => $linksTruncated,
                'external_workbook_link_count' => count($externalLinks),
                'connection_count' => count($connections),
                'query_table_count' => count($queryTables),
                'power_query_or_mashup_count' => count($powerQuery),
                'hyperlink_count' => $settings->atLeast(MetadataProfile::STANDARD) ? count($hyperlinks) : null,
                'warnings' => $settings->atLeast(MetadataProfile::STANDARD)
                    ? ['External targets and connections were inventoried without contacting them.']
                    : ['Quick profile inventories package-level links but does not scan worksheet hyperlinks.'],
            ],
            'hidden_items' => $hiddenItems,
            'comments_notes' => [
                'count' => $settings->atLeast(MetadataProfile::FULL) ? count($comments) : count($commentParts),
                'items' => $commentItems,
                'truncated' => $commentsTruncated,
                'comment_count' => $settings->atLeast(MetadataProfile::FULL) ? count($comments) : null,
                'comment_part_count' => count($commentParts),
                'threaded_comment_part_count' => count($this->paths($entries, '#^xl/threadedComments/#i')),
                'warnings' => $settings->atLeast(MetadataProfile::FULL) ? [] : ['Detailed comment and reviewer data requires the full profile.'],
            ],
            'tracked_changes' => [
                'state' => $trackedState,
                'data' => [
                    'present' => $trackedParts !== [],
                    'count' => count($trackedParts),
                    'items' => array_map(static fn(string $path): array => ['path' => $path], $trackedParts),
                    'warnings' => $trackedParts === [] ? [] : ['Tracked-change package parts are inventoried, but full change-history interpretation is not yet implemented.'],
                ],
            ],
            'embedded_objects' => [
                'count' => $settings->atLeast(MetadataProfile::FULL) ? count($embedded) : count($mediaParts) + count($oleParts),
                'items' => $embeddedItems,
                'truncated' => $embeddedTruncated,
                'image_count' => $settings->atLeast(MetadataProfile::FULL) ? count($images) : count($mediaParts),
                'media_part_count' => count($mediaParts),
                'ole_activex_part_count' => count($oleParts),
                'warnings' => $settings->atLeast(MetadataProfile::FULL) ? [] : ['Quick and standard profiles inventory media package parts without resolving worksheet anchors.'],
            ],
            'calculation' => [
                'formula_count' => $settings->atLeast(MetadataProfile::STANDARD) ? count($formulas) : null,
                'items' => $settings->atLeast(MetadataProfile::FULL) ? $formulaItems : [],
                'count' => $settings->atLeast(MetadataProfile::STANDARD) ? count($formulas) : null,
                'truncated' => $settings->atLeast(MetadataProfile::FULL) && $formulaTruncated,
                'calculation_chain_present' => in_array('xl/calcChain.xml', $entries, true),
                'warnings' => $settings->atLeast(MetadataProfile::STANDARD)
                    ? ['Formulas were read as text and were not calculated or executed.']
                    : ['Quick profile reports calculation settings but does not scan worksheet formulas.'],
            ],
            'print_settings' => [
                'count' => count($printSettings),
                'items' => $settings->atLeast(MetadataProfile::STANDARD) ? $printItems : [],
                'truncated' => $printTruncated,
            ],
            'validation' => [
                'count' => count($validations) + count($conditionalFormatting),
                'items' => $settings->atLeast(MetadataProfile::FULL) ? $validationItems : [],
                'truncated' => $settings->atLeast(MetadataProfile::FULL) && $validationTruncated,
                'data_validation_count' => count($validations),
                'conditional_formatting_count' => count($conditionalFormatting),
            ],
            'pivot_metadata' => [
                'count' => count($pivotParts),
                'items' => $settings->atLeast(MetadataProfile::STANDARD) ? $pivotItems : [],
                'truncated' => $pivotTruncated,
                'pivot_table_part_count' => count($this->paths($pivotParts, '#^xl/pivotTables/#i')),
                'pivot_cache_part_count' => count($this->paths($pivotParts, '#^xl/pivotCache/#i')),
            ],
            'xml_metadata' => [
                'custom_xml_part_count' => count($xmlParts),
                'items' => $settings->includePackageParts() ? $xmlItems : [],
                'count' => count($xmlParts),
                'truncated' => $settings->includePackageParts() && $xmlTruncated,
                'relationships' => $relationshipItems,
                'relationship_count' => $relationshipScan['count'],
                'relationships_truncated' => $relationshipScan['truncated'],
            ],
            'signature_parts' => $signatureParts,
            'states' => [
                'macros' => $macroParts === [] ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL,
                'named_objects' => $namedObjects === [] ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL,
                'links' => $settings->atLeast(MetadataProfile::FULL)
                    ? MetadataSectionState::AVAILABLE
                    : ($settings->atLeast(MetadataProfile::STANDARD)
                        && $externalLinks === []
                        && $connections === []
                        && $queryTables === []
                        && $powerQuery === []
                        && $hyperlinks === []
                            ? MetadataSectionState::AVAILABLE
                            : MetadataSectionState::PARTIAL),
                'hidden_content' => $settings->atLeast(MetadataProfile::STANDARD) ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL,
                'comments_notes' => $settings->atLeast(MetadataProfile::FULL)
                    ? MetadataSectionState::AVAILABLE
                    : ($commentParts === [] ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL),
                'embedded_objects' => $settings->atLeast(MetadataProfile::FULL)
                    ? MetadataSectionState::AVAILABLE
                    : (($mediaParts === [] && $oleParts === []) ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL),
                'calculation' => $settings->atLeast(MetadataProfile::STANDARD) ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL,
                'print_settings' => $settings->atLeast(MetadataProfile::STANDARD) ? MetadataSectionState::AVAILABLE : MetadataSectionState::NOT_SCANNED,
                'validation' => $settings->atLeast(MetadataProfile::STANDARD) ? MetadataSectionState::AVAILABLE : MetadataSectionState::NOT_SCANNED,
                'pivot_metadata' => $pivotParts === [] ? MetadataSectionState::AVAILABLE : MetadataSectionState::PARTIAL,
            ],
            'counts' => [
                'tables' => count($this->paths($entries, '#^xl/tables/[^/]+\.xml$#i')),
                'charts' => count($this->paths($entries, '#^xl/charts/[^/]+\.xml$#i')),
                'pivot_tables' => count($this->paths($entries, '#^xl/pivotTables/[^/]+\.xml$#i')),
                'slicers' => count($this->paths($entries, '#^xl/slicers/[^/]+\.xml$#i')),
                'hidden_rows' => $hiddenRowCount,
                'hidden_columns' => $hiddenColumnCount,
                'formulas' => count($formulas),
                'comments' => $settings->atLeast(MetadataProfile::FULL) ? count($comments) : count($commentParts),
                'hyperlinks' => count($hyperlinks),
                'images' => $settings->atLeast(MetadataProfile::FULL) ? count($images) : count($mediaParts),
            ],
            'warnings' => $warnings,
        ];
    }

    /** @return array<string,mixed> */
    private function securityInformation(
        string $path,
        array $sheets,
        MetadataOptions $settings,
        bool $encrypted,
        ?string $encryptionMode,
        array $featureData
    ): array {
        $worksheetProtection = [];
        $workbook = [];
        $workbookProtected = false;
        if ($sheets !== []) {
            foreach ($sheets as $sheet) {
                try {
                    $protection = (new XlsxProtectionReader($this->resolver))->read($path, (string) $sheet['name']);
                } catch (\Throwable $error) {
                    return [
                        'state' => MetadataSectionState::PARTIAL,
                        'data' => [
                            'file_encrypted' => $encrypted,
                            'encryption_mode' => $encryptionMode,
                            'password_required' => false,
                            'count' => 0,
                            'warnings' => ['Protection metadata could not be completely read: ' . $error->getMessage()],
                        ],
                    ];
                }
                $workbook = (array) ($protection['workbook'] ?? $workbook);
                $workbookProtected = $workbookProtected || (bool) ($protection['workbook_protected'] ?? false);
                if ((bool) ($protection['worksheet_protected'] ?? false)) {
                    $worksheetProtection[] = [
                        'sheet' => (string) $sheet['name'],
                        'settings' => (array) ($protection['worksheet'] ?? []),
                    ];
                }
                if (!$settings->atLeast(MetadataProfile::STANDARD)) {
                    break;
                }
            }
        }

        $signatureParts = (array) ($featureData['signature_parts'] ?? []);

        return [
            'state' => MetadataSectionState::AVAILABLE,
            'data' => [
                'file_encrypted' => $encrypted,
                'encryption_mode' => $encryptionMode,
                'password_required' => false,
                'workbook_protected' => $workbookProtected,
                'workbook' => $workbook,
                'worksheet_protected_count' => count($worksheetProtection),
                'worksheets' => $worksheetProtection,
                'digital_signature_present' => $signatureParts !== [],
                'signature_parts' => array_map(static fn(string $path): array => ['path' => $path], $signatureParts),
                'signature_verification_state' => $signatureParts === [] ? 'not_present' : 'not_verified',
                'count' => ($encrypted ? 1 : 0) + ($workbookProtected ? 1 : 0) + count($worksheetProtection) + count($signatureParts),
                'items' => array_merge($worksheetProtection, array_map(static fn(string $path): array => ['type' => 'digital_signature', 'path' => $path], $signatureParts)),
                'warnings' => $signatureParts === [] ? [] : ['Digital signatures are inventoried but cryptographic verification is not yet implemented.'],
            ],
        ];
    }

    /** @return array<string,mixed> */
    private function scanWorksheetXml(string $xml, string $sheetName, MetadataOptions $settings): array
    {
        $hiddenItems = [];
        $sheetView = $this->firstTagAttributes($xml, 'sheetView');
        $tabSelected = $this->truthy($sheetView['tabSelected'] ?? null);
        $hiddenRows = 0;
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?row\b([^>]*)>/i', $xml, $rows, PREG_SET_ORDER);
        foreach ($rows as $row) {
            $attrs = $this->attributes($row[1]);
            if ($this->truthy($attrs['hidden'] ?? null)) {
                $hiddenRows++;
                if ($settings->atLeast(MetadataProfile::FULL)) {
                    $hiddenItems[] = ['type' => 'hidden_row', 'sheet' => $sheetName, 'row' => isset($attrs['r']) ? (int) $attrs['r'] : null];
                }
            }
        }

        $hiddenColumns = 0;
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?col\b([^>]*)\/?\s*>/i', $xml, $columns, PREG_SET_ORDER);
        foreach ($columns as $column) {
            $attrs = $this->attributes($column[1]);
            if ($this->truthy($attrs['hidden'] ?? null)) {
                $hiddenColumns++;
                if ($settings->atLeast(MetadataProfile::FULL)) {
                    $hiddenItems[] = [
                        'type' => 'hidden_column_range',
                        'sheet' => $sheetName,
                        'minimum_column' => isset($attrs['min']) ? (int) $attrs['min'] : null,
                        'maximum_column' => isset($attrs['max']) ? (int) $attrs['max'] : null,
                    ];
                }
            }
        }

        $print = [];
        foreach (['pageMargins', 'pageSetup', 'printOptions', 'sheetPr'] as $tag) {
            $attrs = $this->firstTagAttributes($xml, $tag);
            if ($attrs !== []) {
                $print[] = ['type' => $tag, 'sheet' => $sheetName, 'settings' => $this->scalarAttributes($attrs)];
            }
        }
        $headerFooter = $this->tagBlock($xml, 'headerFooter');
        if ($headerFooter !== null) {
            $print[] = [
                'type' => 'header_footer',
                'sheet' => $sheetName,
                'odd_header' => $this->tagText($headerFooter, 'oddHeader'),
                'odd_footer' => $this->tagText($headerFooter, 'oddFooter'),
                'even_header' => $this->tagText($headerFooter, 'evenHeader'),
                'even_footer' => $this->tagText($headerFooter, 'evenFooter'),
                'first_header' => $this->tagText($headerFooter, 'firstHeader'),
                'first_footer' => $this->tagText($headerFooter, 'firstFooter'),
            ];
        }

        $validations = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dataValidation\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dataValidation\s*>|<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dataValidation\b([^>]*)\/>/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes(($match[1] ?? '') !== '' ? $match[1] : ($match[3] ?? ''));
            $body = (string) ($match[2] ?? '');
            $validations[] = [
                'type' => 'data_validation',
                'sheet' => $sheetName,
                'range' => $attrs['sqref'] ?? null,
                'validation_type' => $attrs['type'] ?? 'none',
                'operator' => $attrs['operator'] ?? null,
                'allow_blank' => $this->truthy($attrs['allowBlank'] ?? null),
                'formula1' => $this->tagText($body, 'formula1'),
                'formula2' => $this->tagText($body, 'formula2'),
                'settings' => $this->scalarAttributes($attrs),
            ];
        }

        $conditional = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?conditionalFormatting\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?conditionalFormatting\s*>/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cfRule\b([^>]*)>/i', $match[2], $rules, PREG_SET_ORDER);
            $conditional[] = [
                'type' => 'conditional_formatting',
                'sheet' => $sheetName,
                'range' => $attrs['sqref'] ?? null,
                'rule_count' => count($rules),
                'rules' => $settings->atLeast(MetadataProfile::FULL)
                    ? array_map(fn(array $rule): array => $this->scalarAttributes($this->attributes($rule[1])), $rules)
                    : [],
            ];
        }

        $hyperlinks = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?hyperlink\b([^>]*)\/?\s*>/i', $xml, $hyperlinkMatches, PREG_SET_ORDER);
        foreach ($hyperlinkMatches as $match) {
            $attrs = $this->attributes($match[1]);
            $hyperlinks[] = [
                'sheet' => $sheetName,
                'cell' => $attrs['ref'] ?? null,
                'location' => $attrs['location'] ?? null,
                'display' => $attrs['display'] ?? null,
                'tooltip' => $attrs['tooltip'] ?? null,
                'relationship_id' => $attrs['id'] ?? null,
                'target' => null,
                'type' => isset($attrs['location']) ? 'internal' : 'relationship',
            ];
        }

        $formulas = [];
        if ($settings->atLeast(MetadataProfile::FULL)) {
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\s*>/is', $xml, $cells, PREG_SET_ORDER);
            foreach ($cells as $cell) {
                $formula = $this->tagText($cell[2], 'f');
                if ($formula === null) {
                    continue;
                }
                $attrs = $this->attributes($cell[1]);
                $formulas[] = [
                    'sheet' => $sheetName,
                    'cell' => $attrs['r'] ?? null,
                    'formula' => $formula,
                    'formula_type' => $this->firstTagAttributes($cell[2], 'f')['t'] ?? 'normal',
                ];
            }
        } else {
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f\b/i', $xml, $formulaMatches);
            for ($i = 0, $count = count($formulaMatches[0] ?? []); $i < $count; $i++) {
                $formulas[] = ['sheet' => $sheetName];
            }
        }

        return [
            'tab_selected' => $tabSelected,
            'hidden_items' => $hiddenItems,
            'hidden_row_count' => $hiddenRows,
            'hidden_column_count' => $hiddenColumns,
            'print_settings' => $print,
            'validations' => $validations,
            'conditional_formatting' => $conditional,
            'hyperlinks' => $hyperlinks,
            'formulas' => $formulas,
        ];
    }

    /** @return array<string,mixed> */
    private function workbookSettings(string $xml): array
    {
        $workbookPr = $this->firstTagAttributes($xml, 'workbookPr');
        $view = $this->scalarAttributes($this->firstTagAttributes($xml, 'workbookView'));
        $calcPr = $this->scalarAttributes($this->firstTagAttributes($xml, 'calcPr'));
        return [
            'date1904' => $this->truthy($workbookPr['date1904'] ?? null),
            'code_name' => $workbookPr['codeName'] ?? null,
            'view' => $view,
            'calculation' => $calcPr,
        ];
    }

    /** @param list<array<string,mixed>> $sheets @return list<array<string,mixed>> */
    private function definedNames(string $xml, array $sheets): array
    {
        $items = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?definedName\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?definedName\s*>/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $name = (string) ($attrs['name'] ?? '');
            $localId = isset($attrs['localSheetId']) ? (int) $attrs['localSheetId'] : null;
            $scope = $localId === null ? 'workbook' : (($sheets[$localId]['name'] ?? null) ?: ('sheet_index_' . ($localId + 1)));
            $type = match ($name) {
                '_xlnm.Print_Area' => 'print_area',
                '_xlnm.Print_Titles' => 'print_titles',
                default => 'defined_name',
            };
            $items[] = [
                'type' => $type,
                'name' => $name,
                'scope' => $scope,
                'formula_or_reference' => $this->decodeText($match[2]),
                'hidden' => $this->truthy($attrs['hidden'] ?? null),
                'function' => $this->truthy($attrs['function'] ?? null),
                'vb_procedure' => $this->truthy($attrs['vbProcedure'] ?? null),
            ];
        }
        return $items;
    }

    /** @return array<string,mixed> */
    private function readCoreProperties(string $xml): array
    {
        $mapping = [
            'title' => 'title',
            'subject' => 'subject',
            'creator' => 'creator',
            'keywords' => 'keywords',
            'description' => 'description',
            'lastModifiedBy' => 'last_modified_by',
            'revision' => 'revision',
            'created' => 'created',
            'modified' => 'modified',
            'lastPrinted' => 'last_printed',
            'category' => 'category',
            'contentStatus' => 'content_status',
            'identifier' => 'identifier',
            'language' => 'language',
            'version' => 'version',
        ];
        $result = [];
        foreach ($mapping as $tag => $key) {
            $result[$key] = $this->tagText($xml, $tag);
        }
        return $result;
    }

    /** @return array<string,mixed> */
    private function readApplicationProperties(string $xml): array
    {
        return [
            'application' => $this->tagText($xml, 'Application'),
            'application_version' => $this->tagText($xml, 'AppVersion'),
            'company' => $this->tagText($xml, 'Company'),
            'manager' => $this->tagText($xml, 'Manager'),
            'operating_system_hint' => $this->tagText($xml, 'OperatingSystem'),
            'document_security' => $this->toIntOrNull($this->tagText($xml, 'DocSecurity')),
            'scale_crop' => $this->toBoolOrNull($this->tagText($xml, 'ScaleCrop')),
            'links_up_to_date' => $this->toBoolOrNull($this->tagText($xml, 'LinksUpToDate')),
            'shared_document' => $this->toBoolOrNull($this->tagText($xml, 'SharedDoc')),
            'hyperlinks_changed' => $this->toBoolOrNull($this->tagText($xml, 'HyperlinksChanged')),
            'total_time_minutes' => $this->toIntOrNull($this->tagText($xml, 'TotalTime')),
        ];
    }

    /** @return list<array<string,mixed>> */
    private function readCustomProperties(string $xml): array
    {
        $items = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?property\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?property\s*>/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $body = $match[2];
            if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?([A-Za-z0-9_]+)\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?\1\s*>/is', $body, $valueMatch) === 1) {
                $nativeType = strtolower($valueMatch[1]);
                $raw = $this->decodeText($valueMatch[2]);
            } elseif (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?([A-Za-z0-9_]+)\b[^>]*\/>/is', $body, $valueMatch) === 1) {
                $nativeType = strtolower($valueMatch[1]);
                $raw = '';
            } else {
                $nativeType = 'unknown';
                $raw = '';
            }
            [$type, $value] = $this->customValue($nativeType, $raw);
            $items[] = [
                'name' => (string) ($attrs['name'] ?? ''),
                'type' => $type,
                'value' => $value,
                'native_type' => $nativeType,
                'property_id' => isset($attrs['pid']) ? (int) $attrs['pid'] : null,
                'linked' => isset($attrs['linkTarget']),
                'link_target' => $attrs['linkTarget'] ?? null,
            ];
        }
        return $items;
    }

    /** @return array{string,mixed} */
    private function customValue(string $type, string $raw): array
    {
        return match ($type) {
            'i1', 'i2', 'i4', 'i8', 'int', 'ui1', 'ui2', 'ui4', 'ui8', 'uint' => ['integer', $this->integerMetadataValue($raw)],
            'r4', 'r8', 'decimal', 'cy' => ['float', (float) $raw],
            'bool' => ['boolean', $this->truthy($raw)],
            'filetime', 'date' => ['datetime', $raw],
            'empty', 'null' => ['null', null],
            'lpstr', 'lpwstr', 'bstr' => ['string', $raw],
            'blob', 'oblob', 'stream', 'storage' => ['binary/opaque', null],
            default => ['opaque', null],
        };
    }

    /** @return list<array<string,string>> */
    private function contentTypes(string $xml): array
    {
        $items = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?(?:Default|Override)\b([^>]*)\/?\s*>/i', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $items[] = array_filter([
                'extension' => $attrs['Extension'] ?? null,
                'part_name' => $attrs['PartName'] ?? null,
                'content_type' => $attrs['ContentType'] ?? null,
            ], static fn(mixed $value): bool => $value !== null);
        }
        return $items;
    }

    /** @param list<string> $entries @return array{items:list<array<string,mixed>>,count:int,truncated:bool} */
    private function relationships(ZipArchive $zip, array $entries, int $maxItems): array
    {
        $items = [];
        $count = 0;
        foreach ($entries as $entry) {
            if (!str_ends_with(strtolower($entry), '.rels')) {
                continue;
            }
            $xml = (string) ($zip->getFromName($entry) ?: '');
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Relationship\b([^>]*)\/?\s*>/i', $xml, $matches, PREG_SET_ORDER);
            foreach ($matches as $match) {
                $attrs = $this->attributes($match[1]);
                $count++;
                if (count($items) >= $maxItems) {
                    continue;
                }
                $items[] = [
                    'source_part' => $entry,
                    'id' => $attrs['Id'] ?? null,
                    'type' => $attrs['Type'] ?? null,
                    'target' => $attrs['Target'] ?? null,
                    'target_mode' => $attrs['TargetMode'] ?? 'Internal',
                ];
            }
        }
        return [
            'items' => $items,
            'count' => $count,
            'truncated' => $count > count($items),
        ];
    }

    /** @return array<string,mixed> */
    private function firstVisibleSheet(array $sheets): array
    {
        foreach ($sheets as $sheet) {
            if (($sheet['state'] ?? 'visible') === 'visible') {
                return ['index' => $sheet['index'], 'name' => $sheet['name']];
            }
        }
        return [];
    }

    /** @return list<string> */
    private function keywords(string $keywords): array
    {
        if (trim($keywords) === '') {
            return [];
        }
        return array_values(array_unique(array_filter(array_map('trim', preg_split('/[,;]+/', $keywords) ?: []), static fn(string $value): bool => $value !== '')));
    }

    /** @param list<string> $entries @return list<string> */
    private function paths(array $entries, string $pattern): array
    {
        return array_values(array_filter($entries, static fn(string $entry): bool => preg_match($pattern, $entry) === 1));
    }

    /** @param list<string> $entries */
    private function firstMatching(array $entries, string $pattern): ?string
    {
        foreach ($entries as $entry) {
            if (preg_match($pattern, $entry) === 1) {
                return $entry;
            }
        }
        return null;
    }

    /** @param list<mixed> $items @return array{list<mixed>,bool} */
    private function limit(array $items, int $maximum): array
    {
        if (count($items) <= $maximum) {
            return [array_values($items), false];
        }
        return [array_slice(array_values($items), 0, $maximum), true];
    }

    /** @param array<string,mixed> $values */
    private function nonEmptyCount(array $values): int
    {
        return count(array_filter($values, static fn(mixed $value): bool => $value !== null && $value !== '' && $value !== []));
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

    /** @return array<string,string> */
    private function firstTagAttributes(string $xml, string $tag): array
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b([^>]*)\/?\s*>/is', $xml, $match) !== 1) {
            return [];
        }
        return $this->attributes($match[1]);
    }

    private function tagText(string $xml, string $tag): ?string
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\s*>/is', $xml, $match) !== 1) {
            return null;
        }
        return $this->decodeText($match[1]);
    }

    private function tagBlock(string $xml, string $tag): ?string
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . preg_quote($tag, '/') . '\s*>/is', $xml, $match) !== 1) {
            return null;
        }
        return $match[0];
    }

    private function decodeText(string $value): string
    {
        $value = preg_replace('/<!\[CDATA\[(.*?)\]\]>/s', '$1', $value) ?? $value;
        $value = strip_tags($value);
        return trim(html_entity_decode($value, ENT_QUOTES | ENT_XML1, 'UTF-8'));
    }

    /** @param array<string,string> $attributes @return array<string,mixed> */
    private function scalarAttributes(array $attributes): array
    {
        $result = [];
        foreach ($attributes as $name => $value) {
            if (in_array(strtolower($value), ['true', 'false'], true)) {
                $result[$name] = strtolower($value) === 'true';
            } elseif (is_numeric($value)) {
                $result[$name] = str_contains($value, '.') ? (float) $value : (int) $value;
            } else {
                $result[$name] = $value;
            }
        }
        return $result;
    }

    private function truthy(mixed $value): bool
    {
        return in_array(strtolower(trim((string) $value)), ['1', 'true', 'yes', 'on'], true);
    }

    private function toBoolOrNull(?string $value): ?bool
    {
        return $value === null || trim($value) === '' ? null : $this->truthy($value);
    }

    private function toIntOrNull(?string $value): ?int
    {
        return $value === null || trim($value) === '' ? null : (int) $value;
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

    private function variantFromExtension(string $extension): string
    {
        return in_array($extension, ['xlsx', 'xlsm', 'xltx', 'xltm'], true) ? $extension : 'xlsx';
    }

    private function mimeForVariant(string $variant): string
    {
        return match ($variant) {
            'xlsm' => 'application/vnd.ms-excel.sheet.macroEnabled.12',
            'xltx' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
            'xltm' => 'application/vnd.ms-excel.template.macroEnabled.12',
            default => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        };
    }

    private function openZip(string $path): ZipArchive
    {
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            throw new MnbExcelException('Unable to open XLSX package.');
        }
        return $zip;
    }
}
