<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Security\FormulaGuard;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use ZipArchive;

final class XlsxWriter
{
    /** @var array<int, array<string,mixed>> */
    private array $styleEntries = [];

    /** @var array<string,int> */
    private array $styleIds = [];

    public function write(WorkbookData $workbook, string $path): void
    {
        AtomicFileWriter::writeViaTemp(
            $path,
            function (string $tmp) use ($workbook): void {
                $this->writePackage($workbook, $tmp);
            },
            function (string $tmp) use ($workbook): void {
                $this->validateWrittenWorkbook($workbook, $tmp);
            }
        );
    }

    private function writePackage(WorkbookData $workbook, string $path): void
    {
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode(
                'ext-zip is required to write XLSX files.',
                ErrorCode::EXTENSION_MISSING,
                ['extension' => 'zip']
            );
        }

        $this->buildStyleRegistry($workbook);
        $imagePlan = $this->buildImagePlan($workbook);
        $preservedPackage = $this->loadPreservedPackage($workbook, $imagePlan);

        $zip = new ZipArchive();
        $openResult = $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        if ($openResult !== true) {
            throw MnbExcelException::withCode(
                'Unable to create XLSX file: ' . $path,
                ErrorCode::XLSX_WRITE_FAILED,
                ['path' => $path, 'zip_status' => $openResult]
            );
        }

        $writtenEntries = [];
        $addXml = function (string $entry, string $xml) use ($zip, &$writtenEntries): void {
            $this->addStringToZip($zip, $entry, $xml);
            $writtenEntries[$entry] = true;
        };

        try {
            $addXml('[Content_Types].xml', $this->contentTypes($workbook, $imagePlan, $preservedPackage));
            $addXml('_rels/.rels', $this->rootRels());
            $addXml('docProps/core.xml', $this->coreProps($workbook->metadata));
            $addXml('docProps/app.xml', $this->appProps($workbook->metadata));
            $addXml('xl/workbook.xml', $this->workbookXml($workbook));
            $addXml('xl/_rels/workbook.xml.rels', $this->workbookRels(count($workbook->sheets)));
            $addXml('xl/styles.xml', $this->stylesXml());

            foreach ($workbook->sheets as $index => $sheet) {
                $sheetNumber = $index + 1;
                $hasImages = isset($imagePlan[$sheetNumber]) && $imagePlan[$sheetNumber] !== [];
                $hasHyperlinks = $sheet->hyperlinks !== [];
                $hasComments = $sheet->comments !== [];
                $hasGeneratedSheetRels = $hasImages || $hasHyperlinks || $hasComments;
                $preservedSheet = (!$hasGeneratedSheetRels && is_array($preservedPackage)) ? ($preservedPackage['sheets'][$sheetNumber] ?? null) : null;
                $addXml('xl/worksheets/sheet' . $sheetNumber . '.xml', $this->worksheetXml($sheet, $hasImages, is_array($preservedSheet) ? $preservedSheet : null));

                if ($hasGeneratedSheetRels) {
                    $addXml('xl/worksheets/_rels/sheet' . $sheetNumber . '.xml.rels', $this->sheetRelationships($sheetNumber, $sheet, $hasImages));
                } elseif (is_array($preservedSheet) && ($preservedSheet['rels_xml'] ?? '') !== '') {
                    $entry = 'xl/worksheets/_rels/sheet' . $sheetNumber . '.xml.rels';
                    $addXml($entry, (string) $preservedSheet['rels_xml']);
                }

                if ($hasImages) {
                    $addXml('xl/drawings/drawing' . $sheetNumber . '.xml', $this->drawingXml($imagePlan[$sheetNumber]));
                    $addXml('xl/drawings/_rels/drawing' . $sheetNumber . '.xml.rels', $this->drawingRels($imagePlan[$sheetNumber]));

                    foreach ($imagePlan[$sheetNumber] as $image) {
                        $entry = 'xl/media/' . $image['mediaName'];
                        $this->addFileToZip($zip, $image['path'], $entry);
                        $writtenEntries[$entry] = true;
                    }
                }

                if ($hasComments) {
                    $addXml('xl/comments' . $sheetNumber . '.xml', $this->commentsXml($sheet));
                    $addXml('xl/drawings/vmlDrawing' . $sheetNumber . '.vml', $this->commentsVmlDrawingXml($sheet));
                }
            }

            if (is_array($preservedPackage)) {
                $this->copyPreservedPackageEntries($zip, $preservedPackage, $writtenEntries);
            }
        } catch (\Throwable $e) {
            $zip->close();
            throw $e;
        }

        $closed = $zip->close();
        if ($closed !== true) {
            throw MnbExcelException::withCode(
                'Unable to finalize XLSX zip package: ' . $path,
                ErrorCode::XLSX_ZIP_CLOSE_FAILED,
                ['path' => $path]
            );
        }
    }

    private function addStringToZip(ZipArchive $zip, string $entry, string $contents): void
    {
        if ($zip->addFromString($entry, $contents) !== true) {
            throw MnbExcelException::withCode(
                'Unable to add XLSX zip entry: ' . $entry,
                ErrorCode::XLSX_ZIP_ENTRY_FAILED,
                ['entry' => $entry]
            );
        }
    }

    private function addFileToZip(ZipArchive $zip, string $sourcePath, string $entry): void
    {
        if (!is_file($sourcePath)) {
            throw MnbExcelException::withCode(
                'XLSX media file not found: ' . $sourcePath,
                ErrorCode::FILE_NOT_FOUND,
                ['path' => $sourcePath, 'entry' => $entry]
            );
        }

        if ($zip->addFile($sourcePath, $entry) !== true) {
            throw MnbExcelException::withCode(
                'Unable to add file to XLSX zip entry: ' . $entry,
                ErrorCode::XLSX_ZIP_ENTRY_FAILED,
                ['path' => $sourcePath, 'entry' => $entry]
            );
        }
    }

    private function validateWrittenWorkbook(WorkbookData $workbook, string $path): void
    {
        $settings = $workbook->metadata['_mnb_xlsx_integrity_validation'] ?? ['enabled' => true];
        if (is_array($settings) && array_key_exists('enabled', $settings) && !$settings['enabled']) {
            return;
        }

        $options = is_array($settings) ? $settings : [];
        try {
            (new XlsxIntegrityValidator())->assertValid($path, $options);
        } catch (MnbExcelException $e) {
            if (is_file($path)) {
                @unlink($path);
            }
            throw $e;
        }
    }

    private function buildStyleRegistry(WorkbookData $workbook): void
    {
        $this->styleEntries = [];
        $this->styleIds = [];

        foreach ($workbook->sheets as $sheet) {
            foreach ($sheet->rows as $rowIndex => $row) {
                $r = $rowIndex + 1;
                foreach (array_values($row) as $columnIndex => $_value) {
                    $c = $columnIndex + 1;
                    $style = $this->effectiveCellStyle($sheet, $r, $c, Coordinate::columnIndexToName($c) . $r);
                    $this->styleIdFor($style);
                }
            }
        }
    }

    /** @return array<string,mixed> */
    private function effectiveCellStyle(WorksheetData $sheet, int $row, int $column, string $cellRef): array
    {
        $style = [];

        if ($sheet->headerRowIndex !== null && $row === $sheet->headerRowIndex && $sheet->hasHeader) {
            $style = $this->mergeStyles($style, $this->resolveStyle($sheet, $sheet->headerStyle));
        } else {
            if (isset($sheet->columnStyles[$column])) {
                $style = $this->mergeStyles($style, $this->resolveStyle($sheet, $sheet->columnStyles[$column]));
            }
            if (isset($sheet->rowStyles[$row])) {
                $style = $this->mergeStyles($style, $this->resolveStyle($sheet, $sheet->rowStyles[$row]));
            }
            foreach ($sheet->rangeStyles as $range => $rangeStyle) {
                if ($this->cellInRange($cellRef, (string) $range)) {
                    $style = $this->mergeStyles($style, $this->resolveStyle($sheet, $rangeStyle));
                }
            }
            if (isset($sheet->cellStyles[$cellRef])) {
                $style = $this->mergeStyles($style, $this->resolveStyle($sheet, $sheet->cellStyles[$cellRef]));
            }
        }

        return $style;
    }

    /** @param string|array<string,mixed>|mixed $style */
    private function resolveStyle(WorksheetData $sheet, mixed $style): array
    {
        if (is_string($style)) {
            return $sheet->namedStyles[$style] ?? [];
        }
        if (is_array($style)) {
            return $style;
        }
        return [];
    }

    /** @param array<string,mixed> $style */
    private function styleIdFor(array $style): int
    {
        if ($style === []) {
            return 0;
        }

        $normalized = $this->normalizeStyle($style);
        if ($normalized === []) {
            return 0;
        }

        $key = json_encode($normalized, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
        if ($key === false) {
            $key = serialize($normalized);
        }

        if (!isset($this->styleIds[$key])) {
            $this->styleIds[$key] = count($this->styleEntries) + 1;
            $this->styleEntries[$this->styleIds[$key]] = $normalized;
        }

        return $this->styleIds[$key];
    }

    /** @param array<string,mixed> $style */
    private function normalizeStyle(array $style): array
    {
        $normalized = [];

        foreach ($style as $key => $value) {
            $styleKey = strtolower((string) $key);
            if ($styleKey === 'number_format') {
                $styleKey = 'format';
            }
            if ($styleKey === 'wraptext') {
                $styleKey = 'wrap_text';
            }
            if (is_array($value)) {
                $normalized[$styleKey] = $this->normalizeStyle($value);
            } else {
                $normalized[$styleKey] = $value;
            }
        }

        ksort($normalized);
        return $normalized;
    }

    /** @param array<string,mixed> $base @param array<string,mixed> $override */
    private function mergeStyles(array $base, array $override): array
    {
        foreach ($override as $key => $value) {
            if (isset($base[$key]) && is_array($base[$key]) && is_array($value)) {
                $base[$key] = $this->mergeStyles($base[$key], $value);
            } else {
                $base[$key] = $value;
            }
        }

        return $base;
    }

    /**
     * @return array<int, list<array{path:string,mediaName:string,extension:string,cell:string,width:int,height:int,name:string}>>
     */
    private function buildImagePlan(WorkbookData $workbook): array
    {
        $plan = [];
        $mediaIndex = 1;

        foreach ($workbook->sheets as $sheetIndex => $sheet) {
            $sheetNumber = $sheetIndex + 1;
            foreach ($sheet->images as $imageIndex => $image) {
                $extension = $this->imageExtension($image['path']);
                $size = @getimagesize($image['path']);
                $width = isset($image['width']) ? (int) $image['width'] : (($size !== false && isset($size[0])) ? (int) $size[0] : 120);
                $height = isset($image['height']) ? (int) $image['height'] : (($size !== false && isset($size[1])) ? (int) $size[1] : 80);

                $plan[$sheetNumber][] = [
                    'path' => $image['path'],
                    'mediaName' => 'image' . $mediaIndex . '.' . $extension,
                    'extension' => $extension,
                    'cell' => $image['cell'],
                    'width' => max(1, $width),
                    'height' => max(1, $height),
                    'name' => $image['name'] ?? ('Image ' . ($imageIndex + 1)),
                ];
                $mediaIndex++;
            }
        }

        return $plan;
    }

    /** @param array<int, list<array{extension:string}>> $imagePlan @param array<string,mixed>|null $preservedPackage */
    private function contentTypes(WorkbookData $workbook, array $imagePlan, ?array $preservedPackage = null): string
    {
        $sheetCount = count($workbook->sheets);

        $defaultContentTypes = [
            'rels' => 'application/vnd.openxmlformats-package.relationships+xml',
            'xml' => 'application/xml',
        ];
        foreach ($workbook->sheets as $sheet) {
            if ($sheet->comments !== []) {
                $defaultContentTypes['vml'] = 'application/vnd.openxmlformats-officedocument.vmlDrawing';
                break;
            }
        }
        foreach ($imagePlan as $images) {
            foreach ($images as $image) {
                $defaultContentTypes[$image['extension']] = $this->imageContentType($image['extension']);
            }
        }

        if (is_array($preservedPackage)) {
            foreach ((array) ($preservedPackage['content_types']['defaults'] ?? []) as $extension => $contentType) {
                $extension = strtolower((string) $extension);
                if ($extension !== '' && !isset($defaultContentTypes[$extension])) {
                    $defaultContentTypes[$extension] = (string) $contentType;
                }
            }
        }

        $defaults = '';
        foreach ($defaultContentTypes as $extension => $contentType) {
            $defaults .= '<Default Extension="' . $this->esc((string) $extension) . '" ContentType="' . $this->esc((string) $contentType) . '"/>';
        }

        $overrides = [
            '/docProps/core.xml' => 'application/vnd.openxmlformats-package.core-properties+xml',
            '/docProps/app.xml' => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
            '/xl/workbook.xml' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
            '/xl/styles.xml' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
        ];
        for ($i = 1; $i <= $sheetCount; $i++) {
            $overrides['/xl/worksheets/sheet' . $i . '.xml'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
            if (isset($imagePlan[$i]) && $imagePlan[$i] !== []) {
                $overrides['/xl/drawings/drawing' . $i . '.xml'] = 'application/vnd.openxmlformats-officedocument.drawing+xml';
            }
            if (($workbook->sheets[$i - 1]->comments ?? []) !== []) {
                $overrides['/xl/comments' . $i . '.xml'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
            }
        }
        if (is_array($preservedPackage)) {
            foreach ((array) ($preservedPackage['content_types']['overrides'] ?? []) as $partName => $contentType) {
                $partName = '/' . ltrim((string) $partName, '/');
                if (!isset($overrides[$partName]) && $this->preservedOverrideIsUseful($partName, $preservedPackage)) {
                    $overrides[$partName] = (string) $contentType;
                }
            }
        }

        $overrideXml = '';
        foreach ($overrides as $partName => $contentType) {
            $overrideXml .= '<Override PartName="' . $this->esc($partName) . '" ContentType="' . $this->esc($contentType) . '"/>';
        }

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . $defaults
            . $overrideXml
            . '</Types>';
    }

    /**
     * @param array<int, list<array{path:string,mediaName:string,extension:string,cell:string,width:int,height:int,name:string}>> $imagePlan
     * @return array<string,mixed>|null
     */
    private function loadPreservedPackage(WorkbookData $workbook, array $imagePlan): ?array
    {
        $settings = $workbook->metadata['_mnb_preserve_xlsx_package'] ?? null;
        if (!is_array($settings) || !isset($settings['path']) || !is_string($settings['path'])) {
            return null;
        }

        $realPath = realpath($settings['path']);
        if ($realPath === false || !is_file($realPath)) {
            throw MnbExcelException::withCode('Advanced-object template XLSX not found: ' . (string) $settings['path'], ErrorCode::FILE_NOT_FOUND, ['path' => (string) $settings['path']]);
        }

        $zip = new ZipArchive();
        if ($zip->open($realPath) !== true) {
            throw MnbExcelException::withCode('Unable to open advanced-object template XLSX: ' . $realPath, ErrorCode::XLSX_INVALID, ['path' => $realPath]);
        }

        $sheets = [];
        try {
            $sourceSheets = (new XlsxWorkbookResolver())->sheets($realPath);
        } catch (\Throwable) {
            $sourceSheets = [];
        }

        foreach ($workbook->sheets as $index => $_sheet) {
            $sheetNumber = $index + 1;
            $sourceSheet = $sourceSheets[$index] ?? null;
            if (!is_array($sourceSheet) || !($sourceSheet['exists'] ?? false)) {
                continue;
            }

            $sourceSheetPath = (string) $sourceSheet['path'];
            $sheetXml = $zip->getFromName($sourceSheetPath);
            if ($sheetXml === false || trim($sheetXml) === '') {
                continue;
            }

            $relsPath = $this->relsPathForPart($sourceSheetPath);
            $relsXml = $zip->getFromName($relsPath);
            $elements = ((bool) ($settings['preserve_sheet_elements'] ?? true))
                ? $this->extractPreservedSheetElements($sheetXml)
                : [];

            $sheets[$sheetNumber] = [
                'source_sheet_path' => $sourceSheetPath,
                'source_rels_path' => $relsPath,
                'rels_xml' => ((bool) ($settings['preserve_sheet_relationships'] ?? true)) && $relsXml !== false ? $relsXml : '',
                'elements' => $elements,
                'requires_relationships' => $this->elementsRequireRelationships($elements),
                'image_conflict' => isset($imagePlan[$sheetNumber]) && $imagePlan[$sheetNumber] !== [],
            ];
        }

        $contentTypesXml = $zip->getFromName('[Content_Types].xml');
        $contentTypes = $contentTypesXml !== false ? $this->parseContentTypes($contentTypesXml) : ['defaults' => [], 'overrides' => []];
        $entries = [];
        if ((bool) ($settings['copy_package_parts'] ?? true)) {
            for ($i = 0; $i < $zip->numFiles; $i++) {
                $entry = (string) $zip->getNameIndex($i);
                if ($this->shouldCopyPreservedEntry($entry)) {
                    $entries[] = $entry;
                }
            }
        }

        $zip->close();

        return [
            'path' => $realPath,
            'settings' => $settings,
            'sheets' => $sheets,
            'entries' => array_values(array_unique($entries)),
            'content_types' => $contentTypes,
        ];
    }

    /** @param array<string,mixed> $preservedPackage @param array<string,true> $writtenEntries */
    private function copyPreservedPackageEntries(ZipArchive $outputZip, array $preservedPackage, array &$writtenEntries): void
    {
        $sourcePath = (string) ($preservedPackage['path'] ?? '');
        if ($sourcePath === '') {
            return;
        }

        $sourceZip = new ZipArchive();
        if ($sourceZip->open($sourcePath) !== true) {
            throw MnbExcelException::withCode('Unable to reopen advanced-object template XLSX: ' . $sourcePath, ErrorCode::XLSX_INVALID, ['path' => $sourcePath]);
        }

        foreach ((array) ($preservedPackage['entries'] ?? []) as $entry) {
            $entry = (string) $entry;
            if ($entry === '' || isset($writtenEntries[$entry])) {
                continue;
            }
            $content = $sourceZip->getFromName($entry);
            if ($content === false) {
                continue;
            }
            $this->addStringToZip($outputZip, $entry, $content);
            $writtenEntries[$entry] = true;
        }

        $sourceZip->close();
    }

    /** @return list<string> */
    private function extractPreservedSheetElements(string $sheetXml): array
    {
        $elements = [];
        $tags = [
            'conditionalFormatting',
            'dataValidations',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'pageSetup',
            'headerFooter',
            'drawing',
            'legacyDrawing',
            'legacyDrawingHF',
            'picture',
            'oleObjects',
            'controls',
            'webPublishItems',
            'tableParts',
            'ignoredErrors',
            'smartTags',
            'extLst',
        ];

        foreach ($tags as $tag) {
            foreach ($this->extractXmlElements($sheetXml, $tag) as $element) {
                // These elements should remain behind the regenerated sheetData so comments,
                // hyperlinks, tables, drawings, validations, and related advanced objects stay linked.
                $elements[] = $element;
            }
        }

        return $elements;
    }

    /** @return list<string> */
    private function extractXmlElements(string $xml, string $tag): array
    {
        preg_match_all('/<' . preg_quote($tag, '/') . '\b[^>]*(?:\/>|>.*?<\/' . preg_quote($tag, '/') . '>)/isu', $xml, $matches);
        return $matches[0] ?? [];
    }

    /** @param list<string> $elements */
    private function elementsRequireRelationships(array $elements): bool
    {
        foreach ($elements as $element) {
            if (str_contains((string) $element, 'r:id=')) {
                return true;
            }
        }
        return false;
    }

    /** @return array{defaults:array<string,string>,overrides:array<string,string>} */
    private function parseContentTypes(string $xml): array
    {
        $defaults = [];
        preg_match_all('/<Default\b[^>]*\/>/iu', $xml, $defaultMatches);
        foreach ($defaultMatches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            if (isset($attrs['Extension'], $attrs['ContentType'])) {
                $defaults[strtolower($attrs['Extension'])] = $attrs['ContentType'];
            }
        }

        $overrides = [];
        preg_match_all('/<Override\b[^>]*\/>/iu', $xml, $overrideMatches);
        foreach ($overrideMatches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            if (isset($attrs['PartName'], $attrs['ContentType'])) {
                $overrides['/' . ltrim($attrs['PartName'], '/')] = $attrs['ContentType'];
            }
        }

        return ['defaults' => $defaults, 'overrides' => $overrides];
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

    /** @param array<string,mixed> $preservedPackage */
    private function preservedOverrideIsUseful(string $partName, array $preservedPackage): bool
    {
        $entry = ltrim($partName, '/');
        if ($this->shouldCopyPreservedEntry($entry)) {
            return true;
        }
        foreach ((array) ($preservedPackage['sheets'] ?? []) as $sheet) {
            if (!is_array($sheet)) {
                continue;
            }
            if (($sheet['source_rels_path'] ?? null) === $entry || ($sheet['source_sheet_path'] ?? null) === $entry) {
                return true;
            }
        }
        return false;
    }

    private function shouldCopyPreservedEntry(string $entry): bool
    {
        $lower = strtolower(str_replace('\\', '/', $entry));
        if ($lower === '' || str_ends_with($lower, '/')) {
            return false;
        }
        foreach ([
            'xl/drawings/',
            'xl/media/',
            'xl/comments',
            'xl/threadedcomments/',
            'xl/persons/',
            'xl/vmldrawings/',
            'xl/tables/',
            'xl/pivottables/',
            'xl/pivotcache/',
            'xl/charts/',
            'xl/embeddings/',
            'xl/ctrlprops/',
            'xl/printersettings/',
            'xl/externallinks/',
            'xl/querytables/',
            'xl/slicers/',
            'xl/timelines/',
            'xl/model/',
            'customxml/',
        ] as $prefix) {
            if (str_starts_with($lower, $prefix)) {
                return true;
            }
        }
        return false;
    }

    private function relsPathForPart(string $partPath): string
    {
        $dir = dirname($partPath);
        $base = basename($partPath);
        return ($dir === '.' ? '_rels/' : $dir . '/_rels/') . $base . '.rels';
    }

    private function rootRels(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            . '</Relationships>';
    }

    private function workbookRels(int $sheetCount): string
    {
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

        for ($i = 1; $i <= $sheetCount; $i++) {
            $rels .= '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }

        $rels .= '<Relationship Id="rId' . ($sheetCount + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        $rels .= '</Relationships>';

        return $rels;
    }

    private function workbookXml(WorkbookData $workbook): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheets>';

        foreach ($workbook->sheets as $index => $sheet) {
            $id = $index + 1;
            $xml .= '<sheet name="' . $this->esc($sheet->name) . '" sheetId="' . $id . '" r:id="rId' . $id . '"/>';
        }

        return $xml . '</sheets></workbook>';
    }

    /** @param array<string,mixed>|null $preservedSheet */
    private function worksheetXml(WorksheetData $sheet, bool $hasImages = false, ?array $preservedSheet = null): string
    {
        $rowCount = count($sheet->rows);
        $columnCount = $this->maxColumnCount($sheet->rows);
        $lastCell = $columnCount > 0 && $rowCount > 0 ? Coordinate::columnIndexToName($columnCount) . $rowCount : 'A1';
        $needsRelationships = $hasImages || $sheet->hyperlinks !== [] || $sheet->comments !== [] || (is_array($preservedSheet) && (bool) ($preservedSheet['requires_relationships'] ?? false));
        $xmlnsR = $needsRelationships ? ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : '';

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' . $xmlnsR . '>'
            . '<dimension ref="A1:' . $lastCell . '"/>';

        if ($sheet->freezeHeader) {
            $freezeRow = $sheet->headerRowIndex ?? 1;
            $topLeft = 'A' . ($freezeRow + 1);
            $xml .= '<sheetViews><sheetView workbookViewId="0"><pane ySplit="' . $freezeRow . '" topLeftCell="' . $topLeft . '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>';
        }

        $colsXml = $this->columnsXml($sheet->columnWidths);
        if ($colsXml !== '') {
            $xml .= $colsXml;
        }

        $xml .= '<sheetData>';
        foreach ($sheet->rows as $rowIndex => $row) {
            $r = $rowIndex + 1;
            $rowAttrs = ' r="' . $r . '"';
            if (isset($sheet->rowHeights[$r])) {
                $rowAttrs .= ' ht="' . $this->number((float) $sheet->rowHeights[$r]) . '" customHeight="1"';
            }

            $xml .= '<row' . $rowAttrs . '>';
            foreach (array_values($row) as $columnIndex => $value) {
                $c = $columnIndex + 1;
                $cellRef = Coordinate::columnIndexToName($c) . $r;
                $styleId = $this->styleIdFor($this->effectiveCellStyle($sheet, $r, $c, $cellRef));
                $xml .= $this->cellXml($cellRef, $value, $styleId);
            }
            $xml .= '</row>';
        }
        $xml .= '</sheetData>';

        if ($sheet->mergeCells !== []) {
            $xml .= '<mergeCells count="' . count($sheet->mergeCells) . '">';
            foreach ($sheet->mergeCells as $range) {
                $xml .= '<mergeCell ref="' . $this->esc($range) . '"/>';
            }
            $xml .= '</mergeCells>';
        }

        if ($sheet->autoFilter && $rowCount > 0 && $columnCount > 0) {
            $filterRow = $sheet->headerRowIndex ?? 1;
            $xml .= '<autoFilter ref="A' . $filterRow . ':' . Coordinate::columnIndexToName($columnCount) . $filterRow . '"/>';
        }

        if ($sheet->hyperlinks !== []) {
            $xml .= '<hyperlinks>';
            foreach (array_values($sheet->hyperlinks) as $index => $link) {
                $relationshipId = $this->hyperlinkRelationshipId($hasImages, $index);
                $xml .= '<hyperlink ref="' . $this->esc($link['cell']) . '" r:id="' . $relationshipId . '"';
                if (($link['display'] ?? '') !== '') {
                    $xml .= ' display="' . $this->esc((string) $link['display']) . '"';
                }
                if (($link['tooltip'] ?? '') !== '') {
                    $xml .= ' tooltip="' . $this->esc((string) $link['tooltip']) . '"';
                }
                $xml .= '/>';
            }
            $xml .= '</hyperlinks>';
        }

        if (is_array($preservedSheet)) {
            foreach ((array) ($preservedSheet['elements'] ?? []) as $element) {
                $xml .= (string) $element;
            }
        }

        if ($hasImages) {
            $xml .= '<drawing r:id="rId1"/>';
        }

        if ($sheet->comments !== []) {
            $xml .= '<legacyDrawing r:id="' . $this->commentsVmlRelationshipId($hasImages, count($sheet->hyperlinks)) . '"/>';
        }

        return $xml . '</worksheet>';
    }

    /** @param array<int|string, float|int> $widths */
    private function columnsXml(array $widths): string
    {
        if ($widths === []) {
            return '';
        }

        $xml = '<cols>';
        foreach ($widths as $column => $width) {
            $index = is_int($column) || ctype_digit((string) $column)
                ? (int) $column
                : Coordinate::columnNameToIndex((string) $column);
            if ($index < 1) {
                continue;
            }
            $xml .= '<col min="' . $index . '" max="' . $index . '" width="' . $this->number((float) $width) . '" customWidth="1"/>';
        }
        return $xml . '</cols>';
    }

    private function cellXml(string $cellRef, mixed $value, int $styleId): string
    {
        $style = $styleId > 0 ? ' s="' . $styleId . '"' : '';

        if ($value instanceof CellValue) {
            return $this->typedCellXml($cellRef, $value, $style);
        }

        if ($value === null || $value === '') {
            return '<c r="' . $cellRef . '"' . $style . '/>';
        }

        if (is_bool($value)) {
            return '<c r="' . $cellRef . '" t="b"' . $style . '><v>' . ($value ? '1' : '0') . '</v></c>';
        }

        if (is_int($value) || is_float($value)) {
            return '<c r="' . $cellRef . '"' . $style . '><v>' . $this->number($value) . '</v></c>';
        }

        $string = (string) $value;
        if (str_starts_with($string, '=')) {
            FormulaGuard::assertSafe(substr($string, 1));
            return '<c r="' . $cellRef . '"' . $style . '><f>' . $this->esc(substr($string, 1)) . '</f></c>';
        }

        return $this->inlineStringCellXml($cellRef, $string, $style);
    }

    private function typedCellXml(string $cellRef, CellValue $cell, string $style): string
    {
        return match ($cell->type()) {
            CellValue::TYPE_BLANK => '<c r="' . $cellRef . '"' . $style . '/>',
            CellValue::TYPE_BOOLEAN => '<c r="' . $cellRef . '" t="b"' . $style . '><v>' . ($cell->value() ? '1' : '0') . '</v></c>',
            CellValue::TYPE_NUMBER => '<c r="' . $cellRef . '"' . $style . '><v>' . $this->number($cell->value()) . '</v></c>',
            CellValue::TYPE_ERROR => '<c r="' . $cellRef . '" t="e"' . $style . '><v>' . $this->esc((string) $cell->value()) . '</v></c>',
            CellValue::TYPE_FORMULA => $this->formulaCellXml($cellRef, $cell, $style),
            CellValue::TYPE_DATE => $this->inlineStringCellXml($cellRef, (string) $cell->displayValue(), $style),
            default => $this->inlineStringCellXml($cellRef, (string) $cell->displayValue(), $style),
        };
    }

    private function formulaCellXml(string $cellRef, CellValue $cell, string $style): string
    {
        FormulaGuard::assertSafe((string) $cell->value(), $cell->options());
        $cached = $cell->cachedValue();
        if ($cached === null || $cached === '') {
            return '<c r="' . $cellRef . '"' . $style . '><f>' . $this->esc((string) $cell->value()) . '</f></c>';
        }
        if (is_bool($cached)) {
            return '<c r="' . $cellRef . '" t="b"' . $style . '><f>' . $this->esc((string) $cell->value()) . '</f><v>' . ($cached ? '1' : '0') . '</v></c>';
        }
        if (is_int($cached) || is_float($cached) || (is_string($cached) && is_numeric($cached))) {
            return '<c r="' . $cellRef . '"' . $style . '><f>' . $this->esc((string) $cell->value()) . '</f><v>' . $this->number($cached) . '</v></c>';
        }

        return '<c r="' . $cellRef . '" t="str"' . $style . '><f>' . $this->esc((string) $cell->value()) . '</f><v>' . $this->esc((string) $cached) . '</v></c>';
    }

    private function inlineStringCellXml(string $cellRef, string $string, string $style): string
    {
        $space = (trim($string) !== $string) ? ' xml:space="preserve"' : '';
        return '<c r="' . $cellRef . '" t="inlineStr"' . $style . '><is><t' . $space . '>' . $this->esc($string) . '</t></is></c>';
    }

    /** @param list<list<mixed>> $rows */
    private function maxColumnCount(array $rows): int
    {
        $max = 0;
        foreach ($rows as $row) {
            $max = max($max, count($row));
        }
        return $max;
    }

    private function stylesXml(): string
    {
        $customFormats = $this->customFormats();
        $numFmtXml = '';
        if ($customFormats !== []) {
            $numFmtXml .= '<numFmts count="' . count($customFormats) . '">';
            foreach ($customFormats as $formatCode => $id) {
                $numFmtXml .= '<numFmt numFmtId="' . $id . '" formatCode="' . $this->esc($formatCode) . '"/>';
            }
            $numFmtXml .= '</numFmts>';
        }

        $fonts = '<fonts count="' . (count($this->styleEntries) + 1) . '"><font><sz val="11"/><name val="Calibri"/></font>';
        $fills = '<fills count="' . (count($this->styleEntries) + 2) . '"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>';
        $borders = '<borders count="' . (count($this->styleEntries) + 1) . '"><border><left/><right/><top/><bottom/><diagonal/></border>';
        $xfs = '<cellXfs count="' . (count($this->styleEntries) + 1) . '"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';

        foreach ($this->styleEntries as $id => $style) {
            $fonts .= $this->fontXml($style);
            $fills .= $this->fillXml($style);
            $borders .= $this->borderXml($style);
            [$numFmtId, $isCustom] = $this->numFmtForStyle($style, $customFormats);
            $xfs .= '<xf numFmtId="' . $numFmtId . '" fontId="' . $id . '" fillId="' . ($id + 1) . '" borderId="' . $id . '" xfId="0" applyFont="1" applyFill="1" applyBorder="1"' . ($numFmtId > 0 ? ' applyNumberFormat="1"' : '') . '>';
            $alignment = $this->alignmentXml($style['alignment'] ?? null);
            if ($alignment !== '') {
                $xfs .= $alignment;
            }
            $xfs .= '</xf>';
            unset($isCustom);
        }

        $fonts .= '</fonts>';
        $fills .= '</fills>';
        $borders .= '</borders>';
        $xfs .= '</cellXfs>';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . $numFmtXml
            . $fonts
            . $fills
            . $borders
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . $xfs
            . '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            . '</styleSheet>';
    }

    /** @return array<string,int> */
    private function customFormats(): array
    {
        $formats = [];
        $nextId = 164;
        foreach ($this->styleEntries as $style) {
            $format = $this->formatCode($style['format'] ?? null);
            if ($format === null || $format === '@') {
                continue;
            }
            if (!isset($formats[$format])) {
                $formats[$format] = $nextId++;
            }
        }
        return $formats;
    }

    /** @param array<string,int> $customFormats @return array{0:int,1:bool} */
    private function numFmtForStyle(array $style, array $customFormats): array
    {
        $format = $this->formatCode($style['format'] ?? null);
        if ($format === null) {
            return [0, false];
        }
        if ($format === '@') {
            return [49, false];
        }
        return [$customFormats[$format] ?? 0, true];
    }

    private function formatCode(mixed $format): ?string
    {
        if ($format === null || $format === false || $format === '') {
            return null;
        }

        $format = (string) $format;
        return match (strtolower($format)) {
            'text' => '@',
            'integer', 'int' => '#,##0',
            'number', 'decimal' => '#,##0.00',
            'currency' => '$#,##0.00',
            'percentage', 'percent' => '0.00%',
            'date' => 'yyyy-mm-dd',
            'datetime' => 'yyyy-mm-dd hh:mm:ss',
            default => $format,
        };
    }

    private function fontXml(array $style): string
    {
        $font = is_array($style['font'] ?? null) ? $style['font'] : [];
        $fontSize = isset($font['size']) && is_numeric($font['size']) ? (float) $font['size'] : 11.0;
        $fontName = isset($font['name']) ? (string) $font['name'] : 'Calibri';
        $fontColor = $this->styleColor($font['color'] ?? null, 'FF000000');

        $xml = '<font>';
        if (($font['bold'] ?? false) === true) {
            $xml .= '<b/>';
        }
        if (($font['italic'] ?? false) === true) {
            $xml .= '<i/>';
        }
        if (($font['underline'] ?? false) === true) {
            $xml .= '<u/>';
        }
        $xml .= '<sz val="' . $this->number($fontSize) . '"/><color rgb="' . $fontColor . '"/><name val="' . $this->esc($fontName) . '"/></font>';

        return $xml;
    }

    private function fillXml(array $style): string
    {
        if (!isset($style['fill']) || $style['fill'] === false || $style['fill'] === '') {
            return '<fill><patternFill patternType="none"/></fill>';
        }

        $fill = $this->styleColor($style['fill'], 'FFFFFFFF');
        return '<fill><patternFill patternType="solid"><fgColor rgb="' . $fill . '"/><bgColor indexed="64"/></patternFill></fill>';
    }

    private function borderXml(array $style): string
    {
        if (!isset($style['border']) || $style['border'] === false) {
            return '<border><left/><right/><top/><bottom/><diagonal/></border>';
        }

        $border = $style['border'];
        $borderColor = 'FFD0D7DE';
        if (is_array($border) && isset($border['color'])) {
            $borderColor = $this->styleColor($border['color'], 'FFD0D7DE');
        }
        $borderStyle = is_array($border) && isset($border['style']) ? (string) $border['style'] : 'thin';
        if (!in_array($borderStyle, ['thin', 'medium', 'dashed', 'dotted', 'thick', 'double'], true)) {
            $borderStyle = 'thin';
        }

        return '<border><left style="' . $borderStyle . '"><color rgb="' . $borderColor . '"/></left><right style="' . $borderStyle . '"><color rgb="' . $borderColor . '"/></right><top style="' . $borderStyle . '"><color rgb="' . $borderColor . '"/></top><bottom style="' . $borderStyle . '"><color rgb="' . $borderColor . '"/></bottom><diagonal/></border>';
    }

    private function alignmentXml(mixed $alignment): string
    {
        if ($alignment === null || $alignment === false) {
            return '';
        }

        $horizontal = 'center';
        $vertical = 'center';
        $wrap = false;

        if (is_string($alignment)) {
            $horizontal = $alignment;
        } elseif (is_array($alignment)) {
            $horizontal = (string) ($alignment['horizontal'] ?? $horizontal);
            $vertical = (string) ($alignment['vertical'] ?? $vertical);
            $wrap = (bool) ($alignment['wrap_text'] ?? $alignment['wrapText'] ?? false);
        }

        $allowedHorizontal = ['left', 'center', 'right', 'justify'];
        $allowedVertical = ['top', 'center', 'bottom'];
        if (!in_array($horizontal, $allowedHorizontal, true)) {
            $horizontal = 'center';
        }
        if (!in_array($vertical, $allowedVertical, true)) {
            $vertical = 'center';
        }

        return '<alignment horizontal="' . $horizontal . '" vertical="' . $vertical . '"' . ($wrap ? ' wrapText="1"' : '') . '/>';
    }

    private function cellInRange(string $cellRef, string $range): bool
    {
        [$cellColumn, $cellRow] = Coordinate::splitCellRef($cellRef);
        [$start, $end] = explode(':', strtoupper($range), 2);
        [$startColumn, $startRow] = Coordinate::splitCellRef($start);
        [$endColumn, $endRow] = Coordinate::splitCellRef($end);

        return $cellColumn >= min($startColumn, $endColumn)
            && $cellColumn <= max($startColumn, $endColumn)
            && $cellRow >= min($startRow, $endRow)
            && $cellRow <= max($startRow, $endRow);
    }

    private function sheetRelationships(int $sheetNumber, WorksheetData $sheet, bool $hasImages): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

        if ($hasImages) {
            $xml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing' . $sheetNumber . '.xml"/>';
        }

        foreach (array_values($sheet->hyperlinks) as $index => $link) {
            $xml .= '<Relationship Id="' . $this->hyperlinkRelationshipId($hasImages, $index) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="' . $this->esc((string) $link['url']) . '" TargetMode="External"/>';
        }

        if ($sheet->comments !== []) {
            $commentsRid = $this->commentsRelationshipId($hasImages, count($sheet->hyperlinks));
            $vmlRid = $this->commentsVmlRelationshipId($hasImages, count($sheet->hyperlinks));
            $xml .= '<Relationship Id="' . $commentsRid . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments' . $sheetNumber . '.xml"/>';
            $xml .= '<Relationship Id="' . $vmlRid . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' . $sheetNumber . '.vml"/>';
        }

        return $xml . '</Relationships>';
    }

    private function hyperlinkRelationshipId(bool $hasImages, int $hyperlinkIndex): string
    {
        return 'rId' . (($hasImages ? 1 : 0) + $hyperlinkIndex + 1);
    }

    private function commentsRelationshipId(bool $hasImages, int $hyperlinkCount): string
    {
        return 'rId' . (($hasImages ? 1 : 0) + $hyperlinkCount + 1);
    }

    private function commentsVmlRelationshipId(bool $hasImages, int $hyperlinkCount): string
    {
        return 'rId' . (($hasImages ? 1 : 0) + $hyperlinkCount + 2);
    }

    private function commentsXml(WorksheetData $sheet): string
    {
        $authors = [];
        $authorIds = [];
        foreach ($sheet->comments as $comment) {
            $author = (string) ($comment['author'] ?? 'MNB PHPExcel');
            if (!array_key_exists($author, $authorIds)) {
                $authorIds[$author] = count($authors);
                $authors[] = $author;
            }
        }

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . '<authors>';
        foreach ($authors as $author) {
            $xml .= '<author>' . $this->esc($author) . '</author>';
        }
        $xml .= '</authors><commentList>';

        foreach ($sheet->comments as $index => $comment) {
            $author = (string) ($comment['author'] ?? 'MNB PHPExcel');
            $authorId = $authorIds[$author] ?? 0;
            $xml .= '<comment ref="' . $this->esc($comment['cell']) . '" authorId="' . $authorId . '" shapeId="' . $index . '">'
                . '<text><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr><t xml:space="preserve">'
                . $this->esc((string) $comment['text'])
                . '</t></r></text></comment>';
        }

        return $xml . '</commentList></comments>';
    }

    private function commentsVmlDrawingXml(WorksheetData $sheet): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">'
            . '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>'
            . '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'
            . '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

        foreach (array_values($sheet->comments) as $index => $comment) {
            [$column, $row] = Coordinate::splitCellRef((string) $comment['cell']);
            $zeroColumn = max(0, $column - 1);
            $zeroRow = max(0, $row - 1);
            $width = isset($comment['width']) ? (float) $comment['width'] : 108.0;
            $height = isset($comment['height']) ? (float) $comment['height'] : 59.25;
            $visible = (bool) ($comment['visible'] ?? false);
            $shapeId = 1025 + $index;
            $toColumn = $zeroColumn + 2;
            $toRow = $zeroRow + 4;

            $xml .= '<v:shape id="_x0000_s' . $shapeId . '" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:' . $this->number($width) . 'pt;height:' . $this->number($height) . 'pt;z-index:' . ($index + 1) . ';visibility:' . ($visible ? 'visible' : 'hidden') . '" fillcolor="#ffffe1" o:insetmode="auto">'
                . '<v:fill color2="#ffffe1"/><v:shadow color="black" obscured="t"/><v:path o:connecttype="none"/>'
                . '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
                . '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
                . '<x:Anchor>' . $zeroColumn . ', 15, ' . $zeroRow . ', 2, ' . $toColumn . ', 15, ' . $toRow . ', 16</x:Anchor>'
                . '<x:AutoFill>False</x:AutoFill><x:Row>' . $zeroRow . '</x:Row><x:Column>' . $zeroColumn . '</x:Column>'
                . ($visible ? '<x:Visible/>' : '')
                . '</x:ClientData></v:shape>';
        }

        return $xml . '</xml>';
    }

    /** @param list<array{mediaName:string}> $images */
    private function drawingRels(array $images): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        foreach ($images as $index => $image) {
            $xml .= '<Relationship Id="rId' . ($index + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . $this->esc($image['mediaName']) . '"/>';
        }
        return $xml . '</Relationships>';
    }

    /** @param list<array{cell:string,width:int,height:int,name:string}> $images */
    private function drawingXml(array $images): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

        foreach ($images as $index => $image) {
            [$col, $row] = Coordinate::splitCellRef($image['cell']);
            $cx = $image['width'] * 9525;
            $cy = $image['height'] * 9525;
            $picId = $index + 1;
            $xml .= '<xdr:oneCellAnchor>'
                . '<xdr:from><xdr:col>' . ($col - 1) . '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' . ($row - 1) . '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                . '<xdr:ext cx="' . $cx . '" cy="' . $cy . '"/>'
                . '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' . $picId . '" name="' . $this->esc($image['name']) . '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
                . '<xdr:blipFill><a:blip r:embed="rId' . $picId . '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                . '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
                . '<xdr:clientData/></xdr:oneCellAnchor>';
        }

        return $xml . '</xdr:wsDr>';
    }

    /** @param array<string,mixed> $metadata */
    private function coreProps(array $metadata = []): string
    {
        $now = gmdate('Y-m-d\\TH:i:s\\Z');
        $creator = (string) ($metadata['creator'] ?? 'MNB PHPExcel');
        $modifiedBy = (string) ($metadata['last_modified_by'] ?? $metadata['modified_by'] ?? $creator);

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            . '<dc:creator>' . $this->esc($creator) . '</dc:creator>'
            . '<cp:lastModifiedBy>' . $this->esc($modifiedBy) . '</cp:lastModifiedBy>';

        foreach ([
            'title' => 'dc:title',
            'subject' => 'dc:subject',
            'description' => 'dc:description',
            'category' => 'cp:category',
            'keywords' => 'cp:keywords',
        ] as $key => $tag) {
            if (isset($metadata[$key]) && $metadata[$key] !== '') {
                $xml .= '<' . $tag . '>' . $this->esc((string) $metadata[$key]) . '</' . $tag . '>';
            }
        }

        return $xml
            . '<dcterms:created xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:created>'
            . '<dcterms:modified xsi:type="dcterms:W3CDTF">' . $now . '</dcterms:modified>'
            . '</cp:coreProperties>';
    }

    /** @param array<string,mixed> $metadata */
    private function appProps(array $metadata = []): string
    {
        $company = (string) ($metadata['company'] ?? '');
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            . '<Application>MNB PHPExcel</Application>'
            . ($company !== '' ? '<Company>' . $this->esc($company) . '</Company>' : '')
            . '</Properties>';
    }

    private function esc(string $value): string
    {
        return htmlspecialchars($value, ENT_XML1 | ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }

    private function styleColor(mixed $color, string $fallback): string
    {
        if (is_array($color)) {
            $color = $color['color'] ?? $color['rgb'] ?? $fallback;
        }
        if (!is_string($color)) {
            return $fallback;
        }
        return $this->normalizeArgb($color, $fallback);
    }

    private function normalizeArgb(string $color, string $fallback = 'FFEEF4FF'): string
    {
        $color = strtoupper(ltrim(trim($color), '#'));
        if (preg_match('/^[0-9A-F]{6}$/', $color)) {
            return 'FF' . $color;
        }
        if (preg_match('/^[0-9A-F]{8}$/', $color)) {
            return $color;
        }
        return $fallback;
    }

    private function imageExtension(string $path): string
    {
        $extension = strtolower(pathinfo($path, PATHINFO_EXTENSION));
        return match ($extension) {
            'jpg', 'jpeg' => 'jpeg',
            'png' => 'png',
            'gif' => 'gif',
            default => throw new MnbExcelException('Unsupported image type for XLSX: ' . $extension),
        };
    }

    private function imageContentType(string $extension): string
    {
        return match ($extension) {
            'jpeg', 'jpg' => 'image/jpeg',
            'png' => 'image/png',
            'gif' => 'image/gif',
            default => 'application/octet-stream',
        };
    }

    private function number(int|float|string $number): string
    {
        if (is_string($number)) {
            $value = trim($number);
            if ($value === '' || !is_numeric($value)) {
                return '0';
            }
            // Normalize decimal separator while preserving significant digits from
            // numeric strings. Reject locale/group separators because XLSX numeric
            // values must be invariant XML numbers.
            $value = str_replace(',', '', $value);
            if (preg_match('/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/', $value) === 1) {
                return $this->trimNumericString($value);
            }
            $number = (float) $value;
        }

        if (is_int($number)) {
            return (string) $number;
        }

        if (!is_finite($number)) {
            return '0';
        }

        // Excel stores numbers with roughly 15 significant digits. Use a compact
        // invariant representation instead of number_format(..., 6), which rounded
        // all decimals to six places and caused silent precision loss.
        $value = sprintf('%.15G', $number);
        return $this->trimNumericString($value);
    }

    private function trimNumericString(string $value): string
    {
        if (stripos($value, 'e') !== false) {
            return strtoupper($value);
        }

        if (str_contains($value, '.')) {
            $value = rtrim(rtrim($value, '0'), '.');
        }

        if ($value === '' || $value === '+' || $value === '-') {
            return '0';
        }

        if ($value === '-0') {
            return '0';
        }

        return $value;
    }
}
