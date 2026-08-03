<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use DateTimeImmutable;
use DateTimeInterface;
use DateTimeZone;
use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\StyleNormalizer;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Reader\XlsxWorkbookResolver;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Security\FormulaGuard;
use Mnb\PHPExcel\Security\DocumentProtection;
use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

final class XlsxWriter
{
    /** @var array<int, array<string,mixed>> */
    private array $styleEntries = [];

    /** @var array<string,int> */
    private array $styleIds = [];

    /** @var array<int,array<string,mixed>> */
    private array $differentialStyles = [];

    /** @var array<string,int> */
    private array $differentialStyleIds = [];

    /** @var array<string,mixed> */
    private array $currentWorkbookMetadata = [];

    public function write(WorkbookData $workbook, string $path): void
    {
        $encryption = $workbook->metadata['_mnb_xlsx_encryption'] ?? null;
        if (is_array($encryption) && (string) ($encryption['password'] ?? '') !== '') {
            AtomicFileWriter::writeViaTemp($path, function (string $tmp) use ($workbook, $encryption): void {
                $plain = tempnam(sys_get_temp_dir(), 'mnb_xlsx_plain_');
                if ($plain === false) {
                    throw new MnbExcelException('Unable to allocate a temporary XLSX package for encryption.');
                }
                try {
                    $this->writePackage($workbook, $plain);
                    $this->validateWrittenWorkbook($workbook, $plain);
                    (new XlsxEncryption())->encryptFile($plain, $tmp, (string) $encryption['password'], $encryption);
                } finally {
                    @unlink($plain);
                }
            }, static function (string $tmp): void {
                if (!(new XlsxEncryption())->isEncryptedFile($tmp)) {
                    throw new MnbExcelException('Encrypted XLSX output did not produce a valid Office encrypted container.');
                }
            });
            return;
        }

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
        $this->currentWorkbookMetadata = $workbook->metadata;
        if (!class_exists(ZipArchive::class)) {
            throw MnbExcelException::withCode(
                'ext-zip is required to write XLSX files.',
                ErrorCode::EXTENSION_MISSING,
                ['extension' => 'zip']
            );
        }

        $this->buildStyleRegistry($workbook);
        $imagePlan = $this->buildImagePlan($workbook);
        $chartPlan = $this->buildChartPlan($workbook);
        $pivotPlan = $this->buildPivotPlan($workbook);
        $preservedPackage = $this->loadPreservedPackage($workbook, $imagePlan, $chartPlan);
        if ($pivotPlan !== [] && is_array($preservedPackage) && (string) ($preservedPackage['workbook_pivot_caches'] ?? '') !== '') {
            throw new MnbExcelException('Generated pivot tables cannot be combined with preserved template pivots in the same write.');
        }

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
            $addXml('[Content_Types].xml', $this->contentTypes($workbook, $imagePlan, $chartPlan, $pivotPlan, $preservedPackage));
            $addXml('_rels/.rels', $this->rootRels());
            $addXml('docProps/core.xml', $this->coreProps($workbook->metadata));
            $addXml('docProps/app.xml', $this->appProps($workbook->metadata));
            $addXml('xl/workbook.xml', $this->workbookXml($workbook, $pivotPlan, $preservedPackage));
            $addXml('xl/_rels/workbook.xml.rels', $this->workbookRels(count($workbook->sheets), $pivotPlan, $preservedPackage));
            $addXml('xl/styles.xml', $this->stylesXml());

            foreach ($workbook->sheets as $index => $sheet) {
                $sheetNumber = $index + 1;
                $sheetImages = $imagePlan[$sheetNumber] ?? [];
                $sheetCharts = $chartPlan[$sheetNumber] ?? [];
                $sheetPivots = $pivotPlan[$sheetNumber] ?? [];
                $hasDrawing = $sheetImages !== [] || $sheetCharts !== [];
                $hasHyperlinks = $sheet->hyperlinks !== [];
                $hasComments = $sheet->comments !== [];
                $hasGeneratedSheetRels = $hasDrawing || $hasHyperlinks || $hasComments || $sheetPivots !== [];
                $preservedSheet = is_array($preservedPackage) ? ($preservedPackage['sheets'][$sheetNumber] ?? null) : null;
                if ($hasGeneratedSheetRels && is_array($preservedSheet) && (bool) ($preservedSheet['requires_relationships'] ?? false)) {
                    throw new MnbExcelException(
                        'The selected template sheet contains preserved relationship-backed objects and cannot also receive newly generated images, charts, hyperlinks, or comments in the same write. '
                        . 'Keep those objects in the template, generate them on another sheet, or disable package preservation for this workbook.'
                    );
                }
                $addXml('xl/worksheets/sheet' . $sheetNumber . '.xml', $this->worksheetXml($sheet, $hasDrawing, $sheetPivots, is_array($preservedSheet) ? $preservedSheet : null));

                if ($hasGeneratedSheetRels) {
                    $addXml('xl/worksheets/_rels/sheet' . $sheetNumber . '.xml.rels', $this->sheetRelationships($sheetNumber, $sheet, $hasDrawing, $sheetPivots));
                } elseif (is_array($preservedSheet) && ($preservedSheet['rels_xml'] ?? '') !== '') {
                    $entry = 'xl/worksheets/_rels/sheet' . $sheetNumber . '.xml.rels';
                    $addXml($entry, (string) $preservedSheet['rels_xml']);
                }

                if ($hasDrawing) {
                    $addXml('xl/drawings/drawing' . $sheetNumber . '.xml', $this->drawingXml($sheetImages, $sheetCharts));
                    $addXml('xl/drawings/_rels/drawing' . $sheetNumber . '.xml.rels', $this->drawingRels($sheetImages, $sheetCharts));

                    foreach ($sheetImages as $image) {
                        $entry = 'xl/media/' . $image['mediaName'];
                        $this->addFileToZip($zip, $image['path'], $entry);
                        $writtenEntries[$entry] = true;
                    }
                    foreach ($sheetCharts as $chart) {
                        $entry = 'xl/charts/' . $chart['partName'];
                        $addXml($entry, $this->chartXml($chart, $sheet->name));
                    }
                }

                if ($hasComments) {
                    $addXml('xl/comments' . $sheetNumber . '.xml', $this->commentsXml($sheet));
                    $addXml('xl/drawings/vmlDrawing' . $sheetNumber . '.vml', $this->commentsVmlDrawingXml($sheet));
                }

                foreach ($sheetPivots as $pivot) {
                    $number = (int) $pivot['pivotNumber'];
                    $addXml('xl/pivotTables/pivotTable' . $number . '.xml', $this->pivotTableXml($pivot));
                    $addXml('xl/pivotTables/_rels/pivotTable' . $number . '.xml.rels', $this->pivotTableRelationships($number));
                    $addXml('xl/pivotCache/pivotCacheDefinition' . $number . '.xml', $this->pivotCacheDefinitionXml($pivot));
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
        $this->differentialStyles = [];
        $this->differentialStyleIds = [];

        foreach ($workbook->sheets as $sheet) {
            foreach ($sheet->conditionalFormats as $rule) {
                if (array_key_exists('style', $rule)) {
                    $this->differentialStyleIdFor($this->resolveStyle($sheet, $rule['style']));
                }
            }
            foreach ($sheet->rows as $rowIndex => $row) {
                $r = $rowIndex + 1;
                foreach (array_values($row) as $columnIndex => $value) {
                    $c = $columnIndex + 1;
                    $style = $this->effectiveCellStyle(
                        $sheet,
                        $r,
                        $c,
                        Coordinate::columnIndexToName($c) . $r,
                        $value
                    );
                    $this->styleIdFor($style);
                }
            }
        }
    }

    /** @return array<string,mixed> */
    private function effectiveCellStyle(
        WorksheetData $sheet,
        int $row,
        int $column,
        string $cellRef,
        mixed $value = null
    ): array
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

        if (!array_key_exists('format', $style)) {
            if ($value instanceof CellValue && $value->type() === CellValue::TYPE_DATE) {
                $style['format'] = (string) ($value->options()['format'] ?? 'm/d/yy');
            } elseif ($value instanceof DateTimeInterface) {
                $style['format'] = $value->format('His.u') === '000000.000000'
                    ? 'm/d/yy'
                    : 'm/d/yy h:mm';
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
        return StyleNormalizer::normalize($style);
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

    /** @return array<int,list<array<string,mixed>>> */
    private function buildChartPlan(WorkbookData $workbook): array
    {
        $plan = [];
        $chartNumber = 1;
        foreach ($workbook->sheets as $sheetIndex => $sheet) {
            $sheetNumber = $sheetIndex + 1;
            foreach ($sheet->charts as $chart) {
                $from = strtoupper((string) ($chart['from'] ?? 'A1'));
                $to = strtoupper((string) ($chart['to'] ?? 'H16'));
                Coordinate::splitCellRef($from);
                Coordinate::splitCellRef($to);
                $chart['partName'] = 'chart' . $chartNumber . '.xml';
                $chart['chartNumber'] = $chartNumber;
                $chart['from'] = $from;
                $chart['to'] = $to;
                $plan[$sheetNumber][] = $chart;
                $chartNumber++;
            }
        }
        return $plan;
    }

    /** @return array<int,list<array<string,mixed>>> */
    private function buildPivotPlan(WorkbookData $workbook): array
    {
        $plan = [];
        $sheetByName = [];
        foreach ($workbook->sheets as $index => $sheet) {
            $sheetByName[$sheet->name] = ['index' => $index + 1, 'sheet' => $sheet];
        }

        $pivotNumber = 1;
        $names = [];
        foreach ($workbook->sheets as $targetIndex => $targetSheet) {
            foreach ($targetSheet->pivotTables as $definition) {
                $sourceName = (string) ($definition['source_sheet'] ?? '');
                if (!isset($sheetByName[$sourceName])) {
                    throw new MnbExcelException('Pivot source worksheet not found: ' . $sourceName);
                }
                /** @var WorksheetData $sourceSheet */
                $sourceSheet = $sheetByName[$sourceName]['sheet'];
                [$startRef, $endRef] = $this->normalizeRange((string) ($definition['source_range'] ?? 'A1:A1'));
                [$startColumn, $startRow] = Coordinate::splitCellRef($startRef);
                [$endColumn, $endRow] = Coordinate::splitCellRef($endRef);
                $firstColumn = min($startColumn, $endColumn);
                $lastColumn = max($startColumn, $endColumn);
                $headerRow = min($startRow, $endRow);
                $headers = [];
                for ($column = $firstColumn; $column <= $lastColumn; $column++) {
                    $value = $sourceSheet->rows[$headerRow - 1][$column - 1] ?? null;
                    $name = trim((string) ($value instanceof CellValue ? $value->displayValue() : $value));
                    $headers[] = $name !== '' ? $name : 'Field' . ($column - $firstColumn + 1);
                }

                $resolveField = function (mixed $field) use ($headers, $firstColumn): int {
                    if (is_int($field) || ctype_digit((string) $field)) {
                        $index = (int) $field - 1;
                    } elseif (preg_match('/^[A-Za-z]{1,3}$/', trim((string) $field)) === 1) {
                        $index = Coordinate::columnNameToIndex((string) $field) - $firstColumn;
                    } else {
                        $index = array_search(strtolower(trim((string) $field)), array_map('strtolower', $headers), true);
                        if ($index === false) {
                            throw new MnbExcelException('Pivot field not found: ' . (string) $field);
                        }
                    }
                    if ($index < 0 || $index >= count($headers)) {
                        throw new MnbExcelException('Pivot field is outside the source range: ' . (string) $field);
                    }
                    return $index;
                };

                $resolveAxis = function (array $items) use ($resolveField): array {
                    $resolved = [];
                    foreach (array_values($items) as $item) {
                        $config = is_array($item) ? $item : ['field' => $item];
                        $field = $resolveField($config['field'] ?? 1);
                        $subtotals = array_values(array_filter(array_map('strtolower', (array) ($config['subtotals'] ?? [])), static fn (string $name): bool => in_array($name, ['sum','count','average','max','min','product','count_nums','std_dev','std_dev_p','var','var_p'], true)));
                        $resolved[] = $config + [
                            'field' => $field,
                            'subtotals' => $subtotals,
                            'sort' => in_array(strtolower((string) ($config['sort'] ?? 'manual')), ['manual','ascending','descending'], true) ? strtolower((string) ($config['sort'] ?? 'manual')) : 'manual',
                            'show_all' => (bool) ($config['show_all'] ?? true),
                            'repeat_labels' => (bool) ($config['repeat_labels'] ?? false),
                            'insert_blank_row' => (bool) ($config['insert_blank_row'] ?? false),
                            'show_items' => array_values((array) ($config['show_items'] ?? [])),
                            'hide_items' => array_values((array) ($config['hide_items'] ?? [])),
                        ];
                    }
                    return $resolved;
                };
                $rows = $resolveAxis((array) ($definition['rows'] ?? []));
                $columns = $resolveAxis((array) ($definition['columns'] ?? []));
                $filters = $resolveAxis((array) ($definition['filters'] ?? []));
                $values = [];
                foreach (array_values((array) ($definition['values'] ?? [])) as $value) {
                    $fieldIndex = $resolveField($value['field'] ?? 1);
                    $function = strtolower((string) ($value['function'] ?? 'sum'));
                    $showDataAs = strtolower((string) ($value['show_data_as'] ?? 'normal'));
                    $allowedShowDataAs = ['normal','difference','percent','percent_diff','run_total','percent_row','percent_col','percent_total','index'];
                    if (!in_array($showDataAs, $allowedShowDataAs, true)) { $showDataAs = 'normal'; }
                    $values[] = [
                        'field' => $fieldIndex,
                        'function' => $function,
                        'name' => (string) ($value['name'] ?? ucfirst(str_replace('_', ' ', $function)) . ' of ' . $headers[$fieldIndex]),
                        'number_format' => (string) ($value['number_format'] ?? ''),
                        'show_data_as' => $showDataAs,
                        'base_field' => isset($value['base_field']) ? $resolveField($value['base_field']) : -1,
                        'base_item' => (int) ($value['base_item'] ?? 1048832),
                    ];
                }

                $targetCell = strtoupper((string) ($definition['target_cell'] ?? 'A1'));
                [$targetColumn, $targetRow] = Coordinate::splitCellRef($targetCell);
                $endTarget = Coordinate::columnIndexToName($targetColumn + max(3, (int) ($definition['width'] ?? 8)) - 1)
                    . ($targetRow + max(5, (int) ($definition['height'] ?? 20)) - 1);
                $name = (string) ($definition['name'] ?? ('PivotTable' . $pivotNumber));
                $baseName = $name;
                $suffix = 2;
                while (isset($names[strtolower($name)])) {
                    $name = $baseName . '_' . $suffix++;
                }
                $names[strtolower($name)] = true;

                $plan[$targetIndex + 1][] = $definition + [
                    'pivotNumber' => $pivotNumber,
                    'cacheId' => $pivotNumber,
                    'name' => $name,
                    'source_sheet' => $sourceName,
                    'source_range' => $startRef . ':' . $endRef,
                    'target_range' => $targetCell . ':' . $endTarget,
                    'headers' => $headers,
                    'rows_resolved' => $rows,
                    'columns_resolved' => $columns,
                    'filters_resolved' => $filters,
                    'values_resolved' => $values,
                    'record_count' => max(0, max($startRow, $endRow) - $headerRow),
                ];
                $pivotNumber++;
            }
        }
        return $plan;
    }

    /** @param array<string,mixed> $pivot */
    private function pivotCacheDefinitionXml(array $pivot): string
    {
        $fields = '';
        foreach ((array) $pivot['headers'] as $header) {
            $fields .= '<cacheField name="' . $this->esc((string) $header) . '" numFmtId="0"><sharedItems/></cacheField>';
        }
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            . ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            . ' saveData="' . ((bool) ($pivot['save_data'] ?? false) ? '1' : '0') . '" refreshOnLoad="' . ((bool) ($pivot['refresh_on_load'] ?? true) ? '1' : '0') . '"'
            . ' recordCount="' . (int) ($pivot['record_count'] ?? 0) . '" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3">'
            . '<cacheSource type="worksheet"><worksheetSource ref="' . $this->esc((string) $pivot['source_range']) . '" sheet="' . $this->esc((string) $pivot['source_sheet']) . '"/></cacheSource>'
            . '<cacheFields count="' . count((array) $pivot['headers']) . '">' . $fields . '</cacheFields>'
            . '</pivotCacheDefinition>';
    }

    /** @param array<string,mixed> $pivot */
    private function pivotTableXml(array $pivot): string
    {
        $rows = (array) ($pivot['rows_resolved'] ?? []);
        $columns = (array) ($pivot['columns_resolved'] ?? []);
        $filters = (array) ($pivot['filters_resolved'] ?? []);
        $values = (array) ($pivot['values_resolved'] ?? []);
        $fieldCount = count((array) ($pivot['headers'] ?? []));
        $rowMap=[];$columnMap=[];$filterMap=[];foreach($rows as $item)$rowMap[(int)$item['field']]=$item;foreach($columns as $item)$columnMap[(int)$item['field']]=$item;foreach($filters as $item)$filterMap[(int)$item['field']]=$item;
        $layout=(string)($pivot['layout']??'compact');$pivotFields='';
        for($index=0;$index<$fieldCount;$index++){
            $config=$rowMap[$index]??$columnMap[$index]??$filterMap[$index]??[];$axis=isset($rowMap[$index])?' axis="axisRow"':(isset($columnMap[$index])?' axis="axisCol"':(isset($filterMap[$index])?' axis="axisPage"':' dataField="1"'));
            $attrs=' showAll="'.((bool)($config['show_all']??false)?'1':'0').'"';
            if(isset($rowMap[$index])){$attrs.=' compact="'.($layout==='compact'?'1':'0').'" outline="'.($layout==='outline'?'1':'0').'"';if((bool)($config['repeat_labels']??$pivot['repeat_item_labels']??false))$attrs.=' itemPageCount="10"';if((bool)($config['insert_blank_row']??false))$attrs.=' insertBlankRow="1"';}
            $sort=(string)($config['sort']??'manual');if($sort!=='manual')$attrs.=' sortType="'.$sort.'"';
            $subtotalAttrs='';foreach((array)($config['subtotals']??[]) as $subtotal){$name=match($subtotal){'average'=>'avgSubtotal','max'=>'maxSubtotal','min'=>'minSubtotal','product'=>'productSubtotal','count','count_nums'=>'countSubtotal','std_dev'=>'stdDevSubtotal','std_dev_p'=>'stdDevPSubtotal','var'=>'varSubtotal','var_p'=>'varPSubtotal',default=>'sumSubtotal'};$subtotalAttrs.=' '.$name.'="1"';}
            $pivotFields.='<pivotField'.$axis.$attrs.$subtotalAttrs.'><items count="1"><item t="default"/></items></pivotField>';
        }
        $rowFields='';foreach($rows as $item)$rowFields.='<field x="'.(int)$item['field'].'"/>';
        $colFields='';foreach($columns as $item)$colFields.='<field x="'.(int)$item['field'].'"/>';if(count($values)>1)$colFields.='<field x="-2"/>';
        $pageFields='';foreach($filters as $item)$pageFields.='<pageField fld="'.(int)$item['field'].'" hier="-1"'.(isset($item['selected_item'])?' item="'.(int)$item['selected_item'].'"':'').'/>';
        $dataFields='';foreach($values as $value){$subtotal=match((string)$value['function']){'average'=>'average','max'=>'max','min'=>'min','product'=>'product','count','count_nums'=>'count','std_dev'=>'stdDev','std_dev_p'=>'stdDevP','var'=>'var','var_p'=>'varP',default=>'sum'};$show=match((string)($value['show_data_as']??'normal')){'difference'=>'difference','percent'=>'percent','percent_diff'=>'percentDiff','run_total'=>'runTotal','percent_row'=>'percentOfRow','percent_col'=>'percentOfCol','percent_total'=>'percentOfTotal','index'=>'index',default=>'normal'};$dataFields.='<dataField name="'.$this->esc((string)$value['name']).'" fld="'.(int)$value['field'].'" subtotal="'.$subtotal.'"'.($show!=='normal'?' showDataAs="'.$show.'" baseField="'.(int)($value['base_field']??-1).'" baseItem="'.(int)($value['base_item']??1048832).'"':'').'/>';}
        $xml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'.'<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="'.$this->esc((string)$pivot['name']).'" cacheId="'.(int)$pivot['cacheId'].'" dataCaption="'.$this->esc((string)($pivot['data_caption']??'Values')).'" rowHeaderCaption="'.$this->esc((string)($pivot['row_header_caption']??'Row Labels')).'" colHeaderCaption="'.$this->esc((string)($pivot['column_header_caption']??'Column Labels')).'" grandTotalCaption="'.$this->esc((string)($pivot['grand_total_caption']??'Grand Total')).'" updatedVersion="8" minRefreshableVersion="3" useAutoFormatting="1" preserveFormatting="'.((bool)($pivot['preserve_formatting']??true)?'1':'0').'" showDrill="'.((bool)($pivot['show_drill']??true)?'1':'0').'" enableWizard="'.((bool)($pivot['show_field_list']??true)?'1':'0').'" showEmptyRow="'.((bool)($pivot['show_empty_rows']??false)?'1':'0').'" showEmptyCol="'.((bool)($pivot['show_empty_columns']??false)?'1':'0').'" rowGrandTotals="'.((bool)($pivot['show_row_grand_totals']??true)?'1':'0').'" colGrandTotals="'.((bool)($pivot['show_column_grand_totals']??true)?'1':'0').'">'.'<location ref="'.$this->esc((string)$pivot['target_range']).'" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>'.'<pivotFields count="'.$fieldCount.'">'.$pivotFields.'</pivotFields>';
        if($rows!==[])$xml.='<rowFields count="'.count($rows).'">'.$rowFields.'</rowFields><rowItems count="1"><i/></rowItems>';
        if($columns!==[]||count($values)>1)$xml.='<colFields count="'.(count($columns)+(count($values)>1?1:0)).'">'.$colFields.'</colFields><colItems count="1"><i/></colItems>';
        if($filters!==[])$xml.='<pageFields count="'.count($filters).'">'.$pageFields.'</pageFields>';
        $xml.='<dataFields count="'.count($values).'">'.$dataFields.'</dataFields>'.'<pivotTableStyleInfo name="'.$this->esc((string)($pivot['style']??'PivotStyleMedium9')).'" showRowHeaders="'.((bool)($pivot['show_row_headers']??true)?'1':'0').'" showColHeaders="'.((bool)($pivot['show_column_headers']??true)?'1':'0').'" showRowStripes="'.((bool)($pivot['show_row_stripes']??false)?'1':'0').'" showColStripes="'.((bool)($pivot['show_column_stripes']??false)?'1':'0').'" showLastColumn="'.((bool)($pivot['show_last_column']??false)?'1':'0').'"/>'.'</pivotTableDefinition>';
        return $xml;
    }

    private function pivotTableRelationships(int $number): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition' . $number . '.xml"/>'
            . '</Relationships>';
    }

    /** @param array<string,mixed> $style */
    private function differentialStyleIdFor(array $style): int
    {
        $normalized = $this->normalizeStyle($style);
        $key = json_encode($normalized, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
        if ($key === false) {
            $key = serialize($normalized);
        }
        if (!isset($this->differentialStyleIds[$key])) {
            $id = count($this->differentialStyles);
            $this->differentialStyleIds[$key] = $id;
            $this->differentialStyles[$id] = $normalized;
        }
        return $this->differentialStyleIds[$key];
    }

    /** @param array<int, list<array{extension:string}>> $imagePlan @param array<int,list<array<string,mixed>>> $chartPlan @param array<int,list<array<string,mixed>>> $pivotPlan @param array<string,mixed>|null $preservedPackage */
    private function contentTypes(WorkbookData $workbook, array $imagePlan, array $chartPlan, array $pivotPlan, ?array $preservedPackage = null): string
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
            if ((isset($imagePlan[$i]) && $imagePlan[$i] !== []) || (isset($chartPlan[$i]) && $chartPlan[$i] !== [])) {
                $overrides['/xl/drawings/drawing' . $i . '.xml'] = 'application/vnd.openxmlformats-officedocument.drawing+xml';
            }
            foreach ($chartPlan[$i] ?? [] as $chart) {
                $overrides['/xl/charts/' . $chart['partName']] = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
            }
            foreach ($pivotPlan[$i] ?? [] as $pivot) {
                $number = (int) $pivot['pivotNumber'];
                $overrides['/xl/pivotTables/pivotTable' . $number . '.xml'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';
                $overrides['/xl/pivotCache/pivotCacheDefinition' . $number . '.xml'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';
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
    private function loadPreservedPackage(WorkbookData $workbook, array $imagePlan, array $chartPlan = []): ?array
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
                'chart_conflict' => isset($chartPlan[$sheetNumber]) && $chartPlan[$sheetNumber] !== [],
            ];
        }

        $contentTypesXml = $zip->getFromName('[Content_Types].xml');
        $contentTypes = $contentTypesXml !== false ? $this->parseContentTypes($contentTypesXml) : ['defaults' => [], 'overrides' => []];
        $workbookPivot = $this->extractWorkbookPivotState(
            (string) ($zip->getFromName('xl/workbook.xml') ?: ''),
            (string) ($zip->getFromName('xl/_rels/workbook.xml.rels') ?: ''),
            count($workbook->sheets)
        );
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
            'workbook_pivot_caches' => $workbookPivot['pivot_caches'],
            'workbook_relationships' => $workbookPivot['relationships'],
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
            if (str_starts_with(strtolower($entry), 'xl/pivotcache/pivotcachedefinition')) {
                $content = $this->rebindPivotCacheDefinition($content, (array) ($preservedPackage['settings'] ?? []));
            }
            $this->addStringToZip($outputZip, $entry, $content);
            $writtenEntries[$entry] = true;
        }

        $sourceZip->close();
    }

    /** @return list<string> */
    private function extractPreservedSheetElements(string $sheetXml): array
    {
        // Keep the original worksheet-schema order. Reordering preserved elements can
        // make otherwise valid templates trigger Excel's repair dialog.
        $tags = [
            'sheetCalcPr',
            'sheetProtection',
            'protectedRanges',
            'scenarios',
            'sortState',
            'dataConsolidate',
            'customSheetViews',
            'phoneticPr',
            'conditionalFormatting',
            'dataValidations',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'pageSetup',
            'headerFooter',
            'rowBreaks',
            'colBreaks',
            'customProperties',
            'cellWatches',
            'ignoredErrors',
            'smartTags',
            'drawing',
            'legacyDrawing',
            'legacyDrawingHF',
            'picture',
            'oleObjects',
            'controls',
            'webPublishItems',
            'tableParts',
            'pivotTableParts',
            'extLst',
        ];

        $pattern = '/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?(' . implode('|', array_map(static fn (string $tag): string => preg_quote($tag, '/'), $tags)) . ')\b[^>]*(?:\/>|>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?\1>)/isu';
        preg_match_all($pattern, $sheetXml, $matches, PREG_OFFSET_CAPTURE);
        $elements = [];
        foreach ($matches[0] ?? [] as $match) {
            if (is_array($match) && isset($match[0], $match[1])) {
                $elements[] = ['xml' => (string) $match[0], 'offset' => (int) $match[1]];
            }
        }
        usort($elements, static fn (array $a, array $b): int => $a['offset'] <=> $b['offset']);
        return array_values(array_map(static fn (array $item): string => $item['xml'], $elements));
    }

    /** @param list<string> $elements */
    private function elementsRequireRelationships(array $elements): bool
    {
        foreach ($elements as $element) {
            if (preg_match('/\br:id\s*=/i', (string) $element) === 1) {
                return true;
            }
        }
        return false;
    }

    /** @return array{pivot_caches:string,relationships:list<array{id:string,type:string,target:string}>} */
    private function extractWorkbookPivotState(string $workbookXml, string $relsXml, int $sheetCount): array
    {
        if ($workbookXml === '' || $relsXml === '') {
            return ['pivot_caches' => '', 'relationships' => []];
        }
        $relationships = [];
        $idMap = [];
        $nextId = $sheetCount + 2;
        preg_match_all('/<Relationship\b[^>]*\/?\s*>/i', $relsXml, $matches);
        foreach ($matches[0] ?? [] as $tag) {
            $attrs = $this->parseXmlAttributes($tag);
            $type = (string) ($attrs['Type'] ?? '');
            if (!str_contains(strtolower($type), '/pivotcache')) {
                continue;
            }
            $oldId = (string) ($attrs['Id'] ?? '');
            $target = (string) ($attrs['Target'] ?? '');
            if ($oldId === '' || $target === '') {
                continue;
            }
            $newId = 'rId' . $nextId++;
            $idMap[$oldId] = $newId;
            $relationships[] = ['id' => $newId, 'type' => $type, 'target' => $target];
        }
        if ($relationships === [] || preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?pivotCaches\b[^>]*>.*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?pivotCaches>/is', $workbookXml, $pivotMatch) !== 1) {
            return ['pivot_caches' => '', 'relationships' => []];
        }
        $pivotCaches = $pivotMatch[0];
        foreach ($idMap as $oldId => $newId) {
            $pivotCaches = preg_replace('/(r:id\s*=\s*["\'])' . preg_quote($oldId, '/') . '(["\'])/i', '$1' . $newId . '$2', $pivotCaches) ?? $pivotCaches;
        }
        return ['pivot_caches' => $pivotCaches, 'relationships' => $relationships];
    }

    /** @param array<string,mixed> $settings */
    private function rebindPivotCacheDefinition(string $xml, array $settings): string
    {
        $sheet = trim((string) ($settings['pivot_source_sheet'] ?? ''));
        $range = strtoupper(trim((string) ($settings['pivot_source_range'] ?? '')));
        if ($sheet !== '' && $range !== '') {
            $xml = preg_replace_callback('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?worksheetSource\b([^>]*)\/?\s*>/i', function (array $match) use ($sheet, $range): string {
                $attrs = $this->parseXmlAttributes('<worksheetSource ' . $match[1] . '>');
                $attrs['sheet'] = $sheet;
                $attrs['ref'] = $range;
                unset($attrs['name']);
                $text = '<worksheetSource';
                foreach ($attrs as $name => $value) {
                    $text .= ' ' . $name . '="' . $this->esc((string) $value) . '"';
                }
                return $text . '/>';
            }, $xml) ?? $xml;
        }
        if ((bool) ($settings['pivot_refresh_on_load'] ?? true)) {
            $xml = preg_replace_callback('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?pivotCacheDefinition\b([^>]*)>/i', function (array $match): string {
                $attrs = $this->parseXmlAttributes('<pivotCacheDefinition ' . $match[1] . '>');
                $attrs['refreshOnLoad'] = '1';
                $attrs['enableRefresh'] = '1';
                $attrs['refreshedBy'] = 'MNB PHPExcel';
                $text = '<pivotCacheDefinition';
                foreach ($attrs as $name => $value) {
                    $text .= ' ' . $name . '="' . $this->esc((string) $value) . '"';
                }
                return $text . '>';
            }, $xml, 1) ?? $xml;
        }
        return $xml;
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

    /** @param array<int,list<array<string,mixed>>> $pivotPlan @param array<string,mixed>|null $preservedPackage */
    private function workbookRels(int $sheetCount, array $pivotPlan, ?array $preservedPackage = null): string
    {
        $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';

        for ($i = 1; $i <= $sheetCount; $i++) {
            $rels .= '<Relationship Id="rId' . $i . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' . $i . '.xml"/>';
        }

        $rels .= '<Relationship Id="rId' . ($sheetCount + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        foreach ($pivotPlan as $pivots) {
            foreach ($pivots as $pivot) {
                $number = (int) $pivot['pivotNumber'];
                $rels .= '<Relationship Id="rIdPivotCache' . $number . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition' . $number . '.xml"/>';
            }
        }
        foreach ((array) ($preservedPackage['workbook_relationships'] ?? []) as $relationship) {
            if (!is_array($relationship)) {
                continue;
            }
            $rels .= '<Relationship Id="' . $this->esc((string) $relationship['id']) . '" Type="' . $this->esc((string) $relationship['type']) . '" Target="' . $this->esc((string) $relationship['target']) . '"/>';
        }
        return $rels . '</Relationships>';
    }

    /** @param array<int,list<array<string,mixed>>> $pivotPlan @param array<string,mixed>|null $preservedPackage */
    private function workbookXml(WorkbookData $workbook, array $pivotPlan, ?array $preservedPackage = null): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        if ((bool) ($workbook->metadata['date1904'] ?? false)) {
            $xml .= '<workbookPr date1904="1"/>';
        }
        $activeSheet = $workbook->metadata['_mnb_active_sheet'] ?? 1;
        if (is_string($activeSheet) && !ctype_digit($activeSheet)) {
            foreach ($workbook->sheets as $index => $candidate) {
                if ($candidate->name === $activeSheet) {
                    $activeSheet = $index + 1;
                    break;
                }
            }
        }
        $activeTab = max(0, min(count($workbook->sheets) - 1, ((int) $activeSheet) - 1));
        $xml .= '<bookViews><workbookView activeTab="' . $activeTab . '"/></bookViews>';
        $workbookProtection = $workbook->metadata['_mnb_workbook_protection'] ?? null;
        if (is_array($workbookProtection) && (string) ($workbookProtection['password'] ?? '') !== '') {
            $xml .= '<workbookProtection' . $this->xmlAttributes(DocumentProtection::workbookAttributes((string) $workbookProtection['password'], $workbookProtection)) . '/>';
        }
        $xml .= '<sheets>';

        $sheetStates = is_array($workbook->metadata['_mnb_sheet_states'] ?? null)
            ? $workbook->metadata['_mnb_sheet_states']
            : [];
        foreach ($workbook->sheets as $index => $sheet) {
            $id = $index + 1;
            $state = (string) ($sheetStates[$sheet->name] ?? $sheetStates[$id] ?? 'visible');
            if (!in_array($state, ['visible', 'hidden', 'veryHidden'], true)) {
                $state = 'visible';
            }
            $xml .= '<sheet name="' . $this->esc($sheet->name) . '" sheetId="' . $id . '" r:id="rId' . $id . '"'
                . ($state !== 'visible' ? ' state="' . $state . '"' : '')
                . '/>';
        }
        $xml .= '</sheets>';
        $generatedPivotCaches = '';
        $pivotCount = 0;
        foreach ($pivotPlan as $pivots) {
            foreach ($pivots as $pivot) {
                $number = (int) $pivot['pivotNumber'];
                $generatedPivotCaches .= '<pivotCache cacheId="' . $number . '" r:id="rIdPivotCache' . $number . '"/>';
                $pivotCount++;
            }
        }
        if ($pivotCount > 0) {
            $xml .= '<pivotCaches count="' . $pivotCount . '">' . $generatedPivotCaches . '</pivotCaches>';
        }
        $pivotCaches = (string) ($preservedPackage['workbook_pivot_caches'] ?? '');
        if ($pivotCaches !== '') {
            $xml .= $pivotCaches;
        }
        return $xml . '<calcPr calcId="191029" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>';
    }

    /** @param list<array<string,mixed>> $pivotTables @param array<string,mixed>|null $preservedSheet */
    private function worksheetXml(WorksheetData $sheet, bool $hasImages = false, ?array $pivotTables = [], ?array $preservedSheet = null): string
    {
        $pivotTables ??= [];
        $rowCount = count($sheet->rows);
        $columnCount = $this->maxColumnCount($sheet->rows);
        $lastCell = $columnCount > 0 && $rowCount > 0 ? Coordinate::columnIndexToName($columnCount) . $rowCount : 'A1';
        $needsRelationships = $hasImages || $sheet->hyperlinks !== [] || $sheet->comments !== [] || $pivotTables !== [] || (is_array($preservedSheet) && (bool) ($preservedSheet['requires_relationships'] ?? false));
        $xmlnsR = $needsRelationships ? ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : '';

        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' . $xmlnsR . '>'
            . '<dimension ref="A1:' . $lastCell . '"/>';

        $freezeRows = $sheet->freezeHeader ? ($sheet->headerRowIndex ?? 1) : max(0, $sheet->freezeRows);
        $freezeColumns = max(0, $sheet->freezeColumns);
        if ($freezeRows > 0 || $freezeColumns > 0) {
            $topLeft = $sheet->freezeTopLeftCell
                ?? Coordinate::columnIndexToName($freezeColumns + 1) . ($freezeRows + 1);
            $activePane = $freezeRows > 0 && $freezeColumns > 0
                ? 'bottomRight'
                : ($freezeColumns > 0 ? 'topRight' : 'bottomLeft');
            $pane = '<pane'
                . ($freezeColumns > 0 ? ' xSplit="' . $freezeColumns . '"' : '')
                . ($freezeRows > 0 ? ' ySplit="' . $freezeRows . '"' : '')
                . ' topLeftCell="' . $this->esc($topLeft) . '" activePane="' . $activePane . '" state="frozen"/>';
            $xml .= '<sheetViews><sheetView workbookViewId="0">' . $pane
                . '<selection pane="' . $activePane . '" activeCell="' . $this->esc($topLeft) . '" sqref="' . $this->esc($topLeft) . '"/>'
                . '</sheetView></sheetViews>';
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
                $styleId = $this->styleIdFor(
                    $this->effectiveCellStyle($sheet, $r, $c, $cellRef, $value)
                );
                $xml .= $this->cellXml($cellRef, $value, $styleId);
            }
            $xml .= '</row>';
        }
        $xml .= '</sheetData>';

        $sheetProtection = $this->sheetProtectionOptions($sheet->name, $this->currentWorkbookMetadata ?? []);
        if ($sheetProtection !== null) {
            $xml .= '<sheetProtection' . $this->xmlAttributes(DocumentProtection::sheetAttributes((string) $sheetProtection['password'], $sheetProtection)) . '/>';
        }

        if ($sheet->mergeCells !== []) {
            $xml .= '<mergeCells count="' . count($sheet->mergeCells) . '">';
            foreach ($sheet->mergeCells as $range) {
                $xml .= '<mergeCell ref="' . $this->esc($range) . '"/>';
            }
            $xml .= '</mergeCells>';
        }

        if ($sheet->autoFilter && $rowCount > 0 && $columnCount > 0) {
            $filterRow = $sheet->headerRowIndex ?? 1;
            $filterRange = $sheet->autoFilterRange
                ?? ('A' . $filterRow . ':' . Coordinate::columnIndexToName($columnCount) . max($filterRow, $rowCount));
            $xml .= $this->autoFilterXml($filterRange, $sheet->filterColumns);
        }

        foreach ($sheet->conditionalFormats as $priority => $rule) {
            $xml .= $this->conditionalFormattingXml($sheet, $rule, $priority + 1);
        }

        if ($sheet->dataValidations !== []) {
            $xml .= $this->dataValidationsXml($sheet->dataValidations);
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

        if ($pivotTables !== []) {
            $xml .= '<pivotTableParts count="' . count($pivotTables) . '">';
            foreach (array_values($pivotTables) as $index => $_pivot) {
                $xml .= '<pivotTablePart r:id="' . $this->pivotRelationshipId($hasImages, count($sheet->hyperlinks), $sheet->comments !== [], $index) . '"/>';
            }
            $xml .= '</pivotTableParts>';
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
            CellValue::TYPE_DATE => $this->dateCellXml($cellRef, $cell, $style),
            default => $this->inlineStringCellXml($cellRef, (string) $cell->displayValue(), $style),
        };
    }

    private function dateCellXml(string $cellRef, CellValue $cell, string $style): string
    {
        $value = $cell->value();
        try {
            $date = $value instanceof DateTimeInterface
                ? DateTimeImmutable::createFromInterface($value)
                : new DateTimeImmutable((string) $value);
        } catch (\Throwable $e) {
            throw new MnbExcelException('Invalid typed date cell value: ' . (string) $value, previous: $e);
        }

        $serial = $this->excelDateSerial($date, (bool) ($this->currentWorkbookMetadata['date1904'] ?? false));
        return '<c r="' . $cellRef . '"' . $style . '><v>' . $this->number($serial) . '</v></c>';
    }

    private function excelDateSerial(DateTimeInterface $date, bool $date1904): float
    {
        // Excel serials are timezone-free wall-clock values. Rebuilding the
        // components in UTC avoids daylight-saving offsets changing the serial.
        $utc = new DateTimeZone('UTC');
        $value = DateTimeImmutable::createFromFormat(
            '!Y-m-d H:i:s.u',
            $date->format('Y-m-d H:i:s.u'),
            $utc
        );
        if (!$value instanceof DateTimeImmutable) {
            throw new MnbExcelException('Unable to normalize typed date cell value.');
        }
        $base = new DateTimeImmutable(
            $date1904 ? '1904-01-01 00:00:00.000000' : '1899-12-31 00:00:00.000000',
            $utc
        );
        $seconds = (float) ($value->format('U.u')) - (float) ($base->format('U.u'));
        $serial = $seconds / 86400;
        if (!$date1904 && $serial >= 60) {
            $serial += 1; // Excel's historic fake 1900-02-29 day.
        }
        return $serial;
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
            $alignment = $this->alignmentXml($style['alignment'] ?? null);
            $protection = $this->protectionXml($style['protection'] ?? null);
            $xfs .= '<xf numFmtId="' . $numFmtId . '" fontId="' . $id . '" fillId="' . ($id + 1) . '" borderId="' . $id . '" xfId="0" applyFont="1" applyFill="1" applyBorder="1"'
                . ($numFmtId > 0 ? ' applyNumberFormat="1"' : '')
                . ($alignment !== '' ? ' applyAlignment="1"' : '')
                . ($protection !== '' ? ' applyProtection="1"' : '')
                . '>';
            if ($alignment !== '') {
                $xfs .= $alignment;
            }
            if ($protection !== '') {
                $xfs .= $protection;
            }
            $xfs .= '</xf>';
            unset($isCustom);
        }

        $fonts .= '</fonts>';
        $fills .= '</fills>';
        $borders .= '</borders>';
        $xfs .= '</cellXfs>';

        $dxfs = '<dxfs count="' . count($this->differentialStyles) . '">';
        foreach ($this->differentialStyles as $style) {
            $dxfs .= $this->dxfXml($style, $customFormats);
        }
        $dxfs .= '</dxfs>';

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            . $numFmtXml
            . $fonts
            . $fills
            . $borders
            . '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            . $xfs
            . '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            . $dxfs
            . '</styleSheet>';
    }

    /** @return array<string,int> */
    private function customFormats(): array
    {
        $formats = [];
        $nextId = 164;
        foreach (array_merge(array_values($this->styleEntries), array_values($this->differentialStyles)) as $style) {
            $format = $this->formatCode($style['format'] ?? null);
            if ($format === null || $this->builtinNumFmtId($format) !== null) {
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
        $builtin = $this->builtinNumFmtId($format);
        if ($builtin !== null) {
            return [$builtin, false];
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
            'general' => null,
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

    private function builtinNumFmtId(string $format): ?int
    {
        return match ($format) {
            'General' => 0, '@' => 49, '0' => 1, '0.00' => 2, '#,##0' => 3, '#,##0.00' => 4,
            '0%' => 9, '0.00%' => 10, '0.00E+00' => 11, '# ?/?' => 12, '# ??/??' => 13,
            'm/d/yy' => 14, 'd-mmm-yy' => 15, 'd-mmm' => 16, 'mmm-yy' => 17,
            'h:mm AM/PM' => 18, 'h:mm:ss AM/PM' => 19, 'h:mm' => 20, 'h:mm:ss' => 21,
            'm/d/yy h:mm' => 22, 'mm:ss' => 45, '[h]:mm:ss' => 46, 'mmss.0' => 47,
            default => null,
        };
    }

    private function fontXml(array $style): string
    {
        $font = is_array($style['font'] ?? null) ? $style['font'] : [];
        $fontSize = isset($font['size']) && is_numeric($font['size']) ? (float) $font['size'] : 11.0;
        $fontName = isset($font['name']) ? (string) $font['name'] : 'Calibri';

        $xml = '<font>';
        foreach (['bold' => 'b', 'italic' => 'i', 'strike' => 'strike', 'outline' => 'outline', 'shadow' => 'shadow', 'condense' => 'condense', 'extend' => 'extend'] as $key => $tag) {
            if (($font[$key] ?? false) === true) {
                $xml .= '<' . $tag . '/>';
            }
        }
        if (array_key_exists('underline', $font) && $font['underline'] !== false && $font['underline'] !== 'none') {
            $underline = $font['underline'] === true ? 'single' : (string) $font['underline'];
            $xml .= $underline === 'single' ? '<u/>' : '<u val="' . $this->esc($underline) . '"/>';
        }
        if (isset($font['vertical_align']) && $font['vertical_align'] !== '') {
            $xml .= '<vertAlign val="' . $this->esc((string) $font['vertical_align']) . '"/>';
        }
        $xml .= '<sz val="' . $this->number($fontSize) . '"/>';
        if (isset($font['color'])) {
            $xml .= $this->colorElement('color', $font['color'], 'FF000000');
        }
        $xml .= '<name val="' . $this->esc($fontName) . '"/>';
        foreach (['family', 'charset', 'scheme'] as $key) {
            if (isset($font[$key]) && $font[$key] !== '') {
                $xml .= '<' . $key . ' val="' . $this->esc((string) $font[$key]) . '"/>';
            }
        }
        return $xml . '</font>';
    }

    private function fillXml(array $style): string
    {
        $fill = $style['fill'] ?? false;
        if ($fill === false || $fill === null || $fill === '') {
            return '<fill><patternFill patternType="none"/></fill>';
        }
        if (!is_array($fill)) {
            $fill = ['type' => 'pattern', 'pattern' => 'solid', 'foreground' => $fill];
        }

        if (($fill['type'] ?? 'pattern') === 'gradient') {
            $attributes = '';
            $gradientType = (string) ($fill['gradient_type'] ?? 'linear');
            if ($gradientType !== 'linear') {
                $attributes .= ' type="' . $this->esc($gradientType) . '"';
            }
            foreach (['degree', 'left', 'right', 'top', 'bottom'] as $key) {
                if (isset($fill[$key]) && is_numeric($fill[$key])) {
                    $attributes .= ' ' . $key . '="' . $this->number((float) $fill[$key]) . '"';
                }
            }
            $xml = '<fill><gradientFill' . $attributes . '>';
            foreach ((array) ($fill['stops'] ?? []) as $stop) {
                if (!is_array($stop) || !isset($stop['color'])) {
                    continue;
                }
                $xml .= '<stop position="' . $this->number((float) ($stop['position'] ?? 0)) . '">' . $this->colorElement('color', $stop['color'], 'FFFFFFFF') . '</stop>';
            }
            return $xml . '</gradientFill></fill>';
        }

        $pattern = (string) ($fill['pattern'] ?? 'solid');
        $xml = '<fill><patternFill patternType="' . $this->esc($pattern) . '">';
        if (isset($fill['foreground'])) {
            $xml .= $this->colorElement('fgColor', $fill['foreground'], 'FFFFFFFF');
        }
        if (isset($fill['background'])) {
            $xml .= $this->colorElement('bgColor', $fill['background'], 'FFFFFFFF');
        } elseif ($pattern === 'solid') {
            $xml .= '<bgColor indexed="64"/>';
        }
        return $xml . '</patternFill></fill>';
    }

    private function borderXml(array $style): string
    {
        $border = $style['border'] ?? false;
        if ($border === false || !is_array($border)) {
            return '<border><left/><right/><top/><bottom/><diagonal/></border>';
        }

        $attributes = '';
        foreach (['diagonal_up' => 'diagonalUp', 'diagonal_down' => 'diagonalDown', 'outline' => 'outline'] as $key => $attribute) {
            if (array_key_exists($key, $border)) {
                $attributes .= ' ' . $attribute . '="' . ((bool) $border[$key] ? '1' : '0') . '"';
            }
        }
        $xml = '<border' . $attributes . '>';
        foreach (['left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal', 'start', 'end'] as $side) {
            $definition = $border[$side] ?? null;
            if (!is_array($definition) || $definition === []) {
                $xml .= '<' . $side . '/>';
                continue;
            }
            $styleName = (string) ($definition['style'] ?? 'thin');
            $xml .= '<' . $side . ($styleName !== '' && $styleName !== 'none' ? ' style="' . $this->esc($styleName) . '"' : '') . '>';
            if (isset($definition['color'])) {
                $xml .= $this->colorElement('color', $definition['color'], 'FFD0D7DE');
            }
            $xml .= '</' . $side . '>';
        }
        return $xml . '</border>';
    }

    private function alignmentXml(mixed $alignment): string
    {
        if ($alignment === null || $alignment === false || $alignment === []) {
            return '';
        }
        if (is_string($alignment)) {
            $alignment = ['horizontal' => $alignment];
        }
        if (!is_array($alignment)) {
            return '';
        }

        $attributes = [];
        foreach (['horizontal', 'vertical'] as $key) {
            if (isset($alignment[$key]) && $alignment[$key] !== '') {
                $attributes[$key] = (string) $alignment[$key];
            }
        }
        foreach (['wrap_text' => 'wrapText', 'shrink_to_fit' => 'shrinkToFit', 'justify_last_line' => 'justifyLastLine'] as $key => $attribute) {
            if (array_key_exists($key, $alignment)) {
                $attributes[$attribute] = (bool) $alignment[$key] ? '1' : '0';
            }
        }
        foreach (['text_rotation' => 'textRotation', 'indent' => 'indent', 'relative_indent' => 'relativeIndent', 'reading_order' => 'readingOrder'] as $key => $attribute) {
            if (isset($alignment[$key]) && is_numeric($alignment[$key])) {
                $attributes[$attribute] = (string) (int) $alignment[$key];
            }
        }
        return $attributes === [] ? '' : '<alignment' . $this->xmlAttributes($attributes) . '/>';
    }

    private function protectionXml(mixed $protection): string
    {
        if (!is_array($protection) || $protection === []) {
            return '';
        }
        $attributes = [];
        foreach (['locked', 'hidden'] as $key) {
            if (array_key_exists($key, $protection)) {
                $attributes[$key] = (bool) $protection[$key] ? '1' : '0';
            }
        }
        return $attributes === [] ? '' : '<protection' . $this->xmlAttributes($attributes) . '/>';
    }

    private function colorElement(string $tag, mixed $color, string $fallback): string
    {
        if (is_string($color)) {
            return '<' . $tag . ' rgb="' . $this->normalizeArgb($color, $fallback) . '"/>';
        }
        if (!is_array($color)) {
            return '<' . $tag . ' rgb="' . $fallback . '"/>';
        }
        $attributes = [];
        if (isset($color['rgb']) && is_string($color['rgb'])) {
            $attributes['rgb'] = $this->normalizeArgb($color['rgb'], $fallback);
        }
        foreach (['indexed', 'theme', 'tint', 'auto'] as $key) {
            if (array_key_exists($key, $color)) {
                $attributes[$key] = is_bool($color[$key]) ? ($color[$key] ? '1' : '0') : (string) $color[$key];
            }
        }
        if ($attributes === []) {
            $attributes['rgb'] = $fallback;
        }
        return '<' . $tag . $this->xmlAttributes($attributes) . '/>';
    }

    /** @param array<string,int> $customFormats */
    private function dxfXml(array $style, array $customFormats): string
    {
        $xml = '<dxf>';
        if (isset($style['font'])) {
            $xml .= $this->fontXml($style);
        }
        if (isset($style['fill'])) {
            $xml .= $this->fillXml($style);
        }
        if (isset($style['border'])) {
            $xml .= $this->borderXml($style);
        }
        $format = $this->formatCode($style['format'] ?? null);
        if ($format !== null) {
            $numFmtId = $format === '@' ? 49 : ($customFormats[$format] ?? 0);
            $xml .= '<numFmt numFmtId="' . $numFmtId . '" formatCode="' . $this->esc($format) . '"/>';
        }
        $alignment = $this->alignmentXml($style['alignment'] ?? null);
        if ($alignment !== '') {
            $xml .= $alignment;
        }
        return $xml . '</dxf>';
    }

    /** @param list<array<string,mixed>> $columns */
    private function autoFilterXml(string $range, array $columns): string
    {
        $xml = '<autoFilter ref="' . $this->esc($range) . '">';
        [$start] = explode(':', $range, 2);
        [$startColumn] = Coordinate::splitCellRef($start);
        foreach ($columns as $criteria) {
            $column = (int) ($criteria['column'] ?? 0);
            if ($column < $startColumn) {
                continue;
            }
            $colId = $column - $startColumn;
            $type = (string) ($criteria['type'] ?? 'values');
            $xml .= '<filterColumn colId="' . $colId . '">';
            if ($type === 'values') {
                $xml .= '<filters' . ((bool) ($criteria['include_blank'] ?? false) ? ' blank="1"' : '') . '>';
                foreach ((array) ($criteria['values'] ?? []) as $value) {
                    $xml .= '<filter val="' . $this->esc((string) $value) . '"/>';
                }
                $xml .= '</filters>';
            } elseif ($type === 'custom') {
                $items = (array) ($criteria['filters'] ?? [$criteria]);
                $xml .= '<customFilters' . ((bool) ($criteria['and'] ?? false) ? ' and="1"' : '') . '>';
                foreach ($items as $item) {
                    if (!is_array($item) || !array_key_exists('value', $item)) {
                        continue;
                    }
                    $operator = (string) ($item['operator'] ?? 'equal');
                    $xml .= '<customFilter operator="' . $this->esc($operator) . '" val="' . $this->esc((string) $item['value']) . '"/>';
                }
                $xml .= '</customFilters>';
            } elseif ($type === 'top10') {
                $xml .= '<top10 val="' . max(1, (int) ($criteria['value'] ?? 10)) . '"'
                    . ((bool) ($criteria['percent'] ?? false) ? ' percent="1"' : '')
                    . ((bool) ($criteria['bottom'] ?? false) ? ' top="0"' : '') . '/>';
            } elseif ($type === 'dynamic') {
                $xml .= '<dynamicFilter type="' . $this->esc((string) ($criteria['dynamic_type'] ?? $criteria['value'] ?? 'today')) . '"/>';
            } elseif ($type === 'color') {
                $xml .= '<colorFilter dxfId="' . max(0, (int) ($criteria['dxf_id'] ?? 0)) . '"'
                    . ((bool) ($criteria['cell_color'] ?? true) ? ' cellColor="1"' : ' cellColor="0"') . '/>';
            }
            $xml .= '</filterColumn>';
        }
        return $xml . '</autoFilter>';
    }

    /** @param array<string,mixed> $rule */
    private function conditionalFormattingXml(WorksheetData $sheet, array $rule, int $priority): string
    {
        $type = (string) ($rule['type'] ?? 'expression');
        $range = (string) ($rule['range'] ?? 'A1');
        $dxfId = null;
        if (array_key_exists('style', $rule)) {
            $dxfId = $this->differentialStyleIdFor($this->resolveStyle($sheet, $rule['style']));
        }
        $attrs = ' priority="' . $priority . '"';
        if ($dxfId !== null) {
            $attrs .= ' dxfId="' . $dxfId . '"';
        }
        $formulaXml = '';
        foreach ((array) ($rule['formulas'] ?? []) as $formula) {
            $formulaXml .= '<formula>' . $this->esc(ltrim((string) $formula, '=')) . '</formula>';
        }

        if ($type === 'cell_is') {
            $operator = (string) ($rule['operator'] ?? 'equal');
            $body = '<cfRule type="cellIs" operator="' . $this->esc($operator) . '"' . $attrs . '>' . $formulaXml . '</cfRule>';
        } elseif ($type === 'expression') {
            $body = '<cfRule type="expression"' . $attrs . '>' . $formulaXml . '</cfRule>';
        } elseif ($type === 'color_scale') {
            $colors = array_values((array) ($rule['colors'] ?? ['#F8696B', '#FFEB84', '#63BE7B']));
            if (count($colors) < 2) {
                $colors = ['#F8696B', '#63BE7B'];
            }
            $cfvos = count($colors) >= 3
                ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>'
                : '<cfvo type="min"/><cfvo type="max"/>';
            $colorXml = '';
            foreach (array_slice($colors, 0, 3) as $color) {
                $colorXml .= '<color rgb="' . $this->normalizeArgb((string) $color, 'FF638EC6') . '"/>';
            }
            $body = '<cfRule type="colorScale"' . $attrs . '><colorScale>' . $cfvos . $colorXml . '</colorScale></cfRule>';
        } elseif ($type === 'data_bar') {
            $body = '<cfRule type="dataBar"' . $attrs . '><dataBar showValue="' . ((bool) ($rule['show_value'] ?? true) ? '1' : '0') . '">'
                . '<cfvo type="min"/><cfvo type="max"/><color rgb="' . $this->normalizeArgb((string) ($rule['color'] ?? '#638EC6'), 'FF638EC6') . '"/>'
                . '</dataBar></cfRule>';
        } elseif ($type === 'icon_set') {
            $iconSet = (string) ($rule['icon_set'] ?? '3TrafficLights1');
            $body = '<cfRule type="iconSet"' . $attrs . '><iconSet iconSet="' . $this->esc($iconSet) . '" showValue="' . ((bool) ($rule['show_value'] ?? true) ? '1' : '0') . '">'
                . '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>'
                . '</iconSet></cfRule>';
        } elseif ($type === 'top10') {
            $body = '<cfRule type="top10" rank="' . max(1, (int) ($rule['rank'] ?? 10)) . '"'
                . ((bool) ($rule['percent'] ?? false) ? ' percent="1"' : '')
                . ((bool) ($rule['bottom'] ?? false) ? ' bottom="1"' : '') . $attrs . '/>';
        } elseif ($type === 'duplicate_values' || $type === 'unique_values') {
            $body = '<cfRule type="' . ($type === 'duplicate_values' ? 'duplicateValues' : 'uniqueValues') . '"' . $attrs . '/>';
        } elseif ($type === 'contains_text') {
            $text = (string) ($rule['text'] ?? '');
            $body = '<cfRule type="containsText" operator="containsText" text="' . $this->esc($text) . '"' . $attrs . '>'
                . ($formulaXml !== '' ? $formulaXml : '<formula>NOT(ISERROR(SEARCH("' . $this->esc(str_replace('"', '""', $text)) . '",' . explode(':', $range)[0] . ')))</formula>')
                . '</cfRule>';
        } else {
            $body = '<cfRule type="timePeriod" timePeriod="' . $this->esc((string) ($rule['period'] ?? 'today')) . '"' . $attrs . '>' . $formulaXml . '</cfRule>';
        }
        return '<conditionalFormatting sqref="' . $this->esc($range) . '">' . $body . '</conditionalFormatting>';
    }

    /** @param list<array<string,mixed>> $validations */
    private function dataValidationsXml(array $validations): string
    {
        $xml = '<dataValidations count="' . count($validations) . '">';
        foreach ($validations as $validation) {
            $type = (string) ($validation['type'] ?? 'custom');
            $operator = isset($validation['operator']) ? ' operator="' . $this->esc((string) $validation['operator']) . '"' : '';
            $attrs = ' type="' . $this->esc($type === 'text_length' ? 'textLength' : $type) . '"'
                . $operator
                . ' allowBlank="' . ((bool) ($validation['allow_blank'] ?? true) ? '1' : '0') . '"'
                . ' showInputMessage="' . ((bool) ($validation['show_input'] ?? true) ? '1' : '0') . '"'
                . ' showErrorMessage="' . ((bool) ($validation['show_error'] ?? true) ? '1' : '0') . '"';
            foreach (['prompt_title' => 'promptTitle', 'prompt' => 'prompt', 'error_title' => 'errorTitle', 'error' => 'error', 'error_style' => 'errorStyle'] as $key => $attribute) {
                if (isset($validation[$key]) && $validation[$key] !== null && $validation[$key] !== '') {
                    $attrs .= ' ' . $attribute . '="' . $this->esc((string) $validation[$key]) . '"';
                }
            }
            $xml .= '<dataValidation' . $attrs . ' sqref="' . $this->esc((string) $validation['range']) . '">';
            if ($type === 'list' && isset($validation['values'])) {
                $list = implode(',', array_map(static fn (mixed $value): string => str_replace('"', '""', (string) $value), (array) $validation['values']));
                $xml .= '<formula1>"' . $this->esc($list) . '"</formula1>';
            } elseif (isset($validation['formula1'])) {
                $xml .= '<formula1>' . $this->esc(ltrim((string) $validation['formula1'], '=')) . '</formula1>';
            }
            if (isset($validation['formula2'])) {
                $xml .= '<formula2>' . $this->esc(ltrim((string) $validation['formula2'], '=')) . '</formula2>';
            }
            $xml .= '</dataValidation>';
        }
        return $xml . '</dataValidations>';
    }

    /** @return array{string,string} */
    private function normalizeRange(string $range): array
    {
        $range = strtoupper(str_replace('$', '', trim($range)));
        if (preg_match('/^([A-Z]{1,3}[1-9][0-9]*)(?::([A-Z]{1,3}[1-9][0-9]*))?$/', $range, $match) !== 1) {
            throw new MnbExcelException('Invalid cell range: ' . $range);
        }
        return [$match[1], $match[2] ?? $match[1]];
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

    /** @param list<array<string,mixed>> $pivotTables */
    private function sheetRelationships(int $sheetNumber, WorksheetData $sheet, bool $hasImages, array $pivotTables = []): string
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

        foreach (array_values($pivotTables) as $index => $pivot) {
            $number = (int) $pivot['pivotNumber'];
            $xml .= '<Relationship Id="' . $this->pivotRelationshipId($hasImages, count($sheet->hyperlinks), $sheet->comments !== [], $index) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable' . $number . '.xml"/>';
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

    private function pivotRelationshipId(bool $hasImages, int $hyperlinkCount, bool $hasComments, int $pivotIndex): string
    {
        return 'rIdPivot' . ($pivotIndex + 1);
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

    /** @param list<array{mediaName:string}> $images @param list<array<string,mixed>> $charts */
    private function drawingRels(array $images, array $charts = []): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        foreach ($images as $index => $image) {
            $xml .= '<Relationship Id="rId' . ($index + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' . $this->esc($image['mediaName']) . '"/>';
        }
        $offset = count($images);
        foreach ($charts as $index => $chart) {
            $xml .= '<Relationship Id="rId' . ($offset + $index + 1) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/' . $this->esc((string) $chart['partName']) . '"/>';
        }
        return $xml . '</Relationships>';
    }

    /** @param list<array{cell:string,width:int,height:int,name:string}> $images @param list<array<string,mixed>> $charts */
    private function drawingXml(array $images, array $charts = []): string
    {
        $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">';

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

        $relationshipOffset = count($images);
        foreach ($charts as $index => $chart) {
            [$fromColumn, $fromRow] = Coordinate::splitCellRef((string) $chart['from']);
            [$toColumn, $toRow] = Coordinate::splitCellRef((string) $chart['to']);
            $shapeId = count($images) + $index + 1;
            $relationshipId = $relationshipOffset + $index + 1;
            $xml .= '<xdr:twoCellAnchor editAs="twoCell">'
                . '<xdr:from><xdr:col>' . ($fromColumn - 1) . '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' . ($fromRow - 1) . '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                . '<xdr:to><xdr:col>' . ($toColumn - 1) . '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' . ($toRow - 1) . '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                . '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' . $shapeId . '" name="' . $this->esc((string) ($chart['title'] ?? ('Chart ' . $shapeId))) . '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>'
                . '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                . '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId' . $relationshipId . '"/></a:graphicData></a:graphic>'
                . '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';
        }

        return $xml . '</xdr:wsDr>';
    }

    /** @param array<string,mixed> $chart */
    private function chartXml(array $chart, string $sheetName): string
    {
        $type = (string) ($chart['type'] ?? 'column');
        $seriesXml = '';
        foreach (array_values((array) ($chart['series'] ?? [])) as $index => $series) {
            $name = (string) ($series['name'] ?? ('Series ' . ($index + 1)));
            $values = $this->chartFormula((string) $series['values'], $sheetName);
            $categories = isset($series['categories']) ? $this->chartFormula((string) $series['categories'], $sheetName) : null;
            $seriesXml .= '<c:ser><c:idx val="' . $index . '"/><c:order val="' . $index . '"/>'
                . '<c:tx><c:v>' . $this->esc($name) . '</c:v></c:tx>';
            if ($categories !== null && $categories !== '') {
                $seriesXml .= ($type === 'scatter'
                    ? '<c:xVal><c:numRef><c:f>' . $this->esc($categories) . '</c:f></c:numRef></c:xVal>'
                    : '<c:cat><c:strRef><c:f>' . $this->esc($categories) . '</c:f></c:strRef></c:cat>');
            }
            $seriesXml .= ($type === 'scatter'
                ? '<c:yVal><c:numRef><c:f>' . $this->esc($values) . '</c:f></c:numRef></c:yVal>'
                : '<c:val><c:numRef><c:f>' . $this->esc($values) . '</c:f></c:numRef></c:val>');
            $seriesXml .= '</c:ser>';
        }

        $varyColors = (bool) ($chart['vary_colors'] ?? false) ? '1' : '0';
        $axisIds = '<c:axId val="48650112"/><c:axId val="48672768"/>';
        $plot = match ($type) {
            'bar' => '<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="' . $varyColors . '"/>' . $seriesXml . $axisIds . '</c:barChart>',
            'line' => '<c:lineChart><c:grouping val="standard"/><c:varyColors val="' . $varyColors . '"/>' . $seriesXml . '<c:marker val="1"/>' . $axisIds . '</c:lineChart>',
            'area' => '<c:areaChart><c:grouping val="standard"/><c:varyColors val="' . $varyColors . '"/>' . $seriesXml . $axisIds . '</c:areaChart>',
            'pie' => '<c:pieChart><c:varyColors val="1"/>' . $seriesXml . '<c:firstSliceAng val="0"/></c:pieChart>',
            'doughnut' => '<c:doughnutChart><c:varyColors val="1"/>' . $seriesXml . '<c:holeSize val="50"/></c:doughnutChart>',
            'scatter' => '<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="' . $varyColors . '"/>' . $seriesXml . $axisIds . '</c:scatterChart>',
            default => '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="' . $varyColors . '"/>' . $seriesXml . $axisIds . '</c:barChart>',
        };
        $axes = in_array($type, ['pie', 'doughnut'], true)
            ? ''
            : ($type === 'scatter' ? $this->scatterChartAxesXml() : $this->chartAxesXml());
        $legendPosition = match (strtolower((string) ($chart['legend'] ?? 'right'))) {
            'left' => 'l', 'top' => 't', 'bottom' => 'b', 'top_right' => 'tr', default => 'r',
        };

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:style val="' . max(1, min(48, (int) ($chart['style'] ?? 10))) . '"/>'
            . '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>' . $this->esc((string) ($chart['title'] ?? 'Chart')) . '</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>'
            . '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' . $plot . $axes . '</c:plotArea>'
            . '<c:legend><c:legendPos val="' . $legendPosition . '"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/>'
            . '</c:chart></c:chartSpace>';
    }

    private function chartAxesXml(): string
    {
        return '<c:catAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>'
            . '<c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
    }

    private function scatterChartAxesXml(): string
    {
        return '<c:valAx><c:axId val="48650112"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="48672768"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
            . '<c:valAx><c:axId val="48672768"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="48650112"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>';
    }

    private function chartFormula(string $range, string $sheetName): string
    {
        $range = trim($range);
        if ($range === '') {
            return '';
        }
        if (str_contains($range, '!')) {
            return ltrim($range, '=');
        }
        $escapedSheet = str_replace("'", "''", $sheetName);
        return "'" . $escapedSheet . "'!" . ltrim($range, '=');
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

    /** @param array<string,mixed> $metadata @return array<string,mixed>|null */
    private function sheetProtectionOptions(string $sheetName, array $metadata): ?array
    {
        $protections = $metadata['_mnb_sheet_protection'] ?? null;
        if (!is_array($protections)) {
            return null;
        }
        $settings = $protections[$sheetName] ?? $protections['*'] ?? null;
        return is_array($settings) && (string) ($settings['password'] ?? '') !== '' ? $settings : null;
    }

    /** @param array<string,string> $attributes */
    private function xmlAttributes(array $attributes): string
    {
        $xml = '';
        foreach ($attributes as $name => $value) {
            $xml .= ' ' . $name . '="' . $this->esc($value) . '"';
        }
        return $xml;
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
