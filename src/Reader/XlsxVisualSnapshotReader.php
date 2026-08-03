<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use DateTimeInterface;
use Mnb\PHPExcel\Core\RichText;
use Mnb\PHPExcel\Core\StyleNormalizer;
use Mnb\PHPExcel\Reader\State\FormulaResult;
use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Snapshot\VisualSnapshot;
use Mnb\PHPExcel\Snapshot\VisualSnapshotReaderInterface;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

/**
 * Reads values, styles, and worksheet presentation settings into a stable,
 * sparse, versioned snapshot that can be written back without application-side
 * coordinate conversion.
 */
final class XlsxVisualSnapshotReader implements VisualSnapshotReaderInterface
{
    public function __construct(
        private readonly XlsxWorkbookResolver $resolver = new XlsxWorkbookResolver(),
        private readonly XlsxReader $reader = new XlsxReader(),
    ) {
    }

    /** @param array<string,mixed> $options @return array<string,mixed> */
    public function visualSnapshot(string $path, array $options = []): array
    {
        [$readablePath, $temporary] = $this->readablePath($path, $options);
        try {
            $realPath = realpath($readablePath);
            if ($realPath === false) {
                throw new MnbExcelException('Invalid XLSX path: ' . $path);
            }
            $zip = new ZipArchive();
            if ($zip->open($realPath) !== true) {
                throw new MnbExcelException('Unable to open XLSX package for visual snapshot.');
            }
            try {
                $stylesXml = $zip->getFromName('xl/styles.xml');
                $workbookXml = $zip->getFromName('xl/workbook.xml');
                $styleMap = XlsxStyleMap::fromXml(is_string($stylesXml) ? $stylesXml : null);
                $sheets = $this->resolver->sheets($realPath);
                $active = $this->resolver->activeSheet($realPath);
                $date1904 = is_string($workbookXml)
                    && preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?workbookPr\b[^>]*date1904\s*=\s*(?:"(?:1|true)"|\'(?:1|true)\')/i', $workbookXml) === 1;

                $styleTable = [];
                $styleIds = [];
                $snapshotSheets = [];
                $warnings = [];
                $maximumCells = max(1, (int) ($options['max_cells'] ?? 1_000_000));
                $seenCells = 0;

                foreach ($sheets as $sheet) {
                    if (!(bool) ($sheet['exists'] ?? false)) {
                        $warnings[] = 'Worksheet XML is missing for ' . (string) $sheet['name'] . '.';
                        continue;
                    }
                    $sheetXml = $zip->getFromName((string) $sheet['path']);
                    if (!is_string($sheetXml)) {
                        continue;
                    }
                    $declaredDimension = $this->dimension($sheetXml);
                    $cellRefs = $this->cellReferences($sheetXml);
                    $seenCells += count($cellRefs);
                    if ($seenCells > $maximumCells) {
                        throw new MnbExcelException('Visual snapshot exceeds max_cells.');
                    }

                    $values = $cellRefs === [] ? [] : $this->reader->readCells(
                        $realPath,
                        array_keys($cellRefs),
                        (string) $sheet['name'],
                        array_replace($options, [
                            'formula_cells' => 'both',
                            'return_datetime' => true,
                            'format_dates' => true,
                            'include_hidden_rows' => true,
                        ])
                    );
                    $metadata = $this->reader->readSheetMetadata($realPath, (string) $sheet['name'], $options);
                    $richObjects = is_array($metadata['rich_text_objects'] ?? null) ? $metadata['rich_text_objects'] : [];
                    $cells = [];
                    foreach ($cellRefs as $coordinate => $styleIndex) {
                        $value = $values[$coordinate] ?? null;
                        $style = StyleNormalizer::normalize($styleMap->styleForIndex($styleIndex));
                        $styleId = $this->registerStyle($style, $styleTable, $styleIds);
                        $richText = $richObjects[$coordinate] ?? null;
                        $cells[$coordinate] = $this->encodeCell(
                            $value,
                            $richText instanceof RichText ? $richText : null,
                            $style,
                            $styleMap->isDateStyle($styleIndex),
                            $styleId,
                        );
                    }

                    $mergedCells = array_values((array) ($metadata['merged_cells'] ?? []));
                    $dimension = $this->actualDimension(array_keys($cellRefs), $mergedCells);
                    $layout = $this->layout($sheetXml);
                    if ($declaredDimension !== $dimension) {
                        $layout['declared_dimension'] = $declaredDimension;
                    }
                    $snapshotSheets[] = [
                        'index' => (int) $sheet['index'],
                        'name' => (string) $sheet['name'],
                        'state' => (string) ($sheet['state'] ?? 'visible'),
                        'dimension' => $dimension,
                        'cells' => $cells,
                        'layout' => $layout,
                        'merged_cells' => $mergedCells,
                        'comments' => $this->normalizeComments((array) ($metadata['comments'] ?? [])),
                        'hyperlinks' => $this->normalizeHyperlinks((array) ($metadata['hyperlinks'] ?? [])),
                        'images' => array_values(array_filter((array) ($metadata['images'] ?? []), 'is_array')),
                        'conditional_formats' => $this->conditionalFormats($sheetXml),
                        'data_validations' => $this->dataValidations($sheetXml),
                        'capabilities' => [
                            'values' => 'available',
                            'styles' => 'available',
                            'layout' => 'available',
                            'conditional_formats' => 'available',
                            'data_validations' => 'available',
                            'images' => 'inventory_only',
                        ],
                    ];
                }
            } finally {
                $zip->close();
            }

            $documentMetadata = [];
            try {
                $meta = (new XlsxMetadataReader($this->resolver))->metaInfo($realPath, ['profile' => 'quick']);
                foreach (['document', 'revision', 'application', 'custom_properties'] as $section) {
                    if (isset($meta[$section]) && is_array($meta[$section])) {
                        $documentMetadata[$section] = $meta[$section];
                    }
                }
            } catch (\Throwable $error) {
                $warnings[] = 'Document metadata could not be added to the visual snapshot: ' . $error->getMessage();
            }

            $snapshot = [
                'schema' => VisualSnapshot::SCHEMA,
                'schema_version' => VisualSnapshot::VERSION,
                'format' => 'xlsx',
                'source' => [
                    'file_name' => basename($path),
                    'size_bytes' => filesize($path) ?: 0,
                    'encrypted' => (new XlsxEncryption())->isEncryptedFile($path),
                ],
                'workbook' => [
                    'active_sheet' => (int) ($active['index'] ?? 1),
                    'active_sheet_name' => (string) ($active['name'] ?? ''),
                    'date1904' => $date1904,
                    'metadata' => $documentMetadata,
                ],
                'styles' => $styleTable,
                'sheets' => $snapshotSheets,
                'capabilities' => [
                    'values' => 'available',
                    'styles' => 'available',
                    'layout' => 'available',
                    'round_trip' => 'supported',
                ],
                'warnings' => $warnings,
            ];
            VisualSnapshot::assertValid($snapshot);
            return $snapshot;
        } finally {
            if ($temporary !== null) {
                @unlink($temporary);
            }
        }
    }

    /** @return array{0:string,1:?string} */
    private function readablePath(string $path, array $options): array
    {
        $encryption = new XlsxEncryption();
        if (!$encryption->isEncryptedFile($path)) {
            return [$path, null];
        }
        $password = (string) ($options['password'] ?? $options['xlsx_password'] ?? '');
        if ($password === '') {
            throw new MnbExcelException('A password is required to snapshot this encrypted XLSX file.');
        }
        $temporary = $encryption->decryptToTemporary($path, $password, $options);
        return [$temporary, $temporary];
    }

    private function dimension(string $xml): string
    {
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dimension\b([^>]*)\/?\s*>/i', $xml, $match) === 1) {
            $attrs = $this->attributes($match[1]);
            if (is_string($attrs['ref'] ?? null) && $attrs['ref'] !== '') {
                $ref = strtoupper(str_replace('$', '', $attrs['ref']));
                return str_contains($ref, ':') ? $ref : 'A1:' . $ref;
            }
        }
        $refs = array_keys($this->cellReferences($xml));
        if ($refs === []) {
            return 'A1:A1';
        }
        $maxColumn = 1;
        $maxRow = 1;
        foreach ($refs as $ref) {
            [$column, $row] = Coordinate::splitCellRef($ref);
            $maxColumn = max($maxColumn, $column);
            $maxRow = max($maxRow, $row);
        }
        return 'A1:' . Coordinate::columnIndexToName($maxColumn) . $maxRow;
    }

    /** @param list<string> $cellRefs @param list<mixed> $mergedCells */
    private function actualDimension(array $cellRefs, array $mergedCells): string
    {
        $minimumColumn = PHP_INT_MAX;
        $minimumRow = PHP_INT_MAX;
        $maximumColumn = 1;
        $maximumRow = 1;
        $references = $cellRefs;
        foreach ($mergedCells as $range) {
            if (!is_string($range) || trim($range) === '') {
                continue;
            }
            foreach (explode(':', strtoupper(str_replace('$', '', $range)), 2) as $reference) {
                $references[] = $reference;
            }
        }
        foreach ($references as $reference) {
            try {
                [$column, $row] = Coordinate::splitCellRef(strtoupper((string) $reference));
            } catch (\Throwable) {
                continue;
            }
            $minimumColumn = min($minimumColumn, $column);
            $minimumRow = min($minimumRow, $row);
            $maximumColumn = max($maximumColumn, $column);
            $maximumRow = max($maximumRow, $row);
        }
        if ($minimumColumn === PHP_INT_MAX) {
            return 'A1:A1';
        }
        $start = Coordinate::columnIndexToName($minimumColumn) . $minimumRow;
        $end = Coordinate::columnIndexToName($maximumColumn) . $maximumRow;
        return $start . ':' . $end;
    }

    /** @return array<string,?int> */
    private function cellReferences(string $xml): array
    {
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\b([^>]*)>/i', $xml, $matches, PREG_SET_ORDER);
        $result = [];
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $ref = strtoupper((string) ($attrs['r'] ?? ''));
            if ($ref === '') {
                continue;
            }
            try {
                Coordinate::splitCellRef($ref);
            } catch (\Throwable) {
                continue;
            }
            $result[$ref] = isset($attrs['s']) && ctype_digit((string) $attrs['s']) ? (int) $attrs['s'] : null;
        }
        return $result;
    }

    /** @param array<string,mixed> $style @param array<string,array<string,mixed>> $table @param array<string,string> $ids */
    private function registerStyle(array $style, array &$table, array &$ids): ?string
    {
        if ($style === []) {
            return null;
        }
        $key = json_encode($style, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE);
        if (!is_string($key)) {
            return null;
        }
        if (!isset($ids[$key])) {
            $id = 's' . (count($table) + 1);
            $ids[$key] = $id;
            $table[$id] = $style;
        }
        return $ids[$key];
    }

    /** @param array<string,mixed> $style @return array<string,mixed> */
    private function encodeCell(mixed $value, ?RichText $richText, array $style, bool $dateStyle, ?string $styleId): array
    {
        $cell = [];
        if ($styleId !== null) {
            $cell['style_id'] = $styleId;
        }
        if ($richText !== null) {
            $cell['type'] = 'rich_text';
            $cell['value'] = $richText->text();
            $cell['rich_text'] = $richText->jsonSerialize();
            return $cell;
        }
        if ($value instanceof FormulaResult) {
            $cell['type'] = 'formula';
            $cell['formula'] = ltrim($value->formula, '=');
            $cell['cached_value'] = $value->cachedValue;
            return $cell;
        }
        if ($value instanceof DateTimeInterface) {
            $cell['type'] = 'date';
            $cell['value'] = $value->format('Y-m-d\TH:i:sP');
            $cell['format'] = (string) ($style['format'] ?? 'yyyy-mm-dd');
            return $cell;
        }
        if ($dateStyle && is_string($value) && $this->looksLikeIsoDate($value)) {
            $cell['type'] = 'date';
            $cell['value'] = $value;
            $cell['format'] = (string) ($style['format'] ?? 'yyyy-mm-dd');
            return $cell;
        }
        if ($value === null) {
            $cell['type'] = 'blank';
            $cell['value'] = null;
        } elseif (is_bool($value)) {
            $cell['type'] = 'boolean';
            $cell['value'] = $value;
        } elseif (is_int($value) || is_float($value)) {
            $cell['type'] = 'number';
            $cell['value'] = $value;
        } else {
            $cell['type'] = 'text';
            $cell['value'] = (string) $value;
        }
        return $cell;
    }

    private function looksLikeIsoDate(string $value): bool
    {
        return preg_match('/^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/', trim($value)) === 1;
    }

    /** @return array<string,mixed> */
    private function layout(string $xml): array
    {
        $widths = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?col\b([^>]*)\/?\s*>/i', $xml, $columns, PREG_SET_ORDER);
        foreach ($columns as $column) {
            $attrs = $this->attributes($column[1]);
            if (!isset($attrs['min'], $attrs['max'], $attrs['width']) || !is_numeric($attrs['width'])) {
                continue;
            }
            for ($index = (int) $attrs['min']; $index <= (int) $attrs['max']; $index++) {
                $widths[Coordinate::columnIndexToName($index)] = (float) $attrs['width'];
            }
        }
        $heights = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?row\b([^>]*)>/i', $xml, $rows, PREG_SET_ORDER);
        foreach ($rows as $row) {
            $attrs = $this->attributes($row[1]);
            if (isset($attrs['r'], $attrs['ht']) && ctype_digit((string) $attrs['r']) && is_numeric($attrs['ht'])) {
                $heights[(int) $attrs['r']] = (float) $attrs['ht'];
            }
        }
        $freeze = ['rows' => 0, 'columns' => 0, 'top_left_cell' => null];
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?pane\b([^>]*)\/?\s*>/i', $xml, $pane) === 1) {
            $attrs = $this->attributes($pane[1]);
            if (($attrs['state'] ?? '') === 'frozen' || ($attrs['state'] ?? '') === 'frozenSplit') {
                $freeze = [
                    'rows' => isset($attrs['ySplit']) ? (int) round((float) $attrs['ySplit']) : 0,
                    'columns' => isset($attrs['xSplit']) ? (int) round((float) $attrs['xSplit']) : 0,
                    'top_left_cell' => isset($attrs['topLeftCell']) ? strtoupper($attrs['topLeftCell']) : null,
                ];
            }
        }
        $filter = null;
        if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?autoFilter\b([^>]*)>/i', $xml, $autoFilter) === 1) {
            $attrs = $this->attributes($autoFilter[1]);
            $filter = isset($attrs['ref']) ? strtoupper($attrs['ref']) : null;
        }
        return [
            'column_widths' => $widths,
            'row_heights' => $heights,
            'freeze_panes' => $freeze,
            'auto_filter_range' => $filter,
        ];
    }

    /** @return list<array<string,mixed>> */
    private function conditionalFormats(string $xml): array
    {
        $result = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?conditionalFormatting\b([^>]*)>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?conditionalFormatting\s*>/is', $xml, $blocks, PREG_SET_ORDER);
        foreach ($blocks as $block) {
            $blockAttrs = $this->attributes($block[1]);
            $range = strtoupper((string) ($blockAttrs['sqref'] ?? 'A1'));
            preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cfRule\b([^>]*)(?:\/>|>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?cfRule\s*>)/is', $block[2], $rules, PREG_SET_ORDER);
            foreach ($rules as $rule) {
                $attrs = $this->attributes($rule[1]);
                $body = (string) ($rule[2] ?? '');
                $type = (string) ($attrs['type'] ?? 'expression');
                $item = ['range' => $range, 'type' => $this->snakeType($type)];
                if (isset($attrs['operator'])) {
                    $item['operator'] = $attrs['operator'];
                }
                preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?formula\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?formula\s*>/is', $body, $formulas);
                if (($formulas[1] ?? []) !== []) {
                    $item['formulas'] = array_map([$this, 'xmlText'], $formulas[1]);
                }
                if ($type === 'colorScale') {
                    preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?color\b([^>]*)\/?\s*>/i', $body, $colors, PREG_SET_ORDER);
                    $item['colors'] = [];
                    foreach ($colors as $color) {
                        $colorAttrs = $this->attributes($color[1]);
                        if (isset($colorAttrs['rgb'])) {
                            $item['colors'][] = '#' . substr(strtoupper($colorAttrs['rgb']), -6);
                        }
                    }
                }
                if ($type === 'dataBar' && preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?color\b([^>]*)\/?\s*>/i', $body, $color) === 1) {
                    $colorAttrs = $this->attributes($color[1]);
                    if (isset($colorAttrs['rgb'])) {
                        $item['color'] = '#' . substr(strtoupper($colorAttrs['rgb']), -6);
                    }
                }
                if ($type === 'iconSet' && preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?iconSet\b([^>]*)>/i', $body, $icon) === 1) {
                    $iconAttrs = $this->attributes($icon[1]);
                    $item['icon_set'] = $iconAttrs['iconSet'] ?? '3TrafficLights1';
                }
                $result[] = $item;
            }
        }
        return $result;
    }

    /** @return list<array<string,mixed>> */
    private function dataValidations(string $xml): array
    {
        $result = [];
        preg_match_all('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dataValidation\b([^>]*)(?:\/>|>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?dataValidation\s*>)/is', $xml, $matches, PREG_SET_ORDER);
        foreach ($matches as $match) {
            $attrs = $this->attributes($match[1]);
            $body = (string) ($match[2] ?? '');
            $item = [
                'range' => strtoupper((string) ($attrs['sqref'] ?? 'A1')),
                'type' => $this->snakeType((string) ($attrs['type'] ?? 'custom')),
                'allow_blank' => $this->truthy($attrs['allowBlank'] ?? true),
                'show_input' => $this->truthy($attrs['showInputMessage'] ?? true),
                'show_error' => $this->truthy($attrs['showErrorMessage'] ?? true),
            ];
            foreach (['operator', 'promptTitle' => 'prompt_title', 'prompt' => 'prompt', 'errorTitle' => 'error_title', 'error' => 'error'] as $source => $target) {
                if (is_int($source)) {
                    $source = $target;
                }
                if (isset($attrs[$source])) {
                    $item[$target] = $attrs[$source];
                }
            }
            foreach (['formula1', 'formula2'] as $tag) {
                if (preg_match('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '\b[^>]*>(.*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?' . $tag . '\s*>/is', $body, $formula) === 1) {
                    $item[$tag] = $this->xmlText($formula[1]);
                }
            }
            $result[] = $item;
        }
        return $result;
    }

    /** @return list<array<string,mixed>> */
    private function normalizeComments(array $comments): array
    {
        $result = [];
        foreach ($comments as $comment) {
            if (!is_array($comment) || !is_string($comment['cell'] ?? null) || !is_string($comment['text'] ?? null)) {
                continue;
            }
            $result[] = [
                'cell' => strtoupper($comment['cell']),
                'author' => (string) ($comment['author'] ?? 'Imported'),
                'text' => $comment['text'],
            ];
        }
        return $result;
    }

    /** @return list<array<string,mixed>> */
    private function normalizeHyperlinks(array $hyperlinks): array
    {
        $result = [];
        foreach ($hyperlinks as $link) {
            if (!is_array($link)) {
                continue;
            }
            $cell = $link['cell'] ?? $link['ref'] ?? null;
            $url = $link['url'] ?? $link['target'] ?? null;
            if (!is_string($cell) || !is_string($url) || $url === '') {
                continue;
            }
            $item = ['cell' => strtoupper($cell), 'url' => $url];
            foreach (['display', 'tooltip'] as $key) {
                if (is_string($link[$key] ?? null)) {
                    $item[$key] = $link[$key];
                }
            }
            $result[] = $item;
        }
        return $result;
    }

    private function snakeType(string $value): string
    {
        return strtolower((string) preg_replace('/(?<!^)[A-Z]/', '_$0', $value));
    }

    /** @return array<string,string> */
    private function attributes(string $source): array
    {
        preg_match_all('/([A-Za-z_][A-Za-z0-9_.:-]*)\s*=\s*("[^"]*"|\'[^\']*\')/s', $source, $matches, PREG_SET_ORDER);
        $result = [];
        foreach ($matches as $match) {
            $name = str_contains($match[1], ':') ? substr($match[1], strrpos($match[1], ':') + 1) : $match[1];
            $result[$name] = html_entity_decode(substr($match[2], 1, -1), ENT_QUOTES | ENT_XML1, 'UTF-8');
        }
        return $result;
    }

    private function truthy(mixed $value): bool
    {
        if (is_bool($value)) {
            return $value;
        }
        return in_array(strtolower((string) $value), ['1', 'true', 'yes', 'on'], true);
    }

    private function xmlText(string $value): string
    {
        return html_entity_decode(strip_tags($value), ENT_QUOTES | ENT_XML1, 'UTF-8');
    }
}
