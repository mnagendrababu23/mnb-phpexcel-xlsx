<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Template;

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Support\Coordinate;
use Mnb\PHPExcel\Support\MnbExcelException;

/** Builds XLSX import templates without depending on the application facade. */
final class XlsxImportTemplateFactory
{
    /**
     * @param array<int|string,string|array<string,mixed>> $columns
     * @param array<string,mixed> $options
     */
    public function create(array $columns, array $options = []): WorkbookData
    {
        if ($columns === []) {
            throw new MnbExcelException('Import template columns cannot be empty.');
        }

        $headers = [];
        $definitions = [];
        foreach ($columns as $key => $definition) {
            if (is_array($definition)) {
                $name = trim((string) ($definition['header'] ?? $definition['name'] ?? $key));
                if ($name === '') {
                    throw new MnbExcelException('Every import template column requires a non-empty header.');
                }
                $headers[] = $name;
                $definitions[] = $definition + ['header' => $name];
                continue;
            }

            $name = trim((string) $definition);
            if ($name === '') {
                throw new MnbExcelException('Every import template column requires a non-empty header.');
            }
            $headers[] = $name;
            $definitions[] = ['header' => $name];
        }

        $instructions = trim((string) ($options['instructions'] ?? ''));
        $rows = [];
        $rowStyles = [];
        $mergeCells = [];
        if ($instructions !== '') {
            $rows[] = [$instructions];
            $rowStyles[1] = $options['instruction_style'] ?? [
                'font' => ['bold' => true, 'color' => '#1F1F1F'],
                'fill' => '#D9EAF7',
                'alignment' => ['wrap_text' => true, 'vertical' => 'center'],
            ];
            if (count($headers) > 1) {
                $mergeCells[] = 'A1:' . Coordinate::columnIndexToName(count($headers)) . '1';
            }
        }

        $headerRow = count($rows) + 1;
        $rows[] = $headers;
        $sampleRows = max(1, (int) ($options['sample_rows'] ?? 1));
        for ($i = 0; $i < $sampleRows; $i++) {
            $sample = [];
            foreach ($definitions as $definition) {
                $sample[] = $i === 0 ? ($definition['example'] ?? '') : '';
            }
            $rows[] = $sample;
        }

        $lastValidationRow = max(1000, (int) ($options['validation_rows'] ?? 10000));
        $dataStartRow = $headerRow + 1;
        $validations = [];
        $comments = [];
        foreach ($definitions as $index => $definition) {
            $column = Coordinate::columnIndexToName($index + 1);
            $range = $column . $dataStartRow . ':' . $column . $lastValidationRow;
            if (isset($definition['list']) && is_array($definition['list'])) {
                $validations[] = [
                    'range' => $range,
                    'type' => 'list',
                    'values' => array_values($definition['list']),
                    'allow_blank' => !((bool) ($definition['required'] ?? false)),
                    'prompt_title' => (string) $definition['header'],
                    'prompt' => $definition['description'] ?? null,
                    'error' => $definition['error'] ?? null,
                ];
            } elseif (isset($definition['validation']) && is_array($definition['validation'])) {
                $validation = $definition['validation'];
                $type = strtolower((string) ($validation['type'] ?? 'custom'));
                if (!in_array($type, ['list', 'whole', 'decimal', 'date', 'time', 'text_length', 'custom'], true)) {
                    throw new MnbExcelException('Unsupported import-template validation type: ' . $type);
                }
                unset($validation['type']);
                $validations[] = ['range' => $range, 'type' => $type] + $validation;
            }

            if ((bool) ($definition['required'] ?? false)) {
                $text = 'Required field';
                if (isset($definition['description']) && trim((string) $definition['description']) !== '') {
                    $text .= ': ' . trim((string) $definition['description']);
                }
                $comments[] = [
                    'cell' => $column . $headerRow,
                    'author' => 'MNB PHPExcel',
                    'text' => $text,
                ];
            }
        }

        $widths = [];
        foreach ($headers as $index => $header) {
            $widths[$index + 1] = max(
                (float) ($options['minimum_column_width'] ?? 12),
                min((float) ($options['maximum_column_width'] ?? 45), strlen($header) + 4)
            );
        }

        $headerStyle = $options['header_style'] ?? [
            'font' => ['bold' => true, 'color' => '#FFFFFF'],
            'fill' => '#1F4E78',
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center', 'wrap_text' => true],
        ];

        $sheet = new WorksheetData(
            name: (string) ($options['sheet_name'] ?? 'Import Template'),
            rows: $rows,
            columns: $headers,
            hasHeader: true,
            freezeHeader: false,
            autoFilter: true,
            headerStyle: $headerStyle,
            mergeCells: $mergeCells,
            columnWidths: $widths,
            rowHeights: $instructions !== '' ? [1 => 36, $headerRow => 24] : [$headerRow => 24],
            headerRowIndex: $headerRow,
            rowStyles: $rowStyles,
            comments: $comments,
            sourceColumnKeys: $headers,
            dataRowStart: $headerRow,
            dataRowCount: $sampleRows,
            freezeRows: $headerRow,
            freezeColumns: 0,
            freezeTopLeftCell: 'A' . ($headerRow + 1),
            autoFilterRange: 'A' . $headerRow . ':' . Coordinate::columnIndexToName(count($headers)) . max($headerRow + $sampleRows, $headerRow + 1),
            dataValidations: $validations,
        );

        return new WorkbookData([$sheet], [
            'title' => (string) ($options['title'] ?? 'Import Template'),
            '_mnb_xlsx_integrity_validation' => ['enabled' => (bool) ($options['integrity_validation'] ?? true)],
        ]);
    }
}
