<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Support;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Writer\XlsxWriter;

/** Internal native fixture builder used by the package compatibility verifier. */
final class XlsxCompatibilityFixtureFactory
{
    public function basic(string $path): void
    {
        $sheet = new WorksheetData(
            name: 'Sheet1',
            rows: [
                ['Name', 'Score', 'Passed'],
                ['Asha', 95, true],
                ['Ravi', 88, true],
            ],
            hasHeader: true,
            freezeHeader: true,
            autoFilter: true,
            headerStyle: $this->headerStyle(),
            columnWidths: [1 => 18, 2 => 12, 3 => 12],
            headerRowIndex: 1,
            dataRowStart: 1,
            dataRowCount: 2,
        );
        $this->write($path, new WorkbookData([$sheet], ['title' => 'Basic compatibility workbook']));
    }

    public function formulasStylesMergedCells(string $path): void
    {
        $sheet = new WorksheetData(
            name: 'Report',
            rows: [
                ['Compatibility Verification', null],
                ['Metric', 'Value'],
                ['Revenue', 12500],
                ['Cost', 7200],
                ['Profit', CellValue::formula('B3-B4', 5300)],
            ],
            hasHeader: true,
            headerStyle: $this->headerStyle(),
            mergeCells: ['A1:B1'],
            columnWidths: [1 => 24, 2 => 16],
            rowStyles: [1 => ['font' => ['bold' => true, 'size' => 16], 'alignment' => ['horizontal' => 'center']]],
            cellStyles: ['B5' => ['font' => ['bold' => true], 'number_format' => '$#,##0.00']],
            headerRowIndex: 2,
            dataRowStart: 2,
            dataRowCount: 3,
            freezeRows: 2,
            autoFilter: true,
            autoFilterRange: 'A2:B5',
        );
        $this->write($path, new WorkbookData([$sheet], ['title' => 'Formula and styles compatibility workbook']));
    }

    public function commentsHyperlinks(string $path): void
    {
        $sheet = new WorksheetData(
            name: 'Resources',
            rows: [
                ['Resource', 'URL', 'Status'],
                ['Repository', 'Open repository', 'Review'],
                ['Documentation', 'Docs', 'Ready'],
            ],
            hasHeader: true,
            headerStyle: $this->headerStyle(),
            columnWidths: [1 => 20, 2 => 24, 3 => 14],
            headerRowIndex: 1,
            dataRowStart: 1,
            dataRowCount: 2,
            hyperlinks: [[
                'cell' => 'B2',
                'url' => 'https://github.com/mnagendrababu23/mnb-phpexcel',
                'display' => 'Open repository',
                'tooltip' => 'Project repository',
            ]],
            comments: [[
                'cell' => 'C2',
                'author' => 'MNB PHPExcel',
                'text' => 'Verify this generated note opens without a repair warning.',
            ]],
        );
        $this->write($path, new WorkbookData([$sheet], ['title' => 'Comments and hyperlinks compatibility workbook']));
    }

    public function preservedAdvancedObjects(string $path, string $template): void
    {
        $templateSheet = new WorksheetData(
            name: 'Sheet1',
            rows: [
                ['Task', 'Link'],
                ['Template', 'Template link'],
            ],
            hasHeader: true,
            headerStyle: $this->headerStyle(),
            headerRowIndex: 1,
            dataRowStart: 1,
            dataRowCount: 1,
            hyperlinks: [[
                'cell' => 'B2',
                'url' => 'https://example.com/template',
                'display' => 'Template link',
            ]],
            comments: [[
                'cell' => 'A2',
                'author' => 'Template',
                'text' => 'Template note to preserve.',
            ]],
        );
        $this->write($template, new WorkbookData([$templateSheet], ['title' => 'Advanced object template']));

        $targetSheet = new WorksheetData(
            name: 'Sheet1',
            rows: [
                ['Task', 'Link'],
                ['Generated', 'Copied package parts'],
            ],
            hasHeader: true,
            headerStyle: $this->headerStyle(),
            headerRowIndex: 1,
            dataRowStart: 1,
            dataRowCount: 1,
        );
        $metadata = [
            'title' => 'Advanced object preservation compatibility workbook',
            '_mnb_preserve_xlsx_package' => [
                'path' => realpath($template) ?: $template,
                'mode' => 'advanced_objects',
                'sheet_match' => 'index',
                'preserve_sheet_relationships' => true,
                'preserve_sheet_elements' => true,
                'copy_package_parts' => true,
            ],
        ];
        $this->write($path, new WorkbookData([$targetSheet], $metadata));
    }

    /** @return array<string,mixed> */
    private function headerStyle(): array
    {
        return [
            'font' => ['bold' => true, 'color' => '#FFFFFF'],
            'fill' => '#1F4E78',
            'alignment' => ['horizontal' => 'center', 'vertical' => 'center'],
        ];
    }

    private function write(string $path, WorkbookData $workbook): void
    {
        (new XlsxWriter())->write($workbook, $path);
    }
}
