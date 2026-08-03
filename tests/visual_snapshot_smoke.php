<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Security\XlsxEncryption;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Writer\XlsxWriter;

$source = sys_get_temp_dir() . '/mnb-visual-source-' . bin2hex(random_bytes(4)) . '.xlsx';
$out = sys_get_temp_dir() . '/mnb-visual-snapshot-' . bin2hex(random_bytes(4)) . '.xlsx';
$styleOut = sys_get_temp_dir() . '/mnb-visual-style-regression-' . bin2hex(random_bytes(4)) . '.xlsx';
$staleDimensionOut = sys_get_temp_dir() . '/mnb-visual-stale-dimension-' . bin2hex(random_bytes(4)) . '.xlsx';
$encryptedOut = sys_get_temp_dir() . '/mnb-visual-encrypted-' . bin2hex(random_bytes(4)) . '.xlsx';
try {
    $sourceWorkbook = new WorkbookData([
        new WorksheetData(
            name: 'Sheet1',
            rows: [
                ['Order ID', 'Order Date', 'Customer', 'Product', 'Quantity', 'Unit Price', 'Total', 'Status'],
                ['ORD-1001', CellValue::date('2026-07-01', ['format' => 'yyyy-mm-dd']), 'Aarav Retail', 'Wireless Keyboard', 4, 2499.0, 9996.0, 'Completed'],
                ['ORD-1002', CellValue::date('2026-07-03', ['format' => 'yyyy-mm-dd']), 'Bright Stores', 'Ergonomic Mouse', 8, 1299.0, 10392.0, 'Completed'],
                ['ORD-1003', CellValue::date('2026-07-05', ['format' => 'yyyy-mm-dd']), 'City Mart', 'USB-C Hub', 5, 3199.0, 15995.0, 'Pending'],
                ['ORD-1004', CellValue::date('2026-07-08', ['format' => 'yyyy-mm-dd']), 'Delta Solutions', 'Laptop Stand', 10, 1899.0, 18990.0, 'Completed'],
            ],
            hasHeader: true,
            headerStyle: [
                'font' => ['bold' => true, 'color' => '#FFFFFF'],
                'fill' => ['color' => '#1F4E78'],
                'alignment' => ['horizontal' => 'center'],
            ],
            headerRowIndex: 1,
            columnWidths: ['A' => 15, 'B' => 14, 'C' => 22, 'D' => 28, 'E' => 12, 'F' => 15, 'G' => 16, 'H' => 14],
            rowHeights: [1 => 30],
            columnStyles: [
                2 => ['format' => 'yyyy-mm-dd'],
                5 => ['format' => '#,##0'],
                6 => ['format' => '₹#,##0.00'],
                7 => ['format' => '₹#,##0.00'],
            ],
            rangeStyles: [
                'A2:H5' => ['borders' => ['all' => ['style' => 'thin']]],
            ],
            freezeRows: 1,
            autoFilter: true,
            autoFilterRange: 'A1:H5',
            conditionalFormats: [[
                'range' => 'G2:G5',
                'type' => 'color_scale',
                'colors' => ['#F8696B', '#FFEB84', '#63BE7B'],
            ]],
        ),
    ]);
    (new XlsxWriter())->write($sourceWorkbook, $source);

    $snapshot = Xlsx::visualSnapshot($source);
    assert($snapshot['schema'] === 'mnb-phpexcel.visual-snapshot');
    assert($snapshot['sheets'][0]['dimension'] === 'A1:H5');
    assert(isset($snapshot['sheets'][0]['cells']['H5']));
    assert(($snapshot['sheets'][0]['cells']['A1']['style_id'] ?? null) !== null);
    $headerStyle = $snapshot['styles'][$snapshot['sheets'][0]['cells']['A1']['style_id']];
    assert(($headerStyle['font']['bold'] ?? false) === true);
    assert(($headerStyle['font']['color']['rgb'] ?? null) === 'FFFFFFFF');
    assert(($headerStyle['fill']['foreground']['rgb'] ?? null) === 'FF1F4E78');
    assert(($snapshot['sheets'][0]['layout']['column_widths']['H'] ?? null) === 14.0);
    assert(($snapshot['sheets'][0]['layout']['row_heights'][1] ?? null) === 30.0);
    assert(($snapshot['sheets'][0]['layout']['freeze_panes']['rows'] ?? null) === 1);
    assert(($snapshot['sheets'][0]['layout']['auto_filter_range'] ?? null) === 'A1:H5');
    assert(($snapshot['sheets'][0]['conditional_formats'][0]['type'] ?? null) === 'color_scale');
    assert(($snapshot['sheets'][0]['cells']['B2']['type'] ?? null) === 'date');

    Xlsx::createFromVisualSnapshot($snapshot, $out);
    $roundTrip = Xlsx::visualSnapshot($out);
    assert($roundTrip['sheets'][0]['dimension'] === 'A1:H5');
    assert(($roundTrip['sheets'][0]['cells']['B2']['type'] ?? null) === 'date');
    assert(($roundTrip['sheets'][0]['layout']['freeze_panes']['rows'] ?? null) === 1);
    assert(($roundTrip['sheets'][0]['layout']['auto_filter_range'] ?? null) === 'A1:H5');
    assert(($roundTrip['sheets'][0]['conditional_formats'][0]['type'] ?? null) === 'color_scale');

    $zip = new ZipArchive();
    assert($zip->open($out) === true);
    $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    $stylesXml = $zip->getFromName('xl/styles.xml');
    $zip->close();
    assert(is_string($sheetXml) && str_contains($sheetXml, '<autoFilter ref="A1:H5"'));
    assert(is_string($sheetXml) && str_contains($sheetXml, '<conditionalFormatting sqref="G2:G5"'));
    assert(is_string($sheetXml) && preg_match('/<c r="B2"[^>]*><v>\d+(?:\.\d+)?<\/v><\/c>/', $sheetXml) === 1);
    assert(is_string($stylesXml) && str_contains($stylesXml, 'FF1F4E78'));

    $styledWorkbook = new WorkbookData([
        new WorksheetData(
            name: 'Visible',
            rows: [
                ['Date', 'Amount'],
                [CellValue::date('2026-07-01'), 1250.5],
            ],
            cellStyles: [
                'A1' => [
                    'font' => ['bold' => true, 'color' => '#FFFFFF'],
                    'fill' => ['color' => '#1F4E78'],
                ],
            ],
            rangeStyles: [
                'A1:B2' => [
                    'borders' => [
                        'all' => ['style' => 'thin', 'color' => '#FF0000'],
                    ],
                ],
            ],
            columnWidths: ['A' => 18, 'B' => 14],
            rowHeights: [1 => 24],
            freezeRows: 1,
            autoFilterRange: 'A1:B2'
        ),
        new WorksheetData(name: 'Hidden', rows: [['secret']]),
    ], [
        '_mnb_active_sheet' => 0,
        '_mnb_sheet_states' => ['Hidden' => 'hidden'],
    ]);
    (new XlsxWriter())->write($styledWorkbook, $styleOut);
    $styledSnapshot = Xlsx::visualSnapshot($styleOut);
    assert(($styledSnapshot['sheets'][1]['state'] ?? null) === 'hidden');
    assert(($styledSnapshot['workbook']['active_sheet_name'] ?? null) === 'Visible');
    $a1Style = $styledSnapshot['styles'][$styledSnapshot['sheets'][0]['cells']['A1']['style_id']];
    foreach (['left', 'right', 'top', 'bottom'] as $side) {
        assert(($a1Style['border'][$side]['style'] ?? null) === 'thin');
    }
    assert(($a1Style['border']['left']['color']['rgb'] ?? null) === 'FFFF0000');
    assert(($styledSnapshot['sheets'][0]['cells']['A2']['type'] ?? null) === 'date');
    $dateStyle = $styledSnapshot['styles'][$styledSnapshot['sheets'][0]['cells']['A2']['style_id']] ?? [];
    assert(($dateStyle['format'] ?? null) === 'm/d/yy');

    $zip = new ZipArchive();
    assert($zip->open($styleOut) === true);
    $styleSheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    $styleStylesXml = $zip->getFromName('xl/styles.xml');
    $zip->close();
    assert(is_string($styleStylesXml) && str_contains($styleStylesXml, '<left style="thin"><color rgb="FFFF0000"/></left>'));
    assert(is_string($styleSheetXml) && preg_match('/<c r="A2"[^>]*><v>\d+(?:\.\d+)?<\/v><\/c>/', $styleSheetXml) === 1);

    assert(copy($styleOut, $staleDimensionOut));
    $zip = new ZipArchive();
    assert($zip->open($staleDimensionOut) === true);
    $staleSheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    assert(is_string($staleSheetXml));
    $staleSheetXml = preg_replace('/<dimension\b[^>]*\/>/', '<dimension ref="A1:XFD1048576"/>', $staleSheetXml, 1);
    assert(is_string($staleSheetXml));
    assert($zip->addFromString('xl/worksheets/sheet1.xml', $staleSheetXml));
    $zip->close();
    $staleSnapshot = Xlsx::visualSnapshot($staleDimensionOut, ['max_cells' => 100]);
    assert(($staleSnapshot['sheets'][0]['dimension'] ?? null) === 'A1:B2');
    assert(($staleSnapshot['sheets'][0]['layout']['declared_dimension'] ?? null) === 'A1:XFD1048576');

    $password = 'Visual-Snapshot-2026!';
    Xlsx::createFromVisualSnapshot($snapshot, $encryptedOut, [
        'password' => $password,
        'encryption_mode' => 'standard',
    ]);
    assert((new XlsxEncryption())->isEncryptedFile($encryptedOut));
    $encryptedSnapshot = Xlsx::visualSnapshot($encryptedOut, ['password' => $password]);
    assert(($encryptedSnapshot['sheets'][0]['dimension'] ?? null) === 'A1:H5');
    $passwordRequired = false;
    try {
        Xlsx::visualSnapshot($encryptedOut);
    } catch (Throwable) {
        $passwordRequired = true;
    }
    assert($passwordRequired);

    echo "visual_snapshot_smoke passed\n";
} finally {
    @unlink($source);
    @unlink($out);
    @unlink($styleOut);
    @unlink($staleDimensionOut);
    @unlink($encryptedOut);
}
