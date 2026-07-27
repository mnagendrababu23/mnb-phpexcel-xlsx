<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Format\Xlsx;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');
$sheet = $argv[2] ?? 1;

print_r([
    // Rows containing at least one worksheet cell.
    'filled_rows' => Xlsx::rowCount($path, $sheet),

    // Physical <row> XML elements, including styled empty rows.
    'physical_rows' => Xlsx::rowCount($path, $sheet, ['mode' => 'physical']),

    // Highest row index referenced in worksheet XML.
    'last_used_row' => Xlsx::rowCount($path, $sheet, ['mode' => 'last_row']),

    // Fastest mode: last row declared by the worksheet dimension.
    'declared_last_row' => Xlsx::rowCount($path, $sheet, ['mode' => 'declared']),

    // Filled-row count for every worksheet.
    'all_sheets' => Xlsx::rowCounts($path),
]);
