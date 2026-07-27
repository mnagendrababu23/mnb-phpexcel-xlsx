<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Format\Xlsx;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');

// Reads filesystem and OOXML package metadata. Worksheet rows and cell values
// are not converted into PHP arrays.
$info = Xlsx::fileInfo($path);

print_r([
    'name' => $info['name'],
    'size_bytes' => $info['size_bytes'],
    'encrypted' => $info['encrypted'],
    'sheet_count' => $info['sheet_count'],
    'sheet_names' => $info['sheet_names'],
    'zip_entries' => $info['zip_entries'],
    'has_macros' => $info['has_macros'],
    'properties' => $info['properties'],
]);
