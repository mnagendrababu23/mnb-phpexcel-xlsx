<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Format\Xlsx;

$path = $argv[1] ?? (__DIR__ . '/workbook.xlsx');

// Fast mode reads sheet names, visibility, dimensions, and worksheet XML sizes.
$fast = Xlsx::sheetsInfo($path);

// Accurate mode streams worksheet XML tags to count rows and cells. It still
// does not hydrate workbook rows or decode cell values.
$accurate = Xlsx::sheetsInfo($path, [
    'accurate_row_count' => true,
]);

print_r([
    'fast' => $fast,
    'accurate' => $accurate,
]);
