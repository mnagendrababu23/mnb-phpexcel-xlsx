<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorkbookFactory;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Writer\XlsxWriter;

$directory = sys_get_temp_dir() . '/mnb-xlsx-quick-info-' . bin2hex(random_bytes(5));
if (!mkdir($directory, 0775, true) && !is_dir($directory)) {
    throw new RuntimeException('Unable to create test directory.');
}
$path = $directory . '/quick-info.xlsx';

try {
    $orders = WorkbookFactory::worksheet([
        ['Order ID' => 1001, 'Customer' => 'Alice', 'Total' => 125.50],
        ['Order ID' => 1002, 'Customer' => 'Bob', 'Total' => 80.00],
    ], 'Orders', true);
    $notes = WorkbookFactory::worksheet([
        ['Message' => 'Generated for quick-info testing'],
    ], 'Notes', true);

    (new XlsxWriter())->write(new WorkbookData([$orders, $notes]), $path);

    $file = Xlsx::fileInfo($path);
    assert($file['status'] === 'ok');
    assert($file['sheet_count'] === 2);
    assert($file['sheet_names'] === ['Orders', 'Notes']);
    assert($file['size_bytes'] > 0);

    $fast = Xlsx::sheetsInfo($path);
    assert(count($fast) === 2);
    assert($fast[0]['name'] === 'Orders');
    assert($fast[0]['declared_last_row'] === 3);
    assert($fast[0]['accurate_row_count'] === false);

    $ordersInfo = Xlsx::sheetInfo($path, 'Orders', ['accurate_row_count' => true]);
    assert($ordersInfo['filled_rows'] === 3);
    assert($ordersInfo['physical_rows'] === 3);
    assert($ordersInfo['columns'] === 3);
    assert($ordersInfo['cells'] === 9);

    assert(Xlsx::rowCount($path, 'Orders') === 3);
    assert(Xlsx::rowCount($path, 'Orders', ['mode' => 'physical']) === 3);
    assert(Xlsx::rowCount($path, 'Orders', ['mode' => 'last_row']) === 3);
    assert(Xlsx::rowCount($path, 'Orders', ['mode' => 'declared']) === 3);
    assert(Xlsx::rowCounts($path) === ['Orders' => 3, 'Notes' => 2]);

    echo "quick_info_smoke passed\n";
} finally {
    @unlink($path);
    @rmdir($directory);
}
