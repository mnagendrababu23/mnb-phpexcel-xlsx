<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Format\Xlsx;

$source = $argv[1] ?? '';
if ($source === '' || !is_file($source)) {
    fwrite(STDERR, "Usage: php examples/unified_metadata.php workbook.xlsx [updated.xlsx]\n");
    exit(1);
}

$report = Xlsx::metaInfo($source, [
    'profile' => 'standard',
]);

echo json_encode($report, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE), PHP_EOL;

$destination = $argv[2] ?? '';
if ($destination !== '') {
    Xlsx::updateMetaInfo($source, $destination, [
        'document' => [
            'title' => 'Updated with MNB PHPExcel',
            'creator' => 'MNB PHPExcel',
        ],
        'custom_properties' => [
            'Metadata Schema' => '1.0',
        ],
    ]);
    echo 'Updated workbook: ' . $destination . PHP_EOL;
}
