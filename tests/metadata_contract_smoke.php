<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Writer\XlsxWriter;

function metadata_check(bool $condition, string $message): void
{
    if (!$condition) {
        throw new RuntimeException($message);
    }
}

$directory = sys_get_temp_dir() . '/mnb-xlsx-metadata-' . bin2hex(random_bytes(5));
if (!mkdir($directory, 0775, true) && !is_dir($directory)) {
    throw new RuntimeException('Unable to create metadata smoke-test directory.');
}
$path = $directory . '/metadata.xlsx';

try {
    $data = new WorksheetData(
        'Data',
        [
            ['Name', 'Amount', 'Result'],
            ['Alice', 10, CellValue::formula('B2*2', 20)],
            ['Bob', 15, 30],
        ],
        hasHeader: true,
        autoFilter: true,
        hyperlinks: [
            ['cell' => 'A2', 'url' => 'https://example.com', 'display' => 'Alice'],
        ],
        comments: [
            ['cell' => 'B2', 'author' => 'Reviewer', 'text' => 'Check this amount.'],
        ],
        conditionalFormats: [
            ['type' => 'cell_is', 'range' => 'B2:B3', 'operator' => 'greaterThan', 'formulas' => ['12']],
        ],
        dataValidations: [
            ['type' => 'list', 'range' => 'A2:A3', 'values' => ['Alice', 'Bob']],
        ],
    );
    $hidden = new WorksheetData('HiddenSheet', [['Secret'], ['x']]);
    $workbook = new WorkbookData([$data, $hidden], [
        'title' => 'Metadata Test',
        'subject' => 'Unified metadata',
        'creator' => 'MNB',
        'last_modified_by' => 'QA',
        'company' => 'MNB Technologies',
        '_mnb_workbook_protection' => ['password' => 'secret', 'lockStructure' => true],
        '_mnb_sheet_protection' => ['Data' => ['password' => 'secret', 'sheet' => true]],
    ]);
    (new XlsxWriter())->write($workbook, $path);

    Xlsx::updateMetaInfo($path, $path, [
        'custom_properties' => [
            'Project ID' => 'PRJ-1001',
            'Budget' => 1250.5,
        ],
        'workbook' => [
            'sheet_visibility' => ['HiddenSheet' => 'hidden'],
        ],
    ]);

    $quick = Xlsx::metaInfo($path, ['profile' => 'quick']);
    metadata_check($quick['schema_version'] === '1.0', 'Unexpected metadata schema version.');
    metadata_check($quick['document']['title'] === 'Metadata Test', 'Quick profile did not read title.');
    metadata_check($quick['document']['company'] === 'MNB Technologies', 'Quick profile did not read company.');
    metadata_check($quick['workbook']['sheet_count'] === 2, 'Quick profile did not read worksheets.');
    metadata_check($quick['hidden_content']['hidden_sheet_count'] === 1, 'Quick profile did not report hidden sheet.');
    metadata_check($quick['validation']['state'] === 'not_scanned', 'Quick profile scanned validation metadata.');
    metadata_check($quick['calculation']['formula_count'] === null, 'Quick profile incorrectly reported zero formulas without scanning.');
    metadata_check($quick['links']['hyperlink_count'] === null, 'Quick profile incorrectly reported zero hyperlinks without scanning.');
    metadata_check($quick['workbook']['selected_sheets'] === null, 'Quick profile should not fabricate selected worksheet metadata.');

    $standard = Xlsx::read($path)->metaInfo(['profile' => 'standard']);
    metadata_check($standard['validation']['data_validation_count'] === 1, 'Data validation count is incorrect.');
    metadata_check($standard['validation']['conditional_formatting_count'] === 1, 'Conditional-format count is incorrect.');
    metadata_check($standard['calculation']['formula_count'] === 1, 'Formula count is incorrect.');
    metadata_check($standard['links']['hyperlink_count'] === 1, 'Hyperlink count is incorrect.');
    metadata_check($standard['security']['workbook_protected'] === true, 'Workbook protection was not detected.');
    metadata_check($standard['security']['worksheet_protected_count'] === 1, 'Sheet protection was not detected.');

    $full = Xlsx::metaInfo($path, ['profile' => 'full']);
    metadata_check($full['comments_notes']['comment_count'] === 1, 'Comment count is incorrect.');
    metadata_check($full['calculation']['items'][0]['cell'] === 'C2', 'Formula inventory has the wrong cell.');
    metadata_check($full['custom_properties']['count'] === 2, 'Custom property count is incorrect.');
    metadata_check($full['capabilities']['document']['read'] === true, 'Document metadata read capability is missing.');
    metadata_check($full['capabilities']['document']['write'] === true, 'Document metadata write capability is missing.');

    $limited = Xlsx::metaInfo($path, ['profile' => 'full', 'max_items' => 1]);
    metadata_check($limited['custom_properties']['count'] === 2, 'Limited profile lost the complete custom-property count.');
    metadata_check(count($limited['custom_properties']['items']) === 1, 'max_items was not enforced.');
    metadata_check($limited['custom_properties']['truncated'] === true, 'Truncation was not reported.');

    $forensic = Xlsx::metaInfo($path, ['profile' => 'forensic']);
    metadata_check(is_string($forensic['file']['sha256']) && strlen($forensic['file']['sha256']) === 64, 'Forensic profile did not calculate SHA-256.');
    metadata_check($forensic['format_details']['items'] !== [], 'Forensic profile did not return package parts.');

    echo "metadata_contract_smoke passed\n";
} finally {
    @unlink($path);
    @rmdir($directory);
}
