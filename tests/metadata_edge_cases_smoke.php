<?php

declare(strict_types=1);

require __DIR__ . '/../vendor/autoload.php';

use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Support\Zip\ZipArchive;
use Mnb\PHPExcel\Support\XlsxIntegrityValidator;
use Mnb\PHPExcel\Writer\XlsxWriter;

function metadata_edge_check(bool $condition, string $message): void
{
    if (!$condition) {
        throw new RuntimeException($message);
    }
}

/** @return array<string,string> */
function metadata_edge_entries(string $path): array
{
    $zip = new ZipArchive();
    metadata_edge_check($zip->open($path) === true, 'Unable to open package: ' . $path);
    $entries = [];
    for ($i = 0; $i < $zip->numFiles; $i++) {
        $name = $zip->getNameIndex($i);
        if (!is_string($name) || $name === '' || str_ends_with($name, '/')) {
            continue;
        }
        $content = $zip->getFromName($name);
        metadata_edge_check(is_string($content), 'Unable to read package entry: ' . $name);
        $entries[$name] = $content;
    }
    $zip->close();
    return $entries;
}

/** @param array<string,string> $entries */
function metadata_edge_write_package(string $path, array $entries): void
{
    $zip = new ZipArchive();
    metadata_edge_check(
        $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE) === true,
        'Unable to create package: ' . $path,
    );
    foreach ($entries as $name => $content) {
        metadata_edge_check($zip->addFromString($name, $content), 'Unable to write package entry: ' . $name);
    }
    $zip->close();
}

$directory = sys_get_temp_dir() . '/mnb-xlsx-metadata-edge-' . bin2hex(random_bytes(5));
metadata_edge_check(mkdir($directory, 0775, true), 'Unable to create edge-case test directory.');
$base = $directory . '/base.xlsx';
$minimal = $directory . '/minimal.xlsx';
$registered = $directory . '/registered.xlsx';
$custom = $directory . '/custom.xlsx';
$customUpdated = $directory . '/custom-updated.xlsx';
$readOnly = $directory . '/read-only.xlsx';
$invalid = $directory . '/invalid.xlsx';

try {
    (new XlsxWriter())->write(
        new WorkbookData([new WorksheetData('Sheet1', [['Value'], [1]])]),
        $base,
    );

    // Metadata parts added to a minimal package must also be registered in
    // [Content_Types].xml and the package-root relationship file.
    $entries = metadata_edge_entries($base);
    unset($entries['docProps/core.xml'], $entries['docProps/app.xml']);
    $entries['[Content_Types].xml'] = preg_replace(
        '#<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Override\b[^>]*PartName=("|\')/docProps/(?:core|app)\.xml\1[^>]*/>#i',
        '',
        $entries['[Content_Types].xml'],
    ) ?? $entries['[Content_Types].xml'];
    $entries['_rels/.rels'] = preg_replace(
        '#<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Relationship\b[^>]*Type=("|\')[^"\']*/(?:core-properties|extended-properties)\1[^>]*/>#i',
        '',
        $entries['_rels/.rels'],
    ) ?? $entries['_rels/.rels'];
    metadata_edge_write_package($minimal, $entries);
    // An orphaned property part must be rejected by integrity validation.
    $brokenEntries = $entries;
    $brokenEntries['docProps/core.xml'] = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>';
    $broken = $directory . '/broken-registration.xlsx';
    metadata_edge_write_package($broken, $brokenEntries);
    $brokenValidation = (new XlsxIntegrityValidator())->validate($broken);
    metadata_edge_check(!(bool) ($brokenValidation['valid'] ?? true), 'Validator accepted an unregistered document property part.');

    Xlsx::updateMetaInfo($minimal, $registered, [
        'document' => ['title' => 'Registered'],
        'application' => ['company' => 'MNB'],
    ]);
    $registeredEntries = metadata_edge_entries($registered);
    metadata_edge_check(str_contains($registeredEntries['[Content_Types].xml'], '/docProps/core.xml'), 'Core properties content type was not registered.');
    metadata_edge_check(str_contains($registeredEntries['[Content_Types].xml'], '/docProps/app.xml'), 'App properties content type was not registered.');
    metadata_edge_check(str_contains($registeredEntries['_rels/.rels'], 'core-properties'), 'Core properties relationship was not registered.');
    metadata_edge_check(str_contains($registeredEntries['_rels/.rels'], 'extended-properties'), 'App properties relationship was not registered.');
    $validation = (new XlsxIntegrityValidator())->validate($registered);
    metadata_edge_check((bool) ($validation['valid'] ?? false), 'Registered metadata package failed integrity validation.');

    // Untouched custom properties must retain their original OOXML value type,
    // body and property ID while unrelated properties are updated.
    $entries = metadata_edge_entries($registered);
    $entries['docProps/custom.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
        . 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        . '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="BigUnsigned"><vt:ui8>18446744073709551615</vt:ui8></property>'
        . '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="Ansi"><vt:lpstr>abc</vt:lpstr></property>'
        . '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="Vector"><vt:vector size="2" baseType="lpwstr"><vt:lpwstr>a</vt:lpwstr><vt:lpwstr>b</vt:lpwstr></vt:vector></property>'
        . '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="EmptyValue"><vt:empty/></property>'
        . '</Properties>';
    // Ensure the injected custom part itself is validly registered.
    if (!str_contains($entries['[Content_Types].xml'], '/docProps/custom.xml')) {
        $entries['[Content_Types].xml'] = preg_replace(
            '/<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Types>\s*$/i',
            '<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>',
            $entries['[Content_Types].xml'],
            1,
        ) ?? $entries['[Content_Types].xml'];
    }
    if (!str_contains($entries['_rels/.rels'], 'custom-properties')) {
        $entries['_rels/.rels'] = preg_replace(
            '/<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?Relationships>\s*$/i',
            '<Relationship Id="rIdCustomProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>',
            $entries['_rels/.rels'],
            1,
        ) ?? $entries['_rels/.rels'];
    }
    metadata_edge_write_package($custom, $entries);

    Xlsx::updateMetaInfo($custom, $customUpdated, [
        'custom_properties' => [
            'Added' => 'new value',
            'LargeSigned' => 2147483648,
            'ExplicitFalse' => false,
        ],
    ]);
    $updatedEntries = metadata_edge_entries($customUpdated);
    $customXml = $updatedEntries['docProps/custom.xml'];
    metadata_edge_check(str_contains($customXml, '<vt:ui8>18446744073709551615</vt:ui8>'), 'Untouched ui8 property was rewritten or truncated.');
    metadata_edge_check(str_contains($customXml, '<vt:lpstr>abc</vt:lpstr>'), 'Untouched lpstr property type was rewritten.');
    metadata_edge_check(str_contains($customXml, '<vt:vector size="2" baseType="lpwstr"><vt:lpwstr>a</vt:lpwstr><vt:lpwstr>b</vt:lpwstr></vt:vector>'), 'Untouched vector property was rewritten.');
    metadata_edge_check(str_contains($customXml, '<vt:i8>2147483648</vt:i8>'), 'Large signed integer was not written as vt:i8.');
    metadata_edge_check(str_contains($customXml, '<vt:bool>false</vt:bool>'), 'Boolean false was not preserved exactly.');

    $customInfo = Xlsx::metaInfo($customUpdated, ['profile' => 'quick']);
    $byName = [];
    foreach ($customInfo['custom_properties']['items'] ?? [] as $item) {
        $byName[$item['name']] = $item;
    }
    metadata_edge_check(($byName['BigUnsigned']['value'] ?? null) === '18446744073709551615', 'Out-of-range ui8 metadata was not returned losslessly as a string.');
    metadata_edge_check(($byName['Vector']['type'] ?? null) === 'opaque', 'Unsupported custom-property vector should be reported as opaque.');
    metadata_edge_check(array_key_exists('value', $byName['Vector']) && $byName['Vector']['value'] === null, 'Opaque custom-property values must not be flattened into misleading text.');
    metadata_edge_check(($byName['EmptyValue']['type'] ?? null) === 'null', 'Self-closing vt:empty custom property was not inventoried.');
    metadata_edge_check(array_key_exists('value', $byName['EmptyValue']) && $byName['EmptyValue']['value'] === null, 'Self-closing vt:empty custom property should have a null value.');

    // fileSharing belongs before sheets in workbook element order, including
    // packages that do not already contain workbookPr.
    $entries = metadata_edge_entries($registered);
    $entries['xl/workbook.xml'] = preg_replace('/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?workbookPr\b[^>]*\/?\s*>/i', '', $entries['xl/workbook.xml']) ?? $entries['xl/workbook.xml'];
    metadata_edge_write_package($readOnly, $entries);
    Xlsx::updateMetaInfo($readOnly, $readOnly, ['workbook' => ['read_only_recommended' => true]]);
    $workbookXml = metadata_edge_entries($readOnly)['xl/workbook.xml'];
    metadata_edge_check(strpos($workbookXml, 'fileSharing') < strpos($workbookXml, '<sheets'), 'fileSharing was inserted after sheets, violating workbook element order.');

    // Invalid scalar values must fail before a destination is published.
    $thrown = false;
    try {
        Xlsx::updateMetaInfo($registered, $invalid, ['custom_properties' => ['Infinite' => INF]]);
    } catch (Throwable) {
        $thrown = true;
    }
    metadata_edge_check($thrown, 'Non-finite custom-property values should be rejected.');
    metadata_edge_check(!is_file($invalid), 'A failed metadata update left a destination file behind.');

    foreach ([
        ['revision' => ['total_editing_time_seconds' => 'not-a-number']],
        ['calculation' => ['calc_id' => 'not-a-number']],
        ['workbook' => ['date1904' => ['invalid']]],
        ['application' => ['company' => ['invalid']]],
    ] as $index => $invalidChanges) {
        $invalidTyped = $directory . '/invalid-type-' . $index . '.xlsx';
        $thrown = false;
        try {
            Xlsx::updateMetaInfo($registered, $invalidTyped, $invalidChanges);
        } catch (Throwable) {
            $thrown = true;
        }
        metadata_edge_check($thrown, 'Invalid metadata value type was accepted for case ' . $index . '.');
        metadata_edge_check(!is_file($invalidTyped), 'Invalid metadata value type left an output file for case ' . $index . '.');
    }

    echo "metadata_edge_cases_smoke passed\n";
} finally {
    $items = glob($directory . '/*');
    if (is_array($items)) {
        foreach ($items as $item) {
            @unlink($item);
        }
    }
    @rmdir($directory);
}
