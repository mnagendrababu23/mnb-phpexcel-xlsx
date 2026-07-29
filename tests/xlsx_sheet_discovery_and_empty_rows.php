<?php

declare(strict_types=1);

$coreRoot = dirname(__DIR__, 2) . '/mnb-phpexcel-core';
$xlsxRoot = dirname(__DIR__);

spl_autoload_register(static function (string $class) use ($coreRoot, $xlsxRoot): void {
    $prefix = 'Mnb\\PHPExcel\\';
    if (!str_starts_with($class, $prefix)) {
        return;
    }

    $relative = str_replace('\\', '/', substr($class, strlen($prefix))) . '.php';
    foreach ([$xlsxRoot . '/src/' . $relative, $coreRoot . '/src/' . $relative] as $path) {
        if (is_file($path)) {
            require $path;
            return;
        }
    }
});

use Mnb\PHPExcel\Format\Xlsx;
use Mnb\PHPExcel\Support\EmptyWorksheetException;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\Zip\ZipArchive;

$path = tempnam(sys_get_temp_dir(), 'mnb-active-sheet-');
if ($path === false) {
    throw new RuntimeException('Unable to create temporary XLSX path.');
}
@unlink($path);
$path .= '.xlsx';

$parts = [
    '[Content_Types].xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        . '<Default Extension="xml" ContentType="application/xml"/>'
        . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        . '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        . '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        . '</Types>',
    '_rels/.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        . '</Relationships>',
    'xl/workbook.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        . '<bookViews><workbookView activeTab="1"/></bookViews>'
        . '<sheets>'
        . '<sheet name="Cover" sheetId="1" r:id="rId1"/>'
        . '<sheet name="Data" sheetId="2" r:id="rId2"/>'
        . '<sheet name="Empty" sheetId="3" r:id="rId3"/>'
        . '</sheets></workbook>',
    'xl/_rels/workbook.xml.rels' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        . '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        . '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>'
        . '</Relationships>',
    'xl/worksheets/sheet1.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        . '<row r="1"><c r="A1" t="inlineStr"><is><t>Welcome</t></is></c></row>'
        . '</sheetData></worksheet>',
    'xl/worksheets/sheet2.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        . '<row r="1"><c r="A1" t="inlineStr"><is><t>id</t></is></c><c r="B1" t="inlineStr"><is><t>name</t></is></c></row>'
        . '<row r="2"><c r="A2"><v>1</v></c><c r="B2" t="inlineStr"><is><t>Ada</t></is></c></row>'
        . '<row r="3"><c r="A3"><v>2</v></c><c r="B3" t="inlineStr"><is><t>Linus</t></is></c></row>'
        . '</sheetData></worksheet>',
    'xl/worksheets/sheet3.xml' => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        . '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>',
];

$zip = new ZipArchive();
if ($zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
    throw new RuntimeException('Unable to create XLSX fixture.');
}
foreach ($parts as $name => $content) {
    if (!$zip->addFromString($name, $content)) {
        throw new RuntimeException('Unable to add XLSX fixture entry: ' . $name);
    }
}
$zip->close();

try {
    $session = Xlsx::read($path);

    assert($session->sheetNames() === ['Cover', 'Data', 'Empty']);
    assert($session->hasSheet('Data'));
    assert($session->hasSheet(3));
    assert(!$session->hasSheet('Missing'));
    assert($session->sheetIfExists('Missing') === null);

    assert($session->activeSheetInfo() === ['index' => 2, 'name' => 'Data']);
    assert(($session->inspect()['active_sheet'] ?? null) === ['index' => 2, 'name' => 'Data']);
    assert($session->activeSheetIndex() === 2);
    assert($session->activeSheetName() === 'Data');

    $active = $session->activeSheet()->withHeaderRow(1);
    assert($active->hasRows());
    assert(!$active->isEmpty());
    assert($active->countRows() === 2);
    assert($active->first() === ['id' => 1, 'name' => 'Ada']);

    $empty = $session->sheet('Empty');
    assert(!$empty->hasRows());
    assert($empty->isEmpty());
    assert($empty->countRows() === 0);

    try {
        $empty->requireRows();
        throw new RuntimeException('Expected EmptyWorksheetException was not thrown.');
    } catch (EmptyWorksheetException $e) {
        assert($e->getErrorCode() === ErrorCode::SHEET_EMPTY);
        assert(($e->context()['selected_sheet'] ?? null) === 'Empty');
    }

    echo "XLSX sheet discovery, active-sheet, and empty-row checks passed\n";
} finally {
    @unlink($path);
}
