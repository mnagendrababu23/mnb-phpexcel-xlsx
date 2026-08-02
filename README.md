# MNB PHPExcel XLSX

Independent native XLSX reader/writer with streaming, formulas, styles, charts, pivots, comments, hyperlinks, protection, and Office password encryption.

```bash
composer require mnb/mnb-phpexcel-xlsx:^2.0
```

For an XLSX-only installation, use `Mnb\PHPExcel\Format\Xlsx`—the `MnbExcel` facade belongs to the application package.

```php
use Mnb\PHPExcel\Format\Xlsx;

$rows = Xlsx::read('report.xlsx')
    ->sheet('Data')
    ->withHeaderRow()
    ->toArray();

Xlsx::write($rows, 'report-copy.xlsx', ['with_header' => true]);
```

Import templates are also package-local and do not require the application package:

```php
Xlsx::writeImportTemplate([
    'name' => ['header' => 'Name', 'required' => true],
    'status' => ['header' => 'Status', 'list' => ['Active', 'Inactive']],
], 'import-template.xlsx');
```

This package does not require, suggest, or call PhpSpreadsheet. It requires `ext-iconv`, `ext-openssl`, `ext-libxml`, and `ext-zlib`. `ext-zip` and `ext-xmlreader` remain optional performance extensions; package-local fallbacks are included.

## Worksheet selection and errors

Worksheet numbers are **1-based**. The first worksheet is selected automatically when `sheet()` is omitted.

```php
$session = Xlsx::read('report.xlsx');

$session->sheet(1);        // first worksheet
$session->sheet('Data');   // worksheet by name
$session->sheetNames();    // inspect available names
```

Do not call `sheet()` without an argument. The library throws a `SheetSelectionException` with an actionable message, workbook path, caller location, and stable error code for missing, invalid, or unknown selections.

```php
use Mnb\PHPExcel\Support\MnbExcelException;

try {
    $rows = Xlsx::read('report.xlsx')
        ->sheet('Missing')
        ->withHeaderRow()
        ->toArray();
} catch (MnbExcelException $e) {
    print_r($e->toErrorArray(debug: true));
}
```


## Sheet discovery, active sheet, and empty worksheets

```php
$session = Xlsx::read('report.xlsx');

if (!$session->hasSheet('Orders')) {
    $session = $session->activeSheet();
} else {
    $session = $session->sheet('Orders');
}

echo $session->activeSheetName();
print_r($session->activeSheetInfo());

$data = $session->withHeaderRow(1);
if ($data->isEmpty()) {
    // The worksheet has zero data rows after the header/options are applied.
    return;
}

$data->assertHasRows();
foreach ($data->rows() as $row) {
    // Process data.
}
```

`inspect()` now includes an `active_sheet` entry with a 1-based index and worksheet name. XLSX active-sheet detection reads the workbook's `activeTab`; workbooks without that metadata safely fall back to the first worksheet.

## Unified XLSX metadata

```php
$quick = Xlsx::metaInfo('report.xlsx', ['profile' => 'quick']);
$full = Xlsx::read('report.xlsx')->metaInfo(['profile' => 'full']);
```

Metadata schema `1.0` covers file/package information, document and revision properties, application/custom properties, sheets and hidden content, protection/encryption/signatures, macros, names/tables/charts/pivots, links/connections, comments, embedded objects, formulas/calculation, print settings, validations and XML package details. Every section reports an explicit `state`, `count`, `items`, `truncated`, and `warnings` envelope.

Profiles are `quick`, `standard`, `full`, and `forensic`. Quick mode avoids full worksheet scans. Full mode returns item inventories, and forensic mode can add hashes and package-part details. The reader never executes formulas, VBA, ActiveX, external links, data connections, Power Query or embedded files.

Update metadata atomically:

```php
Xlsx::updateMetaInfo('source.xlsx', 'updated.xlsx', [
    'document' => ['title' => 'Annual Report', 'creator' => 'MNB'],
    'custom_properties' => ['Project ID' => 'PRJ-1001'],
    'workbook' => ['active_sheet' => 'Summary'],
    'calculation' => ['mode' => 'auto'],
]);

Xlsx::removePersonalInfo('updated.xlsx', 'clean.xlsx');
```

Updates preserve unknown OOXML parts, arbitrary namespace prefixes and encryption mode. Encrypted input requires the existing password. The writable sections are document, revision, application, custom properties, workbook settings and calculation settings; feature objects remain owned by their dedicated APIs.
