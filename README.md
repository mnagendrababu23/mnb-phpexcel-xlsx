# MNB PHPExcel XLSX

<<<<<<< HEAD
Native XLSX reader/writer with password encryption, document protection, formulas, pivots, and low-memory streaming for MNB PHPExcel.
Documentation URL: https://mnbphpexcel.space/getting-started/installation
This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.

## Install
=======
Independent native XLSX reader/writer with streaming, formulas, styles, charts, pivots, comments, hyperlinks, protection, and Office password encryption.
>>>>>>> 6bc1ca7 (Release v2.0.0)

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
