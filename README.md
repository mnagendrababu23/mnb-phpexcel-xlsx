# mnb/mnb-phpexcel-xlsx

Native XLSX reader/writer with password encryption, document protection, formulas, pivots, and low-memory streaming for MNB PHPExcel.

This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.

## Install

```bash
composer require mnb/mnb-phpexcel-xlsx
```

See the main project documentation for typed options, streaming reads, and compatibility notes.

## Lightweight file and sheet information

Inspect package metadata, worksheet dimensions, and row counts without loading
worksheet records into PHP arrays:

```php
use Mnb\PHPExcel\Format\Xlsx;

$file = Xlsx::fileInfo('orders.xlsx');
$sheets = Xlsx::sheetsInfo('orders.xlsx');
$orders = Xlsx::sheetInfo('orders.xlsx', 'Orders');
$rows = Xlsx::rowCount('orders.xlsx', 'Orders');
$allRows = Xlsx::rowCounts('orders.xlsx');
```

`rowCount()` supports `filled`, `physical`, `last_row`, and `declared` modes.
The `declared` mode reads the worksheet dimension and is fastest. Other modes
scan worksheet XML without decoding cell values or creating workbook row arrays.

See the scripts in [`examples/`](examples/).

